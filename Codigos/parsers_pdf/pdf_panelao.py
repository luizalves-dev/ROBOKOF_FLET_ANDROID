from __future__ import annotations

from pathlib import Path
import re
import unicodedata
from typing import Iterable, List

from parsers_pdf.pdf_utils import (
    build_intermediate_df,
    clean_text,
    extract_pages_text_detailed,
    normalize_qty,
    only_digits,
)
from terminal_logger import get_terminal_logger

try:
    import pdfplumber  # type: ignore
except ModuleNotFoundError:
    pdfplumber = None  # type: ignore


terminal_log = get_terminal_logger("pdf_panelao")

SUPPLIER_CNPJ_PREFIXES = {"61186888"}
CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
PEDIDO_RE = re.compile(r"PEDIDO\s+DE\s+COMPRA\s*:\s*([0-9]{3,12})", re.IGNORECASE)
ENTREGA_RE = re.compile(r"ENTREGA\s*/\s*RETIRADA\s*:\s*.*?([0-9]{2}/[0-9]{2}/[0-9]{4})", re.IGNORECASE)
ITEM_RE = re.compile(
    r"^\s*(?P<item>\d+)\s+"
    r"(?P<prod>\d{3,8})\s+"
    r"(?P<descricao>.+?)\s+"
    r"(?P<ref>\d{4,8})\s+"
    r"(?P<ean>\d{8,14})\s+"
    r"(?P<ncm>\d{6,10})\s+"
    r"(?P<emb>[A-Z]{1,4}\s*/\s*\d+)\s+"
    r"(?P<qtd_compra>\d+(?:[,.]\d+)?)\s+"
    r"(?P<qtd_total>\d+(?:[,.]\d+)?)\s+",
    re.IGNORECASE,
)
ITEM_CANDIDATO_RE = re.compile(r"^\s*\d+\s+\d{3,8}\s+.+?\s+\d{4,8}\s+\d{8,14}\s+\d{6,10}\s+", re.IGNORECASE)


def _normalizar_linha(valor: object) -> str:
    return re.sub(r"\s+", " ", str(valor or "").replace("\xa0", " ")).strip()


def _chave(valor: object) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).upper().strip()


def _eh_cnpj_cliente(cnpj: str) -> bool:
    digits = only_digits(cnpj)
    return len(digits) == 14 and not any(digits.startswith(prefix) for prefix in SUPPLIER_CNPJ_PREFIXES)


def _primeiro_cnpj_cliente(texto: str) -> str:
    for match in CNPJ_RE.findall(texto or ""):
        cnpj = only_digits(match)
        if _eh_cnpj_cliente(cnpj):
            return cnpj
    return ""


def _pedido(texto: str) -> str:
    match = PEDIDO_RE.search(_chave(texto))
    return only_digits(match.group(1)) if match else ""


def _data_entrega(texto: str) -> str:
    match = ENTREGA_RE.search(_chave(texto))
    return match.group(1) if match else ""


def _qtd_valida(valor: object) -> bool:
    qtd = normalize_qty(str(valor or ""))
    if not qtd:
        return False
    try:
        return float(qtd.replace(",", ".")) > 0
    except Exception:
        return False


def _montar_row(dados: dict[str, str], contexto: dict[str, str], pagina: int, linha_idx: int, linha: str) -> dict[str, str]:
    sku = only_digits(dados.get("ref"))
    ean = only_digits(dados.get("ean"))
    qtd = normalize_qty(dados.get("qtd_compra"))
    cnpj = contexto.get("cnpj", "")
    pedido = contexto.get("pedido", "")
    campos_ok = bool(cnpj and pedido and sku and qtd)
    return {
        "matricula_lida": "",
        "cnpj_lido": cnpj,
        "cnpj_base_lido": cnpj[:8] if cnpj else "",
        "codigo_cliente_lido": cnpj,
        "sku_lido": sku,
        "codigo_sku_lido": sku,
        "ean_lido": ean,
        "codigo_origem_lido": only_digits(dados.get("prod")) or ean,
        "descricao_lida": clean_text(dados.get("descricao")),
        "quantidade_lida": qtd,
        "numero_pedido_lido": pedido,
        "data_entrega_lida": contexto.get("data_entrega", ""),
        "pagina_pdf": str(pagina),
        "linha_origem": str(linha_idx),
        "linha_bruta": linha,
        "origem_extracao": "PDF_PANELAO_DEDICADO",
        "motor_extracao": "pdf_text",
        "status_extracao": "OK" if campos_ok else "VALIDAR_PANELAO",
        "alerta_extracao": "" if campos_ok else "Panelao: conferir CNPJ, pedido, SKU e quantidade extraidos.",
        "modo_rastreabilidade": "NAO",
        "layout_referencia": "PANELAO_PDF",
        "confianca_rastreabilidade": "PARSER_DEDICADO",
    }


def _layout_reconhecido(textos: Iterable[str]) -> bool:
    texto = _chave("\n".join(str(t or "") for t in textos))
    forte = any(token in texto for token in ["PANELAO", "SUPERMERCADO VIEIRA DIAS"])
    apoio = "PEDIDO DE COMPRA" in texto and ("QTDE COMPRA" in texto or "QTDE. COMPRA" in texto or "REF." in texto)
    return bool(forte and apoio)


def _indices_tabela(header: list[str]) -> dict[str, int]:
    normalizados = [_chave(cell) for cell in header]

    def buscar(*tokens: str) -> int:
        for idx, valor in enumerate(normalizados):
            if all(token in valor for token in tokens):
                return idx
        return -1

    ref_idx = buscar("REF")
    qtd_idx = buscar("QTDE", "COMPRA")
    return {
        "prod": buscar("PROD"),
        "descricao": buscar("DESCRICAO"),
        "ref": ref_idx,
        "ean": buscar("EAN"),
        "qtd": qtd_idx,
    }


def _cell(row: list[object], idx: int) -> str:
    if idx < 0 or idx >= len(row):
        return ""
    return _normalizar_linha(row[idx])


def _rows_por_tabela(caminho_arquivo: str, contexto: dict[str, str]) -> tuple[list[dict[str, str]], list[str]]:
    rows: list[dict[str, str]] = []
    alertas: list[str] = []
    if pdfplumber is None:
        return rows, ["PANELAO_TABELA_PDFPLUMBER_INDISPONIVEL"]

    with pdfplumber.open(caminho_arquivo) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            contexto["pedido"] = _pedido(text) or contexto.get("pedido", "")
            contexto["cnpj"] = _primeiro_cnpj_cliente(text) or contexto.get("cnpj", "")
            contexto["data_entrega"] = _data_entrega(text) or contexto.get("data_entrega", "")

            try:
                tables = page.extract_tables() or []
            except Exception as exc:
                alertas.append(f"PANELAO_ERRO_EXTRACT_TABLES | pagina={page_idx} | erro={exc}")
                continue

            for table_idx, table in enumerate(tables, start=1):
                header_idx = -1
                indices: dict[str, int] = {}
                for idx, raw_row in enumerate(table or []):
                    header = [_normalizar_linha(cell) for cell in (raw_row or [])]
                    key = _chave(" ".join(header))
                    if "REF" in key and "EAN" in key and "QTDE" in key and "COMPRA" in key:
                        header_idx = idx
                        indices = _indices_tabela(header)
                        break
                if header_idx < 0 or indices.get("ref", -1) < 0 or indices.get("qtd", -1) < 0:
                    continue

                for row_idx, raw_row in enumerate(table[header_idx + 1 :], start=header_idx + 2):
                    row = list(raw_row or [])
                    ref = only_digits(_cell(row, indices["ref"]))
                    qtd = _cell(row, indices["qtd"])
                    if not ref and not qtd:
                        continue
                    if not ref or not _qtd_valida(qtd):
                        alertas.append(
                            f"PANELAO_ITEM_TABELA_NAO_INTERPRETADO | pagina={page_idx} | tabela={table_idx} | linha={row_idx} | dados={row}"
                        )
                        continue
                    dados = {
                        "prod": _cell(row, indices.get("prod", -1)),
                        "descricao": _cell(row, indices.get("descricao", -1)),
                        "ref": ref,
                        "ean": _cell(row, indices.get("ean", -1)),
                        "qtd_compra": qtd,
                    }
                    linha = " | ".join(_normalizar_linha(cell) for cell in row)
                    item = _montar_row(dados, contexto, page_idx, row_idx, linha)
                    item["motor_extracao"] = "pdfplumber_table"
                    rows.append(item)

    return rows, alertas


def _rows_por_texto(textos: Iterable[str], contexto: dict[str, str]) -> tuple[list[dict[str, str]], list[str], int]:
    rows: list[dict[str, str]] = []
    alertas: list[str] = []
    linhas_lidas = 0

    for page_idx, texto in enumerate(textos, start=1):
        contexto["pedido"] = _pedido(texto) or contexto.get("pedido", "")
        contexto["cnpj"] = _primeiro_cnpj_cliente(texto) or contexto.get("cnpj", "")
        contexto["data_entrega"] = _data_entrega(texto) or contexto.get("data_entrega", "")

        for linha_idx, raw_line in enumerate(str(texto or "").splitlines(), start=1):
            linha = _normalizar_linha(raw_line)
            if not linha:
                continue
            linhas_lidas += 1
            match = ITEM_RE.match(linha)
            if match:
                rows.append(_montar_row(match.groupdict(), contexto, page_idx, linha_idx, linha))
                continue
            if ITEM_CANDIDATO_RE.match(linha):
                alertas.append(f"PANELAO_ITEM_NAO_INTERPRETADO | pagina={page_idx} | linha={linha_idx} | texto={linha[:220]}")

    return rows, alertas, linhas_lidas


def ler_pdf_panelao(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    audit = extract_pages_text_detailed(caminho_arquivo)
    contexto = {"cnpj": "", "pedido": "", "data_entrega": ""}
    alertas: List[str] = []

    if not _layout_reconhecido(audit.paginas):
        mensagem = "Layout invalido ou nao reconhecido para Panelao. Verifique se o PDF enviado corresponde ao padrao Panelao."
        terminal_log.error("[PANELAO] %s | arquivo=%s", mensagem, Path(caminho_arquivo).name)
        df_vazio = build_intermediate_df([], caminho_arquivo, layout_config.get("nome_layout", "PANELAO PDF"))
        return {
            "sucesso": False,
            "mensagem": mensagem,
            "df_intermediario": df_vazio,
            "qtd_linhas_lidas": sum(len(str(t or "").splitlines()) for t in audit.paginas),
            "qtd_itens_extraidos": 0,
            "paginas_pdf_total": audit.total_paginas,
            "paginas_pdf_processadas": audit.paginas_processadas,
            "paginas_pdf_sem_texto": int(sum(1 for a in audit.auditoria if not a.caracteres)),
            "motores_pdf": ", ".join(sorted({a.motor for a in audit.auditoria})),
            "df_auditoria_paginas": audit.auditoria_df(),
            "alertas": sorted({mensagem, *audit.alertas}),
        }

    rows, table_alerts = _rows_por_tabela(caminho_arquivo, contexto)
    alertas.extend(table_alerts)
    linhas_lidas = sum(len(str(t or "").splitlines()) for t in audit.paginas)

    if not rows:
        rows, text_alerts, linhas_lidas = _rows_por_texto(audit.paginas, contexto)
        alertas.extend(text_alerts)

    if rows and any(not row.get("cnpj_lido") for row in rows):
        alertas.append("PANELAO: ha item(ns) sem CNPJ da empresa/entrega identificado; conferir PDF/layout.")
    if rows and any(not row.get("numero_pedido_lido") for row in rows):
        alertas.append("PANELAO: ha item(ns) sem numero de pedido identificado; conferir PDF/layout.")

    df_intermediario = build_intermediate_df(rows, caminho_arquivo, layout_config.get("nome_layout", "PANELAO PDF"))
    alertas_final = sorted({str(a) for a in (alertas + audit.alertas) if str(a).strip()})
    sucesso = not df_intermediario.empty
    terminal_log.info(
        "[PANELAO] arquivo=%s | paginas=%s | linhas_lidas=%s | itens=%s | cnpjs=%s | pedidos=%s | alertas=%s",
        Path(caminho_arquivo).name,
        audit.total_paginas,
        linhas_lidas,
        len(df_intermediario),
        sorted({r.get("cnpj_lido", "") for r in rows if r.get("cnpj_lido")})[:10],
        sorted({r.get("numero_pedido_lido", "") for r in rows if r.get("numero_pedido_lido")})[:10],
        len(alertas_final),
    )
    return {
        "sucesso": sucesso,
        "mensagem": f"Leitura PDF Panelao concluida com {len(df_intermediario)} item(ns)" if sucesso else "Nenhum item extraido do PDF Panelao",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "qtd_itens_extraidos": len(df_intermediario),
        "paginas_pdf_total": audit.total_paginas,
        "paginas_pdf_processadas": audit.paginas_processadas,
        "paginas_pdf_sem_texto": int(sum(1 for a in audit.auditoria if not a.caracteres)),
        "motores_pdf": ", ".join(sorted({a.motor for a in audit.auditoria})),
        "df_auditoria_paginas": audit.auditoria_df(),
        "alertas": alertas_final,
    }
