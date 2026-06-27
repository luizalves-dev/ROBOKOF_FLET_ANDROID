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


terminal_log = get_terminal_logger("pdf_caita")

SUPPLIER_CNPJ_PREFIXES = {"61186888"}
CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
PEDIDO_RE = re.compile(r"NUMERO\s+DO\s+PEDIDO\s*:\s*([0-9]{3,12})", re.IGNORECASE)
PREVISAO_RE = re.compile(r"PREVISAO\s+DE\s+ENTREGA\s*:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", re.IGNORECASE)
EMPRESA_RE = re.compile(
    r"EMPRESA\s+DO\s+PEDIDO\s*:\s*(?P<codigo>\d+)\s+(?P<nome>.*?)(?:\s+NUMERO\s+DO\s+PEDIDO|\s+DATA\s+DO\s+PEDIDO|\s+FORNECEDOR|\s*$)",
    re.IGNORECASE,
)
ITEM_RE = re.compile(
    r"^\s*(?P<codigo>\d{3,9})\s+"
    r"(?P<gtin>\d{8,14})\s+"
    r"(?P<ref>\d{3,8})\s+"
    r"(?P<descricao>.+?)\s+"
    r"(?P<qtd_caixa>\d+(?:[,.]\d+)?)\s+"
    r"(?P<qtd_pcx>\d{1,4})\s+"
    r"(?P<qtd_pedida>\d+(?:[,.]\d+)?)\s+"
    r"(?P<preco>\d{1,9}(?:\.\d{3})*,\d{2,4})(?:\s+|$)",
    re.IGNORECASE,
)
ITEM_CANDIDATO_RE = re.compile(r"^\s*\d{3,9}\s+\d{8,14}\s+\d{3,8}\s+", re.IGNORECASE)


def _normalizar_linha(valor: object) -> str:
    return re.sub(r"\s+", " ", str(valor or "").replace("\xa0", " ")).strip()


def _chave(valor: object) -> str:
    texto = unicodedata.normalize("NFKD", str(valor or ""))
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).upper().strip()


def _eh_cnpj_cliente(cnpj: str) -> bool:
    digits = only_digits(cnpj)
    return len(digits) == 14 and not any(digits.startswith(prefix) for prefix in SUPPLIER_CNPJ_PREFIXES)


def _cnpjs_clientes(texto: str) -> list[str]:
    cnpjs: list[str] = []
    for match in CNPJ_RE.findall(texto or ""):
        cnpj = only_digits(match)
        if _eh_cnpj_cliente(cnpj) and cnpj not in cnpjs:
            cnpjs.append(cnpj)
    return cnpjs


def _primeiro_cnpj_cliente(texto: str) -> str:
    cnpjs = _cnpjs_clientes(texto)
    return cnpjs[0] if cnpjs else ""


def _pedido(texto: str) -> str:
    match = PEDIDO_RE.search(_chave(texto))
    return only_digits(match.group(1)) if match else ""


def _data_entrega(texto: str) -> str:
    match = PREVISAO_RE.search(_chave(texto))
    return match.group(1) if match else ""


def _empresa(texto: str) -> tuple[str, str]:
    match = EMPRESA_RE.search(_chave(texto))
    if not match:
        return "", ""
    return only_digits(match.group("codigo")), clean_text(match.group("nome"))


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
    gtin = only_digits(dados.get("gtin"))
    qtd = normalize_qty(dados.get("qtd_caixa"))
    cnpj = contexto.get("cnpj", "")
    pedido = contexto.get("pedido", "")
    campos_ok = bool(pedido and sku and qtd and (cnpj or contexto.get("cnpj_pendente_pos_processamento") == "SIM"))
    return {
        "matricula_lida": "",
        "cnpj_lido": cnpj,
        "cnpj_base_lido": cnpj[:8] if cnpj else "",
        "codigo_loja_lido": contexto.get("codigo_loja", ""),
        "codigo_cliente_lido": cnpj or contexto.get("codigo_loja", ""),
        "cod_cliente_lido": contexto.get("codigo_loja", ""),
        "loja_lida": contexto.get("loja", ""),
        "texto_loja_lido": contexto.get("loja", ""),
        "sku_lido": sku,
        "codigo_sku_lido": sku,
        "ean_lido": gtin,
        "codigo_origem_lido": only_digits(dados.get("codigo")) or gtin,
        "descricao_lida": clean_text(dados.get("descricao")),
        "quantidade_lida": qtd,
        "numero_pedido_lido": pedido,
        "data_entrega_lida": contexto.get("data_entrega", ""),
        "pagina_pdf": str(pagina),
        "linha_origem": str(linha_idx),
        "linha_bruta": linha,
        "origem_extracao": "PDF_CAITA_DEDICADO",
        "motor_extracao": "pdf_text",
        "status_extracao": "OK" if campos_ok else "VALIDAR_CAITA",
        "alerta_extracao": "" if campos_ok else "Caita: conferir CNPJ, pedido, SKU e quantidade extraidos.",
        "modo_rastreabilidade": "NAO",
        "layout_referencia": "CAITA_PDF",
        "confianca_rastreabilidade": "PARSER_DEDICADO",
    }


def _atualizar_contexto(texto: str, contexto: dict[str, str]) -> None:
    pedido = _pedido(texto)
    if pedido:
        contexto["pedido"] = pedido
    cnpj = _primeiro_cnpj_cliente(texto)
    if cnpj:
        contexto["cnpj"] = cnpj
    data = _data_entrega(texto)
    if data:
        contexto["data_entrega"] = data
    codigo_loja, loja = _empresa(texto)
    if codigo_loja:
        contexto["codigo_loja"] = codigo_loja
    if loja:
        contexto["loja"] = loja


def _layout_reconhecido(textos: Iterable[str]) -> bool:
    texto = _chave("\n".join(str(t or "") for t in textos))
    forte = any(token in texto for token in ["CAITA", "ZA SUPERMERCADOS"])
    apoio = "PEDIDOS DE COMPRA" in texto and "NUMERO DO PEDIDO" in texto and (
        "REFERENCIA" in texto or "QUANT CAIXA" in texto
    )
    return bool(forte and apoio)


def _indices_tabela(header: list[str]) -> dict[str, int]:
    normalizados = [_chave(cell) for cell in header]

    def buscar(*tokens: str) -> int:
        for idx, valor in enumerate(normalizados):
            if all(token in valor for token in tokens):
                return idx
        return -1

    codigo_idx = -1
    for idx, valor in enumerate(normalizados):
        if "CODIGO" in valor and "GTIN" not in valor:
            codigo_idx = idx
            break

    return {
        "codigo": codigo_idx,
        "gtin": buscar("GTIN"),
        "ref": buscar("REFERENCIA"),
        "descricao": buscar("DESCRICAO"),
        "qtd": buscar("QUANT", "CAIXA"),
    }


def _cell(row: list[object], idx: int) -> str:
    if idx < 0 or idx >= len(row):
        return ""
    return _normalizar_linha(row[idx])


def _rows_por_tabela(caminho_arquivo: str, contexto: dict[str, str]) -> tuple[list[dict[str, str]], list[str]]:
    rows: list[dict[str, str]] = []
    alertas: list[str] = []
    if pdfplumber is None:
        return rows, ["CAITA_TABELA_PDFPLUMBER_INDISPONIVEL"]

    with pdfplumber.open(caminho_arquivo) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            _atualizar_contexto(text, contexto)

            try:
                tables = page.extract_tables() or []
            except Exception as exc:
                alertas.append(f"CAITA_ERRO_EXTRACT_TABLES | pagina={page_idx} | erro={exc}")
                continue

            for table_idx, table in enumerate(tables, start=1):
                header_idx = -1
                indices: dict[str, int] = {}
                for idx, raw_row in enumerate(table or []):
                    header = [_normalizar_linha(cell) for cell in (raw_row or [])]
                    key = _chave(" ".join(header))
                    if "REFERENCIA" in key and "QUANT" in key and "CAIXA" in key:
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
                            f"CAITA_ITEM_TABELA_NAO_INTERPRETADO | pagina={page_idx} | tabela={table_idx} | linha={row_idx} | dados={row}"
                        )
                        continue
                    dados = {
                        "codigo": _cell(row, indices.get("codigo", -1)),
                        "gtin": _cell(row, indices.get("gtin", -1)),
                        "ref": ref,
                        "descricao": _cell(row, indices.get("descricao", -1)),
                        "qtd_caixa": qtd,
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
        _atualizar_contexto(texto, contexto)
        for linha_idx, raw_line in enumerate(str(texto or "").splitlines(), start=1):
            linha = _normalizar_linha(raw_line)
            if not linha:
                continue
            linhas_lidas += 1
            _atualizar_contexto(linha, contexto)
            match = ITEM_RE.match(linha)
            if match:
                rows.append(_montar_row(match.groupdict(), contexto, page_idx, linha_idx, linha))
                continue
            if ITEM_CANDIDATO_RE.match(linha):
                alertas.append(f"CAITA_ITEM_NAO_INTERPRETADO | pagina={page_idx} | linha={linha_idx} | texto={linha[:220]}")

    return rows, alertas, linhas_lidas


def _preencher_cnpj_posterior(rows: list[dict[str, str]], textos: Iterable[str], alertas: list[str]) -> None:
    cnpjs = []
    for texto in textos:
        for cnpj in _cnpjs_clientes(texto):
            if cnpj not in cnpjs:
                cnpjs.append(cnpj)

    if not rows:
        return
    if any(row.get("cnpj_lido") for row in rows):
        return
    if len(cnpjs) == 1:
        cnpj = cnpjs[0]
        for row in rows:
            row["cnpj_lido"] = cnpj
            row["cnpj_base_lido"] = cnpj[:8]
            row["codigo_cliente_lido"] = cnpj
            row["status_extracao"] = "OK" if row.get("numero_pedido_lido") and row.get("sku_lido") and row.get("quantidade_lida") else "VALIDAR_CAITA"
            if row["status_extracao"] == "OK":
                row["alerta_extracao"] = ""
        alertas.append("CAITA_CNPJ_APLICADO_POS_PROCESSAMENTO: CNPJ localizado em pagina/rodape e aplicado aos itens sem CNPJ.")
    elif len(cnpjs) > 1:
        alertas.append(
            "CAITA_MULTIPLOS_CNPJS_NO_ARQUIVO: itens sem CNPJ nao foram preenchidos automaticamente para evitar mistura de lojas."
        )


def ler_pdf_caita(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    audit = extract_pages_text_detailed(caminho_arquivo)
    contexto = {"cnpj": "", "pedido": "", "data_entrega": "", "codigo_loja": "", "loja": "", "cnpj_pendente_pos_processamento": "SIM"}
    alertas: List[str] = []

    if not _layout_reconhecido(audit.paginas):
        mensagem = "Layout invalido ou nao reconhecido para Caita. Verifique se o PDF enviado corresponde ao padrao Caita."
        terminal_log.error("[CAITA] %s | arquivo=%s", mensagem, Path(caminho_arquivo).name)
        df_vazio = build_intermediate_df([], caminho_arquivo, layout_config.get("nome_layout", "CAITA PDF"))
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

    _preencher_cnpj_posterior(rows, audit.paginas, alertas)

    if rows and any(not row.get("cnpj_lido") for row in rows):
        alertas.append("CAITA: ha item(ns) sem CNPJ da empresa/entrega identificado; conferir PDF/layout.")
    if rows and any(not row.get("numero_pedido_lido") for row in rows):
        alertas.append("CAITA: ha item(ns) sem numero de pedido identificado; conferir PDF/layout.")

    df_intermediario = build_intermediate_df(rows, caminho_arquivo, layout_config.get("nome_layout", "CAITA PDF"))
    alertas_final = sorted({str(a) for a in (alertas + audit.alertas) if str(a).strip()})
    sucesso = not df_intermediario.empty
    terminal_log.info(
        "[CAITA] arquivo=%s | paginas=%s | linhas_lidas=%s | itens=%s | cnpjs=%s | pedidos=%s | alertas=%s",
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
        "mensagem": f"Leitura PDF Caita concluida com {len(df_intermediario)} item(ns)" if sucesso else "Nenhum item extraido do PDF Caita",
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
