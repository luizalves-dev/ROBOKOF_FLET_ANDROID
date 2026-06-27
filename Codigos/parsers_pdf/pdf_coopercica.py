from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Tuple

from parsers_pdf.pdf_utils import (
    build_intermediate_df,
    clean_text,
    extract_pages_text_detailed,
    only_digits,
)
from terminal_logger import get_terminal_logger

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None


terminal_log = get_terminal_logger("pdf_coopercica")

RE_PEDIDO = re.compile(r"PEDIDO\s+DE\s+COMPRAS\s+(\d+)\s*/\s*([A-Z])", re.I)
RE_CNPJ_FULL = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}\s*-\s*\d{2}")
RE_CNPJ_TRAILING = re.compile(r"(\d{2}\.\d{3}\.\d{3}/\d{4})\s*-\s*(\d{2})")
RE_CNPJ_REVERSED = re.compile(r"-\s*(\d{2})\s+(\d{2}\.\d{3}\.\d{3}/\d{4})")
RE_LINHA_ITEM = re.compile(
    r"^\s*(?P<sku>\d{5,7})\s+\d+\s+.+?\s+(?P<emb>FD|CX|UN)\s+\d+\s+"
    r"(?P<qtd>\d{1,3}(?:\.\d{3})*,\d{2})\b",
    re.I,
)
RE_QTD_DECIMAL = re.compile(r"^\d{1,3}(?:\.\d{3})*(?:,\d{2,3})$")
RE_EMBALAGEM = re.compile(r"^(FD|CX|UN)\s+\d+\b", re.I)

# Regra de negócio validada pela saída de conferência do Coopercica:
# o SKU 56741 (pack promocional Coca-Cola 350ml LV/PG) aparece no PDF,
# mas não deve entrar no Modelo Robô KOF para Enviar. Mantemos rastreio no log/alertas.
SKUS_COOPERCICA_NAO_ENVIAR_MODELO = {"56741"}


def _extract_cnpjs(page_text: str) -> List[str]:
    texto = re.sub(r"\s+", " ", page_text or "")
    encontrados: List[str] = []

    for match in RE_CNPJ_FULL.findall(texto):
        cnpj = only_digits(match)
        if len(cnpj) == 14:
            encontrados.append(cnpj)

    for base, digito in RE_CNPJ_TRAILING.findall(texto):
        cnpj = only_digits(base) + only_digits(digito)
        if len(cnpj) == 14:
            encontrados.append(cnpj)

    for digito, base in RE_CNPJ_REVERSED.findall(texto):
        cnpj = only_digits(base) + only_digits(digito)
        if len(cnpj) == 14:
            encontrados.append(cnpj)

    unicos: List[str] = []
    for cnpj in encontrados:
        if cnpj not in unicos:
            unicos.append(cnpj)
    return unicos


def _extract_header(page_text: str) -> Tuple[str, str]:
    pedido = ""
    cnpj = ""

    match_pedido = RE_PEDIDO.search(page_text or "")
    if match_pedido:
        pedido = f"{clean_text(match_pedido.group(1))}/{clean_text(match_pedido.group(2)).upper()}"

    cnpjs = _extract_cnpjs(page_text)
    if cnpjs:
        cnpj_cliente = next((valor for valor in cnpjs if valor.startswith("50974732")), cnpjs[-1])
        cnpj = only_digits(cnpj_cliente)

    return pedido, cnpj


def _parse_qtd(valor: str) -> str:
    numero = clean_text(valor).replace(".", "").replace(",", ".")
    try:
        qtd = Decimal(numero).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return ""
    if qtd <= 0:
        return ""
    return str(int(qtd))


def _parse_item_linha_flex(linha: str) -> tuple[str, str] | None:
    match_item = RE_LINHA_ITEM.search(linha)
    if match_item:
        return only_digits(match_item.group("sku")), _parse_qtd(match_item.group("qtd"))

    if not re.search(r"\b(FD|CX|UN)\b", linha, re.I):
        return None

    codigos = re.findall(r"\b\d{5,14}\b", linha)
    if not codigos:
        return None

    sku = next((codigo for codigo in codigos if 5 <= len(codigo) <= 7), codigos[0])
    qtds = re.findall(r"\d{1,3}(?:\.\d{3})*(?:,\d{2,3})", linha)
    if not qtds:
        return None

    qtd = _parse_qtd(qtds[-1])
    if not sku or not qtd:
        return None
    return only_digits(sku), qtd


def _fitz_blocks_por_pagina(caminho_arquivo: str) -> dict[int, list[str]]:
    if fitz is None:
        return {}
    blocos_por_pagina: dict[int, list[str]] = {}
    with fitz.open(caminho_arquivo) as doc:
        for page_idx, page in enumerate(doc, start=1):
            blocos_por_pagina[page_idx] = [str(block[4] or "") for block in page.get_text("blocks") or []]
    return blocos_por_pagina


def _parse_itens_bloco_fitz(bloco: str) -> list[tuple[str, str, str]]:
    linhas = [clean_text(linha) for linha in str(bloco or "").splitlines() if clean_text(linha)]
    if not linhas:
        return []

    itens: list[tuple[str, str, str]] = []
    for idx_emb, linha_emb in enumerate(linhas):
        if not RE_EMBALAGEM.match(linha_emb):
            continue

        qtd = ""
        for anterior in reversed(linhas[:idx_emb]):
            if RE_QTD_DECIMAL.match(anterior):
                qtd = _parse_qtd(anterior)
                break

        sku = ""
        limite_sku = idx_emb
        for idx_sku in range(limite_sku - 1, -1, -1):
            candidato = linhas[idx_sku]
            if re.fullmatch(r"\d{5,7}", candidato):
                sku = only_digits(candidato)
                break

        if sku and qtd:
            itens.append((sku, qtd, linha_emb))

    return itens


def _build_row(
    *,
    pedido: str,
    cnpj: str,
    sku: str,
    qtd: str,
    pagina: int,
    motor: str,
    origem: str,
    bruto: str,
    status: str = "OK",
    alerta: str = "",
) -> Dict[str, str]:
    return {
        "matricula_lida": "",
        "cnpj_lido": cnpj,
        "sku_lido": sku,
        "codigo_sku_lido": sku,
        "ean_lido": "",
        "quantidade_lida": qtd,
        "numero_pedido_lido": pedido,
        "data_entrega_lida": "",
        "pagina_pdf": str(pagina),
        "motor_extracao": motor,
        "origem_extracao": origem,
        "linha_bruta": bruto[:500],
        "status_extracao": status,
        "alerta_extracao": alerta,
    }


def _registrar_alerta(alertas: List[str], mensagem: str):
    if mensagem:
        alertas.append(mensagem)
        terminal_log.warning("[COOPERCICA] %s", mensagem)


def ler_pdf_coopercica(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    extracao = extract_pages_text_detailed(caminho_arquivo)
    paginas = extracao.paginas
    auditoria = extracao.auditoria
    blocos_fitz = _fitz_blocks_por_pagina(caminho_arquivo)

    linhas_saida: List[Dict[str, str]] = []
    alertas: List[str] = list(extracao.alertas)
    pedidos_encontrados: Dict[str, int] = {}
    cnpjs_encontrados: Dict[str, int] = {}
    linhas_lidas = 0
    linhas_descartadas = 0
    paginas_com_item = 0
    ultimo_pedido = ""
    ultimo_cnpj = ""

    terminal_log.info(
        "[COOPERCICA] Iniciando leitura integral | arquivo=%s | paginas=%s",
        caminho_arquivo,
        extracao.total_paginas,
    )

    for page_idx, texto in enumerate(paginas, start=1):
        audit = auditoria[page_idx - 1] if page_idx - 1 < len(auditoria) else None
        motor = audit.motor if audit else "desconhecido"
        pedido, cnpj = _extract_header(texto)
        if pedido:
            ultimo_pedido = pedido
            pedidos_encontrados[pedido] = pedidos_encontrados.get(pedido, 0) + 1
        else:
            pedido = ultimo_pedido

        if cnpj:
            ultimo_cnpj = cnpj
            cnpjs_encontrados[cnpj] = cnpjs_encontrados.get(cnpj, 0) + 1
        else:
            cnpj = ultimo_cnpj

        linhas_pagina = 0
        linhas_textuais = [clean_text(l) for l in texto.splitlines() if clean_text(l)]
        for linha in linhas_textuais:
            if not re.match(r"^\d{5,7}\s+\d+\s+", linha):
                continue

            linhas_lidas += 1
            item_flex = _parse_item_linha_flex(linha)
            if not item_flex:
                linhas_descartadas += 1
                _registrar_alerta(alertas, f"Pagina {page_idx}: linha de item nao reconhecida | {linha[:160]}")
                continue

            sku, qtd = item_flex
            if sku in SKUS_COOPERCICA_NAO_ENVIAR_MODELO:
                linhas_descartadas += 1
                _registrar_alerta(
                    alertas,
                    f"Pagina {page_idx}: SKU {sku} ignorado conforme conferência Coopercica; item mantido fora do Modelo Robô KOF. Linha={linha[:160]}",
                )
                continue

            alerta = ""
            if not pedido or not cnpj:
                alerta = (
                    f"Pagina {page_idx}: item extraido com campo de cabecalho ausente "
                    f"(pedido={pedido or '-'}, cnpj={cnpj or '-'})"
                )
                _registrar_alerta(alertas, alerta)

            linhas_pagina += 1
            linhas_saida.append(
                _build_row(
                    pedido=pedido,
                    cnpj=cnpj,
                    sku=sku,
                    qtd=qtd,
                    pagina=page_idx,
                    motor=motor,
                    origem="texto_linha",
                    bruto=linha,
                    status="PENDENTE_CABECALHO" if alerta else "OK",
                    alerta=alerta,
                )
            )

        if linhas_pagina == 0:
            for bloco in blocos_fitz.get(page_idx, []):
                itens_bloco = _parse_itens_bloco_fitz(bloco)
                if not itens_bloco:
                    continue

                for sku, qtd, embalagem in itens_bloco:
                    linhas_lidas += 1
                    if sku in SKUS_COOPERCICA_NAO_ENVIAR_MODELO:
                        linhas_descartadas += 1
                        _registrar_alerta(
                            alertas,
                            f"Pagina {page_idx}: SKU {sku} ignorado por fallback fitz conforme conferência Coopercica; item mantido fora do Modelo Robô KOF.",
                        )
                        continue

                    alerta = ""
                    if not pedido or not cnpj:
                        alerta = (
                            f"Pagina {page_idx}: item bloco/fitz extraido com campo de cabecalho ausente "
                            f"(pedido={pedido or '-'}, cnpj={cnpj or '-'})"
                        )
                        _registrar_alerta(alertas, alerta)

                    linhas_pagina += 1
                    linhas_saida.append(
                        _build_row(
                            pedido=pedido,
                            cnpj=cnpj,
                            sku=sku,
                            qtd=qtd,
                            pagina=page_idx,
                            motor="fitz",
                            origem="fitz_bloco",
                            bruto=bloco,
                            status="PENDENTE_CABECALHO" if alerta else "OK",
                            alerta=alerta or f"Extraido por fallback fitz ({embalagem})",
                        )
                    )

        if linhas_pagina:
            paginas_com_item += 1
        elif texto:
            _registrar_alerta(alertas, f"Pagina {page_idx}: texto extraido, mas nenhum item Coopercica reconhecido")

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    motores = {}
    for item in auditoria:
        motores[item.motor] = motores.get(item.motor, 0) + 1
    paginas_sem_texto = sum(1 for item in auditoria if not item.caracteres)

    resumo_msg = (
        f"paginas={extracao.total_paginas}, processadas={extracao.paginas_processadas}, "
        f"paginas_com_item={paginas_com_item}, itens_extraidos={len(df_intermediario)}, "
        f"candidatos_lidos={linhas_lidas}, descartes={linhas_descartadas}, motores={motores}"
    )
    terminal_log.info("[COOPERCICA] Leitura finalizada | %s", resumo_msg)

    if df_intermediario.empty:
        if cnpjs_encontrados:
            alertas.append(f"CNPJ identificado, mas nenhum item Coopercica foi reconhecido: {sorted(cnpjs_encontrados.keys())}")
        if pedidos_encontrados:
            alertas.append(f"Pedido identificado, mas nenhum item Coopercica foi reconhecido: {sorted(pedidos_encontrados.keys())}")
        return {
            "sucesso": False,
            "mensagem": f"Nenhuma linha valida foi extraida do PDF Coopercica | {resumo_msg}",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "alertas": sorted(set(alertas)),
            "df_auditoria_paginas": extracao.auditoria_df(),
            "paginas_pdf_total": extracao.total_paginas,
            "paginas_pdf_processadas": extracao.paginas_processadas,
            "paginas_pdf_sem_texto": paginas_sem_texto,
            "motores_pdf": str(motores),
            "qtd_itens_extraidos": 0,
            "qtd_itens_ignorados": linhas_descartadas,
        }

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Coopercica concluida com {len(df_intermediario)} linha(s) | {resumo_msg}",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "alertas": sorted(set(alertas)),
        "df_auditoria_paginas": extracao.auditoria_df(),
        "paginas_pdf_total": extracao.total_paginas,
        "paginas_pdf_processadas": extracao.paginas_processadas,
        "paginas_pdf_sem_texto": paginas_sem_texto,
        "motores_pdf": str(motores),
        "qtd_itens_extraidos": len(df_intermediario),
        "qtd_itens_ignorados": linhas_descartadas,
    }
