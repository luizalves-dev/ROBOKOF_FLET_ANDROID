from __future__ import annotations

from decimal import Decimal, InvalidOperation
import re
from typing import Any, Dict, List

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text_detailed, only_digits
from terminal_logger import get_terminal_logger

terminal_log = get_terminal_logger("pdf_super_ubialli")

PEDIDO_RE = re.compile(
    r"(?:N[ºO°]?\s*)?(?:PEDIDO|ORDEM\s+DE\s+COMPRA|OC)\s*(?:N[ºO°]?|NUMERO|NÚMERO)?\s*[:\-]?\s*(\d{4,12})",
    re.I,
)
CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}|\d{14}")
EAN_RE = re.compile(r"\b\d{8,14}\b")
SKU_RE = re.compile(r"\b0*\d{5,6}\b")
PRICE_RE = re.compile(r"^\d{1,7}[,.]\d{3,4}$")
QTY_RE = re.compile(r"^\d{1,7}(?:[,.]\d{1,2})?$")
HEADER_TOKENS = {
    "TOTAL", "TOTAIS", "VALOR", "UNIT", "UNITARIO", "UNITÁRIO", "PRECO", "PREÇO", "DESCONTO",
    "CODIGO", "CÓDIGO", "DESCRICAO", "DESCRIÇÃO", "PRODUTO", "EAN", "BARRAS", "QTDE", "QTD",
}
SUPPLIER_CNPJ_PREFIXES = {"61186888"}


def _decimal(value: Any) -> Decimal:
    text = str(value or "").strip().replace(" ", "")
    if not text:
        return Decimal("0")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        parts = text.split(".")
        if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]):
            text = "".join(parts)
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _qty_text(value: Any) -> str:
    number = _decimal(value)
    if number == number.to_integral_value():
        return str(int(number))
    return format(number.normalize(), "f").replace(".", ",")


def _is_header_or_total(line: str) -> bool:
    upper = line.upper()
    if any(token in upper for token in ("TOTAL GERAL", "TOTAL DO PEDIDO", "VALOR TOTAL", "SUBTOTAL")):
        return True
    if "CNPJ" in upper and not EAN_RE.search(line):
        return True
    return False


def _is_valid_qtd(token: str) -> bool:
    token = str(token or "").strip()
    if not QTY_RE.fullmatch(token):
        return False
    # Super Ubialli trafega em unidade; valores com 3/4 casas normalmente são preço/valor, não quantidade.
    if PRICE_RE.fullmatch(token):
        return False
    return _decimal(token) > 0


def _extract_context(text: str, pedido_atual: str, cnpj_atual: str) -> tuple[str, str]:
    pedido_match = PEDIDO_RE.search(text or "")
    if pedido_match:
        pedido_atual = pedido_match.group(1)

    for match in CNPJ_RE.findall(text or ""):
        cnpj = only_digits(match)
        if len(cnpj) == 14 and not any(cnpj.startswith(prefix) for prefix in SUPPLIER_CNPJ_PREFIXES):
            cnpj_atual = cnpj
            break
    return pedido_atual, cnpj_atual


def _tokens(line: str) -> list[str]:
    return [t.strip() for t in re.split(r"\s+", str(line or "").replace("\xa0", " ").strip()) if t.strip()]


def _extract_item_from_line(line: str) -> Dict[str, str] | None:
    line = " ".join(_tokens(line))
    if not line or _is_header_or_total(line):
        return None

    parts = _tokens(line)
    if len(parts) < 3:
        return None

    # Procura EAN e, a partir dele, busca SKU e QTD sem confundir preço/total com quantidade.
    ean_index = None
    ean = ""
    for idx, token in enumerate(parts):
        if EAN_RE.fullmatch(token):
            ean_index = idx
            ean = only_digits(token)
            break
    if ean_index is None:
        return None

    sku = ""
    sku_index = None
    for idx in range(ean_index + 1, min(len(parts), ean_index + 7)):
        token = parts[idx]
        if SKU_RE.fullmatch(token) and only_digits(token) != ean:
            sku = only_digits(token).lstrip("0") or only_digits(token)
            sku_index = idx
            break

    qtd = ""
    search_start = (sku_index + 1) if sku_index is not None else (ean_index + 1)
    for idx in range(search_start, min(len(parts), search_start + 6)):
        token = parts[idx]
        if _is_valid_qtd(token):
            # Se o próximo token for preço com 3/4 casas, esta é uma forte evidência de que token atual é a QTD.
            qtd = token
            break

    # Fallback: layouts com EAN -> DESCRICAO -> QTD -> PRECO, sem SKU no PDF.
    if not qtd:
        for idx in range(ean_index + 1, min(len(parts), ean_index + 12)):
            token = parts[idx]
            if _is_valid_qtd(token):
                next_token = parts[idx + 1] if idx + 1 < len(parts) else ""
                if PRICE_RE.fullmatch(next_token) or idx == len(parts) - 1:
                    qtd = token
                    break

    if not qtd:
        return None

    desc_tokens = parts[:ean_index]
    # Remove códigos soltos no início da descrição, preservando texto do produto.
    if desc_tokens and re.fullmatch(r"\d{1,7}", desc_tokens[0]):
        codigo_origem = only_digits(desc_tokens[0])
        desc_tokens = desc_tokens[1:]
    else:
        codigo_origem = ""
    descricao = clean_text(" ".join(desc_tokens))
    if not descricao:
        descricao = "ITEM SUPER UBIALLI"

    return {
        "matricula_lida": "",
        "cnpj_lido": "",
        "sku_lido": sku,
        "codigo_sku_lido": sku,
        "ean_lido": ean,
        "codigo_origem_lido": codigo_origem,
        "descricao_lida": descricao,
        "quantidade_lida": _qty_text(qtd),
        "qtd_original": _qty_text(qtd),
        "tipo_qtd_original": "UNIDADE",
        "numero_pedido_lido": "",
        "data_entrega_lida": "",
        "origem_extracao": "PARSER_SUPER_UBIALLI_PDF",
        "status_extracao": "OK",
    }


def ler_pdf_super_ubialli(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    audit = extract_pages_text_detailed(caminho_arquivo)
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedido_atual = ""
    cnpj_atual = ""

    for page_idx, text in enumerate(audit.paginas, start=1):
        pedido_atual, cnpj_atual = _extract_context(text, pedido_atual, cnpj_atual)
        itens_pagina = 0
        for raw_line in str(text or "").splitlines():
            parsed = _extract_item_from_line(raw_line)
            if not parsed:
                continue
            parsed["cnpj_lido"] = cnpj_atual
            parsed["numero_pedido_lido"] = pedido_atual
            parsed["pagina_pdf"] = str(page_idx)
            parsed["linha_bruta"] = clean_text(raw_line)
            rows.append(parsed)
            itens_pagina += 1
        terminal_log.info(
            "[SUPER_UBIALLI] página=%s | itens=%s | pedido_atual=%s | cnpj_atual=%s",
            page_idx,
            itens_pagina,
            pedido_atual or "-",
            cnpj_atual or "-",
        )

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Super Ubialli. Verifique se o PDF enviado corresponde ao padrão esperado."
        return {
            "sucesso": False,
            "mensagem": msg,
            "df_intermediario": build_intermediate_df([], caminho_arquivo, layout_config.get("nome_layout", "")),
            "qtd_linhas_lidas": 0,
            "alertas": [msg],
            "paginas_pdf_total": audit.total_paginas,
            "paginas_pdf_processadas": audit.paginas_processadas,
            "df_auditoria_paginas": audit.auditoria_df(),
        }

    if any(not row.get("cnpj_lido") for row in rows):
        alertas.append("Super Ubialli: há itens sem CNPJ identificado; manter no Excel para validação/de-para.")
    if any(not row.get("numero_pedido_lido") for row in rows):
        alertas.append("Super Ubialli: há itens sem número de pedido identificado; conferir cabeçalho do PDF.")
    if any(not row.get("sku_lido") and row.get("ean_lido") for row in rows):
        alertas.append("Super Ubialli: alguns itens vieram somente com EAN; SKU será resolvido pelo mapa de produtos na conversão.")

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Super Ubialli concluída com {len(rows)} item(ns).",
        "df_intermediario": build_intermediate_df(rows, caminho_arquivo, layout_config.get("nome_layout", "")),
        "qtd_linhas_lidas": len(rows),
        "qtd_itens_extraidos": len(rows),
        "alertas": sorted(set(alertas)),
        "paginas_pdf_total": audit.total_paginas,
        "paginas_pdf_processadas": audit.paginas_processadas,
        "df_auditoria_paginas": audit.auditoria_df(),
    }


# Função exposta para testes técnicos sem PDF real.
def testar_linha_super_ubialli(linha: str) -> Dict[str, str] | None:
    return _extract_item_from_line(linha)
