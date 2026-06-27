from __future__ import annotations

import re
from typing import Dict, List, Optional

from parsers_pdf.pdf_utils import (
    build_intermediate_df,
    clean_text,
    extract_pages_text,
    normalize_qty,
    only_digits,
)
from parsers_pdf.ean_sku_map import resolve_ean_to_sku


RE_PEDIDO = re.compile(r"N[uú]mero do pedido:\s*(\S+)", re.I)
RE_PREVISAO = re.compile(r"Previs[aã]o de entrega:\s*(\d{2}/\d{2}/\d{4})", re.I)
RE_CNPJ = re.compile(r"CNPJ:\s*([\d.\-\/]+)", re.I)

RE_ITEM_START = re.compile(r"^\s*(\d{3,10})\s+(\d{6,14})\s+(.+)$")
RE_DECIMAL = re.compile(r"^\d+(?:\.\d+)*,\d+$")
RE_INT = re.compile(r"^\d+$")

TISCHLER_GABARITOS = ("Gabarito Tischler - ATUALIZADO 2.xlsx",)
TISCHLER_EAN_SKU_OVERRIDES = {
    "7896388010556": "139765",
    "7791540127106": "139738",
    "7898770420042": "119605",
    "7894900664003": "119181",
    "7894900664010": "119182",
    "7894900664027": "119183",
}

IGNORE_PREFIXES = (
    "EMISSÃO DE PEDIDOS DE COMPRAS",
    "PRODUTO QUANT.",
    "Data do pedido:",
    "Tipo:",
    "Bonificação:",
    "Núm. no fornecedor:",
    "Fornecedor:",
    "Origem:",
    "Forma de pagamento:",
    "Condição de pagamento:",
    "Tipo de cobrança:",
    "Transportador:",
    "Comprador:",
    "Empresa do pedido:",
    "Gerou arquivo EDI:",
    "Faturamento:",
    "Referência",
    "Qtd",
    "Cx",
    "Valor Valor",
    "Bruto Líquido",
    "Total Total",
    "Itens Volumes",
    "Quantidade",
    "Tema:",
    "Destino:",
    "Informamos",
    "Observação:",
    "Processado por:",
    "* ",
    "Código Descrição Quantidade",
    "Uni Prç. efetivo Total",
    "TOTAL DAS TROCAS:",
)


def is_ignored_line(line: str) -> bool:
    if not line:
        return True

    for prefix in IGNORE_PREFIXES:
        if line.startswith(prefix):
            return True

    return False


def parse_item_line_tischler(line: str) -> Optional[Dict[str, str]]:
    """
    Exemplo real:
    111796 7894900531015 92520 AGUA MINERAL CRYSTAL 1,5L C/ GAS 06 8 48,000 2,61 125,28 2,61

    Regra do RoboKOF:
    - sku_lido = barcode
    - quantidade_lida = coluna Pedida
    """
    if not line or is_ignored_line(line):
        return None

    m = RE_ITEM_START.match(line)
    if not m:
        return None

    barcode = only_digits(m.group(2))
    resto = clean_text(m.group(3))
    tokens = resto.split()

    if len(tokens) < 6:
        return None

    if not RE_DECIMAL.match(tokens[-1]):
        return None
    if not RE_DECIMAL.match(tokens[-2]):
        return None
    if not RE_DECIMAL.match(tokens[-3]):
        return None
    if not RE_DECIMAL.match(tokens[-4]):
        return None
    if not RE_INT.match(tokens[-5]):
        return None

    quantidade_caixas = tokens[-5].strip()

    if not barcode or not quantidade_caixas:
        return None

    sku_final = resolve_ean_to_sku(barcode, TISCHLER_GABARITOS, TISCHLER_EAN_SKU_OVERRIDES) or barcode

    return {
        "sku_lido": sku_final,
        "codigo_sku_lido": sku_final,
        "ean_lido": barcode,
        "quantidade_lida": quantidade_caixas,
    }


def finalizar_bloco(
    bloco_atual: Optional[Dict[str, object]],
    linhas_saida: List[Dict[str, str]],
    alertas: List[str],
):
    if not bloco_atual:
        return

    pedido = str(bloco_atual.get("numero_pedido_lido", "")).strip()
    data = str(bloco_atual.get("data_entrega_lida", "")).strip()
    cnpj = str(bloco_atual.get("cnpj_lido", "")).strip()
    itens = bloco_atual.get("itens", [])

    if not itens:
        alertas.append(f"Pedido {pedido or '[sem número]'} sem itens reconhecidos")
        return

    if not cnpj:
        alertas.append(f"Pedido {pedido or '[sem número]'} sem CNPJ identificado")

    if not data:
        alertas.append(f"Pedido {pedido or '[sem número]'} sem data de entrega identificada")

    for item in itens:
        linhas_saida.append({
            "matricula_lida": "",
            "cnpj_lido": cnpj,
            "sku_lido": item.get("sku_lido", ""),
            "codigo_sku_lido": item.get("codigo_sku_lido", item.get("sku_lido", "")),
            "ean_lido": item.get("ean_lido", ""),
            "quantidade_lida": item.get("quantidade_lida", ""),
            "numero_pedido_lido": pedido,
            "data_entrega_lida": data,
        })


def ler_pdf_tischler(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    paginas = extract_pages_text(caminho_arquivo)

    linhas_saida: List[Dict[str, str]] = []
    alertas: List[str] = []

    bloco_atual: Optional[Dict[str, object]] = None
    em_endereco = False
    em_trocas = False
    pending_item_line: Optional[str] = None

    for texto in paginas:
        linhas = [clean_text(x) for x in texto.splitlines() if clean_text(x)]

        for linha in linhas:
            m_pedido = RE_PEDIDO.search(linha)
            if m_pedido:
                finalizar_bloco(bloco_atual, linhas_saida, alertas)
                bloco_atual = {
                    "numero_pedido_lido": clean_text(m_pedido.group(1)),
                    "data_entrega_lida": "",
                    "cnpj_lido": "",
                    "itens": [],
                }
                em_endereco = False
                em_trocas = False
                pending_item_line = None
                continue

            if bloco_atual is None:
                continue

            m_prev = RE_PREVISAO.search(linha)
            if m_prev and not bloco_atual["data_entrega_lida"]:
                bloco_atual["data_entrega_lida"] = clean_text(m_prev.group(1))
                continue

            if "ENDEREÇO DE ENTREGA" in linha.upper():
                em_endereco = True
                em_trocas = False
                pending_item_line = None
                continue

            if "TROCAS DO FORNECEDOR" in linha.upper():
                em_trocas = True
                pending_item_line = None
                continue

            if "TOTAL DAS TROCAS" in linha.upper():
                em_trocas = False
                pending_item_line = None
                continue

            if em_endereco:
                m_cnpj = RE_CNPJ.search(linha)
                if m_cnpj:
                    bloco_atual["cnpj_lido"] = only_digits(m_cnpj.group(1))
                continue

            if em_trocas:
                continue

            if is_ignored_line(linha):
                continue

            item = None

            if pending_item_line:
                combinado = f"{pending_item_line} {linha}"
                item = parse_item_line_tischler(combinado)
                pending_item_line = None

            if item is None:
                item = parse_item_line_tischler(linha)

            if item is None and RE_ITEM_START.match(linha):
                pending_item_line = linha
                continue

            if item is not None:
                bloco_atual["itens"].append(item)

    finalizar_bloco(bloco_atual, linhas_saida, alertas)

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    sucesso = not df_intermediario.empty
    mensagem = (
        f"Leitura PDF Tischler concluída com {len(df_intermediario)} linha(s)"
        if sucesso
        else "Nenhuma linha válida foi extraída do PDF Tischler"
    )

    return {
        "sucesso": sucesso,
        "mensagem": mensagem,
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": len(df_intermediario),
        "alertas": alertas,
    }
