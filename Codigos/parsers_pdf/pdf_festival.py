from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Tuple

from pdf_alert_utils import linha_item_com_alerta
from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text, only_digits


RE_PEDIDO = re.compile(r"PEDIDO\s+DE\s+COMPRAS\s+(\d+\s*/\s*[A-Z])", re.I)
RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}")
RE_PREVISAO_ENTREGA = re.compile(r"Previs[aã]o\s+de\s+entrega\s+(\d{2}/\d{2}/\d{4})", re.I)
RE_LINHA_ITEM = re.compile(
    r"^\s*(?P<sku>\d{5,6})(?:\s+(?P<seq>\d{1,7})|(?P<seq_glued>\d))?\s*"
    r"(?P<desc>.+?)\s+(?P<emb>CX|FD|UN|PC|PT)\s+\d+\s+"
    r"(?P<qtd>\d{1,3}(?:\.\d{3})*,\d{2})\b",
    re.I,
)


def _parse_qtd(valor: str) -> str:
    texto = clean_text(valor).replace(".", "").replace(",", ".")
    try:
        qtd = Decimal(texto).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return ""
    if qtd <= 0:
        return ""
    return str(int(qtd))


def _extract_header(page_text: str) -> Tuple[str, str]:
    pedido = ""
    cnpj = ""

    match_pedido = RE_PEDIDO.search(page_text)
    if match_pedido:
        pedido = clean_text(match_pedido.group(1)).replace(" ", "").upper()

    cnpjs = RE_CNPJ.findall(page_text)
    if cnpjs:
        cnpj_cliente = next((valor for valor in cnpjs if only_digits(valor).startswith("78116670")), cnpjs[-1])
        cnpj = only_digits(cnpj_cliente)

    return pedido, cnpj


def _extract_data_entrega(page_text: str) -> str:
    match_data = RE_PREVISAO_ENTREGA.search(page_text)
    return clean_text(match_data.group(1)) if match_data else ""


def ler_pdf_festival(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    paginas = extract_pages_text(caminho_arquivo)

    datas_por_pedido: Dict[Tuple[str, str], str] = {}
    for texto in paginas:
        pedido, cnpj = _extract_header(texto)
        data_entrega = _extract_data_entrega(texto)
        if pedido and cnpj and data_entrega:
            datas_por_pedido[(pedido, cnpj)] = data_entrega

    linhas_saida: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedidos_encontrados = {}
    cnpjs_encontrados = {}
    vistos = set()
    linhas_lidas = 0
    linhas_validas = 0
    linhas_descartadas = 0

    for page_idx, texto in enumerate(paginas, start=1):
        pedido, cnpj = _extract_header(texto)
        if pedido:
            pedidos_encontrados[pedido] = pedidos_encontrados.get(pedido, 0) + 1
        if cnpj:
            cnpjs_encontrados[cnpj] = cnpjs_encontrados.get(cnpj, 0) + 1

        data_entrega = datas_por_pedido.get((pedido, cnpj), "")

        for linha in [clean_text(l) for l in texto.splitlines() if clean_text(l)]:
            if not re.match(r"^\d{5,7}", linha):
                continue

            match_item = RE_LINHA_ITEM.search(linha)
            if not match_item:
                continue

            linhas_lidas += 1
            chave_linha = (page_idx, pedido, cnpj, linha)
            if chave_linha in vistos:
                continue
            vistos.add(chave_linha)

            sku = only_digits(match_item.group("sku"))
            qtd = _parse_qtd(match_item.group("qtd"))

            if not pedido or not cnpj or not sku or not qtd or not data_entrega:
                linhas_descartadas += 1
                alerta = (
                    f"Pagina {page_idx}: item mantido para validacao por campo obrigatorio ausente "
                    f"(pedido={pedido or '-'}, cnpj={cnpj or '-'}, sku={sku or '-'}, "
                    f"qtd={qtd or '-'}, data={data_entrega or '-'})"
                )
                alertas.append(alerta)
                linhas_saida.append(linha_item_com_alerta(
                    caminho_arquivo=caminho_arquivo,
                    layout_usado=layout_config.get("nome_layout", ""),
                    pagina_pdf=page_idx,
                    linha_bruta=linha,
                    alerta=alerta,
                    cnpj_lido=cnpj,
                    sku_lido=sku,
                    quantidade_lida=qtd,
                    numero_pedido_lido=pedido,
                    data_entrega_lida=data_entrega,
                ))
                continue

            linhas_validas += 1
            linhas_saida.append(
                {
                    "matricula_lida": "",
                    "cnpj_lido": cnpj,
                    "sku_lido": sku,
                    "quantidade_lida": qtd,
                    "numero_pedido_lido": pedido,
                    "data_entrega_lida": data_entrega,
                }
            )

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    print("\n" + "=" * 100)
    print("DEBUG FESTIVAL")
    print("pedidos encontrados:", sorted(pedidos_encontrados.keys()))
    print("cnpjs encontrados:", sorted(cnpjs_encontrados.keys()))
    print("datas por pedido:", {f"{p}|{c}": d for (p, c), d in sorted(datas_por_pedido.items())})
    print("linhas brutas de item:", linhas_lidas)
    print("linhas validas:", linhas_validas)
    print("descartes no parser:", linhas_descartadas)
    print("linhas intermediarias totais:", len(df_intermediario))
    print("amostra df_intermediario:")
    print(df_intermediario.head(10).to_string(index=False) if not df_intermediario.empty else "<vazio>")
    print("=" * 100)

    if df_intermediario.empty:
        return {
            "sucesso": False,
            "mensagem": "Nenhuma linha valida foi extraida do PDF Festival",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "alertas": sorted(set(alertas)),
        }

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Festival concluida com {len(df_intermediario)} linha(s)",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "alertas": sorted(set(alertas)),
    }
