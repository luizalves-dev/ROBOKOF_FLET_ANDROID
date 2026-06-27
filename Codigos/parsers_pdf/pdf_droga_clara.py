from __future__ import annotations

import re
from typing import Dict, List, Tuple

from pdf_alert_utils import linha_item_com_alerta
from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text, only_digits


RE_PEDIDO = re.compile(r"Numero\s+Pedido:\s*(\d+)", re.I)
RE_CNPJ_FILIAL = re.compile(r"Filial:\s*\d+\s+([\d./-]+)", re.I)
RE_LINHA_ITEM = re.compile(
    r"^\s*(?P<sku>\d{5,6})\s+.+?\s+(?P<und>UN|FD|CX|SDF)\s+"
    r"\d+\s*x\s*\d+\s+(?P<qtd>\d{1,8})\s+\d+[,.]\d{4}\b",
    re.I,
)


def _extract_header(page_text: str) -> Tuple[str, str]:
    pedido = ""
    cnpj = ""

    match_pedido = RE_PEDIDO.search(page_text)
    if match_pedido:
        pedido = clean_text(match_pedido.group(1))

    match_cnpj = RE_CNPJ_FILIAL.search(page_text)
    if match_cnpj:
        cnpj = only_digits(match_cnpj.group(1))

    return pedido, cnpj


def ler_pdf_droga_clara(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    paginas = extract_pages_text(caminho_arquivo)

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

        for linha in [clean_text(l) for l in texto.splitlines() if clean_text(l)]:
            if not re.match(r"^\d{5,6}\s+", linha):
                continue

            linhas_lidas += 1
            chave_linha = (page_idx, pedido, cnpj, linha)
            if chave_linha in vistos:
                continue
            vistos.add(chave_linha)

            match_item = RE_LINHA_ITEM.search(linha)
            if not match_item:
                linhas_descartadas += 1
                alertas.append(f"Pagina {page_idx}: linha de item nao reconhecida | {linha[:120]}")
                continue

            sku = only_digits(match_item.group("sku"))
            qtd = only_digits(match_item.group("qtd"))

            if not pedido or not cnpj or not sku or not qtd or int(qtd) <= 0:
                linhas_descartadas += 1
                alerta = (
                    f"Pagina {page_idx}: item mantido para validacao por campo obrigatorio ausente "
                    f"(pedido={pedido or '-'}, cnpj={cnpj or '-'}, sku={sku or '-'}, qtd={qtd or '-'})"
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
                    quantidade_lida=str(int(qtd)) if qtd and qtd.isdigit() and int(qtd) > 0 else qtd,
                    numero_pedido_lido=pedido,
                ))
                continue

            linhas_validas += 1
            linhas_saida.append(
                {
                    "matricula_lida": "",
                    "cnpj_lido": cnpj,
                    "sku_lido": sku,
                    "quantidade_lida": str(int(qtd)),
                    "numero_pedido_lido": pedido,
                    "data_entrega_lida": "",
                }
            )

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    print("\n" + "=" * 100)
    print("DEBUG DROGA CLARA")
    print("pedidos encontrados:", sorted(pedidos_encontrados.keys()))
    print("cnpjs encontrados:", sorted(cnpjs_encontrados.keys()))
    print("data remessa:", layout_config.get("regra_data_entrega", "D+1"))
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
            "mensagem": "Nenhuma linha valida foi extraida do PDF Droga Clara",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "alertas": sorted(set(alertas)),
        }

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Droga Clara concluida com {len(df_intermediario)} linha(s)",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "alertas": sorted(set(alertas)),
    }
