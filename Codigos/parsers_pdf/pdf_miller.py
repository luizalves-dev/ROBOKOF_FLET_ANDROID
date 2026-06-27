from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text, only_digits
from parsers_pdf.ean_sku_map import resolve_ean_to_sku


RE_PEDIDO = re.compile(r"N[uú]mero\s+do\s+pedido:\s*(\d+)", re.I)
RE_PREVISAO = re.compile(r"Previs[aã]o\s+de\s+entrega:\s*(\d{2}/\d{2}/\d{4})", re.I)
RE_CNPJ = re.compile(r"CNPJ:\s*([\d./-]+)", re.I)

MILLER_GABARITOS = ("GABARITO MILLER 1 1.xlsx",)
MILLER_EAN_SKU_OVERRIDES = {
    "7894900530032": "92521",
    "7898770420042": "119605",
    "7894900664003": "119181",
    "7894900664010": "119182",
    "7894900664027": "119183",
    "78934115": "139367",
}


def _extract_referencia_from_line(line: str) -> str:
    """Extrai a coluna Referencia do item, usada como SKU/EAN no RoboKOF."""
    tokens = re.findall(r"\b\d+\b", line)
    if len(tokens) >= 2:
        referencia = only_digits(tokens[1])
        if 8 <= len(referencia) <= 14:
            return referencia

    m = re.search(r"\b\d{13,14}\b", line)
    if m:
        return only_digits(m.group(0))

    m = re.search(r"\b789\d{5}\b", line)
    if m:
        return only_digits(m.group(0))

    return ""


def _extract_qtd_caixas(line: str) -> str:
    encontrados = re.findall(r"\b(\d{1,6})\s*-\s*CX\b", line, re.I)
    if not encontrados:
        return ""
    return str(int(encontrados[-1]))


def _extract_cnpj_entrega(page_text: str) -> str:
    partes = re.split(r"ENDERE[CÇ]O\s+DE\s+ENTREGA", page_text, flags=re.I)
    if len(partes) < 2:
        return ""

    bloco = partes[-1]
    m = RE_CNPJ.search(bloco)
    return only_digits(m.group(1)) if m else ""


def _extract_header(page_text: str) -> Tuple[str, str]:
    pedido = ""
    data_entrega = ""

    m_pedido = RE_PEDIDO.search(page_text)
    if m_pedido:
        pedido = clean_text(m_pedido.group(1))

    m_data = RE_PREVISAO.search(page_text)
    if m_data:
        data_entrega = clean_text(m_data.group(1))

    return pedido, data_entrega


def ler_pdf_miller(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    paginas = extract_pages_text(caminho_arquivo)

    pedidos: Dict[str, Dict[str, object]] = {}
    pedido_atual: Optional[str] = None
    linhas_lidas = 0
    linhas_descartadas = 0
    alertas: List[str] = []
    vistos = set()

    for page_idx, texto in enumerate(paginas, start=1):
        pedido, data_entrega = _extract_header(texto)
        if pedido:
            pedido_atual = pedido
            pedidos.setdefault(
                pedido_atual,
                {
                    "pedido": pedido_atual,
                    "data_entrega": data_entrega,
                    "cnpj": "",
                    "itens": {},
                    "descartes": [],
                },
            )
            if data_entrega:
                pedidos[pedido_atual]["data_entrega"] = data_entrega

        if not pedido_atual:
            continue

        cnpj = _extract_cnpj_entrega(texto)
        if cnpj:
            pedidos[pedido_atual]["cnpj"] = cnpj

        for linha in [clean_text(l) for l in texto.splitlines() if clean_text(l)]:
            if "- CX" not in linha.upper():
                continue

            linhas_lidas += 1
            chave_linha = (page_idx, pedido_atual, linha)
            if chave_linha in vistos:
                continue
            vistos.add(chave_linha)

            referencia = _extract_referencia_from_line(linha)
            qtd = _extract_qtd_caixas(linha)
            if not referencia or not qtd:
                linhas_descartadas += 1
                alertas.append(f"Pedido {pedido_atual}: linha sem Referencia/EAN ou quantidade CX | {linha[:120]}")
                pedidos[pedido_atual]["descartes"].append({"sku": "", "ean": referencia, "qtd": qtd, "linha_bruta": linha})
                continue

            sku_final = resolve_ean_to_sku(referencia, MILLER_GABARITOS, MILLER_EAN_SKU_OVERRIDES) or referencia
            itens = pedidos[pedido_atual]["itens"]
            item_info = itens.setdefault(sku_final, {"qtd": 0, "ean": referencia})
            item_info["qtd"] = int(item_info.get("qtd", 0)) + int(qtd)
            if referencia and not item_info.get("ean"):
                item_info["ean"] = referencia

    linhas_saida: List[Dict[str, str]] = []

    for pedido, dados in pedidos.items():
        cnpj = clean_text(dados.get("cnpj", ""))
        data_entrega = clean_text(dados.get("data_entrega", ""))
        itens = dados.get("itens", {})
        descartes = dados.get("descartes", [])

        if not cnpj:
            alertas.append(f"Pedido {pedido}: CNPJ de entrega nao identificado")
        if not data_entrega:
            alertas.append(f"Pedido {pedido}: previsao de entrega nao identificada")

        for sku, info in sorted(itens.items(), key=lambda x: int(x[0]) if str(x[0]).isdigit() else 10**18):
            qtd = info.get("qtd", 0) if isinstance(info, dict) else info
            ean = info.get("ean", "") if isinstance(info, dict) else ""
            linhas_saida.append(
                {
                    "matricula_lida": "",
                    "cnpj_lido": cnpj,
                    "sku_lido": clean_text(sku),
                    "codigo_sku_lido": clean_text(sku),
                    "ean_lido": clean_text(ean),
                    "quantidade_lida": str(qtd),
                    "numero_pedido_lido": clean_text(pedido),
                    "data_entrega_lida": data_entrega,
                }
            )

        for descarte in descartes:
            linhas_saida.append(
                {
                    "matricula_lida": "",
                    "cnpj_lido": cnpj,
                    "sku_lido": clean_text(descarte.get("sku", "")),
                    "codigo_sku_lido": clean_text(descarte.get("sku", "")),
                    "ean_lido": clean_text(descarte.get("ean", "")),
                    "linha_bruta": clean_text(descarte.get("linha_bruta", "")),
                    "alerta_extracao": "Linha sem Referencia/EAN ou quantidade CX",
                    "quantidade_lida": clean_text(descarte.get("qtd", "")),
                    "numero_pedido_lido": clean_text(pedido),
                    "data_entrega_lida": data_entrega,
                }
            )

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    print("\n" + "=" * 100)
    print("DEBUG MILLER")
    print("pedidos encontrados:", list(pedidos.keys()))
    for pedido, dados in pedidos.items():
        print(
            f"pedido={pedido} | cnpj={dados.get('cnpj', '')} | "
            f"data={dados.get('data_entrega', '')} | itens={len(dados.get('itens', {}))}"
        )
    print("linhas brutas com - CX:", linhas_lidas)
    print("linhas intermediarias totais:", len(df_intermediario))
    print("descartes no parser:", linhas_descartadas)
    print("amostra df_intermediario:")
    print(df_intermediario.head(10).to_string(index=False) if not df_intermediario.empty else "<vazio>")
    print("=" * 100)

    if df_intermediario.empty:
        return {
            "sucesso": False,
            "mensagem": "Nenhuma linha valida foi extraida do PDF Miller",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "alertas": sorted(set(alertas)),
        }

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Miller concluida com {len(df_intermediario)} linha(s)",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "alertas": sorted(set(alertas)),
    }
