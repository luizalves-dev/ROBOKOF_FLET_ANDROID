from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import pandas as pd

from layout_standard import normalize_intermediate_columns


def _texto(valor) -> str:
    return str(valor or "").strip()


def alerta_para_linha(alerta: str, caminho_arquivo: str = "", layout_usado: str = "") -> dict:
    texto = _texto(alerta)
    pagina = ""
    linha_bruta = ""

    m_pag = re.search(r"P[áa]gina\s+(\d+)", texto, re.I)
    if m_pag:
        pagina = m_pag.group(1)

    if "|" in texto:
        linha_bruta = texto.split("|", 1)[1].strip()

    return {
        "matricula_lida": "",
        "cnpj_lido": "",
        "sku_lido": "",
        "quantidade_lida": "",
        "numero_pedido_lido": "",
        "data_entrega_lida": "",
        "arquivo_origem": Path(str(caminho_arquivo)).name if caminho_arquivo else "",
        "layout_usado": layout_usado,
        "pagina_pdf": pagina,
        "linha_origem": "",
        "linha_bruta": linha_bruta,
        "origem_extracao": "ALERTA_PARSER",
        "status_extracao": "ALERTA_LEITURA",
        "alerta_extracao": texto,
        "motivo_descarte": "ALERTA_LEITURA_PARSER",
    }


def alertas_para_dataframe(alertas: Iterable[object] | None, caminho_arquivo: str = "", layout_usado: str = "") -> pd.DataFrame:
    linhas = []
    for alerta in alertas or []:
        texto = _texto(alerta)
        if not texto:
            continue
        # Inclui alertas de parser/linha/pagina no Excel, sem criar item falso de pedido.
        if any(token in texto.upper() for token in [
            "LINHA DE ITEM", "ITEM DESCARTADO", "PAGINA", "PÁGINA", "SEM TEXTO", "NENHUM ITEM",
            "SEM CNPJ", "SEM NUMERO", "SEM NÚMERO", "SEM DATA", "NAO RECONHECIDA", "NÃO RECONHECIDA",
        ]):
            linhas.append(alerta_para_linha(texto, caminho_arquivo, layout_usado))
    return normalize_intermediate_columns(pd.DataFrame(linhas), arquivo_origem=caminho_arquivo, layout_usado=layout_usado) if linhas else pd.DataFrame()


def linha_item_com_alerta(
    *,
    caminho_arquivo: str,
    layout_usado: str,
    pagina_pdf: str | int = "",
    linha_bruta: str = "",
    alerta: str,
    cnpj_lido: str = "",
    sku_lido: str = "",
    quantidade_lida: str = "",
    numero_pedido_lido: str = "",
    data_entrega_lida: str = "",
    descricao_lida: str = "",
    codigo_sku_lido: str = "",
    ean_lido: str = "",
    codigo_origem_lido: str = "",
) -> dict:
    return {
        "matricula_lida": "",
        "cnpj_lido": str(cnpj_lido or "").strip(),
        "sku_lido": str(sku_lido or "").strip(),
        "codigo_sku_lido": str(codigo_sku_lido or sku_lido or "").strip(),
        "ean_lido": str(ean_lido or "").strip(),
        "codigo_origem_lido": str(codigo_origem_lido or "").strip(),
        "descricao_lida": str(descricao_lida or "").strip(),
        "quantidade_lida": str(quantidade_lida or "").strip(),
        "numero_pedido_lido": str(numero_pedido_lido or "").strip(),
        "data_entrega_lida": str(data_entrega_lida or "").strip(),
        "arquivo_origem": Path(str(caminho_arquivo)).name if caminho_arquivo else "",
        "layout_usado": layout_usado,
        "pagina_pdf": str(pagina_pdf or "").strip(),
        "linha_origem": "",
        "linha_bruta": str(linha_bruta or "").strip(),
        "origem_extracao": "PARSER_ITEM_COM_ALERTA",
        "motor_extracao": "PDF",
        "status_extracao": "PENDENTE_VALIDACAO",
        "alerta_extracao": str(alerta or "").strip(),
    }
