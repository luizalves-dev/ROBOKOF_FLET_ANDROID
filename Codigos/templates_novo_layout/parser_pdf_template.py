# Autor: Kauê Melo
from __future__ import annotations

import pandas as pd


def ler_pdf_nova_rede(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    """Template seguro para layout PDF.

    Ler todas as páginas. Se uma página continuar o pedido anterior, manter o
    número anterior até aparecer novo número de pedido.
    """
    rows = []
    alertas = []

    # 1. Abrir PDF completo.
    # 2. Ler página a página.
    # 3. Detectar pedido, CNPJ correto, SKU/EAN e quantidade correta.
    # 4. Registrar página/linha/origem.

    df = pd.DataFrame(rows)
    return {
        "sucesso": not df.empty,
        "mensagem": f"NOVA_REDE: {len(df)} item(ns) extraído(s)",
        "df_intermediario": df,
        "qtd_linhas_lidas": len(df),
        "qtd_itens_extraidos": len(df),
        "alertas": alertas,
        "df_auditoria_paginas": pd.DataFrame(),
    }
