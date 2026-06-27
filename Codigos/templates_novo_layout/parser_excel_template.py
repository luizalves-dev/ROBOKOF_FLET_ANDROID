# Autor: Kauê Melo
from __future__ import annotations

import pandas as pd


def ler_excel_nova_rede(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    """Template seguro para layout Excel.

    Substitua NOVA_REDE pelo nome real da rede e mantenha o retorno no padrão
    df_intermediario do Robô KOF.
    """
    rows = []
    alertas = []

    # 1. Ler todas as abas/linhas.
    # 2. Identificar cabeçalho e colunas reais.
    # 3. Montar rows com campos *_lido.
    # 4. Registrar alertas, nunca descartar sem rastreio.

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
