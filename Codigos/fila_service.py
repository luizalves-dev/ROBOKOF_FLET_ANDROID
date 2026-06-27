from __future__ import annotations

from pathlib import Path
from typing import Dict

import pandas as pd

import config


def obter_caminho_fila() -> Path:
    return config.FILA_DIR / config.FILA_FILE_NAME


def criar_fila_se_nao_existir() -> Path:
    caminho = obter_caminho_fila()
    caminho.parent.mkdir(parents=True, exist_ok=True)
    if not caminho.exists():
        pd.DataFrame(columns=config.FILA_COLUMNS).to_excel(caminho, index=False)
    return caminho


def carregar_fila() -> pd.DataFrame:
    caminho = criar_fila_se_nao_existir()
    df = pd.read_excel(caminho, dtype=str).fillna("")
    if df.empty and list(df.columns) != config.FILA_COLUMNS:
        df = pd.DataFrame(columns=config.FILA_COLUMNS)
    validar_colunas_fila(df)
    return df


def validar_colunas_fila(df_fila: pd.DataFrame) -> None:
    faltantes = [c for c in config.FILA_COLUMNS if c not in df_fila.columns]
    if faltantes:
        raise ValueError(f"A fila não contém as colunas esperadas: {faltantes}")


def preparar_linhas_para_fila(df_final: pd.DataFrame) -> pd.DataFrame:
    faltantes = [c for c in config.FILA_COLUMNS if c not in df_final.columns]
    if faltantes:
        raise ValueError(f"DataFrame final sem colunas obrigatórias: {faltantes}")
    return df_final[config.FILA_COLUMNS].copy()


def salvar_fila(df_fila: pd.DataFrame) -> Path:
    caminho = criar_fila_se_nao_existir()
    df_fila.to_excel(caminho, index=False)
    return caminho


def inserir_na_fila(df_novos_pedidos: pd.DataFrame) -> Dict[str, object]:
    df_atual = carregar_fila()
    df_novos = preparar_linhas_para_fila(df_novos_pedidos).fillna("")
    df_final = pd.concat([df_atual, df_novos], ignore_index=True)
    caminho = salvar_fila(df_final)
    return {
        "sucesso": True,
        "linhas_inseridas": len(df_novos),
        "caminho": str(caminho),
    }
