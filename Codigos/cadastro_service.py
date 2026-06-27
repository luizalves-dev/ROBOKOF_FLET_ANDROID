from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

import config

REQUIRED_LAYOUT_COLUMNS = [
    "layout_id", "cliente_id", "nome_layout", "tipo_arquivo", "tipo_cliente_destino",
    "ativo", "aceita_multiplos_pedidos", "aceita_multiplos_clientes", "sheet_nome",
    "header_linha", "inicio_dados_linha", "regra_data_entrega", "coluna_data_entrega", "observacoes"
]
REQUIRED_MAPEAMENTO_COLUMNS = [
    "mapeamento_id", "layout_id", "campo_destino", "origem_tipo", "origem_valor",
    "obrigatorio", "tipo_dado", "transformacao", "observacoes"
]


def _carregar_csv(caminho: Path) -> pd.DataFrame:
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    try:
        return pd.read_csv(caminho, dtype=str, encoding="utf-8").fillna("")
    except UnicodeDecodeError:
        return pd.read_csv(caminho, dtype=str, encoding="latin-1").fillna("")


def carregar_clientes() -> pd.DataFrame:
    return _carregar_csv(config.CADASTROS_DIR / "clientes.csv")


def carregar_layouts() -> pd.DataFrame:
    return _carregar_csv(config.CADASTROS_DIR / "layouts.csv")


def carregar_mapeamentos() -> pd.DataFrame:
    return _carregar_csv(config.CADASTROS_DIR / "mapeamento_campos.csv")


def carregar_regras_conversao() -> pd.DataFrame:
    caminho = config.CADASTROS_DIR / "regras_conversao.csv"
    return _carregar_csv(caminho) if caminho.exists() else pd.DataFrame()


def listar_layouts_ativos(tipo_arquivo: Optional[str] = None) -> pd.DataFrame:
    df = carregar_layouts()
    df = df[df["ativo"].astype(str) == "1"]
    if tipo_arquivo:
        df = df[df["tipo_arquivo"].str.upper() == tipo_arquivo.upper()]
    return df.reset_index(drop=True)


def buscar_layout(layout_id: str | int | None = None, nome_layout: Optional[str] = None) -> Optional[Dict[str, str]]:
    df = carregar_layouts()
    if layout_id is not None:
        filtrado = df[df["layout_id"].astype(str) == str(layout_id)]
    elif nome_layout:
        filtrado = df[df["nome_layout"].astype(str) == str(nome_layout)]
    else:
        return None
    if filtrado.empty:
        return None
    return filtrado.iloc[0].to_dict()


def buscar_mapeamentos_do_layout(layout_id: str | int) -> pd.DataFrame:
    df = carregar_mapeamentos()
    return df[df["layout_id"].astype(str) == str(layout_id)].reset_index(drop=True)


def validar_estrutura_cadastros() -> List[str]:
    erros: List[str] = []
    arquivos_regras = [
        (config.CADASTROS_DIR / "layouts.csv", REQUIRED_LAYOUT_COLUMNS),
        (config.CADASTROS_DIR / "mapeamento_campos.csv", REQUIRED_MAPEAMENTO_COLUMNS),
    ]
    for caminho, cols in arquivos_regras:
        if not caminho.exists():
            erros.append(f"Cadastro ausente: {caminho.name}")
            continue
        df = pd.read_csv(caminho, nrows=0)
        faltantes = [c for c in cols if c not in df.columns]
        if faltantes:
            erros.append(f"{caminho.name}: colunas ausentes {faltantes}")
    return erros
