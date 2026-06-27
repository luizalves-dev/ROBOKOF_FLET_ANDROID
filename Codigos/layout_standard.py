from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd


BASE_INTERMEDIATE_COLUMNS = [
    "matricula_lida",
    "cnpj_lido",
    "sku_lido",
    "quantidade_lida",
    "numero_pedido_lido",
    "data_entrega_lida",
    "gln_lido",
    "cnpj_base_lido",
    "codigo_loja_lido",
    "codigo_cliente_lido",
    "cod_cliente_lido",
    "loja_lida",
    "texto_loja_lido",
    "arquivo_origem",
    "layout_usado",
]

AUDIT_INTERMEDIATE_COLUMNS = [
    "descricao_lida",
    "codigo_sku_lido",
    "ean_lido",
    "codigo_origem_lido",
    "pagina_pdf",
    "linha_origem",
    "linha_bruta",
    "origem_extracao",
    "motor_extracao",
    "status_extracao",
    "alerta_extracao",
    "modo_rastreabilidade",
    "layout_referencia",
    "confianca_rastreabilidade",
    "motivo_rastreabilidade",
    "centro_lido",
    "centro_referencia_conversao",
    "qtd_original",
    "tipo_qtd_original",
    "fator_conversao",
    "qtd_convertida",
    "qtd_final",
    "status_conversao",
    "tipo_regra_conversao",
    "regra_aplicada_conversao",
    "origem_regra_conversao",
    "prioridade_regra_conversao",
    "observacao_conversao",
    "status_identidade",
    "motivo_identidade",
    "alerta_identidade",
]

STANDARD_INTERMEDIATE_COLUMNS = BASE_INTERMEDIATE_COLUMNS + AUDIT_INTERMEDIATE_COLUMNS


def normalize_intermediate_columns(
    df: pd.DataFrame | None,
    *,
    arquivo_origem: str | Path | None = None,
    layout_usado: str | None = None,
    preserve_extra: bool = True,
) -> pd.DataFrame:
    """Padrao definitivo para retorno de parsers do Robo KOF.

    Mantem as colunas obrigatorias para fila/validacao e adiciona campos de
    auditoria que ajudam a comparar layouts que trabalham com SKU, EAN ou
    descricao. Colunas extras do parser sao preservadas no final.
    """
    if df is None:
        df = pd.DataFrame()
    else:
        df = df.copy()

    for col in STANDARD_INTERMEDIATE_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    if arquivo_origem is not None:
        nome = Path(str(arquivo_origem)).name
        df["arquivo_origem"] = df["arquivo_origem"].astype(str).where(
            df["arquivo_origem"].astype(str).str.strip() != "",
            nome,
        )

    if layout_usado is not None:
        layout = str(layout_usado or "")
        df["layout_usado"] = df["layout_usado"].astype(str).where(
            df["layout_usado"].astype(str).str.strip() != "",
            layout,
        )

    if "codigo_sku_lido" in df.columns:
        df["codigo_sku_lido"] = df["codigo_sku_lido"].astype(str).where(
            df["codigo_sku_lido"].astype(str).str.strip() != "",
            df["sku_lido"].astype(str),
        )

    ordered = list(STANDARD_INTERMEDIATE_COLUMNS)
    if preserve_extra:
        ordered += [col for col in df.columns if col not in ordered]
    return df[ordered].fillna("")


def empty_intermediate_df() -> pd.DataFrame:
    return pd.DataFrame(columns=STANDARD_INTERMEDIATE_COLUMNS)


def audit_summary_from_df(df: pd.DataFrame | None) -> dict:
    if df is None or df.empty:
        return {
            "itens_extraidos": 0,
            "itens_com_alerta": 0,
            "cnpjs_sem_matricula": 0,
        }
    work = normalize_intermediate_columns(df)
    return {
        "itens_extraidos": len(work),
        "itens_com_alerta": int((work["alerta_extracao"].astype(str).str.strip() != "").sum()),
        "cnpjs_sem_matricula": int(
            work.loc[
                (work["cnpj_lido"].astype(str).str.strip() != "")
                & (work["matricula_lida"].astype(str).str.strip() == ""),
                "cnpj_lido",
            ].nunique()
        ),
    }


def unique_alerts(values: Iterable[object] | None) -> list[str]:
    return sorted({str(value).strip() for value in (values or []) if str(value).strip()})
