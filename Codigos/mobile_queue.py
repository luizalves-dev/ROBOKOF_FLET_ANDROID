"""Liberação segura de um Excel de validação para a fila oficial RoboKOF."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd

import config

SHEETS_PREFERIDAS = (
    "FILA_KOF_VALIDACAO",
    "Modelo Robô KOF para Enviar",
    "MODELO_ROBOKOF",
    "PEDIDOS",
)
VALID_VALUES = {"SIM", "S", "YES", "Y", "VALIDADO", "OK", "CONFERIDO"}
STATUS_CONVERSAO_OK = {"", "OK CONVERTIDO", "OK SEM CONVERSÃO", "OK SEM CONVERSAO"}
COLUMN_ALIASES = {
    "Matricula": ("Matricula", "Matrícula", "MATRICULA"),
    "Sku": ("Sku", "SKU", "Código SKU", "Codigo SKU"),
    "Qtd": ("Qtd", "QTD", "Quantidade", "QTD Final"),
    "Nº Pedido": ("Nº Pedido", "N° Pedido", "Numero Pedido", "Número Pedido", "Nº do Pedido"),
    "Data remessa": ("Data remessa", "Data Remessa", "Data de remessa", "Data entrega", "Data Entrega"),
}


def _texto(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    value = str(value).strip()
    return value[:-2] if re.fullmatch(r"\d+\.0", value) else value


def _find_column(df: pd.DataFrame, aliases: tuple[str, ...]) -> str | None:
    normal = {str(c).strip().casefold(): c for c in df.columns}
    for alias in aliases:
        if alias.casefold() in normal:
            return str(normal[alias.casefold()])
    return None


def _read_best_sheet(path: Path) -> tuple[str, pd.DataFrame]:
    book = pd.ExcelFile(path)
    for sheet in SHEETS_PREFERIDAS:
        if sheet in book.sheet_names:
            if sheet == "Modelo Robô KOF para Enviar":
                # Alguns modelos antigos gravam o cabeçalho na quinta linha.
                for header in (0, 4, 3):
                    df = pd.read_excel(path, sheet_name=sheet, dtype=str, header=header).fillna("")
                    if _find_column(df, COLUMN_ALIASES["Matricula"]):
                        return sheet, df
            return sheet, pd.read_excel(path, sheet_name=sheet, dtype=str).fillna("")
    for sheet in book.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str).fillna("")
        if all(_find_column(df, aliases) for aliases in COLUMN_ALIASES.values()):
            return sheet, df
    raise ValueError("Nenhuma aba compatível com a fila KOF foi encontrada no Excel.")


def montar_fila_do_excel_validacao(caminho: str | Path) -> dict[str, object]:
    path = Path(caminho)
    if not path.exists():
        raise FileNotFoundError(f"Excel de validação não encontrado: {path}")

    sheet, source = _read_best_sheet(path)
    renamed: dict[str, str] = {}
    for target, aliases in COLUMN_ALIASES.items():
        found = _find_column(source, aliases)
        if not found:
            raise ValueError(f"Coluna obrigatória ausente: {target}")
        renamed[found] = target
    df = source.rename(columns=renamed).copy()

    validation_col = next(
        (c for c in config.ROBOKOF_COLUNAS_VALIDACAO_MANUAL if c in df.columns),
        None,
    )
    status_conversion_col = _find_column(df, ("Status Conversão", "Status Conversao"))

    safe_rows: list[dict[str, str]] = []
    alerts: list[str] = []
    for index, row in df.iterrows():
        line = index + 2
        values = {col: _texto(row.get(col, "")) for col in config.FILA_COLUMNS}
        reasons: list[str] = []
        if not values["Matricula"] or values["Matricula"].upper() in {"A CADASTRAR", "PENDENTE"}:
            reasons.append("matrícula ausente/pendente")
        if not values["Sku"]:
            reasons.append("SKU ausente")
        try:
            qtd = float(values["Qtd"].replace(".", "").replace(",", "."))
            if qtd <= 0 or not qtd.is_integer():
                reasons.append("quantidade deve ser inteira e maior que zero")
            else:
                values["Qtd"] = str(int(qtd))
        except Exception:
            reasons.append("quantidade inválida")
        if not values["Nº Pedido"]:
            reasons.append("pedido ausente")
        if not values["Data remessa"]:
            reasons.append("data de remessa ausente")
        if validation_col:
            status = _texto(row.get(validation_col, "")).upper()
            if status not in VALID_VALUES:
                reasons.append(f"não validado manualmente em '{validation_col}'")
        if status_conversion_col:
            status_conversion = _texto(row.get(status_conversion_col, "")).upper()
            if status_conversion not in STATUS_CONVERSAO_OK:
                reasons.append(f"conversão pendente: {status_conversion}")

        if reasons:
            alerts.append(f"Linha {line}: " + "; ".join(reasons))
            continue
        safe_rows.append(values)

    safe_df = pd.DataFrame(safe_rows, columns=config.FILA_COLUMNS)
    return {
        "df": safe_df,
        "alerts": alerts,
        "sheet": sheet,
        "total_rows": len(df),
        "safe_rows": len(safe_df),
        "blocked_rows": len(alerts),
        "manual_validation_column": validation_col or "",
    }
