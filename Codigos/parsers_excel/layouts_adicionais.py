from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, Iterable, List

import pandas as pd

from layout_standard import STANDARD_INTERMEDIATE_COLUMNS, normalize_intermediate_columns

INTERMEDIATE_COLUMNS = STANDARD_INTERMEDIATE_COLUMNS


def _only_digits(value: Any) -> str:
    return re.sub(r"\D+", "", str(value or "")).strip()


def _safe_int(value: Any) -> str:
    """Normaliza quantidade sem transformar 1.0 em 10.

    Correção importante para layouts matriciais Excel, especialmente Rede VIP:
    quando o Excel entrega a célula como número/float, o separador decimal já
    vem no padrão Python (1.0, 2.0). A regra antiga removia todo ponto antes
    de converter, gerando 1.0 -> 10.
    """
    if value is None or pd.isna(value):
        return ""

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            number = float(value)
        except Exception:
            return ""
    else:
        text = str(value).strip()
        if not text or text.lower() in {"nan", "none"}:
            return ""
        text = text.replace("R$", "").replace(" ", "")
        if "," in text and "." in text:
            text = text.replace(".", "").replace(",", ".")
        elif "," in text:
            text = text.replace(",", ".")
        elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
            text = text.replace(".", "")
        try:
            number = float(text)
        except ValueError:
            return ""

    if number <= 0:
        return ""
    return str(int(number)) if float(number).is_integer() else str(number).replace(".", ",")


def _ensure_columns(rows: List[Dict[str, str]], caminho_arquivo: str, nome_layout: str) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    return normalize_intermediate_columns(df, arquivo_origem=Path(caminho_arquivo).name, layout_usado=nome_layout)


def _result(caminho_arquivo: str, layout_config: Dict[str, str], rows: List[Dict[str, str]], rede: str, alertas: Iterable[str] | None = None):
    df = _ensure_columns(rows, caminho_arquivo, layout_config.get("nome_layout", ""))
    sucesso = not df.empty
    return {
        "sucesso": sucesso,
        "mensagem": (
            f"Leitura Excel {rede} concluida com {len(df)} linha(s)"
            if sucesso
            else f"Layout invalido ou nao reconhecido para {rede}. Verifique se o arquivo enviado corresponde ao padrao esperado."
        ),
        "df_intermediario": df if sucesso else None,
        "qtd_linhas_lidas": len(df),
        "alertas": sorted({a for a in (alertas or []) if a}),
    }


def _row(matricula: Any, sku: Any, qtd: Any, pedido: Any = "") -> Dict[str, str] | None:
    matricula_lida = _only_digits(matricula)
    sku_lido = _only_digits(sku)
    qtd_lida = _safe_int(qtd)

    if not matricula_lida or not sku_lido or not qtd_lida:
        return None

    return {
        "matricula_lida": matricula_lida,
        "cnpj_lido": "",
        "sku_lido": sku_lido,
        "codigo_sku_lido": sku_lido,
        "ean_lido": "",
        "descricao_lida": "",
        "quantidade_lida": qtd_lida,
        "numero_pedido_lido": str(pedido or "").strip(),
        "data_entrega_lida": "",
    }



def _valor_preenchido(value: Any) -> bool:
    if value is None or pd.isna(value):
        return False
    text = str(value).strip()
    return bool(text and text.lower() not in {"nan", "none"})


def ler_excel_vip(caminho_arquivo: str, layout_config: Dict[str, str], mapeamentos_df=None):
    df_raw = pd.read_excel(caminho_arquivo, header=None)
    header_row = None
    for idx, raw_row in df_raw.iterrows():
        if any(str(value).strip().upper() == "SKU" for value in raw_row.values):
            header_row = idx
            break

    if header_row is None or header_row <= 0:
        return _result(caminho_arquivo, layout_config, [], "Rede VIP", ["Rede VIP: cabecalho SKU nao encontrado"])

    matriculas = df_raw.iloc[header_row - 1].tolist()
    df = pd.read_excel(caminho_arquivo, header=header_row)
    df.columns = [str(col).strip() for col in df.columns]
    if "SKU" not in df.columns:
        return _result(caminho_arquivo, layout_config, [], "Rede VIP", ["Rede VIP: coluna SKU nao encontrada"])

    alertas: List[str] = []
    col_map: Dict[str, str] = {}
    for col, matricula in zip(df.columns, matriculas):
        if str(col).strip().upper() == "SKU":
            continue
        matricula_digits = _only_digits(matricula)
        if len(matricula_digits) >= 6:
            col_map[col] = matricula_digits

    if not col_map:
        alertas.append("Rede VIP: nenhuma coluna de matrícula reconhecida")

    # Ordem oficial da automação VIP: agrupa por matrícula/loja e lista todos os
    # SKUs daquela loja antes de seguir para a próxima coluna. Isso mantém a
    # mesma saída dos arquivos resultado_coca*.xlsx enviados para conferência.
    dados = df.dropna(how="all")
    rows: List[Dict[str, str]] = []
    for col, matricula in col_map.items():
        for linha_idx, record in dados.iterrows():
            sku = record.get("SKU", "")
            valor = record.get(col, "")
            parsed = _row(matricula, sku, valor, pedido=".")
            if parsed:
                rows.append(parsed)
            elif _valor_preenchido(valor):
                alertas.append(f"Rede VIP: linha {linha_idx + 1} coluna {col} ignorada por SKU/matrícula/QTD inválido | sku={sku} matricula={matricula} valor={valor}")

    return _result(caminho_arquivo, layout_config, rows, "Rede VIP", alertas)


def ler_excel_mano_doces(caminho_arquivo: str, layout_config: Dict[str, str], mapeamentos_df=None):
    df = pd.read_excel(caminho_arquivo, header=None)
    if df.shape[0] < 5 or df.shape[1] < 3:
        return _result(caminho_arquivo, layout_config, [], "Mano Doces", ["Mano Doces: estrutura minima nao encontrada"])

    matriculas = df.iloc[2]
    dados = df.iloc[4:].copy()
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []

    for col in range(2, df.shape[1]):
        matricula = _only_digits(matriculas[col])
        if not matricula:
            if any(_valor_preenchido(v) for v in dados.iloc[:, col].tolist()):
                alertas.append(f"Mano Doces: coluna {col + 1} possui quantidades, mas a matrícula do cabeçalho está vazia")
            continue
        for linha_idx, record in dados.iterrows():
            valor = record.iloc[col]
            parsed = _row(matricula, record.iloc[0], valor)
            if parsed:
                rows.append(parsed)
            elif _valor_preenchido(valor):
                alertas.append(f"Mano Doces: linha {linha_idx + 1} coluna {col + 1} ignorada por SKU/QTD inválido | matricula={matricula} valor={valor}")

    return _result(caminho_arquivo, layout_config, rows, "Mano Doces", alertas)


def _buffon_sku(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    match = re.search(r"\d{5,}", str(value))
    return match.group(0) if match else ""


def _buffon_header_row(df: pd.DataFrame) -> int | None:
    for idx in range(len(df)):
        count_skus = sum(1 for value in df.iloc[idx] if _buffon_sku(value))
        if count_skus >= 3:
            return idx
    return None


def ler_excel_buffon(caminho_arquivo: str, layout_config: Dict[str, str], mapeamentos_df=None):
    df = pd.read_excel(caminho_arquivo, header=None)
    header_row = _buffon_header_row(df)
    if header_row is None:
        return _result(caminho_arquivo, layout_config, [], "Buffon", ["Buffon: linha de SKUs nao encontrada"])

    sku_cols: Dict[int, str] = {}
    for col in df.columns:
        sku = _buffon_sku(df.iloc[header_row, col])
        if sku:
            sku_cols[col] = sku

    rows: List[Dict[str, str]] = []
    alertas: List[str] = []
    for idx in range(header_row + 1, len(df)):
        matricula = _only_digits(df.iloc[idx, 1] if df.shape[1] > 1 else "")
        if not matricula:
            if any(_valor_preenchido(df.iloc[idx, col]) for col in sku_cols):
                alertas.append(f"Buffon: linha {idx + 1} possui quantidades, mas matrícula está vazia")
            continue
        for col, sku in sku_cols.items():
            valor = df.iloc[idx, col]
            parsed = _row(matricula, sku, valor)
            if parsed:
                rows.append(parsed)
            elif _valor_preenchido(valor):
                alertas.append(f"Buffon: linha {idx + 1} coluna {col + 1} ignorada por QTD inválida | matricula={matricula} sku={sku} valor={valor}")

    return _result(caminho_arquivo, layout_config, rows, "Buffon", alertas)
