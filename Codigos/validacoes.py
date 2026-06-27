import pandas as pd
import re
from decimal import Decimal, InvalidOperation
from typing import Tuple, List

from utils import normalize_sku, parse_qtd_to_int, sanitize_pedido, normalize_date_remessa, clean_str

REQUIRED_COLS = ["Matricula", "Sku", "Qtd", "Nº Pedido", "Data remessa"]


def normalizar_matricula_para_validacao(valor) -> str:
    """Normaliza matrícula para evitar falha de GLN por formatação do Excel.

    Exemplos tratados:
    - 1700273623
    - 1700273623.0
    - 1.700273623E+09
    - espaços/pontos/hífens acidentais
    """
    s = clean_str(valor)
    if not s or s.lower() in {"nan", "none"}:
        return ""

    normalizado = s.replace(",", ".")
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?(?:[eE][+-]?\d+)?", normalizado):
        try:
            dec = Decimal(normalizado)
            if dec == dec.to_integral_value():
                return str(int(dec))
        except (InvalidOperation, ValueError):
            pass

    if re.fullmatch(r"\d+\.0+", s):
        return s.split(".", 1)[0]

    return re.sub(r"\D+", "", s)


def normalizar_pedido_para_validacao(pedido_raw: str) -> str:
    """
    Regra especial:
    - se vier "." -> mantém "."
    - se vier vazio -> retorna ""
    - senão aplica sanitize_pedido normalmente
    """
    s = clean_str(pedido_raw)

    if s == ".":
        return "."

    if s == "":
        return ""

    return sanitize_pedido(s)


def _consolidar_skus_duplicados(df: pd.DataFrame) -> pd.DataFrame:
    """Consolida duplicidades operacionais de SKU sem bloquear a geração de TXT.

    Em algumas redes, como Daher e outros layouts já importados para a fila operacional,
    o mesmo SKU pode aparecer mais de uma vez no mesmo pedido/matrícula. O formato final
    do RoboKOF/TXT trabalha melhor com uma linha por SKU. Em vez de descartar a menor
    quantidade ou criar erro técnico, somamos as quantidades por:

        Matrícula + Nº Pedido + SKU + Data remessa

    Assim o robô não perde volume, não gera linha duplicada para o mesmo item e não
    bloqueia pedidos futuros por uma duplicidade que é tratável por consolidação.
    """
    if df is None or df.empty:
        return df

    group_cols = ["Matricula", "Pedido_norm", "Sku_norm", "Data_norm"]
    missing = [c for c in group_cols + ["Qtd_int"] if c not in df.columns]
    if missing:
        return df

    linhas = []
    for _, g in df.groupby(group_cols, dropna=False, sort=False):
        base = g.iloc[0].copy()
        if len(g) > 1:
            total = pd.to_numeric(g["Qtd_int"], errors="coerce").fillna(0).sum()
            base["Qtd_int"] = int(total)
            base["Qtd"] = str(int(total))
            base["Observacao Validacao"] = (
                f"SKU_DUPLICADO_CONSOLIDADO | linhas={len(g)} | qtd_total={int(total)}"
            )
        linhas.append(base)

    return pd.DataFrame(linhas).reset_index(drop=True)


def normalize_and_validate(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Retorna:
      df_ok: linhas elegíveis para gerar RoboKOF (após regras)
      df_err: linhas registradas como erro (coluna "Status Erro")

    Regra importante de produção:
    - duplicidade de SKU dentro do mesmo pedido/matrícula/data é consolidada somando QTD;
    - erros bloqueantes continuam indo para df_err para evitar TXT com item incompleto.
    """
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f"Colunas ausentes no arquivo de entrada: {missing}. Encontradas: {list(df.columns)}"
        )

    df = df.copy()

    # limpa espaços
    for c in REQUIRED_COLS:
        df[c] = df[c].apply(clean_str)

    # normaliza matrícula antes de consultar/gerar TXT, evitando falha quando o Excel
    # entrega 1700273623.0, notação científica ou pontuação acidental.
    df["Matricula"] = df["Matricula"].apply(normalizar_matricula_para_validacao)

    # normaliza
    df["Sku_norm"] = df["Sku"].apply(normalize_sku)
    df["Pedido_norm"] = df["Nº Pedido"].apply(normalizar_pedido_para_validacao)

    # qtd
    qtd_int, qtd_err = [], []
    for v in df["Qtd"].tolist():
        q, e = parse_qtd_to_int(v)
        qtd_int.append(q)
        qtd_err.append(e)
    df["Qtd_int"] = qtd_int
    df["Qtd_err"] = qtd_err

    # data
    data_norm, data_err = [], []
    for v in df["Data remessa"].tolist():
        d, e = normalize_date_remessa(v)
        data_norm.append(d)
        data_err.append(e)
    df["Data_norm"] = data_norm
    df["Data_err"] = data_err

    # força reavaliação segura da data
    df["Data_norm"] = df["Data_norm"].astype(str).replace("nan", "").str.strip()
    df.loc[df["Data_norm"] == "", "Data_norm"] = None

    err_parts = []

    # erros base
    # OBS:
    # "." agora é aceito e NÃO deve virar erro
    base_err_mask = (
        (df["Matricula"] == "")
        | (df["Sku_norm"] == "")
        | (df["Pedido_norm"] == "")
        | (df["Qtd_int"].isna())
        | (df["Data_norm"].isna())
    )

    if base_err_mask.any():
        tmp = df[base_err_mask].copy()

        def build_status(r):
            reasons: List[str] = []

            if r["Matricula"] == "":
                reasons.append("MATRICULA_VAZIA")

            if r["Sku_norm"] == "":
                reasons.append("SKU_VAZIO")

            if r["Pedido_norm"] == "":
                reasons.append("PEDIDO_VAZIO")

            if pd.isna(r["Qtd_int"]):
                reasons.append(str(r["Qtd_err"] or "QTD_INVALIDA"))

            if pd.isna(r["Data_norm"]):
                reasons.append(str(r["Data_err"] or "DATA_INVALIDA"))

            return " | ".join(reasons) if reasons else "ERRO_VALIDACAO"

        tmp["Status Erro"] = tmp.apply(build_status, axis=1)
        err_parts.append(tmp)
        df = df[~base_err_mask].copy()

    # SKU muito longo -> remove e registra erro bloqueante
    sku_len_mask = df["Sku_norm"].apply(lambda x: len(x) > 14)
    if sku_len_mask.any():
        tmp = df[sku_len_mask].copy()
        tmp["Status Erro"] = "SKU_MAIS_14"
        err_parts.append(tmp)
        df = df[~sku_len_mask].copy()

    # Duplicidade de SKU por (Matricula, Pedido_norm, Sku_norm, Data_norm)
    # Agora é consolidada por soma, em vez de virar erro técnico.
    df_ok = _consolidar_skus_duplicados(df)

    df_err = (
        pd.concat(err_parts, ignore_index=True)
        if err_parts
        else pd.DataFrame()
    )

    # prepara df_ok final
    if df_ok is not None and not df_ok.empty:
        df_ok["Sku"] = df_ok["Sku_norm"]
        df_ok["Nº Pedido"] = df_ok["Pedido_norm"]
        df_ok["Qtd"] = df_ok["Qtd_int"].astype(int)
        df_ok["Data remessa"] = df_ok["Data_norm"]
        df_ok_final = df_ok[REQUIRED_COLS].copy()
    else:
        df_ok_final = pd.DataFrame(columns=REQUIRED_COLS)

    # prepara df_err final
    if not df_err.empty:
        cols = [
            "Matricula",
            "Sku",
            "Sku_norm",
            "Qtd",
            "Qtd_int",
            "Nº Pedido",
            "Pedido_norm",
            "Data remessa",
            "Data_norm",
            "Status Erro",
        ]
        df_err_final = df_err[cols].copy()
    else:
        df_err_final = pd.DataFrame(columns=["Status Erro"])

    return df_ok_final, df_err_final
