from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd

from layout_standard import normalize_intermediate_columns
from terminal_logger import get_terminal_logger

terminal_log = get_terminal_logger("excel_generico_homologacao")

ALIASES = {
    "cnpj_lido": ["CNPJ", "CNPJ CLIENTE", "CPF/CNPJ", "CGC"],
    "matricula_lida": ["MATRICULA", "MATRÍCULA", "COD CLIENTE", "CODIGO CLIENTE", "CÓDIGO CLIENTE", "CLIENTE", "LOJA"],
    "sku_lido": ["SKU", "COD SKU", "CODIGO SKU", "CÓDIGO SKU", "COD PRODUTO", "CODIGO PRODUTO", "CÓDIGO PRODUTO", "COD FORN", "COD.FORN", "COD FAB", "COD.FAB", "REF", "REFERENCIA", "REFERÊNCIA", "ITEM"],
    "codigo_sku_lido": ["SKU", "COD SKU", "CODIGO SKU", "CÓDIGO SKU", "COD PRODUTO", "CODIGO PRODUTO", "CÓDIGO PRODUTO", "COD FORN", "COD.FORN", "COD FAB", "COD.FAB", "REF", "REFERENCIA", "REFERÊNCIA", "ITEM"],
    "ean_lido": ["EAN", "DUN", "COD BARRAS", "CODIGO BARRAS", "CÓDIGO BARRAS", "BARRAS", "GTIN"],
    "quantidade_lida": ["QTD", "QTDE", "QUANT", "QUANTIDADE", "QTD PEDIDA", "QUANTIDADE PEDIDA", "EMB", "QTDE EMB", "QTD EMB", "CAIXA", "CX"],
    "numero_pedido_lido": ["PEDIDO", "N PEDIDO", "Nº PEDIDO", "NO PEDIDO", "NUM PEDIDO", "NUMERO PEDIDO", "NÚMERO PEDIDO", "ORDEM", "PO"],
    "data_entrega_lida": ["DATA", "ENTREGA", "DATA ENTREGA", "DT ENTREGA", "REMESSA", "DATA REMESSA", "PREVISAO", "PREVISÃO"],
    "descricao_lida": ["DESCRICAO", "DESCRIÇÃO", "PRODUTO", "ITEM", "NOME PRODUTO"],
}

CNPJ_RE = re.compile(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b")
NUM_RE = re.compile(r"(?<!\d)\d{1,14}(?:[\.,]\d{1,3})?(?!\d)")


def _norm(valor: object) -> str:
    texto = str(valor or "").upper()
    tabela = str.maketrans("ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ", "AAAAAEEEEIIIIOOOOOUUUUCN")
    texto = texto.translate(tabela)
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def _digits(valor: object) -> str:
    return re.sub(r"\D+", "", str(valor or ""))


def _normalizar_qtd(valor: object) -> str:
    texto = str(valor or "").strip()
    if not texto:
        return ""
    texto = texto.replace(" ", "")
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")
    try:
        numero = float(texto)
        if numero.is_integer():
            return str(int(numero))
        return str(numero).replace(".", ",")
    except Exception:
        return _digits(texto)


def _mapear_colunas_por_header(df: pd.DataFrame) -> tuple[int, Dict[str, int], int]:
    """Acha linha de cabeçalho e colunas prováveis por aliases."""
    melhor_linha = -1
    melhor_score = 0
    melhor_map: Dict[str, int] = {}
    limite = min(len(df), 60)
    aliases_norm = {campo: [_norm(a) for a in aliases] for campo, aliases in ALIASES.items()}

    for idx in range(limite):
        valores = [_norm(v) for v in df.iloc[idx].fillna("").tolist()]
        mapa: Dict[str, int] = {}
        score = 0
        for campo, aliases in aliases_norm.items():
            for col_idx, valor in enumerate(valores):
                if not valor:
                    continue
                if any(valor == alias or alias in valor or valor in alias for alias in aliases):
                    mapa[campo] = col_idx
                    score += 1
                    break
        # sku/qtd são essenciais para ser considerado bom cabeçalho
        if "quantidade_lida" in mapa and ("sku_lido" in mapa or "codigo_sku_lido" in mapa or "ean_lido" in mapa):
            score += 4
        if score > melhor_score:
            melhor_score = score
            melhor_linha = idx
            melhor_map = mapa
    return melhor_linha, melhor_map, melhor_score


def _extrair_por_header(df: pd.DataFrame, aba: str, caminho: str, nome_layout: str) -> pd.DataFrame:
    header_idx, mapa, score = _mapear_colunas_por_header(df)
    if header_idx < 0 or score < 4:
        return pd.DataFrame()

    linhas: List[Dict[str, str]] = []
    dados = df.iloc[header_idx + 1 :].reset_index(drop=True)
    for rel_idx, row in dados.iterrows():
        item: Dict[str, str] = {}
        for campo, col_idx in mapa.items():
            if col_idx < len(row):
                item[campo] = str(row.iloc[col_idx] or "").strip()

        sku = _digits(item.get("sku_lido") or item.get("codigo_sku_lido") or item.get("ean_lido"))
        qtd = _normalizar_qtd(item.get("quantidade_lida"))
        if not sku and not qtd and not _digits(item.get("cnpj_lido")) and not _digits(item.get("matricula_lida")):
            continue

        item["sku_lido"] = sku or item.get("sku_lido", "")
        item["codigo_sku_lido"] = item.get("codigo_sku_lido") or sku
        item["ean_lido"] = _digits(item.get("ean_lido"))
        item["quantidade_lida"] = qtd
        item["cnpj_lido"] = _digits(item.get("cnpj_lido"))
        item["matricula_lida"] = _digits(item.get("matricula_lida"))
        item["numero_pedido_lido"] = _digits(item.get("numero_pedido_lido"))
        item["aba_origem"] = aba
        item["linha_origem"] = str(header_idx + rel_idx + 2)
        item["origem_extracao"] = "EXCEL_GENERICO_HOMOLOGACAO_HEADER"
        item["status_extracao"] = "VALIDAR_HOMOLOGACAO"
        item["alerta_extracao"] = "Excel em homologacao/rastreabilidade: conferir colunas, SKU, QTD, pedido, CNPJ/GLN e matrícula antes de TXT/fila."
        linhas.append(item)

    return normalize_intermediate_columns(pd.DataFrame(linhas), arquivo_origem=Path(caminho).name, layout_usado=nome_layout)


def _extrair_por_linha(df: pd.DataFrame, aba: str, caminho: str, nome_layout: str) -> pd.DataFrame:
    linhas: List[Dict[str, str]] = []
    cnpj_ctx = ""
    matricula_ctx = ""
    pedido_ctx = ""
    for idx, row in df.fillna("").iterrows():
        texto = " ".join(str(v) for v in row.tolist() if str(v).strip())
        texto = re.sub(r"\s+", " ", texto).strip()
        if not texto:
            continue
        cnpj = CNPJ_RE.search(texto)
        if cnpj:
            cnpj_ctx = _digits(cnpj.group(0))
        # matrícula/pedido por contexto só quando há palavra de apoio próxima.
        if re.search(r"MATR[IÍ]CULA|COD\.?\s*CLIENTE|CLIENTE", texto, flags=re.I):
            nums = [_digits(n) for n in NUM_RE.findall(texto)]
            nums = [n for n in nums if 6 <= len(n) <= 12]
            if nums:
                matricula_ctx = nums[0]
        if re.search(r"PEDIDO|ORDEM|PO", texto, flags=re.I):
            nums = [_digits(n) for n in NUM_RE.findall(texto)]
            nums = [n for n in nums if 4 <= len(n) <= 12]
            if nums:
                pedido_ctx = nums[0]

        nums = [_digits(n) for n in NUM_RE.findall(texto)]
        nums = [n for n in nums if n]
        if len(nums) < 2:
            continue
        # Item candidato: primeiro código médio/longo como SKU/EAN e último número pequeno como QTD.
        sku = next((n for n in nums if 3 <= len(n) <= 14 and n not in {cnpj_ctx, pedido_ctx, matricula_ctx}), "")
        qtd = ""
        for n in reversed(nums):
            if n in {sku, cnpj_ctx, pedido_ctx, matricula_ctx}:
                continue
            if len(n) <= 6:
                qtd = n
                break
        if not sku or not qtd:
            continue
        linhas.append({
            "cnpj_lido": cnpj_ctx,
            "matricula_lida": matricula_ctx,
            "sku_lido": sku,
            "codigo_sku_lido": sku,
            "ean_lido": sku if len(sku) == 13 else "",
            "quantidade_lida": qtd,
            "numero_pedido_lido": pedido_ctx,
            "descricao_lida": texto[:240],
            "aba_origem": aba,
            "linha_origem": str(idx + 1),
            "origem_extracao": "EXCEL_GENERICO_HOMOLOGACAO_LINHA",
            "status_extracao": "VALIDAR_HOMOLOGACAO",
            "alerta_extracao": "Excel em homologacao/rastreabilidade sem cabecalho seguro: conferir item antes de TXT/fila.",
        })
    return normalize_intermediate_columns(pd.DataFrame(linhas), arquivo_origem=Path(caminho).name, layout_usado=nome_layout)


def ler_excel_generico_homologacao(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    nome_layout = str(layout_config.get("nome_layout", "RASTREABILIDADE EXCEL Homologacao Generica"))
    alertas = [
        f"LAYOUT_EXCEL_EM_HOMOLOGACAO: {nome_layout}. Conferir obrigatoriamente o Excel antes de TXT/fila.",
        "REFERENCIA_RASTREABILIDADE: EXCEL_GENERICO_HOMOLOGACAO",
    ]
    partes: List[pd.DataFrame] = []
    auditoria: List[dict] = []
    try:
        planilhas = pd.read_excel(caminho_arquivo, sheet_name=None, header=None, dtype=str)
    except Exception as exc:
        terminal_log.exception("[EXCEL_GENERICO] Falha ao abrir arquivo: %s", caminho_arquivo)
        return {"sucesso": False, "mensagem": str(exc), "df_intermediario": pd.DataFrame(), "qtd_linhas_lidas": 0, "alertas": alertas + [str(exc)]}

    for aba, df in planilhas.items():
        if df is None or df.dropna(how="all").empty:
            auditoria.append({"aba": aba, "linhas": 0, "status": "VAZIA"})
            continue
        df_header = _extrair_por_header(df, str(aba), caminho_arquivo, nome_layout)
        if df_header.empty:
            df_linha = _extrair_por_linha(df, str(aba), caminho_arquivo, nome_layout)
            df_extraido = df_linha
            modo = "LINHA"
        else:
            df_extraido = df_header
            modo = "HEADER"
        if not df_extraido.empty:
            partes.append(df_extraido)
        auditoria.append({"aba": aba, "linhas": len(df), "itens_extraidos": len(df_extraido), "modo": modo, "status": "OK"})

    df_final = pd.concat(partes, ignore_index=True, sort=False) if partes else pd.DataFrame()
    if df_final.empty:
        alertas.append("Nenhum item extraido pelo parser Excel generico de homologacao. Selecionar layout especifico ou criar parser dedicado.")

    terminal_log.info(
        "[EXCEL_GENERICO] arquivo=%s | abas=%s | itens=%s",
        Path(caminho_arquivo).name,
        len(planilhas),
        len(df_final),
    )
    return {
        "sucesso": not df_final.empty,
        "mensagem": "Leitura Excel generica de homologacao concluida" if not df_final.empty else "Nenhum item extraido no Excel em homologacao",
        "df_intermediario": df_final,
        "qtd_linhas_lidas": len(df_final),
        "qtd_itens_extraidos": len(df_final),
        "qtd_linhas_planilha": int(sum(len(df) for df in planilhas.values() if df is not None)),
        "df_auditoria_paginas": pd.DataFrame(auditoria),
        "alertas": sorted({str(a) for a in alertas if str(a).strip()}),
    }
