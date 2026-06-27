from __future__ import annotations

from pathlib import Path
from typing import Dict

import pandas as pd

from layout_standard import STANDARD_INTERMEDIATE_COLUMNS, normalize_intermediate_columns
from parsers_excel.coelho_diniz import ler_excel_coelho_diniz
from parsers_excel.layouts_adicionais import ler_excel_buffon, ler_excel_mano_doces, ler_excel_vip
from parsers_excel.excel_emop_parteka import ler_excel_emop_parteka
from parsers_excel.excel_iquegami import ler_excel_iquegami
from parsers_excel.excel_generico_homologacao import ler_excel_generico_homologacao
from terminal_logger import get_terminal_logger
from pdf_alert_utils import alertas_para_dataframe


terminal_log = get_terminal_logger("leitor_excel")

INTERMEDIATE_COLUMNS = STANDARD_INTERMEDIATE_COLUMNS


def carregar_planilha(caminho, sheet_nome, header):
    import pandas as pd

    if sheet_nome is None or str(sheet_nome).strip() == "":
        # Lê todas as abas úteis quando o layout não fixa uma aba específica.
        # Isso evita truncar Excel multi-abas na primeira planilha sem avisar.
        planilhas = pd.read_excel(caminho, sheet_name=None, header=header, dtype=str)
        partes = []
        for nome_aba, df_aba in planilhas.items():
            if df_aba is None or df_aba.dropna(how="all").empty:
                continue
            df_aba = df_aba.copy()
            df_aba["_aba_origem"] = nome_aba
            partes.append(df_aba)
        if not partes:
            return pd.DataFrame()
        return pd.concat(partes, ignore_index=True, sort=False)

    # usa aba definida no layout ou índice explícito
    return pd.read_excel(caminho, sheet_name=sheet_nome, header=header, dtype=str)


def normalizar_nomes_colunas(df):
    df.columns = [
        str(col)
        .strip()
        .upper()
        .replace("Á", "A")
        .replace("Ã", "A")
        .replace("Ç", "C")
        for col in df.columns
    ]
    return df


def validar_colunas_obrigatorias(df: pd.DataFrame, mapeamentos_df: pd.DataFrame) -> None:
    faltantes = []
    for _, row in mapeamentos_df.iterrows():
        if str(row.get("origem_tipo", "")).upper() == "COLUNA" and str(row.get("obrigatorio", "")) == "1":
            col = str(row.get("origem_valor", "")).strip()
            if col and col not in df.columns:
                faltantes.append(col)
    if faltantes:
        raise ValueError(f"Colunas obrigatórias ausentes: {sorted(set(faltantes))}")


def extrair_campos_mapeados(df: pd.DataFrame, mapeamentos_df: pd.DataFrame) -> pd.DataFrame:
    resultado = pd.DataFrame(index=df.index)
    for _, row in mapeamentos_df.iterrows():
        destino = str(row.get("campo_destino", "")).strip()
        origem_tipo = str(row.get("origem_tipo", "")).strip().upper()
        origem_valor = str(row.get("origem_valor", "")).strip()
        if not destino:
            continue
        if origem_tipo == "COLUNA":
            colunas_possiveis = [
                c.strip().upper()
                for c in origem_valor.split("|")
            ]

            coluna_encontrada = None

            for col in df.columns:
                col_norm = (
                    str(col)
                    .strip()
                    .upper()
                    .replace("Á", "A")
                    .replace("Ã", "A")
                    .replace("Ç", "C")
                )
                if col_norm in colunas_possiveis:
                    coluna_encontrada = col
                    break

            if coluna_encontrada:
                resultado[destino] = df[coluna_encontrada]
            else:
                resultado[destino] = ""

        elif origem_tipo == "FIXO":
            resultado[destino] = origem_valor
    return resultado


def garantir_colunas_intermediarias(df: pd.DataFrame) -> pd.DataFrame:
    return normalize_intermediate_columns(df)


def adicionar_metadados(df: pd.DataFrame, caminho_arquivo: str, nome_layout: str) -> pd.DataFrame:
    return normalize_intermediate_columns(df, arquivo_origem=Path(caminho_arquivo).name, layout_usado=nome_layout)


def limpar_linhas_totalmente_vazias(df: pd.DataFrame) -> pd.DataFrame:
    campos_dados = [
        "matricula_lida",
        "cnpj_lido",
        "sku_lido",
        "codigo_sku_lido",
        "ean_lido",
        "quantidade_lida",
        "numero_pedido_lido",
        "data_entrega_lida",
    ]
    return df[~df[campos_dados].fillna("").astype(str).apply(lambda r: all(v.strip() == "" for v in r), axis=1)].reset_index(drop=True)


def ler_excel_cliente(caminho_arquivo: str, layout_config: Dict[str, str], mapeamentos_df: pd.DataFrame) -> Dict[str, object]:
    try:
        nome_layout = str(layout_config.get("nome_layout", "")).strip().upper()
        terminal_log.info("[EXCEL] Arquivo recebido: %s", caminho_arquivo)
        terminal_log.info("[EXCEL] Layout selecionado: %s", layout_config.get("nome_layout", ""))

        if "RASTREABILIDADE" in nome_layout and ("HOMOLOGACAO" in nome_layout or "GENERICA" in nome_layout):
            terminal_log.info("[EXCEL] Parser acionado: RASTREABILIDADE EXCEL GENERICA")
            return _finalizar_excel(ler_excel_generico_homologacao(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        if "COELHO DINIZ" in nome_layout:
            terminal_log.info("[EXCEL] Parser acionado: COELHO DINIZ")
            return _finalizar_excel(ler_excel_coelho_diniz(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        if "VIP" in nome_layout:
            terminal_log.info("[EXCEL] Parser acionado: REDE VIP")
            return _finalizar_excel(ler_excel_vip(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        if "MANO DOCES" in nome_layout or "MANOS DOCES" in nome_layout:
            terminal_log.info("[EXCEL] Parser acionado: MANO DOCES")
            return _finalizar_excel(ler_excel_mano_doces(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        if "BUFFON" in nome_layout:
            terminal_log.info("[EXCEL] Parser acionado: BUFFON")
            return _finalizar_excel(ler_excel_buffon(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        if "PARTEKA" in nome_layout or "E.M.O.P" in nome_layout or "EMOP" in nome_layout:
            terminal_log.info("[EXCEL] Parser acionado: GRUPO E.M.O.P. / PARTEKA")
            return _finalizar_excel(ler_excel_emop_parteka(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        if "IQUEGAMI" in nome_layout or "IKEGAMI" in nome_layout or "YQUEGAMI" in nome_layout:
            terminal_log.info("[EXCEL] Parser acionado: REDE IQUEGAMI")
            return _finalizar_excel(ler_excel_iquegami(caminho_arquivo, layout_config, mapeamentos_df), caminho_arquivo, layout_config)

        header_linha = int(str(layout_config.get("header_linha", "1") or "1")) - 1
        sheet_nome = layout_config.get("sheet_nome")
        df = carregar_planilha(caminho_arquivo, sheet_nome, header_linha)
        df = normalizar_nomes_colunas(df)
        validar_colunas_obrigatorias(df, mapeamentos_df)
        df_extraido = extrair_campos_mapeados(df, mapeamentos_df)
        df_extraido = garantir_colunas_intermediarias(df_extraido)
        df_extraido = adicionar_metadados(df_extraido, caminho_arquivo, layout_config.get("nome_layout", ""))
        df_extraido = limpar_linhas_totalmente_vazias(df_extraido)
        terminal_log.info("[EXCEL] Leitura concluida: %s linha(s) extraida(s).", len(df_extraido))
        return _finalizar_excel({
            "sucesso": True,
            "mensagem": "Leitura realizada com sucesso",
            "df_intermediario": df_extraido,
            "qtd_linhas_lidas": len(df_extraido),
            "qtd_linhas_planilha": len(df),
            "alertas": [],
        }, caminho_arquivo, layout_config)
    except Exception as e:
        terminal_log.exception("[EXCEL] Erro ao ler arquivo Excel: %s", caminho_arquivo)
        return {
            "sucesso": False,
            "mensagem": str(e),
            "df_intermediario": None,
            "qtd_linhas_lidas": 0,
            "alertas": [str(e)],
        }


def _finalizar_excel(resultado: Dict[str, object], caminho_arquivo: str, layout_config: Dict[str, str]) -> Dict[str, object]:
    df_intermediario = resultado.get("df_intermediario")
    if df_intermediario is not None:
        df_padrao = normalize_intermediate_columns(
            df_intermediario,
            arquivo_origem=caminho_arquivo,
            layout_usado=layout_config.get("nome_layout", ""),
        )
        resultado["df_intermediario"] = df_padrao
        resultado["qtd_itens_extraidos"] = len(df_padrao)
        resultado["qtd_linhas_lidas"] = max(int(resultado.get("qtd_linhas_lidas", 0) or 0), len(df_padrao))
    df_alertas = alertas_para_dataframe(
        resultado.get("alertas", []) or [],
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )
    if df_alertas is not None and not df_alertas.empty:
        resultado["df_itens_ignorados"] = df_alertas
        resultado["qtd_itens_ignorados"] = max(
            int(resultado.get("qtd_itens_ignorados", 0) or 0),
            len(df_alertas),
        )

    terminal_log.info(
        "[EXCEL] Resultado leitura | layout=%s | sucesso=%s | linhas_planilha=%s | itens=%s | alertas=%s | ignorados=%s",
        layout_config.get("nome_layout", ""),
        resultado.get("sucesso"),
        resultado.get("qtd_linhas_planilha", ""),
        resultado.get("qtd_itens_extraidos", resultado.get("qtd_linhas_lidas", 0)),
        len(resultado.get("alertas", []) or []),
        resultado.get("qtd_itens_ignorados", 0),
    )
    return resultado
