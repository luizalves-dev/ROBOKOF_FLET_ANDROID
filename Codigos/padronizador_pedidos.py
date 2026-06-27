from __future__ import annotations

from datetime import datetime
from typing import Dict, List
from pathlib import Path

import re
import unicodedata
import pandas as pd

import config
import gln_service
import depara_clientes_service
from tratador_datas import resolver_data_entrega
from terminal_logger import get_terminal_logger
from conversao_service import aplicar_regras_conversao
from identidade_pedidos import validar_identidade_pedidos


terminal_log = get_terminal_logger("padronizador")


def _get_depara_clientes_path():
    """Retorna o caminho do cadastro complementar sem quebrar em versões parciais do config.py.

    Motivo: se o usuário aplicar um pacote incremental que sobrescreva o config.py
    sem a constante DE_PARA_CLIENTES_PATH, o processamento não deve parar com
    AttributeError. O fallback sempre aponta para Cadastros/de_para_clientes.csv.
    """
    caminho = getattr(config, "DE_PARA_CLIENTES_PATH", None) or getattr(config, "DEPARA_CLIENTES_PATH", None)
    if caminho:
        return Path(caminho)

    root_dir = Path(getattr(config, "ROOT_DIR", Path(__file__).resolve().parents[1]))
    caminho_fallback = root_dir / "Cadastros" / "de_para_clientes.csv"
    terminal_log.warning(
        "[PADRONIZADOR] config.DE_PARA_CLIENTES_PATH ausente; usando fallback: %s",
        caminho_fallback,
    )
    return caminho_fallback


def normalizar_campos_basicos(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()
    return df


def _candidatos_chave_cliente(row: pd.Series) -> list[str]:
    campos = [
        "cnpj_lido",
        "cnpj_oficial_lido",
        "cnpj_base_lido",
        "cnpj_raiz_lido",
        "cnpj_sem_dv_lido",
        "gln_lido",
        "texto_loja_lido",
        "codigo_loja_lido",
        "codigo_cliente_lido",
        "codigo_origem_lido",
        "cod_cliente_lido",
        "cliente_lido",
        "loja_lida",
    ]
    candidatos: list[str] = []
    for campo in campos:
        valor = str(row.get(campo, "") or "").strip()
        if valor and valor not in candidatos:
            candidatos.append(valor)
    return candidatos


def _resolver_cliente_linha(
    row: pd.Series,
    layout_config: Dict[str, str],
    mapa_cnpj: Dict[str, str],
    depara_entries: list[depara_clientes_service.DeParaEntry],
) -> dict[str, str]:
    """Resolve matrícula/CNPJ sem criar regra fixa por rede.

    Ordem de prioridade:
    1. Layout que já traz matrícula diretamente;
    2. BASE de GLNS.xlsx para CNPJ oficial;
    3. Cadastros/de_para_clientes.csv para GLN/código próprio/CNPJ alternativo;
    4. matrícula_lida como fallback rastreável, quando existir.
    """
    tipo = str(layout_config.get("tipo_cliente_destino", "")).upper().strip()
    nome_layout = str(layout_config.get("nome_layout", "") or "").strip()
    cnpj_lido = str(row.get("cnpj_lido", "") or "").strip()
    matricula_lida = gln_service.only_digits(row.get("matricula_lida", ""))

    resultado = {
        "matricula_final": "",
        "cnpj_oficial_final": depara_clientes_service.only_digits(cnpj_lido),
        "tipo_chave_depara": depara_clientes_service.infer_tipo_chave(cnpj_lido) if cnpj_lido else "",
        "chave_lida_depara": cnpj_lido,
        "status_depara_cliente": "",
        "observacao_depara_cliente": "",
    }

    if tipo == "MATRICULA":
        resultado["matricula_final"] = matricula_lida
        resultado["status_depara_cliente"] = "MATRICULA_DIRETA" if matricula_lida else "MATRICULA_VAZIA"
        return resultado

    if tipo == "CNPJ":
        chaves = _candidatos_chave_cliente(row)

        # MABY/SPAL usa CNPJ de 15 dígitos no PDF/arquivo esperado e o
        # cadastro validado fica no de_para_clientes.csv. Tentar a BASE GLNS
        # primeiro gera falso alerta de CNPJ sem matrícula, apesar de o de/para
        # complementar estar correto. Para esta rede, prioriza o cadastro
        # complementar e só depois cai no GLN geral.
        if any(chave in nome_layout.upper() for chave in ["MABY", "ESTRELA"]):
            match_rede = depara_clientes_service.buscar_depara_por_chaves(nome_layout, chaves, entries=depara_entries)
            if match_rede:
                resultado.update({
                    "matricula_final": match_rede.get("matricula", ""),
                    "cnpj_oficial_final": match_rede.get("cnpj_oficial", "") or depara_clientes_service.only_digits(cnpj_lido),
                    "tipo_chave_depara": match_rede.get("tipo_chave", ""),
                    "chave_lida_depara": match_rede.get("chave_lida", "") or cnpj_lido,
                    "status_depara_cliente": "DEPARA_CLIENTES",
                    "observacao_depara_cliente": (
                        f"Resolvido via de_para_clientes.csv | rede={match_rede.get('rede', '')} | "
                        f"tipo={match_rede.get('tipo_chave', '')} | chave={match_rede.get('chave_lida', '')}"
                    ),
                })
                return resultado

        matricula_gln = gln_service.buscar_matricula_por_cnpj(
            cnpj_lido,
            gln_base_path=config.GLN_BASE_PATH,
            sheet_name=config.GLN_SHEET_NAME,
            col_cnpj=config.GLN_COL_CNPJ,
            col_matricula=config.GLN_COL_MATRICULA,
            mapa_cache=mapa_cnpj,
        )
        if matricula_gln:
            resultado["matricula_final"] = matricula_gln
            resultado["status_depara_cliente"] = "BASE_GLNS_CNPJ"
            return resultado

        match = depara_clientes_service.buscar_depara_por_chaves(nome_layout, chaves, entries=depara_entries)
        if match:
            resultado.update({
                "matricula_final": match.get("matricula", ""),
                "cnpj_oficial_final": match.get("cnpj_oficial", "") or depara_clientes_service.only_digits(cnpj_lido),
                "tipo_chave_depara": match.get("tipo_chave", ""),
                "chave_lida_depara": match.get("chave_lida", "") or cnpj_lido,
                "status_depara_cliente": "DEPARA_CLIENTES",
                "observacao_depara_cliente": (
                    f"Resolvido via de_para_clientes.csv | rede={match.get('rede', '')} | "
                    f"tipo={match.get('tipo_chave', '')} | chave={match.get('chave_lida', '')}"
                ),
            })
            return resultado

        if matricula_lida:
            resultado["matricula_final"] = matricula_lida
            resultado["status_depara_cliente"] = "MATRICULA_LIDA_LAYOUT"
            resultado["observacao_depara_cliente"] = "Matrícula veio do próprio layout/parser; recomendado centralizar no de_para_clientes.csv."
            return resultado

        resultado["status_depara_cliente"] = "A_CADASTRAR"
        resultado["observacao_depara_cliente"] = "Chave de cliente sem matrícula localizada; cadastrar no de_para_clientes.csv ou BASE de GLNS.xlsx."
        return resultado

    resultado["status_depara_cliente"] = "TIPO_CLIENTE_NAO_SUPORTADO"
    return resultado


def resolver_matricula_final(df: pd.DataFrame, layout_config: Dict[str, str]) -> pd.DataFrame:
    nome_layout = str(layout_config.get("nome_layout", "") or "").upper()
    tipo_cliente = str(layout_config.get("tipo_cliente_destino", "") or "").upper().strip()

    # Layouts que já trazem matrícula direta, como Rede VIP/Iquegami/Coelho,
    # não precisam carregar BASE GLNS nem de_para_clientes.csv. Isso evita
    # lentidão e logs desnecessários sem alterar a regra de negócio.
    if tipo_cliente == "MATRICULA":
        mapa = {}
        depara_entries = []
    elif "MABY" in nome_layout or "ESTRELA" in nome_layout:
        # MABY/ESTRELA usam de/para complementar validado no de_para_clientes.csv.
        # Evita alertas falsos e evita que a BASE GLNS prevaleça sobre a regra da rede.
        mapa = {}
        depara_entries = depara_clientes_service.carregar_depara_clientes(_get_depara_clientes_path())
    else:
        mapa = gln_service.load_cnpj_to_matricula_map(
            config.GLN_BASE_PATH,
            config.GLN_SHEET_NAME,
            config.GLN_COL_CNPJ,
            config.GLN_COL_MATRICULA,
        )
        depara_entries = depara_clientes_service.carregar_depara_clientes(_get_depara_clientes_path())
    df = df.copy()
    resolucoes = df.apply(lambda row: _resolver_cliente_linha(row, layout_config, mapa, depara_entries), axis=1)
    df_res = pd.DataFrame(list(resolucoes), index=df.index, dtype=str) if len(resolucoes) else pd.DataFrame(index=df.index)
    for coluna in [
        "matricula_final",
        "cnpj_oficial_final",
        "tipo_chave_depara",
        "chave_lida_depara",
        "status_depara_cliente",
        "observacao_depara_cliente",
    ]:
        df[coluna] = df_res[coluna].fillna("").astype(str) if coluna in df_res.columns else ""
    return df


def resolver_numero_pedido(df: pd.DataFrame, layout_config: Dict[str, str]) -> pd.DataFrame:
    df = df.copy()
    tipo_arquivo = str(layout_config.get("tipo_arquivo", "")).upper()
    pedido = df.get("numero_pedido_lido", pd.Series([""] * len(df), index=df.index)).fillna("").astype(str).str.strip()
    if tipo_arquivo == "EXCEL":
        pedido = pedido.replace("", ".")
    df["numero_pedido_final"] = pedido
    return df


def resolver_data_remessa(df: pd.DataFrame, layout_config: Dict[str, str], data_base: datetime | None = None) -> pd.DataFrame:
    df = df.copy()
    regra = layout_config.get("regra_data_entrega", "D+1")
    df["data_remessa_final"] = df["data_entrega_lida"].apply(lambda valor: resolver_data_entrega(regra, valor, data_base))
    return df


def _status_conversao_bloqueia_fila(row: pd.Series) -> bool:
    """Retorna True quando a linha ainda não pode ir para fila/TXT.

    Itens com conversão pendente devem aparecer no Excel de validação, mas não
    podem alimentar o Modelo Robô KOF/Fila automaticamente. Isso evita que uma
    quantidade em unidade seja tratada como caixaria quando o fator não foi
    encontrado ou quando a regra ficou ambígua.
    """
    status = str(row.get("status_conversao", "") or "").strip().upper()
    status = status.replace("NAO", "NÃO").replace("CONVERSAO", "CONVERSÃO")
    if not status:
        return False
    return status not in {"OK CONVERTIDO", "OK SEM CONVERSÃO"}




def _upper_sem_acento_local(value) -> str:
    text = str(value or "").strip().upper()
    return unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("ASCII")


def _is_layout_dedicado_sem_conversao_texto(value) -> bool:
    t = _upper_sem_acento_local(value)
    return any(nome in t for nome in ("ALABARCE", "MONACO", "PRIMATO", "DAHER", "SEMPRE VALE", "BAKLIZI", "BAZKILI", "MABY", "BOZZA", "BOZA", "ESTRELA", "COOPERCICA", "REDE VIP", "IQUEGAMI", "REDE IQUEGAMI"))

def _only_digits_local(value) -> str:
    return re.sub(r"\D+", "", str(value or ""))


def _linha_alabarce_pronta_para_modelo(row: pd.Series) -> bool:
    """Liberação do Modelo para layouts dedicados sem conversão.

    Mantém o nome antigo para compatibilidade, mas agora cobre também
    Monaco/Mônaco e Primato. RASTREABILIDADE_LAYOUT é alerta de auditoria
    quando os campos de negócio estão completos; não deve zerar o modelo.
    """
    texto = " ".join([
        str(row.get("layout_usado", "") or ""),
        str(row.get("layout_referencia", "") or ""),
        str(row.get("alerta_extracao", "") or ""),
        str(row.get("alerta_identidade", "") or ""),
        str(row.get("motivo_identidade", "") or ""),
        str(row.get("status_conversao", "") or ""),
        str(row.get("regra_aplicada_conversao", "") or ""),
    ])
    if not _is_layout_dedicado_sem_conversao_texto(texto):
        return False

    bloqueadores = [
        "SKU_EAN_NAO_LOCALIZADO", "SKU_VAZIO", "QTD_VAZIA", "QTD_NAO_NUMERICA",
        "QTD_ZERO_OU_NEGATIVA", "PEDIDO_VAZIO", "PEDIDO_AUSENTE",
        "CNPJ_COM_TAMANHO_INVALIDO", "CHAVE_CLIENTE_NAO_LOCALIZADA",
        "LAYOUT_INVALIDO", "LAYOUT INVALIDO", "CONVERSAO_PENDENTE",
        "ALERTA - NAO CONVERTIDO", "VALIDAR CONVERSAO", "QTD_PARECE_PRECO_CUSTO",
        "TOTAL_DIVERGENTE",
    ]
    texto_norm = _upper_sem_acento_local(texto)
    if any(tok in texto_norm for tok in bloqueadores):
        return False

    matricula = str(row.get("matricula_final", "") or row.get("matricula_lida", "") or "").strip()
    cnpj = _only_digits_local(row.get("cnpj_oficial_final", "") or row.get("cnpj_lido", "") or row.get("cnpj_base_lido", ""))
    sku = _only_digits_local(row.get("sku_lido", "") or row.get("codigo_sku_lido", ""))
    qtd = str(row.get("qtd_final", "") or row.get("quantidade_lida", "") or "").strip()
    pedido = str(row.get("numero_pedido_final", "") or row.get("numero_pedido_lido", "") or "").strip()
    data = str(row.get("data_remessa_final", "") or row.get("data_entrega_lida", "") or "").strip()
    if not matricula:
        return False
    layout_maby = "MABY" in _upper_sem_acento_local(texto)
    if len(cnpj) != 14 and not (layout_maby and len(cnpj) == 15 and cnpj.startswith("01695774")):
        return False
    if not sku:
        return False
    if not qtd:
        return False
    if not pedido:
        return False
    if not data:
        return False
    return True


def _liberar_identidade_alabarce_sem_conversao(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    for idx, row in df.iterrows():
        if _linha_alabarce_pronta_para_modelo(row):
            # Liberação segura para layouts dedicados sem conversão: alerta de
            # rastreabilidade/auditoria não pode esvaziar o Modelo Robô KOF
            # quando os campos mínimos estão preenchidos.
            df.at[idx, "status_identidade"] = "OK"
            df.at[idx, "motivo_identidade"] = "OK"
            df.at[idx, "alerta_identidade"] = ""
            layout_atual = str(df.at[idx, "layout_usado"] if "layout_usado" in df.columns else "")
            layout_ref = str(df.at[idx, "layout_referencia"] if "layout_referencia" in df.columns else "")
            if _upper_sem_acento_local(layout_atual).startswith("RASTREABILIDADE ->") and layout_ref:
                df.at[idx, "layout_usado"] = layout_ref
    return df


def _status_identidade_bloqueia_fila(row: pd.Series) -> bool:
    status = str(row.get("status_identidade", "") or "").strip().upper()
    if not status:
        return False
    if status in {"OK", "IDENTIDADE OK"}:
        return False
    if _linha_alabarce_pronta_para_modelo(row):
        return False

    # Correção específica Alabarce:
    # O layout é homologado, sem conversão, e pode ser identificado por
    # de/para/rastreabilidade de CNPJ. Quando o único apontamento é
    # RASTREABILIDADE_LAYOUT, não deve bloquear o Modelo Robô KOF.
    texto = " ".join([
        str(row.get("layout_usado", "") or ""),
        str(row.get("alerta_extracao", "") or ""),
        str(row.get("alerta_identidade", "") or ""),
        str(row.get("motivo_identidade", "") or ""),
    ]).upper()
    if _is_layout_dedicado_sem_conversao_texto(texto) and "RASTREABILIDADE_LAYOUT" in _upper_sem_acento_local(texto):
        bloqueadores_reais = [
            "SKU_EAN_NAO_LOCALIZADO",
            "SKU_VAZIO",
            "QTD_VAZIA",
            "QTD_NAO_NUMERICA",
            "QTD_ZERO_OU_NEGATIVA",
            "MATRICULA_NAO_LOCALIZADA",
            "MATRICULA_NAO_ENCONTRADA",
            "PEDIDO_VAZIO",
            "PEDIDO_AUSENTE",
            "CNPJ_COM_TAMANHO_INVALIDO",
            "CHAVE_CLIENTE_NAO_LOCALIZADA",
            "LAYOUT_EM_HOMOLOGACAO",
            "STATUS_EXTRACAO_REQUER_VALIDACAO",
        ]
        if not any(tok in texto for tok in bloqueadores_reais):
            return False

    return True


def validar_linhas_negocio(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    validas = []
    descartadas = []
    for _, row in df.iterrows():
        motivo = ""
        if _status_identidade_bloqueia_fila(row):
            motivo = str(row.get("motivo_identidade", "") or "IDENTIDADE_PENDENTE_VALIDACAO")
        elif not str(row.get("sku_lido", "")).strip():
            motivo = "SKU_VAZIO"
        elif not str(row.get("quantidade_lida", "")).strip():
            motivo = "QTD_VAZIA"
        elif not str(row.get("matricula_final", "")).strip():
            motivo = "MATRICULA_NAO_ENCONTRADA"
        elif not str(row.get("numero_pedido_final", "")).strip():
            motivo = "PEDIDO_VAZIO"
        elif not str(row.get("data_remessa_final", "")).strip():
            motivo = "DATA_INVALIDA"
        elif _status_conversao_bloqueia_fila(row):
            motivo = "CONVERSAO_PENDENTE_VALIDACAO"

        data = row.to_dict()
        data["motivo_descarte"] = motivo
        if motivo:
            descartadas.append(data)
        else:
            validas.append(data)

    return pd.DataFrame(validas), pd.DataFrame(descartadas)


def _coluna_preferencial(df: pd.DataFrame, nomes: list[str], fallback: str = "") -> pd.Series:
    serie = pd.Series([fallback] * len(df), index=df.index, dtype=str)
    for nome in nomes:
        if nome in df.columns:
            candidato = df[nome].fillna("").astype(str).str.strip()
            serie = serie.where(serie.astype(str).str.strip() != "", candidato)
    return serie


def montar_dataframe_fila(df_validas: pd.DataFrame) -> pd.DataFrame:
    if df_validas.empty:
        return pd.DataFrame(columns=config.FILA_COLUMNS)
    qtd_final = _coluna_preferencial(df_validas, ["qtd_final", "qtd_convertida", "quantidade_lida"])
    df = pd.DataFrame({
        "Matricula": df_validas["matricula_final"].astype(str),
        "Sku": df_validas["sku_lido"].astype(str),
        "Qtd": qtd_final.astype(str),
        "Nº Pedido": df_validas["numero_pedido_final"].astype(str),
        "Data remessa": df_validas["data_remessa_final"].astype(str),
    })
    return df[config.FILA_COLUMNS]


def padronizar_pedidos(df_intermediario: pd.DataFrame, layout_config: Dict[str, str], data_base: datetime | None = None) -> Dict[str, object]:
    if df_intermediario is None or df_intermediario.empty:
        return {
            "sucesso": False,
            "mensagem": "Nenhum dado intermediário para padronizar",
            "df_final": pd.DataFrame(columns=config.FILA_COLUMNS),
            "qtd_linhas_entrada": 0,
            "qtd_linhas_validas": 0,
            "qtd_linhas_descartadas": 0,
            "alertas": ["Sem linhas válidas"],
            "df_descartadas": pd.DataFrame(),
            "df_validas": pd.DataFrame(),
            "resumo_conversao": {
                "layout_possui_conversao": False,
                "itens_convertidos": 0,
                "itens_nao_convertidos": 0,
                "itens_validar_conversao": 0,
                "itens_sem_conversao": 0,
            },
        }

    df = normalizar_campos_basicos(df_intermediario)
    df = validar_identidade_pedidos(df, layout_config)
    df = resolver_matricula_final(df, layout_config)
    df, resumo_conversao = aplicar_regras_conversao(df, layout_config)
    df = resolver_numero_pedido(df, layout_config)
    df = resolver_data_remessa(df, layout_config, data_base)
    df = _liberar_identidade_alabarce_sem_conversao(df)
    df_validas, df_descartadas = validar_linhas_negocio(df)
    df_final = montar_dataframe_fila(df_validas)

    alertas: List[str] = []
    alertas.extend(resumo_conversao.get("alertas_conversao", []))
    if not df_descartadas.empty:
        contagem = df_descartadas["motivo_descarte"].value_counts().to_dict()
        alertas.extend([f"{motivo}: {qtd}" for motivo, qtd in contagem.items()])
        if "MATRICULA_NAO_ENCONTRADA" in contagem:
            cnpjs = []
            if "cnpj_lido" in df_descartadas.columns:
                mask = df_descartadas["motivo_descarte"].astype(str) == "MATRICULA_NAO_ENCONTRADA"
                cnpjs = sorted({str(v).strip() for v in df_descartadas.loc[mask, "cnpj_lido"].tolist() if str(v).strip()})
            terminal_log.warning(
                "[PADRONIZADOR] Matricula nao encontrada | linhas=%s | cnpjs=%s",
                contagem.get("MATRICULA_NAO_ENCONTRADA", 0),
                cnpjs[:20],
            )

    sucesso = not df_final.empty
    mensagem = "Padronização concluída" if sucesso else "Nenhuma linha válida após padronização"
    return {
        "sucesso": sucesso,
        "mensagem": mensagem,
        "df_final": df_final,
        "qtd_linhas_entrada": len(df_intermediario),
        "qtd_linhas_validas": len(df_final),
        "qtd_linhas_descartadas": len(df_descartadas),
        "alertas": alertas,
        "df_descartadas": df_descartadas,
        "df_validas": df_validas,
        "resumo_conversao": resumo_conversao,
        "layout_possui_conversao": resumo_conversao.get("layout_possui_conversao", False),
        "qtd_itens_convertidos": resumo_conversao.get("itens_convertidos", 0),
        "qtd_itens_nao_convertidos": resumo_conversao.get("itens_nao_convertidos", 0),
        "qtd_itens_validar_conversao": resumo_conversao.get("itens_validar_conversao", 0),
        "qtd_itens_sem_conversao": resumo_conversao.get("itens_sem_conversao", 0),
    }
