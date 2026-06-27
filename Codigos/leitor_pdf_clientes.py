import unicodedata

from parsers_pdf.pdf_tischler import ler_pdf_tischler
from parsers_pdf.pdf_alabarce import ler_pdf_alabarce
from parsers_pdf.pdf_barracao import ler_pdf_barracao
from parsers_pdf.pdf_baklizi import ler_pdf_baklizi
from parsers_pdf.pdf_coopercica import ler_pdf_coopercica
from parsers_pdf.pdf_droga_clara import ler_pdf_droga_clara
from parsers_pdf.pdf_festival import ler_pdf_festival
from parsers_pdf.pdf_galassi import ler_pdf_galassi
from parsers_pdf.pdf_miller import ler_pdf_miller
from parsers_pdf.pdf_rede_italo import ler_pdf_rede_italo
from parsers_pdf.pdf_estrela import ler_pdf_estrela
from parsers_pdf.pdf_coasul import ler_pdf_coasul
from parsers_pdf.pdf_sempre_vale import ler_pdf_sempre_vale
from parsers_pdf.pdf_generico_homologacao import ler_pdf_generico_homologacao

# Importacoes opcionais adicionadas nas evolucoes recentes.
# Regra corporativa: o Robô KOF nao pode deixar de abrir caso um patch tenha sido
# aplicado sem copiar algum parser novo. Nesses casos, o layout cai em
# homologacao controlada, registra alerta no Excel/log e preserva os itens para
# validacao, em vez de quebrar a tela na inicializacao.
try:
    from parsers_pdf.pdf_emop_parteka import ler_pdf_emop_parteka
    _ERRO_IMPORT_EMOP_PARTEKA = None
except Exception as _exc_emop_parteka:
    _ERRO_IMPORT_EMOP_PARTEKA = _exc_emop_parteka

    def ler_pdf_emop_parteka(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
        resultado = ler_pdf_generico_homologacao(
            caminho_arquivo,
            layout_config,
            mapeamentos_df,
            referencia="EMOP_PARTEKA_HOMOLOGACAO_CONTROLADA_SEM_PARSER_DEDICADO",
        )
        alertas = list(resultado.get("alertas", []) or [])
        alertas.append(
            "PARSER_EMOP_PARTEKA_NAO_CARREGADO: arquivo Codigos/parsers_pdf/pdf_emop_parteka.py ausente ou com erro. "
            "O Robô KOF manteve a leitura em homologação controlada; conferir Excel antes de TXT/fila. "
            f"Detalhe técnico: {_ERRO_IMPORT_EMOP_PARTEKA}"
        )
        resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
        return resultado

try:
    from parsers_pdf.pdf_super_ubialli import ler_pdf_super_ubialli
    _ERRO_IMPORT_SUPER_UBIALLI = None
except Exception as _exc_super_ubialli:
    _ERRO_IMPORT_SUPER_UBIALLI = _exc_super_ubialli

    def ler_pdf_super_ubialli(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
        resultado = ler_pdf_generico_homologacao(
            caminho_arquivo,
            layout_config,
            mapeamentos_df,
            referencia="SUPER_UBIALLI_HOMOLOGACAO_CONTROLADA_SEM_PARSER_DEDICADO",
        )
        alertas = list(resultado.get("alertas", []) or [])
        alertas.append(
            "PARSER_SUPER_UBIALLI_NAO_CARREGADO: arquivo Codigos/parsers_pdf/pdf_super_ubialli.py ausente ou com erro. "
            "O Robô KOF manteve a leitura em homologação controlada; conferir Excel antes de TXT/fila. "
            f"Detalhe técnico: {_ERRO_IMPORT_SUPER_UBIALLI}"
        )
        resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
        return resultado
from parsers_pdf.pdf_maby import ler_pdf_maby
from parsers_pdf.pdf_supermais import ler_pdf_supermais
from parsers_pdf.pdf_panelao import ler_pdf_panelao
from parsers_pdf.pdf_caita import ler_pdf_caita
from parsers_pdf.pdf_layouts_adicionais import (
    ler_pdf_bozza,
    ler_pdf_daher,
    ler_pdf_indiana,
    ler_pdf_kacula,
    ler_pdf_monaco,
    ler_pdf_primato,
    ler_pdf_superlar,
)

from layout_standard import normalize_intermediate_columns
from parsers_pdf.pdf_utils import extract_pages_text_detailed
from terminal_logger import get_terminal_logger
from pdf_alert_utils import alertas_para_dataframe


terminal_log = get_terminal_logger("leitor_pdf")


def _normalizar_nome_layout_pdf(valor: object) -> str:
    texto = str(valor or "").strip().upper()
    texto = unicodedata.normalize("NFKD", texto).encode("ASCII", "ignore").decode("ASCII")
    return texto



def _finalizar_pdf(resultado: dict, layout_config: dict, caminho_arquivo: str = "") -> dict:
    nome_layout = layout_config.get("nome_layout", "")
    df_intermediario = resultado.get("df_intermediario")
    if df_intermediario is not None:
        resultado["df_intermediario"] = normalize_intermediate_columns(
            df_intermediario,
            arquivo_origem=caminho_arquivo,
            layout_usado=nome_layout,
        )
        resultado["qtd_itens_extraidos"] = len(resultado["df_intermediario"])
        resultado["qtd_linhas_lidas"] = max(
            int(resultado.get("qtd_linhas_lidas", 0) or 0),
            len(resultado["df_intermediario"]),
        )

    auditoria_existente = resultado.get("df_auditoria_paginas")
    auditoria_vazia = auditoria_existente is None or bool(getattr(auditoria_existente, "empty", False))
    if caminho_arquivo and auditoria_vazia:
        try:
            audit = extract_pages_text_detailed(caminho_arquivo)
            df_audit = audit.auditoria_df()
            resultado["df_auditoria_paginas"] = df_audit
            resultado["paginas_pdf_total"] = audit.total_paginas
            resultado["paginas_pdf_processadas"] = audit.paginas_processadas
            resultado["paginas_pdf_sem_texto"] = int((df_audit["caracteres"] == 0).sum()) if not df_audit.empty else 0
            resultado["motores_pdf"] = (
                ", ".join(f"{motor}:{qtd}" for motor, qtd in df_audit["motor"].value_counts().to_dict().items())
                if not df_audit.empty and "motor" in df_audit.columns
                else ""
            )
            alertas = list(resultado.get("alertas", []) or [])
            alertas.extend(audit.alertas)
            resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
        except Exception as exc:
            terminal_log.exception("[PDF] Falha ao montar auditoria detalhada do PDF: %s", caminho_arquivo)
            alertas = list(resultado.get("alertas", []) or [])
            alertas.append(f"AUDITORIA_PDF_FALHOU: {exc}")
            resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})

    # Registra alertas de parser em DataFrame para aparecerem no Excel de validacao,
    # sem transformar linha nao reconhecida em pedido falso.
    df_alertas_parser = alertas_para_dataframe(
        resultado.get("alertas", []) or [],
        caminho_arquivo,
        nome_layout,
    )
    if df_alertas_parser is not None and not df_alertas_parser.empty:
        resultado["df_itens_ignorados"] = df_alertas_parser
        resultado["qtd_itens_ignorados"] = max(
            int(resultado.get("qtd_itens_ignorados", 0) or 0),
            len(df_alertas_parser),
        )

    sem_texto = int(resultado.get("paginas_pdf_sem_texto", 0) or 0)
    total_paginas = int(resultado.get("paginas_pdf_total", 0) or 0)
    itens_extraidos = int(resultado.get("qtd_itens_extraidos", 0) or 0)
    if total_paginas and sem_texto == total_paginas and itens_extraidos == 0:
        resultado["sucesso"] = False
        resultado["mensagem"] = (
            "PDF sem texto extraivel. O arquivo parece ser imagem/escaneado. "
            "Ative/instale OCR para leitura automatica ou envie PDF editavel."
        )
        alertas = list(resultado.get("alertas", []) or [])
        alertas.append("PDF_IMAGEM_OU_SEM_TEXTO_EXTRAIVEL")
        resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})

    terminal_log.info(
        "[PDF] Resultado leitura | layout=%s | sucesso=%s | linhas=%s | itens=%s | paginas=%s/%s | alertas=%s | ignorados=%s",
        nome_layout,
        resultado.get("sucesso"),
        resultado.get("qtd_linhas_lidas", 0),
        resultado.get("qtd_itens_extraidos", resultado.get("qtd_linhas_lidas", 0)),
        resultado.get("paginas_pdf_processadas", ""),
        resultado.get("paginas_pdf_total", ""),
        len(resultado.get("alertas", []) or []),
        resultado.get("qtd_itens_ignorados", 0),
    )
    return resultado




def _finalizar_generico_homologacao(caminho_arquivo: str, layout_config: dict, referencia: str = "RASTREABILIDADE") -> dict:
    return _finalizar_pdf(
        ler_pdf_generico_homologacao(caminho_arquivo, layout_config, None, referencia=referencia),
        layout_config,
        caminho_arquivo,
    )


def _ler_max_center_com_fallback(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    """Max Center foi cadastrado como candidato semelhante ao Superlar.

    Primeiro tenta o parser Superlar. Se não extrair item, cai para o parser genérico
    de homologação, sempre com alerta de conferência manual.
    """
    resultado = ler_pdf_superlar(caminho_arquivo, layout_config, mapeamentos_df)
    df = resultado.get("df_intermediario")
    if resultado.get("sucesso") and df is not None and not getattr(df, "empty", True):
        alertas = list(resultado.get("alertas", []) or [])
        alertas.append("MAX_CENTER_ALIAS_SUPERLAR: leitura feita com parser Superlar por similaridade. Conferir Excel de validação.")
        resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
        return resultado
    terminal_log.warning("[PDF] Max Center sem itens no parser Superlar; usando parser genérico de homologação.")
    return ler_pdf_generico_homologacao(caminho_arquivo, layout_config, mapeamentos_df, referencia="SUPERLAR")

def ler_pdf_cliente(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    try:
        nome_layout = _normalizar_nome_layout_pdf(layout_config.get("nome_layout", ""))
        terminal_log.info("[PDF] Arquivo recebido: %s", caminho_arquivo)
        terminal_log.info("[PDF] Layout selecionado: %s", layout_config.get("nome_layout", ""))

        if "TISCHLER" in nome_layout:
            return _finalizar_pdf(ler_pdf_tischler(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "ALABARCE" in nome_layout:
            return _finalizar_pdf(ler_pdf_alabarce(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "BAKLIZI" in nome_layout:
            return _finalizar_pdf(ler_pdf_baklizi(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "BARRACAO" in nome_layout or "BARRAC" in nome_layout:
            return _finalizar_pdf(ler_pdf_barracao(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "MILLER" in nome_layout:
            return _finalizar_pdf(ler_pdf_miller(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "COOPERCICA" in nome_layout:
            return _finalizar_pdf(ler_pdf_coopercica(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "FESTIVAL" in nome_layout:
            return _finalizar_pdf(ler_pdf_festival(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "GALASSI" in nome_layout:
            return _finalizar_pdf(ler_pdf_galassi(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "DROGA CLARA" in nome_layout:
            return _finalizar_pdf(ler_pdf_droga_clara(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "REDE ITALO" in nome_layout:
            return _finalizar_pdf(ler_pdf_rede_italo(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "ESTRELA" in nome_layout:
            return _finalizar_pdf(ler_pdf_estrela(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "COASUL" in nome_layout:
            return _finalizar_pdf(ler_pdf_coasul(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "SUPER UBIALLI" in nome_layout or "UBIALLI" in nome_layout:
            return _finalizar_pdf(ler_pdf_super_ubialli(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if any(chave in nome_layout for chave in ["TESTE REDE ALFA", "TESTE PONTO CERTO", "REDE CELEIRO"]):
            return _finalizar_generico_homologacao(caminho_arquivo, layout_config, referencia=f"{nome_layout}_CONVERSAO_HOMOLOGACAO")

        if "SEMPRE VALE" in nome_layout:
            return _finalizar_pdf(ler_pdf_sempre_vale(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)


        if "SUPERMAIS" in nome_layout or "SUPER MAIS" in nome_layout:
            return _finalizar_pdf(ler_pdf_supermais(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "PANELAO" in nome_layout or "PANELÃO" in nome_layout:
            return _finalizar_pdf(ler_pdf_panelao(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "CAITA" in nome_layout or "CAITÁ" in nome_layout:
            return _finalizar_pdf(ler_pdf_caita(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "RASTREABILIDADE" in nome_layout and ("HOMOLOGACAO" in nome_layout or "GENERICA" in nome_layout):
            return _finalizar_generico_homologacao(caminho_arquivo, layout_config, referencia="RASTREABILIDADE_GENERICA_MANUAL")

        if "PASSARELA CENTER" in nome_layout or "PASSARELA" in nome_layout:
            return _finalizar_generico_homologacao(caminho_arquivo, layout_config, referencia="PASSARELA_CENTER_HOMOLOGACAO_OCR")

        if "MAX CENTER" in nome_layout or "MAXCENTER" in nome_layout:
            return _finalizar_pdf(_ler_max_center_com_fallback(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "MABY" in nome_layout:
            return _finalizar_pdf(ler_pdf_maby(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "E.M.O.P" in nome_layout or "EMOP" in nome_layout or "PARTEKA" in nome_layout:
            return _finalizar_pdf(ler_pdf_emop_parteka(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "BELTRAME" in nome_layout:
            return _finalizar_generico_homologacao(caminho_arquivo, layout_config, referencia="BELTRAME_EAN_SKU_SEM_CONVERSAO")

        if "BOZZA" in nome_layout:
            return _finalizar_pdf(ler_pdf_bozza(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "MONACO" in nome_layout:
            return _finalizar_pdf(ler_pdf_monaco(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "DAHER" in nome_layout:
            return _finalizar_pdf(ler_pdf_daher(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "PRIMATO" in nome_layout:
            return _finalizar_pdf(ler_pdf_primato(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "SUPERLAR" in nome_layout:
            return _finalizar_pdf(ler_pdf_superlar(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "INDIANA" in nome_layout:
            return _finalizar_pdf(ler_pdf_indiana(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        if "KACULA" in nome_layout:
            return _finalizar_pdf(ler_pdf_kacula(caminho_arquivo, layout_config, mapeamentos_df), layout_config, caminho_arquivo)

        terminal_log.error("[PDF] Leitura nao implementada para o layout: %s", layout_config.get("nome_layout", ""))
        return {
            "sucesso": False,
            "mensagem": f"Leitura de PDF nao implementada para o layout: {layout_config.get('nome_layout', '')}",
            "df_intermediario": None,
            "qtd_linhas_lidas": 0,
            "alertas": [],
        }

    except Exception as e:
        terminal_log.exception("[PDF] Erro ao ler arquivo PDF: %s", caminho_arquivo)
        return {
            "sucesso": False,
            "mensagem": str(e),
            "df_intermediario": None,
            "qtd_linhas_lidas": 0,
            "alertas": [str(e)],
        }
