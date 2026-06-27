from pathlib import Path
import traceback
import re
from datetime import datetime

import config
from leitor_fila import listar_excels, ler_arquivo_entrada
from validacoes import normalize_and_validate
from gln_service import load_gln_map, only_digits
from gerador_robokof import gerar_arquivo_robokof
from gerador_erro import gerar_arquivo_erro
from gerador_txt_edi import gerar_txt_edi_do_excel
from terminal_logger import get_terminal_logger


terminal_log = get_terminal_logger("operacao")


def nome_arquivo_safe(s: str) -> str:
    s = str(s)
    return re.sub(r"[^0-9A-Za-z]+", "", s)


def normalizar_pedido_regra_especial(matricula, pedido) -> str:
    """
    Regra:
    se o pedido vier "." ou vazio,
    gera: matricula + data de hoje no formato DDMMAAAA

    Exemplo:
    matrícula = 7777777777
    hoje = 17/03/2026
    resultado = 777777777717032026
    """
    mat_s = str(matricula).strip()
    ped_s = str(pedido).strip()

    if ped_s == "." or ped_s == "":
        data_hoje = datetime.now().strftime("%d%m%Y")
        return f"{mat_s}{data_hoje}"

    return ped_s


def arquivo_excel_valido(path: Path) -> bool:
    nome = path.name
    if nome.startswith("~$"):
        return False
    if path.suffix.lower() not in [".xlsx", ".xls"]:
        return False
    return True


def obter_arquivos_da_fila():
    """
    Prioridade:
    1. arquivo oficial da fila (config.FILA_FILE_NAME)
    2. fallback para outros excels da pasta
    """
    arquivo_principal_fila = config.FILA_DIR / config.FILA_FILE_NAME

    if arquivo_principal_fila.exists() and arquivo_excel_valido(arquivo_principal_fila):
        return [arquivo_principal_fila]

    files = [f for f in listar_excels(config.FILA_DIR) if arquivo_excel_valido(f)]
    return files



def _status_erros_unicos(df_err) -> set[str]:
    """Extrai status de erro normalizados para decidir bloqueio operacional."""
    if df_err is None or getattr(df_err, "empty", True) or "Status Erro" not in df_err.columns:
        return set()
    status: set[str] = set()
    for raw in df_err["Status Erro"].fillna("").astype(str).tolist():
        for parte in re.split(r"\s*\|\s*", raw):
            parte = parte.strip().upper()
            if parte:
                status.add(parte)
    return status


def _df_erros_bloqueantes(df_err):
    """Retorna apenas erros que devem impedir TXT/fila.

    Duplicidades operacionais tratáveis não devem travar a geração. A consolidação de
    SKU duplicado ocorre em validacoes.py, mas esta função mantém compatibilidade com
    filas antigas que ainda tragam status de duplicidade.
    """
    if df_err is None or getattr(df_err, "empty", True) or "Status Erro" not in df_err.columns:
        return df_err.iloc[0:0].copy() if df_err is not None and hasattr(df_err, "iloc") else df_err

    nao_bloqueantes = {str(s).strip().upper() for s in getattr(config, "ROBOKOF_STATUS_ERRO_NAO_BLOQUEANTES", set())}

    def eh_bloqueante(raw) -> bool:
        partes = [p.strip().upper() for p in re.split(r"\s*\|\s*", str(raw or "")) if p.strip()]
        if not partes:
            return True
        return any(p not in nao_bloqueantes for p in partes)

    mask = df_err["Status Erro"].apply(eh_bloqueante)
    return df_err[mask].copy()


def _gerar_relatorio_validacao_fila(input_path: Path, df_err, *, tipo: str = "ERRO") -> Path | None:
    """Mantém o padrão de Excel único: não cria ERR separado para fila.

    Os erros/alertas devem permanecer no próprio Excel consolidado de validação
    gerado antes da fila/TXT. Esta função fica compatível com chamadas antigas,
    mas não grava arquivo adicional.
    """
    if df_err is None or getattr(df_err, "empty", True):
        return None
    terminal_log.warning(
        "[OPERACAO][EXCEL_UNICO] %s registro(s) de %s encontrados em %s; não será gerado Excel ERR separado.",
        len(df_err),
        tipo,
        input_path,
    )
    return None


def validar_liberacao_manual_para_txt(df_in, input_path: Path | None = None, df_ok=None, df_err=None):
    """Libera ou bloqueia a geração RoboKOF/TXT com segurança.

    Regra principal:
    - Se existir coluna de validação manual, ela continua obrigatória.
    - Se não existir, permite compatibilidade com o arquivo operacional legado
      Pedidos_RoboKOF.xlsx quando existem linhas válidas e não há erro técnico bloqueante.

    Isso evita o erro que bloqueava a geração mesmo com GLN/matrícula corretos na base,
    sem abrir mão da conferência contra linhas inválidas.
    """
    if not getattr(config, "ROBOKOF_EXIGIR_VALIDACAO_MANUAL", True):
        terminal_log.warning("[OPERACAO] Trava de validacao manual desabilitada por configuracao.")
        return

    colunas_existentes = [c for c in getattr(config, "ROBOKOF_COLUNAS_VALIDACAO_MANUAL", []) if c in df_in.columns]
    if colunas_existentes:
        coluna = colunas_existentes[0]
        valores_ok = getattr(config, "ROBOKOF_VALORES_VALIDACAO_MANUAL", {"SIM", "VALIDADO", "OK", "CONFERIDO"})
        status = df_in[coluna].fillna("").astype(str).str.strip().str.upper()
        pendentes = df_in[~status.isin(valores_ok)]
        if not pendentes.empty:
            raise ValueError(
                f"Geração de RoboKOF/TXT bloqueada: {len(pendentes)} linha(s) da fila não estão validadas manualmente "
                f"na coluna '{coluna}'. Use SIM/VALIDADO/OK/CONFERIDO somente após conferir o Excel de validação."
            )

        terminal_log.info("[OPERACAO] Validação manual confirmada para %s linha(s) da fila.", len(df_in))
        return

    # Compatibilidade com o fluxo operacional antigo: o usuário já separa o arquivo final
    # dentro de Resultados/Arquivos Fila/Pedidos_RoboKOF.xlsx. Nesse caso, sem coluna
    # de conferência, liberamos quando existem linhas válidas e os erros encontrados são apenas não bloqueantes/auditáveis.
    permitir_legado = getattr(config, "ROBOKOF_PERMITIR_FILA_LEGADA_SEM_COLUNA_VALIDACAO", True)
    arquivo_oficial = False
    if input_path is not None:
        try:
            arquivo_oficial = (
                Path(input_path).name.lower() == str(config.FILA_FILE_NAME).lower()
                and Path(input_path).resolve().parent == Path(config.FILA_DIR).resolve()
            )
        except Exception:
            arquivo_oficial = Path(input_path).name.lower() == str(config.FILA_FILE_NAME).lower()

    df_bloqueantes = _df_erros_bloqueantes(df_err)
    erros_tecnicos = 0 if df_err is None else len(df_err)
    erros_bloqueantes = 0 if df_bloqueantes is None else len(df_bloqueantes)
    linhas_validas = 0 if df_ok is None else len(df_ok)

    if permitir_legado and arquivo_oficial and erros_bloqueantes == 0 and linhas_validas > 0:
        terminal_log.warning(
            "[OPERACAO] Fila sem coluna de validação manual liberada em modo compatibilidade | "
            "arquivo=%s | linhas_validadas=%s | alertas_nao_bloqueantes=%s",
            input_path,
            linhas_validas,
            erros_tecnicos,
        )
        return

    if erros_bloqueantes > 0:
        status = sorted(_status_erros_unicos(df_bloqueantes))
        raise ValueError(
            f"Geração de RoboKOF/TXT bloqueada: a fila não possui coluna de conferência manual "
            f"e contém {erros_bloqueantes} linha(s) com erro técnico bloqueante. "
            f"Status encontrados: {status}. Corrija o Excel de validação antes de gerar TXT."
        )

    raise ValueError(
        "Geração de RoboKOF/TXT bloqueada: a fila não possui coluna de conferência manual. "
        "Inclua uma coluna 'Validado Manualmente' com valor SIM/VALIDADO após conferir o Excel de validação, "
        "ou use o arquivo oficial Resultados/Arquivos Fila/Pedidos_RoboKOF.xlsx com 100% das linhas tecnicamente válidas."
    )


def processar_arquivo(input_path: Path, gln_map: dict):
    terminal_log.info("[OPERACAO] Iniciando processamento da fila: %s", input_path)
    df_in = ler_arquivo_entrada(input_path, sheet_name=config.INPUT_SHEET_NAME)
    terminal_log.info("[OPERACAO] Linhas recebidas na fila: %s", len(df_in) if df_in is not None else 0)
    df_ok, df_err = normalize_and_validate(df_in)
    terminal_log.info(
        "[OPERACAO] Validacao da fila concluida | validas=%s | erros=%s | status_erros=%s",
        len(df_ok) if df_ok is not None else 0,
        len(df_err) if df_err is not None else 0,
        sorted(_status_erros_unicos(df_err)),
    )

    df_bloqueantes_pre = _df_erros_bloqueantes(df_err)
    if df_err is not None and not df_err.empty:
        # Gera sempre um diagnóstico/auditoria quando houver linhas fora do TXT.
        # Se forem somente status não bloqueantes, o TXT segue para as linhas válidas,
        # mas a pendência fica registrada para conferência.
        tipo_relatorio = "ERRO" if (df_bloqueantes_pre is not None and not df_bloqueantes_pre.empty) else "ALERTA"
        _gerar_relatorio_validacao_fila(input_path, df_err if tipo_relatorio == "ALERTA" else df_bloqueantes_pre, tipo=tipo_relatorio)

    validar_liberacao_manual_para_txt(df_in, input_path=input_path, df_ok=df_ok, df_err=df_err)

    resumo = {
        "robokof_gerados": 0,
        "txts_gerados": 0,
        "erros_gerados": 0,
    }

    if df_ok is not None and not df_ok.empty:
        df_ok = df_ok.copy()
        df_ok["Nº Pedido"] = df_ok.apply(
            lambda row: normalizar_pedido_regra_especial(row["Matricula"], row["Nº Pedido"]),
            axis=1
        )

    err_groups = {}
    if df_err is not None and not df_err.empty:
        if "Pedido_norm" in df_err.columns:
            for (mat, ped), g in df_err.groupby(["Matricula", "Pedido_norm"], dropna=False):
                err_groups[(str(mat).strip(), str(ped).strip())] = g.copy()
        else:
            for (mat, ped), g in df_err.groupby(["Matricula", "Nº Pedido"], dropna=False):
                err_groups[(str(mat).strip(), str(ped).strip())] = g.copy()

    if df_ok is None or df_ok.empty:
        return resumo

    # normaliza matrícula no ponto final também, para não falhar em GLN por causa de
    # formato Excel como 1700273623.0 ou notação científica.
    df_ok["Matricula"] = df_ok["Matricula"].apply(only_digits)

    for (mat, ped), g in df_ok.groupby(["Matricula", "Nº Pedido"], dropna=False):
        mat_s = only_digits(mat)
        ped_s = str(ped).strip()

        gln_val = gln_map.get(mat_s)
        if not gln_val:
            exemplos = list(gln_map.keys())[:5]
            terminal_log.error(
                "[OPERACAO] GLN nao encontrado para matricula: %s | base_carregada=%s matriculas | exemplos=%s",
                mat_s,
                len(gln_map),
                exemplos,
            )
            raise ValueError(
                f"GLN não encontrado na BASE de GLNS para a matrícula {mat_s}. "
                "A matrícula passou pela validação do pedido, mas não possui GLN na coluna configurada da BASE de GLNS. "
                "Confira a BASE de GLNS.xlsx: coluna B = matrícula e coluna A = GLN."
            )

        out_name = f"RoboKOF_{nome_arquivo_safe(mat_s)}_{nome_arquivo_safe(ped_s)}.xlsx"
        out_path = config.OUT_ROBOKOF_DIR / out_name

        rows = []
        for _, row in g.iterrows():
            rows.append({
                "Matricula": mat_s,
                "Sku": str(row["Sku"]),
                "Qtd": int(row["Qtd"]),
                "Pedido": ped_s,
                "Data": str(row["Data remessa"]),
            })

        gerar_arquivo_robokof(
            template_path=config.TEMPLATE_PATH,
            out_path=out_path,
            sheet_name=config.TEMPLATE_ORDEM_SHEET,
            header_candidates=config.HEADER_MAP,
            rows=rows,
            gln_value=gln_val,
            tipo_solicitacao_value=config.TIPO_SOLICITACAO_VALUE,
            forma_pagamento_value=config.FORMA_PAGAMENTO_VALUE,
        )
        resumo["robokof_gerados"] += 1
        terminal_log.info("[OPERACAO] Excel RoboKOF gerado: %s", out_path)

        txt_name = f"{nome_arquivo_safe(mat_s)}_{nome_arquivo_safe(ped_s)}.txt"
        txt_path = config.OUT_TXT_DIR / txt_name
        gerar_txt_edi_do_excel(out_path, txt_path)
        resumo["txts_gerados"] += 1
        terminal_log.info("[OPERACAO] TXT EDI gerado: %s", txt_path)

        gerr = err_groups.get((mat_s, ped_s))
        if gerr is not None and not gerr.empty:
            terminal_log.warning(
                "[OPERACAO][EXCEL_UNICO] %s erro(s)/alerta(s) do pedido %s/%s mantidos no Excel de validação; ERR separado não será gerado.",
                len(gerr),
                mat_s,
                ped_s,
            )

    return resumo


def processar_lista_arquivos(files, gln_map, progress_callback=None):
    resumo = {
        "robokof_gerados": 0,
        "txts_gerados": 0,
        "erros_gerados": 0,
    }

    total = len(files)

    for i, f in enumerate(files, start=1):
        try:
            if progress_callback:
                progress_callback(i - 1, total, f"Processando arquivo {i}/{total}: {f.name}")

            terminal_log.info("[OPERACAO] Processando arquivo %s/%s: %s", i, total, f.name)
            print(f"Processando: {f.name}")
            r = processar_arquivo(f, gln_map)

            if r:
                resumo["robokof_gerados"] += r.get("robokof_gerados", 0)
                resumo["txts_gerados"] += r.get("txts_gerados", 0)
                resumo["erros_gerados"] += r.get("erros_gerados", 0)

        except Exception as e:
            terminal_log.exception("[OPERACAO] Erro ao processar arquivo: %s", f)
            print(f"Erro ao processar {f.name}: {e}")
            traceback.print_exc()

    if progress_callback:
        progress_callback(total, total, "Processamento concluído.")

    return resumo


def main(progress_callback=None):
    terminal_log.info("[OPERACAO] Iniciando fluxo operacional RoboKOF.")
    config.OUT_ROBOKOF_DIR.mkdir(parents=True, exist_ok=True)
    config.OUT_ERRO_DIR.mkdir(parents=True, exist_ok=True)
    config.OUT_TXT_DIR.mkdir(parents=True, exist_ok=True)

    gln_map = load_gln_map(
        config.GLN_BASE_PATH,
        sheet_name=config.GLN_SHEET_NAME,
        col_gln=config.GLN_COL_GLN,
        col_matricula=config.GLN_COL_MATRICULA,
    )

    files = obter_arquivos_da_fila()
    terminal_log.info("[OPERACAO] Arquivos localizados na fila: %s", len(files))

    if not files:
        if progress_callback:
            progress_callback(0, 1, "Nenhum arquivo encontrado na fila.")
        terminal_log.warning("[OPERACAO] Nenhum arquivo encontrado em: %s", config.FILA_DIR)
        print(f"Nenhum arquivo encontrado em: {config.FILA_DIR}")
        return {
            "robokof_gerados": 0,
            "txts_gerados": 0,
            "erros_gerados": 0,
        }

    # tenta primeiro do jeito oficial
    try:
        return processar_lista_arquivos(files, gln_map, progress_callback=progress_callback)
    except Exception as e:
        terminal_log.exception("[OPERACAO] Falha ao processar arquivo principal da fila.")
        print(f"Falha ao processar arquivo principal da fila: {e}")
        traceback.print_exc()

    # fallback: tenta outros excels da pasta
    try:
        files_fallback = [
            f for f in listar_excels(config.FILA_DIR)
            if arquivo_excel_valido(f)
        ]

        if not files_fallback:
            raise RuntimeError("Nenhum arquivo alternativo encontrado para fallback.")

        return processar_lista_arquivos(files_fallback, gln_map, progress_callback=progress_callback)

    except Exception as e:
        terminal_log.exception("[OPERACAO] Falha geral no processamento da fila.")
        print(f"Falha geral no processamento da fila: {e}")
        traceback.print_exc()

        if progress_callback:
            progress_callback(1, 1, "Erro no processamento da fila.")

        return {
            "robokof_gerados": 0,
            "txts_gerados": 0,
            "erros_gerados": 0,
        }


if __name__ == "__main__":
    main()
