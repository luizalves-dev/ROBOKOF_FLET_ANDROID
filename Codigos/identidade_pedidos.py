from __future__ import annotations

import re
import unicodedata
from decimal import Decimal, InvalidOperation
from typing import Any, Dict

import pandas as pd

from terminal_logger import get_terminal_logger

terminal_log = get_terminal_logger("identidade")

STATUS_OK = "OK"
STATUS_VALIDAR = "VALIDAR IDENTIDADE"
STATUS_ALERTA = "ALERTA IDENTIDADE"


def only_digits(value: Any) -> str:
    return re.sub(r"\D+", "", str(value or ""))


def _txt(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()




def _upper_sem_acento(value: Any) -> str:
    text = _txt(value).upper()
    return unicodedata.normalize("NFKD", text).encode("ASCII", "ignore").decode("ASCII")


def _layout_homologado_sem_conversao_nome(texto: str) -> bool:
    t = _upper_sem_acento(texto)
    return any(nome in t for nome in ("ALABARCE", "MONACO", "PRIMATO", "DAHER", "SEMPRE VALE", "BAKLIZI", "BAZKILI", "MABY", "BOZZA", "BOZA", "ESTRELA", "COOPERCICA", "REDE VIP", "IQUEGAMI", "REDE IQUEGAMI"))


def _is_contexto_iquegami(row: pd.Series | None, layout_config: Dict[str, Any]) -> bool:
    """Identifica com segurança linhas/layouts da Rede Iquegami.

    Usado apenas para liberar o marcador técnico de pedido "." quando a própria
    rede não envia número de pedido. Não altera a regra de outras redes.
    """
    partes = [
        _txt(layout_config.get("nome_layout", "")),
        _txt(layout_config.get("layout", "")),
        _txt(layout_config.get("observacoes", "")),
    ]
    if row is not None:
        partes.extend([
            _txt(row.get("layout_usado", "")),
            _txt(row.get("layout_referencia", "")),
            _txt(row.get("origem_extracao", "")),
            _txt(row.get("origem_regra_conversao", "")),
            _txt(row.get("regra_aplicada_conversao", "")),
        ])
    texto = _upper_sem_acento(" ".join(partes))
    return "IQUEGAMI" in texto

def _parse_numero_brasil(value: Any) -> Decimal | None:
    text = _txt(value)
    if not text:
        return None
    if re.search(r"[A-Za-z$%]", text):
        return None
    text = text.replace("R$", "").replace(" ", "")
    # 1.320 deve ser milhar em pedidos; 1.320,50 deve virar 1320.50.
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    elif re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
        text = text.replace(".", "")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _format_decimal(value: Decimal | None) -> str:
    if value is None:
        return ""
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f").replace(".", ",")


def _parece_preco_em_vez_de_quantidade(raw: Any) -> bool:
    """Detecta QTD capturada de coluna de custo/preço.

    A regra é intencionalmente conservadora: só alerta quando o valor tem vírgula
    com 3+ casas decimais e magnitude pequena, padrão comum em custo unitário
    como 2,4733 / 6,0917 / 12,5267. Isso evita o erro do Superlar se repetir em
    outros layouts com colunas próximas de quantidade e preço.
    """
    text = _txt(raw).replace(" ", "")
    if re.fullmatch(r"\d{1,4},\d{3,6}", text):
        val = _parse_numero_brasil(text)
        return val is not None and Decimal("0") < val < Decimal("10000")
    return False


def _status_inseguro(status: str) -> bool:
    s = _txt(status).upper()
    if not s:
        return False
    return any(tok in s for tok in ["ALERTA", "ERRO", "PENDENTE", "VALIDAR", "HOMOLOG"])


def _layout_em_homologacao(layout_config: Dict[str, Any]) -> bool:
    nome = _upper_sem_acento(layout_config.get("nome_layout", ""))
    # Layouts dedicados e homologados dentro do Robô KOF não podem ser tratados
    # como TESTE/HOMOLOGAÇÃO apenas por terem sido identificados por rastreabilidade.
    if _layout_homologado_sem_conversao_nome(nome):
        return False
    return any(tok in nome for tok in ["HOMOLOG", "RASTREABILIDADE", "TESTE ", "TESTE_"])



def _is_layout_homologado_sem_conversao(row: pd.Series, layout_config: Dict[str, Any]) -> bool:
    texto = " ".join([
        _txt(layout_config.get("nome_layout", "")),
        _txt(row.get("layout_usado", "")),
        _txt(row.get("layout_referencia", "")),
        _txt(row.get("regra_aplicada_conversao", "")),
        _txt(row.get("origem_regra_conversao", "")),
    ])
    return _layout_homologado_sem_conversao_nome(texto)


def _is_layout_alabarce(row: pd.Series, layout_config: Dict[str, Any]) -> bool:
    # Mantido por compatibilidade com validações existentes: agora representa
    # layouts dedicados sem conversão que aceitam GTIN/EAN como rastreio.
    return _is_layout_homologado_sem_conversao(row, layout_config)


def _linha_layout_apta_para_liberar(row: pd.Series, layout_config: Dict[str, Any], alerta: str = "", motivo: str = "") -> bool:
    """Liberação defensiva para layouts dedicados sem conversão.

    Alabarce, Monaco/Mônaco e Primato podem entrar por rastreabilidade/de-para.
    Quando parser dedicado já trouxe CNPJ, SKU/EAN, quantidade e pedido,
    RASTREABILIDADE_LAYOUT é auditoria e não deve esvaziar o Modelo Robô KOF.
    Erros reais continuam bloqueando.
    """
    texto = " ".join([
        _txt(layout_config.get("nome_layout", "")),
        _txt(row.get("layout_usado", "")),
        _txt(row.get("layout_referencia", "")),
        _txt(row.get("alerta_extracao", "")),
        _txt(alerta),
        _txt(motivo),
    ])
    if not _layout_homologado_sem_conversao_nome(texto):
        return False

    bloqueadores_reais = [
        "SKU_EAN_NAO_LOCALIZADO", "SKU_VAZIO", "QTD_VAZIA", "QTD_NAO_NUMERICA",
        "QTD_ZERO_OU_NEGATIVA", "PEDIDO_VAZIO", "PEDIDO_AUSENTE",
        "CNPJ_COM_TAMANHO_INVALIDO", "CHAVE_CLIENTE_NAO_LOCALIZADA",
        "LAYOUT_INVALIDO", "LAYOUT INVALIDO", "ERRO", "QTD_PARECE_PRECO_CUSTO",
        "TOTAL_DIVERGENTE",
    ]
    texto_norm = _upper_sem_acento(texto)
    if any(tok in texto_norm for tok in bloqueadores_reais):
        return False

    # Falta de matrícula não impede a leitura/validação; a etapa de fila decide
    # se a linha entra no modelo ou vai para Cadastrar CNPJ.
    cnpj = only_digits(row.get("cnpj_lido", "") or row.get("cnpj_base_lido", ""))
    sku = only_digits(row.get("sku_lido", "") or row.get("codigo_sku_lido", ""))
    ean = only_digits(row.get("ean_lido", ""))
    raw_pedido = _txt(row.get("numero_pedido_lido", ""))
    pedido = only_digits(raw_pedido)
    pedido_ponto_iquegami = raw_pedido == "." and _is_contexto_iquegami(row, layout_config)
    qtd_num = _parse_numero_brasil(row.get("quantidade_lida", ""))

    layout_maby = "MABY" in _upper_sem_acento(" ".join([_txt(layout_config.get("nome_layout", "")), _txt(row.get("layout_usado", "")), _txt(row.get("layout_referencia", ""))]))
    if len(cnpj) != 14 and not (layout_maby and len(cnpj) == 15 and cnpj.startswith("01695774")):
        return False
    if not (sku or ean):
        return False
    if not pedido_ponto_iquegami and (not pedido or len(pedido) < 3):
        return False
    if qtd_num is None or qtd_num <= 0:
        return False
    return True


def _linha_alabarce_apta_para_liberar(row: pd.Series, layout_config: Dict[str, Any], alerta: str = "", motivo: str = "") -> bool:
    return _linha_layout_apta_para_liberar(row, layout_config, alerta=alerta, motivo=motivo)


def _alerta_auditoria_nao_bloqueante_layout(alerta: str) -> bool:
    """Alertas que devem aparecer no Excel, mas não podem bloquear o Modelo.

    Baklizi/Bazkili usa Cod Forn como SKU oficial. Portanto REF divergente é
    auditoria, não erro. Quando o parser recupera SKU por REF porque o Cod Forn
    veio deslocado/vazio, a linha pode ir ao modelo desde que o SKU final esteja
    preenchido. RASTREABILIDADE_LAYOUT também é só informativo em layouts
    dedicados homologados.
    """
    texto = _upper_sem_acento(alerta)
    if not texto.strip():
        return True
    partes = [p.strip() for p in re.split(r"\s*\|\s*", texto) if p.strip()]
    seguros = (
        "RASTREABILIDADE_LAYOUT",
        "REF_DIFERENTE_DO_COD_FORN_MANTIDO_COMO_AUDITORIA",
        "SKU_RECUPERADO_POR_REF_COD_FORN_AUSENTE",
    )
    bloqueantes = (
        "PENDENTE_SKU",
        "SKU_COD_FORN_AUSENTE_VALIDAR_MANUALMENTE",
        "SKU_EAN_NAO_LOCALIZADO",
        "SKU_VAZIO",
        "QTD_VAZIA",
        "QTD_NAO_NUMERICA",
        "ERRO",
        "INVALID",
        "INVALIDO",
        "CONVERSAO_PENDENTE",
        "VALIDAR_CONVERSAO",
    )
    if any(b in texto for b in bloqueantes):
        return False
    return all(any(s in p for s in seguros) for p in partes)



def _validar_linha(row: pd.Series, layout_config: Dict[str, Any]) -> tuple[str, str, str, dict[str, str]]:
    layout_nome = _txt(layout_config.get("nome_layout", ""))
    tipo_cliente = _txt(layout_config.get("tipo_cliente_destino", "")).upper()
    observacoes: list[str] = []
    updates: dict[str, str] = {}

    matricula = only_digits(row.get("matricula_lida", ""))
    cnpj = only_digits(row.get("cnpj_lido", ""))
    gln = only_digits(row.get("gln_lido", ""))
    cod_cliente = _txt(row.get("codigo_cliente_lido", "") or row.get("codigo_loja_lido", "") or row.get("cod_cliente_lido", ""))
    sku = only_digits(row.get("sku_lido", "") or row.get("codigo_sku_lido", ""))
    ean = only_digits(row.get("ean_lido", ""))
    raw_pedido = _txt(row.get("numero_pedido_lido", ""))
    pedido = only_digits(raw_pedido)
    pedido_ponto_iquegami = raw_pedido in {"", "."} and _is_contexto_iquegami(row, layout_config)
    qtd_raw = _txt(row.get("quantidade_lida", ""))
    status_extracao = _txt(row.get("status_extracao", ""))
    alerta_extracao = _txt(row.get("alerta_extracao", ""))

    # Quando o parser colocou EAN na coluna SKU, preserva o EAN e deixa a
    # conversão/mapa resolver o SKU real. Não perde o código original.
    if not ean and len(sku) in {8, 12, 13, 14}:
        updates["ean_lido"] = sku
        if len(sku) >= 12:
            observacoes.append("SKU_LIDO_PARECE_EAN: preservado também em EAN para conversão/rastreio.")

    # Pedido: evita usar CNPJ, EAN, SKU ou quantidade como número de pedido.
    if pedido in {"", "0"}:
        obs_layout = _txt(layout_config.get("observacoes", "")).upper()
        pedido_opcional_excel = (
            str(layout_config.get("tipo_arquivo", "")).upper() == "EXCEL"
            and raw_pedido == "."
            and any(flag in obs_layout for flag in ["PEDIDO_OPCIONAL", "SEM PEDIDO POR LINHA", "PEDIDO OPCIONAL"])
        )
        if pedido_opcional_excel or pedido_ponto_iquegami:
            # Rede Iquegami: quando a origem vem sem número de pedido, o ponto
            # é o marcador técnico oficial. Isso libera a identidade sem afetar
            # outras redes e mantém o pedido rastreável no Excel/Fila.
            updates["numero_pedido_lido"] = "."
        elif str(layout_config.get("tipo_arquivo", "")).upper() == "EXCEL" and raw_pedido == ".":
            observacoes.append("PEDIDO_AUSENTE_EM_EXCEL: layout Excel sem pedido por linha; validar agrupamento antes de TXT/fila.")
        else:
            observacoes.append("PEDIDO_VAZIO_OU_NAO_LOCALIZADO")
    else:
        if len(pedido) < 3 or len(pedido) > 20:
            observacoes.append(f"PEDIDO_COM_TAMANHO_ATIPICO: {pedido}")
        if cnpj and pedido == cnpj:
            observacoes.append("PEDIDO_PARECE_CNPJ: validar posição do número do pedido.")
        if ean and pedido == ean:
            observacoes.append("PEDIDO_PARECE_EAN: validar posição do número do pedido.")
        if sku and pedido == sku and len(sku) >= 5:
            observacoes.append("PEDIDO_PARECE_SKU: validar posição do número do pedido.")

    # Cliente: o pedido não deve sumir se faltar matrícula, mas a identidade
    # do cliente precisa ficar rastreável para Cadastrar CNPJ/de-para.
    if tipo_cliente == "MATRICULA":
        if not matricula:
            observacoes.append("MATRICULA_NAO_LOCALIZADA_NO_LAYOUT")
        elif len(matricula) < 6 or len(matricula) > 12:
            observacoes.append(f"MATRICULA_COM_TAMANHO_ATIPICO: {matricula}")
    else:
        if cnpj:
            layout_maby = "MABY" in _upper_sem_acento(" ".join([layout_nome, _txt(row.get("layout_usado", "")), _txt(row.get("layout_referencia", ""))]))
            if len(cnpj) != 14 and not (layout_maby and len(cnpj) == 15 and cnpj.startswith("01695774")):
                observacoes.append(f"CNPJ_COM_TAMANHO_INVALIDO: {cnpj}")
        elif not (gln or cod_cliente or matricula):
            observacoes.append("CHAVE_CLIENTE_NAO_LOCALIZADA: sem CNPJ/GLN/código loja/matrícula.")

    # Produto: SKU, EAN ou código origem devem estar presentes. Conversão pode
    # preencher SKU via EAN depois, por isso EAN válido não bloqueia aqui.
    codigo_origem = only_digits(row.get("codigo_origem_lido", ""))
    is_alabarce = _is_layout_alabarce(row, layout_config)
    if not (sku or ean or codigo_origem):
        observacoes.append("SKU_EAN_NAO_LOCALIZADO")
    if not is_alabarce:
        if sku and len(sku) in {13, 14} and not ean:
            observacoes.append("SKU_COM_TAMANHO_DE_EAN: conferir coluna SKU/EAN.")
        if ean and len(ean) not in {8, 12, 13, 14}:
            observacoes.append(f"EAN_COM_TAMANHO_ATIPICO: {ean}")
    else:
        # Alabarce/Bluesoft é SKU/Ref + quantidade por loja em caixaria.
        # O GTIN Unitário pode ter 8, 11, 12, 13 ou 14 dígitos e é apenas rastreio;
        # não deve bloquear o Modelo Robô KOF nem impedir que todas as quantidades entrem.
        pass

    # Quantidade: não deixa preço/custo virar quantidade final silenciosamente.
    qtd_num = _parse_numero_brasil(qtd_raw)
    if qtd_raw == "":
        observacoes.append("QTD_VAZIA")
    elif qtd_num is None:
        observacoes.append(f"QTD_NAO_NUMERICA: {qtd_raw}")
    elif qtd_num <= 0:
        observacoes.append(f"QTD_ZERO_OU_NEGATIVA: {qtd_raw}")
    else:
        updates["quantidade_lida"] = _format_decimal(qtd_num)
        if not is_alabarce and _parece_preco_em_vez_de_quantidade(qtd_raw):
            observacoes.append(f"QTD_PARECE_PRECO_CUSTO: {qtd_raw}. Validar coluna de quantidade do layout.")

    layout_usado_linha = _txt(row.get("layout_usado", ""))
    texto_layout_dedicado = " ".join([layout_nome, layout_usado_linha])
    alerta_up = _upper_sem_acento(alerta_extracao)
    rastreabilidade_alabarce_ok = (
        _layout_homologado_sem_conversao_nome(texto_layout_dedicado)
        and "RASTREABILIDADE_LAYOUT" in alerta_up
        and not any(tok in alerta_up for tok in ["ERRO", "INVALID", "INVALIDO", "NAO HOMOLOG", "HOMOLOGACAO CONTROLADA"])
    )

    alerta_apenas_auditoria = (
        _layout_homologado_sem_conversao_nome(texto_layout_dedicado)
        and _alerta_auditoria_nao_bloqueante_layout(alerta_extracao)
    )
    if _status_inseguro(status_extracao):
        # Para Baklizi/Bazkili, ALERTA só pode bloquear quando existe pendência
        # real. Alertas de auditoria/rastreabilidade não devem zerar o modelo.
        if not alerta_apenas_auditoria:
            observacoes.append(f"STATUS_EXTRACAO_REQUER_VALIDACAO: {status_extracao}")
    if alerta_extracao:
        # Layouts dedicados homologados podem carregar alertas informativos,
        # como RASTREABILIDADE_LAYOUT, REF divergente e SKU recuperado por REF.
        # Eles ficam registrados no Excel, mas não bloqueiam o Modelo Robô KOF
        # quando SKU/QTD/Pedido/CNPJ estão completos.
        if not (rastreabilidade_alabarce_ok or alerta_apenas_auditoria):
            observacoes.append(alerta_extracao)
    if _layout_em_homologacao(layout_config):
        observacoes.append(f"LAYOUT_EM_HOMOLOGACAO_OU_TESTE: {layout_nome}. Conferir identidade antes de fila/TXT.")

    # Trava extra: se o único motivo restante for a rastreabilidade informativa
    # da Alabarce, libera a identidade como OK.
    if rastreabilidade_alabarce_ok:
        observacoes = [
            obs for obs in observacoes
            if "RASTREABILIDADE_LAYOUT" not in str(obs).upper()
            and "RASTREABILIDADE" not in str(obs).upper()
        ]

    observacoes = [obs for obs in dict.fromkeys(observacoes) if obs]
    if observacoes:
        return STATUS_VALIDAR, " | ".join(observacoes), observacoes[0].split(":", 1)[0], updates
    return STATUS_OK, "OK", "", updates


def validar_identidade_pedidos(df: pd.DataFrame | None, layout_config: Dict[str, Any]) -> pd.DataFrame:
    """Aplica uma camada corporativa de identidade antes da padronização.

    Não exclui linhas. Apenas marca riscos em status_identidade/alerta_identidade
    para aparecerem no Excel e impedir envio à fila quando a identidade estiver
    insegura. Isso protege todos os layouts, inclusive novos/homologação.
    """
    if df is None or df.empty:
        return df if df is not None else pd.DataFrame()

    work = df.copy().fillna("")
    for col in ["status_identidade", "motivo_identidade", "alerta_identidade"]:
        if col not in work.columns:
            work[col] = ""

    total_alertas = 0
    for idx, row in work.iterrows():
        status, alerta, motivo, updates = _validar_linha(row, layout_config)
        for col, value in updates.items():
            if col not in work.columns:
                work[col] = ""
            work.at[idx, col] = value

        # Defesa final: Alabarce homologada, sem conversão e com campos-chave
        # completos não deve ficar bloqueada por alerta informativo de rastreabilidade.
        row_corrigida = work.loc[idx]
        if _linha_layout_apta_para_liberar(row_corrigida, layout_config, alerta=alerta, motivo=motivo):
            status, motivo, alerta = STATUS_OK, "OK", ""

        work.at[idx, "status_identidade"] = status
        work.at[idx, "motivo_identidade"] = motivo
        work.at[idx, "alerta_identidade"] = alerta
        if status != STATUS_OK:
            total_alertas += 1

    terminal_log.info(
        "[IDENTIDADE] layout=%s | linhas=%s | pendentes_validacao=%s",
        layout_config.get("nome_layout", ""),
        len(work),
        total_alertas,
    )
    return work
