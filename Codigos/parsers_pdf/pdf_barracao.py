from __future__ import annotations

import re
from typing import Dict, List, Tuple

from pdf_alert_utils import linha_item_com_alerta
from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text, only_digits
from terminal_logger import get_terminal_logger


# Barracão / RP One
# Regra de negócio preservada:
# - CNPJ da SPAL/fornecedor 61.186.888 deve ser ignorado;
# - CNPJ válido é o da loja no bloco Empresa;
# - SKU oficial é Cod Forn/Cod Form;
# - Quantidade final é Qtde Emb, pois o layout já vem em caixaria;
# - linhas com Cod Forn ausente ou quantidade zerada seguem para validação/alerta,
#   sem desaparecer e sem alimentar fila/TXT antes da validação manual.

SUPPLIER_CNPJ_PREFIX = "61186888"
SUPPLIER_CNPJ_FULL = "61186888013929"

RE_PEDIDO_NUMERO = re.compile(r"N[úu]mero\s+do\s+Pedido:\s*(\d+)", re.I)
RE_PEDIDO_FALLBACK = re.compile(r"\bPedido:\s*(\d+)", re.I)
RE_CNPJ = re.compile(r"CNPJ:\s*([\d./-]+)", re.I)
RE_EMPRESA_CNPJ = re.compile(r"Empresa:\s*.*?CNPJ:\s*([\d./-]+)", re.I | re.S)
RE_DT_ENTREGA = re.compile(r"Dt\.\s*Entrega:\s*(\d{2}/\d{2}/\d{2,4})", re.I)
terminal_log = get_terminal_logger("pdf_barracao")


RE_LINHA_ITEM = re.compile(
    r"^\s*(?P<codigo>\d+)\s+"
    r"(?P<ean>\d{8,14})"
    r"(?:\s+(?P<sku>\d{4,6}))?\s+"
    r"(?P<descricao_marca>.+?)\s+"
    r"(?P<quant>\d{1,6}(?:\.\d{3})*,\d{3})\s+"
    r"(?P<qtd_emb>\d+)\s+"
    r"(?P<emb>[A-Z]{1,5}/\d+)\s+",
    re.I,
)

RE_CANDIDATO_ITEM = re.compile(r"^\s*\d+\s+\d{8,14}\b")


def _is_supplier_cnpj(cnpj: str) -> bool:
    digitos = only_digits(cnpj)
    return digitos == SUPPLIER_CNPJ_FULL or digitos.startswith(SUPPLIER_CNPJ_PREFIX)


def _extract_header(page_text: str) -> Tuple[str, str, str]:
    """Extrai pedido, CNPJ da loja e data de entrega da página.

    Em páginas de continuação o PDF traz novamente apenas o CNPJ do fornecedor.
    Nesses casos retornamos CNPJ vazio para preservar o CNPJ da loja da página anterior.
    """
    pedido = ""
    cnpj_loja = ""
    data_entrega = ""

    match_pedido = RE_PEDIDO_NUMERO.search(page_text) or RE_PEDIDO_FALLBACK.search(page_text)
    if match_pedido:
        pedido = clean_text(match_pedido.group(1))

    match_data = RE_DT_ENTREGA.search(page_text)
    if match_data:
        data_entrega = clean_text(match_data.group(1))

    header_text = page_text.split("Código", 1)[0] if "Código" in page_text else page_text[:2500]

    match_empresa = RE_EMPRESA_CNPJ.search(header_text)
    if match_empresa:
        candidato = only_digits(match_empresa.group(1))
        if candidato and not _is_supplier_cnpj(candidato):
            cnpj_loja = candidato

    if not cnpj_loja:
        cnpjs = [only_digits(cnpj) for cnpj in RE_CNPJ.findall(header_text)]
        cnpjs_loja = [cnpj for cnpj in cnpjs if cnpj and not _is_supplier_cnpj(cnpj)]
        if cnpjs_loja:
            cnpj_loja = cnpjs_loja[-1]

    return pedido, cnpj_loja, data_entrega


def _qtd_int(qtd: str) -> int | None:
    qtd_limpa = only_digits(qtd)
    if not qtd_limpa:
        return None
    try:
        return int(qtd_limpa)
    except Exception:
        return None


def _descricao_limpa(match_item: re.Match[str]) -> str:
    return clean_text(match_item.group("descricao_marca") or "")


def _montar_linha_alerta(
    *,
    caminho_arquivo: str,
    layout_config: dict,
    page_idx: int,
    linha: str,
    alerta: str,
    current_cnpj: str,
    current_pedido: str,
    current_data_entrega: str,
    match_item: re.Match[str] | None,
) -> Dict[str, str]:
    sku = ""
    ean = ""
    codigo_origem = ""
    qtd = ""
    descricao = ""
    if match_item:
        sku = only_digits(match_item.group("sku") or "")
        ean = only_digits(match_item.group("ean") or "")
        codigo_origem = only_digits(match_item.group("codigo") or "")
        qtd_int = _qtd_int(match_item.group("qtd_emb") or "")
        qtd = str(qtd_int) if qtd_int is not None else ""
        descricao = _descricao_limpa(match_item)

    return linha_item_com_alerta(
        caminho_arquivo=caminho_arquivo,
        layout_usado=layout_config.get("nome_layout", ""),
        pagina_pdf=page_idx,
        linha_bruta=linha,
        alerta=alerta,
        cnpj_lido=current_cnpj,
        sku_lido=sku,
        quantidade_lida=qtd,
        numero_pedido_lido=current_pedido,
        data_entrega_lida=current_data_entrega,
        descricao_lida=descricao,
        codigo_sku_lido=sku,
        ean_lido=ean,
        codigo_origem_lido=codigo_origem,
    )


def ler_pdf_barracao(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    paginas = extract_pages_text(caminho_arquivo)

    linhas_saida: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedidos_encontrados = {}
    cnpjs_encontrados = {}
    linhas_lidas = 0
    linhas_validas = 0
    linhas_pendentes = 0
    linhas_descartadas = 0
    current_pedido = ""
    current_cnpj = ""
    current_data_entrega = ""

    for page_idx, texto in enumerate(paginas, start=1):
        pedido, cnpj_loja, data_entrega = _extract_header(texto)
        if pedido:
            current_pedido = pedido
            pedidos_encontrados[pedido] = pedidos_encontrados.get(pedido, 0) + 1
        if cnpj_loja:
            current_cnpj = cnpj_loja
            cnpjs_encontrados[cnpj_loja] = cnpjs_encontrados.get(cnpj_loja, 0) + 1
        if data_entrega:
            current_data_entrega = data_entrega

        for linha in [clean_text(l) for l in texto.splitlines() if clean_text(l)]:
            if not RE_CANDIDATO_ITEM.match(linha):
                continue

            linhas_lidas += 1
            match_item = RE_LINHA_ITEM.search(linha)
            if not match_item:
                linhas_descartadas += 1
                alerta = f"Pagina {page_idx}: linha de item nao reconhecida | {linha[:180]}"
                alertas.append(alerta)
                linhas_saida.append(_montar_linha_alerta(
                    caminho_arquivo=caminho_arquivo,
                    layout_config=layout_config,
                    page_idx=page_idx,
                    linha=linha,
                    alerta=alerta,
                    current_cnpj=current_cnpj,
                    current_pedido=current_pedido,
                    current_data_entrega=current_data_entrega,
                    match_item=None,
                ))
                continue

            sku = only_digits(match_item.group("sku") or "")
            ean = only_digits(match_item.group("ean") or "")
            codigo_origem = only_digits(match_item.group("codigo") or "")
            qtd_int = _qtd_int(match_item.group("qtd_emb") or "")
            qtd = str(qtd_int) if qtd_int is not None else ""
            descricao = _descricao_limpa(match_item)

            motivos: list[str] = []
            if not current_pedido:
                motivos.append("pedido ausente")
            if not current_cnpj:
                motivos.append("CNPJ loja ausente")
            if not sku:
                motivos.append("Cod Forn/SKU ausente")
            if qtd_int is None:
                motivos.append("Qtde Emb nao identificada")
            elif qtd_int <= 0:
                motivos.append("Qtde Emb zerada")

            if motivos:
                linhas_pendentes += 1
                alerta = (
                    f"Pagina {page_idx}: item mantido para validacao por campo obrigatorio/operacional pendente "
                    f"({'; '.join(motivos)}) | pedido={current_pedido or '-'} | cnpj={current_cnpj or '-'} | "
                    f"sku={sku or '-'} | qtd={qtd or '-'}"
                )
                alertas.append(alerta)
                linhas_saida.append(_montar_linha_alerta(
                    caminho_arquivo=caminho_arquivo,
                    layout_config=layout_config,
                    page_idx=page_idx,
                    linha=linha,
                    alerta=alerta,
                    current_cnpj=current_cnpj,
                    current_pedido=current_pedido,
                    current_data_entrega=current_data_entrega,
                    match_item=match_item,
                ))
                continue

            linhas_validas += 1
            linhas_saida.append(
                {
                    "matricula_lida": "",
                    "cnpj_lido": current_cnpj,
                    "sku_lido": sku,
                    "codigo_sku_lido": sku,
                    "ean_lido": ean,
                    "codigo_origem_lido": codigo_origem,
                    "descricao_lida": descricao,
                    "pagina_pdf": str(page_idx),
                    "linha_bruta": linha,
                    "origem_extracao": "PDF_BARRACAO_RP_ONE",
                    "motor_extracao": "PDF_TEXTO",
                    "status_extracao": "OK",
                    "alerta_extracao": "",
                    "quantidade_lida": qtd,
                    "qtd_original": qtd,
                    "tipo_qtd_original": "CAIXARIA_QTDE_EMB",
                    "qtd_final": qtd,
                    "status_conversao": "OK SEM CONVERSÃO",
                    "regra_aplicada_conversao": "BARRACAO_QTDE_EMB_SEM_CONVERSAO",
                    "origem_regra_conversao": "PARSER_BARRACAO",
                    "numero_pedido_lido": current_pedido,
                    "data_entrega_lida": current_data_entrega,
                }
            )

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    terminal_log.info(
        "[BARRACAO] Leitura finalizada | pedidos=%s | cnpjs_loja=%s | linhas_lidas=%s | "
        "linhas_validas=%s | pendentes=%s | descartes=%s | linhas_intermediarias=%s | regra_data=%s",
        sorted(pedidos_encontrados.keys()),
        sorted(cnpjs_encontrados.keys()),
        linhas_lidas,
        linhas_validas,
        linhas_pendentes,
        linhas_descartadas,
        len(df_intermediario),
        layout_config.get("regra_data_entrega", "D+1"),
    )
    if not df_intermediario.empty:
        terminal_log.info("[BARRACAO] Amostra df_intermediario:\n%s", df_intermediario.head(10).to_string(index=False))

    if df_intermediario.empty:
        return {
            "sucesso": False,
            "mensagem": "Nenhuma linha valida foi extraida do PDF Barracao",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "qtd_itens_extraidos": 0,
            "qtd_itens_ignorados": linhas_descartadas,
            "alertas": sorted(set(alertas)),
        }

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Barracao concluida com {len(df_intermediario)} linha(s) para validacao",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "qtd_itens_extraidos": len(df_intermediario),
        "qtd_itens_ignorados": linhas_descartadas,
        "alertas": sorted(set(alertas)),
    }
