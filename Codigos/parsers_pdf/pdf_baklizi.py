from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Optional

from pdf_alert_utils import linha_item_com_alerta
from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text_detailed, only_digits

# -----------------------------------------------------------------------------
# BAKLIZI / BAZKILI — TOTVS Varejo Supermercados / RELPEDSUPRIM
# -----------------------------------------------------------------------------
# Regra preservada da automacao original do Kaue Melo e ajustada ao Robô KOF:
# - PDF textual TOTVS, podendo conter vários pedidos no mesmo arquivo;
# - cada página repete o cabeçalho do pedido, mas as linhas de item podem continuar
#   em páginas posteriores;
# - o CNPJ correto é o CNPJ da loja/faturamento BAKLIZI, nunca o CNPJ do fornecedor
#   SPAL/Coca-Cola FEMSA;
# - SKU oficial é a coluna Cod Forn. O REF da descrição é apenas auditoria/fallback
#   quando o Cod Forn não veio na primeira coluna;
# - EAN vem nas linhas "EANs:" subsequentes ao item;
# - quantidade final é a coluna Qtde, não Qtd.Canc.;
# - sem conversão unidade->caixaria. Status: OK SEM CONVERSÃO.

RE_PEDIDO = re.compile(r"PEDIDO\s+DE\s+COMPRAS\s+(?P<pedido>\d{3,8}\s*/\s*[A-Z])", re.I)
RE_CNPJ_MASK = re.compile(r"(?P<base>\d{2}\.\d{3}\.\d{3}/\d{4})\s*-?\s*(?P<dv>\d{2})", re.I)
RE_ITEM = re.compile(
    r"^\s*(?P<cod1>\d{3,6})\s+(?:(?P<cod2>\d{1,6})\s+)?(?P<desc>.+?)\s+"
    r"(?P<emb>CX|FD|UN|UND|PC|PT)\s+(?P<emb_qtd>\d{1,3})\s+"
    r"(?P<qtd>\d{1,9}(?:\.\d{3})*,\d{2})\s+"
    r"(?P<valor_unit>\d{1,9}(?:\.\d{3})*,\d{4})\b",
    re.I,
)
RE_REF = re.compile(r"REF:\s*(\d{3,6})", re.I)
RE_EAN = re.compile(r"EANs?:\s*([0-9,\s]+)", re.I)
RE_TOTAL = re.compile(r"^TOTAIS\s+(?P<qtd>\d{1,9}(?:\.\d{3})*,\d{2})\b", re.I)

FORNECEDOR_PREFIXOS_CNPJ = ("61186888",)
BAKLIZI_PREFIXOS_CNPJ = ("00610350",)

BAKLIZI_CNPJ_MATRICULA = {
    "00610350000250": "7140054659",
    "00610350000170": "7140053557",
    "00610350000331": "7140053322",
    "00610350000501": "7140053589",
    "00610350000684": "7140053588",
    "00610350001141": "7140036471",
    "00610350001060": "7140036470",
    "00610350000765": "7140034189",
    "00610350001818": "7140037015",
    "00610350001737": "7140037165",
    "00610350001494": "7140290631",
    "00610350001656": "7140112118",
    "00610350001575": "7140037016",
    "00610350002032": "7140121490",
    "00610350002202": "7140343869",
}

SKIP_PREFIXES = (
    "PEDIDO DE COMPRAS",
    "FORNECEDOR",
    "R. SOCIAL",
    "ENDEREÇO",
    "ENDERECO",
    "BAIRRO",
    "CIDADE",
    "CEP",
    "ENDEREÇO PARA ENTREGA",
    "ENDERECO PARA ENTREGA",
    "ENDEREÇO PARA COBRANÇA",
    "ENDERECO PARA COBRANCA",
    "TRANSPORTADOR",
    "COD FORN",
    "A RECEBER",
    "(UNT.)",
    "EANS:",
    "TOTAIS",
    "DADOS ADICIONAIS",
    "VALOR TOTAL DO PEDIDO",
    "PRAZO PARA PAGAMENTO",
    "DESCONTO FINANCEIRO",
    "TAXA DE VENDOR",
    "TAXA DE COMPROR",
    "DATA DA EMISSÃO",
    "DATA DA EMISSAO",
    "PREVISÃO DE ENTREGA",
    "PREVISAO DE ENTREGA",
    "PESO TOTAL PEDIDO",
    "VOLUME TOTAL PEDIDO",
    "DATA LIMITE PARA ENTREGA",
    "CONDIÇÃO DO FRETE",
    "CONDICAO DO FRETE",
    "NRO PEDIDO NO FORNECEDOR",
    "NRO REQUISIÇÃO",
    "NRO REQUISICAO",
    "JUSTIFICATIVA",
    "OBSERVAÇÕES",
    "OBSERVACOES",
    "TOTVS VAREJO",
    "SUPERMERCADO BAKLIZI",
    "ITAQUI,",
    "SAO BORJA,",
    "SÃO BORJA,",
    "QUARAI,",
    "BARRA DO QUARAI,",
    "ALEGRETE,",
)


def _norm_line(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def _parse_decimal_br(value: str) -> Decimal:
    text = str(value or "").strip().replace(".", "").replace(",", ".")
    if not text:
        return Decimal("0")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _parse_qtd(value: str) -> str:
    number = _parse_decimal_br(value)
    if number <= 0:
        return ""
    # O PDF trabalha com caixas/embalagens. Quando vier inteiro com ,00, manter inteiro.
    if number == number.to_integral_value():
        return str(int(number))
    return format(number.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP).normalize(), "f").replace(".", ",")


def _pedido_limpo(value: str) -> str:
    return clean_text(value).replace(" ", "").upper()


def _cnpj_digits(value: str) -> str:
    d = only_digits(value)
    if not d:
        return ""
    if len(d) > 14:
        d = d[-14:]
    return d.zfill(14)


def _is_fornecedor_spal(cnpj: str) -> bool:
    c = _cnpj_digits(cnpj)
    return any(c.startswith(prefix) for prefix in FORNECEDOR_PREFIXOS_CNPJ)


def _extract_cnpj_cliente(text: str) -> str:
    """Extrai o CNPJ da loja Baklizi, ignorando o CNPJ do fornecedor SPAL."""
    encontrados: List[str] = []
    for match in RE_CNPJ_MASK.finditer(text or ""):
        cnpj = _cnpj_digits(f"{match.group('base')}-{match.group('dv')}")
        if cnpj:
            encontrados.append(cnpj)

    for cnpj in encontrados:
        if any(cnpj.startswith(prefix) for prefix in BAKLIZI_PREFIXOS_CNPJ):
            return cnpj

    for cnpj in encontrados:
        if not _is_fornecedor_spal(cnpj):
            return cnpj

    return ""


def _matricula(cnpj: str) -> str:
    return BAKLIZI_CNPJ_MATRICULA.get(_cnpj_digits(cnpj), "")


def _first_ean(contexto: List[str]) -> str:
    for linha in contexto:
        match = RE_EAN.search(linha)
        if not match:
            continue
        nums = [only_digits(x) for x in re.split(r"[,;\s]+", match.group(1))]
        nums = [x for x in nums if len(x) >= 8]
        return nums[0] if nums else ""
    return ""


def _ref_from_context(linha: str, contexto: List[str]) -> str:
    texto = " ".join([linha] + list(contexto))
    match = RE_REF.search(texto)
    return only_digits(match.group(1)) if match else ""


def _next_context(lines: List[str], idx: int, max_ahead: int = 8) -> List[str]:
    contexto: List[str] = []
    for j in range(idx + 1, min(len(lines), idx + 1 + max_ahead)):
        linha = lines[j]
        if RE_PEDIDO.search(linha) or RE_ITEM.match(linha) or linha.upper().startswith("TOTAIS"):
            break
        contexto.append(linha)
        if linha.upper().startswith("EANS:"):
            break
    return contexto


def _descricao_completa(match: re.Match, contexto: List[str]) -> str:
    partes = [_norm_line(match.group("desc"))]
    for linha in contexto:
        if linha.upper().startswith("EANS:"):
            break
        partes.append(linha)
    return _norm_line(" ".join(partes))


def _is_ignorable_line(line: str) -> bool:
    up = _norm_line(line).upper()
    if not up:
        return True
    return any(up.startswith(prefix) for prefix in SKIP_PREFIXES)


def _row_item(
    *,
    caminho_arquivo: str,
    layout_config: dict,
    pagina_pdf: int,
    linha_origem: int,
    linha_bruta: str,
    pedido: str,
    cnpj: str,
    sku: str,
    qtd: str,
    seq: str,
    ean: str,
    descricao: str,
    alerta: str,
) -> Dict[str, str]:
    status = "OK" if sku and qtd and pedido and cnpj else "ALERTA"
    alerta_final = alerta
    if not pedido:
        alerta_final = " | ".join([a for a in [alerta_final, "PEDIDO_NAO_LOCALIZADO"] if a])
    if not cnpj:
        alerta_final = " | ".join([a for a in [alerta_final, "CNPJ_LOJA_NAO_LOCALIZADO"] if a])
    if cnpj and _is_fornecedor_spal(cnpj):
        alerta_final = " | ".join([a for a in [alerta_final, "CNPJ_FORNECEDOR_SPAL_IGNORADO_VALIDAR_CNPJ_LOJA"] if a])
    if not sku:
        alerta_final = " | ".join([a for a in [alerta_final, "SKU_COD_FORN_AUSENTE_VALIDAR_MANUALMENTE"] if a])
    if not qtd:
        alerta_final = " | ".join([a for a in [alerta_final, "QTD_NAO_LOCALIZADA"] if a])

    return {
        "matricula_lida": _matricula(cnpj),
        "cnpj_lido": _cnpj_digits(cnpj),
        "sku_lido": sku,
        "codigo_sku_lido": sku,
        "ean_lido": ean,
        "codigo_origem_lido": seq,
        "cod_forn_lido": sku,
        "referencia_lida": sku,
        "descricao_lida": descricao,
        "quantidade_lida": qtd,
        "numero_pedido_lido": _pedido_limpo(pedido),
        "data_entrega_lida": "",
        "pagina_pdf": str(pagina_pdf),
        "linha_origem": str(linha_origem),
        "linha_bruta": linha_bruta,
        "origem_extracao": "PDF_TEXT_BAKLIZI_TOTVS",
        "motor_extracao": "pdfplumber/fitz/ocr",
        "status_extracao": status,
        "alerta_extracao": alerta_final,
        "qtd_original": qtd,
        "tipo_qtd_original": "CAIXARIA",
        "fator_conversao": "1",
        "qtd_convertida": qtd,
        "qtd_final": qtd,
        "status_conversao": "OK SEM CONVERSÃO" if qtd else "VALIDAR CONVERSÃO",
        "tipo_regra_conversao": "SEM_CONVERSAO",
        "regra_aplicada_conversao": "BAKLIZI_QTD_JA_EM_CAIXARIA_COLUNA_QTDE",
        "origem_regra_conversao": "PARSER_BAKLIZI_DEDICADO",
        "observacao_conversao": "Baklizi/Bazkili: quantidade lida diretamente da coluna Qtde do PDF TOTVS; Qtd.Canc. ignorada; sem conversão unidade-caixaria.",
        "layout_referencia": "BAKLIZI PDF",
        "arquivo_origem": caminho_arquivo,
    }


def ler_pdf_baklizi(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    extracao_pdf = extract_pages_text_detailed(caminho_arquivo)
    paginas = extracao_pdf.paginas

    linhas_saida: List[Dict[str, str]] = []
    alertas: List[str] = list(extracao_pdf.alertas)
    pedidos_encontrados: set[str] = set()
    cnpjs_encontrados: set[str] = set()
    totais_pdf: Dict[str, Decimal] = {}
    totais_lidos: Dict[str, Decimal] = {}
    itens_por_pedido: Dict[str, int] = {}
    sem_sku_por_pedido: Dict[str, int] = {}
    itens_vistos: set[tuple[str, str, str, str, str]] = set()
    linhas_lidas = 0

    for pagina_idx, texto in enumerate(paginas, start=1):
        pedido_match = RE_PEDIDO.search(texto or "")
        pedido_atual = _pedido_limpo(pedido_match.group("pedido")) if pedido_match else ""
        cnpj_atual = _extract_cnpj_cliente(texto or "")
        if pedido_atual:
            pedidos_encontrados.add(pedido_atual)
        if cnpj_atual:
            cnpjs_encontrados.add(cnpj_atual)

        linhas = [_norm_line(linha) for linha in (texto or "").splitlines() if _norm_line(linha)]
        for linha_idx, linha in enumerate(linhas, start=1):
            linhas_lidas += 1
            if _is_ignorable_line(linha):
                total_match = RE_TOTAL.match(linha)
                if total_match and pedido_atual:
                    totais_pdf[pedido_atual] = _parse_decimal_br(total_match.group("qtd"))
                continue

            item_match = RE_ITEM.match(linha)
            if not item_match:
                continue

            contexto = _next_context(linhas, linha_idx - 1)
            cod1 = only_digits(item_match.group("cod1"))
            cod2 = only_digits(item_match.group("cod2") or "")
            ref = _ref_from_context(linha, contexto)
            seq = cod2 or cod1
            alerta_linha = ""

            # Regra crítica: Cod Forn é o SKU oficial. REF nunca deve substituir o Cod Forn
            # quando a primeira coluna possui SKU válido. REF fica apenas como auditoria.
            if len(cod1) >= 4:
                sku = cod1
                if ref and ref != sku:
                    alerta_linha = f"REF_DIFERENTE_DO_COD_FORN_MANTIDO_COMO_AUDITORIA: ref={ref}"
            elif ref:
                sku = ref
                alerta_linha = f"SKU_RECUPERADO_POR_REF_COD_FORN_AUSENTE: cod_inicial={cod1}"
            else:
                sku = ""
                alerta_linha = f"PENDENTE_SKU_COD_FORN_AUSENTE: cod_inicial={cod1}"

            qtd = _parse_qtd(item_match.group("qtd"))
            ean = _first_ean(contexto)
            descricao = _descricao_completa(item_match, contexto)

            chave = (pedido_atual, cnpj_atual, sku, qtd, str(pagina_idx))
            if chave in itens_vistos:
                alerta_linha = " | ".join([a for a in [alerta_linha, "DUPLICIDADE_POTENCIAL_MANTIDA_PARA_VALIDACAO"] if a])
            else:
                itens_vistos.add(chave)

            row = _row_item(
                caminho_arquivo=caminho_arquivo,
                layout_config=layout_config,
                pagina_pdf=pagina_idx,
                linha_origem=linha_idx,
                linha_bruta=linha,
                pedido=pedido_atual,
                cnpj=cnpj_atual,
                sku=sku,
                qtd=qtd,
                seq=seq,
                ean=ean,
                descricao=descricao,
                alerta=alerta_linha,
            )
            linhas_saida.append(row)
            if pedido_atual:
                totais_lidos[pedido_atual] = totais_lidos.get(pedido_atual, Decimal("0")) + _parse_decimal_br(qtd)
                itens_por_pedido[pedido_atual] = itens_por_pedido.get(pedido_atual, 0) + 1
                if not sku:
                    sem_sku_por_pedido[pedido_atual] = sem_sku_por_pedido.get(pedido_atual, 0) + 1

    if not pedidos_encontrados:
        alertas.append("BAKLIZI_NENHUM_PEDIDO_IDENTIFICADO: esperado cabeçalho PEDIDO DE COMPRAS 00000/L ou /C.")
    if not cnpjs_encontrados:
        alertas.append("BAKLIZI_NENHUM_CNPJ_CLIENTE_IDENTIFICADO: validar extração do cabeçalho Dados para Faturamento.")
    if not linhas_saida:
        alertas.append("BAKLIZI_NENHUMA_LINHA_ITEM_EXTRAIDA: verificar se o PDF é imagem/escaneado ou layout diferente.")

    for pedido, qtd_pdf in sorted(totais_pdf.items()):
        qtd_lida = totais_lidos.get(pedido, Decimal("0"))
        if qtd_pdf != qtd_lida:
            alertas.append(
                f"BAKLIZI_TOTAL_DIVERGENTE | pedido={pedido} | qtd_pdf={int(qtd_pdf)} | qtd_lida={int(qtd_lida)} | conferir Excel de validação."
            )
        elif sem_sku_por_pedido.get(pedido, 0):
            alertas.append(
                f"BAKLIZI_TOTAL_OK_COM_PENDENCIA_SKU | pedido={pedido} | qtd_pdf={int(qtd_pdf)} | linhas_sem_sku={sem_sku_por_pedido[pedido]} | item mantido somente para validação."
            )
        else:
            alertas.append(
                f"BAKLIZI_TOTAL_OK | pedido={pedido} | itens={itens_por_pedido.get(pedido, 0)} | qtd={int(qtd_lida)}."
            )

    for cnpj in sorted(cnpjs_encontrados):
        if not _matricula(cnpj):
            alertas.append(f"BAKLIZI_CNPJ_SEM_MATRICULA | cnpj={cnpj} | manter no Excel/Cadastrar CNPJ como A CADASTRAR.")

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", "BAKLIZI PDF"),
    )

    sucesso = not df_intermediario.empty
    mensagem = (
        f"Leitura PDF BAKLIZI concluída: {len(df_intermediario)} linha(s), {len(pedidos_encontrados)} pedido(s), {len(cnpjs_encontrados)} CNPJ(s)."
        if sucesso
        else "Nenhuma linha válida foi extraída do PDF BAKLIZI."
    )

    return {
        "sucesso": sucesso,
        "mensagem": mensagem,
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "qtd_itens_extraidos": len(df_intermediario),
        "paginas_pdf_total": extracao_pdf.total_paginas,
        "paginas_pdf_processadas": extracao_pdf.paginas_processadas,
        "qtd_itens_ignorados": 0,
        "alertas": sorted({str(a) for a in alertas if str(a).strip()}),
        "resumo_pedidos": {
            pedido: {
                "itens": itens_por_pedido.get(pedido, 0),
                "qtd_lida": str(int(totais_lidos.get(pedido, Decimal("0")))),
                "qtd_pdf": str(int(totais_pdf.get(pedido, Decimal("0")))) if pedido in totais_pdf else "",
                "linhas_sem_sku": sem_sku_por_pedido.get(pedido, 0),
            }
            for pedido in sorted(pedidos_encontrados)
        },
    }
