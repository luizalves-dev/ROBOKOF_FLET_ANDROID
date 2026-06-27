from __future__ import annotations

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
import re
from typing import Dict, List, Optional

from pdf_alert_utils import linha_item_com_alerta
from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text_detailed, only_digits


# ---------------------------------------------------------------------------
# SEMPRE VALE / TOTVS Varejo
# ---------------------------------------------------------------------------
# Regra preservada da automação original do Kauê Melo:
# - PDF textual com vários pedidos dentro do mesmo arquivo;
# - número do pedido no cabeçalho: "PEDIDO DE COMPRAS 1779217/M";
# - CNPJ da loja/faturamento no corpo do pedido;
# - SKU oficial = primeira coluna Cod Forn;
# - quantidade final em caixaria = campo de quantidade com 2 casas decimais;
# - quando uma linha possui quantidade, mas não possui SKU na coluna Cod Forn,
#   o item não pode sumir: ele entra no Excel como PENDENTE_VALIDACAO/Alertas.

RE_PEDIDO = re.compile(r"PEDIDO\s+DE\s+COMPRAS\s+([0-9]+\s*/\s*[A-Z])", re.I)
RE_CNPJ = re.compile(r"CNPJ\s+(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", re.I)
RE_ITEM = re.compile(
    r"^\s*(?P<sku>\d{4,6})\s+(?P<seq>\d{1,6})\s+.+?\s+"
    r"(?P<qtd>\d{1,9}(?:\.\d{3})*,\d{2})\s+(?P<valor>\d{1,9}(?:\.\d{3})*,\d{4})\b",
    re.I,
)
RE_ITEM_SEM_SKU = re.compile(
    r"^\s*(?P<referencia>\d{4,6})\s+.+?\s+"
    r"(?P<qtd>\d{1,9}(?:\.\d{3})*,\d{2})\s+(?P<valor>\d{1,9}(?:\.\d{3})*,\d{4})\b",
    re.I,
)
RE_TOTAL = re.compile(r"^TOTAIS\s+(?P<qtd>\d{1,9}(?:\.\d{3})*,\d{2})\b", re.I)

SKIP_PREFIXES = (
    "PEDIDO DE COMPRAS",
    "FORNECEDOR",
    "R. Social",
    "Endereço",
    "Bairro",
    "Cidade",
    "Cep",
    "Telefone",
    "Transportador",
    "Cod Forn",
    "a Receber",
    "(Unt.)",
    "OBSERVAÇÃO PADRÃO DO PEDIDO",
    "TOTVS Varejo",
    "DADOS ADICIONAIS",
    "Prazo para pagamento",
    "Desconto Financeiro",
    "Taxa de vendor",
    "Taxa de compror",
    "Data da emissão",
    "Previsão de entrega",
    "Peso total pedido",
    "Volume total pedido",
    "Condição do frete",
    "Nro pedido no fornecedor",
    "Nro Requisição",
    "LIMEIRA,",
    "PAULINIA,",
    "Recebimento",
    "LOJAS -",
    "CD -",
    "ENDEREÇO PARA ENTREGA",
    "ENDEREÇO PARA COBRANÇA",
)

SEMPRE_VALE_CNPJ_MATRICULA = {
    "62488937000105": "700251669",
    "62488937000962": "700308941",
    "62488937000296": "700318860",
    "62488937001934": "7110033583",
    "62488937002825": "7110197882",
    "62488937003120": "7110451189",
    "62488937003392": "7110484386",
}


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


def _parse_qtd_int(value: str) -> str:
    number = _parse_decimal_br(value)
    if number <= 0:
        return ""
    rounded = number.to_integral_value(rounding=ROUND_HALF_UP)
    return str(int(rounded))


def _pedido_limpo(value: str) -> str:
    return clean_text(value).replace(" ", "").upper()


def _cnpj_digits(value: str) -> str:
    return only_digits(value)


def _matricula(cnpj: str) -> str:
    return SEMPRE_VALE_CNPJ_MATRICULA.get(_cnpj_digits(cnpj), "")


def _row_ok(
    *,
    caminho_arquivo: str,
    layout_config: dict,
    pagina_pdf: int,
    linha_origem: int,
    linha_bruta: str,
    cnpj: str,
    pedido: str,
    sku: str,
    qtd: str,
    seq: str = "",
) -> Dict[str, str]:
    qtd_final = _parse_qtd_int(qtd)
    cnpj_final = _cnpj_digits(cnpj)
    return {
        "matricula_lida": _matricula(cnpj_final),
        "cnpj_lido": cnpj_final,
        "sku_lido": only_digits(sku),
        "codigo_sku_lido": only_digits(sku),
        "ean_lido": "",
        "codigo_origem_lido": only_digits(seq),
        "descricao_lida": linha_bruta[:180],
        "quantidade_lida": qtd_final,
        "numero_pedido_lido": pedido,
        "data_entrega_lida": "",
        "pagina_pdf": str(pagina_pdf),
        "linha_origem": str(linha_origem),
        "linha_bruta": linha_bruta,
        "origem_extracao": "PDF_TEXT_SEMPRE_VALE",
        "motor_extracao": "pdfplumber",
        "status_extracao": "OK",
        "alerta_extracao": "",
        "qtd_original": qtd_final,
        "tipo_qtd_original": "CAIXARIA",
        "qtd_final": qtd_final,
        "status_conversao": "OK SEM CONVERSÃO",
        "tipo_regra_conversao": "SEM_CONVERSAO",
        "regra_aplicada_conversao": "SEMPRE_VALE_QTD_JA_EM_CAIXARIA",
        "origem_regra_conversao": "PARSER_SEMPRE_VALE_DEDICADO",
        "observacao_conversao": "Sempre Vale: quantidade lida diretamente do PDF/TOTVS; sem conversão unidade-caixaria.",
        "arquivo_origem": "",
        "layout_usado": layout_config.get("nome_layout", ""),
    }


def _should_skip_line(line: str) -> bool:
    if not line:
        return True
    if line.startswith("EANs:"):
        return True
    return any(line.startswith(prefix) for prefix in SKIP_PREFIXES)


def ler_pdf_sempre_vale(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    extracao_pdf = extract_pages_text_detailed(caminho_arquivo)
    paginas = extracao_pdf.paginas

    rows: List[Dict[str, str]] = []
    alertas: List[str] = list(extracao_pdf.alertas)

    current_pedido: Optional[str] = None
    current_cnpj: Optional[str] = None
    order_qtd_ok = Decimal("0")
    order_qtd_sem_sku = Decimal("0")
    order_itens_ok = 0
    order_itens_sem_sku = 0
    order_total_pdf: Optional[Decimal] = None
    order_pages: set[int] = set()
    pedidos_encontrados: set[str] = set()
    cnpjs_encontrados: set[str] = set()
    linhas_lidas = 0
    linhas_validas = 0
    linhas_alerta = 0

    def flush_order():
        nonlocal order_qtd_ok, order_qtd_sem_sku, order_itens_ok, order_itens_sem_sku, order_total_pdf, order_pages
        if not current_pedido:
            order_qtd_ok = Decimal("0")
            order_qtd_sem_sku = Decimal("0")
            order_itens_ok = 0
            order_itens_sem_sku = 0
            order_total_pdf = None
            order_pages = set()
            return

        if order_total_pdf is None:
            alertas.append(
                f"Sempre Vale pedido {current_pedido}: SEM_TOTAL_NO_PDF para conferência | paginas={sorted(order_pages)}"
            )
        else:
            total_lido = order_qtd_ok + order_qtd_sem_sku
            if total_lido != order_total_pdf:
                alertas.append(
                    f"Sempre Vale pedido {current_pedido}: TOTAL_DIVERGENTE | "
                    f"qtd_sku={int(order_qtd_ok)} | qtd_sem_sku={int(order_qtd_sem_sku)} | "
                    f"qtd_pdf={int(order_total_pdf)} | paginas={sorted(order_pages)}"
                )
            elif order_itens_sem_sku:
                alertas.append(
                    f"Sempre Vale pedido {current_pedido}: COM_ITENS_SEM_SKU | "
                    f"itens_sem_sku={order_itens_sem_sku} | qtd_sem_sku={int(order_qtd_sem_sku)} | "
                    f"qtd_pdf={int(order_total_pdf)}"
                )

        order_qtd_ok = Decimal("0")
        order_qtd_sem_sku = Decimal("0")
        order_itens_ok = 0
        order_itens_sem_sku = 0
        order_total_pdf = None
        order_pages = set()

    for page_idx, texto in enumerate(paginas, start=1):
        texto = texto or ""
        if not texto.strip():
            alertas.append(f"Sempre Vale página {page_idx}: SEM_TEXTO_EXTRAIVEL")
            continue

        pedido_match = RE_PEDIDO.search(texto)
        page_pedido = _pedido_limpo(pedido_match.group(1)) if pedido_match else current_pedido

        cnpj_matches = RE_CNPJ.findall(texto)
        page_cnpj = _cnpj_digits(cnpj_matches[-1]) if cnpj_matches else current_cnpj

        if page_pedido and current_pedido and page_pedido != current_pedido:
            flush_order()

        if page_pedido:
            current_pedido = page_pedido
            pedidos_encontrados.add(current_pedido)
        if page_cnpj:
            current_cnpj = page_cnpj
            cnpjs_encontrados.add(current_cnpj)
        if current_pedido:
            order_pages.add(page_idx)

        if not current_pedido:
            alertas.append(f"Sempre Vale página {page_idx}: PEDIDO_NAO_IDENTIFICADO")
        if not current_cnpj:
            alertas.append(f"Sempre Vale página {page_idx}: CNPJ_NAO_IDENTIFICADO")
        elif not _matricula(current_cnpj):
            alertas.append(f"Sempre Vale página {page_idx}: CNPJ_SEM_MATRICULA | cnpj={current_cnpj}")

        for line_no, raw_line in enumerate(texto.splitlines(), start=1):
            line = _norm_line(raw_line)
            if not line:
                continue

            total_match = RE_TOTAL.match(line)
            if total_match:
                order_total_pdf = _parse_decimal_br(total_match.group("qtd"))
                continue

            if _should_skip_line(line):
                continue

            if not re.match(r"^\d{4,6}\s+", line):
                continue

            linhas_lidas += 1
            item_match = RE_ITEM.match(line)
            if item_match:
                qtd_txt = item_match.group("qtd")
                qtd_decimal = _parse_decimal_br(qtd_txt)
                qtd_final = _parse_qtd_int(qtd_txt)
                if not qtd_final:
                    alerta = f"Sempre Vale página {page_idx}: QTD_INVALIDA | {line[:160]}"
                    alertas.append(alerta)
                    rows.append(linha_item_com_alerta(
                        caminho_arquivo=caminho_arquivo,
                        layout_usado=layout_config.get("nome_layout", ""),
                        pagina_pdf=page_idx,
                        linha_bruta=line,
                        alerta=alerta,
                        cnpj_lido=current_cnpj or "",
                        sku_lido=only_digits(item_match.group("sku")),
                        quantidade_lida=qtd_txt,
                        numero_pedido_lido=current_pedido or "",
                        descricao_lida=line[:180],
                        codigo_sku_lido=only_digits(item_match.group("sku")),
                        codigo_origem_lido=only_digits(item_match.group("seq")),
                    ))
                    linhas_alerta += 1
                    continue

                rows.append(_row_ok(
                    caminho_arquivo=caminho_arquivo,
                    layout_config=layout_config,
                    pagina_pdf=page_idx,
                    linha_origem=line_no,
                    linha_bruta=line,
                    cnpj=current_cnpj or "",
                    pedido=current_pedido or "",
                    sku=item_match.group("sku"),
                    qtd=qtd_txt,
                    seq=item_match.group("seq"),
                ))
                order_qtd_ok += qtd_decimal.to_integral_value(rounding=ROUND_HALF_UP)
                order_itens_ok += 1
                linhas_validas += 1
                continue

            sem_sku_match = RE_ITEM_SEM_SKU.match(line)
            if sem_sku_match:
                referencia = only_digits(sem_sku_match.group("referencia"))
                qtd_txt = sem_sku_match.group("qtd")
                qtd_final = _parse_qtd_int(qtd_txt)
                qtd_decimal = _parse_decimal_br(qtd_txt).to_integral_value(rounding=ROUND_HALF_UP)
                alerta = (
                    "SEMPRE_VALE_SEM_SKU_COD_FORN: linha possui quantidade, mas não possui SKU na coluna Cod Forn; "
                    "mantida no Excel para validação manual e não deve entrar na fila/TXT até correção."
                )
                rows.append(linha_item_com_alerta(
                    caminho_arquivo=caminho_arquivo,
                    layout_usado=layout_config.get("nome_layout", ""),
                    pagina_pdf=page_idx,
                    linha_bruta=line,
                    alerta=alerta,
                    cnpj_lido=current_cnpj or "",
                    sku_lido="",
                    quantidade_lida=qtd_final,
                    numero_pedido_lido=current_pedido or "",
                    descricao_lida=line[:180],
                    codigo_sku_lido="",
                    codigo_origem_lido=referencia,
                ))
                order_qtd_sem_sku += qtd_decimal
                order_itens_sem_sku += 1
                linhas_alerta += 1
                continue

            alerta = f"Sempre Vale página {page_idx}: LINHA_DE_ITEM_NAO_RECONHECIDA | {line[:180]}"
            alertas.append(alerta)
            rows.append(linha_item_com_alerta(
                caminho_arquivo=caminho_arquivo,
                layout_usado=layout_config.get("nome_layout", ""),
                pagina_pdf=page_idx,
                linha_bruta=line,
                alerta=alerta,
                cnpj_lido=current_cnpj or "",
                quantidade_lida="",
                numero_pedido_lido=current_pedido or "",
                descricao_lida=line[:180],
            ))
            linhas_alerta += 1

    flush_order()

    df_intermediario = build_intermediate_df(rows, caminho_arquivo, layout_config.get("nome_layout", ""))

    # Mesmo linhas pendentes de SKU precisam manter a identificação do cliente.
    # Isso garante que o Excel de validação mostre CNPJ + matrícula e não trate o item
    # como perda silenciosa da loja.
    if not df_intermediario.empty and "cnpj_lido" in df_intermediario.columns and "matricula_lida" in df_intermediario.columns:
        for idx, row in df_intermediario.iterrows():
            if str(row.get("matricula_lida", "") or "").strip():
                continue
            matricula = _matricula(str(row.get("cnpj_lido", "") or ""))
            if matricula:
                df_intermediario.at[idx, "matricula_lida"] = matricula

    print("\n" + "=" * 100)
    print("DEBUG SEMPRE VALE ROBO KOF")
    print("paginas lidas:", len(paginas))
    print("pedidos encontrados:", sorted(pedidos_encontrados))
    print("cnpjs encontrados:", sorted(cnpjs_encontrados))
    print("linhas brutas candidatas:", linhas_lidas)
    print("linhas validas:", linhas_validas)
    print("linhas alerta/sem sku:", linhas_alerta)
    print("linhas intermediarias:", len(df_intermediario))
    print("=" * 100)

    if df_intermediario.empty:
        msg = "Layout inválido ou não reconhecido para Sempre Vale. Verifique se o arquivo enviado corresponde ao padrão TOTVS esperado."
        return {
            "sucesso": False,
            "mensagem": msg,
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "alertas": sorted(set(alertas or [msg])),
        }

    return {
        "sucesso": True,
        "mensagem": (
            f"Leitura PDF Sempre Vale concluída com {linhas_validas} item(ns) válido(s) "
            f"e {linhas_alerta} linha(s) para validação manual."
        ),
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "alertas": sorted(set(alertas)),
    }
