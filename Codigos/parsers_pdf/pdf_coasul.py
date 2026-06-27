from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Optional

import pandas as pd

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text_detailed, only_digits


RE_PEDIDO = re.compile(r"ORDEM\s*N[º°O]?\s*[:.]?\s*(?P<pedido>\d{3,12})", re.I)
RE_CNPJ = re.compile(r"(?P<cnpj>\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}|\d{14})")
RE_ITEM = re.compile(
    r"^\s*(?P<item>\d{1,4})\s+"
    r"(?P<ean>\d{8,14})\s+"
    r"(?P<produto>.+?)\s+"
    r"(?P<embalagem>UNID|UN|CX|FD|FARDO|LTA|LT|PET|PC|KG|GRF|GFA|VD|TP)\s+"
    r"(?P<qtd>\d{1,9}(?:[\.,]\d+)?)\s+"
    r"(?P<unitario>\d{1,9}(?:[\.,]\d{2,6})?)\s+"
    r"(?P<total>\d{1,12}(?:[\.,]\d{2,6})?)\s*$",
    re.I,
)


def _decimal_br(value: object) -> Optional[Decimal]:
    text = str(value or "").strip()
    if not text:
        return None
    # Coasul usa ponto como decimal nos valores monetários e quantidade inteira.
    # Mantém compatibilidade com vírgula caso venha em outra extração.
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _fmt_decimal(value: Decimal | None) -> str:
    if value is None:
        return ""
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.normalize(), "f").replace(".", ",")


def _extract_pedido(texto: str) -> str:
    match = RE_PEDIDO.search(texto or "")
    return only_digits(match.group("pedido")) if match else ""


def _extract_cnpj_faturamento(texto: str) -> str:
    """Retorna o CNPJ do bloco DADOS PARA FATURAMENTO.

    No PDF Coasul aparecem dois CNPJs: primeiro o cliente Coasul e depois a
    SPAL/fornecedor. A regra corporativa do Robô KOF é usar o CNPJ do cliente.
    """
    page = texto or ""
    bloco = page
    if re.search(r"DADOS\s+PARA\s+FATURAMENTO", page, flags=re.I):
        bloco = re.split(r"DADOS\s+PARA\s+FATURAMENTO", page, flags=re.I, maxsplit=1)[-1]
    # Corta antes do bloco do fornecedor quando possível.
    bloco_cliente = re.split(r"EMPRESA\.:\s*SPAL|SPAL\s+IND", bloco, flags=re.I, maxsplit=1)[0]
    candidatos = [only_digits(m.group("cnpj")) for m in RE_CNPJ.finditer(bloco_cliente)]
    candidatos = [c for c in candidatos if len(c) == 14]
    if candidatos:
        return candidatos[0]
    # Fallback: primeiro CNPJ da página. No layout Coasul ele é o cliente.
    candidatos = [only_digits(m.group("cnpj")) for m in RE_CNPJ.finditer(page)]
    candidatos = [c for c in candidatos if len(c) == 14]
    return candidatos[0] if candidatos else ""


def _parse_item_line(linha: str) -> Optional[Dict[str, str]]:
    match = RE_ITEM.match(linha or "")
    if not match:
        return None
    qtd = _decimal_br(match.group("qtd"))
    unitario = _decimal_br(match.group("unitario"))
    total = _decimal_br(match.group("total"))
    produto = re.sub(r"\s+", " ", match.group("produto") or "").strip()
    return {
        "item": only_digits(match.group("item")),
        "ean": only_digits(match.group("ean")),
        "produto": produto,
        "embalagem": clean_text(match.group("embalagem")).upper(),
        "qtd": _fmt_decimal(qtd),
        "unitario": _fmt_decimal(unitario),
        "valor_total": _fmt_decimal(total),
    }


def _linha_parece_item(linha: str) -> bool:
    # Evita erro silencioso: linhas que parecem item mas não casaram com o regex
    # aparecem na aba ALERTAS/ERROS do Excel de validação.
    return bool(re.match(r"^\s*\d{1,4}\s+\d{8,14}\s+", linha or ""))


def ler_pdf_coasul(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    audit = extract_pages_text_detailed(caminho_arquivo)
    linhas_saida: List[Dict[str, str]] = []
    linhas_alerta: List[Dict[str, str]] = []
    alertas: List[str] = list(audit.alertas)
    pedidos_encontrados: set[str] = set()
    cnpjs_encontrados: set[str] = set()
    linhas_lidas = 0
    current_pedido = ""
    current_cnpj = ""

    for page_idx, texto in enumerate(audit.paginas, start=1):
        pedido = _extract_pedido(texto)
        cnpj = _extract_cnpj_faturamento(texto)
        if pedido:
            current_pedido = pedido
            pedidos_encontrados.add(pedido)
        elif not current_pedido:
            alertas.append(f"Pagina {page_idx}: numero da Ordem Coasul nao localizado.")
        if cnpj:
            current_cnpj = cnpj
            cnpjs_encontrados.add(cnpj)
        elif not current_cnpj:
            alertas.append(f"Pagina {page_idx}: CNPJ de faturamento Coasul nao localizado.")

        for line_idx, raw in enumerate((texto or "").splitlines(), start=1):
            linha = clean_text(raw)
            if not linha:
                continue
            item = _parse_item_line(linha)
            if not item:
                if _linha_parece_item(linha):
                    linhas_alerta.append(
                        {
                            "arquivo_origem": str(caminho_arquivo),
                            "layout_usado": layout_config.get("nome_layout", "REDE COASUL PDF Conversao"),
                            "pagina_pdf": str(page_idx),
                            "linha_origem": str(line_idx),
                            "linha_bruta": linha,
                            "status_extracao": "ALERTA",
                            "alerta_extracao": "Linha com formato de item Coasul não reconhecida integralmente; validar layout/regex.",
                            "origem_extracao": "PDF_COASUL_REGEX_ALERTA",
                        }
                    )
                continue
            linhas_lidas += 1
            obs_alerta: list[str] = []
            if not current_pedido:
                obs_alerta.append("pedido_nao_localizado")
            if not current_cnpj:
                obs_alerta.append("cnpj_nao_localizado")
            if not item.get("ean"):
                obs_alerta.append("ean_cod_barras_nao_localizado")
            if not item.get("qtd"):
                obs_alerta.append("quantidade_nao_localizada")

            linhas_saida.append(
                {
                    "matricula_lida": "",
                    "cnpj_lido": current_cnpj,
                    # O SKU definitivo será resolvido na conversão usando o EAN/COD. BARRAS no mapa.
                    # Enquanto isso, mantemos o EAN como SKU temporário para não perder rastreabilidade.
                    "sku_lido": item["ean"],
                    "codigo_sku_lido": item["ean"],
                    "ean_lido": item["ean"],
                    "descricao_lida": item["produto"],
                    "codigo_origem_lido": item["ean"],
                    "quantidade_lida": item["qtd"],
                    "qtd_original": item["qtd"],
                    "tipo_qtd_original": "UNIDADE",
                    "fator_conversao": "",
                    "qtd_convertida": "",
                    "qtd_final": "",
                    "status_conversao": "",
                    "tipo_regra_conversao": "",
                    "origem_regra_conversao": "",
                    "observacao_conversao": "Coasul: quantidade lida da coluna QTDE em unidade; conversão por EAN/COD. BARRAS via mapa de produtos.",
                    "centro_lido": "",
                    "numero_pedido_lido": current_pedido,
                    "data_entrega_lida": "",
                    "pagina_pdf": str(page_idx),
                    "linha_origem": str(line_idx),
                    "linha_bruta": linha,
                    "origem_extracao": "PDF_COASUL_ORDEM_COMPRA",
                    "motor_extracao": "pdfplumber/fitz/ocr",
                    "status_extracao": "OK" if not obs_alerta else "ALERTA",
                    "alerta_extracao": " | ".join(obs_alerta),
                    "item_pdf": item["item"],
                    "embalagem_lida": item["embalagem"],
                    "unitario_lido": item["unitario"],
                    "valor_total_lido": item["valor_total"],
                }
            )

    df_intermediario = build_intermediate_df(linhas_saida, caminho_arquivo, layout_config.get("nome_layout", "REDE COASUL PDF Conversao"))
    df_alertas = pd.DataFrame(linhas_alerta)
    auditoria_df = audit.auditoria_df()
    paginas_sem_texto = int((auditoria_df["caracteres"] == 0).sum()) if not auditoria_df.empty and "caracteres" in auditoria_df.columns else 0
    motores = ", ".join(sorted(set(auditoria_df["motor"].astype(str)))) if not auditoria_df.empty and "motor" in auditoria_df.columns else ""

    if df_intermediario.empty:
        return {
            "sucesso": False,
            "mensagem": "Nenhum item extraido do PDF Coasul.",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": linhas_lidas,
            "qtd_itens_extraidos": 0,
            "qtd_itens_ignorados": len(df_alertas),
            "alertas": sorted(set(alertas + ["COASUL_SEM_ITENS_EXTRAIDOS"])),
            "df_auditoria_paginas": auditoria_df,
            "df_itens_ignorados": df_alertas,
            "paginas_pdf_total": audit.total_paginas,
            "paginas_pdf_processadas": audit.paginas_processadas,
            "paginas_pdf_sem_texto": paginas_sem_texto,
            "motores_pdf": motores,
        }

    if df_alertas.empty:
        df_alertas = pd.DataFrame(columns=["arquivo_origem", "layout_usado", "pagina_pdf", "linha_origem", "linha_bruta", "status_extracao", "alerta_extracao", "origem_extracao"])

    return {
        "sucesso": True,
        "mensagem": f"Leitura PDF Coasul concluida com {len(df_intermediario)} item(ns), {len(pedidos_encontrados)} pedido(s) e {len(cnpjs_encontrados)} CNPJ(s).",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": linhas_lidas,
        "qtd_itens_extraidos": len(df_intermediario),
        "qtd_itens_ignorados": len(df_alertas),
        "alertas": sorted(set(alertas)),
        "df_auditoria_paginas": auditoria_df,
        "df_itens_ignorados": df_alertas,
        "paginas_pdf_total": audit.total_paginas,
        "paginas_pdf_processadas": audit.paginas_processadas,
        "paginas_pdf_sem_texto": paginas_sem_texto,
        "motores_pdf": motores,
    }
