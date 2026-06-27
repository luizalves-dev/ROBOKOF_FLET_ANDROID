from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd

from parsers_pdf.pdf_utils import build_intermediate_df, extract_pages_text_detailed, normalize_qty, only_digits
from terminal_logger import get_terminal_logger


terminal_log = get_terminal_logger("pdf_generico_homologacao")

CNPJ_COMPLETO_RE = re.compile(r"\b(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})\b")
CNPJ_BASE_RE = re.compile(r"\b(\d{2}\.?\d{3}\.?\d{3}/?\d{4})(?!-?\d{2})\b")
GLN_RE = re.compile(r"\b(\d{13})\b")
PEDIDO_RE = re.compile(
    r"(?:PEDIDO(?:\s+DE\s+COMPRA)?|N[ºO°]?\s*PEDIDO|NUM(?:ERO)?\s*PEDIDO|ORDEM|PO)\D{0,25}(\d{4,14})",
    re.IGNORECASE,
)

# Palavras que normalmente indicam cabecalho/rodape. No modo forçado, elas nao
# bloqueiam a linha completamente; apenas reduzem prioridade para evitar item falso.
HEADER_WORDS = {
    "CNPJ", "ENDERECO", "ENDEREÇO", "FORNECEDOR", "CLIENTE", "TOTAL", "SUBTOTAL", "PEDIDO",
    "PAGINA", "PÁGINA", "CONDICAO", "CONDIÇÃO", "PRAZO", "ENTREGA", "FATURAMENTO",
}
ITEM_HINTS = {
    "SKU", "COD", "CODIGO", "CÓDIGO", "PRODUTO", "DESCRICAO", "DESCRIÇÃO", "QTD", "QTDE",
    "QUANT", "QUANTIDADE", "EMB", "CAIXA", "CX", "EAN", "DUN", "REF", "REFERENCIA", "REFERÊNCIA",
}


def _digits(value: object) -> str:
    return only_digits(value)


def _normalizar_qtd(value: object) -> str:
    texto = str(value or "").strip()
    if not texto:
        return ""
    qtd = normalize_qty(texto)
    return qtd or _digits(texto)


def _extrair_cnpj_linha(linha: str) -> Tuple[str, str]:
    """Retorna CNPJ completo e CNPJ-base, quando existirem.

    Alguns PDFs trazem apenas a base de 12 digitos ou separam o DV do CNPJ,
    como acontece em layouts parecidos com Coopercica.
    """
    completo = ""
    base = ""
    m = CNPJ_COMPLETO_RE.search(linha)
    if m:
        completo = _digits(m.group(1))
        base = completo[:12] if len(completo) >= 12 else ""
        return completo, base

    # Ex.: "CNPJ -08 50.974.732/0009" -> base + DV
    m_dv_antes = re.search(
        r"CNPJ\s*[-:]?\s*(\d{2})\s+(\d{2}\.?\d{3}\.?\d{3}/?\d{4})",
        linha,
        flags=re.IGNORECASE,
    )
    if m_dv_antes:
        dv = _digits(m_dv_antes.group(1))
        base = _digits(m_dv_antes.group(2))
        if len(base) == 12 and len(dv) == 2:
            return base + dv, base

    m = CNPJ_BASE_RE.search(linha)
    if m:
        base = _digits(m.group(1))
    return "", base


def _extrair_pedido_linha(linha: str) -> str:
    m = PEDIDO_RE.search(linha)
    if m:
        return _digits(m.group(1))
    return ""


def _score_linha_item(linha: str, numeros: List[str]) -> int:
    up = linha.upper()
    score = 0
    if any(hint in up for hint in ITEM_HINTS):
        score += 3
    if len(numeros) >= 2:
        score += 2
    if any(3 <= len(n) <= 8 for n in numeros):
        score += 1
    if any(len(n) == 13 for n in numeros):
        score += 1
    if any(word in up for word in HEADER_WORDS):
        score -= 2
    return score


def _numeric_tokens_with_positions(linha: str) -> List[Tuple[int, str]]:
    tokens: List[Tuple[int, str]] = []
    for idx, token in enumerate(re.split(r"\s+", linha)):
        clean = token.strip(";:|()[]{}")
        # Mantém 24,00 / 24 / 2.473,33 / 16.329, mas ignora códigos muito longos puros.
        if re.fullmatch(r"\d{1,7}(?:[\.,]\d{1,4})?", clean or ""):
            tokens.append((idx, clean))
    return tokens


def _looks_money(value: str) -> bool:
    text = str(value or "").strip()
    return bool(re.fullmatch(r"\d{1,7}[\.,]\d{2,4}", text))


def _choose_qty_token(linha: str, numeros_qtd: List[str], numeros_cod: List[str], ean: str) -> str:
    """Escolhe quantidade evitando preço/custo/total em layouts de conversão.

    Muitos pedidos em homologação seguem a ordem:
    código/EAN + descrição + QTD + preço/custo + total.
    O fallback antigo pegava o último número da linha e, em alguns layouts,
    isso capturava o total monetário. Aqui damos prioridade ao número anterior
    aos dois últimos valores monetários e também ao primeiro número útil após o EAN.
    """
    codigos_set = {re.sub(r"\D", "", n) for n in numeros_cod if n}
    tokens = _numeric_tokens_with_positions(linha)
    if not tokens:
        return ""

    # Remove números que são exatamente códigos longos já usados como EAN/SKU.
    filtered: List[Tuple[int, str]] = []
    for idx, tok in tokens:
        clean_digits = re.sub(r"\D", "", tok)
        if len(clean_digits) >= 8 and clean_digits in codigos_set:
            continue
        filtered.append((idx, tok))

    if not filtered:
        return ""

    # Regra QTD + preço + total: escolhe o número imediatamente antes dos dois últimos valores monetários.
    if len(filtered) >= 3 and _looks_money(filtered[-1][1]) and _looks_money(filtered[-2][1]):
        candidate = filtered[-3][1]
        normalized = _normalizar_qtd(candidate)
        if normalized:
            return normalized

    # Quando houver EAN, prioriza número útil após o EAN e antes de preço/total.
    if ean:
        parts = re.split(r"\s+", linha)
        ean_idx = next((i for i, part in enumerate(parts) if re.sub(r"\D", "", part) == ean), -1)
        after_ean = [(idx, tok) for idx, tok in filtered if idx > ean_idx] if ean_idx >= 0 else filtered
        if len(after_ean) >= 3 and _looks_money(after_ean[-1][1]) and _looks_money(after_ean[-2][1]):
            normalized = _normalizar_qtd(after_ean[-3][1])
            if normalized:
                return normalized

    # Fallback conservador: número pequeno no final, diferente do SKU/EAN, mas evita totais com valor alto quando houver alternativas.
    for _idx, num in reversed(filtered):
        clean = re.sub(r"\D", "", num)
        if clean in codigos_set or len(clean) > 7:
            continue
        normalized = _normalizar_qtd(num)
        if normalized:
            return normalized
    return ""


def _escolher_sku_qtd(linha: str) -> tuple[str, str, str]:
    """Escolhe SKU/EAN e QTD em modo homologação.

    O objetivo não é homologar layout definitivo, e sim montar uma planilha de
    validação quando o arquivo é desconhecido. Por isso a função é mais flexível,
    mas sempre marca a linha como VALIDAR_HOMOLOGACAO.
    """
    linha = re.sub(r"\s+", " ", str(linha or "")).strip()
    if len(linha) < 4:
        return "", "", ""

    # Mantém números com separador decimal para quantidade e números puros para códigos.
    numeros_qtd = re.findall(r"(?<!\d)(\d{1,7}(?:[\.,]\d{1,3})?)(?!\d)", linha)
    numeros_cod = re.findall(r"(?<!\d)(\d{3,14})(?!\d)", linha)

    if not numeros_cod and not numeros_qtd:
        return "", "", ""

    score = _score_linha_item(linha, numeros_cod)

    # Código: prioriza EAN/DUN 13, depois códigos médios que não pareçam CNPJ-base.
    ean = next((n for n in numeros_cod if len(n) == 13), "")
    sku = ""
    for cand in numeros_cod:
        if len(cand) in {12, 14}:  # normalmente CNPJ/base; não prioriza como SKU
            continue
        sku = cand
        break
    if not sku:
        sku = ean or (numeros_cod[0] if numeros_cod else "")

    # Quantidade: evita capturar preço/custo/total quando a linha tem QTD + preço + total.
    qtd = _choose_qty_token(linha, numeros_qtd, numeros_cod, ean)

    # Se a linha tem indício de item, mas a QTD não ficou clara, ainda cria linha
    # para validação com QTD vazia. Assim o Excel mostra a linha e o motivo.
    if score < 1 and not (sku and qtd):
        return "", "", ean
    return sku, qtd, ean


def _montar_linha_forcada(
    linha: str,
    *,
    pagina_idx: int,
    linha_idx: int,
    cnpj_atual: str,
    cnpj_base_atual: str,
    gln_atual: str,
    pedido_atual: str,
    referencia: str,
) -> dict[str, str] | None:
    linha_limpa = re.sub(r"\s+", " ", str(linha or "")).strip()
    if not linha_limpa:
        return None

    sku, qtd, ean = _escolher_sku_qtd(linha_limpa)
    if not sku and not qtd and not ean:
        return None

    alerta = [
        "Layout em homologacao/rastreabilidade: conferir SKU, QTD, pedido, CNPJ/GLN e matrícula antes de TXT/fila."
    ]
    if not sku:
        alerta.append("SKU não identificado com segurança; validar pela linha bruta.")
    if not qtd:
        alerta.append("QTD não identificada com segurança; validar pela linha bruta.")
    if not (cnpj_atual or cnpj_base_atual or gln_atual):
        alerta.append("Sem CNPJ/GLN no contexto; preencher de/para se aplicável.")
    if not pedido_atual:
        alerta.append("Sem Nº Pedido no contexto da linha.")

    return {
        "cnpj_lido": cnpj_atual,
        "cnpj_base_lido": cnpj_base_atual,
        "gln_lido": gln_atual,
        "codigo_cliente_lido": gln_atual or cnpj_base_atual,
        "sku_lido": sku,
        "codigo_sku_lido": sku,
        "ean_lido": ean,
        "quantidade_lida": qtd,
        "qtd_original": qtd,
        "tipo_qtd_original": "UNIDADE",
        "numero_pedido_lido": pedido_atual,
        "descricao_lida": linha_limpa[:240],
        "linha_bruta": linha_limpa,
        "pagina_pdf": str(pagina_idx),
        "linha_origem": str(linha_idx),
        "origem_extracao": "PDF_GENERICO_HOMOLOGACAO_FORCADO",
        "status_extracao": "VALIDAR_HOMOLOGACAO",
        "alerta_extracao": " | ".join(alerta),
        "modo_rastreabilidade": "SIM",
        "layout_referencia": referencia,
        "confianca_rastreabilidade": "HOMOLOGACAO_FORCADA",
    }


def ler_pdf_generico_homologacao(
    caminho_arquivo: str,
    layout_config: dict,
    mapeamentos_df=None,
    *,
    referencia: str = "RASTREABILIDADE",
) -> dict:
    """Parser genérico para homologação/rastreabilidade.

    Ele é usado quando o layout ainda não foi homologado ou quando a similaridade
    não encontrou sugestão segura. A regra é gerar um Excel de validação sempre
    que houver algum texto/linha candidata, sem liberar TXT/fila automaticamente.
    """
    nome_layout = str(layout_config.get("nome_layout", "Layout em homologacao"))
    audit = extract_pages_text_detailed(caminho_arquivo)
    rows: List[Dict[str, str]] = []
    alertas = [
        f"LAYOUT_EM_HOMOLOGACAO: {nome_layout}. Conferir obrigatoriamente o Excel antes de TXT/fila.",
        f"REFERENCIA_RASTREABILIDADE: {referencia}",
        "PARSER_GENERICO_FORCADO: arquivo sem layout seguro; linhas candidatas foram planilhadas para validação manual.",
    ]
    cnpj_atual = ""
    cnpj_base_atual = ""
    gln_atual = ""
    pedido_atual = ""
    linhas_texto = 0
    linhas_candidatas = 0

    for pagina_idx, texto in enumerate(audit.paginas, start=1):
        linhas = str(texto or "").splitlines()
        for linha_idx, linha in enumerate(linhas, start=1):
            linha_limpa = re.sub(r"\s+", " ", str(linha or "")).strip()
            if not linha_limpa:
                continue
            linhas_texto += 1

            cnpj, cnpj_base = _extrair_cnpj_linha(linha_limpa)
            if cnpj:
                cnpj_atual = cnpj
                cnpj_base_atual = cnpj_base or cnpj[:12]
            elif cnpj_base:
                cnpj_base_atual = cnpj_base

            glns = [g for g in GLN_RE.findall(linha_limpa) if g not in {cnpj_atual, cnpj_base_atual}]
            if glns:
                gln_atual = glns[0]

            pedido = _extrair_pedido_linha(linha_limpa)
            if pedido:
                pedido_atual = pedido

            item = _montar_linha_forcada(
                linha_limpa,
                pagina_idx=pagina_idx,
                linha_idx=linha_idx,
                cnpj_atual=cnpj_atual,
                cnpj_base_atual=cnpj_base_atual,
                gln_atual=gln_atual,
                pedido_atual=pedido_atual,
                referencia=referencia,
            )
            if item:
                linhas_candidatas += 1
                rows.append(item)

    # Último fallback: se há texto, mas nenhuma linha candidata, cria até 30 linhas
    # brutas para que o Excel mostre evidência do que foi lido. Essas linhas entram
    # com SKU/QTD vazios e serão sinalizadas na validação, sem virar fila.
    if not rows and linhas_texto:
        max_raw = 30
        count = 0
        for pagina_idx, texto in enumerate(audit.paginas, start=1):
            for linha_idx, linha in enumerate(str(texto or "").splitlines(), start=1):
                linha_limpa = re.sub(r"\s+", " ", str(linha or "")).strip()
                if not linha_limpa or len(linha_limpa) < 5:
                    continue
                rows.append({
                    "cnpj_lido": "",
                    "cnpj_base_lido": "",
                    "gln_lido": "",
                    "codigo_cliente_lido": "",
                    "sku_lido": "",
                    "codigo_sku_lido": "",
                    "ean_lido": "",
                    "quantidade_lida": "",
                    "numero_pedido_lido": "",
                    "descricao_lida": linha_limpa[:240],
                    "linha_bruta": linha_limpa,
                    "pagina_pdf": str(pagina_idx),
                    "linha_origem": str(linha_idx),
                    "origem_extracao": "PDF_GENERICO_HOMOLOGACAO_LINHA_BRUTA",
                    "status_extracao": "VALIDAR_HOMOLOGACAO",
                    "alerta_extracao": "Linha bruta planilhada porque o layout não teve sugestão segura. Criar parser específico ou selecionar layout correto.",
                    "modo_rastreabilidade": "SIM",
                    "layout_referencia": referencia,
                    "confianca_rastreabilidade": "LINHA_BRUTA",
                })
                count += 1
                if count >= max_raw:
                    break
            if count >= max_raw:
                break
        alertas.append("Nenhum item candidato seguro foi extraido; linhas brutas foram enviadas ao Excel para análise manual.")

    df = build_intermediate_df(rows, caminho_arquivo, nome_layout)
    if df.empty:
        alertas.append(
            "Nenhum texto/item extraído pelo parser genérico. Se for PDF imagem, instalar OCR/Tesseract ou enviar PDF editável."
        )

    terminal_log.info(
        "[PDF_GENERICO] arquivo=%s | layout=%s | referencia=%s | paginas=%s | texto_linhas=%s | candidatas=%s | linhas_excel=%s",
        Path(caminho_arquivo).name,
        nome_layout,
        referencia,
        audit.total_paginas,
        linhas_texto,
        linhas_candidatas,
        len(df),
    )
    return {
        "sucesso": not df.empty,
        "mensagem": "Leitura genérica de homologação concluída" if not df.empty else "Nenhum item/texto extraído no layout em homologação",
        "df_intermediario": df,
        "qtd_linhas_lidas": len(df),
        "qtd_itens_extraidos": len(df),
        "paginas_pdf_total": audit.total_paginas,
        "paginas_pdf_processadas": audit.paginas_processadas,
        "paginas_pdf_sem_texto": int(sum(1 for a in audit.auditoria if not a.caracteres)),
        "motores_pdf": ", ".join(sorted({a.motor for a in audit.auditoria})),
        "df_auditoria_paginas": audit.auditoria_df(),
        "alertas": sorted({str(a) for a in (alertas + audit.alertas) if str(a).strip()}),
    }
