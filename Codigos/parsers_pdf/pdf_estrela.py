from __future__ import annotations

import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from typing import Dict, List, Optional

import pandas as pd

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, only_digits

try:
    import pdfplumber  # type: ignore
except ModuleNotFoundError:  # pragma: no cover
    pdfplumber = None  # type: ignore


"""
Parser robusto - Rede Estrela / TOTVS Varejo Supermercados / RELPEDSUPRIM.

Regra de negócio preservada:
- Layout textual TOTVS; pode conter vários pedidos no mesmo PDF e várias páginas por pedido.
- Pedido: cabeçalho "PEDIDO DE COMPRAS 000000/M" ou "000000/L".
- CNPJ correto: bloco direito "DADOS PARA FATURAMENTO", CNPJ da rede Estrela 55.624.498/*.
- Ignorar CNPJ do fornecedor/SPAL 61.186.888/0140-62.
- SKU oficial: primeira coluna "Cod Forn".
- Quantidade final: coluna "Qtde" já em caixaria, sem conversão unidade-caixa.
- EANs: somente auditoria; quando houver mais de um EAN, preserva o primeiro em ean_lido e todos em eans_lidos.
- O parser não descarta silenciosamente página sem item: registra auditoria/alerta.

Autor: Kauê Melo
"""

BR_NUMBER_RE = re.compile(r"^\d{1,3}(?:\.\d{3})*,\d{2}$|^\d+,\d{2}$")
SKU_RE = re.compile(r"^\d{5,6}$")
INT_RE = re.compile(r"^\d+$")
EMB_RE = re.compile(r"^[A-Z]{1,3}$")
RE_PEDIDO_TEXT = re.compile(r"PEDIDO\s+DE\s+COMPRAS\s+(\d{5,8})\s*/?\s*([A-Z])?", re.I)
RE_CNPJ_FORMATADO = re.compile(r"(\d{2}\.\d{3}\.\d{3}/\d{4})\s*-?\s*(\d{2})?")
RE_EANS = re.compile(r"EANs?:\s*([\d,\s]+)", re.I)

DEPARA_ESTRELA_CNPJ_MATRICULA: Dict[str, str] = {
    "55624498000155": "7120082797",
    "55624498000236": "7120084566",
    "55624498000317": "7120086587",
    "55624498000406": "7120090091",
    "55624498000589": "7120092980",
    "55624498000660": "7120105467",
    "55624498000740": "7120096312",
    "55624498000821": "7120098530",
    "55624498000902": "7120102100",
    "55624498001208": "7120311023",
    "55624498001399": "7120459608",
    "55624498001470": "7120523932",
}


def _parse_decimal_br(value: object) -> Optional[Decimal]:
    text = str(value or "").strip()
    if not text:
        return None
    text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _fmt_decimal(value: Decimal | None) -> str:
    if value is None:
        return ""
    if value == value.to_integral_value():
        return str(int(value))
    return format(value.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP).normalize(), "f").replace(".", ",")


def _fmt_qtd(value: str) -> str:
    return _fmt_decimal(_parse_decimal_br(value))


def _group_words_by_line(words: List[dict], tolerance: float = 2.5) -> List[dict]:
    lines: List[dict] = []
    for word in sorted(words, key=lambda item: (float(item.get("top", 0)), float(item.get("x0", 0)))):
        top = float(word.get("top", 0))
        placed = False
        for line in lines:
            if abs(float(line["top"]) - top) <= tolerance:
                line["words"].append(word)
                line["top"] = (float(line["top"]) * (len(line["words"]) - 1) + top) / len(line["words"])
                placed = True
                break
        if not placed:
            lines.append({"top": top, "words": [word]})

    for line in lines:
        line["words"].sort(key=lambda item: float(item.get("x0", 0)))
        line["text"] = " ".join(str(word.get("text", "")) for word in line["words"])
    return lines


def _find_word_in_x_range(line: dict, xmin: float, xmax: float, pattern: Optional[re.Pattern] = None) -> str:
    candidates = [
        word for word in line.get("words", [])
        if xmin <= float(word.get("x0", 0)) <= xmax
    ]
    if pattern is not None:
        candidates = [word for word in candidates if pattern.match(str(word.get("text", "")))]
    return str(candidates[0].get("text", "")) if candidates else ""


def _extract_pedido_words(words: List[dict], page_text: str) -> str:
    candidates = []
    for word in words:
        text = str(word.get("text", "")).strip()
        if float(word.get("top", 999)) < 55 and re.match(r"^\d{5,8}/?$", text):
            candidates.append(word)

    if candidates:
        word = max(candidates, key=lambda item: float(item.get("x0", 0)))
        number = only_digits(word.get("text", ""))
        suffix = ""
        for other in words:
            same_line = abs(float(other.get("top", 0)) - float(word.get("top", 0))) < 4
            close_right = float(word.get("x1", 0)) < float(other.get("x0", 0)) < float(word.get("x1", 0)) + 35
            if same_line and close_right and re.match(r"^[A-Z]$", str(other.get("text", ""))):
                suffix = str(other.get("text", ""))
                break
        return f"{number}/{suffix}" if suffix else number

    match = RE_PEDIDO_TEXT.search(page_text or "")
    if match:
        numero = only_digits(match.group(1))
        sufixo = clean_text(match.group(2) or "")
        return f"{numero}/{sufixo}" if sufixo else numero
    return ""


def _extract_cliente_cnpj_words(words: List[dict], page_text: str) -> str:
    # Preferir o bloco direito de faturamento. O fornecedor/SPAL também aparece no cabeçalho e deve ser ignorado.
    for word in sorted(words, key=lambda item: (float(item.get("top", 0)), float(item.get("x0", 0)))):
        if str(word.get("text", "")).upper() == "CNPJ" and float(word.get("x0", 0)) > 600:
            same_line = sorted(
                [
                    other for other in words
                    if abs(float(other.get("top", 0)) - float(word.get("top", 0))) < 4
                    and float(other.get("x0", 0)) > float(word.get("x1", 0))
                ],
                key=lambda item: float(item.get("x0", 0)),
            )
            joined = " ".join(str(item.get("text", "")) for item in same_line[:8])
            match = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4})\s*-\s*(\d{2})", joined)
            if match:
                cnpj = only_digits(match.group(1) + match.group(2))
                if cnpj.startswith("55624498"):
                    return cnpj
            digits = "".join(re.findall(r"\d+", joined))
            if len(digits) >= 14:
                cnpj = digits[-14:]
                if cnpj.startswith("55624498"):
                    return cnpj

    candidatos: list[str] = []
    for match in RE_CNPJ_FORMATADO.finditer(page_text or ""):
        cnpj = only_digits((match.group(1) or "") + (match.group(2) or ""))
        if len(cnpj) == 14:
            candidatos.append(cnpj)

    # CNPJ pode sair quebrado pelo pdfplumber como "CNPJ -55 55.624.498/0001".
    for match in re.finditer(r"CNPJ\s*-?\s*(\d{0,2})\s*(\d{2}\.\d{3}\.\d{3}/\d{4})\s*-?\s*(\d{0,2})", page_text or "", flags=re.I):
        prefixo = only_digits(match.group(1))
        base = only_digits(match.group(2))
        sufixo = only_digits(match.group(3))
        cnpj = (prefixo + base + sufixo)[-14:]
        if len(cnpj) == 14:
            candidatos.append(cnpj)

    estrela = [cnpj for cnpj in candidatos if cnpj.startswith("55624498")]
    return estrela[-1] if estrela else ""


def _extract_items_from_page_words(words: List[dict]) -> List[dict]:
    lines = _group_words_by_line(words)
    items: List[dict] = []

    for line in lines:
        top = float(line.get("top", 0))
        if top < 145 or top > 570:
            continue

        line_words = line.get("words", [])
        if not line_words:
            continue

        first = line_words[0]
        first_text = str(first.get("text", "")).strip()
        if not (float(first.get("x0", 999)) < 50 and SKU_RE.fullmatch(first_text)):
            continue

        qtd_original = _find_word_in_x_range(line, 270, 325, BR_NUMBER_RE)
        if not qtd_original:
            continue

        embalagem = _find_word_in_x_range(line, 208, 236, EMB_RE)
        fator_embalagem = _find_word_in_x_range(line, 232, 250, INT_RE)
        codigo_produto_cliente = _find_word_in_x_range(line, 75, 128, INT_RE)
        descricao_linha = " ".join(
            str(word.get("text", ""))
            for word in line_words
            if 115 <= float(word.get("x0", 0)) < 220
        ).strip()

        items.append({
            "sku": only_digits(first_text),
            "qtd": _fmt_qtd(qtd_original),
            "qtd_original": qtd_original,
            "embalagem": clean_text(embalagem).upper(),
            "fator_embalagem": _fmt_decimal(_parse_decimal_br(fator_embalagem)) if fator_embalagem else "",
            "codigo_produto_cliente": only_digits(codigo_produto_cliente),
            "descricao_linha": descricao_linha,
            "top": top,
        })

    return items


def _extract_eans_after_items(words: List[dict], items: List[dict]) -> None:
    """Associa a próxima linha EANs ao item imediatamente anterior.

    Usa coordenada Y para funcionar mesmo quando a descrição do item quebrou em várias linhas.
    """
    if not items:
        return
    lines = _group_words_by_line(words)
    item_positions = sorted([(float(item.get("top", 0)), idx) for idx, item in enumerate(items)], key=lambda x: x[0])
    for line in lines:
        line_text = str(line.get("text", ""))
        match = RE_EANS.search(line_text)
        if not match:
            continue
        y = float(line.get("top", 0))
        previous = [pair for pair in item_positions if pair[0] < y]
        if not previous:
            continue
        _, item_idx = previous[-1]
        eans = [only_digits(value) for value in re.split(r"[,\s]+", match.group(1)) if only_digits(value)]
        if not eans:
            continue
        items[item_idx]["ean_lido"] = eans[0]
        items[item_idx]["eans_lidos"] = ", ".join(eans)


def ler_pdf_estrela(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    linhas_saida: List[Dict[str, str]] = []
    alertas: List[str] = []
    audit_rows: List[Dict[str, object]] = []
    pedidos_encontrados: set[str] = set()
    cnpjs_encontrados: set[str] = set()
    paginas_com_itens = 0
    paginas_sem_itens = 0
    current_pedido = ""
    current_cnpj = ""

    if pdfplumber is None:
        df_vazio = build_intermediate_df([], caminho_arquivo, layout_config.get("nome_layout", "REDE ESTRELA PDF"))
        return {
            "sucesso": False,
            "mensagem": "pdfplumber nao esta instalado; nao foi possivel aplicar o parser robusto Rede Estrela.",
            "df_intermediario": df_vazio,
            "qtd_linhas_lidas": 0,
            "qtd_itens_extraidos": 0,
            "qtd_itens_ignorados": 0,
            "alertas": sorted(set(alertas + ["ESTRELA_PDFPLUMBER_INDISPONIVEL"])),
            "df_auditoria_paginas": pd.DataFrame(audit_rows),
            "paginas_pdf_total": len(audit_rows),
            "paginas_pdf_processadas": len(audit_rows),
            "paginas_pdf_sem_texto": int(sum(1 for row in audit_rows if int(row.get("caracteres", 0) or 0) == 0)),
        }

    try:
        with pdfplumber.open(str(caminho_arquivo)) as pdf:
            for page_idx, page in enumerate(pdf.pages, start=1):
                try:
                    page_text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
                except Exception as exc_text:
                    page_text = ""
                    alertas.append(f"Pagina {page_idx}: erro ao extrair texto pdfplumber: {exc_text}")
                try:
                    words = page.extract_words(x_tolerance=1, y_tolerance=3, use_text_flow=False) or []
                except TypeError:
                    words = page.extract_words(x_tolerance=1, y_tolerance=3) or []
                audit_rows.append({
                    "pagina": page_idx,
                    "motor": "pdfplumber_words",
                    "caracteres": len(page_text or ""),
                    "tabelas": 0,
                    "blocos": len(words),
                    "status": "OK" if (page_text or words) else "ALERTA",
                    "alerta": "" if (page_text or words) else f"Pagina {page_idx}: PAGINA_SEM_TEXTO_EXTRAIVEL",
                })

                pedido = _extract_pedido_words(words, page_text) or current_pedido
                cnpj = _extract_cliente_cnpj_words(words, page_text) or current_cnpj

                if pedido:
                    current_pedido = pedido
                    pedidos_encontrados.add(pedido)
                elif not current_pedido:
                    alertas.append(f"Pagina {page_idx}: numero do pedido nao localizado.")

                if cnpj:
                    current_cnpj = cnpj
                    cnpjs_encontrados.add(cnpj)
                elif not current_cnpj:
                    alertas.append(f"Pagina {page_idx}: CNPJ cliente nao localizado.")

                items = _extract_items_from_page_words(words)
                _extract_eans_after_items(words, items)
                if not items:
                    paginas_sem_itens += 1
                    continue
                paginas_com_itens += 1

                for idx_item, item in enumerate(items, start=1):
                    obs_alerta: List[str] = []
                    if not current_pedido:
                        obs_alerta.append("pedido_nao_localizado")
                    if not current_cnpj:
                        obs_alerta.append("cnpj_nao_localizado")
                    if not item.get("sku"):
                        obs_alerta.append("sku_cod_forn_nao_localizado")
                    if not item.get("qtd"):
                        obs_alerta.append("qtd_nao_localizada")

                    matricula = DEPARA_ESTRELA_CNPJ_MATRICULA.get(current_cnpj, "")
                    if current_cnpj and not matricula:
                        obs_alerta.append("cnpj_sem_depara_estrela")

                    linhas_saida.append({
                        "matricula_lida": matricula,
                        "cnpj_lido": current_cnpj,
                        "sku_lido": item.get("sku", ""),
                        "codigo_sku_lido": item.get("sku", ""),
                        "ean_lido": item.get("ean_lido", ""),
                        "descricao_lida": item.get("descricao_linha", ""),
                        "codigo_origem_lido": item.get("sku", ""),
                        "quantidade_lida": item.get("qtd", ""),
                        "qtd_original": item.get("qtd_original", item.get("qtd", "")),
                        "tipo_qtd_original": "CAIXARIA",
                        "fator_conversao": "",
                        "qtd_convertida": "",
                        "qtd_final": item.get("qtd", ""),
                        "status_conversao": "OK SEM CONVERSÃO",
                        "tipo_regra_conversao": "SEM_CONVERSAO",
                        "origem_regra_conversao": "PDF_ESTRELA_TOTVS_QTDE_CAIXARIA",
                        "observacao_conversao": "Layout Estrela usa coluna Qtde como caixaria; sem divisão unidade-caixa.",
                        "centro_lido": "",
                        "numero_pedido_lido": current_pedido,
                        "data_entrega_lida": "",
                        "pagina_pdf": str(page_idx),
                        "linha_origem": str(idx_item),
                        "linha_bruta": str(item.get("descricao_linha", "") or "").strip(),
                        "origem_extracao": "PDF_ESTRELA_TOTVS_RELPEDSUPRIM_COORDENADAS",
                        "motor_extracao": "pdfplumber_words",
                        "status_extracao": "OK" if not obs_alerta else "ALERTA",
                        "alerta_extracao": " | ".join(obs_alerta),
                        "embalagem_lida": item.get("embalagem", ""),
                        "fator_embalagem_lido": item.get("fator_embalagem", ""),
                        "codigo_produto_cliente_lido": item.get("codigo_produto_cliente", ""),
                        "eans_lidos": item.get("eans_lidos", ""),
                    })
    except Exception as exc:  # noqa: BLE001 - Robô KOF deve registrar alerta e gerar Excel de validação, sem crash silencioso.
        alertas.append(f"ESTRELA_ERRO_PARSER_COORDENADAS: {exc}")

    nome_layout = layout_config.get("nome_layout", "REDE ESTRELA PDF")
    df_intermediario = build_intermediate_df(linhas_saida, caminho_arquivo, nome_layout)

    if df_intermediario.empty:
        return {
            "sucesso": False,
            "mensagem": "Nenhum item extraido do PDF Rede Estrela",
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": 0,
            "qtd_itens_extraidos": 0,
            "qtd_itens_ignorados": 0,
            "alertas": sorted(set(alertas + ["ESTRELA_SEM_ITENS_EXTRAIDOS"])),
            "df_auditoria_paginas": pd.DataFrame(audit_rows),
            "paginas_pdf_total": len(audit_rows),
            "paginas_pdf_processadas": len(audit_rows),
            "paginas_pdf_sem_texto": int(sum(1 for row in audit_rows if int(row.get("caracteres", 0) or 0) == 0)),
        }

    return {
        "sucesso": True,
        "mensagem": (
            f"Leitura PDF Rede Estrela concluida com {len(df_intermediario)} item(ns), "
            f"{len(pedidos_encontrados)} pedido(s), {len(cnpjs_encontrados)} CNPJ(s), "
            f"{paginas_com_itens} pagina(s) com itens e {paginas_sem_itens} pagina(s) sem itens novos."
        ),
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": len(df_intermediario),
        "qtd_itens_extraidos": len(df_intermediario),
        "qtd_itens_ignorados": 0,
        "alertas": sorted(set(alertas)),
        "df_auditoria_paginas": pd.DataFrame(audit_rows),
        "paginas_pdf_total": len(audit_rows),
        "paginas_pdf_processadas": len(audit_rows),
        "paginas_pdf_sem_texto": int(sum(1 for row in audit_rows if int(row.get("caracteres", 0) or 0) == 0)),
    }
