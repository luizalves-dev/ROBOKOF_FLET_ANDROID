from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Tuple

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text_detailed, only_digits

try:
    import pdfplumber  # type: ignore
except ModuleNotFoundError:
    pdfplumber = None  # type: ignore


def _ensure_pdfplumber():
    if pdfplumber is None:
        raise RuntimeError("pdfplumber nao esta disponivel. Instale com: python -m pip install pdfplumber")


def _result(
    *,
    caminho_arquivo: str,
    layout_config: dict,
    rows: List[Dict[str, str]],
    rede: str,
    alertas: Iterable[str] | None = None,
    erro_layout: str = "",
):
    alertas_final = sorted({a for a in (alertas or []) if a})
    df_intermediario = build_intermediate_df(rows, caminho_arquivo, layout_config.get("nome_layout", ""))

    if erro_layout:
        return {
            "sucesso": False,
            "mensagem": erro_layout,
            "df_intermediario": df_intermediario,
            "qtd_linhas_lidas": 0,
            "alertas": alertas_final or [erro_layout],
        }

    sucesso = not df_intermediario.empty
    return {
        "sucesso": sucesso,
        "mensagem": (
            f"Leitura PDF {rede} concluida com {len(df_intermediario)} linha(s)"
            if sucesso
            else f"Nenhuma linha valida foi extraida do PDF {rede}"
        ),
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": len(df_intermediario),
        "alertas": alertas_final,
    }


def _br_decimal(value: Any) -> Decimal:
    text = str(value or "").strip().replace(" ", "")
    if not text:
        return Decimal("0")

    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        parts = text.split(".")
        if len(parts) > 1 and all(len(part) == 3 for part in parts[1:]):
            text = "".join(parts)

    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def _qty_text(value: Any) -> str:
    number = _br_decimal(value)
    if number == number.to_integral_value():
        return str(int(number))
    return format(number.normalize(), "f").replace(".", ",")


def _page_texts(caminho_arquivo: str, *, x_tolerance: int = 1, y_tolerance: int = 3) -> List[str]:
    return extract_pages_text_detailed(caminho_arquivo).paginas


def _group_words_by_line(words: List[dict], tolerance: float = 3.0) -> List[List[dict]]:
    lines: List[dict] = []
    for word in sorted(words, key=lambda w: (float(w.get("top", 0)), float(w.get("x0", 0)))):
        top = float(word.get("top", 0))
        placed = False
        for line in lines:
            if abs(float(line["top"]) - top) <= tolerance:
                line["words"].append(word)
                line["tops"].append(top)
                line["top"] = sum(line["tops"]) / len(line["tops"])
                placed = True
                break
        if not placed:
            lines.append({"top": top, "tops": [top], "words": [word]})

    return [sorted(line["words"], key=lambda w: float(w.get("x0", 0))) for line in sorted(lines, key=lambda l: l["top"])]


def _build_row(
    cnpj: str,
    sku: str,
    qtd: Any,
    pedido: str,
    matricula: str = "",
    descricao: str = "",
    ean: str = "",
    codigo_origem: str = "",
    *,
    pagina_pdf: Any = "",
    linha_origem: Any = "",
    linha_bruta: str = "",
    origem_extracao: str = "PDF_TEXT",
    alerta_extracao: str = "",
) -> Dict[str, str]:
    """Monta uma linha no padrão intermediário do Robô KOF.

    Regras corporativas preservadas:
    - não descarta item por falta de matrícula;
    - mantém CNPJ/SKU/EAN/QTD/pedido rastreáveis;
    - informa página/linha/linha bruta para auditoria no Excel de validação.
    """
    sku_final = only_digits(sku) or clean_text(sku)
    qtd_final = _qty_text(qtd)
    return {
        "matricula_lida": only_digits(matricula),
        "cnpj_lido": only_digits(cnpj),
        "sku_lido": sku_final,
        "codigo_sku_lido": sku_final,
        "ean_lido": only_digits(ean) or clean_text(ean),
        "codigo_origem_lido": only_digits(codigo_origem) or clean_text(codigo_origem),
        "descricao_lida": clean_text(descricao),
        "quantidade_lida": qtd_final,
        "numero_pedido_lido": clean_text(pedido),
        "data_entrega_lida": "",
        "pagina_pdf": clean_text(pagina_pdf),
        "linha_origem": clean_text(linha_origem),
        "linha_bruta": clean_text(linha_bruta),
        "origem_extracao": origem_extracao,
        "motor_extracao": "pdfplumber",
        "status_extracao": "OK",
        "alerta_extracao": clean_text(alerta_extracao),
        "qtd_original": qtd_final,
        "tipo_qtd_original": "CAIXARIA",
        "qtd_final": qtd_final,
        "status_conversao": "OK SEM CONVERSÃO",
        "tipo_regra_conversao": "SEM_CONVERSAO",
        "regra_aplicada_conversao": "PADRAO_LAYOUT_SEM_REGRA_CONVERSAO",
        "origem_regra_conversao": "PARSER_LAYOUT_DEDICADO",
        "observacao_conversao": "Layout dedicado sem conversão unidade-caixaria; quantidade já lida em caixaria/embalagem.",
    }


# ---------------------------------------------------------------------------
# BOZZA
# ---------------------------------------------------------------------------

BOZZA_ORDER_RE = re.compile(r"Numero\s+Pedido\s*:\s*(\d+)|N\S*mero\s+Pedido\s*:\s*(\d+)", re.I)
BOZZA_PROTOCOLO_RE = re.compile(r"Protocolo\s*:\s*([^|\n]+)", re.I)
BOZZA_CNPJ_RE = re.compile(r"CNPJ/CPF\s*:\s*([\d./-]+)", re.I)
BOZZA_DATA_ENTREGA_RE = re.compile(r"Dt\s+Entrega\s*:\s*([0-9]{2}/[0-9]{2}/[0-9]{4})", re.I)
BOZZA_NUM_RE = r"(?:\d{1,3}(?:\.\d{3})*,\d{2,3}|\d+,\d{2,3}|\d+)"
BOZZA_ITEM_RE = re.compile(
    r"^\s*(?P<codigo_entidade>\d+)\s+"
    r"(?P<sku>\d{4,8})\s+"
    r"(?P<ean>\d{8,14})\s+"
    r"(?P<descricao>.+?)\s+"
    rf"(?P<qtd_emb>{BOZZA_NUM_RE})\s+"
    rf"(?P<qtd_unitaria>{BOZZA_NUM_RE})\s+"
    rf"(?P<valor_unitario>{BOZZA_NUM_RE})\s+"
    r"(?P<valor_pedido>\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2})\s*$",
    re.I,
)
BOZZA_CNPJ_MATRICULA = {
    "73419905000174": "7120019738",
    "73419905000417": "7120173165",
    "73419905000506": "7120173142",
    "73419905000689": "7120174539",
    "73419905000760": "7120174532",
    "73419905000921": "7120468819",
    "73419905001065": "7120468685",
    "73419905001146": "7120468683",
    "73419905001227": "7120472555",
}


def _bozza_order(text: str) -> str:
    match = BOZZA_ORDER_RE.search(text or "")
    if not match:
        return ""
    return clean_text(next(group for group in match.groups() if group))


def _bozza_protocolo(text: str) -> str:
    match = BOZZA_PROTOCOLO_RE.search(text or "")
    return clean_text(match.group(1)) if match else ""


def _bozza_data_entrega(text: str) -> str:
    match = BOZZA_DATA_ENTREGA_RE.search(text or "")
    return clean_text(match.group(1)) if match else ""


def _bozza_item(line: str) -> Dict[str, str] | None:
    linha = " ".join(str(line or "").split())
    match = BOZZA_ITEM_RE.match(linha)
    if not match:
        return None
    dados = {k: clean_text(v) for k, v in match.groupdict().items()}
    if _br_decimal(dados.get("qtd_emb")) <= 0:
        return None
    return dados


def _bozza_cnpj_formatado(cnpj: str) -> str:
    digits = only_digits(cnpj)
    if len(digits) != 14:
        return digits
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"


def ler_pdf_bozza(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedido_atual = ""
    cnpj_atual = ""
    protocolo_atual = ""
    data_entrega_atual = ""
    linhas_aparencia_item = 0

    for pagina_idx, text in enumerate(_page_texts(caminho_arquivo), start=1):
        pedido_atual = _bozza_order(text) or pedido_atual
        protocolo_atual = _bozza_protocolo(text) or protocolo_atual
        data_entrega_atual = _bozza_data_entrega(text) or data_entrega_atual
        cnpj_match = BOZZA_CNPJ_RE.search(text)
        cnpj_atual = only_digits(cnpj_match.group(1)) if cnpj_match else cnpj_atual

        for linha_idx, raw_line in enumerate(str(text or "").splitlines(), start=1):
            linha = " ".join(raw_line.split())
            parsed = _bozza_item(linha)
            if not parsed:
                if re.match(r"^\s*\d+\s+\d{4,8}\s+\d{8,14}\s+", linha):
                    linhas_aparencia_item += 1
                    alertas.append(
                        f"BOZZA_ITEM_NAO_INTERPRETADO | pagina={pagina_idx} | linha={linha_idx} | texto={linha[:180]}"
                    )
                continue

            row = _build_row(
                cnpj_atual,
                parsed.get("sku", ""),
                parsed.get("qtd_emb", ""),
                pedido_atual,
                BOZZA_CNPJ_MATRICULA.get(cnpj_atual, ""),
                descricao=parsed.get("descricao", ""),
                ean=parsed.get("ean", ""),
                codigo_origem=parsed.get("codigo_entidade", ""),
                pagina_pdf=pagina_idx,
                linha_origem=linha_idx,
                linha_bruta=linha,
                origem_extracao="PDF_BOZZA_LINKERP",
            )
            row.update({
                "data_entrega_lida": data_entrega_atual,
                "codigo_entidade_lido": parsed.get("codigo_entidade", ""),
                "codigo_barra_lido": parsed.get("ean", ""),
                "qtd_emb_lida": _qty_text(parsed.get("qtd_emb", "")),
                "qtd_unitaria_lida": _qty_text(parsed.get("qtd_unitaria", "")),
                "valor_unitario_lido": clean_text(parsed.get("valor_unitario", "")),
                "valor_pedido_lido": clean_text(parsed.get("valor_pedido", "")),
                "protocolo_lido": protocolo_atual,
                "cnpj_formatado_lido": _bozza_cnpj_formatado(cnpj_atual),
                "layout_referencia": "BOZZA PDF",
                "confianca_rastreabilidade": "PARSER_DEDICADO",
            })
            rows.append(row)

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Bozza. Verifique se o PDF enviado corresponde ao padrao esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Bozza", erro_layout=msg)

    if linhas_aparencia_item:
        alertas.append(f"Bozza: {linhas_aparencia_item} linha(s) com aparencia de item nao foram interpretadas.")
    if any(not row["cnpj_lido"] for row in rows):
        alertas.append("Bozza: ha itens sem CNPJ identificado")
    if any(not row["numero_pedido_lido"] for row in rows):
        alertas.append("Bozza: ha itens sem numero de pedido identificado")

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Bozza", alertas=alertas)


# ---------------------------------------------------------------------------
# MONACO
# ---------------------------------------------------------------------------

MONACO_ITEM_RE = re.compile(
    r"^\s*(?P<codigo>\d{5,6})\s+"
    r"(?P<cod_barras>\d{8,14})\s+"
    r"(?P<cod_forn>[\d,;/\-\s]+?)\s+"
    r"(?P<descricao>.+?)\s+"
    r"(?P<qtde_emb>\d+)\s+"
    r"(?P<quant>\d{1,3}(?:\.\d{3})*,\d{3}|\d+,\d{3})\s+"
    r"(?P<emb>[A-Z]{1,4}/\d+)\s+",
    re.I,
)
MONACO_PEDIDO_RE = re.compile(r"Numero\s+do\s+Pedido\s*:\s*(\d+)|N\S*mero\s+do\s+Pedido\s*:\s*(\d+)|Pedido\s*:\s*(\d{5,})", re.I)
MONACO_CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
MONACO_CNPJ_MATRICULA = {
    "25080804000179": "7140134157",
    "25080804000411": "7140267712",
    "25080804000500": "7140360262",
    "25080804000683": "7140379936",
    "25080804000764": "7140420384",
    "25080804000926": "7140467428",
    "25080804001060": "7140482852",
}


def _first_group(match) -> str:
    return clean_text(next((group for group in match.groups() if group), ""))


def _monaco_pedido(text: str) -> str:
    match = MONACO_PEDIDO_RE.search(text or "")
    return _first_group(match) if match else ""


MONACO_FORNECEDOR_PREFIXOS = ("61186888",)  # SPAL/Coca-Cola FEMSA: nunca é CNPJ de loja Monaco


def _is_cnpj_fornecedor_monaco(cnpj: str) -> bool:
    digits = only_digits(cnpj).zfill(14)
    return any(digits.startswith(prefix) for prefix in MONACO_FORNECEDOR_PREFIXOS)


def _monaco_cnpj(text: str) -> str:
    """Extrai somente o CNPJ da loja Monaco/Macamo.

    O PDF também traz o CNPJ do fornecedor SPAL/Coca-Cola FEMSA. Essa função
    prioriza o bloco Empresa e ignora qualquer CNPJ com prefixo 61.186.888,
    evitando que a página seja enviada como fornecedor e caia em A CADASTRAR.
    """
    text = text or ""
    blocos: list[str] = []
    m_bloco = re.search(
        r"Empresa\s*:\s*(.*?)(?:Dt\.\s*Pedido|Frete\s*:|C[oó]digo\s+Cod|Codigo\s+Cod)",
        text,
        flags=re.I | re.S,
    )
    if m_bloco:
        blocos.append(m_bloco.group(1))
    blocos.append(text)

    for block in blocos:
        for raw in MONACO_CNPJ_RE.findall(block):
            digits = only_digits(raw).zfill(14)
            if len(digits) == 14 and not _is_cnpj_fornecedor_monaco(digits):
                return digits
    return ""


def _monaco_sku(cod_forn: str) -> str:
    candidates = re.findall(r"\d{4,8}", cod_forn or "")
    return candidates[0] if candidates else only_digits(cod_forn)


def ler_pdf_monaco(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedido_atual = ""
    cnpj_atual = ""

    for pagina, text in enumerate(_page_texts(caminho_arquivo), start=1):
        pedido_atual = _monaco_pedido(text) or pedido_atual
        cnpj_novo = _monaco_cnpj(text)
        if cnpj_novo and not _is_cnpj_fornecedor_monaco(cnpj_novo):
            cnpj_atual = cnpj_novo

        for linha_num, raw_line in enumerate((text or "").splitlines(), start=1):
            line = " ".join(raw_line.split())
            match = MONACO_ITEM_RE.match(line)
            if not match:
                continue
            cod_forn_bruto = match.group("cod_forn")
            sku = _monaco_sku(cod_forn_bruto)
            qtd = match.group("qtde_emb")
            alerta_linha = ""
            if len(re.findall(r"\d{4,8}", cod_forn_bruto or "")) > 1:
                alerta_linha = f"SKU_DUPLO_COD_FORN: bruto={cod_forn_bruto}; usado={sku}"
            rows.append(
                _build_row(
                    cnpj_atual,
                    sku,
                    qtd,
                    pedido_atual,
                    MONACO_CNPJ_MATRICULA.get(cnpj_atual, ""),
                    descricao=match.group("descricao"),
                    ean=match.group("cod_barras"),
                    codigo_origem=match.group("codigo"),
                    pagina_pdf=pagina,
                    linha_origem=linha_num,
                    linha_bruta=line,
                    origem_extracao="PDF_TEXT_MONACO",
                    alerta_extracao=alerta_linha,
                )
            )

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Monaco. Verifique se o PDF enviado corresponde ao padrao esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Monaco", erro_layout=msg)

    if any(not row["cnpj_lido"] for row in rows):
        alertas.append("Monaco: ha itens sem CNPJ identificado")
    if any(not row["numero_pedido_lido"] for row in rows):
        alertas.append("Monaco: ha itens sem numero de pedido identificado")

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Monaco", alertas=alertas)


# ---------------------------------------------------------------------------
# DAHER
# ---------------------------------------------------------------------------

# Regra preservada da automação original do Kauê Melo:
# - PDF textual com múltiplos pedidos no mesmo arquivo;
# - pedido no cabeçalho: "Pedido de Compras N° 10978";
# - CNPJ válido da loja fica antes do bloco "Destinatário"; o CNPJ da SPAL/fornecedor
#   aparece depois e deve ser ignorado como cliente;
# - SKU oficial = coluna Ref.;
# - EAN = primeira coluna;
# - quantidade final = coluna Qtd.;
# - sem conversão unidade-caixaria.

DAHER_ORDER_RE = re.compile(r"Pedido\s+de\s+Compras\s+N[^0-9]*(\d+)", re.I)
DAHER_CNPJ_RE = re.compile(r"CNPJ:\s*([\d.\-/]+)", re.I)
DAHER_CNPJ_MATRICULA = {
    "45291341000100": "7120018178",
    "45291341000363": "7120018197",
    "45291341000444": "7120012567",
    "45291341000606": "7120013557",
    "11209903000101": "7120005120",
    "11209903000365": "7120221386",
    "39420850000184": "7120318947",
}
DAHER_CNPJ_FORNECEDOR_PREFIXOS = ("61186888",)
DAHER_ITEM_RE = re.compile(
    r"^(?P<ean>\d{8,14})\s+"
    r"(?P<descricao>.+?)\s+"
    r"(?P<plu>\d{3,8})\s+"
    r"(?P<ref>\d{3,8})\s+"
    r"(?P<emb>[A-Z]{1,4}\d{0,3})\s+"
    r"(?P<qtd>\d+(?:[.,]\d+)?)\s+R\$",
    re.I,
)


def _daher_context(text: str, pedido_anterior: str, cnpj_anterior: str) -> Tuple[str, str]:
    pedido_match = DAHER_ORDER_RE.search(text or "")
    pedido = pedido_match.group(1) if pedido_match else pedido_anterior

    bloco_emitente = re.split(r"Destinat\S*rio:", text or "", maxsplit=1, flags=re.I)[0]
    candidatos = [only_digits(c) for c in DAHER_CNPJ_RE.findall(bloco_emitente)]
    if not candidatos:
        candidatos = [only_digits(c) for c in DAHER_CNPJ_RE.findall(text or "")]
    candidatos = [c for c in candidatos if len(c) == 14 and not c.startswith(DAHER_CNPJ_FORNECEDOR_PREFIXOS)]
    cnpj = candidatos[0] if candidatos else cnpj_anterior
    return pedido, cnpj


def _daher_items(text: str) -> List[Dict[str, str]]:
    header = re.search(r"EAN\s+Descri\S+\s+PLU\s+Ref\.\s+Emb\.\s+Qtd\.", text or "", flags=re.I)
    if not header:
        return []
    tail = text[header.end():]
    total = re.search(r"\nTotal:", tail, flags=re.I)
    if total:
        tail = tail[:total.start()]

    items: List[Dict[str, str]] = []
    for line_no, raw_line in enumerate(tail.splitlines(), start=1):
        line = " ".join(str(raw_line or "").replace("\xa0", " ").split())
        if not line:
            continue
        match = DAHER_ITEM_RE.match(line)
        if not match:
            # Não registra qualquer linha do rodapé como item inválido; apenas linhas
            # com cara de EAN no início são úteis para investigação.
            if re.match(r"^\d{8,14}\s+", line):
                items.append({
                    "sku": "",
                    "qtd": "",
                    "descricao": line[:180],
                    "ean": only_digits(line.split()[0]),
                    "plu": "",
                    "emb": "",
                    "line_no": str(line_no),
                    "linha_bruta": line,
                    "alerta": "DAHER_LINHA_DE_ITEM_NAO_RECONHECIDA",
                })
            continue
        items.append({
            "sku": match.group("ref"),
            "qtd": match.group("qtd"),
            "descricao": match.group("descricao"),
            "ean": match.group("ean"),
            "plu": match.group("plu"),
            "emb": match.group("emb"),
            "line_no": str(line_no),
            "linha_bruta": line,
            "alerta": "",
        })
    return items


def ler_pdf_daher(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    extracao_pdf = extract_pages_text_detailed(caminho_arquivo)
    rows: List[Dict[str, str]] = []
    alertas: List[str] = list(extracao_pdf.alertas)
    pedido_atual = ""
    cnpj_atual = ""
    paginas_com_item = 0
    paginas_sem_item_util = 0

    for page_idx, text in enumerate(extracao_pdf.paginas, start=1):
        text = text or ""
        pedido_atual, cnpj_atual = _daher_context(text, pedido_atual, cnpj_atual)
        itens = _daher_items(text)
        if itens:
            paginas_com_item += 1
        else:
            paginas_sem_item_util += 1

        if itens and not pedido_atual:
            alertas.append(f"Daher página {page_idx}: PEDIDO_NAO_IDENTIFICADO em página com item.")
        if itens and not cnpj_atual:
            alertas.append(f"Daher página {page_idx}: CNPJ_NAO_IDENTIFICADO em página com item.")
        if cnpj_atual and cnpj_atual not in DAHER_CNPJ_MATRICULA:
            alertas.append(f"Daher página {page_idx}: CNPJ_SEM_MATRICULA | cnpj={cnpj_atual}")

        for item in itens:
            alerta_linha = item.get("alerta", "")
            sku = item.get("sku", "")
            qtd = item.get("qtd", "")
            if not sku or not qtd:
                alerta = alerta_linha or "DAHER_ITEM_PENDENTE_VALIDACAO"
                rows.append(
                    _build_row(
                        cnpj_atual,
                        sku,
                        qtd,
                        pedido_atual,
                        DAHER_CNPJ_MATRICULA.get(cnpj_atual, ""),
                        descricao=item.get("descricao", ""),
                        ean=item.get("ean", ""),
                        codigo_origem=item.get("plu", ""),
                        pagina_pdf=page_idx,
                        linha_origem=item.get("line_no", ""),
                        linha_bruta=item.get("linha_bruta", ""),
                        origem_extracao="PDF_TEXT_DAHER_ALERTA",
                        alerta_extracao=alerta,
                    )
                )
                alertas.append(f"Daher página {page_idx}: {alerta} | {item.get('linha_bruta', '')[:160]}")
                continue

            rows.append(
                _build_row(
                    cnpj_atual,
                    sku,
                    qtd,
                    pedido_atual,
                    DAHER_CNPJ_MATRICULA.get(cnpj_atual, ""),
                    descricao=item.get("descricao", ""),
                    ean=item.get("ean", ""),
                    codigo_origem=item.get("plu", ""),
                    pagina_pdf=page_idx,
                    linha_origem=item.get("line_no", ""),
                    linha_bruta=item.get("linha_bruta", ""),
                    origem_extracao="PDF_TEXT_DAHER",
                )
            )

    if not rows:
        msg = "Layout inválido ou não reconhecido para Daher. Verifique se o PDF enviado corresponde ao padrão esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Daher", erro_layout=msg)

    if any(not row["cnpj_lido"] for row in rows):
        alertas.append("Daher: há itens sem CNPJ identificado")
    if any(not row["numero_pedido_lido"] for row in rows):
        alertas.append("Daher: há itens sem número de pedido identificado")

    alertas.append(
        f"Daher auditoria: páginas lidas={extracao_pdf.total_paginas}; "
        f"páginas com item={paginas_com_item}; páginas sem item útil/rodapé={paginas_sem_item_util}; "
        f"linhas extraídas={len(rows)}."
    )

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Daher", alertas=alertas)


# ---------------------------------------------------------------------------
# PRIMATO
# ---------------------------------------------------------------------------

PRIMATO_ORDER_RE = re.compile(r"MERC\s*-\s*PEDIDO\s+DE\s+COMPRA\s+N\S*\s*(\d+)", re.I)
PRIMATO_CNPJ_RE = re.compile(r"CNPJ\s*:\s*([0-9.\-/]{13,18})\s+IE\s*:", re.I)
PRIMATO_CNPJ_MATRICULA = {
    "02168202001144": "7120163679",
    "02168202001225": "7120009289",
    "02168202001578": "7120299672",
    "02168202001730": "7120028069",
    "02168202002205": "7120069029",
    "02168202002540": "7120191214",
    "02168202002701": "7120195261",
    "02168202003279": "7120287692",
}
PRIMATO_ITEM_RE = re.compile(
    r"^\s*(?P<sku>\d{5,6})\s+"
    r"(?P<cod_barras>\d{6,14})\s+"
    r"(?P<descricao>.+?)\s+"
    r"(?P<qtd>\d{1,9}(?:\.\d{3})*,\d{2})\s+"
    r"C/\s*(?P<embalagem>\d+)\s+"
    r"(?P<valor_unitario>\d{1,9}(?:\.\d{3})*,\d{2})\s+"
    r"(?P<desc_item>\d{1,9}(?:\.\d{3})*,\d{2})\s+"
    r"(?P<valor_total>\d{1,12}(?:\.\d{3})*,\d{2})\s*$",
    re.I,
)


def _primato_cnpj(text: str) -> str:
    match = PRIMATO_CNPJ_RE.search(text or "")
    if match:
        return only_digits(match.group(1)).zfill(14)
    for raw in re.findall(r"CNPJ\s*:\s*([0-9.\-/]{13,18})", text or "", flags=re.I):
        digits = only_digits(raw).zfill(14)
        if len(digits) == 14 and digits != "61186888014496":
            return digits
    return ""


def ler_pdf_primato(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedido_atual = ""
    cnpj_atual = ""

    for pagina, text in enumerate(_page_texts(caminho_arquivo, x_tolerance=1, y_tolerance=3), start=1):
        pedido = PRIMATO_ORDER_RE.search(text or "")
        if pedido:
            pedido_atual = pedido.group(1)
        cnpj_atual = _primato_cnpj(text) or cnpj_atual

        for linha_num, raw_line in enumerate((text or "").splitlines(), start=1):
            line = " ".join(raw_line.split())
            match = PRIMATO_ITEM_RE.match(line)
            if not match:
                continue
            rows.append(
                _build_row(
                    cnpj_atual,
                    match.group("sku"),
                    match.group("qtd"),
                    pedido_atual,
                    PRIMATO_CNPJ_MATRICULA.get(cnpj_atual, ""),
                    descricao=match.group("descricao"),
                    ean=match.group("cod_barras"),
                    codigo_origem=match.group("sku"),
                    pagina_pdf=pagina,
                    linha_origem=linha_num,
                    linha_bruta=line,
                    origem_extracao="PDF_TEXT_PRIMATO",
                )
            )

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Primato. Verifique se o PDF enviado corresponde ao padrao esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Primato", erro_layout=msg)

    if any(not row["cnpj_lido"] for row in rows):
        alertas.append("Primato: ha itens sem CNPJ identificado")
    if any(not row["numero_pedido_lido"] for row in rows):
        alertas.append("Primato: ha itens sem numero de pedido identificado")

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Primato", alertas=alertas)


# ---------------------------------------------------------------------------
# SUPERLAR
# ---------------------------------------------------------------------------

SUPERLAR_CNPJ_RE = re.compile(r"CNPJ:\s*([0-9./-]{14,20})", re.I)
SUPERLAR_PEDIDO_RE = re.compile(r"N\S{0,3}\s*PEDIDO\s*:\s*(\d+)", re.I)
SUPERLAR_SUPPLIER_CNPJ = "61186888000274"
# O layout Superlar/VR Software traz a linha no padrao:
# CODIGO DESCRICAO EAN EMBALAGEM SKU QTDE CUSTO FRETE DESC TOTAL
# A versao anterior lia a coluna de custo/preco como quantidade em alguns PDFs
# porque a coluna de QTDE fica imediatamente apos o SKU e antes do custo.
SUPERLAR_ITEM_TEXT_RE = re.compile(
    r"^\s*"
    r"(?P<codigo_ext>\d{5,6})\s+"
    r"(?P<descricao>.+?)\s*"
    r"(?P<ean>\d{8,14})\s+"
    r"(?P<embalagem>[A-Z]{1,4}/\d{3,4})\s+"
    r"(?P<sku>0*\d{5,6})\s+"
    r"(?P<qtd>\d{1,6}(?:[,.]\d{1,2})?)\s+"
    r"(?P<custo>\d{1,6}(?:[,.]\d{2,4}))\b",
    re.I,
)
SUPERLAR_EMBALAGEM_RE = re.compile(r"^[A-Z]{1,4}/\d{3,4}$", re.I)
SUPERLAR_DECIMAL_4_RE = re.compile(r"^\d{1,6}[,.]\d{4}$")
SUPERLAR_EAN_GRUDADO_RE = re.compile(
    r"(?P<prefixo>[A-Za-zÀ-ÿ0-9.,;:%ºª()/+-])(?P<ean>\d{8,14})(?=\s+[A-Z]{1,4}/\d{3,4}\s+0*\d{5,6}\s+\d)",
    re.I,
)
SUPERLAR_CNPJ_MATRICULA = {
    "21676733000200": "7110070348",
    "21676733000463": "7110378581",
    "21676733000625": "7110569403",
    "21676733000706": "7110665004",
    "16938900000114": "7110010324",
    "21676733000110": "7110054527",
    "21676733000544": "7110535516",
    "21676733000382": "7110282768",
}


def _superlar_context(text: str, pedido_atual: str, cnpj_atual: str) -> Tuple[str, str]:
    pedido_match = SUPERLAR_PEDIDO_RE.search(text or "")
    if pedido_match:
        pedido_atual = pedido_match.group(1)

    for match in SUPERLAR_CNPJ_RE.findall(text or ""):
        cnpj = only_digits(match)
        if len(cnpj) == 14 and cnpj != SUPERLAR_SUPPLIER_CNPJ:
            cnpj_atual = cnpj
            break
    return pedido_atual, cnpj_atual


def _superlar_qtd_valida(texto: str) -> bool:
    """Evita capturar custo/preco como quantidade no Superlar.

    No PDF Superlar, valores como 2,4733, 6,0917 e 12,5267 sao custo/preco
    unitario, nao quantidade. A quantidade vem antes do custo, normalmente como
    inteiro ou decimal com ate 2 casas.
    """
    valor = clean_text(texto)
    if not valor:
        return False
    if SUPERLAR_DECIMAL_4_RE.fullmatch(valor):
        return False
    return _br_decimal(valor) > 0


def _superlar_item_from_line(line: str) -> Dict[str, str] | None:
    linha = " ".join(str(line or "").replace(" ", " ").split())
    # Alguns PDFs Superlar/VR Software trazem o EAN colado no fim da descrição,
    # ex.: "473ML1220000250031 CX/0006 000000119170 3 ...".
    # Sem essa normalização, a linha é perdida e o total fica abaixo da
    # conferência enviada pelo usuário.
    linha = SUPERLAR_EAN_GRUDADO_RE.sub(r"\g<prefixo> \g<ean>", linha)
    if not linha:
        return None
    if any(token in linha.upper() for token in ("TOTAIS", "TOTAL LIQUIDO", "PESO BRUTO", "VALOR FRETE")):
        return None

    match = SUPERLAR_ITEM_TEXT_RE.search(linha)
    if match:
        qtd = match.group("qtd")
        if not _superlar_qtd_valida(qtd):
            return None
        return {
            "sku": (only_digits(match.group("sku")).lstrip("0") or only_digits(match.group("sku"))),
            "qtd": qtd,
            "ean": only_digits(match.group("ean")),
            "descricao": clean_text(match.group("descricao")),
            "codigo_origem": only_digits(match.group("codigo_ext")),
            "embalagem": clean_text(match.group("embalagem")),
            "linha_original": linha,
        }

    # Fallback tokenizado para casos em que o texto venha desalinhado, mas ainda
    # mantendo a regra estrutural: EAN -> embalagem -> SKU -> QTDE -> custo.
    parts = linha.split()
    for idx, token in enumerate(parts):
        if not (re.fullmatch(r"\d{8,14}", token or "") and idx + 4 < len(parts)):
            continue
        embalagem = parts[idx + 1]
        sku = parts[idx + 2]
        qtd = parts[idx + 3]
        custo = parts[idx + 4]
        if not SUPERLAR_EMBALAGEM_RE.fullmatch(embalagem or ""):
            continue
        if not re.fullmatch(r"0*\d{5,6}", sku or ""):
            continue
        if not _superlar_qtd_valida(qtd):
            continue
        if not re.fullmatch(r"\d{1,6}(?:[,.]\d{2,4})", custo or ""):
            continue
        descricao = " ".join(parts[1:idx]) if len(parts) > 1 else ""
        return {
            "sku": (only_digits(sku).lstrip("0") or only_digits(sku)),
            "qtd": qtd,
            "ean": only_digits(token),
            "descricao": clean_text(descricao),
            "codigo_origem": only_digits(parts[0]),
            "embalagem": clean_text(embalagem),
            "linha_original": linha,
        }
    return None


def _superlar_items_from_text(text: str, pagina_pdf: int | str = "") -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for linha_num, raw_line in enumerate(str(text or "").splitlines(), start=1):
        parsed = _superlar_item_from_line(raw_line)
        if parsed:
            parsed["pagina_pdf"] = str(pagina_pdf or "")
            parsed["linha_origem"] = str(linha_num)
            parsed["linha_bruta"] = str(raw_line or "")
            rows.append(parsed)
    return rows


def _superlar_items_from_page(page, text: str = "", pagina_pdf: int | str = "") -> List[Dict[str, str]]:
    # Preferir o texto linear do PDF: ele preserva a ordem correta das colunas
    # e impede confundir QTDE com CUSTO/COMPRA.
    rows = _superlar_items_from_text(text or "", pagina_pdf=pagina_pdf)
    if rows:
        return rows

    # Fallback por palavras somente para PDFs cujo extract_text venha vazio.
    # Ainda assim, a linha montada e validada pela regra textual acima evita
    # capturar valores de custo com 4 casas decimais como quantidade.
    words = page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
    for line_words in _group_words_by_line(words, tolerance=3.0):
        line_text = " ".join(clean_text(word.get("text", "")) for word in line_words if clean_text(word.get("text", "")))
        parsed = _superlar_item_from_line(line_text)
        if parsed:
            parsed["pagina_pdf"] = str(pagina_pdf or "")
            parsed["linha_origem"] = "WORDS"
            parsed["linha_bruta"] = line_text
            rows.append(parsed)

    return rows


def ler_pdf_superlar(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    _ensure_pdfplumber()
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []
    pedido_atual = ""
    cnpj_atual = ""

    with pdfplumber.open(caminho_arquivo) as pdf:
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            pedido_atual, cnpj_atual = _superlar_context(text, pedido_atual, cnpj_atual)

            itens_pagina = _superlar_items_from_page(page, text, pagina_pdf=getattr(page, "page_number", ""))
            for item in itens_pagina:
                obs_descricao = item.get("descricao", "")
                embalagem = item.get("embalagem", "")
                if embalagem:
                    obs_descricao = f"{obs_descricao} | EMB={embalagem}" if obs_descricao else f"EMB={embalagem}"
                rows.append(
                    _build_row(
                        cnpj_atual,
                        item.get("sku", ""),
                        item.get("qtd", ""),
                        pedido_atual,
                        SUPERLAR_CNPJ_MATRICULA.get(cnpj_atual, ""),
                        descricao=obs_descricao,
                        ean=item.get("ean", ""),
                        codigo_origem=item.get("codigo_origem", ""),
                        pagina_pdf=item.get("pagina_pdf", ""),
                        linha_origem=item.get("linha_origem", ""),
                        linha_bruta=item.get("linha_bruta", ""),
                        origem_extracao="PDF_TEXT_SUPERLAR",
                    )
                )

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Superlar. Verifique se o PDF enviado corresponde ao padrao esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Superlar", erro_layout=msg)

    if any(not row["cnpj_lido"] for row in rows):
        alertas.append("Superlar: ha itens sem CNPJ identificado")
    if any(not row["numero_pedido_lido"] for row in rows):
        alertas.append("Superlar: ha itens sem numero de pedido identificado")

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Superlar", alertas=alertas)


# ---------------------------------------------------------------------------
# INDIANA
# ---------------------------------------------------------------------------

INDIANA_PEDIDO_RE = re.compile(r"Numero\s*Pedido:\s*(\d{5,12})", re.I)
INDIANA_CNPJ_RE = re.compile(r"(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2})")
INDIANA_ITEM_RE = re.compile(r"^\s*(\d{4,9})\b.*?\b(\d{1,5})\s+0,0000\b")


def ler_pdf_indiana(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    rows: List[Dict[str, str]] = []
    alertas: List[str] = []

    for text in _page_texts(caminho_arquivo):
        pedido_match = INDIANA_PEDIDO_RE.search(text)
        cnpj_match = INDIANA_CNPJ_RE.search(text)
        if not pedido_match or not cnpj_match:
            continue

        pedido = pedido_match.group(1)
        cnpj = only_digits(cnpj_match.group(1))
        for line in text.splitlines():
            item = INDIANA_ITEM_RE.search(line)
            if not item:
                continue
            rows.append(_build_row(cnpj, item.group(1), item.group(2), pedido))

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Indiana. Verifique se o PDF enviado corresponde ao padrao esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Indiana", erro_layout=msg)

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Indiana", alertas=alertas)


# ---------------------------------------------------------------------------
# KACULA
# ---------------------------------------------------------------------------

KACULA_PEDIDO_RE = re.compile(r"\d+\s*/\s*L")
KACULA_CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}\s*-?\s*\d{2}")


def ler_pdf_kacula(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    rows: List[Dict[str, str]] = []
    pedido_atual = ""
    cnpj_atual = ""
    sku_atual = ""

    for text in _page_texts(caminho_arquivo):
        for raw_line in text.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            if not pedido_atual:
                pedido_match = KACULA_PEDIDO_RE.search(line)
                if pedido_match:
                    pedido_atual = re.sub(r"\s+", "", pedido_match.group(0))

            if "CNPJ" in line.upper():
                for cnpj in KACULA_CNPJ_RE.findall(line):
                    if not only_digits(cnpj).startswith("61186888"):
                        cnpj_atual = only_digits(cnpj)

            if "EAN" in line.upper():
                continue

            sku_match = re.match(r"^(\d{5,6})\b", line)
            if sku_match:
                sku_atual = sku_match.group(1)

            qtd_match = re.search(r"\b\d+,\d{2}\b", line)
            if sku_atual and qtd_match:
                rows.append(_build_row(cnpj_atual, sku_atual, qtd_match.group(0), pedido_atual))
                sku_atual = ""

    if not rows:
        msg = "Layout invalido ou nao reconhecido para Kacula. Verifique se o PDF enviado corresponde ao padrao esperado."
        return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=[], rede="Kacula", erro_layout=msg)

    alertas = []
    if any(not row["cnpj_lido"] for row in rows):
        alertas.append("Kacula: ha itens sem CNPJ identificado")
    if any(not row["numero_pedido_lido"] for row in rows):
        alertas.append("Kacula: ha itens sem numero de pedido identificado")

    return _result(caminho_arquivo=caminho_arquivo, layout_config=layout_config, rows=rows, rede="Kacula", alertas=alertas)
