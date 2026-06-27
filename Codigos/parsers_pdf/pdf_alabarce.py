from __future__ import annotations

import re
from datetime import datetime, timedelta
from typing import Dict, List, Tuple

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, only_digits
from terminal_logger import get_terminal_logger

try:
    import pdfplumber  # type: ignore
except ModuleNotFoundError:
    pdfplumber = None  # type: ignore

terminal_log = get_terminal_logger("pdf_alabarce")

# Exemplo do layout Bluesoft:
# Pedido - 794327 (SPAL INDUSTRIA BRAS.) - Emitido em 13/05/2026
RE_PEDIDO = re.compile(r"Pedido\s*-\s*(\d{4,})", re.I)

# Observações podem vir como:
# LOJA1-CENTRO-... CNPJ.00.203.057/0001-98
# LOJA5-MOGI BERTIOGA-... CNPJ.00.203.057/0007-83
RE_LOJA_CNPJ = re.compile(
    r"LOJA\s*0?(\d{1,2})\s*[-–:]?\s*(.*?)\s*CNPJ\.?\s*([\d./-]{10,25})",
    re.I | re.S,
)

RE_LOJA_CABECALHO = re.compile(r"LJ\s*0?(\d{1,2})", re.I)

# De/para manual reforçado no próprio parser para rastreabilidade da Alabarce.
# O de/para oficial também fica em Cadastros/de_para_clientes.csv.
DE_PARA_MANUAL_ALABARCE = {
    "00203057000198": "700114177",
    "00203057000430": "700207032",
    "00203057000511": "7110079716",
    "00203057000600": "7110073125",
    "00203057000783": "7110436423",
}

NOMES_LOJAS_ALABARCE = {
    "LJ01": "CENTRO",
    "LJ02": "SOCORRO",
    "LJ03": "HENRIQUE PERES",
    "LJ04": "NOVA MOGILAR",
    "LJ05": "MOGI BERTIOGA",
}

# Marcadores de cabeçalho/rodapé/seções que NÃO são itens.
# Importante: não usar tokens genéricos por substring, como "REF" e "CUSTO".
# No Bluesoft Alabarce, muitos produtos começam com "REFRIG." e o código antigo
# pulava páginas inteiras porque encontrava "REF" dentro de "REFRIG.".
IGNORE_TOKENS = (
    "SUBTOTAL",
    "TOTAL DO PEDIDO",
    "TOTAL(QUANT",
    "TOTAL(VALOR",
    "OBSERVAÇÕES",
    "OBSERVACOES",
    "LOCAIS DE ENTREGA",
    "DADOS DE TRANSPORTADORAS",
    "TRANSPORTADORA",
    "TROCAS",
    "TOTAL QUANTIDADE",
    "CUSTO MEDIO",
)


def _eh_linha_controle_alabarce(texto: str) -> bool:
    """Retorna True para linhas de cabeçalho, subtotal, total, rodapé ou seções.

    A validação é propositalmente restrita para não bloquear descrições de produto.
    Exemplo real corrigido: "REFRIG.COCA..." contém as letras "REF", mas é item.
    """
    texto_up = _upper(texto)
    if not texto_up:
        return True

    if any(token in texto_up for token in IGNORE_TOKENS):
        return True

    # Cabeçalho da tabela: pode vir tudo na mesma linha.
    if (
        re.search(r"\bREF\b", texto_up)
        and ("GTIN" in texto_up or "DESCRI" in texto_up)
    ):
        return True

    if "LOJAS" in texto_up and ("GTIN" in texto_up or "DESCRI" in texto_up):
        return True

    # Linhas de identificação/rodapé do PDF.
    if texto_up.startswith("PEDIDO -") or texto_up.startswith("PÁGINA ") or texto_up.startswith("PAGINA "):
        return True
    if texto_up.startswith("BLUESOFT ERP"):
        return True

    return False


def _linha_inicia_bloco_nao_item(texto: str) -> bool:
    """Indica que, daquele ponto em diante na página, acabou a tabela de itens."""
    texto_up = _upper(texto)
    return any(
        marcador in texto_up
        for marcador in (
            "SUBTOTAL",
            "TOTAL DO PEDIDO",
            "OBSERVAÇÕES",
            "OBSERVACOES",
            "LOCAIS DE ENTREGA",
            "DADOS DE TRANSPORTADORAS",
            "TROCAS",
        )
    )


def ensure_pdfplumber():
    if pdfplumber is None:
        raise RuntimeError(
            "pdfplumber não está disponível. Instale com: python -m pip install pdfplumber"
        )


def _norm_text(value) -> str:
    return clean_text(value).replace("\n", " ").strip()


def _upper(value) -> str:
    return _norm_text(value).upper()


def _extract_texto_total(caminho_arquivo: str) -> str:
    ensure_pdfplumber()

    partes = []
    with pdfplumber.open(caminho_arquivo) as pdf:
        for page in pdf.pages:
            partes.append(page.extract_text() or "")
    return "\n".join(partes)


def _extract_pedido(texto_total: str) -> str:
    m = RE_PEDIDO.search(texto_total)
    return clean_text(m.group(1)) if m else ""


def _limpar_nome_loja(texto: str) -> str:
    texto = re.sub(r"\s+", " ", str(texto or "").replace("\n", " ")).strip(" -–:")
    texto = re.sub(r"CNPJ.*$", "", texto, flags=re.I).strip(" -–:")
    return texto.upper()


def _extract_lojas_cnpj(texto_total: str) -> Dict[str, Dict[str, str]]:
    lojas: Dict[str, Dict[str, str]] = {}

    for loja_num, nome, cnpj in RE_LOJA_CNPJ.findall(texto_total):
        loja = f"LJ{int(loja_num):02d}"
        cnpj_digits = only_digits(cnpj).zfill(14)
        nome_limpo = NOMES_LOJAS_ALABARCE.get(loja, "") or _limpar_nome_loja(nome)
        lojas[loja] = {
            "loja": loja,
            "nome": nome_limpo,
            "cnpj": cnpj_digits,
            "matricula_manual": DE_PARA_MANUAL_ALABARCE.get(cnpj_digits, ""),
        }

    # Garante cadastro manual conhecido mesmo quando a linha Observações vier truncada.
    # O parser só usa a loja se a coluna existir no cabeçalho da tabela.
    fallback_cnpj_por_loja = {
        "LJ01": "00203057000198",
        "LJ02": "00203057000430",
        "LJ03": "00203057000600",
        "LJ04": "00203057000511",
        "LJ05": "00203057000783",
    }
    for loja, cnpj in fallback_cnpj_por_loja.items():
        lojas.setdefault(
            loja,
            {
                "loja": loja,
                "nome": NOMES_LOJAS_ALABARCE.get(loja, ""),
                "cnpj": cnpj,
                "matricula_manual": DE_PARA_MANUAL_ALABARCE.get(cnpj, ""),
            },
        )
    return lojas


def _data_d_plus_1() -> str:
    return (datetime.now() + timedelta(days=1)).strftime("%d/%m/%Y")


def _to_int(value) -> int:
    texto = clean_text(value).replace("\n", "").strip()
    if not texto:
        return 0
    # Campo de loja deve ser inteiro. Não aceitar custo/preço com vírgula decimal.
    texto = texto.replace(".", "")
    if re.fullmatch(r"\d+", texto):
        return int(texto)
    if re.fullmatch(r"\d+,0+", texto):
        return int(texto.split(",", 1)[0])
    return 0


def _find_idx(row: List[str], patterns: Tuple[str, ...]) -> int | None:
    for idx, cell in enumerate(row):
        cell_up = _upper(cell)
        if any(pat in cell_up for pat in patterns):
            return idx
    return None


def _parse_loja_label(label: str) -> str:
    m = RE_LOJA_CABECALHO.search(label or "")
    if not m:
        return ""
    return f"LJ{int(m.group(1)):02d}"


def _detectar_mapa_tabela(tabela: List[List[str]]) -> dict | None:
    """Detecta colunas reais do layout Alabarce.

    Ponto crítico corrigido: nem todo pedido traz todas as lojas na tabela.
    Alguns arquivos começam em LJ02 e não possuem a coluna LJ01. Antes o parser
    assumia posição fixa e deslocava as quantidades para lojas erradas.
    """
    for idx, row in enumerate(tabela[:25]):
        row_norm = [_norm_text(c) for c in row]
        row_join = " ".join(row_norm).upper()
        if "REF" not in row_join or "GTIN" not in row_join or "DESCRI" not in row_join:
            continue

        header_1 = row_norm
        header_2 = [_norm_text(c) for c in tabela[idx + 1]] if idx + 1 < len(tabela) else []
        ref_idx = _find_idx(header_1, ("REF",))
        gtin_idx = _find_idx(header_1, ("GTIN",))
        desc_idx = _find_idx(header_1, ("DESCRI",))
        totais_idx = None
        loja_cols: List[Tuple[int, str, str]] = []

        limite_lojas = ref_idx if ref_idx is not None else min(len(header_2), len(header_1))
        for col_idx in range(limite_lojas):
            label = header_2[col_idx] if col_idx < len(header_2) else ""
            loja = _parse_loja_label(label)
            if loja:
                loja_cols.append((col_idx, loja, label))
            elif "TOTAL" in _upper(label) or "TOTAIS" in _upper(label):
                totais_idx = col_idx

        if ref_idx is None or not loja_cols:
            continue

        if totais_idx is None:
            for col_idx in range(limite_lojas):
                label = header_2[col_idx] if col_idx < len(header_2) else ""
                if "TOTAL" in _upper(label) or "TOTAIS" in _upper(label):
                    totais_idx = col_idx
                    break

        return {
            "header_idx": idx,
            "data_inicio_idx": idx + 2,
            "ref_idx": ref_idx,
            "gtin_idx": gtin_idx,
            "desc_idx": desc_idx,
            "totais_idx": totais_idx,
            "loja_cols": loja_cols,
        }
    return None


def _is_item_row(row: List[str], ref_idx: int) -> bool:
    if not row or ref_idx >= len(row):
        return False

    texto = " ".join(_upper(x) for x in row if x)
    if not texto:
        return False

    if _eh_linha_controle_alabarce(texto):
        return False

    ref = only_digits(row[ref_idx])
    if not ref or len(ref) < 4 or len(ref) > 7:
        return False

    return True



def _linha_alabarce_padrao(
    *,
    loja: str,
    qtd: int,
    sku_ref: str,
    gtin_unitario: str,
    descricao: str,
    loja_info: Dict[str, str],
    page_idx: int,
    row_idx: str,
    linha_bruta: str,
    origem_extracao: str,
    alerta_linha: str = "",
) -> Dict[str, object]:
    """Monta a linha intermediária da Alabarce de forma única.

    A Alabarce trafega por SKU/Ref e quantidade em caixaria por loja.
    O GTIN Unitário fica apenas para rastreabilidade e nunca deve bloquear a fila.
    """
    cnpj = str(loja_info.get("cnpj", ""))
    nome_loja = loja_info.get("nome", NOMES_LOJAS_ALABARCE.get(loja, ""))
    return {
        "matricula_lida": str(loja_info.get("matricula_manual", "")),
        "cnpj_lido": cnpj,
        "cnpj_base_lido": cnpj,
        "sku_lido": sku_ref,
        "codigo_sku_lido": sku_ref,
        "ean_lido": "",
        "gtin_unitario_lido": gtin_unitario,
        "descricao_lida": descricao,
        "quantidade_lida": str(qtd),
        "qtd_original": str(qtd),
        "qtd_final": str(qtd),
        "tipo_qtd_original": "CAIXARIA",
        "fator_conversao": "",
        "qtd_convertida": "",
        "status_conversao": "OK SEM CONVERSÃO",
        "tipo_regra_conversao": "SEM_CONVERSAO",
        "regra_aplicada_conversao": "ALABARCE_SEM_CONVERSAO_PARSER | SEM_CONVERSAO | prioridade PARSER",
        "origem_regra_conversao": "PARSER_ALABARCE_BLUESOFT_CAIXARIA",
        "prioridade_regra_conversao": "PARSER",
        "observacao_conversao": "Layout Alabarce já envia quantidade por loja em caixaria; conversão por mapa/fator não se aplica.",
        "layout_usa_conversao": "NAO",
        "numero_pedido_lido": "",
        "data_entrega_lida": "",
        "codigo_loja_lido": loja,
        "loja_lida": loja,
        "texto_loja_lido": f"{loja} - {nome_loja}".strip(" -"),
        "pagina_pdf": str(page_idx),
        "linha_origem": str(row_idx),
        "linha_bruta": linha_bruta,
        "origem_extracao": origem_extracao,
        "layout_usado": "ALABARCE PDF",
        "status_identidade": "OK",
        "motivo_identidade": "OK",
        "alerta_identidade": "",
        "status_extracao": "OK" if not alerta_linha else "VALIDAR",
        "alerta_extracao": alerta_linha,
    }


def _group_words_lines(words: List[dict], tol: float = 3.0) -> List[dict]:
    """Agrupa palavras do pdfplumber em linhas visuais.

    Esse fallback é importante para PDFs Bluesoft da Alabarce com muitas páginas:
    em alguns arquivos, extract_tables() detecta apenas a primeira tabela ou perde
    a continuação nas páginas 2, 3, 4 e 5. A leitura por coordenada mantém as
    colunas de lojas alinhadas e não depende do texto linear perder os espaços.
    """
    linhas: List[dict] = []
    for word in sorted(words or [], key=lambda w: (float(w.get("top", 0)), float(w.get("x0", 0)))):
        top = float(word.get("top", 0))
        for linha in linhas:
            if abs(float(linha["top"]) - top) <= tol:
                n = int(linha.get("n", 1))
                linha["top"] = (float(linha["top"]) * n + top) / (n + 1)
                linha["n"] = n + 1
                linha["words"].append(word)
                break
        else:
            linhas.append({"top": top, "n": 1, "words": [word]})

    for linha in linhas:
        linha["words"].sort(key=lambda w: float(w.get("x0", 0)))
        linha["text"] = " ".join(str(w.get("text", "")) for w in linha["words"]).strip()
    return sorted(linhas, key=lambda l: float(l.get("top", 0)))


def _word_is_int(texto: str) -> bool:
    txt = str(texto or "").replace(".", "").strip()
    return bool(re.fullmatch(r"\d+", txt))


def _word_to_int(texto: str) -> int:
    if not _word_is_int(texto):
        return 0
    try:
        return int(str(texto).replace(".", ""))
    except Exception:
        return 0


def _detectar_cabecalho_por_palavras(page) -> dict | None:
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False) or []
    if not words:
        return None
    linhas = _group_words_lines(words)
    header = None
    for linha in linhas:
        texto_up = _upper(linha.get("text", ""))
        if "REF" in texto_up and "GTIN" in texto_up and "DESCRI" in texto_up:
            header = linha
            break
    if not header:
        return None

    header_top = float(header.get("top", 0))
    zona = [w for w in words if header_top - 6 <= float(w.get("top", 0)) <= header_top + 45]

    loja_cols: List[Tuple[float, str]] = []
    vistos = set()
    for w in zona:
        m = RE_LOJA_CABECALHO.search(str(w.get("text", "")))
        if not m:
            continue
        loja = f"LJ{int(m.group(1)):02d}"
        if loja in vistos:
            continue
        vistos.add(loja)
        loja_cols.append(((float(w.get("x0", 0)) + float(w.get("x1", 0))) / 2.0, loja))
    loja_cols.sort(key=lambda item: item[0])

    total_x = None
    for w in zona:
        txt = _upper(w.get("text", ""))
        if txt in {"TOTAIS", "TOTAL"}:
            total_x = (float(w.get("x0", 0)) + float(w.get("x1", 0))) / 2.0
            break

    ref_words = [w for w in header["words"] if _upper(w.get("text", "")).startswith("REF")]
    gtin_words = [w for w in header["words"] if _upper(w.get("text", "")).startswith("GTIN")]
    desc_words = [w for w in header["words"] if _upper(w.get("text", "")).startswith("DESCRI")]
    if not ref_words or not loja_cols:
        return None

    return {
        "words": words,
        "linhas": linhas,
        "header_top": header_top,
        "loja_cols": loja_cols,
        "total_x": total_x,
        "ref_x": float(ref_words[0].get("x0", 0)),
        "gtin_x": float(gtin_words[0].get("x0", 0)) if gtin_words else float(ref_words[0].get("x0", 0)) + 80,
        "desc_x": float(desc_words[0].get("x0", 0)) if desc_words else float(ref_words[0].get("x0", 0)) + 130,
    }


def _extract_itens_por_palavras(caminho_arquivo: str, lojas_info: Dict[str, Dict[str, str]]) -> tuple[List[Dict[str, object]], List[str], dict]:
    ensure_pdfplumber()
    linhas_saida: List[Dict[str, object]] = []
    alertas: List[str] = []
    auditoria = {
        "paginas_lidas_palavras": 0,
        "paginas_com_cabecalho_palavras": 0,
        "linhas_item_palavras": 0,
        "lojas_detectadas_palavras": set(),
        "itens_por_pagina_palavras": {},
    }

    with pdfplumber.open(caminho_arquivo) as pdf:
        auditoria["paginas_lidas_palavras"] = len(pdf.pages)
        for page_idx, page in enumerate(pdf.pages, start=1):
            mapa = _detectar_cabecalho_por_palavras(page)
            if not mapa:
                auditoria["itens_por_pagina_palavras"][page_idx] = 0
                continue

            auditoria["paginas_com_cabecalho_palavras"] += 1
            loja_cols = list(mapa.get("loja_cols") or [])
            for _, loja in loja_cols:
                auditoria["lojas_detectadas_palavras"].add(loja)
            centros = list(loja_cols)
            if mapa.get("total_x") is not None:
                centros.append((float(mapa["total_x"]), "TOTAIS"))

            ref_x = float(mapa.get("ref_x", 0))
            desc_x = float(mapa.get("desc_x", ref_x + 130))
            header_top = float(mapa.get("header_top", 0))
            itens_pagina = 0

            for line_index, linha in enumerate(mapa.get("linhas") or [], start=1):
                top = float(linha.get("top", 0))
                texto_linha = str(linha.get("text", ""))
                texto_up = _upper(texto_linha)
                if top < header_top + 25:
                    continue
                # Ao encontrar subtotal/total/observações, encerra a leitura da área de itens
                # desta página para não transformar totais ou locais de entrega em SKU.
                if _linha_inicia_bloco_nao_item(texto_linha):
                    break
                if _eh_linha_controle_alabarce(texto_linha):
                    continue
                if not texto_linha:
                    continue

                sku_word = None
                sku_ref = ""
                for w in linha.get("words") or []:
                    d = only_digits(w.get("text", ""))
                    x0 = float(w.get("x0", 0))
                    if 4 <= len(d) <= 7 and (ref_x - 25) <= x0 <= (ref_x + 90):
                        sku_word = w
                        sku_ref = d
                        break
                if not sku_word or not sku_ref:
                    continue

                qtd_por_loja: Dict[str, int] = {}
                total_lido = 0
                for w in linha.get("words") or []:
                    if float(w.get("x1", 0)) >= ref_x - 5:
                        continue
                    if not _word_is_int(w.get("text", "")):
                        continue
                    qtd = _word_to_int(w.get("text", ""))
                    if qtd <= 0:
                        continue
                    centro_x = (float(w.get("x0", 0)) + float(w.get("x1", 0))) / 2.0
                    if not centros:
                        continue
                    alvo_x, alvo = min(centros, key=lambda item: abs(float(item[0]) - centro_x))
                    if abs(float(alvo_x) - centro_x) > 24:
                        continue
                    if alvo == "TOTAIS":
                        total_lido = qtd
                    else:
                        qtd_por_loja[alvo] = qtd

                if not qtd_por_loja:
                    continue

                # GTIN é apenas rastreio. Pode vir quebrado em duas linhas; não bloqueia.
                gtin_unitario = ""
                for w in linha.get("words") or []:
                    x0 = float(w.get("x0", 0))
                    if x0 > float(sku_word.get("x1", 0)) and x0 < desc_x:
                        dig = only_digits(w.get("text", ""))
                        if len(dig) >= 6:
                            gtin_unitario = dig
                            break

                descricao_words = [
                    str(w.get("text", ""))
                    for w in linha.get("words") or []
                    if float(w.get("x0", 0)) >= desc_x
                ]
                descricao = _norm_text(" ".join(descricao_words))

                # Defesa adicional: item real precisa ter descrição textual.
                # Isso evita ler linhas numéricas de subtotal/total como se fossem SKU.
                if not re.search(r"[A-ZÁÉÍÓÚÂÊÔÃÕÇ]", descricao.upper()):
                    continue

                total_calc = sum(qtd_por_loja.values())
                alerta_linha = ""
                if total_lido and total_lido != total_calc:
                    alerta_linha = f"ALABARCE_TOTAL_DIVERGENTE: soma_lojas={total_calc} total_pdf={total_lido}"
                    alertas.append(f"Página {page_idx} linha visual {line_index}: {alerta_linha}")

                for loja, qtd in qtd_por_loja.items():
                    loja_info = lojas_info.get(loja, {})
                    if not loja_info.get("cnpj"):
                        alertas.append(f"CNPJ não identificado para {loja}")
                    linhas_saida.append(
                        _linha_alabarce_padrao(
                            loja=loja,
                            qtd=qtd,
                            sku_ref=sku_ref,
                            gtin_unitario=gtin_unitario,
                            descricao=descricao,
                            loja_info=loja_info,
                            page_idx=page_idx,
                            row_idx=f"palavra:{line_index}",
                            linha_bruta=texto_linha,
                            origem_extracao="pdfplumber.extract_words.coordenadas_todas_paginas",
                            alerta_linha=alerta_linha,
                        )
                    )
                itens_pagina += 1
                auditoria["linhas_item_palavras"] += 1
            auditoria["itens_por_pagina_palavras"][page_idx] = itens_pagina

    auditoria["lojas_detectadas_palavras"] = sorted(auditoria["lojas_detectadas_palavras"])
    return linhas_saida, alertas, auditoria


def _deduplicar_linhas_alabarce(linhas: List[Dict[str, object]]) -> List[Dict[str, object]]:
    saida: List[Dict[str, object]] = []
    vistos = set()
    for linha in linhas:
        # Deduplicação entre extract_tables e extract_words.
        # A descrição pode variar porque a leitura por palavras também captura preço/total
        # na mesma linha; por isso a chave segura usa página + loja + SKU + quantidade.
        chave = (
            str(linha.get("pagina_pdf", "")),
            str(linha.get("codigo_loja_lido", "")),
            str(linha.get("sku_lido", "")),
            str(linha.get("quantidade_lida", "")),
        )
        if chave in vistos:
            continue
        vistos.add(chave)
        saida.append(linha)
    return saida


def _extract_itens_por_tabela(caminho_arquivo: str, lojas_info: Dict[str, Dict[str, str]]) -> tuple[List[Dict[str, object]], List[str], dict]:
    ensure_pdfplumber()

    linhas_saida: List[Dict[str, object]] = []
    alertas: List[str] = []
    auditoria = {
        "paginas_lidas": 0,
        "tabelas_lidas": 0,
        "tabelas_alabarce": 0,
        "linhas_tabela_lidas": 0,
        "lojas_detectadas_tabelas": set(),
    }
    linhas_processadas = set()

    with pdfplumber.open(caminho_arquivo) as pdf:
        auditoria["paginas_lidas"] = len(pdf.pages)
        for page_idx, page in enumerate(pdf.pages, start=1):
            tabelas = page.extract_tables() or []
            auditoria["tabelas_lidas"] += len(tabelas)

            for tabela in tabelas:
                tabela = [[clean_text(c) for c in (row or [])] for row in tabela if row]
                mapa = _detectar_mapa_tabela(tabela)
                if not mapa:
                    continue

                auditoria["tabelas_alabarce"] += 1
                ref_idx = int(mapa["ref_idx"])
                gtin_idx = mapa.get("gtin_idx")
                desc_idx = mapa.get("desc_idx")
                totais_idx = mapa.get("totais_idx")
                loja_cols = list(mapa.get("loja_cols") or [])

                for _, loja, _ in loja_cols:
                    auditoria["lojas_detectadas_tabelas"].add(loja)

                for row_idx, row in enumerate(tabela[int(mapa["data_inicio_idx"]):], start=int(mapa["data_inicio_idx"]) + 1):
                    auditoria["linhas_tabela_lidas"] += 1
                    if not _is_item_row(row, ref_idx):
                        continue

                    sku_ref = only_digits(row[ref_idx])
                    # Alabarce é layout SKU/Ref. O campo GTIN Unitário é apenas informativo.
                    # Não alimentar ean_lido para não bloquear itens com GTIN curto/atípico
                    # (ex.: Monster 70847034803). A fila KOF deve usar o SKU da coluna Ref.
                    gtin_unitario = only_digits(row[gtin_idx]) if isinstance(gtin_idx, int) and gtin_idx < len(row) else ""
                    ean = ""
                    descricao = _norm_text(row[desc_idx]) if isinstance(desc_idx, int) and desc_idx < len(row) else ""

                    qtd_por_loja: Dict[str, int] = {}
                    for col_idx, loja, label in loja_cols:
                        qtd = _to_int(row[col_idx] if col_idx < len(row) else "")
                        if qtd > 0:
                            qtd_por_loja[loja] = qtd

                    if not qtd_por_loja:
                        continue

                    total_lido = _to_int(row[totais_idx] if isinstance(totais_idx, int) and totais_idx < len(row) else "")
                    total_calc = sum(qtd_por_loja.values())
                    alerta_linha = ""
                    if total_lido and total_lido != total_calc:
                        alerta_linha = f"ALABARCE_TOTAL_DIVERGENTE: soma_lojas={total_calc} total_pdf={total_lido}"
                        alertas.append(f"Página {page_idx} linha {row_idx}: {alerta_linha}")

                    for loja, qtd in qtd_por_loja.items():
                        loja_info = lojas_info.get(loja, {})
                        cnpj = str(loja_info.get("cnpj", ""))
                        if not cnpj:
                            alertas.append(f"CNPJ não identificado para {loja}")

                        chave = (page_idx, row_idx, loja, sku_ref, qtd)
                        if chave in linhas_processadas:
                            continue
                        linhas_processadas.add(chave)

                        linhas_saida.append(
                            {
                                "matricula_lida": str(loja_info.get("matricula_manual", "")),
                                "cnpj_lido": cnpj,
                                "cnpj_base_lido": cnpj,
                                "sku_lido": sku_ref,
                                "codigo_sku_lido": sku_ref,
                                # Pedido Alabarce trafega por SKU/Ref. GTIN fica em coluna própria de rastreio.
                                "ean_lido": "",
                                "gtin_unitario_lido": gtin_unitario,
                                "descricao_lida": descricao,
                                "quantidade_lida": str(qtd),
                                "qtd_original": str(qtd),
                                "qtd_final": str(qtd),
                                # Alabarce/Bluesoft já trafega a quantidade em caixaria por loja.
                                # Não aplicar mapa de produtos, fator ou divisão por apresentação.
                                "tipo_qtd_original": "CAIXARIA",
                                "fator_conversao": "",
                                "qtd_convertida": "",
                                "status_conversao": "OK SEM CONVERSÃO",
                                "tipo_regra_conversao": "SEM_CONVERSAO",
                                "regra_aplicada_conversao": "ALABARCE_SEM_CONVERSAO_PARSER | SEM_CONVERSAO | prioridade PARSER",
                                "origem_regra_conversao": "PARSER_ALABARCE_BLUESOFT_CAIXARIA",
                                "prioridade_regra_conversao": "PARSER",
                                "observacao_conversao": "Layout Alabarce já envia quantidade por loja em caixaria; conversão por mapa/fator não se aplica.",
                                "layout_usa_conversao": "NAO",
                                "numero_pedido_lido": "",  # preenchido depois
                                "data_entrega_lida": "",  # preenchido depois
                                "codigo_loja_lido": loja,
                                "loja_lida": loja,
                                "texto_loja_lido": f"{loja} - {loja_info.get('nome', NOMES_LOJAS_ALABARCE.get(loja, ''))}".strip(" -"),
                                "pagina_pdf": str(page_idx),
                                "linha_origem": str(row_idx),
                                "linha_bruta": " | ".join(_norm_text(c) for c in row),
                                "origem_extracao": "pdfplumber.extract_tables.dynamic_lojas",
                                "layout_usado": "ALABARCE PDF",
                                "status_identidade": "OK",
                                "motivo_identidade": "OK",
                                "alerta_identidade": "",
                                "status_extracao": "OK" if not alerta_linha else "VALIDAR",
                                "alerta_extracao": alerta_linha,
                            }
                        )

    # Fallback corporativo: leitura por coordenadas em TODAS as páginas.
    # Isso corrige PDFs Alabarce longos em que o extract_tables() reconhece só a primeira página
    # ou perde tabelas de continuação. A deduplicação mantém o resultado sem repetir itens quando
    # os dois motores leem a mesma linha.
    try:
        linhas_palavras, alertas_palavras, auditoria_palavras = _extract_itens_por_palavras(caminho_arquivo, lojas_info)
        alertas.extend(alertas_palavras)
        paginas_tabela = sorted({str(l.get("pagina_pdf", "")) for l in linhas_saida if l.get("pagina_pdf")})
        paginas_palavras = sorted({str(l.get("pagina_pdf", "")) for l in linhas_palavras if l.get("pagina_pdf")})
        linhas_saida = _deduplicar_linhas_alabarce(linhas_saida + linhas_palavras)
        auditoria["fallback_palavras_ativo"] = True
        auditoria["paginas_com_itens_tabela"] = paginas_tabela
        auditoria["paginas_com_itens_palavras"] = paginas_palavras
        auditoria["linhas_palavras_lidas"] = len(linhas_palavras)
        auditoria["itens_por_pagina_palavras"] = auditoria_palavras.get("itens_por_pagina_palavras", {})
        for loja in auditoria_palavras.get("lojas_detectadas_palavras", []):
            auditoria["lojas_detectadas_tabelas"].add(loja)
    except Exception as exc:
        auditoria["fallback_palavras_ativo"] = False
        auditoria["fallback_palavras_erro"] = str(exc)
        alertas.append(f"Fallback por coordenadas Alabarce falhou: {exc}")

    auditoria["lojas_detectadas_tabelas"] = sorted(auditoria["lojas_detectadas_tabelas"])
    return linhas_saida, alertas, auditoria


def ler_pdf_alabarce(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    texto_total = _extract_texto_total(caminho_arquivo)

    numero_pedido = _extract_pedido(texto_total)
    lojas_info = _extract_lojas_cnpj(texto_total)
    data_entrega = _data_d_plus_1()

    alertas: List[str] = []

    if not numero_pedido:
        alertas.append("Número do pedido não identificado no cabeçalho 'Pedido -'.")

    if not lojas_info:
        alertas.append("Nenhum CNPJ de loja identificado no bloco Observações; usado fallback manual Alabarce quando aplicável.")

    linhas_saida, alertas_tabela, auditoria = _extract_itens_por_tabela(caminho_arquivo, lojas_info)
    alertas.extend(alertas_tabela)

    for linha in linhas_saida:
        linha["numero_pedido_lido"] = numero_pedido
        linha["data_entrega_lida"] = data_entrega

    if not linhas_saida:
        alertas.append("Nenhum item reconhecido nas tabelas Alabarce. Conferir se o PDF veio em imagem ou layout diferente.")

    lojas_detectadas = auditoria.get("lojas_detectadas_tabelas", [])
    lojas_com_cnpj = sorted(lojas_info.keys())
    rastreabilidade_msg = (
        "ALABARCE_RASTREABILIDADE: "
        f"pedido={numero_pedido or 'NAO_IDENTIFICADO'} | "
        f"lojas_tabela={','.join(lojas_detectadas) if lojas_detectadas else 'nenhuma'} | "
        f"lojas_cnpj={','.join(lojas_com_cnpj) if lojas_com_cnpj else 'nenhuma'} | "
        f"tabelas_alabarce={auditoria.get('tabelas_alabarce', 0)} | "
        f"paginas={auditoria.get('paginas_lidas', 0)} | "
        "conversao=NAO_APLICAVEL_CAIXARIA"
    )
    auditoria["rastreabilidade_layout"] = rastreabilidade_msg
    terminal_log.info("[ALABARCE] %s", rastreabilidade_msg)

    df_intermediario = build_intermediate_df(
        linhas_saida,
        caminho_arquivo,
        layout_config.get("nome_layout", ""),
    )

    sucesso = not df_intermediario.empty
    mensagem = (
        f"Leitura PDF Alabarce concluída com {len(df_intermediario)} linha(s) | "
        f"pedido={numero_pedido or 'NAO_IDENTIFICADO'} | lojas={','.join(lojas_detectadas) if lojas_detectadas else 'nenhuma'}"
        if sucesso
        else "Nenhuma linha válida foi extraída do PDF Alabarce"
    )

    terminal_log.info("[ALABARCE] %s", mensagem)

    return {
        "sucesso": sucesso,
        "mensagem": mensagem,
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": len(df_intermediario),
        "alertas": sorted({a for a in alertas if str(a).strip()}),
        "auditoria": auditoria,
    }
