from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

from parsers_pdf.pdf_utils import build_intermediate_df, clean_text, extract_pages_text_detailed, only_digits
from terminal_logger import get_terminal_logger


terminal_log = get_terminal_logger("pdf_maby")

# Regra de negócio MABY SUPERMERCADOS / SPAL
# - PDF textual com várias lojas no mesmo arquivo.
# - Cada loja/pedido pode ocupar 3 ou 4 páginas; não confiar apenas em "Pagina: 1/3".
# - CNPJ correto é o campo "CNPJ da Empresa" do bloco da loja.
# - Código da primeira coluna é apenas código interno do cliente.
# - Produto oficial para o Robô KOF neste layout é o GTIN/EAN, não o SKU interno.
# - Quantidade final já vem em caixas na coluna "Quanti".
# - Linhas sem GTIN/EAN não entram no modelo; ficam como alerta/itens ignorados.

ITEM_RE = re.compile(
    r"^\s*(?P<codigo_cliente>\d{4,8})\s+"
    r"(?P<ean>\d{8,14})\s+"
    r"(?P<descricao>.+?)\s+"
    r"(?P<qtd>\d+(?:[,.]\d+)?)\s+"
    r"(?P<unidade>[A-Z]{1,5})\s+"
    r"(?P<fator>\d+(?:[,.]\d+)?)\s+"
    r"(?P<valor_un>-?[\d.]+,\d{2,4})\s+"
    r"(?P<valor_cx>-?[\d.]+,\d{2,4})\s+"
    r"(?P<bonif>\d+(?:[,.]\d+)?)\s+"
    r"(?P<valor_total>-?[\d.]+,\d{2})\s*$",
    re.IGNORECASE,
)

# Linha com código interno, descrição e quantidade, porém sem GTIN/EAN.
ITEM_SEM_EAN_RE = re.compile(
    r"^\s*(?P<codigo_cliente>\d{4,8})\s+"
    r"(?P<descricao>[A-Za-zÀ-ÿ].+?)\s+"
    r"(?P<qtd>\d+(?:[,.]\d+)?)\s+"
    r"(?P<unidade>[A-Z]{1,5})\s+"
    r"(?P<fator>\d+(?:[,.]\d+)?)\s+"
    r"(?P<valor_un>-?[\d.]+,\d{2,4})\s+"
    r"(?P<valor_cx>-?[\d.]+,\d{2,4})\s+"
    r"(?P<bonif>\d+(?:[,.]\d+)?)\s+"
    r"(?P<valor_total>-?[\d.]+,\d{2})\s*$",
    re.IGNORECASE,
)

CNPJ_EMPRESA_RE = re.compile(r"CNPJ\s+da\s+Empresa\s*:\s*([0-9./\-]+)", re.IGNORECASE)
CNPJ_COBRANCA_RE = re.compile(r"\bCNPJ\s*:\s*([0-9./\-]+)", re.IGNORECASE)
PEDIDO_FOOTER_RE = re.compile(r"^\s*Pedido\s*:\s*([0-9]{6,20})\b", re.IGNORECASE)
PEDIDO_NUMERO_RE = re.compile(r"N[úu]mero\s+do\s+Pedido\s*:?\s*([0-9]{6,20})?", re.IGNORECASE)
PREVISAO_RE = re.compile(r"Previs[aã]o\s+de\s+Entrega\s*:?\s*([0-9]{2}/[0-9]{2}/[0-9]{4})?", re.IGNORECASE)
LOJA_RE = re.compile(r"^\s*(?P<loja>\d{3})\s+MABY\s+SUPERMERCADOS", re.IGNORECASE)
DATA_RE = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")


def _normalizar_linha(linha: object) -> str:
    return re.sub(r"\s+", " ", str(linha or "")).strip()


def _qtd_text(value: object) -> str:
    texto = str(value or "").strip()
    if not texto:
        return ""
    if "," not in texto and "." in texto:
        try:
            numero = float(texto)
            if numero.is_integer():
                return str(int(numero))
        except Exception:
            pass
    if re.fullmatch(r"\d+,0+", texto):
        return texto.split(",", 1)[0]
    return texto


def _normalizar_cnpj_maby(value: object) -> str:
    """Preserva a chave CNPJ exatamente como o PDF Maby apresenta.

    O layout Maby/SPAL usado pelo cliente traz CNPJ da Empresa com 15 dígitos
    em alguns arquivos, por exemplo 016957746000128. Essa é a chave validada no
    de/para da rede e não pode ser reduzida para 14 dígitos no parser, pois isso
    muda a rastreabilidade e quebra a matrícula esperada.
    """
    return only_digits(value)


def _valor_proxima_linha(linhas: List[str], indice_atual: int, padrao: re.Pattern[str]) -> str:
    for proximo in range(indice_atual + 1, min(indice_atual + 4, len(linhas))):
        candidato = _normalizar_linha(linhas[proximo])
        if not candidato:
            continue
        m = padrao.search(candidato)
        if m:
            return m.group(0)
        # Se encontrou novo rótulo antes do valor, interrompe.
        if re.search(r"Empresa do Pedido|Data de Faturamento|Fornecedor|Forma de Pagamento|Bonifica", candidato, re.I):
            break
    return ""


def _atualizar_contexto_linha(linhas: List[str], idx: int, contexto: Dict[str, str]) -> None:
    linha = _normalizar_linha(linhas[idx])
    if not linha:
        return

    loja_match = LOJA_RE.search(linha)
    if loja_match:
        contexto["codigo_loja"] = only_digits(loja_match.group("loja"))
        contexto["loja"] = f"MABY LOJA {contexto['codigo_loja']}"

    cnpj_match = CNPJ_EMPRESA_RE.search(linha) or CNPJ_COBRANCA_RE.search(linha)
    if cnpj_match:
        cnpj = _normalizar_cnpj_maby(cnpj_match.group(1))
        if cnpj:
            contexto["cnpj"] = cnpj
            contexto["cnpj_base"] = cnpj[:12] if len(cnpj) >= 12 else ""

    pedido_match = PEDIDO_NUMERO_RE.search(linha)
    if pedido_match:
        pedido = only_digits(pedido_match.group(1) or "")
        if not pedido:
            pedido = _valor_proxima_linha(linhas, idx, re.compile(r"\b\d{6,20}\b"))
        if pedido and len(only_digits(pedido)) >= 6:
            contexto["pedido"] = only_digits(pedido)

    footer_match = PEDIDO_FOOTER_RE.search(linha)
    if footer_match:
        pedido_footer = only_digits(footer_match.group(1))
        if len(pedido_footer) >= 6:
            contexto["pedido"] = pedido_footer

    data_match = PREVISAO_RE.search(linha)
    if data_match:
        data = data_match.group(1) or _valor_proxima_linha(linhas, idx, DATA_RE)
        if data:
            contexto["data_entrega"] = data


def _montar_row_item(dados: Dict[str, str], linha: str, pagina: int, linha_idx: int, contexto: Dict[str, str]) -> Dict[str, str]:
    codigo_cliente = only_digits(dados.get("codigo_cliente"))
    ean = only_digits(dados.get("ean"))
    cnpj = contexto.get("cnpj", "")
    codigo_loja = contexto.get("codigo_loja", "")

    # Nesta rede o Robô deve receber EAN no campo de produto. O código interno do
    # PDF fica preservado em codigo_origem_lido/sku_origem_pdf_lido para auditoria.
    return {
        "matricula_lida": "",
        "cnpj_lido": cnpj,
        "cnpj_base_lido": contexto.get("cnpj_base", ""),
        "gln_lido": "",
        "codigo_loja_lido": codigo_loja,
        "codigo_cliente_lido": cnpj or codigo_loja,
        "cod_cliente_lido": codigo_loja,
        "loja_lida": contexto.get("loja", ""),
        "texto_loja_lido": contexto.get("loja", ""),
        "sku_lido": ean,
        "codigo_sku_lido": ean,
        "ean_lido": ean,
        "codigo_origem_lido": codigo_cliente,
        "sku_origem_pdf_lido": codigo_cliente,
        "descricao_lida": clean_text(dados.get("descricao")),
        "quantidade_lida": _qtd_text(dados.get("qtd")),
        "numero_pedido_lido": contexto.get("pedido", ""),
        "data_entrega_lida": contexto.get("data_entrega", ""),
        "pagina_pdf": str(pagina),
        "linha_origem": str(linha_idx),
        "linha_bruta": linha,
        "origem_extracao": "PDF_MABY_DEDICADO_EAN",
        "status_extracao": "OK" if cnpj and ean and dados.get("qtd") and contexto.get("pedido") else "VALIDAR_MABY",
        "alerta_extracao": "" if cnpj and ean and dados.get("qtd") and contexto.get("pedido") else "Maby: conferir CNPJ, pedido, EAN e quantidade; linha extraída por parser dedicado.",
        "modo_rastreabilidade": "NAO",
        "layout_referencia": "MABY_SUPERMERCADOS",
        "confianca_rastreabilidade": "PARSER_DEDICADO",
    }


def _parse_textos(textos: Iterable[str], caminho_arquivo: str, layout_config: dict) -> Tuple[list[dict[str, str]], list[str]]:
    rows: list[dict[str, str]] = []
    alertas: list[str] = []
    contexto: Dict[str, str] = {
        "cnpj": "",
        "cnpj_base": "",
        "pedido": "",
        "data_entrega": "",
        "codigo_loja": "",
        "loja": "",
    }
    linhas_sem_ean = 0
    linhas_nao_interpretadas = 0

    for pagina_idx, texto in enumerate(textos, start=1):
        linhas = [_normalizar_linha(l) for l in str(texto or "").splitlines()]
        itens_pagina = 0
        for linha_idx, linha in enumerate(linhas, start=1):
            if not linha:
                continue

            _atualizar_contexto_linha(linhas, linha_idx - 1, contexto)

            match = ITEM_RE.match(linha)
            if match:
                rows.append(_montar_row_item(match.groupdict(), linha, pagina_idx, linha_idx, contexto))
                itens_pagina += 1
                continue

            if ITEM_SEM_EAN_RE.match(linha):
                linhas_sem_ean += 1
                alertas.append(
                    f"MABY_ITEM_SEM_EAN_BLOQUEADO | pagina={pagina_idx} | linha={linha_idx} | texto={linha[:180]}"
                )
                continue

            if re.match(r"^\s*\d{4,8}\s+\d{8,14}\s+", linha):
                linhas_nao_interpretadas += 1
                alertas.append(
                    f"MABY_ITEM_NAO_INTERPRETADO | pagina={pagina_idx} | linha={linha_idx} | texto={linha[:180]}"
                )

        if itens_pagina == 0:
            terminal_log.info(
                "[MABY] página sem item extraído | arquivo=%s | pagina=%s | cnpj=%s | pedido=%s",
                Path(caminho_arquivo).name,
                pagina_idx,
                contexto.get("cnpj", ""),
                contexto.get("pedido", ""),
            )

    if linhas_sem_ean:
        alertas.append(
            f"MABY: {linhas_sem_ean} linha(s) sem GTIN/EAN foram mantidas fora do modelo e enviadas para alerta/validação."
        )
    if linhas_nao_interpretadas:
        alertas.append(f"MABY: {linhas_nao_interpretadas} linha(s) com aparência de item não foram interpretadas.")
    if rows and any(not row.get("cnpj_lido") for row in rows):
        alertas.append("MABY: há item(ns) sem CNPJ da empresa identificado; conferir PDF/layout.")
    if rows and any(not row.get("numero_pedido_lido") for row in rows):
        alertas.append("MABY: há item(ns) sem número de pedido identificado; conferir PDF/layout.")
    if rows and any(not row.get("ean_lido") for row in rows):
        alertas.append("MABY: há item(ns) sem EAN; não liberar para fila/TXT sem correção manual.")

    return rows, sorted({a for a in alertas if a})


def ler_pdf_maby(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None) -> dict:
    audit = extract_pages_text_detailed(caminho_arquivo)
    rows, alertas = _parse_textos(audit.paginas, caminho_arquivo, layout_config)
    df_intermediario = build_intermediate_df(rows, caminho_arquivo, layout_config.get("nome_layout", "MABY SUPERMERCADOS PDF"))

    alertas_final = sorted({str(a) for a in (alertas + audit.alertas) if str(a).strip()})
    sucesso = not df_intermediario.empty
    terminal_log.info(
        "[MABY] arquivo=%s | paginas=%s | itens=%s | cnpjs=%s | pedidos=%s | eans_vazios=%s | alertas=%s",
        Path(caminho_arquivo).name,
        audit.total_paginas,
        len(df_intermediario),
        sorted({r.get("cnpj_lido", "") for r in rows if r.get("cnpj_lido")})[:20],
        sorted({r.get("numero_pedido_lido", "") for r in rows if r.get("numero_pedido_lido")})[:20],
        sum(1 for r in rows if not r.get("ean_lido")),
        len(alertas_final),
    )
    return {
        "sucesso": sucesso,
        "mensagem": f"Leitura PDF Maby concluida com {len(df_intermediario)} item(ns) com EAN" if sucesso else "Nenhum item com EAN extraido do PDF Maby",
        "df_intermediario": df_intermediario,
        "qtd_linhas_lidas": len(df_intermediario),
        "qtd_itens_extraidos": len(df_intermediario),
        "paginas_pdf_total": audit.total_paginas,
        "paginas_pdf_processadas": audit.paginas_processadas,
        "paginas_pdf_sem_texto": int(sum(1 for a in audit.auditoria if not a.caracteres)),
        "motores_pdf": ", ".join(sorted({a.motor for a in audit.auditoria})),
        "df_auditoria_paginas": audit.auditoria_df(),
        "alertas": alertas_final,
    }
