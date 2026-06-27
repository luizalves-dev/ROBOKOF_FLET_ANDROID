from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
import re
from difflib import SequenceMatcher
from typing import Dict, Iterable, List, Optional

import pandas as pd

import cadastro_service
import config
import depara_clientes_service
from terminal_logger import get_terminal_logger

try:
    from parsers_pdf.pdf_utils import extract_pages_text_detailed
except Exception:  # pragma: no cover
    extract_pages_text_detailed = None  # type: ignore


terminal_log = get_terminal_logger("rastreabilidade")

ARQUIVO_CADASTRO = "rastreabilidade_layouts.csv"
LIMIAR_AUTO_PADRAO = 75
LIMIAR_SUGESTAO_PADRAO = 55
MAX_CARACTERES_AMOSTRA = 50000

STOPWORDS = {
    "PDF", "EXCEL", "PADRAO", "PADRÃO", "LAYOUT", "MATRICIAL", "PEDIDO", "PEDIDOS",
    "CLIENTE", "CLIENTES", "ROBO", "ROBÔ", "KOF", "IMPORTACAO", "IMPORTAÇÃO",
}

GENERICOS_PDF = [
    "CNPJ", "PEDIDO", "PEDIDO DE COMPRA", "FORNECEDOR", "PRODUTO", "DESCRICAO", "DESCRIÇÃO",
    "SKU", "COD", "COD FORN", "COD.FORN", "REFERENCIA", "REFERÊNCIA", "QTDE", "QTD", "QUANTIDADE",
    "PREVISAO", "PREVISÃO", "ENTREGA", "UNIDADE", "EMBALAGEM", "TOTAL",
]

GENERICOS_EXCEL = [
    "MATRICULA", "MATRÍCULA", "SKU", "QTD", "QTDE", "QUANTIDADE", "PEDIDO", "ENTREGA", "DATA",
    "CNPJ", "EAN", "CODIGO", "CÓDIGO", "PRODUTO", "DESCRICAO", "DESCRIÇÃO",
]

# Palavras que indicam que o layout depende de conversao unidade-caixaria ou de mapa de produto.
# Estes layouts podem ate aparecer na rastreabilidade, mas nao sao aplicados automaticamente.
MARCADORES_CONVERSAO = [
    "CONVERSAO", "CONVERSÃO", "UNIDADE PARA CAIXA", "UNIDADE-CAIXA", "UNIDADE -> CAIXA",
    "QUANT_CONVERTIDA", "MAPA DE PRODUTOS", "DIVISAO", "DIVISÃO", "FATOR", "EAN->SKU",
]


@dataclass
class ResultadoRastreabilidade:
    sucesso: bool
    aplicar_automaticamente: bool
    layout_id_referencia: str = ""
    nome_layout_referencia: str = ""
    tipo_arquivo: str = ""
    confianca: int = 0
    limiar_auto: int = LIMIAR_AUTO_PADRAO
    limiar_sugestao: int = LIMIAR_SUGESTAO_PADRAO
    motivo: str = ""
    tokens_encontrados: str = ""
    status: str = "SEM_SUGESTAO"
    observacao: str = ""
    arquivo: str = ""

    def to_dict(self) -> Dict[str, object]:
        return asdict(self)


def _normalizar_texto(valor: object) -> str:
    texto = str(valor or "").upper()
    tabela = str.maketrans(
        "ÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇÑ",
        "AAAAAEEEEIIIIOOOOOUUUUCN",
    )
    texto = texto.translate(tabela)
    texto = re.sub(r"[^A-Z0-9]+", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def _tokenizar(valor: object) -> List[str]:
    texto = _normalizar_texto(valor)
    tokens = [t for t in texto.split() if len(t) >= 3 and t not in STOPWORDS]
    return list(dict.fromkeys(tokens))


def _split_tokens(valor: object) -> List[str]:
    texto = str(valor or "")
    partes = re.split(r"[|,;\n]+", texto)
    tokens = []
    for parte in partes:
        parte_norm = _normalizar_texto(parte)
        if parte_norm:
            tokens.append(parte_norm)
    return list(dict.fromkeys(tokens))


def _contem(texto_norm: str, token_norm: str) -> bool:
    if not token_norm:
        return False
    if " " in token_norm:
        return token_norm in texto_norm
    return re.search(rf"\b{re.escape(token_norm)}\b", texto_norm) is not None


def _ler_texto_pdf(caminho: str) -> tuple[str, pd.DataFrame]:
    if extract_pages_text_detailed is None:
        return "", pd.DataFrame([{"status": "SEM_LEITOR_PDF", "mensagem": "Leitor detalhado de PDF indisponivel"}])
    try:
        resultado = extract_pages_text_detailed(caminho)
        texto = "\n".join(resultado.paginas)[:MAX_CARACTERES_AMOSTRA]
        return texto, resultado.auditoria_df()
    except Exception as exc:
        terminal_log.warning("[RASTREABILIDADE] Falha ao ler amostra PDF %s: %s", caminho, exc)
        return "", pd.DataFrame([{"status": "ERRO_LEITURA_AMOSTRA", "mensagem": str(exc)}])


def _ler_texto_excel(caminho: str) -> tuple[str, pd.DataFrame]:
    try:
        planilhas = pd.read_excel(caminho, sheet_name=None, header=None, dtype=str, nrows=80)
        partes = []
        auditoria = []
        for nome_aba, df in planilhas.items():
            if df is None or df.dropna(how="all").empty:
                auditoria.append({"aba": nome_aba, "linhas_amostra": 0, "colunas_amostra": 0, "status": "VAZIA"})
                continue
            df = df.fillna("")
            partes.append(str(nome_aba))
            partes.append(" ".join(str(c) for c in df.columns))
            partes.append("\n".join(" | ".join(str(v) for v in row) for row in df.head(80).values.tolist()))
            auditoria.append({"aba": nome_aba, "linhas_amostra": len(df), "colunas_amostra": len(df.columns), "status": "OK"})
        return "\n".join(partes)[:MAX_CARACTERES_AMOSTRA], pd.DataFrame(auditoria)
    except Exception as exc:
        terminal_log.warning("[RASTREABILIDADE] Falha ao ler amostra Excel %s: %s", caminho, exc)
        return "", pd.DataFrame([{"status": "ERRO_LEITURA_AMOSTRA", "mensagem": str(exc)}])


def extrair_amostra_arquivo(caminho: str, tipo_arquivo: str) -> tuple[str, pd.DataFrame]:
    tipo = str(tipo_arquivo or "").upper()
    if tipo == "PDF":
        return _ler_texto_pdf(caminho)
    if tipo == "EXCEL":
        return _ler_texto_excel(caminho)
    return "", pd.DataFrame([{"status": "TIPO_NAO_SUPORTADO", "mensagem": tipo}])


def _layout_bloqueado_conversao(row: Dict[str, str]) -> bool:
    if str(row.get("ativo", "")).strip() != "1":
        return True
    texto = _normalizar_texto(" ".join(str(row.get(c, "")) for c in ["nome_layout", "observacoes", "regra_data_entrega"]))

    # Bloqueio explícito para layouts fora do escopo atual.
    if "REDE ITALO" in texto:
        return True

    # Cuidado: muitos layouts novos trazem a frase "sem conversão unidade-caixa"
    # na observação. Isso NÃO pode bloquear a rastreabilidade. Só bloqueamos
    # quando há dependência real de conversão/mapa/fator.
    if any(frase in texto for frase in [
        "SEM CONVERSAO",
        "SEM CONVERSAO UNIDADE",
        "SEM DIVISAO",
        "SEM DIVIDIR",
        "SEM CONVERTER",
    ]):
        texto_limpo = texto
        for frase in ["SEM CONVERSAO", "SEM DIVISAO", "SEM DIVIDIR", "SEM CONVERTER"]:
            texto_limpo = texto_limpo.replace(frase, "")
    else:
        texto_limpo = texto

    bloqueios_reais = [
        "QUANT CONVERTIDA",
        "MAPA DE PRODUTOS",
        "UNIDADE PARA CAIXA",
        "UNIDADE CAIXA COM CONVERSAO",
        "COM CONVERSAO",
        "DEPENDE DE CONVERSAO",
        "CONVERSAO POR SKU",
        "CONVERSAO POR EAN",
        "DIVISAO POR",
        "DIVIDIR QUANTIDADE",
        "FATOR",
    ]
    return any(_normalizar_texto(marcador) in texto_limpo for marcador in bloqueios_reais)


def _carregar_cadastro_rastreabilidade() -> pd.DataFrame:
    caminho = config.CADASTROS_DIR / ARQUIVO_CADASTRO
    if caminho.exists():
        try:
            return pd.read_csv(caminho, dtype=str, encoding="utf-8-sig").fillna("")
        except UnicodeDecodeError:
            return pd.read_csv(caminho, dtype=str, encoding="latin-1").fillna("")
    return pd.DataFrame()


def _gerar_cadastro_padrao() -> pd.DataFrame:
    layouts = cadastro_service.carregar_layouts().fillna("")
    linhas = []
    for _, row in layouts.iterrows():
        dados = row.to_dict()
        if str(dados.get("ativo", "")).strip() != "1":
            continue
        tipo = str(dados.get("tipo_arquivo", "")).upper().strip()
        if tipo not in {"PDF", "EXCEL"}:
            continue
        nome = str(dados.get("nome_layout", ""))
        palavras_nome = [t for t in _tokenizar(nome) if t not in {"PDF", "EXCEL"}]
        genericos = GENERICOS_PDF if tipo == "PDF" else GENERICOS_EXCEL
        obs = str(dados.get("observacoes", ""))
        apoio = list(dict.fromkeys(genericos + _tokenizar(obs)))
        bloqueado = _layout_bloqueado_conversao(dados)
        linhas.append(
            {
                "layout_id_referencia": str(dados.get("layout_id", "")),
                "nome_layout_referencia": nome,
                "tipo_arquivo": tipo,
                "ativo_rastreabilidade": "0" if bloqueado else "1",
                "limiar_auto": str(LIMIAR_AUTO_PADRAO),
                "limiar_sugestao": str(LIMIAR_SUGESTAO_PADRAO),
                "palavras_chave_fortes": "|".join(palavras_nome),
                "palavras_chave_apoio": "|".join(apoio[:80]),
                "observacoes_rastreabilidade": (
                    "Bloqueado para rastreabilidade automatica por conversao/unidade ou layout inativo."
                    if bloqueado else
                    "Layout referencia para arquivos parecidos. Usar somente com Excel de validacao e conferencia manual."
                ),
            }
        )
    return pd.DataFrame(linhas)


def carregar_regras_rastreabilidade() -> pd.DataFrame:
    cadastro = _carregar_cadastro_rastreabilidade()
    if cadastro.empty:
        cadastro = _gerar_cadastro_padrao()
    colunas = [
        "layout_id_referencia", "nome_layout_referencia", "tipo_arquivo", "ativo_rastreabilidade",
        "limiar_auto", "limiar_sugestao", "palavras_chave_fortes", "palavras_chave_apoio",
        "observacoes_rastreabilidade",
    ]
    for col in colunas:
        if col not in cadastro.columns:
            cadastro[col] = ""
    return cadastro[colunas].fillna("")


def _pontuar_candidato(texto_norm: str, nome_arquivo_norm: str, regra: Dict[str, str]) -> tuple[int, List[str], str]:
    fortes = _split_tokens(regra.get("palavras_chave_fortes", ""))
    apoio = _split_tokens(regra.get("palavras_chave_apoio", ""))
    nome_layout_norm = _normalizar_texto(regra.get("nome_layout_referencia", ""))

    fortes_encontradas = [t for t in fortes if _contem(texto_norm, t) or _contem(nome_arquivo_norm, t)]
    apoio_encontradas = [t for t in apoio if _contem(texto_norm, t)]

    score_forte = 0
    if fortes:
        score_forte = min(45, round(45 * len(fortes_encontradas) / max(1, len(fortes))))

    score_apoio = min(38, len(apoio_encontradas) * 4)
    score_nome_arquivo = 0
    for token in fortes:
        if _contem(nome_arquivo_norm, token):
            score_nome_arquivo += 8
    score_nome_arquivo = min(14, score_nome_arquivo)

    similaridade_nome = int(SequenceMatcher(None, nome_arquivo_norm, nome_layout_norm).ratio() * 12) if nome_arquivo_norm else 0
    score = min(100, score_forte + score_apoio + score_nome_arquivo + similaridade_nome)

    motivo = (
        f"fortes={len(fortes_encontradas)}/{len(fortes)}; "
        f"apoio={len(apoio_encontradas)}; arquivo={score_nome_arquivo}; similaridade_nome={similaridade_nome}"
    )
    return score, list(dict.fromkeys(fortes_encontradas + apoio_encontradas))[:30], motivo




def _somente_digitos(valor: object) -> str:
    return re.sub(r"\D+", "", str(valor or ""))


def _extrair_identificadores_cliente(texto: str, nome_arquivo: str = "") -> List[str]:
    """Extrai chaves candidatas para de/para sem depender do layout.

    A rastreabilidade usa estes identificadores apenas como atalho seguro:
    CNPJ completo, CNPJ-base, GLN/texto de loja e matrículas já cadastradas.
    A aplicação automática só ocorre se o de/para localizar uma rede e se essa
    rede possuir layout ativo do mesmo tipo de arquivo.
    """
    bruto = f"{texto or ''}\n{nome_arquivo or ''}"
    candidatos: List[str] = []

    def adicionar(valor: object) -> None:
        digitos = _somente_digitos(valor)
        if not digitos:
            return
        # Evita poluir com datas, páginas, códigos internos e quantidades pequenas.
        # Ex.: "Pagina 000001" não pode virar chave "1" e acionar LJ01 de outra rede.
        if len(digitos) < 7 or len(digitos) > 18:
            return
        if len(digitos) <= 7 and digitos.startswith("0"):
            return
        if digitos not in candidatos:
            candidatos.append(digitos)

    # CNPJ com pontuação: 12.345.678/0001-99
    for match in re.finditer(r"\b\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}\b", bruto):
        cnpj = _somente_digitos(match.group(0))
        adicionar(cnpj)
        if len(cnpj) >= 14:
            adicionar(cnpj[:12])

    # Caso Coopercica e similares: "CNPJ -08 50.974.732/0009".
    for match in re.finditer(r"CNPJ\s*[-:]?\s*(\d{2})\s+(\d{2}\.?\d{3}\.?\d{3}/?\d{4})", bruto, flags=re.IGNORECASE):
        dv = _somente_digitos(match.group(1))
        base = _somente_digitos(match.group(2))
        if len(base) == 12 and len(dv) == 2:
            adicionar(base + dv)
            adicionar(base)

    # Sequências numéricas livres. São candidatas; o de/para decide se conhece.
    # Inclui: CNPJ sem máscara (14), GLN/texto loja (13), CNPJ-base (12), matrículas 7-10+.
    for match in re.finditer(r"(?<!\d)\d{6,18}(?!\d)", bruto):
        valor = match.group(0)
        adicionar(valor)
        if len(valor) >= 14:
            adicionar(valor[:12])
        if len(valor) == 13:
            adicionar(valor[:12])

    return candidatos[:120]


def _layout_ativo_por_rede(rede: str, tipo: str) -> Optional[Dict[str, str]]:
    rede_norm = _normalizar_texto(rede)
    tipo_norm = str(tipo or "").upper().strip()
    if not rede_norm or tipo_norm not in {"PDF", "EXCEL"}:
        return None

    try:
        layouts = cadastro_service.carregar_layouts().fillna("")
    except Exception:
        terminal_log.exception("[RASTREABILIDADE] Falha ao carregar layouts para resolver rede por de/para.")
        return None

    candidatos: List[tuple[int, Dict[str, str]]] = []
    for _, row in layouts.iterrows():
        dados = {str(k): str(v or "") for k, v in row.to_dict().items()}
        if str(dados.get("ativo", "")).strip() != "1":
            continue
        if str(dados.get("tipo_arquivo", "")).upper().strip() != tipo_norm:
            continue
        if _layout_bloqueado_conversao(dados):
            continue
        nome_norm = _normalizar_texto(dados.get("nome_layout", ""))
        if not nome_norm:
            continue
        score = 0
        if nome_norm == rede_norm:
            score = 100
        elif rede_norm in nome_norm or nome_norm in rede_norm:
            score = 90
        else:
            tokens_rede = set(rede_norm.split())
            tokens_nome = set(nome_norm.split()) - {"PDF", "EXCEL", "LAYOUT"}
            inter = tokens_rede & tokens_nome
            if inter:
                score = 50 + len(inter) * 5
        if score:
            candidatos.append((score, dados))

    if not candidatos:
        return None
    candidatos.sort(key=lambda item: item[0], reverse=True)
    return candidatos[0][1]




def _layout_ativo_por_nome_exato(nome_layout_desejado: str, tipo: str) -> Optional[Dict[str, str]]:
    """Busca layout ativo por nome exato/contido, sem depender de de/para."""
    desejado = _normalizar_texto(nome_layout_desejado)
    tipo_norm = str(tipo or "").upper().strip()
    if not desejado or tipo_norm not in {"PDF", "EXCEL"}:
        return None
    try:
        layouts = cadastro_service.carregar_layouts().fillna("")
    except Exception:
        terminal_log.exception("[RASTREABILIDADE] Falha ao carregar layouts para assinatura dedicada.")
        return None
    candidatos: List[tuple[int, Dict[str, str]]] = []
    for _, row in layouts.iterrows():
        dados = {str(k): str(v or "") for k, v in row.to_dict().items()}
        if str(dados.get("ativo", "")).strip() != "1":
            continue
        if str(dados.get("tipo_arquivo", "")).upper().strip() != tipo_norm:
            continue
        nome_norm = _normalizar_texto(dados.get("nome_layout", ""))
        if not nome_norm:
            continue
        score = 0
        if nome_norm == desejado:
            score = 100
        elif desejado in nome_norm or nome_norm in desejado:
            score = 90
        if score:
            candidatos.append((score, dados))
    if not candidatos:
        return None
    candidatos.sort(key=lambda item: item[0], reverse=True)
    return candidatos[0][1]


def _parece_excel_rede_vip(texto: str, arquivo_nome: str = "") -> bool:
    """Assinatura forte para Excel matricial da Rede VIP/Coca.

    O arquivo pode chegar como coca.xlsx, sem o nome VIP. A estrutura segura é:
    linha de matrículas 714..., cabeçalho SKU/DESCRIÇÃO/CUSTO/quantidade e lojas
    em colunas. Essa assinatura evita cair no fallback genérico de Excel.
    """
    texto_norm = _normalizar_texto(texto)
    bruto = f"{texto or ''}\n{arquivo_nome or ''}"
    matriculas = set(re.findall(r"(?<!\d)714\d{7}(?!\d)", bruto))
    lojas_vip = [
        "AEROPORTO", "ALICAR", "ANITA", "BRINO", "BELA VISTA",
        "INTERCAP", "MOINHOS", "PAINEIRA", "PASQUALINI",
        "PLANETARIO", "PLANETARIO", "RAMIRO", "PORTO", "VIP",
    ]
    lojas_encontradas = [loja for loja in lojas_vip if loja in texto_norm]
    cabecalho_produto = (
        "SKU" in texto_norm
        and ("DESCRICAO" in texto_norm or "DESCRICAO" in texto_norm)
        and ("CUSTO" in texto_norm or "EMBALAGEM" in texto_norm or "QUANTIDADE" in texto_norm)
    )
    return cabecalho_produto and len(matriculas) >= 3 and len(set(lojas_encontradas)) >= 2


def _rastrear_por_assinatura_dedicada(texto: str, tipo: str, arquivo_nome: str) -> Optional[ResultadoRastreabilidade]:
    """Identificação forte antes do de/para genérico.

    Inclui Excel VIP matricial sem depender do nome do arquivo e PDFs com
    assinatura dedicada já homologada.
    """
    tipo_norm = str(tipo or "").upper().strip()

    if tipo_norm == "EXCEL" and _parece_excel_rede_vip(texto, arquivo_nome):
        layout = _layout_ativo_por_nome_exato("REDE VIP Excel Matricial", tipo_norm) or _layout_ativo_por_nome_exato("REDE VIP", tipo_norm)
        if layout:
            return ResultadoRastreabilidade(
                sucesso=True,
                aplicar_automaticamente=True,
                layout_id_referencia=str(layout.get("layout_id", "")),
                nome_layout_referencia=str(layout.get("nome_layout", "REDE VIP Excel Matricial")),
                tipo_arquivo=tipo_norm,
                confianca=99,
                limiar_auto=95,
                limiar_sugestao=70,
                motivo=(
                    "Layout definido por assinatura dedicada Rede VIP Excel: matriz SKU x lojas/matriculas; "
                    "cabecalho SKU/DESCRICAO/CUSTO/quantidade + multiplas matriculas 714 em colunas."
                ),
                tokens_encontrados="SKU | DESCRICAO | CUSTO | QUANTIDADE | MATRICULAS 714 | LOJAS VIP",
                status="RASTREABILIDADE_ASSINATURA_AUTO",
                observacao=(
                    "Layout aplicado automaticamente por assinatura forte da Rede VIP Excel. "
                    "Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
                ),
                arquivo=arquivo_nome,
            )

    if tipo_norm != "PDF":
        return None

    texto_norm = _normalizar_texto(texto)
    digitos = _somente_digitos(f"{texto or ''} {arquivo_nome or ''}")

    assinatura_bozza = (
        ("LINKERP VERSAO" in texto_norm or "LINKERP" in texto_norm)
        and ("SUPERMERCADO BOZA" in texto_norm or "SUPERMERCADO BOZZA" in texto_norm or "73419905" in digitos)
        and "DADOS ENTREGA" in texto_norm
        and "CNPJ CPF" in texto_norm
        and "NUMERO PEDIDO" in texto_norm
        and "CODIGO CODIGO CODIGO DESCRICAO QTDE QTDE VALOR VALOR" in texto_norm
    )
    if assinatura_bozza:
        layout = _layout_ativo_por_nome_exato("BOZZA PDF", tipo_norm) or _layout_ativo_por_nome_exato("BOZA PDF", tipo_norm)
        if layout:
            return ResultadoRastreabilidade(
                sucesso=True,
                aplicar_automaticamente=True,
                layout_id_referencia=str(layout.get("layout_id", "")),
                nome_layout_referencia=str(layout.get("nome_layout", "BOZZA PDF")),
                tipo_arquivo=tipo_norm,
                confianca=99,
                limiar_auto=95,
                limiar_sugestao=70,
                motivo=(
                    "Layout definido por assinatura dedicada Bozza/LinkERP: LinkERP + Supermercado Boza + "
                    "Dados entrega + CNPJ/CPF + Numero Pedido + tabela Codigo/Codigo/Codigo/Descricao/Qtde/Qtde."
                ),
                tokens_encontrados="LINKERP | SUPERMERCADO BOZA | DADOS ENTREGA | CNPJ/CPF | NUMERO PEDIDO | CODIGO CODIGO CODIGO",
                status="RASTREABILIDADE_ASSINATURA_AUTO",
                observacao=(
                    "Layout aplicado automaticamente por assinatura forte da rede Bozza/Boza. "
                    "Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
                ),
                arquivo=arquivo_nome,
            )

    assinatura_maby = (
        "MABY SUPERMERCADOS" in texto_norm
        and "PEDIDOS DE COMPRA" in texto_norm
        and "CNPJ DA EMPRESA" in texto_norm
        and "CODIGO GTIN DESCRICAO" in texto_norm
        and "11169" in digitos
    )
    if assinatura_maby:
        layout = _layout_ativo_por_nome_exato("MABY SUPERMERCADOS PDF Homologacao", tipo_norm) or _layout_ativo_por_nome_exato("MABY SUPERMERCADOS", tipo_norm)
        if layout:
            return ResultadoRastreabilidade(
                sucesso=True,
                aplicar_automaticamente=True,
                layout_id_referencia=str(layout.get("layout_id", "")),
                nome_layout_referencia=str(layout.get("nome_layout", "MABY SUPERMERCADOS PDF Homologacao")),
                tipo_arquivo=tipo_norm,
                confianca=99,
                limiar_auto=95,
                limiar_sugestao=70,
                motivo=(
                    "Layout definido por assinatura dedicada Maby/SPAL: MABY SUPERMERCADOS + "
                    "PEDIDOS DE COMPRA + CNPJ da Empresa + tabela Codigo/GTIN/Descricao/Quanti."
                ),
                tokens_encontrados="MABY SUPERMERCADOS | PEDIDOS DE COMPRA | CNPJ DA EMPRESA | GTIN | 11169",
                status="RASTREABILIDADE_ASSINATURA_AUTO",
                observacao=(
                    "Layout aplicado automaticamente por assinatura forte da rede Maby Supermercados. "
                    "Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
                ),
                arquivo=arquivo_nome,
            )

    assinatura_estrela = (
        "SUPERMERCADO ESTRELA" in texto_norm
        and "PEDIDO DE COMPRAS" in texto_norm
        and "DADOS PARA FATURAMENTO" in texto_norm
        and "COD FORN" in texto_norm
        and "EAN" in texto_norm
        and "TOTVS VAREJO SUPERMERCADOS" in texto_norm
    )
    if assinatura_estrela:
        layout = _layout_ativo_por_nome_exato("REDE ESTRELA PDF", tipo_norm) or _layout_ativo_por_nome_exato("ESTRELA PDF", tipo_norm)
        if layout:
            return ResultadoRastreabilidade(
                sucesso=True,
                aplicar_automaticamente=True,
                layout_id_referencia=str(layout.get("layout_id", "")),
                nome_layout_referencia=str(layout.get("nome_layout", "REDE ESTRELA PDF")),
                tipo_arquivo=tipo_norm,
                confianca=99,
                limiar_auto=95,
                limiar_sugestao=70,
                motivo=(
                    "Layout definido por assinatura dedicada Estrela/TOTVS: SUPERMERCADO ESTRELA + "
                    "PEDIDO DE COMPRAS + Dados para Faturamento + Cod Forn + EANs + RELPEDSUPRIM."
                ),
                tokens_encontrados="SUPERMERCADO ESTRELA | PEDIDO DE COMPRAS | DADOS PARA FATURAMENTO | COD FORN | EANs | TOTVS",
                status="RASTREABILIDADE_ASSINATURA_AUTO",
                observacao=(
                    "Layout aplicado automaticamente por assinatura forte da rede Estrela. "
                    "Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
                ),
                arquivo=arquivo_nome,
            )

    assinatura_droga_clara = (
        "14169897" in digitos
        and "RELATORIO DE PEDIDO DE COMPRAS" in texto_norm
        and "FILIAL" in texto_norm
        and "CODIGO DESCRICAO" in texto_norm
        and "FABRICANTE" in texto_norm
    )
    if not assinatura_droga_clara:
        return None

    layout = _layout_ativo_por_nome_exato("DROGA CLARA PDF", tipo_norm)
    if not layout:
        return None

    return ResultadoRastreabilidade(
        sucesso=True,
        aplicar_automaticamente=True,
        layout_id_referencia=str(layout.get("layout_id", "")),
        nome_layout_referencia=str(layout.get("nome_layout", "DROGA CLARA PDF")),
        tipo_arquivo=tipo_norm,
        confianca=99,
        limiar_auto=95,
        limiar_sugestao=70,
        motivo=(
            "Layout definido por assinatura dedicada Droga Clara: raiz CNPJ 14169897 + "
            "Relatorio de Pedido de Compras + Filial + tabela Codigo/Descricao/Fabricante/Qtd."
        ),
        tokens_encontrados="14169897 | RELATORIO DE PEDIDO DE COMPRAS | FILIAL | CODIGO DESCRICAO | FABRICANTE | QTD",
        status="RASTREABILIDADE_ASSINATURA_AUTO",
        observacao=(
            "Layout aplicado automaticamente por assinatura forte da Droga Clara. "
            "Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
        ),
        arquivo=arquivo_nome,
    )

def _rastrear_por_depara(texto: str, tipo: str, arquivo_nome: str) -> Optional[ResultadoRastreabilidade]:
    identificadores = _extrair_identificadores_cliente(texto, arquivo_nome)
    if not identificadores:
        return None

    try:
        match = depara_clientes_service.buscar_depara_por_chaves("", identificadores)
    except Exception:
        terminal_log.exception("[RASTREABILIDADE] Falha ao consultar de/para por CNPJ/GLN/matricula.")
        return None

    if not match:
        return None

    layout = _layout_ativo_por_rede(match.get("rede", ""), tipo)
    if not layout:
        terminal_log.info(
            "[RASTREABILIDADE] De/para encontrou rede, mas sem layout ativo do tipo %s | rede=%s | arquivo=%s",
            tipo,
            match.get("rede", ""),
            arquivo_nome,
        )
        return None

    nome_layout = str(layout.get("nome_layout", ""))
    layout_id = str(layout.get("layout_id", ""))
    if not nome_layout or not layout_id:
        return None

    chave_match = str(match.get("chave_match") or match.get("chave_lida") or "")
    tipo_chave = str(match.get("tipo_chave") or "CHAVE")
    matricula = str(match.get("matricula") or "")
    cnpj_oficial = str(match.get("cnpj_oficial") or "")
    rede = str(match.get("rede") or "")

    return ResultadoRastreabilidade(
        sucesso=True,
        aplicar_automaticamente=True,
        layout_id_referencia=layout_id,
        nome_layout_referencia=nome_layout,
        tipo_arquivo=tipo,
        confianca=98,
        limiar_auto=95,
        limiar_sugestao=70,
        motivo=(
            "Layout definido por de/para CNPJ/GLN/matricula; "
            f"rede={rede}; tipo_chave={tipo_chave}; chave_match={chave_match}; "
            f"matricula={matricula}; cnpj_oficial={cnpj_oficial}"
        ),
        tokens_encontrados=" | ".join([x for x in [rede, tipo_chave, chave_match, matricula, cnpj_oficial] if x]),
        status="RASTREABILIDADE_DEPARA_AUTO",
        observacao=(
            "Layout aplicado automaticamente por CNPJ/GLN/matricula ja cadastrada no de/para. "
            "Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
        ),
        arquivo=arquivo_nome,
    )


def _layout_generico_homologacao(tipo: str) -> Optional[Dict[str, str]]:
    """Retorna layout genérico de homologação para uso manual/rastreado.

    Esse fallback é propositalmente conservador: não homologa rede nova e não
    libera fila/TXT. Ele só permite gerar Excel de validação quando nenhum
    layout parecido atingiu confiança segura, evitando bloquear a análise do
    usuário.
    """
    tipo_norm = str(tipo or "").upper().strip()
    if tipo_norm not in {"PDF", "EXCEL"}:
        return None
    try:
        layouts = cadastro_service.carregar_layouts().fillna("")
    except Exception:
        terminal_log.exception("[RASTREABILIDADE] Falha ao carregar layouts para fallback genérico.")
        return None

    candidatos: List[tuple[int, Dict[str, str]]] = []
    for _, row in layouts.iterrows():
        dados = {str(k): str(v or "") for k, v in row.to_dict().items()}
        if str(dados.get("ativo", "")).strip() != "1":
            continue
        if str(dados.get("tipo_arquivo", "")).upper().strip() != tipo_norm:
            continue
        if _layout_bloqueado_conversao(dados):
            continue
        nome_norm = _normalizar_texto(dados.get("nome_layout", ""))
        obs_norm = _normalizar_texto(dados.get("observacoes", ""))
        texto = f"{nome_norm} {obs_norm}"
        score = 0
        # Prioriza exclusivamente o layout genérico de rastreabilidade.
        # Outras redes em homologação (Maby, Passarela, Beltrame etc.) não devem
        # ser usadas como fallback universal de arquivo desconhecido.
        if "RASTREABILIDADE" in nome_norm and "HOMOLOGACAO" in nome_norm:
            score = 100
        elif "RASTREABILIDADE" in nome_norm and "GENERICA" in nome_norm:
            score = 98
        elif "HOMOLOGACAO GENERICA" in nome_norm:
            score = 90
        elif "GENERICA" in nome_norm and "RASTREABILIDADE" in texto:
            score = 85
        if score:
            candidatos.append((score, dados))

    if not candidatos:
        return None
    candidatos.sort(key=lambda item: item[0], reverse=True)
    return candidatos[0][1]


def criar_rastreabilidade_generica(
    caminho: str,
    tipo_arquivo: str,
    *,
    motivo_original: str = "",
    aplicar_automaticamente: bool = False,
) -> Optional[ResultadoRastreabilidade]:
    """Cria resultado de rastreabilidade genérica para validação manual.

    Usado quando o usuário escolhe explicitamente "Rastreabilidade" ou quando
    a similaridade não alcança uma sugestão segura. O status gerado deixa claro
    que é homologação/rastreabilidade, não layout definitivo.
    """
    arquivo = Path(str(caminho))
    tipo = str(tipo_arquivo or "").upper().strip()
    layout = _layout_generico_homologacao(tipo)
    if not layout:
        return None

    nome_layout = str(layout.get("nome_layout", "")).strip()
    layout_id = str(layout.get("layout_id", "")).strip()
    if not nome_layout or not layout_id:
        return None

    return ResultadoRastreabilidade(
        sucesso=True,
        aplicar_automaticamente=bool(aplicar_automaticamente),
        layout_id_referencia=layout_id,
        nome_layout_referencia=nome_layout,
        tipo_arquivo=tipo,
        confianca=35,
        limiar_auto=99,
        limiar_sugestao=1,
        motivo=(
            "Fallback genérico de rastreabilidade/homologação. "
            "Nenhum layout específico atingiu confiança segura."
            + (f" Motivo original: {motivo_original}" if motivo_original else "")
        ),
        tokens_encontrados="RASTREABILIDADE_GENERICA | VALIDACAO_MANUAL | SEM_HOMOLOGACAO_DEFINITIVA",
        status="RASTREABILIDADE_GENERICA_MANUAL" if aplicar_automaticamente else "SUGESTAO_RASTREABILIDADE_GENERICA",
        observacao=(
            "Layout genérico de homologação aplicado/selecionado por rastreabilidade. "
            "Conferir SKU, QTD, pedido, CNPJ/GLN e matrícula no Excel antes de qualquer TXT/fila. "
            "Se o layout se repetir com segurança, criar parser específico depois."
        ),
        arquivo=arquivo.name,
    )

def rastrear_layout_arquivo(caminho: str, tipo_arquivo: str, *, permitir_auto: bool = True) -> ResultadoRastreabilidade:
    arquivo = Path(str(caminho))
    tipo = str(tipo_arquivo or "").upper().strip()
    base = ResultadoRastreabilidade(False, False, tipo_arquivo=tipo, arquivo=arquivo.name)
    if tipo not in {"PDF", "EXCEL"}:
        base.motivo = "Tipo de arquivo fora da rastreabilidade."
        return base

    texto, auditoria = extrair_amostra_arquivo(str(arquivo), tipo)
    texto_norm = _normalizar_texto(texto)
    nome_arquivo_norm = _normalizar_texto(arquivo.stem)
    if not texto_norm and not nome_arquivo_norm:
        base.motivo = "Nao foi possivel extrair amostra para rastreabilidade."
        base.status = "SEM_AMOSTRA"
        return base

    # Primeiro atalho seguro: assinaturas dedicadas e depois CNPJ/GLN/matrícula.
    # A assinatura dedicada precisa vir antes do de/para para evitar falso positivo
    # por chaves curtas de outra rede, como 000001/pagina -> LJ01/Alabarce.
    if permitir_auto:
        resultado_assinatura = _rastrear_por_assinatura_dedicada(texto, tipo, arquivo.name)
        if resultado_assinatura is not None:
            terminal_log.info(
                "[RASTREABILIDADE] layout por assinatura dedicada | arquivo=%s | layout=%s | motivo=%s",
                arquivo.name,
                resultado_assinatura.nome_layout_referencia,
                resultado_assinatura.motivo,
            )
            return resultado_assinatura

        resultado_depara = _rastrear_por_depara(texto, tipo, arquivo.name)
        if resultado_depara is not None:
            terminal_log.info(
                "[RASTREABILIDADE] layout por de/para | arquivo=%s | layout=%s | motivo=%s",
                arquivo.name,
                resultado_depara.nome_layout_referencia,
                resultado_depara.motivo,
            )
            return resultado_depara

    regras = carregar_regras_rastreabilidade()
    regras = regras[
        (regras["tipo_arquivo"].astype(str).str.upper() == tipo)
        & (regras["ativo_rastreabilidade"].astype(str).str.strip() == "1")
    ].reset_index(drop=True)
    if regras.empty:
        base.motivo = "Nenhum layout ativo de rastreabilidade para este tipo."
        base.status = "SEM_REGRAS"
        return base

    candidatos = []
    for _, row in regras.iterrows():
        regra = row.to_dict()
        score, tokens, motivo = _pontuar_candidato(texto_norm, nome_arquivo_norm, regra)
        candidatos.append((score, tokens, motivo, regra))

    candidatos.sort(key=lambda item: item[0], reverse=True)
    melhor_score, tokens, motivo, regra = candidatos[0]
    segundo_score = candidatos[1][0] if len(candidatos) > 1 else 0
    limiar_auto = int(str(regra.get("limiar_auto") or LIMIAR_AUTO_PADRAO))
    limiar_sugestao = int(str(regra.get("limiar_sugestao") or LIMIAR_SUGESTAO_PADRAO))
    margem = melhor_score - segundo_score

    # Aplicacao automatica somente quando a confianca e boa e a sugestao nao empatou com outro layout.
    aplicar = bool(permitir_auto and melhor_score >= limiar_auto and margem >= 8)
    sucesso = melhor_score >= limiar_sugestao
    status = "RASTREABILIDADE_AUTO" if aplicar else ("SUGESTAO_RASTREABILIDADE" if sucesso else "SEM_SUGESTAO")

    resultado = ResultadoRastreabilidade(
        sucesso=sucesso,
        aplicar_automaticamente=aplicar,
        layout_id_referencia=str(regra.get("layout_id_referencia", "")),
        nome_layout_referencia=str(regra.get("nome_layout_referencia", "")),
        tipo_arquivo=tipo,
        confianca=int(melhor_score),
        limiar_auto=limiar_auto,
        limiar_sugestao=limiar_sugestao,
        motivo=f"{motivo}; margem={margem}; segundo_score={segundo_score}",
        tokens_encontrados=" | ".join(tokens),
        status=status,
        observacao=(
            "Layout aplicado por rastreabilidade. Conferir obrigatoriamente o Excel de validacao antes de TXT/fila."
            if aplicar else
            "Sugestao de layout por rastreabilidade. Confirmar manualmente antes de processar."
            if sucesso else
            "Sem similaridade suficiente para sugerir layout."
        ),
        arquivo=arquivo.name,
    )

    # Se não houver sugestão segura, ainda assim disponibiliza um layout
    # genérico de homologação para o usuário validar o Excel. Isso corrige o
    # bloqueio em arquivos parecidos/desconhecidos que antes exibiam apenas
    # "sem sugestão segura" e não permitiam processar para validação.
    if not resultado.sucesso:
        fallback = criar_rastreabilidade_generica(
            str(arquivo),
            tipo,
            motivo_original=resultado.motivo or resultado.observacao,
            aplicar_automaticamente=False,
        )
        if fallback is not None:
            terminal_log.info(
                "[RASTREABILIDADE] fallback generico disponivel | arquivo=%s | tipo=%s | layout=%s",
                arquivo.name,
                tipo,
                fallback.nome_layout_referencia,
            )
            return fallback

    terminal_log.info(
        "[RASTREABILIDADE] arquivo=%s | tipo=%s | melhor=%s | score=%s | auto=%s | motivo=%s",
        arquivo.name,
        tipo,
        resultado.nome_layout_referencia,
        resultado.confianca,
        resultado.aplicar_automaticamente,
        resultado.motivo,
    )
    return resultado




def _layout_rastreado_seguro_sem_prefixo(nome_ref: str) -> bool:
    nome = _normalizar_texto(nome_ref)
    seguros = (
        "ALABARCE", "DROGA CLARA", "MONACO", "PRIMATO",
        "DAHER", "SEMPRE VALE", "BAKLIZI", "BAZKILI", "MABY", "BOZZA", "BOZA", "ESTRELA", "REDE VIP",
    )
    return any(chave in nome for chave in seguros)

def aplicar_rastreabilidade_no_item(item: Dict[str, object], resultado: ResultadoRastreabilidade) -> Dict[str, object]:
    item["layout_id"] = resultado.layout_id_referencia
    if _layout_rastreado_seguro_sem_prefixo(resultado.nome_layout_referencia):
        item["layout_nome"] = resultado.nome_layout_referencia
        item["status"] = "Pronto para validação"
        item["mensagem"] = (
            f"Layout homologado aplicado com {resultado.confianca}% de confianca: "
            f"{resultado.nome_layout_referencia}. Conferir Excel antes de TXT/fila."
        )
    else:
        item["layout_nome"] = f"RASTREABILIDADE -> {resultado.nome_layout_referencia}"
        item["status"] = "Pronto para validação rastreada"
        item["mensagem"] = (
            f"Rastreabilidade aplicada com {resultado.confianca}% de confianca. "
            f"Layout referencia: {resultado.nome_layout_referencia}. Conferir Excel antes de TXT/fila."
        )
    item["modo_rastreabilidade"] = "SIM"
    item["rastreabilidade"] = resultado.to_dict()
    alertas = list(item.get("alertas", []) or [])
    alertas.append(
        f"RASTREABILIDADE_LAYOUT: layout referencia={resultado.nome_layout_referencia}; confianca={resultado.confianca}%; tokens={resultado.tokens_encontrados}"
    )
    item["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
    return item


def registrar_sugestao_no_item(item: Dict[str, object], resultado: ResultadoRastreabilidade) -> Dict[str, object]:
    item["rastreabilidade"] = resultado.to_dict()
    if resultado.sucesso:
        item["status"] = "Sugestão de rastreabilidade"
        item["mensagem"] = (
            f"Sugestao: {resultado.nome_layout_referencia} ({resultado.confianca}% de confianca). "
            "Use 'Aplicar rastreabilidade' ou selecione o layout manualmente."
        )
        alertas = list(item.get("alertas", []) or [])
        alertas.append(
            f"SUGESTAO_RASTREABILIDADE: {resultado.nome_layout_referencia}; confianca={resultado.confianca}%; tokens={resultado.tokens_encontrados}"
        )
        item["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
    return item


def df_rastreabilidade(resultado: Dict[str, object] | ResultadoRastreabilidade | None) -> pd.DataFrame:
    if resultado is None:
        return pd.DataFrame()
    dados = resultado.to_dict() if isinstance(resultado, ResultadoRastreabilidade) else dict(resultado)
    if not dados:
        return pd.DataFrame()
    return pd.DataFrame([dados])
