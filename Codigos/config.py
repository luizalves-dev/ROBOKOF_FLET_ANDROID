import os
from pathlib import Path

try:
    from dotenv import load_dotenv  # type: ignore
    load_dotenv()
except Exception:
    # O projeto continua funcionando sem python-dotenv; variaveis de ambiente seguem opcionais.
    pass

# ROOT_DIR dinâmico: por padrão usa a pasta do projeto, mas permite sobrescrever por variável.
# Isso evita quebrar quando o ZIP é extraído em qualquer pasta local/rede.
ROOT_DIR = Path(os.getenv("ROBOKOF_ROOT", Path(__file__).resolve().parents[1])).resolve()

# Aplicativo móvel: integrações SAP e Outlook foram removidas.

# Controle obrigatório: TXT/fila final somente após validação manual.
ROBOKOF_EXIGIR_VALIDACAO_MANUAL = os.getenv("ROBOKOF_EXIGIR_VALIDACAO_MANUAL", "1").strip() not in {"0", "false", "False", "NAO", "NÃO", "nao", "não"}
ROBOKOF_COLUNAS_VALIDACAO_MANUAL = ["Validado Manualmente", "Status Validação", "Status Validacao", "VALIDADO_MANUALMENTE"]
ROBOKOF_VALORES_VALIDACAO_MANUAL = {"SIM", "S", "YES", "Y", "VALIDADO", "OK", "CONFERIDO"}
# Compatibilidade operacional:
# Quando o arquivo Pedidos_RoboKOF.xlsx já está dentro de Resultados/Arquivos Fila
# e passa 100% na validação técnica, o robô pode liberar TXT mesmo que a coluna
# "Validado Manualmente" ainda não exista. Mantém a trava para arquivos com erro.
ROBOKOF_PERMITIR_FILA_LEGADA_SEM_COLUNA_VALIDACAO = os.getenv(
    "ROBOKOF_PERMITIR_FILA_LEGADA_SEM_COLUNA_VALIDACAO", "1"
).strip() not in {"0", "false", "False", "NAO", "NÃO", "nao", "não"}

# Status que não devem bloquear a geração quando o arquivo oficial da fila já está tecnicamente consolidado.
# A duplicidade de SKU passou a ser consolidada por soma em validacoes.py, mas estes status ficam
# como compatibilidade caso algum arquivo antigo ainda traga alertas de duplicidade.
ROBOKOF_STATUS_ERRO_NAO_BLOQUEANTES = {
    "SKU_DUPLICADO_CONSOLIDADO",
    "SKU_DUPLICADO_QTD_MENOR",
    "SKU_DUPLICADO",
    # Em produção, algumas filas legadas podem carregar uma linha residual/auxiliar
    # com matrícula, pedido, qtd e data, mas sem SKU. Isso não deve bloquear
    # todo o lote quando existem linhas válidas para o mesmo arquivo oficial.
    # A linha sem SKU continua sendo registrada no Excel de diagnóstico/erro e
    # NÃO entra no TXT.
    "SKU_VAZIO",
    "LINHA_RESIDUAL_SEM_SKU",
}


# =========================
# Pastas do projeto
# =========================
FILA_DIR = ROOT_DIR / "Resultados" / "Arquivos Fila"
OUT_ROBOKOF_DIR = ROOT_DIR / "Resultados" / "Arquivos RoboKOF"
OUT_TXT_DIR = ROOT_DIR / "Resultados" / "Arquivos TXT EDI"
OUT_ERRO_DIR = ROOT_DIR / "Resultados" / "Arquivos ERRO"

# =========================
# Expansão RoboKOF 3.0 - Interface/Layout
# Mantém ROOT_DIR, base de GLNs e caminhos operacionais já definidos acima.
# =========================
RESULTADOS_DIR = ROOT_DIR / "Resultados"
LOGS_DIR = RESULTADOS_DIR / "Logs"
CADASTROS_DIR = ROOT_DIR / "Cadastros"
DADOS_DIR = ROOT_DIR / "Dados"
ENTRADAS_CLIENTES_EXCEL_DIR = ROOT_DIR / "Entradas_Clientes" / "Excel"
ENTRADAS_CLIENTES_PDF_DIR = ROOT_DIR / "Entradas_Clientes" / "PDF"
ERROS_DIR = OUT_ERRO_DIR
ARQUIVOS_BASE_DIR = ROOT_DIR / "Arquivos Base"
PEDIDOS_A_VALIDAR_DIR = RESULTADOS_DIR / "pedidos_a_validar"
PEDIDOS_VALIDADOS_DIR = RESULTADOS_DIR / "pedidos_validados"
PEDIDOS_COM_ERRO_DIR = RESULTADOS_DIR / "pedidos_com_erro"
PROCESSAMENTO_IMPORTACOES_CLIENTES_LEGACY_DIR = RESULTADOS_DIR / "processamento_importacoes_clientes"
PROCESSAMENTO_IMPORTACOES_CLIENTES_DIR = RESULTADOS_DIR / "validacoes_clientes"
HISTORICO_IMPORTACOES_DIR = RESULTADOS_DIR / "historico_importacoes"
TEMP_DIR = RESULTADOS_DIR / "temp"
BH_OUTPUT_ROOT = PEDIDOS_A_VALIDAR_DIR / "Rede_BH"
BH_BASE_PATH = ARQUIVOS_BASE_DIR / "BH_CNPJ_MATRICULA_BASE.txt"
FILA_FILE_NAME = "Pedidos_RoboKOF.xlsx"
FILA_COLUMNS = ["Matricula", "Sku", "Qtd", "Nº Pedido", "Data remessa"]

for _dir in [
    FILA_DIR, OUT_ROBOKOF_DIR, OUT_TXT_DIR, OUT_ERRO_DIR, LOGS_DIR,
    CADASTROS_DIR, ENTRADAS_CLIENTES_EXCEL_DIR, ENTRADAS_CLIENTES_PDF_DIR,
    PEDIDOS_A_VALIDAR_DIR, PEDIDOS_VALIDADOS_DIR, PEDIDOS_COM_ERRO_DIR,
    PROCESSAMENTO_IMPORTACOES_CLIENTES_LEGACY_DIR, PROCESSAMENTO_IMPORTACOES_CLIENTES_DIR,
    HISTORICO_IMPORTACOES_DIR, TEMP_DIR, BH_OUTPUT_ROOT,
]:
    _dir.mkdir(parents=True, exist_ok=True)


TEMPLATE_PATH = ROOT_DIR / "Arquivos Base" / "Exemplo envio.xlsx"
GLN_BASE_PATH = ROOT_DIR / "Arquivos Base" / "BASE de GLNS.xlsx"
# Base complementar para chaves próprias de clientes/layouts.
# Use para GLN, código de loja, código cliente, CNPJ-base ou matrícula específica
# que não exista/venha de forma diferente na BASE de GLNS.xlsx.
DE_PARA_CLIENTES_PATH = ROOT_DIR / "Cadastros" / "de_para_clientes.csv"
# Alias de compatibilidade para evitar quebra em códigos antigos/novos que usem outro nome.
DEPARA_CLIENTES_PATH = DE_PARA_CLIENTES_PATH

# =========================
# Conversão unidade -> caixaria
# =========================
REGRAS_CONVERSAO_PATH = CADASTROS_DIR / "regras_conversao.csv"
# O mapa enviado pelo usuário pode ficar dentro do projeto, sem sobrescrever o arquivo original.
# Ordem de uso no conversao_service:
# 1) variável ROBOKOF_MAPA_PRODUTOS; 2) cópia local em Arquivos Base; 3) caminho legado operacional.
MAPA_PRODUTOS_OFICIAL_DIR = ARQUIVOS_BASE_DIR / "Mapa de Produtos"
MAPA_PRODUTOS_OFICIAL_PATH = MAPA_PRODUTOS_OFICIAL_DIR / "mapa de produtos mais atualizado 08.04.xlsx"
# Mantido por compatibilidade com versões anteriores do Robô KOF. O caminho oficial novo fica em Arquivos Base/Mapa de Produtos/.
MAPA_PRODUTOS_PROJETO_PATH = MAPA_PRODUTOS_OFICIAL_PATH if MAPA_PRODUTOS_OFICIAL_PATH.exists() else (ARQUIVOS_BASE_DIR / "mapa de produtos mais atualizado 08.04.xlsx")
_MAPA_PRODUTOS_LEGADO_ENV = os.getenv("ROBOKOF_MAPA_PRODUTOS_LEGADO", "").strip()
MAPA_PRODUTOS_LEGADO_PATH = Path(_MAPA_PRODUTOS_LEGADO_ENV).expanduser() if _MAPA_PRODUTOS_LEGADO_ENV else MAPA_PRODUTOS_OFICIAL_PATH
MAPA_PRODUTOS_PATH = Path(os.getenv("ROBOKOF_MAPA_PRODUTOS", str(MAPA_PRODUTOS_PROJETO_PATH if MAPA_PRODUTOS_PROJETO_PATH.exists() else MAPA_PRODUTOS_LEGADO_PATH)))
MAPA_PRODUTOS_PADRAO = MAPA_PRODUTOS_PATH
CENTRO_CONVERSAO_PADRAO = os.getenv("ROBOKOF_CENTRO_CONVERSAO_PADRAO", "BAAI")
# Referências específicas por rede/layout. Celeiro deve usar Chapecó como centro base de conversão.
CENTRO_CONVERSAO_CELEIRO = os.getenv("ROBOKOF_CENTRO_CONVERSAO_CELEIRO", "BFDZ")
CENTRO_CONVERSAO_CELEIRO_LABEL = os.getenv("ROBOKOF_CENTRO_CONVERSAO_CELEIRO_LABEL", "Chapecó")
try:
    MAPA_PRODUTOS_OFICIAL_DIR.mkdir(parents=True, exist_ok=True)
except Exception:
    pass


INPUT_SHEET_NAME = None
TEMPLATE_ORDEM_SHEET = "Ordem"

GLN_SHEET_NAME = "BASE de GLNS"
GLN_COL_GLN = "A"
GLN_COL_MATRICULA = "B"
GLN_COL_CNPJ = "I"

HEADER_MAP = {
    "Matricula": ["Matrícula ou CNPJ", "Matricula", "Matrícula", "CNPJ"],
    "DescricaoSku": ["Descrição SKU", "Descricao SKU", "DescricaoSku"],
    "Sku": ["SKU", "Sku"],
    "Qtd": ["Quantidade em Caixas", "Quantidade", "Qtd"],
    "Pedido": ["Numero Pedido Cliente", "Número Pedido Cliente", "Pedido Cliente", "Nº Pedido", "Numero Pedido"],
    "Data": ["Data Entrega", "Data de Entrega", "Data remessa", "Data Remessa"],
    "TipoSolicitacao": ["Tipo Solicitação", "Tipo Solicitacao"],
    "FormaPagamento": ["Forma de Pagamento", "Forma Pagamento"],
}

TIPO_SOLICITACAO_VALUE = "ORDEM VENDA"
FORMA_PAGAMENTO_VALUE = "Boleto"

TXT_ENCODING = "latin-1"

# Registro 02 real validado
TXT_REG02_FIXO = "0221530302026040100000000000012 00000000000000000021"
TXT_REG02_TAMANHO = 400

# Registro 03
TXT_REG03_CONDICAO = "CIF"
TXT_REG03_TAMANHO = 400

# Registro 04
TXT_REG04_DESC_ANTES_PONTO = 38
TXT_REG04_DESC_TAMANHO = 35
TXT_REG04_QTD_TAMANHO = 17
TXT_REG04_QTD_BONIF_TAMANHO = 17
TXT_REG04_ESPACOS_APOS_CLIENTE = 6
TXT_REG04_FILLER_FINAL = 107

# Registro 09
TXT_REG09_TAMANHO = 400

# Header 1
TXT_REG01_TIPO_PEDIDO = "001"
TXT_REG01_LISTA_PRECO = "0" * 20
TXT_REG01_BLOCO_FIXO_MEIO = "789890573008406589791000121"
