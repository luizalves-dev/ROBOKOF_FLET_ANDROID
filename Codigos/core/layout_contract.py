# Autor: Kauê Melo
"""Contrato oficial para novos layouts do Robô KOF.

Use este arquivo como referência de regra de negócio. Ele não força mudança nos
parsers antigos, mas define o padrão para novos cadastros.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class LayoutItem:
    rede: str
    layout: str
    arquivo_origem: str
    linha_origem: int | str = ""
    pagina_origem: int | str = ""
    cnpj: str = ""
    matricula: str = ""
    sku: str = ""
    ean: str = ""
    descricao: str = ""
    quantidade_original: Any = ""
    quantidade_final: Any = ""
    numero_pedido: str = ""
    status_extracao: str = "OK"
    status_conversao: str = "OK SEM CONVERSÃO"
    regra_aplicada: str = ""
    alerta: str = ""
    extras: dict[str, Any] = field(default_factory=dict)


COLUNAS_MODELO_ROBO_KOF = [
    "CNPJ",
    "MATRICULA",
    "SKU",
    "QTD",
    "NUMERO_PEDIDO",
]

ABAS_OBRIGATORIAS_EXCEL_VALIDACAO = [
    "Modelo Robô KOF para Enviar",
    "Validação do Pedido",
    "Alertas/Erros",
]

STATUS_CONVERSAO_PADRAO = {
    "OK_CONVERTIDO": "OK CONVERTIDO",
    "OK_SEM_CONVERSAO": "OK SEM CONVERSÃO",
    "ALERTA_NAO_CONVERTIDO": "ALERTA - NÃO CONVERTIDO",
    "VALIDAR_CONVERSAO": "VALIDAR CONVERSÃO",
}

REGRAS_GERAIS = [
    "Nunca descartar linha silenciosamente.",
    "Nunca gerar TXT/fila antes do Excel de validação manual.",
    "CNPJ sem matrícula deve ficar como A CADASTRAR e ir para aba Cadastrar CNPJ.",
    "Cada layout deve ser isolado em parser próprio ou bridge seguro.",
    "Não alterar layout homologado para corrigir outro layout.",
]
