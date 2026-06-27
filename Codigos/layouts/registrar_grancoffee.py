# -*- coding: utf-8 -*-
"""
Registro do layout GRANCOFFEE / SPAL no Robô KOF
Autor: Kauê Melo

Use este arquivo como apoio para cadastrar o layout no registry/roteador do projeto.
Ele não altera nada sozinho; apenas expõe funções padronizadas para importação.
"""
from __future__ import annotations

from .rede_grancoffee import (
    LAYOUT_NOME,
    detectar_layout,
    identificar_grancoffee,
    processar_grancoffee,
    processar_layout,
)

LAYOUT_INFO = {
    "nome": "GRANCOFFEE",
    "aliases": ["GRANCOFFEE", "GRAN COFFEE", "PEDIDOS GRANCOFFEE", "PEDIDO SPAL"],
    "extensoes": [".msg", ".xlsm", ".xlsx"],
    "descricao": "Grancoffee/SPAL: .msg com anexos Excel; data oficial da remessa vem do corpo do e-mail por matrícula + pedido.",
    "detectar": detectar_layout,
    "processar": processar_layout,
}


def registrar(registry):
    """Registra o layout em dicionários/registries simples do projeto.

    Compatível com rotas que usam algo como:
        LAYOUTS["GRANCOFFEE"] = {...}

    Caso o projeto use outro padrão, copie apenas o import e a chamada de processamento:
        from layouts.rede_grancoffee import detectar_layout, processar_layout
    """
    if registry is None:
        return LAYOUT_INFO
    if isinstance(registry, dict):
        registry["GRANCOFFEE"] = LAYOUT_INFO
        return registry
    if hasattr(registry, "registrar"):
        registry.registrar("GRANCOFFEE", detectar_layout, processar_layout)
        return registry
    raise TypeError("Registry não reconhecido para cadastro automático da Grancoffee.")


def obter_layout():
    return LAYOUT_INFO.copy()
