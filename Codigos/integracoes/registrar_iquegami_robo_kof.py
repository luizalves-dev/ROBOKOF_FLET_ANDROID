# Autor: Kauê Melo
"""Registro opcional da Rede Iquegami para projetos Robô KOF que usam dicionário central.

Uso opcional dentro do registrador central, se o seu projeto não tiver descoberta automática:
    from integracoes.registrar_iquegami_robo_kof import registrar_iquegami
    registrar_iquegami(LAYOUTS)

O patch principal não precisa alterar os demais layouts.
"""
from __future__ import annotations

from integracoes.rede_iquegami import parser_iquegami_kof

def registrar_iquegami(registry: dict) -> dict:
    registry.setdefault("IQUEGAMI", parser_iquegami_kof.processar_lote)
    registry.setdefault("REDE IQUEGAMI", parser_iquegami_kof.processar_lote)
    registry.setdefault("iquegami", parser_iquegami_kof.processar_lote)
    return registry
