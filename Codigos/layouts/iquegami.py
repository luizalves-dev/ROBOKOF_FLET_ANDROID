# Autor: Kauê Melo
# Bridge seguro para o Robô KOF reconhecer a Rede Iquegami sem mexer nos demais layouts.
from __future__ import annotations

try:
    from integracoes.rede_iquegami.parser_iquegami_kof import *  # noqa: F401,F403
except Exception:  # fallback quando o projeto usa outro path de import
    import sys
    from pathlib import Path
    raiz = Path(__file__).resolve()
    for parent in [raiz.parent, *raiz.parents]:
        if (parent / "integracoes" / "rede_iquegami" / "parser_iquegami_kof.py").exists():
            sys.path.insert(0, str(parent))
            break
    from integracoes.rede_iquegami.parser_iquegami_kof import *  # noqa: F401,F403
