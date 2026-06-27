"""Diagnóstico interno sem subprocessos nem recursos Windows."""
from __future__ import annotations

import importlib
from pathlib import Path

import cadastro_service
import config


def run_health_check() -> dict[str, object]:
    checks: list[dict[str, str]] = []

    def add(name: str, ok: bool, detail: str) -> None:
        checks.append({"item": name, "status": "OK" if ok else "ERRO", "detalhe": detail})

    for name, path in (
        ("Cadastros", config.CADASTROS_DIR),
        ("Arquivos Base", config.ARQUIVOS_BASE_DIR),
        ("Base GLN", config.GLN_BASE_PATH),
        ("Template RoboKOF", config.TEMPLATE_PATH),
    ):
        add(name, Path(path).exists(), str(path))

    cadastro_errors = cadastro_service.validar_estrutura_cadastros()
    add("Estrutura dos cadastros", not cadastro_errors, "; ".join(cadastro_errors) or "Estrutura válida")

    layouts = cadastro_service.listar_layouts_ativos()
    add("Layouts ativos", not layouts.empty, f"{len(layouts)} layout(s) ativo(s)")

    for module_name in (
        "leitor_pdf_clientes",
        "leitor_excel_clientes",
        "importador_clientes",
        "gerador_txt_edi",
        "layouts.rede_bh.processor",
    ):
        try:
            importlib.import_module(module_name)
            add(f"Módulo {module_name}", True, "Importação concluída")
        except Exception as exc:
            add(f"Módulo {module_name}", False, str(exc))

    errors = sum(1 for c in checks if c["status"] == "ERRO")
    return {"checks": checks, "errors": errors, "ok": errors == 0}
