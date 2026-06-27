"""Inicialização de armazenamento do RoboKOF para desktop e Android."""
from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

APP_DATA_FOLDER = "ROBOKOF"
SEED_FOLDERS = ("Cadastros", "Arquivos Base", "Dados")


def project_dir() -> Path:
    return Path(__file__).resolve().parent


def assets_dir() -> Path:
    configured = os.getenv("FLET_ASSETS_DIR", "").strip()
    if configured:
        candidate = Path(configured).expanduser().resolve()
        if candidate.exists():
            return candidate
    return project_dir() / "assets"


def runtime_root() -> Path:
    configured = os.getenv("ROBOKOF_ROOT", "").strip()
    if configured:
        return Path(configured).expanduser().resolve()

    app_storage = os.getenv("FLET_APP_STORAGE_DATA", "").strip()
    if app_storage:
        return (Path(app_storage).expanduser().resolve() / APP_DATA_FOLDER)

    return project_dir() / ".robokof_data"


def _copy_missing_tree(source: Path, target: Path, *, overwrite: bool = False) -> None:
    if not source.exists():
        return
    for item in source.rglob("*"):
        relative = item.relative_to(source)
        destination = target / relative
        if item.is_dir():
            destination.mkdir(parents=True, exist_ok=True)
            continue
        destination.parent.mkdir(parents=True, exist_ok=True)
        if overwrite or not destination.exists():
            shutil.copy2(item, destination)


def prepare_runtime(*, overwrite_seed: bool = False) -> Path:
    """Cria a área gravável e instala bases/cadastros incorporados ao APK."""
    root = runtime_root()
    root.mkdir(parents=True, exist_ok=True)
    seed_root = assets_dir() / "bootstrap"
    for folder in SEED_FOLDERS:
        _copy_missing_tree(seed_root / folder, root / folder, overwrite=overwrite_seed)

    # Pastas operacionais. O config cria as subpastas detalhadas depois.
    for folder in ("Resultados", "Entradas_Clientes", "Dados"):
        (root / folder).mkdir(parents=True, exist_ok=True)

    os.environ["ROBOKOF_ROOT"] = str(root)
    return root


def install_code_path() -> Path:
    code_dir = project_dir() / "Codigos"
    value = str(code_dir)
    if value not in sys.path:
        sys.path.insert(0, value)
    return code_dir
