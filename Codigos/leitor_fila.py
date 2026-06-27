from pathlib import Path
from typing import List
import pandas as pd

def listar_excels(fila_dir: Path) -> List[Path]:
    if not fila_dir.exists():
        raise FileNotFoundError(f"Pasta de fila não encontrada: {fila_dir}")
    files: List[Path] = []
    for p in fila_dir.iterdir():
        if p.is_file() and p.suffix.lower() in [".xlsx", ".xlsm"]:
            if p.name.startswith("~$"):
                continue
            files.append(p)
    return sorted(files)

def ler_arquivo_entrada(path: Path, sheet_name=None) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
    if isinstance(df, dict):
        df = df[list(df.keys())[0]]
    return df
