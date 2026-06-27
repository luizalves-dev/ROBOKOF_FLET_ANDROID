import re
import math
from datetime import datetime
from typing import Optional, Tuple
import pandas as pd

def clean_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()

def normalize_sku(sku: str) -> str:
    sku = clean_str(sku)

    # remove formato científico se vier
    if "E+" in sku.upper():
        try:
            sku = "{:.0f}".format(float(sku))
        except:
            pass

    # mantém só números
    sku = re.sub(r"\D", "", sku)

    return sku

def parse_qtd_to_int(qtd_raw) -> Tuple[Optional[int], Optional[str]]:
    """
    Regras:
      1) < 1  => 1
      2) >= 1 e quebrado => arredondar para baixo (floor)
    Aceita:
      - número (int/float) vindo do pandas
      - string com vírgula (ex.: "5,8")
      - string com ponto (ex.: "0.5")  [casos raros]
      - string com milhar (ex.: "1.234,56")
    """
    if qtd_raw is None:
        return None, "QTD_VAZIA"

    # NaN do pandas
    try:
        if pd.isna(qtd_raw):
            return None, "QTD_VAZIA"
    except Exception:
        pass

    # Se já veio como número, usa direto (NÃO remover ponto!)
    if isinstance(qtd_raw, (int, float)) and not isinstance(qtd_raw, bool):
        x = float(qtd_raw)
    else:
        s = str(qtd_raw).strip()
        if s == "" or s.lower() == "nan":
            return None, "QTD_VAZIA"

        # Normalização para string numérica
        s = s.replace(" ", "")

        if "," in s:
            # se tem "." e "," assume "." milhar e "," decimal
            if "." in s:
                s = s.replace(".", "")
            s = s.replace(",", ".")
        # se não tem vírgula, mantém ponto como decimal (se existir)

        try:
            x = float(s)
        except Exception:
            return None, f"QTD_INVALIDA: {qtd_raw}"

    # aplica regras
    if x <= 0:
        return 1, None
    if x < 1:
        return 1, None
    return int(math.floor(x)), None

def sanitize_pedido(pedido_raw: str) -> str:
    """
    Remove espaços e caracteres especiais.
    Mantém letras e números. Ex: "AB-12 34" -> "AB1234"
    """
    s = clean_str(pedido_raw)
    s2 = re.sub(r"[^0-9A-Za-z]+", "", s)
    return s2

def normalize_date_remessa(date_raw):
    if date_raw is None:
        return None, "DATA_VAZIA"

    # pandas / Excel datetime
    if isinstance(date_raw, (datetime, pd.Timestamp)):
        return date_raw.strftime("%d/%m/%Y"), None

    # string
    s = str(date_raw).strip()

    if s == "" or s.lower() == "nan":
        return None, "DATA_VAZIA"

    # troca separador
    s = s.replace(".", "/")

    # tenta dd/mm/aaaa
    try:
        dt = datetime.strptime(s, "%d/%m/%Y")
        return dt.strftime("%d/%m/%Y"), None
    except ValueError:
        return None, "DATA_INVALIDA"

    
def pad_gln_14(gln_raw: str) -> str:
    s = clean_str(gln_raw)
    s = re.sub(r"[^0-9]+", "", s)
    return s.zfill(14)
