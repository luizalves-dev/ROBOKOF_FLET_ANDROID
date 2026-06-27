from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
import io
import os
import re
from typing import List

import pandas as pd

from layout_standard import STANDARD_INTERMEDIATE_COLUMNS, empty_intermediate_df as _empty_standard_df, normalize_intermediate_columns
from terminal_logger import get_terminal_logger

try:
    import pdfplumber  # type: ignore
except ModuleNotFoundError:
    pdfplumber = None  # type: ignore

try:
    import fitz  # PyMuPDF
except ModuleNotFoundError:
    fitz = None  # type: ignore

try:
    from PIL import Image  # type: ignore
except ModuleNotFoundError:
    Image = None  # type: ignore

try:
    import pytesseract  # type: ignore
except ModuleNotFoundError:
    pytesseract = None  # type: ignore


terminal_log = get_terminal_logger("pdf_utils")


INTERMEDIATE_COLUMNS = STANDARD_INTERMEDIATE_COLUMNS


@dataclass
class PdfPageAudit:
    pagina: int
    motor: str
    caracteres: int
    tabelas: int = 0
    blocos: int = 0
    status: str = "OK"
    alerta: str = ""


@dataclass
class PdfExtractionResult:
    paginas: List[str]
    auditoria: List[PdfPageAudit]

    @property
    def total_paginas(self) -> int:
        return len(self.paginas)

    @property
    def paginas_processadas(self) -> int:
        return len(self.auditoria)

    @property
    def alertas(self) -> List[str]:
        return [item.alerta for item in self.auditoria if item.alerta]

    def auditoria_df(self) -> pd.DataFrame:
        return pd.DataFrame([asdict(item) for item in self.auditoria])


def ensure_pdfplumber():
    if pdfplumber is None:
        raise RuntimeError(
            "pdfplumber nao esta disponivel. Instale com: python -m pip install pdfplumber"
        )


def ensure_pdf_reader():
    if pdfplumber is None and fitz is None:
        raise RuntimeError("Nenhum leitor de PDF disponivel. Instale pdfplumber ou PyMuPDF.")


def clean_text(value) -> str:
    return str(value or "").strip()


def only_digits(value) -> str:
    return re.sub(r"\D+", "", str(value or "")).strip()


def normalize_qty(value: str) -> str:
    """
    48,000   -> 48
    1.172,00 -> 1172
    """
    texto = str(value or "").strip()
    if not texto:
        return ""

    texto = texto.replace(".", "").replace(",", ".")
    try:
        numero = float(texto)
        if numero.is_integer():
            return str(int(numero))
        return str(numero).replace(".", ",")
    except Exception:
        return ""


def _extract_with_fitz(caminho_arquivo: str) -> tuple[List[str], List[int]]:
    if fitz is None:
        return [], []

    paginas: List[str] = []
    blocos: List[int] = []
    with fitz.open(caminho_arquivo) as doc:
        for page in doc:
            texto = page.get_text("text") or ""
            paginas.append(texto)
            try:
                blocos.append(len(page.get_text("blocks") or []))
            except Exception:
                blocos.append(0)
    return paginas, blocos



def _ocr_habilitado() -> bool:
    return os.getenv("ROBOKOF_ENABLE_OCR", "1").strip().lower() not in {"0", "false", "nao", "não", "n"}


def _extract_page_with_ocr(caminho_arquivo: str, page_index_zero: int) -> tuple[str, str]:
    """Fallback OCR opcional para PDFs imagem.

    Requer PyMuPDF + Pillow + pytesseract e o executável Tesseract instalado no Windows.
    Quando não estiver disponível, não quebra o fluxo: devolve texto vazio e um alerta claro.
    """
    if not _ocr_habilitado():
        return "", "OCR_DESABILITADO"
    if fitz is None or Image is None or pytesseract is None:
        return "", "OCR_INDISPONIVEL_DEPENDENCIAS"
    try:
        with fitz.open(caminho_arquivo) as doc:
            if page_index_zero < 0 or page_index_zero >= len(doc):
                return "", "OCR_PAGINA_INVALIDA"
            page = doc[page_index_zero]
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                texto = pytesseract.image_to_string(img, lang=os.getenv("ROBOKOF_OCR_LANG", "por+eng")) or ""
            except Exception:
                texto = pytesseract.image_to_string(img) or ""
        return texto, "OCR_OK" if texto.strip() else "OCR_SEM_TEXTO"
    except Exception as exc:
        return "", f"OCR_FALHOU: {exc}"


def extract_pages_text_detailed(caminho_arquivo: str) -> PdfExtractionResult:
    ensure_pdf_reader()
    caminho = str(caminho_arquivo)
    paginas: List[str] = []
    auditoria: List[PdfPageAudit] = []
    textos_fitz: List[str] = []
    blocos_fitz: List[int] = []

    if fitz is not None:
        try:
            textos_fitz, blocos_fitz = _extract_with_fitz(caminho)
        except Exception as exc:
            terminal_log.warning("[PDF] Fallback fitz indisponivel para %s: %s", caminho, exc)
            textos_fitz, blocos_fitz = [], []

    if pdfplumber is None:
        paginas = textos_fitz
        for idx, texto in enumerate(paginas, start=1):
            motor = "fitz"
            alerta = ""
            if not texto:
                texto_ocr, status_ocr = _extract_page_with_ocr(caminho, idx - 1)
                if texto_ocr:
                    texto = texto_ocr
                    paginas[idx - 1] = texto
                    motor = "ocr"
                    alerta = f"Pagina {idx}: fitz sem texto; usado OCR"
                else:
                    alerta = f"Pagina {idx}: PAGINA_SEM_TEXTO_EXTRAIVEL ({status_ocr})"
            auditoria.append(
                PdfPageAudit(
                    pagina=idx,
                    motor=motor,
                    caracteres=len(texto or ""),
                    blocos=blocos_fitz[idx - 1] if idx - 1 < len(blocos_fitz) else 0,
                    status="OK" if texto else "ALERTA",
                    alerta=alerta,
                )
            )
        terminal_log.info("[PDF] Leitura concluida com fitz | paginas=%s", len(paginas))
        return PdfExtractionResult(paginas=paginas, auditoria=auditoria)

    with pdfplumber.open(caminho) as pdf:
        for idx, page in enumerate(pdf.pages, start=1):
            texto = ""
            tabelas = 0
            alerta = ""
            motor = "pdfplumber"

            try:
                texto = page.extract_text() or ""
            except Exception as exc:
                alerta = f"Pagina {idx}: erro extract_text pdfplumber: {exc}"

            try:
                tabelas = len(page.extract_tables() or [])
            except Exception:
                tabelas = 0

            if not texto and idx - 1 < len(textos_fitz):
                texto_fitz = textos_fitz[idx - 1] or ""
                if texto_fitz:
                    texto = texto_fitz
                    motor = "fitz"
                    alerta = f"Pagina {idx}: pdfplumber sem texto; usado fallback fitz"

            if not texto:
                texto_ocr, status_ocr = _extract_page_with_ocr(caminho, idx - 1)
                if texto_ocr:
                    texto = texto_ocr
                    motor = "ocr"
                    alerta = f"Pagina {idx}: pdfplumber/fitz sem texto; usado OCR"
                else:
                    alerta = alerta or f"Pagina {idx}: PAGINA_SEM_TEXTO_EXTRAIVEL ({status_ocr})"

            paginas.append(texto)
            auditoria.append(
                PdfPageAudit(
                    pagina=idx,
                    motor=motor,
                    caracteres=len(texto or ""),
                    tabelas=tabelas,
                    blocos=blocos_fitz[idx - 1] if idx - 1 < len(blocos_fitz) else 0,
                    status="OK" if texto else "ALERTA",
                    alerta=alerta,
                )
            )

    motores = {}
    for item in auditoria:
        motores[item.motor] = motores.get(item.motor, 0) + 1
    sem_texto = sum(1 for item in auditoria if not item.caracteres)
    terminal_log.info(
        "[PDF] Leitura detalhada concluida | arquivo=%s | paginas=%s | processadas=%s | motores=%s | sem_texto=%s",
        caminho,
        len(paginas),
        len(auditoria),
        motores,
        sem_texto,
    )
    return PdfExtractionResult(paginas=paginas, auditoria=auditoria)


def extract_pages_text(caminho_arquivo: str) -> List[str]:
    return extract_pages_text_detailed(caminho_arquivo).paginas


def empty_intermediate_df() -> pd.DataFrame:
    return _empty_standard_df()


def build_intermediate_df(rows: list[dict], caminho_arquivo: str, nome_layout: str) -> pd.DataFrame:
    if not rows:
        return normalize_intermediate_columns(
            empty_intermediate_df(),
            arquivo_origem=caminho_arquivo,
            layout_usado=nome_layout,
        )

    df = pd.DataFrame(rows)
    return normalize_intermediate_columns(df, arquivo_origem=Path(caminho_arquivo).name, layout_usado=nome_layout)
