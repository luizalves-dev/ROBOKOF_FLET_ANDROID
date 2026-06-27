from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd

try:
    import config
except Exception:  # pragma: no cover
    config = None

from terminal_logger import get_terminal_logger


terminal_log = get_terminal_logger("depara_clientes")

COLUNAS_DEPARA = [
    "REDE",
    "TIPO_CHAVE",
    "CHAVE_LIDA",
    "CNPJ_OFICIAL",
    "MATRICULA",
    "NOME_LOJA",
    "STATUS",
    "OBSERVACAO",
    "FONTE",
    "DATA_ATUALIZACAO",
]

COLUNAS_PENDENCIA = [
    "Rede/Layout",
    "Tipo da Chave",
    "Chave Lida",
    "CNPJ Lido",
    "CNPJ Oficial",
    "Matrícula",
    "Status",
    "Nº do Pedido",
    "Observação",
]

STATUS_ATIVO = {"", "ATIVO", "A CADASTRAR", "CADASTRADO", "OK", "VALIDADO"}
STATUS_IGNORAR = {"INATIVO", "IGNORAR", "DESATIVADO", "CANCELADO"}
_DEPARA_CACHE: dict[tuple[str, str, str], list[DeParaEntry]] = {}


def _file_cache_key(path: Path) -> tuple[str, str, str]:
    path = Path(path)
    try:
        resolved = str(path.resolve())
    except Exception:
        resolved = str(path)
    try:
        stat = path.stat()
        return resolved, str(stat.st_mtime_ns), str(stat.st_size)
    except Exception:
        return resolved, "", ""


def limpar_cache_depara(path: Path | None = None):
    if path is None:
        _DEPARA_CACHE.clear()
        return
    path = Path(path)
    prefixo = str(path.resolve()) if path.exists() else str(path)
    for key in list(_DEPARA_CACHE.keys()):
        if key and key[0] == prefixo:
            _DEPARA_CACHE.pop(key, None)



def _path_default() -> Path:
    if config is not None and hasattr(config, "DE_PARA_CLIENTES_PATH"):
        return Path(config.DE_PARA_CLIENTES_PATH)
    return Path(__file__).resolve().parents[1] / "Cadastros" / "de_para_clientes.csv"


def clean_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def only_digits(value) -> str:
    return re.sub(r"\D+", "", clean_text(value))


def normalize_rede(value) -> str:
    text = clean_text(value).upper()
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\b(PDF|EXCEL|LAYOUT|RASTREABILIDADE)\b", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_key(value) -> str:
    text = clean_text(value).upper()
    if not text:
        return ""
    digits = only_digits(text)
    if digits:
        return digits
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text.strip()


def key_variants(value) -> list[str]:
    key = normalize_key(value)
    if not key:
        return []
    variants = [key]
    if key.isdigit():
        # Segurança de identificação: nunca transformar códigos curtos
        # com zero à esquerda em chaves genéricas como "1", "2", "15" etc.
        # Isso causava falso de/para em rastreabilidade: PDF Droga Clara com
        # "Pagina 000001" era comparado com nome de loja "LJ01" da Alabarce
        # e o layout virava ALABARCE PDF indevidamente.
        if len(key) <= 6:
            candidatos = []
        else:
            stripped = key.lstrip("0") or "0"
            candidatos = []
            if len(stripped) >= 6:
                candidatos.append(stripped)
            z14 = key.zfill(14) if 8 <= len(key) <= 14 else ""
            z13 = key.zfill(13) if 8 <= len(key) <= 13 else ""
            candidatos.extend([v for v in [z14, z13] if v])
        # CNPJ completo -> também tenta a base de 12 dígitos, útil quando o PDF traz CNPJ sem DV.
        if len(key) >= 14:
            candidatos.append(key[:12])
        # GLN / texto de loja de 13 dígitos -> também tenta os 12 primeiros dígitos.
        if len(key) == 13:
            candidatos.append(key[:12])
        # CNPJ-base de 12 dígitos -> mantém base e versões sem zero à esquerda.
        if len(key) == 12:
            stripped = key.lstrip("0") or "0"
            if len(stripped) >= 6:
                candidatos.append(stripped)

        # Ajuste específico MABY/SPAL: alguns PDFs antigos extraem 15 dígitos
        # para a raiz 01695774, com um dígito excedente logo após a raiz.
        # Ex.: 016957746000128 -> CNPJ oficial 01695774000128.
        # Também gera a variante inversa para aceitar de/para preenchido no formato bruto do PDF.
        if key.startswith("01695774"):
            if len(key) == 15:
                candidatos.append(key[:8] + key[9:])
            elif len(key) == 14:
                candidatos.append(key[:8] + "6" + key[8:])

        for variant in candidatos:
            if variant and variant not in variants:
                variants.append(variant)
    return variants


def infer_tipo_chave(value, fallback: str = "") -> str:
    explicit = clean_text(fallback).upper().replace(" ", "_")
    if explicit:
        return explicit
    digits = only_digits(value)
    if len(digits) == 14:
        return "CNPJ"
    if len(digits) == 13:
        return "GLN"
    if len(digits) >= 8:
        return "COD_CLIENTE"
    if digits:
        return "COD_LOJA"
    return "CHAVE"


@dataclass(frozen=True)
class DeParaEntry:
    rede: str
    tipo_chave: str
    chave_lida: str
    cnpj_oficial: str
    matricula: str
    nome_loja: str = ""
    status: str = "ATIVO"
    observacao: str = ""
    fonte: str = ""
    data_atualizacao: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "REDE": self.rede,
            "TIPO_CHAVE": self.tipo_chave,
            "CHAVE_LIDA": self.chave_lida,
            "CNPJ_OFICIAL": self.cnpj_oficial,
            "MATRICULA": self.matricula,
            "NOME_LOJA": self.nome_loja,
            "STATUS": self.status or "ATIVO",
            "OBSERVACAO": self.observacao,
            "FONTE": self.fonte,
            "DATA_ATUALIZACAO": self.data_atualizacao,
        }


def ensure_depara_file(path: Path | None = None) -> Path:
    path = Path(path or _path_default())
    path.parent.mkdir(parents=True, exist_ok=True)
    if not path.exists():
        with path.open("w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=COLUNAS_DEPARA, delimiter=";")
            writer.writeheader()
    return path


def _row_to_entry(row: dict) -> DeParaEntry | None:
    rede = clean_text(row.get("REDE") or row.get("Rede/Layout") or row.get("rede") or row.get("layout"))
    chave = clean_text(row.get("CHAVE_LIDA") or row.get("Chave Lida") or row.get("CNPJ Lido") or row.get("CNPJ") or row.get("cnpj_lido"))
    cnpj_oficial = only_digits(row.get("CNPJ_OFICIAL") or row.get("CNPJ Oficial") or row.get("CNPJ") or row.get("cnpj_oficial"))
    matricula = only_digits(row.get("MATRICULA") or row.get("Matrícula") or row.get("Matricula") or row.get("Matrícula Encontrada") or row.get("matricula"))
    status = clean_text(row.get("STATUS") or row.get("Status") or "ATIVO").upper()
    if status in STATUS_IGNORAR:
        return None
    if not chave:
        chave = cnpj_oficial
    if not chave or not matricula:
        return None
    return DeParaEntry(
        rede=rede,
        tipo_chave=infer_tipo_chave(chave, row.get("TIPO_CHAVE") or row.get("Tipo da Chave")),
        chave_lida=normalize_key(chave),
        cnpj_oficial=cnpj_oficial,
        matricula=matricula,
        nome_loja=clean_text(row.get("NOME_LOJA") or row.get("Nome Loja") or row.get("Loja")),
        status=status if status in STATUS_ATIVO else "ATIVO",
        observacao=clean_text(row.get("OBSERVACAO") or row.get("Observação") or row.get("Observacao")),
        fonte=clean_text(row.get("FONTE") or row.get("Fonte")),
        data_atualizacao=clean_text(row.get("DATA_ATUALIZACAO") or row.get("Data Atualização") or row.get("Data Atualizacao")),
    )


def carregar_depara_clientes(path: Path | None = None) -> list[DeParaEntry]:
    path = ensure_depara_file(path)
    cache_key = _file_cache_key(path)
    cached = _DEPARA_CACHE.get(cache_key)
    if cached is not None:
        return cached

    entries: list[DeParaEntry] = []
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                entry = _row_to_entry(row)
                if entry:
                    entries.append(entry)
    except Exception:
        terminal_log.exception("[DEPARA] Falha ao carregar de/para complementar: %s", path)
        return []
    _DEPARA_CACHE.clear()
    _DEPARA_CACHE[cache_key] = entries
    terminal_log.info("[DEPARA] De/para complementar carregado/cacheado | registros=%s | arquivo=%s", len(entries), path)
    return entries


def _score_rede(rede_layout: str, rede_entry: str) -> int:
    layout = normalize_rede(rede_layout)
    entry = normalize_rede(rede_entry)
    if not entry:
        return 1
    if layout == entry:
        return 100
    if entry and (entry in layout or layout in entry):
        return 80
    tokens_layout = set(layout.split())
    tokens_entry = set(entry.split())
    inter = tokens_layout & tokens_entry
    return 30 + len(inter) if inter else 0


def buscar_depara_por_chaves(
    rede_layout: str,
    chaves: Sequence[object],
    entries: list[DeParaEntry] | None = None,
    path: Path | None = None,
) -> dict[str, str]:
    entries = entries if entries is not None else carregar_depara_clientes(path)
    variants: list[str] = []
    chaves_exatas: list[str] = []
    for chave in chaves:
        chave_norm = normalize_key(chave)
        if chave_norm and chave_norm not in chaves_exatas:
            chaves_exatas.append(chave_norm)
        for variant in key_variants(chave):
            if variant not in variants:
                variants.append(variant)
    if not variants:
        return {}

    def _qualidade_match(entry: DeParaEntry, matched_key: str) -> int:
        """Prioriza de/para exato antes de variantes por raiz/base.

        Isso é crítico para MABY/SPAL: os CNPJs do PDF podem vir com 15
        dígitos e várias lojas compartilham os 12 primeiros dígitos. Sem esta
        prioridade, uma busca pela loja 002 podia bater na base de 12 dígitos
        da loja 001 e puxar matrícula errada.
        """
        entry_exatos = {normalize_key(entry.chave_lida), normalize_key(entry.cnpj_oficial)}
        if any(chave in entry_exatos for chave in chaves_exatas):
            return 1000
        if matched_key in chaves_exatas and len(matched_key) >= 8:
            return 900
        if matched_key in entry_exatos and len(matched_key) >= 8:
            return 850
        if matched_key.isdigit() and len(matched_key) >= 14:
            return 800
        if matched_key.isdigit() and len(matched_key) == 13:
            return 700
        if matched_key.isdigit() and len(matched_key) == 12:
            return 300
        return 100

    candidatos: list[tuple[int, int, DeParaEntry, str]] = []
    candidatos_globais: list[tuple[int, int, DeParaEntry, str]] = []
    for entry in entries:
        entry_keys = []
        # Inclui a matrícula como chave de busca complementar.
        # Segurança: quando a busca cair em fallback global, a função só aplica
        # automaticamente se a chave for única; isso evita escolher layout errado
        # quando um número semelhante aparecer em mais de uma rede.
        for origem in (entry.chave_lida, entry.cnpj_oficial, entry.nome_loja, entry.matricula):
            for key in key_variants(origem):
                if key not in entry_keys:
                    entry_keys.append(key)
        matches = [key for key in variants if key in entry_keys]
        if not matches:
            continue
        matched_key = sorted(matches, key=lambda key: _qualidade_match(entry, key), reverse=True)[0]
        qualidade = _qualidade_match(entry, matched_key)
        score = _score_rede(rede_layout, entry.rede)
        if score > 0:
            candidatos.append((qualidade, score, entry, matched_key))
        else:
            # Fallback controlado: se a rede/layout não bater, só usa quando a chave é exata
            # e não houver candidato melhor da própria rede. Isso ajuda layouts rastreados/aliases.
            candidatos_globais.append((qualidade, 1, entry, matched_key))

    if not candidatos and candidatos_globais:
        # Evita matrícula errada em chave repetida em mais de uma rede.
        chaves_unicas = {(entry.chave_lida, entry.cnpj_oficial, entry.matricula) for _, _, entry, _ in candidatos_globais}
        if len(chaves_unicas) == 1:
            candidatos = candidatos_globais
        else:
            terminal_log.warning(
                "[DEPARA] Chave localizada em múltiplas redes fora do layout; cadastro não aplicado automaticamente | layout=%s | chaves=%s",
                rede_layout,
                variants[:5],
            )
            return {}

    if not candidatos:
        return {}
    candidatos.sort(key=lambda item: (item[0], item[1]), reverse=True)
    qualidade, score, entry, matched_key = candidatos[0]
    return {
        "rede": entry.rede,
        "tipo_chave": entry.tipo_chave,
        "chave_lida": entry.chave_lida,
        "chave_match": matched_key,
        "cnpj_oficial": entry.cnpj_oficial,
        "matricula": entry.matricula,
        "nome_loja": entry.nome_loja,
        "status": entry.status,
        "observacao": entry.observacao,
        "fonte": entry.fonte,
        "score_rede": str(score),
        "qualidade_match": str(qualidade),
    }


def _entry_identity(entry: DeParaEntry) -> tuple[str, str, str]:
    return (normalize_rede(entry.rede), normalize_key(entry.chave_lida), only_digits(entry.matricula))


def salvar_depara_clientes(entries: Iterable[DeParaEntry], path: Path | None = None) -> Path:
    path = ensure_depara_file(path)
    unique: dict[tuple[str, str, str], DeParaEntry] = {}
    for entry in entries:
        if not entry.chave_lida or not entry.matricula:
            continue
        unique[_entry_identity(entry)] = entry
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=COLUNAS_DEPARA, delimiter=";")
        writer.writeheader()
        for entry in sorted(unique.values(), key=lambda e: (normalize_rede(e.rede), e.tipo_chave, e.chave_lida, e.matricula)):
            writer.writerow(entry.as_dict())
    limpar_cache_depara(path)
    terminal_log.info("[DEPARA] Base complementar salva | registros=%s | arquivo=%s", len(unique), path)
    return path


def adicionar_ou_atualizar_entries(novas: Iterable[DeParaEntry], path: Path | None = None) -> tuple[int, int, Path]:
    path = ensure_depara_file(path)
    existentes = carregar_depara_clientes(path)
    mapa = {_entry_identity(entry): entry for entry in existentes}
    adicionados = 0
    atualizados = 0
    for nova in novas:
        if not nova.data_atualizacao:
            nova = DeParaEntry(
                rede=nova.rede,
                tipo_chave=nova.tipo_chave,
                chave_lida=nova.chave_lida,
                cnpj_oficial=nova.cnpj_oficial,
                matricula=nova.matricula,
                nome_loja=nova.nome_loja,
                status=nova.status,
                observacao=nova.observacao,
                fonte=nova.fonte,
                data_atualizacao=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        key = _entry_identity(nova)
        if key in mapa:
            atualizados += 1
        else:
            adicionados += 1
        mapa[key] = nova
    salvar_depara_clientes(mapa.values(), path)
    return adicionados, atualizados, path


def _ler_sheet_validacao(path_excel: Path) -> pd.DataFrame:
    for sheet in ("Cadastrar CNPJ", "Cadastrar CNPJ / GLN", "CADASTRAR_CNPJ", "CNPJ_SEM_MATRICULA"):
        try:
            return pd.read_excel(path_excel, sheet_name=sheet, dtype=str).fillna("")
        except Exception:
            continue
    raise ValueError("Nao encontrei aba 'Cadastrar CNPJ' no Excel de validacao.")


def importar_pendencias_validacao(path_excel: str | Path, path: Path | None = None) -> dict[str, object]:
    """Importa as linhas preenchidas da aba Cadastrar CNPJ para a base complementar.

    A rotina só grava linhas que tenham matrícula preenchida. Se o CNPJ oficial for
    preenchido, ele passa a ser o CNPJ de saída; se não for preenchido, mantém a
    chave lida para rastreabilidade e exige revisão posterior.
    """
    path_excel = Path(path_excel)
    if not path_excel.exists():
        raise FileNotFoundError(f"Excel de validacao nao encontrado: {path_excel}")
    df = _ler_sheet_validacao(path_excel)
    novas: list[DeParaEntry] = []
    for _, row in df.iterrows():
        row_dict = {str(k): clean_text(v) for k, v in row.to_dict().items()}
        status = clean_text(row_dict.get("Status") or row_dict.get("STATUS")).upper()
        if status in {"OK", "SEM_PENDENCIA"}:
            continue
        matricula = only_digits(row_dict.get("Matrícula") or row_dict.get("Matricula") or row_dict.get("Matrícula Encontrada"))
        if not matricula:
            continue
        chave = row_dict.get("Chave Lida") or row_dict.get("CNPJ Lido") or row_dict.get("CNPJ")
        cnpj_oficial = only_digits(row_dict.get("CNPJ Oficial") or row_dict.get("CNPJ_OFICIAL") or row_dict.get("CNPJ"))
        rede = row_dict.get("Rede/Layout") or row_dict.get("REDE") or ""
        if not chave and not cnpj_oficial:
            continue
        novas.append(
            DeParaEntry(
                rede=rede,
                tipo_chave=infer_tipo_chave(chave or cnpj_oficial, row_dict.get("Tipo da Chave")),
                chave_lida=normalize_key(chave or cnpj_oficial),
                cnpj_oficial=cnpj_oficial,
                matricula=matricula,
                status="ATIVO",
                observacao=row_dict.get("Observação") or row_dict.get("Observacao") or "Cadastro importado do Excel de validacao",
                fonte=path_excel.name,
                data_atualizacao=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        )
    adicionados, atualizados, destino = adicionar_ou_atualizar_entries(novas, path)
    return {
        "arquivo": str(destino),
        "linhas_lidas": len(df),
        "linhas_validas_para_importar": len(novas),
        "adicionados": adicionados,
        "atualizados": atualizados,
    }
