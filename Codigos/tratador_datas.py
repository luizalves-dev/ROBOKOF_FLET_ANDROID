from __future__ import annotations

from datetime import datetime, timedelta
from typing import Optional

DIAS_SEMANA = {
    "SEGUNDA": 0,
    "SEGUNDA-FEIRA": 0,
    "TERCA": 1,
    "TERÇA": 1,
    "TERCA-FEIRA": 1,
    "TERÇA-FEIRA": 1,
    "QUARTA": 2,
    "QUARTA-FEIRA": 2,
    "QUINTA": 3,
    "QUINTA-FEIRA": 3,
    "SEXTA": 4,
    "SEXTA-FEIRA": 4,
    "SABADO": 5,
    "SÁBADO": 5,
    "DOMINGO": 6,
}


def formatar_data_fila(data_obj: datetime) -> str:
    return data_obj.strftime("%d.%m.%Y")


def ajustar_se_domingo(data_obj: datetime) -> datetime:
    return data_obj + timedelta(days=1) if data_obj.weekday() == 6 else data_obj


def calcular_dmais1(data_base: Optional[datetime] = None) -> datetime:
    base = data_base or datetime.today()
    return ajustar_se_domingo(base + timedelta(days=1))


def interpretar_dia_semana(texto: str) -> Optional[int]:
    if not texto:
        return None
    chave = str(texto).strip().upper()
    return DIAS_SEMANA.get(chave)


def proxima_ocorrencia_dia_semana(texto_dia: str, data_base: Optional[datetime] = None) -> Optional[datetime]:
    base = data_base or datetime.today()
    dia_alvo = interpretar_dia_semana(texto_dia)
    if dia_alvo is None:
        return None
    dias_a_frente = (dia_alvo - base.weekday()) % 7
    if dias_a_frente == 0:
        dias_a_frente = 7
    return ajustar_se_domingo(base + timedelta(days=dias_a_frente))


def tentar_ler_data_texto(valor: str) -> Optional[datetime]:
    if not valor:
        return None
    texto = str(valor).strip()
    formatos = ["%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]
    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            continue
    return None


def resolver_data_entrega(regra_data_entrega: str, valor_lido: str | None = None, data_base: Optional[datetime] = None) -> Optional[str]:
    regra = (regra_data_entrega or "").strip().upper()
    if regra == "D+1":
        return formatar_data_fila(calcular_dmais1(data_base))
    if regra == "COLUNA":
        if not valor_lido:
            return None
        data = tentar_ler_data_texto(valor_lido)
        if data is None:
            data = proxima_ocorrencia_dia_semana(str(valor_lido), data_base)
        if data is None:
            return None
        return formatar_data_fila(ajustar_se_domingo(data))
    return None
