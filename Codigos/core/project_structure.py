# Autor: Kauê Melo
"""Estrutura central do Robô KOF.

Este módulo não altera o fluxo dos layouts. Ele apenas centraliza nomes de pastas
para documentação, manutenção e validação técnica.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class RoboKOFFolders:
    root: Path

    @property
    def codigos(self) -> Path: return self.root / "Codigos"
    @property
    def cadastros(self) -> Path: return self.root / "Cadastros"
    @property
    def arquivos_base(self) -> Path: return self.root / "Arquivos Base"
    @property
    def entradas(self) -> Path: return self.root / "Entradas_Clientes"
    @property
    def resultados(self) -> Path: return self.root / "Resultados"
    @property
    def documentacao(self) -> Path: return self.root / "_documentacao"

    def required_dirs(self) -> list[Path]:
        return [
            self.codigos,
            self.cadastros,
            self.arquivos_base,
            self.entradas / "Excel",
            self.entradas / "PDF",
            self.entradas / "Outlook",
            self.resultados / "Arquivos Fila",
            self.resultados / "Arquivos RoboKOF",
            self.resultados / "Arquivos TXT EDI",
            self.resultados / "Arquivos ERRO",
            self.resultados / "Logs",
            self.resultados / "pedidos_a_validar",
            self.resultados / "pedidos_validados",
            self.resultados / "pedidos_com_erro",
            self.resultados / "validacoes_clientes",
            self.resultados / "historico_importacoes",
            self.resultados / "temp",
        ]

    def required_files(self) -> list[Path]:
        return [
            self.root / "requirements.txt",
            self.cadastros / "clientes.csv",
            self.cadastros / "layouts.csv",
            self.cadastros / "mapeamento_campos.csv",
            self.cadastros / "rastreabilidade_layouts.csv",
            self.cadastros / "de_para_clientes.csv",
            self.cadastros / "regras_conversao.csv",
            self.arquivos_base / "BASE de GLNS.xlsx",
            self.arquivos_base / "Exemplo envio.xlsx",
        ]


def get_project_root(start: Path | None = None) -> Path:
    """Retorna a raiz do projeto a partir de um arquivo dentro de Codigos."""
    base = (start or Path(__file__)).resolve()
    for parent in [base.parent, *base.parents]:
        if (parent / "Codigos").exists() and (parent / "Cadastros").exists():
            return parent
    return Path(__file__).resolve().parents[2]
