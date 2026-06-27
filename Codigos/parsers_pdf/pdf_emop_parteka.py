from __future__ import annotations

from parsers_pdf.pdf_generico_homologacao import ler_pdf_generico_homologacao


def ler_pdf_emop_parteka(caminho_arquivo: str, layout_config: dict, mapeamentos_df=None):
    """Parser controlado para Rede EMOP / Parteka.

    Até o momento não há amostra/código antigo suficiente para uma extração dedicada
    100% determinística. Por segurança, usa o motor genérico de homologação com
    referência explícita EMOP/PARTEKA, preservando todos os itens encontrados no
    Excel de validação e registrando alerta de conferência manual.
    """
    resultado = ler_pdf_generico_homologacao(
        caminho_arquivo,
        layout_config,
        mapeamentos_df,
        referencia="EMOP_PARTEKA_HOMOLOGACAO_CONTROLADA",
    )
    alertas = list(resultado.get("alertas", []) or [])
    alertas.append(
        "EMOP/PARTEKA em homologação controlada: não foi localizado código antigo/amostra suficiente para parser dedicado; conferir CNPJ, SKU/EAN, QTD e pedido antes de TXT/fila."
    )
    resultado["alertas"] = sorted({str(a) for a in alertas if str(a).strip()})
    return resultado
