import shutil
from datetime import datetime

import config


def mover_para_historico():
    """
    Move os arquivos gerados para o histórico do dia:
    - Arquivos RoboKOF + TXT -> Historico/DD-MM-AAAA/Pedidos
    - Arquivos de erro -> Historico/DD-MM-AAAA/Erros
    """
    data_hoje = datetime.now().strftime("%d-%m-%Y")

    base_historico = config.ROOT_DIR / "Resultados" / "Historico" / data_hoje
    pasta_pedidos = base_historico / "Pedidos"
    pasta_erros = base_historico / "Erros"

    pasta_pedidos.mkdir(parents=True, exist_ok=True)
    pasta_erros.mkdir(parents=True, exist_ok=True)

    qtd_robokof = 0
    qtd_txt = 0
    qtd_erros = 0

    for arquivo in config.OUT_ROBOKOF_DIR.glob("*.xlsx"):
        destino = pasta_pedidos / arquivo.name
        if arquivo.exists():
            shutil.move(str(arquivo), str(destino))
            qtd_robokof += 1

    for arquivo in config.OUT_TXT_DIR.glob("*.txt"):
        destino = pasta_pedidos / arquivo.name
        if arquivo.exists():
            shutil.move(str(arquivo), str(destino))
            qtd_txt += 1

    for arquivo in config.OUT_ERRO_DIR.glob("*.xlsx"):
        destino = pasta_erros / arquivo.name
        if arquivo.exists():
            shutil.move(str(arquivo), str(destino))
            qtd_erros += 1

    return {
        "data": data_hoje,
        "pasta_pedidos": str(pasta_pedidos),
        "pasta_erros": str(pasta_erros),
        "robokof_movidos": qtd_robokof,
        "txt_movidos": qtd_txt,
        "erros_movidos": qtd_erros,
    }