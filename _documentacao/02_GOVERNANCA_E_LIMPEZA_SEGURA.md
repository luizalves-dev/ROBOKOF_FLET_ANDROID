# Governança e limpeza segura

Não apagar manualmente arquivos sem classificar.

## Pode ficar fora do pacote limpo

- Histórico antigo de `Resultados/`
- PDFs/Excels de entrada já processados
- `__pycache__`
- `*.pyc`
- backups `.bak*`
- patches incrementais antigos em TXT
- ZIPs antigos dentro do projeto

## Nunca apagar sem backup

- `Codigos/`
- `Cadastros/`
- `Arquivos Base/`
- `requirements.txt`
- `.env.example`

## Processo recomendado

1. Rodar diagnóstico.
2. Conferir relatório.
3. Mover para quarentena.
4. Testar o robô.
5. Só depois descartar definitivamente.
