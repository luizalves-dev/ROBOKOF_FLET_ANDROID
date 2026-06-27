# Relatório de limpeza e organização — Robô KOF

Autor: Kauê Melo
Data: 2026-05-29

## Objetivo

Organizar o pacote do Robô KOF com foco produtivo, reduzir resíduos operacionais e preservar os layouts que já funcionam.

## Limpezas aplicadas

- Removidos `__pycache__` e arquivos `.pyc`.
- Limpas entradas antigas de cliente em `Entradas_Clientes/`, mantendo somente estrutura e `.gitkeep`.
- Limpas saídas/históricos/logs antigos em `Resultados/`, mantendo somente estrutura e `.gitkeep`.
- Removida duplicidade do mapa de produtos da raiz de `Arquivos Base`, mantendo a cópia oficial em `Arquivos Base/Mapa de Produtos/`.
- Scripts de validação/teste movidos da raiz de `Codigos/` para `Codigos/scripts_manutencao/validadores_legados/`.
- Patches/testes legados Baklizi movidos para `Codigos/scripts_manutencao/patches_legados_baklizi/`.

## Grancoffee preservado e integrado

- Mantido parser dedicado: `Codigos/layouts/rede_grancoffee.py`.
- Mantido registrador técnico: `Codigos/layouts/registrar_grancoffee.py`.
- Cadastrado layout em `Cadastros/layouts.csv` para `EXCEL` e `MSG`.
- Cadastrada rastreabilidade em `Cadastros/rastreabilidade_layouts.csv`.
- Cadastrado mapeamento mínimo em `Cadastros/mapeamento_campos.csv`.
- Importador do Outlook passa a salvar corpo do e-mail no metadata `.outlook.json`, permitindo que a Grancoffee use a data oficial do corpo do e-mail mesmo quando só o anexo `.xlsm/.xlsx` é importado.
- Importação dedicada Grancoffee gera Excel de validação primeiro e não alimenta fila/TXT automaticamente.

## Arquivos removidos/limpos

Total de entradas operacionais removidas: 169.

Esses itens eram cache, saída antiga, histórico, logs ou duplicidade de base. Nenhum parser produtivo foi removido.
