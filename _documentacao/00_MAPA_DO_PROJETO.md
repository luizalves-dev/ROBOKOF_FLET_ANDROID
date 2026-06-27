# Mapa do projeto — Robô KOF

## Pastas principais

- `Codigos/`: código-fonte.
- `Cadastros/`: CSVs oficiais de configuração.
- `Arquivos Base/`: bases de apoio pesadas e gabaritos.
- `Entradas_Clientes/`: arquivos recebidos para processamento.
- `Resultados/`: saídas geradas pelo robô.
- `_documentacao/`: documentação operacional e técnica.
- `_quarentena_segura/`: local reservado para limpeza controlada.

## Onde cadastrar novos layouts

- PDF: `Codigos/parsers_pdf/pdf_nome_rede.py`
- Excel: `Codigos/parsers_excel/excel_nome_rede.py`
- Cadastro: `Cadastros/layouts.csv`
- Mapeamento: `Cadastros/mapeamento_campos.csv`
- Rastreabilidade: `Cadastros/rastreabilidade_layouts.csv`
- De/para: `Cadastros/de_para_clientes.csv`

## Organização de produção aplicada

- `Codigos/`: mantém apenas o motor produtivo e módulos de layout.
- `Codigos/scripts_manutencao/validadores_legados/`: validadores e diagnósticos antigos fora da raiz produtiva.
- `Codigos/scripts_manutencao/patches_legados_baklizi/`: scripts pontuais preservados fora dos parsers produtivos.
- `Resultados/` e `Entradas_Clientes/`: entregues limpos, com estrutura pronta para uso.
- Grancoffee: layout dedicado em `Codigos/layouts/rede_grancoffee.py`, cadastrado para Excel/MSG e com metadata Outlook para ler o corpo do e-mail.
