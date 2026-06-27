# Códigos — RoboKOF Mobile

Esta pasta contém o motor de conversão e a interface Flet do RoboKOF.

## Organização

- `app_flet.py`: interface Android-first.
- `importador_clientes.py`: orquestração da importação e validação.
- `leitor_pdf_clientes.py` e `parsers_pdf/`: leitura dos layouts PDF.
- `leitor_excel_clientes.py` e `parsers_excel/`: leitura dos layouts Excel.
- `mobile_queue.py`: conferência e liberação segura para a fila.
- `operacao_service.py`: geração RoboKOF e TXT EDI.
- `layouts/rede_bh/`: processamento dedicado Rede BH.
- `pdfplumber/`: cópia incorporada do leitor PDF puro em Python, sem dependência obrigatória de PDFium.

## Regras de manutenção

- Não alterar o parser de uma rede para corrigir outra.
- Novos layouts devem permanecer isolados em `parsers_pdf/`, `parsers_excel/` ou `layouts/`.
- Toda saída deve passar pelo Excel de validação antes da fila/TXT.
- Não adicionar caminhos fixos do Windows nem dependências de Tkinter, SAP ou Outlook.

A entrada do aplicativo fica em `../main.py`.
