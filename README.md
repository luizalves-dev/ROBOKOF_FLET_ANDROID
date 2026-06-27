# RoboKOF Mobile — Flet / Android

Aplicativo local para converter pedidos de clientes em arquivos de validação, fila RoboKOF e TXT EDI. Esta edição substitui integralmente a interface Tkinter por Flet e foi estruturada para empacotamento como APK.

## O que foi mantido

- Leitura de pedidos em PDF e Excel.
- Identificação automática e seleção manual de layout.
- Todos os 43 layouts ativos cadastrados no projeto.
- De/para de CNPJ, matrícula, EAN e SKU.
- Extração, normalização, consolidação e auditoria dos itens.
- Geração do Excel de validação antes da fila.
- Reimportação do Excel validado com bloqueio de linhas inseguras.
- Fila operacional RoboKOF.
- Geração dos arquivos RoboKOF e TXT EDI.
- Fluxo dedicado da Rede BH, incluindo lote, duplicidades, consolidado e pacote ZIP.
- Bases e cadastros locais incorporados ao aplicativo.

## O que foi removido

- Automação SAP.
- Integração e automação Outlook.
- Dependências exclusivas do Windows: Tkinter, pywin32, pythoncom e automação de tela.

## Fluxo do aplicativo

1. Abra **Converter** e selecione um ou mais PDFs/Excels.
2. Confirme o layout sugerido ou selecione o layout correto.
3. Toque em **Processar**.
4. Abra/compartilhe o Excel gerado em **Saídas** e faça a conferência.
5. Reimporte o Excel validado e confirme a inclusão das linhas seguras na fila.
6. Em **Operação**, gere os arquivos RoboKOF e os TXTs.
7. Compartilhe os resultados pelo seletor nativo do Android.

Nenhum arquivo lido é enviado diretamente para TXT. A validação intermediária permanece obrigatória.

## Estrutura principal

```text
ROBOKOF_FLET_ANDROID/
├── main.py                       # entrada Flet
├── bootstrap.py                  # armazenamento local e instalação das bases
├── Codigos/
│   ├── app_flet.py               # interface Android-first
│   ├── importador_clientes.py    # orquestração dos layouts
│   ├── leitor_pdf_clientes.py    # leitura de PDFs
│   ├── leitor_excel_clientes.py  # leitura de planilhas
│   ├── mobile_queue.py           # validação segura antes da fila
│   ├── operacao_service.py       # RoboKOF + TXT
│   └── layouts/rede_bh/          # fluxo dedicado BH
├── assets/
│   ├── icon.png
│   └── bootstrap/                # bases/cadastros copiados na primeira execução
├── pyproject.toml
├── requirements.txt
├── build_apk.bat
└── build_apk.sh
```

## Executar no computador

Recomendado: Python 3.12 ou 3.13.

### Windows PowerShell

```powershell
cd C:\caminho\ROBOKOF_FLET_ANDROID
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

### Linux / WSL

```bash
cd /caminho/ROBOKOF_FLET_ANDROID
python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

## Gerar o APK

O projeto já contém metadados, ícone, dependências e scripts de build.

### Windows

```powershell
.\build_apk.bat
```

### Linux / WSL

```bash
./build_apk.sh
```

Ou diretamente:

```bash
flet build apk . --yes
```

O APK é criado em `build/apk/`. No primeiro build, o Flet pode instalar o Flutter SDK necessário.

## Armazenamento no Android

Na primeira execução, as bases de `assets/bootstrap` são copiadas para a área gravável privada do aplicativo. As pastas operacionais e os resultados são criados nessa área. O usuário acessa os arquivos por **Compartilhar** ou **Salvar**, sem depender de caminhos do Windows.

## Leitura de PDF no APK

A extração textual, de palavras e de tabelas usa `pdfplumber`/`pdfminer.six`, incorporados de forma compatível com o empacotamento móvel. `PyMuPDF` não é obrigatório.

Limitação conhecida: PDF composto apenas por imagem, sem camada de texto, exigia Tesseract instalado externamente no Windows. Esse OCR externo não faz parte do APK. PDFs textuais e planilhas, que representam o fluxo normal dos layouts cadastrados, permanecem suportados.

## Segurança operacional

O módulo `mobile_queue.py` bloqueia linhas com:

- matrícula ausente ou pendente;
- SKU ausente ou pendente;
- quantidade inválida, fracionada ou menor/igual a zero;
- pedido ou data de remessa ausentes;
- conversão pendente;
- status manual de validação reprovado, quando presente.

## Observações

- Arquivos históricos de pedidos de clientes não foram incluídos no pacote final.
- Alterações futuras de bases podem ser feitas dentro do armazenamento do app ou substituindo os arquivos em `assets/bootstrap` antes de um novo build.
- O diagnóstico interno verifica cadastros, bases essenciais, layouts e módulos de leitura/geração.
