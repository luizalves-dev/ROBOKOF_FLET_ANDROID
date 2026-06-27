# Correção Grancoffee - corpo do e-mail e metadata Outlook

## Problema corrigido
Quando a Grancoffee era processada a partir dos anexos `.xlsm` importados do Outlook, a rotina tentava ler o arquivo `<anexo>.outlook.json`, mas o módulo `json` não estava importado em `Codigos/layouts/rede_grancoffee.py`. Com isso, o log registrava:

`Metadata Outlook inválida para Grancoffee: name 'json' is not defined`

Como consequência, a data oficial da remessa não era carregada do corpo do e-mail e o Excel ficava com alerta de data não encontrada.

## Correções aplicadas
- Importação explícita de `json` no layout Grancoffee.
- Normalização do corpo de e-mail em texto puro e HTML.
- Leitura mais robusta da tabela de remessa mesmo quando o Outlook compacta a linha da tabela.
- Deduplicação de remessas por matrícula + pedido + data, evitando repetição quando o mesmo corpo acompanha vários anexos.
- Fallback seguro para tentar capturar o e-mail selecionado/aberto no Outlook quando o usuário processar somente os anexos Excel.
- Contagem das linhas geradas no `Modelo Robô KOF para Enviar` dentro do resultado da importação, para não parecer que o lote ficou vazio.

## Regra preservada
A data oficial continua sendo a DATA DA REMESSA do corpo do e-mail, cruzada por matrícula + pedido. A data interna do anexo permanece como rastreabilidade.
