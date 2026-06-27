# Ajuste Grancoffee - Data Remessa com ponto

## Ajuste aplicado
A saída da Grancoffee foi ajustada para gravar a Data Remessa no padrão `dd.mm.aaaa`, por exemplo `30.05.2026`, em vez de `dd/mm/aaaa`.

## Motivo
O Robô KOF pode quebrar ou interpretar incorretamente a data quando a coluna vem com `/`. Para evitar isso, o layout agora grava a data como texto no Excel de validação.

## Onde foi aplicado
- Aba `Modelo Robô KOF para Enviar`: coluna `DATA REMESSA`.
- Aba `Validação do Pedido`: `Data Entrega Anexo` e `Data Remessa E-mail`.
- Aba `Resumo por Loja`: `Data Remessa`.
- Aba `Email x Remessa`: `Data Remessa`.

## Regra preservada
A data oficial continua sendo a DATA DA REMESSA do corpo do e-mail, cruzada por matrícula + pedido. A data do anexo continua sendo apenas rastreabilidade.
