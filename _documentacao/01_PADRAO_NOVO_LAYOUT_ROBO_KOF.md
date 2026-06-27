# Padrão oficial para cadastrar nova rede/layout

## Antes de programar

1. Identificar a rede.
2. Identificar tipo de arquivo: PDF, XLS, XLSX, XLSM.
3. Verificar se tem múltiplos pedidos no mesmo arquivo.
4. Verificar se tem múltiplas páginas/abas.
5. Localizar número do pedido.
6. Localizar CNPJ correto da loja, ignorando CNPJ fornecedor/SPAL.
7. Definir se o código oficial é SKU, EAN, Ref., Cod Forn ou Código Externo.
8. Definir se a quantidade já vem em caixaria ou precisa conversão.
9. Mapear de/para CNPJ x matrícula.
10. Definir alertas e pendências.

## Saída obrigatória

O Excel de validação precisa conter, no mínimo:

- `Modelo Robô KOF para Enviar`
- `Validação do Pedido`
- `Alertas/Erros`
- `Cadastrar CNPJ`, quando aplicável
- logs/pendências quando aplicável

## Regra de segurança

Erro cadastral, GLN pendente, matrícula bloqueada ou CNPJ sem de/para não deve esvaziar o modelo se houver campos estruturais mínimos: matrícula ou CNPJ, SKU, quantidade e pedido.
