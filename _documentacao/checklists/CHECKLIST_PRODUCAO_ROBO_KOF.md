# Checklist de produção — Robô KOF

Antes de usar em produção:

1. Executar `python Codigos/scripts_manutencao/validar_projeto_robo_kof.py`.
2. Confirmar `STATUS_FINAL;APROVADO` no relatório gerado em `Resultados/Logs`.
3. Processar o pedido e revisar o Excel em `Resultados/pedidos_a_validar`.
4. Validar abas obrigatórias:
   - Modelo Robô KOF para Enviar;
   - Validação do Pedido;
   - Alertas/Erros ou Alertas e Erros;
   - Cadastrar CNPJ quando houver pendência cadastral;
   - Logs/Resumo quando aplicável.
5. Gerar TXT/fila KOF somente depois da validação manual.

Observação: para Grancoffee, conferir principalmente o cruzamento **matrícula + pedido + data da remessa do corpo do e-mail**.
