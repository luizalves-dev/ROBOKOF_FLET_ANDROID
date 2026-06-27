# Relatório de reorganização — 2026-05-27

## O que foi melhorado

- Projeto entregue como base operacional limpa.
- Pastas de entrada e saída recriadas vazias e padronizadas.
- Históricos e resultados antigos removidos do pacote limpo.
- Caches Python removidos.
- Backups e patches incrementais antigos removidos.
- README principal refeito.
- Scripts de manutenção adicionados.
- Template oficial para novos layouts adicionado.
- Integração Iquegami preservada e com de/para mais robusto.

## O que não foi mexido

- Lógica dos layouts PDF/Excel existentes.
- Cadastros oficiais em CSV.
- Bases oficiais em `Arquivos Base`.
- Interface principal `app_gui.py`, que já possui responsividade por redimensionamento.

## Risco controlado

O pacote não inclui histórico operacional antigo. Caso precise consultar algo histórico, manter o ZIP anterior como arquivo morto separado.
