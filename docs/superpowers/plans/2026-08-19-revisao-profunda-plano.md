# Plano — Revisão profunda e sistemática dos 6 instrumentos

## Objetivo

Ir além das revisões ad-hoc já feitas nesta sessão (checklist de ~19
perguntas, achados soltos por instrumento) e cobrir sistematicamente
categorias inteiras de erro que já provaram existir em mais de um lugar:
compressão de conteúdo (Obrigações, Rescisão), numeração/alíneas órfãs
(item 10 do TR-Bens, Seção 14 da Nota Técnica), e diagramação de tabelas
divergente do anexo oficial (Matriz de Riscos). A ideia é procurar essas
mesmas categorias de erro em todo lugar que ainda não foi olhado, em vez
de esperar você apontar cada uma.

## Por que um plano antes de rodar

O workflow de ~19 perguntas já mostrou que perguntas abertas demais geram
achados difíceis de agir (ex.: "aditivos" sem base clara). Este plano
divide o trabalho em buscas objetivas, cada uma com um critério concreto
de "achei / não achei", para que os achados saiam prontos para virar
correção, não apenas uma lista de dúvidas.

## Fase 1 — Auditoria de numeração e alíneas órfãs (6 agentes, 1 por instrumento)

Cada agente varre o `.py` de um instrumento inteiro procurando:
- Alíneas (a, b, c...) que começam no meio do alfabeto sem as anteriores
  (o padrão exato do bug do item 10 do TR-Bens — "h)" sem a) a g)).
- Itens numerados (1., 2., 3...) com saltos ou repetições.
- Cabeçalhos do formulário (`js/dados-*.js`, id da seção) que não batem
  com o número do título gerado no `.docx` (o padrão exato do bug da
  Seção 14 da Nota Técnica).
- Referências cruzadas ("conforme item X.Y") que apontam para um item
  que não existe mais ou tem número diferente no texto atual.

Saída: lista de {instrumento, localização exata (arquivo:linha),
descrição do problema, sugestão de correção}.

## Fase 2 — Restauração de Obrigações da Contratante/Contratada (4 agentes, 1 por instrumento com esse conteúdo)

Já confirmado como bug em TR-Bens (achado original desta sessão, ainda
não corrigido). Cada agente:
1. Lê a Seção de Obrigações da Contratante e da Contratada, tanto no
   `.py` quanto na minuta-fonte SEI correspondente.
2. Compara item a item: toda alínea da fonte está presente no gerado?
3. Se houver compressão (prosa substituindo lista de alíneas), produz o
   texto completo já formatado em alíneas, pronto para colar no `.py`,
   igual ao que foi feito manualmente para a Rescisão.

Saída: para cada instrumento com lacuna, o texto de substituição pronto
(não apenas "está faltando X").

## Fase 3 — Comparação visual/estrutural dos anexos restantes contra os `.odt` oficiais (6-8 agentes)

Expande o método usado para a Matriz de Riscos (extrair `.odt` com
marcação de tabela/linha/célula e comparar contra o `.py`) para os
anexos ainda não comparados dessa forma:
- Planilha de Custos e Formação de Preços
- Declaração de Contratos Firmados
- Mapa Comparativo de Propostas
- Contrato Simplificado
- Termo de Cooperação Técnica da Conta Vinculada
- IMR (Instrumento de Medição de Resultado) — conferir se a correção
  desta sessão manteve fidelidade total à estrutura oficial

Cada agente roda o mesmo script Python de extração usado para a Matriz
de Riscos (`table:table` / `table:table-row` / `table:table-cell` como
marcadores) no `.odt` correspondente, e compara coluna a coluna, seção a
seção, contra o que o `.py` gera hoje.

## Fase 4 — Consistência cruzada entre instrumentos (2-3 agentes)

Para cláusulas que deveriam dizer a mesma coisa em instrumentos
diferentes (reajuste/repactuação, garantia, rescisão, dotação
orçamentária, sanções, nepotismo), um agente compara o texto gerado
lado a lado entre TR-Bens/Serviços/Dispensa/PB e aponta divergências não
justificadas pela natureza do objeto (bens x serviços x obras) — ex.: um
instrumento cita um artigo que outro não cita para a mesma regra, ou um
tem uma exceção que falta no outro.

## Fase 5 — Verificação cética dos achados (pipeline, um verificador por achado)

Mesmo padrão já usado nesta sessão: cada achado das Fases 1-4 é conferido
por um segundo agente independente, que releem a fonte primária (RLC ou
minuta/`.odt`) antes de confirmar. Só achados confirmados chegam ao
relatório final.

## O que fica fora deste plano

- SRP/TIC (spec e plano já existem, só falta executar).
- "Aditivos contratuais" no TR-Serviços — sem base clara na fonte, seguirá
  em aberto até haver uma referência concreta.
- Qualquer sugestão de valores numéricos (multas, prazos) sem grounding
  em fonte real — a Fase 2/3 só produz texto quando há uma fonte real
  para transcrever.

## Estimativa de escala

~20-25 agentes de busca/comparação (Fases 1-4) + verificação cética por
achado confirmado (Fase 5, tipicamente 15-30 agentes adicionais,
dependendo de quantos achados aparecerem). Maior que os workflows
anteriores desta sessão, proporcional ao pedido de revisão "bem feita e
aprofundada".
