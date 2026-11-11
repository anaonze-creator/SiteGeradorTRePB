# Plano consolidado — vedações, modelo de proposta, tabela solta, DEMO, avaliação de imóveis, numeração

> **Para quem for executar:** este plano cobre 7 frentes de tamanhos muito diferentes — da correção de uma linha até a criação de um instrumento novo. Cada fase é independente e pode ser executada, revisada e commitada separadamente. As fases 1-3 são as mais baratas/seguras; as fases 4-6 exigem mais decisão de conteúdo; a fase 7 é uma decisão arquitetural que merece uma conversa própria antes de começar.

**Contexto:** levantamento feito em cima de:
- `documentos de referencia/rlc.txt` (RLC 10.901 completo)
- `documentos de referencia/SEI_43738970_..._Projeto_Basico__Licitacao_.pdf` — PB real de manutenção predial com posto de mão de obra residente (dedicação exclusiva)
- `documentos de referencia/SEI_55236534_..._Termo_de_Referencia__TR____Servicos.pdf` — TR real de laudo de avaliação de imóveis
- Os 4 templates Python já existentes (`templates/build_template_*.py`) e os achados da auditoria contra os checklists do RLC (Anexos VIII-A/B/C e IX) feita na rodada anterior

---

## Fase 1 — Vedações à Contratada (todos os 6 instrumentos)

**Problema:** só o TR-Dispensa (ramo Contrato Simplificado) tem uma seção "DAS VEDAÇÕES" própria, com apenas 4 itens (caução/operação financeira, interromper execução, trabalho de menor, subcontratação total). TR-Bens, TR-Serviços e o corpo do PB não têm nenhuma seção equivalente — o conteúdo de vedação fica disperso (um pouco em Obrigações, um pouco em Sanções) e incompleto.

**Fonte no RLC:** art. 3º, XXVIII/XXI (contratos devem prever "cláusula com as vedações à contratada, incluindo a proibição ao nepotismo"), e o trecho do art. 155 sobre obrigações da Contratada em obras/serviços de engenharia (linhas 4780-4809 de `rlc.txt`), que já traz item d) "é vedado à contratada que seus empregados solicitem serviços, materiais ou equipamentos às empresas terceirizadas que prestam serviços à Conab" — hoje ausente em todos os nossos templates. O exemplo real `SEI_43738970` confirma esse item (18.3) e mais vedações típicas de mão de obra exclusiva (ver Fase 4).

**Vedações padrão a consolidar** (a reunir num único bloco "DAS VEDAÇÕES" por instrumento, citando o artigo correto do RLC para cada uma):
1. Caucionar ou utilizar o Contrato para qualquer operação financeira (já existe no TR-Dispensa)
2. Interromper a execução contratual sob alegação de inadimplemento da Contratante, salvo previsão legal (já existe)
3. Empregar menor de 18 anos em trabalho noturno/perigoso/insalubre; empregar menor de 16 anos, salvo aprendiz a partir de 14 (já existe, mas hoje só no Dispensa — nos outros instrumentos essa vedação está "perdida" dentro de Obrigações, sem o rótulo "vedação")
4. Subcontratar total ou parcialmente fora do permitido no instrumento (já existe)
5. **Novo:** vedar que empregados da contratada solicitem serviços/materiais/equipamentos a empresas terceirizadas que prestam serviço à Conab (art. 155/RLC, confirmado no exemplo real)
6. **Novo:** vedação ao nepotismo — hoje já existe como seção própria ("DA VEDAÇÃO AO NEPOTISMO"); decidir se species/mantém separada ou se passa a ser um item dentro de "DAS VEDAÇÕES" (recomendo manter separada, é mais visível e já está testada)

**Ação por instrumento:**
- `templates/build_template_tr_bens.py`: criar nova seção "DAS VEDAÇÕES" (renumerando as seções seguintes) com os itens 1-5 aplicáveis a bens.
- `templates/build_template_tr_servicos.py`: idem.
- `templates/build_template_tr_dispensa.py`: expandir a seção 22 já existente (corpo + réplica no Anexo VI) com o item novo (5).
- `templates/build_template_pb.py`: criar seção "DAS VEDAÇÕES" no corpo (hoje só existe um equivalente fraco na Minuta de Contrato, Cláusula Vigésima) e reforçar a cláusula da Minuta de Contrato.
- Testar: gerar os 4 templates e conferir que a nova seção aparece uma vez só, sem duplicar conteúdo já existente em Obrigações.

**Tamanho estimado:** médio — 4 arquivos, conteúdo bem definido, sem necessidade de novos campos de formulário (tudo é texto fixo).

---

## Fase 2 — Revisão do Modelo de Apresentação de Proposta (todos os 6 instrumentos)

**Problema:** o conteúdo já foi implementado numa rodada anterior desta sessão, mas o usuário pede nova revisão — provavelmente porque, com todas as correções de numeração/conteúdo feitas depois, pode ter ficado defasado ou inconsistente entre instrumentos.

**Ação:**
1. Ler o Anexo de Modelo de Apresentação de Proposta de cada um dos 4 templates principais (TR-Bens Anexo II, TR-Serviços Anexo II, TR-Dispensa Anexo II/Seção 5.1.1, PB Anexo III) lado a lado.
2. Conferir consistência de conteúdo obrigatório (todos devem ter, quando aplicável): descrição do objeto, quantidades, valor unitário/global (conforme o caso), prazo de entrega/execução, prazo de vigência, prazo de garantia (mín. 60 dias), prazo de validade da proposta (mín. 60 dias), custos inclusos, dados da proponente — isso já foi pedido e implementado, então o objetivo aqui é achar **divergências de redação entre instrumentos que deveriam dizer a mesma coisa** e **campos que a tabela de itens tem num instrumento e não tem em outro sem motivo** (ex.: TR-Bens tem coluna "Valor Unitário", TR-Serviços não — já sabido e correto, mas confirmar se todos os casos análogos seguem a mesma lógica).
3. Conferir se a tabela de itens de cada anexo bate com o que o RLC/minuta-fonte pede (nº de colunas, cabeçalhos exatos) — a Fase de auditoria anterior já sinalizou isso parcialmente.

**Tamanho estimado:** pequeno a médio — é revisão + ajustes pontuais, não reconstrução.

---

## Fase 3 — Corrigir a tabela solta de "Endereço" / "Observação"

**Causa raiz identificada:** o campo `tabela-linhas` (endereços adicionais, usado em TR-Bens/TR-Serviços/TR-Dispensa Seção 6) sempre imprime a frase de introdução ("Quando o objeto envolver a entrega/execução em mais de um endereço, os locais adicionais são os seguintes:") **e** a tabela com o cabeçalho "Endereço | Observação", mesmo quando o usuário não adicionou nenhuma linha — porque `inserirLinhasTabela` (`js/gerador-docx.js`) só remove a linha-modelo quando `linhas` está vazio, mas nunca remove a tabela inteira nem a frase de introdução. Resultado: documento comum (sem endereços extras) mostra uma tabela vazia, só com cabeçalho, pendurada no meio da Seção 6.

**Correção proposta:**
- Envolver a frase de introdução + a tabela (cabeçalho incluso) num bookmark próprio do campo tabela-linhas (ex.: `${campo.id}_BLOCO`), do jeito que já existe pra `_LINHA_MODELO`.
- Em `gerador-docx.js`, quando `linhas.length === 0`, remover o bloco inteiro (frase + tabela) em vez de só deixar o cabeçalho sem linhas.
- Repetir a correção nos 3 templates que usam essa tabela (TR-Bens, TR-Serviços, TR-Dispensa) e no template genérico de infrações extras (Seção 15/16, mesmo padrão `tabela-linhas`) se sofrer do mesmo problema — checar `S15_InfracoesExtras`/`S16_InfracoesExtras` também.
- Adicionar teste automatizado cobrindo "0 linhas → tabela e frase somem" e "≥1 linha → tabela e frase aparecem".

**Tamanho estimado:** pequeno, mas mexe no motor (`js/gerador-docx.js`) — testar bem antes de tocar nos 3 templates.

---

## Fase 4 — Robustecer Obrigações da Contratada em contratos com Dedicação Exclusiva de Mão de Obra (DEMO)

**O que o exemplo `SEI_43738970` ensina:** o PB real de manutenção predial tem uma Seção 3.5 declarando explicitamente o enquadramento como mão de obra exclusiva (art. 19 do RLC, com 3 critérios objetivos: empregados à disposição nas dependências da Conab, não compartilhamento de recursos entre contratos, e possibilidade de fiscalização pela Conab sobre distribuição/controle/supervisão dos RH). A partir daí, a Seção 18 (Obrigações da Contratada) tem **48 alíneas** (I a XLVIII), a maior parte delas específica de DEMO e hoje **ausente** do nosso template:

- REP biométrico de ponto (art. 74, §2º, CLT) com acesso da fiscalização
- Substituição de empregado em até 2h por falta/atraso; substituição automática após 3 advertências
- Comunicação prévia de 30 dias de férias e quem vai substituir
- Exclusão do Simples Nacional por cessão de mão de obra (LC 123/2006, art. 18, §5º-C) + comprovação do ofício à Receita Federal
- Cartão Cidadão CEF em até 60 dias; acesso dos empregados ao extrato de contribuições previdenciárias
- Autorização da Conab para descontar e pagar salários/FGTS/INSS diretamente aos empregados em caso de inadimplência da contratada, com retenção cautelar em juízo quando não for possível pagar diretamente
- Retenção de fatura proporcional ao inadimplemento das obrigações trabalhistas/FGTS, com prazo de regularização de 15 dias e notificação ao sindicato da categoria
- Vedação a que empregados solicitem serviços/materiais a terceirizadas da Conab (ver Fase 1)
- Plantão de emergência 24h
- Vedação a "dobra" de turno pelo mesmo empregado

**Ação:**
1. Confirmar no RLC os artigos que amparam cada item acima (a maioria decorre da legislação trabalhista/CLT/LC 123 citada diretamente no exemplo, não do RLC propriamente — então a citação correta é a lei específica, não "art. X do RLC").
2. Criar um novo campo de gate: `S_ExigeDedicacaoExclusiva` (Sim/Não) em TR-Serviços, TR-Dispensa e PB — hoje esses instrumentos só têm o reajuste/repactuação amarrado indiretamente à DEMO (`S12_TipoCorrecao`/`S13_TipoCorrecao` = "Repactuação"), mas não uma declaração formal da Seção 3/Caracterização do Objeto nem um bloco de obrigações reforçado.
3. Quando "Sim": acrescentar, na Seção "Caracterização do Objeto", a declaração dos 3 critérios do art. 19 do RLC (igual ao 3.5.1-3.5.3 do exemplo); e, na Seção de Obrigações da Contratada, acrescentar o bloco de alíneas específicas de DEMO listado acima (adaptando numeração/lugar conforme o instrumento).
4. Avaliar se vale a pena reaproveitar o `S12_TipoCorrecao`/`S13_TipoCorrecao` = "Repactuação" como o próprio gate (em vez de criar um campo novo), para não duplicar a pergunta ao usuário — **recomendo isso**, já que semanticamente repactuação só se aplica quando há DEMO.

**Tamanho estimado:** grande — é o item de conteúdo mais substancial do plano. Fazer instrumento por instrumento (TR-Serviços primeiro, por já ter a distinção reajuste/repactuação pronta; depois TR-Dispensa, que tem a mesma estrutura; PB por último, que hoje nem tem esse eixo Reajuste x Repactuação declarado explicitamente na Caracterização do Objeto).

---

## Fase 5 — Novo modelo: TR-Dispensa para Avaliação de Imóveis

**O que o exemplo `SEI_55236534` mostra:** é um Termo de Referência de contratação de laudo de avaliação imobiliária, com estrutura muito específica: exigências técnicas do laudo citando a NBR 14653 (partes 1 e 2) e a Norma de Alienação de Bens Imóveis da Conab (NOC 60.208) artigo por artigo — conteúdo completamente diferente de qualquer coisa que já modelamos (nenhum dos 6 instrumentos atuais cobre esse objeto). O próprio documento nota que a contratação pode se enquadrar em dispensa de licitação por valor (linha 1102 do texto extraído).

**Minha avaliação:** faz sentido como um **modelo específico** — mas antes de decidir se é (a) um TR-Dispensa totalmente novo e independente, ou (b) uma variante/`blocoTexto` dentro do TR-Dispensa já existente (ativada por um campo tipo "Esta contratação é de laudo de avaliação imobiliária?"), preciso entender melhor o tamanho da diferença. Pelo que vi no trecho lido, a Seção 1 (Objeto) já é estruturalmente diferente (tabela de item único com uma lista extensa de exigências técnicas do laudo, citando NOC 60.208 alínea por alínea) — isso não cabe bem como "mais um campo opcional" dentro do TR-Dispensa genérico; é conteúdo extenso e de domínio muito específico (perícia de avaliação imobiliária).

**Recomendação:** tratar como um **7º instrumento independente** (`TR-Dispensa — Avaliação de Imóveis`, ou nome similar), reaproveitando a infraestrutura do TR-Dispensa (Contrato Simplificado, Sanções, Rescisão, etc. — tudo isso é genérico e já pronto) mas com Seção 1 (Objeto) e a lista de exigências técnicas do laudo (NOC 60.208 + NBR 14653) como conteúdo fixo próprio.

**Isto é grande o suficiente para merecer seu próprio ciclo de brainstorming → spec → plano**, não cabe bem dentro deste plano consolidado. Sugiro que, quando chegar a vez desta fase, eu leia o `SEI_55236534` inteiro (só li o início) e traga uma proposta de design específica antes de qualquer código.

**Tamanho estimado:** grande, novo instrumento — escopo próprio.

---

## Fase 6 — Priorização dos gaps do checklist do RLC (retomado da rodada anterior)

Já levantado na auditoria anterior; ação aqui é só **decidir a ordem**, não descobrir de novo. Lembrete dos itens, do mais barato ao mais caro:
1. Citar Resolução CGPAR 29/2022 e Acórdão 980/2023-TCU quando TIC = Sim (barato)
2. Justificativa de impacto orçamentário-financeiro (barato, texto fixo + 1 campo)
3. Documentos de qualificação técnica para habilitação em TR-Bens/TR-Serviços, no padrão do TR-Dispensa (médio)
4. Controles específicos de dispensa: limite de soma de valores do art. 416, aviso de 3 dias úteis (IN SEGES 67/21), cartão de pagamento + PNCP (médio)
5. Cabeçalho/preâmbulo da Minuta de Contrato (nº processo, nº contrato, qualificação das partes) — TR-Dispensa e PB (médio)
6. ~~Minuta de Contrato para TR-Bens e TR-Serviços~~ — **removido do escopo.** Confirmado no RLC (art. 447, §2º, e glossário de "Edital ou Instrumento Convocatório") que a elaboração do Contrato em licitações normais (Pregão) é competência da área de contratos/setor administrativo, e a minuta é anexada ao Edital — não ao TR, e não é atribuição da área demandante/técnica. TR-Dispensa e PB já têm Minuta porque, nas contratações diretas (art. 447, §1º), o processo é impulsionado pela própria área demandante, sem Edital.
7. Instrumento de Inexigibilidade — não existe hoje (muito grande, novo instrumento — categoria própria, distinta da Fase 5). **Confirmado que Avaliação de Imóveis (Fase 5) é hipótese de dispensa (art. 416, por valor), não de inexigibilidade** — inexigibilidade (art. 421, II, a) só cabe em avaliações de natureza singular com notória especialização e inviabilidade real de competição, caso raro. As duas fases seguem como projetos separados.

---

## Fase 7 — Automação de numeração (arquitetura)

**Não iniciar sem conversa própria.** Hoje cada número de seção/alínea (`"7.2."`, `"c)"`, etc.) é uma string fixa escrita no código Python — foi a causa de boa parte dos bugs encontrados e corrigidos ao longo desta sessão (parágrafos que sumem e deixam buraco na numeração, letras fora de ordem, etc.).

Duas abordagens possíveis, a debater antes de escrever qualquer plano de implementação:
- **(a) Numerador em runtime no JS**: o gerador de `.docx` já sabe, no momento da geração, quais blocos condicionais sobreviveram — poderia percorrer o documento final e renumerar sequências de `add_heading`/parágrafos marcados como "item de lista numerada", em vez de confiar em texto fixo. Mais robusto, mas exige marcar cada item numerado com um tipo/nível na função Python (`add_item_numerado(nivel, texto)` em vez de escrever o número na string).
- **(b) Numeração automática nativa do Word** (`w:numPr`/`w:numId`, listas numeradas OOXML de verdade) em vez de texto — o Word renumera sozinho e nunca fica errado, mas complica a lógica de bookmark/remoção de blocos condicionais que hoje depende de recortar texto bruto, e also complica a leitura em `js/preview-docx.js` (o motor de pré-visualização, que hoje só entende texto).

Recomendo (a) como primeiro passo, por ser incremental e não quebrar a lógica de remoção condicional já validada — mas isso é uma decisão de arquitetura que merece brainstorming dedicado antes de eu montar o plano de implementação.

---

## Resumo — ordem sugerida de execução

| Fase | Tamanho | Risco | Depende de decisão prévia? |
|---|---|---|---|
| 3. Tabela solta | Pequeno | Baixo | Não — pode começar já |
| 1. Vedações | Médio | Baixo | Não — pode começar já |
| 2. Modelo de Proposta | Pequeno/Médio | Baixo | Não — pode começar já |
| 6. Gaps do checklist (itens 1-2) | Pequeno | Baixo | Não |
| 4. DEMO | Grande | Médio (muito conteúdo novo) | Não, mas benefica de decidir se reaproveita `TipoCorrecao` como gate |
| 6. Gaps do checklist (itens 3-5) | Médio | Baixo | Não |
| 5. Avaliação de imóveis | Grande | Alto (instrumento novo) | Sim — brainstorming próprio |
| 6. Gaps do checklist (itens 6-7) | Grande | Alto (instrumentos/anexos novos) | Sim |
| 7. Numeração automática | Arquitetural | Alto | Sim — brainstorming próprio |

Minha sugestão: começar pelas fases 3, 1, 2 e os itens baratos da 6 nesta ordem (todas de baixo risco, sem decisões pendentes), depois ir para a Fase 4 (DEMO), e só então decidir entre Fase 5, o resto da Fase 6, e a Fase 7 — essas três exigem conversa antes de virar código.
