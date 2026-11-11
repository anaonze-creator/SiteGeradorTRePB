# Design — Motor Projeto Básico (PB)

## Contexto e motivação

TR-Bens, TR-Serviços e TR-Serviços por Dispensa estão completos e mesclados
no `master`. O próximo instrumento é o Projeto Básico (PB), fonte normativa
`modelo_institucional_projeto_basico.txt` da skill `gepea-instrumentos`
(1.151 linhas), já validado pela GEPEA e citado na spec-mãe do projeto
(`2026-08-16-gerador-tr-pb-design.md`) como base para o PB. O texto foi lido
por inteiro — corpo completo das 34 seções e cabeçalho de todos os anexos —
antes de qualquer decisão de arquitetura.

O PB é estruturado segundo os incisos I a XXX do art. 157 do RLC, um
instrumento diferente em natureza do TR: cobre obras e serviços de
engenharia (regime de execução, ART/RRT no CREA/CAU, BIM, garantia mínima
de obra, orçamentação com BDI), não apenas bens/serviços comuns. O formato
de preenchimento também é diferente — `[CAMPO]` entre colchetes em vez de
`____` — mas isso é só cosmético para o motor: o bookmark envolve o mesmo
jeito o texto de placeholder, qualquer que seja seu formato visual.

## Achado que corrige a spec-mãe

A spec-mãe registrou que "apenas 3 seções em 35 variam de fato por tipo de
objeto" (Amostra/Demonstração Prévia, Critérios de Similaridade, Quadro de
Inovações). Essa contagem foi feita sem ler o documento completo. A leitura
integral mostra mais pontos de variação reais — vistoria (decide até qual
anexo de declaração entra), módulo BIM (Seções 4 e 18 e o Anexo II-A),
tipo de objeto (obra/serviço comum/serviço não comum de engenharia,
cascateando para modalidade licitatória e enquadramento), subcontratação —
mas a conclusão arquitetural da spec-mãe continua válida: **nenhuma seção
inteira é removida por variante de tipo de objeto; o que muda é o conteúdo
dentro da seção**, resolvido pelo mesmo mecanismo de variante/bloco já
usado nos três TRs. Só o Quadro de Inovações (Anexo II-A) é removido por
inteiro quando o objeto não é obra/serviço de engenharia — mesma técnica
`blocosAdicionais` + sentinela já usada para o Anexo I do TR-Dispensa.

## Achado que reduz o escopo

O próprio texto do modelo institucional (item 6 das instruções ao
elaborador) diz que as planilhas de Orçamento, Curva ABC, Cronograma
Físico-Financeiro e BDI **devem ser elaboradas no sistema oficial de
orçamentação (Orçafascio)**, exportadas em PDF e anexadas ao processo SEI
como documentos próprios — ou seja, **fora** do Projeto Básico gerado por
este site. Confirmado pelos próprios títulos dos anexos correspondentes no
modelo: "ANEXO X — MEMÓRIA DE CÁLCULO E PREMISSAS ORÇAMENTÁRIAS
(EXEMPLO)", "ANEXO XI — CURVA ABC (EXEMPLO)", "ANEXO XIV — ORÇAMENTO
ANALÍTICO (EXEMPLO)", "ANEXO XV — CRONOGRAMA FÍSICO-FINANCEIRO (EXEMPLO)" —
marcados como exemplo/referência, não como anexo a preencher e emitir.

**Decisão de escopo:** esses 4 anexos ficam fora do gerador — o corpo do
PB apenas os cita como anexados externamente (mesmo texto do modelo: "constam
do [Anexo X]..."), sem gerar conteúdo nenhum para eles. Isso é consistente
com a decisão já registrada na spec-mãe ("Fora do escopo da v0: Geração
automática de conteúdo técnico de anexos").

## Escopo

**Dentro do escopo:**
- As 34 seções do corpo do PB.
- Os anexos efetivamente gerados: I (Visão Geral/Caracterização), II
  (Matriz de Riscos) + II-A (Quadro de Inovações, condicional), III
  (Modelo de Apresentação de Proposta), IV (Declaração de Vistoria), V
  (Declaração de Dispensa de Vistoria — só um dos dois, IV ou V, entra no
  documento final, decidido pela mesma vistoria obrigatória/facultativa),
  VI (Declaração de Contratos Firmados), VII (IMR), VIII (Placa de
  Obra/Serviço de Engenharia).
- Nenhum mecanismo novo no motor — reaproveita bloco multi-parágrafo,
  variante nomeada, `blocosAdicionais`, `duplicarPara` e `tabela-linhas`
  tal como estão.

**Fora do escopo:**
- Anexos IX (Especificações Técnicas), X (Memória de Cálculo), XI (Curva
  ABC), XIV (Orçamento Analítico), XV (Cronograma Físico-Financeiro) — o
  próprio modelo institucional os trata como documentos externos
  (Orçafascio), citados no corpo do PB mas não gerados por este site.
- Pré-preenchimento cruzado entre corpo e anexos (equivalente à Fase 2 do
  TR-Dispensa) — fica para uma rodada futura, seguindo o mesmo padrão já
  validado (`duplicarPara`/`blocosAdicionais`) se a demanda aparecer.

## Mecanismos (todos reaproveitados, nenhum novo)

- **Tipo de objeto** (`PB01_TipoObjeto`, variantes `["Obra", "Serviço comum de engenharia", "Serviço não comum de engenharia"]`,
  `blocoTexto`): controla o enquadramento (art. 3º), a modalidade
  licitatória sugerida, e — via `blocosAdicionais` — a presença do Anexo
  II-A (Quadro de Inovações), que só existe para obra/serviço de
  engenharia.
- **Vistoria** (`PB09_TipoVistoria`, variantes `["Obrigatória", "Facultativa"]`,
  `blocoTexto` + `blocosAdicionais`): mesmo padrão do TR-Serviços/Dispensa
  para o texto "poderá/deverá" — mas aqui também decide, via
  `blocosAdicionais`, qual anexo sobrevive (Anexo IV — Declaração de
  Vistoria — sempre presente; Anexo V — Declaração de Dispensa — só
  relevante quando facultativa, mas o modelo trata como dois documentos
  separados por opção do licitante, não por escolha do elaborador; a
  decisão fina de como modelar isso fica para o plano).
- **Módulo BIM** (`PB04_AdotaBIM`, Sim/Não, `blocoTexto` +
  `blocosAdicionais` cobrindo Seção 4.X, Seção 18.2 e menções cruzadas):
  mesmo padrão do bloco de natureza intelectual do TR-Serviços/Dispensa.
- **Amostra/Demonstração Prévia** (Seção 10) e **Critérios de
  Similaridade** (Seção 11): Sim/Não simples, texto de não aplicação já
  redigido na própria minuta quando "Não" — mesmo padrão dos campos
  binários já usados nos TRs.
- **Subcontratação** (Seção 31): Sim/Não com percentual, idêntico ao
  padrão dos três TRs.
- **Garantia contratual** (Seção 22): percentual editável (como TR-Bens/
  TR-Dispensa, não fixo como TR-Serviços).
- **Resumo da Seção 3** (tabela com modalidade, regime, critério de
  julgamento, serviço continuado, subcontratação, participação ME/EPP):
  não é um mecanismo novo — é `duplicarPara` aplicado a cada um dos campos
  já decididos em outras seções, preenchendo a mesma tabela-resumo sem
  pedir para a elaboradora repetir a resposta.
- **Sanções** (Seção 25): mesma estrutura de percentuais em branco dos
  TRs, mas sem a tabela dupla de infrações graduadas — o PB não tem
  Tabela 1/Tabela 2, então o campo `tabela-linhas` não é necessário aqui
  (confirmado pela leitura: a Seção 25 lista só os 5 tipos de sanção com
  percentual direto, sem grau de infração).

## Testes

Mesmo critério das specs anteriores: gerar o documento cobrindo os dois
lados de cada condicional (tipo de objeto nas três variantes, vistoria
obrigatória/facultativa, BIM sim/não, amostra e similaridade sim/não,
subcontratação sim/não) e abrir no Word real, conferindo que o Anexo II-A
aparece só para obra/serviço de engenharia e que o anexo de vistoria certo
acompanha a escolha.

## Decisões em aberto (não bloqueiam o início do plano)

- Modelagem exata do par Anexo IV/Anexo V (Declaração de Vistoria vs.
  Declaração de Dispensa de Vistoria) — o modelo institucional sugere que
  a escolha entre os dois é do **licitante** no ato da vistoria, não da
  elaboradora no momento de gerar o PB; pode fazer mais sentido incluir os
  **dois** anexos sempre (limpos, como formulário) em vez de condicionar
  um deles à resposta de `PB09_TipoVistoria`. A decidir na escrita do
  plano.
- Regime de execução (Seção 1.3): 5 opções (empreitada por preço unitário/
  global/integral/semi-integrada/contratação integrada) — tratar como
  `select` de substituição direta (sem bloco de texto alternativo), já que
  a minuta não redige parágrafos diferentes por regime, só troca o nome no
  meio da frase.
- Extração fina campo-a-campo das 34 seções (perguntas, `ajuda`, tipos)
  fica para a escrita do plano de implementação, como já foi feito nos
  três TRs.
