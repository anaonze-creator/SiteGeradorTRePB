# TR-Dispensa — Avaliação de Imóveis: Design

## Contexto

O Gerador de TR e PB cobre hoje 6 instrumentos (TR-Bens, TR-Serviços, TR-Dispensa,
Projeto Básico, Nota de Demanda, Nota Técnica). A Fase 5 do plano
`docs/superpowers/plans/2026-08-20-vedacoes-proposta-demo-avaliacao-numeracao.md`
propôs um sétimo instrumento, dedicado à contratação de laudo de avaliação
imobiliária por dispensa de licitação — hoje inexistente no sistema.

O usuário forneceu um documento real da Conab como referência-fonte:
`documentos de referencia/SEI_55236534_Conab___RLC__Termo_de_Referencia__TR____Servicos.pdf`
(SEI 55236534, Processo 21447.001437/2026-39), um TR-Serviços completo,
contratado por dispensa de licitação, especificamente para "elaboração de
Laudo de Avaliação Imobiliária".

Confirmado neste ciclo (ver Fase 6, item 7, do plano acima): avaliação de
imóveis, no caso comum, é hipótese de **dispensa** (art. 416, I/II, do RLC,
por valor), não de inexigibilidade — a inexigibilidade (art. 421, II, "a")
só se aplicaria em caso raro de inviabilidade real de competição com
profissional de notória especialização. Este instrumento é, portanto,
uma variante do TR-Dispensa, não do futuro instrumento de Inexigibilidade
(estes seguem como projetos separados).

## Decisão de escopo

**Instrumento novo e separado**, não uma variante condicional dentro do
TR-Dispensa genérico. Arquivos próprios, card próprio na tela inicial,
seguindo o mesmo padrão usado quando o TR-Dispensa foi criado a partir do
TR-Serviços (arquivo de dados + template Python próprios, boilerplate
comum copiado e adaptado, não compartilhado via parâmetro condicional).

Razão: o conteúdo específico deste objeto (checklist de laudo, matriz de
riscos pré-preenchida, critério de qualificação técnica em m²) não faz
sentido como ramificação condicional dentro de um TR-Dispensa genérico
que atende a qualquer objeto de serviço — poluiria o instrumento genérico
com campos e textos que só existem para este nicho.

## Arquivos novos

- `templates/build_template_tr_dispensa_avaliacao_imoveis.py` — gera
  `Template_TR_Dispensa_Avaliacao_Imoveis.docx`.
- `js/dados-tr-dispensa-avaliacao-imoveis.js` — exporta
  `SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS`.
- Registro em `js/app.js`/`js/triagem.js` (card na tela inicial, ao lado
  dos outros 6instrumentos) — seguir o padrão usado para os demais.
- `tests/dados-tr-dispensa-avaliacao-imoveis.test.mjs` e testes de geração
  em `tests/gerador-docx.test.mjs` (ou arquivo próprio), seguindo o padrão
  dos demais instrumentos (contagem de seções, geração básica, blocos
  condicionais).

## Estrutura do documento

Numeração de seções e conteúdo devem seguir o documento de referência
(21 seções + 3 a 5 anexos, dependendo da decisão de Minuta de Contrato
abaixo). Cada seção abaixo indica se o conteúdo é **novo** (específico
deste objeto, extraído do documento de referência) ou **reaproveitado**
(copiado/adaptado do `build_template_tr_dispensa.py` genérico, já
existente e testado).

### Seção 1 — Do Objeto (NOVO)

Tabela de item único (Item, Localização, Especificações, Valor Total),
onde a coluna "Especificações" contém a checklist fixa do conteúdo
mínimo do laudo (itens a) a u"), transcrita integralmente do documento
de referência, citando ABNT NBR 14653-1, ABNT NBR 14653-2 e a Norma de
Alienação de Bens Imóveis da Conab (NOC 60.208), Capítulo V, item 3,
alíneas "a" a "k" artigo por artigo. Conteúdo **fixo** (texto oficial
normativo, não deve ser removível pelo elaborador) — apenas dois pontos
editáveis, herdados do próprio modelo oficial:
- item "l") grau mínimo de fundamentação exigido — campo de texto.
- valor estimado da contratação — campo numérico (padrão já usado nos
  outros instrumentos: `S01_ValorEstimado` + `S01_MetodologiaEstimativa`).

Demais itens da Seção 1 (regime de execução, critério de julgamento,
recursos orçamentários) seguem o padrão já usado no TR-Dispensa genérico.

### Seção 2 — Da Justificativa (REAPROVEITADO)

Idêntico ao TR-Dispensa genérico já existente, incluindo os itens já
adicionados nesta sessão: exceção de exclusividade ME/EPP (2.3),
impacto orçamentário-financeiro (2.4-2.5 renumerados conforme necessário)
e os 3 controles de dispensa do Anexo VIII-B (somatório de valores, aviso
de 3 dias úteis, cartão de pagamento/PNCP). A motivação (2.1) fica em
branco para o elaborador preencher (mas pode receber um texto-sugestão
folgado, já que o documento de referência tem uma motivação padrão sobre
"baixa disponibilidade de engenheiros e arquitetos da própria Conab" —
avaliar se vale a pena oferecer como sugestão via `ajuda`).

### Seção 3 — Do Objetivo (REAPROVEITADO)

Idêntico ao padrão genérico.

### Seção 4 — Da Vigência do Contrato (REAPROVEITADO)

Idêntico ao TR-Dispensa genérico (regime "até 1 ano, prorrogável" ou
"plurianual", campo de prazo de prorrogação já adicionado nesta sessão).
O documento de referência usa 180 dias fixos, mas isso é só o valor
daquele caso concreto — o campo deve continuar editável como no genérico.

### Seção 5 — Da Seleção do Fornecedor (MISTO)

- **5.1 Habilitação Jurídica/Fiscal/Econômico-Financeira**: reaproveitar
  o bloco padrão já construído nesta sessão para TR-Bens/TR-Serviços
  (Fase 6, item 3) — SICAF, documentos por natureza jurídica, registro
  especial de órgão regulador, regularidade fiscal/trabalhista,
  qualificação econômico-financeira (índices LG/SG/LC), aceite de
  somatório de atestados, consulta a cadastros de idoneidade, critério
  de inabilitação. Padroniza este instrumento com os demais.
- **5.2 Qualificação Técnica (NOVO, mais específico que o padrão)**:
  atestado de execução de laudos de avaliação de imóveis em área mínima
  de `____` m² (campo), limitada a no máximo 50% da área total do imóvel
  objeto — critério de compatibilidade domain-specific do documento de
  referência (não é o critério genérico de "quantitativo mínimo" usado
  nos outros TRs). Inclui também a comprovação, na assinatura do
  Contrato, de profissional habilitado com registro ativo em CREA, CAU,
  CFT, CRECI ou outro conselho com atribuição para avaliação de imóveis,
  no estado onde se localiza o imóvel.
- **5.3 Vistoria**: reaproveitar o padrão genérico (Obrigatória/Facultativa
  com texto "recomenda-se", já ajustado nesta sessão).

### Seção 6 — Da Descrição dos Serviços (NOVO)

- 6.1 Classificação dos serviços como serviço comum (art. 3º, XIII, RLC)
  e ausência de vínculo empregatício — texto fixo, adaptável do padrão
  genérico "6.1" já usado no TR-Serviços/TR-Dispensa (não é
  completamente novo, mas a redação do documento de referência é
  ligeiramente diferente e deve prevalecer aqui).
- 6.2 Detalhamento — remete à tabela da Seção 1 (texto fixo curto).
- 6.3 Local e Prazo de Execução — campos de local, horário, prazo de
  execução em dias, prazo para início dos serviços após assinatura
  (padrão já usado, reaproveitável).
- 6.4 Garantia dos Serviços — período mínimo de garantia (campo,
  documento de referência usa 1 ano) e prazo de correção/refazimento
  (campo) — mesmo padrão do TR-Serviços/TR-Dispensa genérico
  (`S06_GarantiaServicosMeses`-equivalente).
- 6.5 Recebimento dos Serviços — entrega impressa e digital (.pdf),
  recebimento provisório/definitivo com prazos — reaproveitar o padrão
  genérico já existente (`S06_PrazoRecebimentoProvisorioDias`/
  `DefinitivoDias`).

### Seção 7 — Da Garantia Contratual (NOVO — fixo)

Texto fixo: "Não será exigida garantia contratual." Sem campo
condicional Sim/Não (decisão do usuário: reflete a natureza do serviço,
sem produto físico a garantir).

### Seções 8 a 12 (REAPROVEITADO)

Acompanhamento e Fiscalização, Obrigações da Contratante, Obrigações da
Contratada, Pagamento, Reajuste — idênticos ao padrão genérico do
TR-Dispensa/TR-Serviços já existente e testado. Conferir se o índice de
reajuste deve ser o INCC (Índice Nacional de Custo da Construção, usado
no documento de referência por se tratar de serviço técnico de
engenharia) em vez do índice genérico — decisão de implementação, não
де design (usar o campo de índice já existente, com INCC como sugestão
de `ajuda` ou padrão pré-selecionado, se o campo for um `select`).

### Seção 13 — Da Subcontratação (NOVO — fixo)

Texto fixo: "Não será admitida a subcontratação do objeto", com a
justificativa específica do documento de referência (responsabilidade
técnica direta da Contratada, garantia de qualidade/confiabilidade). Sem
campo condicional Sim/Não (decisão do usuário).

### Seções 14 a 26 (REAPROVEITADO)

Alteração Subjetiva, Vedações (incluindo a seção nova desta sessão),
Sanções Administrativas, Rescisão Contratual, Vedação ao Nepotismo,
Sustentabilidade Ambiental, Matriz de Riscos e Responsabilidades (texto
da cláusula, não a tabela — ver Anexo I), Disposições Finais — idênticos
ao padrão genérico do TR-Dispensa já existente e testado.

### Anexo I — Matriz de Riscos (NOVO — pré-preenchida)

Ao contrário dos demais instrumentos (que usam a tabela oficial em
branco do RLC, função `add_matriz_de_riscos_rlc`), este instrumento usa
uma matriz **pré-preenchida e específica do domínio**, reaproveitando o
conhecimento já validado no documento de referência: 5 riscos, todos na
fase "Gestão do contrato" (não há riscos de planejamento/seleção de
fornecedor, porque a contratação é por dispensa):

1. Atraso na execução do serviço (baixa produtividade) — P2×I4=8
2. Inexecução do serviço (falta de planejamento do prestador) — P2×I5=10
3. Laudo entregue em desacordo com o conteúdo mínimo (falha técnica) — P2×I4=8
4. Escassez de dados de mercado comparáveis na região — P2×I3=6
5. Impossibilidade de vistoria interna do imóvel (ocupação irregular) — P2×I3=6

Mais a observação fixa: "Considerando que a presente contratação se
enquadra em hipótese de dispensa de licitação, em razão do valor
estimado, a matriz de riscos foi estruturada apenas para a fase de
gestão contratual, não contemplando etapas de planejamento ou de seleção
de fornecedor. [...]"

As escalas de Probabilidade/Impacto/Nível de Risco (tabelas auxiliares)
são as mesmas de todos os outros instrumentos — reaproveitar
`add_tabela` com o mesmo conteúdo já usado em `add_matriz_de_riscos_rlc`,
sem reaproveitar a própria função (que monta a tabela principal em
branco de 10 colunas — aqui a tabela principal é fixa/pré-preenchida,
estrutura diferente).

### Anexo II — Modelo de Apresentação de Proposta (NOVO — reaproveitando a Seção 1)

Reaproveita a mesma tabela-checklist da Seção 1 (item único, mesma
checklist a-u). Não é um "Modelo de Apresentação de Proposta" genérico
como os outros instrumentos (com "Item, Especificações, Valor Total")
— é uma cópia funcional do quadro da Seção 1, porque a proposta do
licitante precisa reproduzir a mesma especificação de conteúdo mínimo do
laudo.

### Anexo III — Declarações de Vistoria (REAPROVEITADO)

Modelos A (Declaração de Realização de Vistoria Técnica) e B
(Declaração de Pleno Conhecimento das Condições) — mesmo padrão já
usado no TR-Dispensa/TR-Serviços genérico.

### Anexos IV/V — Minuta de Contrato (REAPROVEITADO, decisão do usuário)

Incluir os mesmos Anexos V/VI do TR-Dispensa genérico (Contrato
Simplificado / Contrato completo, condicionados ao campo
`S07_TipoInstrumentoContratual`), incluindo o cabeçalho/preâmbulo
completo (nº processo, nº contrato, qualificação de CONTRATANTE e
CONTRATADA) já adicionado nesta sessão (Fase 6, item 5). O documento de
referência não mostrava essa minuta porque provavelmente foi extraído
antes dessa etapa do processo administrativo — não é evidência de que
o instrumento não precise dela.

## Convenção de nomes de bookmark

Seguir o padrão já estabelecido: prefixo `S` + número da seção (ex.:
`S01_ValorEstimado`, `S05_AreaMinimaAvaliada`, `S06_GarantiaServicosMeses`),
reaproveitando literalmente os nomes já usados no TR-Dispensa genérico
onde o campo for idêntico (para poder copiar blocos de código quase sem
alteração), e criando nomes novos só para os campos genuinamente novos
(ex.: `S05_AreaMinimaAvaliadaM2`, `S01_GrauMinimoFundamentacao`).

## Fora de escopo

- Não mexe no TR-Dispensa genérico existente (arquivos separados).
- Não decide neste momento a arquitetura de numeração automática (Fase 7
  do plano, brainstorm próprio).
- Não implementa o instrumento de Inexigibilidade (mesma categoria de
  "novo instrumento", mas projeto totalmente separado).
