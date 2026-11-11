# Design — Nota de Demanda e Nota Técnica

## Contexto e motivação

Dois novos instrumentos, revertendo a exclusão de escopo original da Nota
Técnica (registrada na spec-mãe). Fontes lidas por inteiro antes de
qualquer decisão:

- `SEI_55430846_Conab___RLC__Nota_de_Demanda.pdf` — modelo em branco.
- `SEI_55430881_Conab___RLC__Nota_Tecnica.pdf` — modelo em branco,
  praticamente idêntico ao da Nota de Demanda (mesma estrutura, mesmo
  texto de exemplo, sem diferenciação técnica).
- `SEI_54089094_Conab___RLC__Nota_Tecnica.pdf` — Nota Técnica **real**,
  preenchida e assinada (reforma da SUREG/RO, GEPEA), usada como
  referência de estrutura correta.
- `10.901_Regulamento_RLC_30.12.2024_versao4.pdf` — arts. 108, 110, 112
  a 124, 415, 416, cross-checados diretamente.

## Achado que corrige os dois modelos em branco

Os dois modelos em branco fornecidos têm o mesmo defeito: tratam Nota de
Demanda e Nota Técnica como idênticas, sem refletir a distinção que o
próprio RLC estabelece (art. 114: Nota de Demanda é funcional/área
demandante; arts. 116-117: Nota Técnica é técnica/área técnica). O
documento real (SEI 54089094) mostra a estrutura correta: quando o
objeto é obra ou serviço de engenharia, a Nota Técnica ganha 4 seções
extras (14 a 17), fundamentadas no art. 123, V a VIII do RLC (Subseção
II — Da Elaboração da Nota Técnica para Obras e Serviços de Engenharia),
ausentes dos dois modelos em branco. Confirmado por leitura direta do
art. 123 e do art. 416, I (citado corretamente no exemplo real para o
limite de dispensa de obras/serviços de engenharia, R$ 140.000,00,
diferente do art. 416, II usado pelo TR-Serviços/TR-Dispensa,
R$ 70.000,00).

Os dois modelos em branco também têm dois erros pontuais, que não devem
ser herdados: citação truncada `"(Art. )"` no item 10 (falta o número), e
a frase de encerramento do item 11 ("Para este objeto não existe a
necessidade...") sem o conector "OU" antes dela, diferente do item 10
que usa "OU" corretamente na mesma situação.

## Decisão estrutural

Nota de Demanda e Nota Técnica **compartilham o mesmo esqueleto de 13
seções**, derivado do art. 108, §1º, I a XII (Estudos Preliminares) mais
o art. 115, III (identificação da modalidade) — isso é literalmente o
que o art. 112 do RLC determina: "os Estudos Preliminares serão
inseridos na Nota de Demanda ou Nota Técnica". Não há base normativa
para uma estrutura de seções diferente entre os dois.

Mas **compartilhar o esqueleto não significa serem o mesmo documento** —
esse foi exatamente o defeito dos dois modelos em branco (achado
registrado acima). O RLC diferencia os dois de forma pontual e
verificável, e essa diferença precisa aparecer em três lugares
concretos do gerador, não só na estrutura de seções:

**1. A pergunta de cada uma das 13 seções-base é redigida diferente nos
dois instrumentos**, refletindo o próprio texto do art. 117: "a área
técnica deverá abordar todos os elementos de ordem técnica da
contratação e apresentar, **tecnicamente**, todos os itens mínimos
constantes na Nota de Demanda". A Nota de Demanda pergunta pelo
conteúdo administrativo/funcional tal como o art. 108/115 o descreve; a
Nota Técnica pergunta pelo mesmo item, mas exigindo tratamento técnico
explícito (especificações, normas técnicas aplicáveis, riscos técnicos,
memórias de cálculo, conforme o caso) — não é um reforço de estilo, é a
aplicação direta do verbo "tecnicamente" que o art. 117 usa. Exemplos
(lista completa das 13 na seção "Modelo de dados" abaixo):

  - Seção 1 — ND: "Qual é a situação-problema que originou a
    necessidade desta contratação, e quais os impactos para a Conab?"
    NT: "Descreva tecnicamente a situação-problema (condições físicas,
    funcionais ou normativas constatadas, de preferência com registro
    de vistoria) que origina a necessidade desta contratação, e seus
    impactos para a Conab."
  - Seção 7 — ND: "Descreva a solução como um todo: os elementos que
    devem ser produzidos, contratados ou executados." NT: "Descreva
    tecnicamente a solução como um todo, incluindo as normas técnicas
    aplicáveis (ABNT e correlatas) e os elementos que devem ser
    produzidos, contratados ou executados."

**2. O bloco de assinatura reflete a autoria distinta dos arts. 114 e
116-117**: a Nota de Demanda é assinada por empregado(s) da área
demandante ("Responsável(is) pelo estudo — Área Demandante"); a Nota
Técnica é assinada por empregado(s) da área técnica, com cargo/função
técnica ("Responsável(is) pelo estudo — Área Técnica"), como no exemplo
real (dois Analistas Engenheiros do GEPEA/SUPAD).

**3. A Nota Técnica, adicionalmente, ganha um bloco condicional de 4
seções extras (14-17)** — mesmo mecanismo `blocoTexto` já usado no PB
(`PB01_NaturezaObjeto`) — controlado por um campo "O objeto é obra ou
serviço de engenharia?": se Sim, inclui Adequação Técnica, Funcionalidade,
Requisitos Ambientais e Adequação às Normas Vigentes (art. 123, V-VIII);
se Não, essas seções são removidas do documento final. A Nota de Demanda
não tem esse bloco — o art. 123 (Subseção II) só rege a elaboração da
Nota Técnica. Nenhum mecanismo novo de motor é necessário —
`gerador-docx.js`/`modelo.js` já suportam isso integralmente.

## Modelo de dados

Cada uma das 13 seções-base é implementada como **um campo de texto
livre por seção** (tipo `especificacao`, `input: "textarea"`), com
`ajuda` citando a nota de referência oficial correspondente (já
transcrita do RLC nesta sessão, nunca inventada) — mesma lógica já usada
em `PB06_MetodoConstrutivo`. Isso reflete fielmente a natureza do
documento fonte: texto narrativo corrido com colchetes de preenchimento
livre, não um formulário de campos discretos como o TR. As tabelas
"meramente ilustrativas" (estimativa de quantidades, preços) citadas no
modelo ficam fora do escopo de geração automática — mesma decisão já
tomada para os anexos de cálculo do PB (Orçafascio) — o elaborador as
descreve em prosa no campo correspondente.

Três seções têm sub-estrutura Sim/Não com `mostrarSe` (parcelamento,
adequação do ambiente, contratações correlatas), replicando o padrão
"a) ... OU b) não há necessidade" do texto fonte — corrigindo a falta do
conector "OU" encontrada no modelo em branco.

Campos (prefixo `ND` para Nota de Demanda, `NT` para Nota Técnica —
mesmos nomes de seção, prefixo diferente). Citação de art. 115 só nas
seções em que ele realmente tem um item correspondente (1, 4, 5, 9, 13 —
ver tabela de conferência já apresentada na conversa); as demais citam
só o art. 108. **Correção em relação aos dois modelos em branco**: a
Seção 7 cita `art. 115, IV` (definição do objeto), não `V` (que já é o
item da Seção 4, quantidade/data) — o modelo de Nota de Demanda citava
só `V` (errado) e o de Nota Técnica citava `IV E V` (redundante).

| # | Seção | Base legal | Pergunta — Nota de Demanda | Pergunta — Nota Técnica |
|---|---|---|---|---|
| 1 | Necessidade de Contratação | art. 108, I; art. 115, I | Qual é a situação-problema que originou a necessidade desta contratação, e quais os impactos para a Conab? | Descreva tecnicamente a situação-problema (condições físicas, funcionais ou normativas constatadas, de preferência com registro de vistoria) que origina a necessidade desta contratação, e seus impactos para a Conab. |
| 2 | Referência a Outros Instrumentos de Planejamento | art. 108, II | A contratação está alinhada a algum instrumento de planejamento da Conab (Plano de Desenvolvimento Institucional, Planejamento Estratégico) ou política pública? | Do ponto de vista técnico, a contratação está alinhada a algum instrumento de planejamento da Conab ou política pública? Indique eventuais diretrizes técnicas internas aplicáveis. |
| 3 | Requisitos da Contratação | art. 108, III | Quais os requisitos necessários ao atendimento da necessidade (incluindo critérios de sustentabilidade e vigência pretendida do Contrato)? | Quais os requisitos técnicos necessários ao atendimento da necessidade (especificações técnicas, critérios de sustentabilidade aplicáveis e vigência tecnicamente adequada ao objeto)? |
| 4 | Estimativa das Quantidades | art. 108, IV; art. 115, V | Descreva o método usado para estimar as quantidades do objeto, com a data prevista de execução. | Descreva tecnicamente o método e a memória de cálculo usados para estimar as quantidades do objeto, com a data prevista de execução. |
| 5 | Levantamento de Mercado e Justificativa da Solução | art. 108, V; art. 115, II | Qual solução foi identificada como a que melhor atende à contratação, e por quê? | Tecnicamente, qual solução foi identificada como a que melhor atende à contratação, com base em quais referências técnicas, normas ou contratações similares? |
| 6 | Estimativas de Preços ou Preços Referenciais | art. 108, VI | Qual metodologia foi usada para estimar o preço de referência? | Qual metodologia técnica foi usada para estimar o preço de referência (fontes, memória de cálculo)? |
| 7 | Descrição da Solução como um Todo | art. 108, VII; art. 115, IV | Descreva a solução como um todo: os elementos que devem ser produzidos, contratados ou executados. | Descreva tecnicamente a solução como um todo, incluindo as normas técnicas aplicáveis (ABNT e correlatas) e os elementos que devem ser produzidos, contratados ou executados. |
| 8 | Parcelamento ou Não da Solução | art. 108, VIII | A solução será parcelada (licitada por item)? Justifique. | A solução será parcelada (licitada por item)? Justifique tecnicamente (viabilidade técnica do parcelamento, risco de perda de escala ou de compatibilização entre partes). |
| 9 | Demonstrativo dos Resultados Pretendidos | art. 108, IX; art. 115, VI | Quais os benefícios/resultados pretendidos com a contratação, e a quem se destinam? | Quais os resultados técnicos pretendidos com a contratação (ganhos de desempenho, conformidade normativa, durabilidade), e a quem se destinam? |
| 10 | Providências para Adequação do Ambiente da Conab | art. 108, X | Há necessidade de adequação do ambiente da Conab para a execução da contratação? | Há necessidade técnica de adequação do ambiente da Conab (acesso, infraestrutura, capacitação) para a execução da contratação? |
| 11 | Contratações Correlatas ou Interdependentes | art. 108, XI | Há contratos correlatos ou interdependentes a esta contratação? | Há contratos ou projetos técnicos correlatos ou interdependentes a esta contratação? |
| 12 | Declaração da Viabilidade | art. 108, XII | A contratação é viável? Justifique. | A contratação é tecnicamente viável? Justifique com base nos elementos técnicos apurados. |
| 13 | Identificação da Modalidade de Licitação | art. 115, III | Qual modalidade de licitação será utilizada, e por quê? | Qual modalidade de licitação será utilizada, e por quê? |

Campos técnicos (`select`/`numero`) por seção, além do textarea acima:
`08_Parcelamento` (select Sim/Não, `blocoTexto`) + `08_Justificativa`
(`mostrarSe: Sim`); `10_AdequacaoAmbiente` (select Sim/Não,
`blocoTexto`) + `10_DetalheAdequacao` (`mostrarSe: Sim`);
`11_ContratacoesCorrelatas` (select Sim/Não, `blocoTexto`) +
`11_DetalheCorrelatas` (`mostrarSe: Sim`); `12_Viabilidade` (select
Viável/Inviável, `blocoTexto`) + `12_Justificativa`; `13_Modalidade`
(select Pregão Eletrônico/Competição Pública/RDC/Dispensa de Licitação)
+ `13_Justificativa`. Réplica exata do padrão "a) ... OU b) não há
necessidade" do texto fonte para os três blocos Sim/Não — corrigindo a
falta do conector "OU" encontrada no modelo em branco (item 11).

Somente na Nota Técnica:

- `NT_ObraServicoEngenharia` — select Sim/Não (`blocoTexto`,
  `blocosAdicionais` removendo as 4 seções extras quando "Não")
- `NT14_AdequacaoTecnica`, `NT15_Funcionalidade`,
  `NT16_RequisitosAmbientais`, `NT17_AdequacaoNormasVigentes` — textarea,
  `mostrarSe: { NT_ObraServicoEngenharia: "Sim" }`

Cabeçalho comum a ambos: identificação do objeto, número do processo SEI,
data — mesmos campos de identificação já usados nos outros instrumentos
(reaproveitar convenção, não reinventar).

## Fora do escopo

- Geração automática das tabelas ilustrativas de estimativa de
  quantidades/preços (Local/Item/Total, Produto/Quantidade) — ficam como
  texto livre no campo correspondente, mesma decisão já tomada para os
  anexos de cálculo de engenharia do PB.
- Um terceiro "modo combinado" do art. 118 (área demandante e área
  técnica assinam só a Nota Técnica, dispensando a Nota de Demanda) —
  fica para decisão futura; por ora os dois instrumentos são gerados
  separadamente, como hoje TR e PB são.

## Testes

Mesma convenção dos outros instrumentos: `lerTemplateNotaDemanda`/
`gerarXmlNotaDemanda` e equivalentes para Nota Técnica, cobrindo: geração
básica, remoção correta dos 3 blocos condicionais Sim/Não com `mostrarSe`,
e o bloco condicional de 4 seções extras da Nota Técnica aparecendo só
quando `NT_ObraServicoEngenharia = "Sim"`.
