# Design — Motor TR-Serviços

## Contexto e motivação

O TR-Bens está completo e mesclado no `master`: motor de dados declarativo
(`js/modelo.js` + `js/dados-tr-bens.js`), painel de preenchimento com
sumário clicável, e geração de `.docx` por manipulação de bookmarks
(`js/gerador-docx.js`) sobre `templates/Template_TR_Bens.docx`.

O próximo instrumento é o TR-Serviços, minuta oficial fornecida em
`documentos de referencia/SEI_55379668_Conab___RLC__Termo_de_Referencia__TR____Servicos.pdf`.
O texto foi extraído com `pdftotext -layout -enc UTF-8` e lido seção a
seção — nada neste documento foi inventado; toda referência a texto da
minuta cita o trecho real.

O TR-Serviços compartilha a mesma espinha de 20 seções do TR-Bens (mesmos
títulos, mesma ordem, só a Seção 6 muda de nome: "DA DESCRIÇÃO DOS
SERVIÇOS"), mas é significativamente mais rico: qualificação técnica e
vistoria na Seção 5, bifurcação real de redação na Seção 11, bloco de
natureza intelectual na Seção 10, tabela dupla de sanções na Seção 15, e 5
anexos (contra 2 do TR-Bens).

O motor de condicionais do TR-Bens (`mostrarSe` + bookmarks `_BLOCO_SIM` /
`_BLOCO_NAO`, um parágrafo por lado) não cobre integralmente essas
diferenças. Este documento desenha as extensões mínimas necessárias.

## Escopo

**Dentro do escopo:**
- Duas primitivas novas no motor de condicionais (bloco multi-parágrafo,
  variante nomeada — ver "Mecanismos"), generalizando o que já existe sem
  reescrevê-lo.
- Um tipo de campo novo (`tabela-linhas`) para a lista dinâmica de infrações
  extras da Seção 15.
- `js/dados-tr-servicos.js`, espelhando a estrutura de
  `js/dados-tr-bens.js`, com as 20 seções e os campos derivados da leitura
  integral da minuta SEI 55379668.
- `templates/build_template_tr_servicos.py` e o `.docx` gerado por ele,
  análogos ao par existente para Bens, cobrindo as 20 seções e os 5 anexos.
- Generalização de `js/app.js` e `js/triagem.js` para operar sobre mais de
  um instrumento/variante (hoje hardcoded em TR-Bens).

**Fora do escopo (herda as decisões já registradas na spec-mãe
`2026-08-16-gerador-tr-pb-design.md`):**
- Geração automática de conteúdo técnico de anexo (ex.: preenchimento de
  indicadores do IMR) — os 5 anexos entram limpos, na forma oficial.
- TR-Serviços por Dispensa e PB — ficam para specs futuras próprias.
- Heurística completa de triagem entre as 3 variantes de TR + PB — a Tela 1
  ganha um segundo caminho ("Avançar para o TR-Serviços"), mas a
  classificação automática completa (bem vs. serviço comum vs. serviço com
  mão de obra exclusiva vs. obra/engenharia) é decisão futura; por ora a
  elaboradora escolhe diretamente a variante.

## Mecanismos novos no motor de condicionais

### 1. Bloco condicional multi-parágrafo

**Problema:** `localizarParagrafoDoBookmark` (`js/gerador-docx.js`) hoje
resolve cada lado de um condicional como exatamente um `<w:p>...</w:p>` —
suficiente para a Seção 13 do TR-Bens (um parágrafo por lado), insuficiente
para:
- Seção 11.2-b do TR-Serviços: quando a fiscalização é só Fiscal Funcional,
  vale a alínea "b" única; quando é Comissão de Fiscalização, valem cinco
  alíneas inteiras (b, b.1, b.2, b.3, c) — texto real: *"Quando a
  fiscalização for composta apenas por Fiscal Funcional, deverá ser
  utilizada a primeira redação (somente alínea 'b'). Quando se tratar de
  Comissão de Fiscalização, deverá ser utilizada a segunda redação (alíneas
  'b', 'b.1', 'b.2', 'b.3' e 'c'...)"*.
- Seção 10, bloco de natureza intelectual: alíneas u, u.1, v, v.1, v.2, w.

**Decisão:** o bookmark do Word já suporta abranger múltiplos parágrafos
nativamente (`bookmarkStart`/`bookmarkEnd` não precisam estar no mesmo
`<w:p>`). A função de busca passa a localizar o `<w:p>` que contém o
`bookmarkStart` e, separadamente, o `<w:p>` que contém o `bookmarkEnd` de
mesmo `w:id`, e o range removido/mantido é do início do primeiro ao fim do
segundo, inclusive. Um bloco de um parágrafo só continua funcionando igual
(os dois extremos coincidem) — mudança retrocompatível, sem tocar em
`dados-tr-bens.js` nem no restante do motor Sim/Não.

### 2. Variante nomeada (N-vias, não só Sim/Não)

**Problema:** a convenção atual (`nomeBlocoSim`/`nomeBlocoNao`,
`controladoresCondicionais` assumindo resposta `"Sim"`) é binária por
construção. A Seção 11 bifurca por `S08_TipoFiscalizacao` (valores "Fiscal
Funcional" / "Comissão de Fiscalização") e a Seção 5 por `S05_TipoVistoria`
("Obrigatória" / "Facultativa") — nenhum dos dois é literalmente Sim/Não.

**Decisão:** generalizar para `nomeBlocoVariante(campoId, valorOpcao)`
(bookmark passa a se chamar `${campoId}_BLOCO_${valorOpcao normalizado}`).
Um campo de parâmetro declara `variantes: [...]` em vez de depender
implicitamente de `"Sim"`/`"Não"`; o gerador resolve a resposta atual e
remove todos os blocos cujo `valorOpcao` não corresponde. O par Sim/Não do
TR-Bens passa a ser só o caso de 2 variantes chamadas "Sim" e "Não" — mesmo
código, sem migração de dado necessária (os nomes de bookmark existentes no
`Template_TR_Bens.docx`, `_BLOCO_SIM`/`_BLOCO_NAO`, continuam batendo porque
"Sim"/"Não" normalizam para os mesmos sufixos).

### 3. Campo `tipo: "tabela-linhas"` (Seção 15, infrações extras)

A Tabela 1 (15.8.1-j) tem 9 infrações com grau (1 a 5) fixo, texto oficial
da minuta — viram parte estática do `.docx`, sem campo associado. A nota
*"Poderão ser incluídas na tabela de infrações outras condutas que se
entenderem necessárias..."* é o único ponto dinâmico: uma lista de linhas
extras que a elaboradora adiciona quando o objeto exige.

```js
{ id: "S15_InfracoesExtras", tipo: "tabela-linhas",
  colunas: [
    { id: "descricao", rotulo: "Descrição da conduta", input: "texto" },
    { id: "grau", rotulo: "Grau", input: "select", opcoes: ["1", "2", "3", "4", "5"] },
  ],
  pergunta: "Infrações adicionais à Tabela 1 (além das 9 já previstas na minuta)" }
```

No gerador, cada linha de `respostas.S15_InfracoesExtras` (array de
`{descricao, grau}`) vira uma nova `<w:tr>` clonada de uma linha-modelo
marcada por bookmark próprio na Tabela 1 do template, inserida antes do
fechamento da tabela. É o único ponto do motor com geração de linha
dinâmica; os outros três mecanismos reaproveitam bookmark existente.

A Tabela 2 (grau → percentual) é 100% fixa (0,2% / 0,4% / 0,8% / 1,6% /
3,2%, por força do art. 578 do RLC) — não vira campo. O único parâmetro
associado é `S15_BaseCalculoMulta` (variantes: "valor do Contrato" / "valor
anual do Contrato" / "valor mensal do Contrato"), que resolve as reticências
repetidas no texto das alíneas "e", "f" e na própria Tabela 2, usando o
mecanismo do item 2.

## Campos por seção — pontos de maior densidade

A extração completa da minuta orienta `dados-tr-servicos.js`; os pontos
abaixo concentram a diferença real em relação ao TR-Bens (as demais 13
seções seguem o mesmo padrão simples já usado em Bens — parâmetro ou
especificação direta, sem mecanismo novo):

- **Seção 5** — Qualificação Técnica (prazo mínimo de experiência,
  quantitativo mínimo do serviço principal) e Vistoria
  (`S05_TipoVistoria`, mecanismo 2; texto da 5.2.1 e a existência da
  alínea a.1 dependem dela).
- **Seção 8** — `S08_TipoFiscalizacao` (mecanismo 2), controlador
  compartilhado com o bloco da Seção 11.
- **Seção 10** — `S10_NaturezaIntelectual` (Sim/Não simples, mecanismo 1
  para o bloco u–w).
- **Seção 11** — bifurcação de redação inteira controlada por
  `S08_TipoFiscalizacao` (mecanismo 1 + 2 combinados: bloco multi-parágrafo
  por variante).
- **Seção 15** — `S15_BaseCalculoMulta` (mecanismo 2) e
  `S15_InfracoesExtras` (mecanismo 3), além dos campos de percentual
  simples já no padrão do TR-Bens (multa por não regularização, pela
  recusa em assinar etc., que aqui têm exatamente o mesmo formato de
  15.8.1-a/b/c).

## Anexos

Os 5 anexos (Matriz de Riscos, Modelo de Apresentação de Proposta,
Declaração de Vistoria, Declaração de Contratos Firmados, Modelo de IMR)
seguem o mesmo padrão já usado nos 2 anexos do TR-Bens: **conteúdo estático
e limpo**, embutido no `.docx` pelo script de build Python
(`build_template_tr_servicos.py`), sem passar pelo motor de dados — exceto
os mesmos bookmarks triviais já usados em Bens para campos já coletados em
outra parte do formulário (objeto, endereço, unidade demandante), repetidos
no cabeçalho de cada anexo aplicável. Nenhum anexo tem campo de
especificação próprio; a Declaração de Contratos Firmados e o IMR mantêm as
fórmulas e o exemplo preenchido da minuta como texto fixo, não como dado.

## Plumbing: generalizar `app.js` e `triagem.js`

Hoje `js/app.js` fixa `INSTRUMENTO = "TR-Bens"`, importa só
`SECOES_TR_BENS`, busca só `templates/Template_TR_Bens.docx` e usa uma
única chave de `localStorage`. Para o TR-Serviços entrar como segunda
opção real (não uma reescrita paralela do app), `app.js` passa a receber o
instrumento escolhido na triagem como parâmetro de inicialização e resolve
a partir dele: qual `SECOES_*` renderizar, qual `.docx` buscar, qual chave
de autosave/nome de arquivo de download usar. `js/triagem.js` ganha um
segundo botão ("Avançar para o TR-Serviços") ao lado do já existente,
adiando a heurística automática completa de classificação (fora de escopo,
ver acima).

## Testes

Mesmo critério já usado no TR-Bens (projeto estático, risco concentrado na
geração do `.docx`, não em lógica de negócio): gerar o documento com um
caso preenchido por completo — incluindo os dois lados de cada condicional
novo (vistoria obrigatória e facultativa; fiscal único e comissão; com e
sem natureza intelectual; com e sem infrações extras na Tabela 1) — e abrir
no Word real, conferindo que:
- os blocos multi-parágrafo removem o lado certo por inteiro, sem deixar
  parágrafo órfão nem duplicar numeração de alínea;
- as linhas extras da Tabela 1 aparecem formatadas igual às 9 fixas;
- o TR-Bens continua gerando exatamente como antes (regressão do
  mecanismo 1/2 generalizado).

## Decisões em aberto (não bloqueiam o início da implementação)

- Composição variável da Comissão de Fiscalização (a minuta permite
  "conforme composição da referida comissão", sugerindo que nem toda
  comissão tem os três fiscais Funcional+Técnico+Setorial) — v0 trata como
  binário (Fiscal Funcional único vs. Comissão completa de 3); variação
  parcial da comissão fica para uma iteração futura, se a demanda real
  aparecer.
- Texto exato do local/endereço da vistoria (linha 5.2.1: "das ..... horas
  às ...... horas", "telefone (....)") — provavelmente reaproveita campos
  já coletados de local de execução (Seção 6) mais um campo de horário e
  telefone próprios; a decidir na extração fina de `dados-tr-servicos.js`.
