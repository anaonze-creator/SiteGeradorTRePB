# Design — Refinamento visual (estilo AGU + opções visíveis + tipografia)

## Contexto e motivação

O redesenho "cards elevados" (direção A) já está implementado nesta mesma
branch (`feature/redesenho-visual-cards`, ainda não mesclada): cada uma das
20 seções tem sombra própria. Ao revisitar o site da AGU
(`cgu.agu.gov.br/edital`) como referência, uma inspeção do CSS computado
real (não só visual) revelou a diferença estrutural: a AGU usa **uma única
elevação forte**, no cartão externo que envolve o formulário inteiro
(`box-shadow: 0 4px 16px rgba(0,0,0,.08)`); por dentro, cada grupo de campos
("Modalidade", "Tipo de Órgão" etc.) é um subcartão **plano**, sem sombra —
fundo `#f8f8f8`, borda `1px solid #ededed`. Com 20 seções, dar sombra a cada
uma faz as elevações competirem entre si e ficarem sutis demais para se
notar — foi exatamente o problema que motivou o achado do fundo branco
corrigido na rodada anterior.

Duas exigências adicionais da usuária, ao ver o mockup comparativo:
1. Campos com poucas opções (Sim/Não, Média/Mediana/Menor Preço etc.) devem
   ficar **visíveis de cara**, sem precisar abrir um menu — hoje são
   `<select>` nativos, que exigem um clique para revelar as opções.
2. Os títulos de seção ("01. DO OBJETO") estão com fonte grande demais —
   herdam o tamanho padrão do navegador para `<h2>`, sem nenhum ajuste.

Decisão validada via mockup na prancheta visual (opção única, sem A/B desta
vez, já que era um refinamento direto da direção já escolhida): aprovado
com "melhor, pode continuar dessa forma".

## Escopo

**Dentro do escopo:**

- **Cartão externo único** (`.cartao-painel`) envolvendo todo o conteúdo da
  Tela 2 (sumário + seções + barra de progresso + botão "Gerar Documento"),
  com sombra forte (`0 4px 16px rgba(0,0,0,.08)`) e um cabeçalho interno
  (`.cabecalho-painel`) com o título do instrumento.
- **Seções e sumário viram subcartões planos**: `.card-secao` e
  `.sumario-card` perdem a sombra própria e ganham fundo `#f8f8f8` + borda
  `1px solid #ededed`, já que agora estão aninhados dentro do cartão externo
  elevado — repetir sombra nos dois níveis voltaria ao problema original.
- **Campos `select` com até 4 opções viram grupos de botões clicáveis**
  ("pills"), visíveis imediatamente, sem exigir abrir um menu. O limite de 4
  é deliberado: todos os campos `select` do TR-Bens hoje têm 2 ou 3 opções,
  mas um campo futuro com uma lista longa (ex.: lista de UFs) deve continuar
  usando `<select>` nativo — um muro de botões não escala.
- **Tipografia menor nos títulos de seção**: de tamanho padrão de `<h2>`
  do navegador (tipicamente 24–32px) para `13px`, negrito, maiúsculas,
  com leve espaçamento entre letras — mais discreto, mais parecido com o
  peso visual que a AGU dá aos rótulos de grupo.
- **Rótulos de campo também reduzidos** para `13px` (hoje herdam o tamanho
  padrão do navegador para `<label>`,~16px).

**Fora do escopo:**

- Mudança de comportamento de geração do documento — os "pills" só mudam
  a forma de captura da resposta na tela; o valor guardado em `respostas`
  e usado por `js/gerador-docx.js` continua sendo a mesma string de opção
  (`"Sim"`, `"Não"`, `"Média"` etc.), sem nenhuma mudança no gerador.
- Header de página (logo + título global) — já existe e não é redundante
  o suficiente para justificar remoção; o cabeçalho interno do cartão passa
  a mostrar o título do *instrumento* específico (ex.: "Termo de Referência
  — Aquisição de Bens"), papel diferente do cabeçalho de página (que mostra
  o nome do *site*).

## Comportamento dos "pills" — esclarecimento importante

Diferente do `<select>` nativo, que sempre mostra visualmente a primeira
opção mesmo antes de qualquer interação (uma ilusão de resposta já dada, já
que `respostas[campo.id]` continua `undefined` até o usuário realmente
interagir), o grupo de pills **não realça nenhuma opção até o elaborador
clicar em uma**. Isso é mais honesto com o estado real dos dados — não é uma
regressão, é uma correção de um comportamento ambíguo que já existia.

## Implementação

`css/estilo.css` — reescrita das regras `.card-secao`/`.sumario-card`
(sombra → fundo+borda plana), acréscimo de `.cartao-painel`,
`.cabecalho-painel`, `.corpo-painel`, `.opcoes-pill`/`.pill`/`.pill.ativo`,
e redução de fonte em `.card-secao h2` e nos `label` dentro de
`.campo-parametro`/`.campo-especificacao`.

`js/app.js` — o template de `montarTelaPainel` passa a envolver
`#sumario`, `#secoes`, a barra de progresso e o botão dentro de
`<div class="cartao-painel"><div class="cabecalho-painel">...</div>
<div class="corpo-painel">...</div></div>`.

`js/render.js` — `criarCampoInput` ganha um novo ramo: quando
`campo.input === "select"` e `campo.opcoes.length <= 4`, renderiza um
`<div class="opcoes-pill">` com um `<button type="button" class="pill">`
por opção, em vez do `<select>`. Cada botão, ao ser clicado, marca a si
mesmo como `.ativo` (removendo a marcação dos irmãos) e chama
`aoMudar(campo.id, opcao)` — mesmo contrato já usado pelos demais campos.
Quando há mais de 4 opções, o comportamento atual (`<select>`) é mantido
sem alteração.

## Testes

Mudança de comportamento real desta vez (não só CSS): o clique nos pills
precisa disparar `aoMudar` corretamente. Como o projeto não usa jsdom
(convenção já estabelecida, ver plano anterior), a verificação é manual no
navegador — já validada informalmente durante o brainstorming, será
reconfirmada ao vivo depois da implementação, do mesmo jeito que a rodada
anterior. Os 31 testes automatizados existentes não devem quebrar, porque
nenhuma função pura (`js/modelo.js`) muda; a lógica nova mora inteiramente
em `criarCampoInput`, uma função de DOM já não coberta por teste
automatizado.
