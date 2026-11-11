# Design — Redesenho visual (cards elevados)

## Contexto e motivação

O motor e o fluxo do gerador de TR-Bens já estão funcionando (ver
`docs/superpowers/plans/2026-08-16-motor-tr-bens.md`, mesclado ao `master`).
Visualmente, porém, o site está cru: sem separação clara entre seções, sem
respiro adequado, e a distinção parâmetro/especificação hoje é uma borda
fina de 3px que passa despercebida. Este design cobre só a camada visual —
nenhuma mudança de comportamento ou de arquitetura de dados.

Referência normativa: paleta e princípios de espaçamento do gov.br Design
System (`gov.br/ds`), o mesmo Design System já usado pelo site (Task 1).
Cores confirmadas na documentação oficial: azul interativo principal
`#1351b4` (Blue Warm Vivid 70), fundo escuro institucional `#071D41`,
amarelo de alerta `#FFCD07` (Yellow Vivid 20), verde de sucesso `#168821`.

## Processo de decisão

Três direções visuais foram comparadas lado a lado via mockup HTML
(prancheta visual de brainstorming), mantendo a mesma arquitetura de
informação em todas (sumário à esquerda, seções à direita, sem seção
oculta): A) cards elevados com sombra e respiro amplo; B) denso e
institucional, mais compacto; C) institucional colorido, com blocos de cor
de fundo. A usuária escolheu a opção **A — Cards elevados**, confirmado
tanto em texto quanto nos cliques registrados na prancheta.

Uma pergunta de escopo surgiu do próprio mockup: a opção A mostrava um
"✓" verde nas seções do sumário já completas — isso exige lógica nova
(calcular quando uma seção está completa), não é puramente visual. Decisão:
**fora do escopo desta rodada**, fica só o CSS/layout. O check de progresso
por seção pode ser uma rodada futura.

## Escopo

**Dentro do escopo** (só `css/estilo.css`, com ajustes pontuais de estrutura
HTML em `js/render.js`/`index.html` onde o CSS exigir um wrapper a mais):

- Cards de seção: fundo branco, `border-radius: 8px`, sombra sutil
  (`box-shadow: 0 2px 8px rgba(0,0,0,.08)`), respiro interno de 14–16px,
  espaçamento de 10px entre cards.
- Distinção parâmetro/especificação: borda esquerda de 3px por bloco de
  campo — azul (`#1351b4`) para parâmetro, amarelo (`#FFCD07`) para
  especificação — substituindo a borda fina atual, que era pouco visível.
- Sumário lateral: vira um cartão próprio (fundo branco, padding, sombra
  leve), sem cálculo de progresso por seção.
- Barra de progresso: hoje é um `<progress>` HTML puro. Achado já
  registrado no ledger da Task 11 (revisão final): o texto "X de Y campos
  preenchidos" definido via `elementoBarra.textContent` é o *fallback
  content* de um `<progress>`, que nenhum navegador comum renderiza —
  fica invisível na prática. Esta rodada troca por uma barra estilizada
  com `<div>` (trilho + preenchimento) e o texto ao lado, visível de
  verdade, resolvendo esse achado.

**Fora do escopo** (explicitamente adiado, não é omissão):

- Check de progresso por seção no sumário (exige lógica nova de cálculo de
  conclusão por seção, não só estilo).
- Qualquer mudança de comportamento, de arquitetura de dados, ou nova
  seção/campo.
- As variantes B e C, descartadas nesta decisão.

## Logo oficial da Conab

A usuária forneceu a logomarca oficial (`logo-horizontal-por-extenso-2.pdf`,
na raiz do projeto), substituindo o placeholder criado na Task 1
(`assets/logo-conab.svg`, um retângulo azul com o texto "CONAB"). Extraída
via PyMuPDF (renderização a 8x + corte automático pela caixa delimitadora
do canal alfa) para `assets/logo-conab.png` (2799×659px, fundo
transparente). `index.html` já foi atualizado para referenciar o PNG novo;
o placeholder SVG foi removido. Nenhuma mudança de CSS necessária além da
já existente (`header-logo { height: 40px }`), que funciona igual para PNG.

## Implementação

Concentrado em `css/estilo.css`:

```css
.card-secao {
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0,0,0,.08);
  padding: 16px;
  margin-bottom: 10px;
}
.campo-parametro { border-left: 3px solid #1351b4; padding-left: 8px; margin-bottom: 8px; }
.campo-especificacao { border-left: 3px solid #FFCD07; padding-left: 8px; margin-bottom: 8px; }
.sumario-card { background: #fff; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,.1); padding: 8px; margin-bottom: 10px; }
.barra-progresso-trilho { height: 8px; background: #E6E6E6; border-radius: 4px; overflow: hidden; margin-bottom: 4px; }
.barra-progresso-preenchimento { height: 100%; background: #168821; transition: width .2s; }
.barra-progresso-texto { font-size: 13px; color: #333333; margin-bottom: 10px; }
```

Em `js/render.js`, `renderizarSecoes` passa a envolver cada seção num
`<section class="card-secao">` (hoje já existe `bloco.className = "br-accordion"`
— adiciona `card-secao` junto, sem remover a classe do gov.br DS). Em
`atualizarProgresso`, troca a manipulação de atributos do `<progress>` por
manipulação de `width` do `.barra-progresso-preenchimento` e `textContent`
do `.barra-progresso-texto`. Em `index.html`/`js/app.js`, o markup do
`<progress id="barra-progresso">` é substituído pela estrutura de trilho +
texto equivalente, mantendo o mesmo `id` onde possível para minimizar o
diff.

`renderizarSumario`, em `js/render.js`, ganha o wrapper `sumario-card` ao
redor da lista existente.

## Testes

Mudança puramente visual/estrutural, sem lógica nova — não introduz
comportamento a testar automaticamente. Verificação: abrir o site num
navegador real (como já fizemos para a Task 11) e conferir visualmente que
os cards aparecem com sombra, a barra de progresso mostra o texto
corretamente (o que hoje está quebrado), e a distinção parâmetro/
especificação fica visualmente clara. Os testes automatizados existentes
(31 testes) não devem quebrar, já que nenhuma função pura muda de
assinatura ou comportamento — só o HTML/CSS que `render.js` produz.
