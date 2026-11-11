# Redesenho Visual (Cards Elevados) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Apply the "cards elevados" visual direction (chosen via visual
brainstorming) to the already-working TR-Bens generator: elevated white
cards for each section, a card-styled sumário, and a real visible progress
bar (replacing the currently-invisible native `<progress>` fallback text).

**Architecture:** Pure CSS/markup change on top of the existing engine —
no new pure functions, no new data, no behavior change. `js/render.js`
gains classes/structure on elements it already creates; `js/modelo.js` and
every field/section behavior stay untouched.

**Tech Stack:** Vanilla CSS, vanilla DOM (same as the rest of the project).
No build step, no new dependency.

## Global Constraints

- No servidor, sem build, sem npm no artefato final (unchanged project rule).
- Nenhuma mudança de comportamento — só aparência. Os 31 testes existentes
  (`node --test tests/*.test.mjs`) devem continuar passando sem alteração,
  porque nenhuma assinatura de função pura muda.
- Cores usadas devem vir da paleta oficial gov.br DS já confirmada: azul
  interativo `#1351b4`, amarelo de alerta `#FFCD07`, verde de sucesso
  `#168821`, cinza de trilho `#E6E6E6`, texto `#333333`.
- Manter o `id="barra-progresso"` no elemento externo da barra de progresso
  (não renomear) para minimizar o diff em `js/app.js`, que já chama
  `atualizarProgresso(SECOES_TR_BENS, respostas, document.getElementById("barra-progresso"))`
  duas vezes e não deve precisar mudar.

---

### Task 1: Cards elevados para seções e sumário

**Files:**
- Modify: `css/estilo.css`
- Modify: `js/render.js` (`renderizarSumario`, `renderizarSecoes`)

**Interfaces:**
- Consumes: nada novo — `renderizarSumario(secoes, container)` e
  `renderizarSecoes(secoes, respostas, container, aoMudar)` mantêm as
  mesmas assinaturas já usadas por `js/app.js`.
- Produces: nenhuma interface nova — só classes CSS adicionais nos
  elementos que essas funções já criam.

- [ ] **Step 1: Adicionar as classes de card ao CSS**

Abrir `css/estilo.css` (conteúdo atual, para referência exata):

```css
.header-logo { height: 40px; margin-right: 1rem; }
.header-titulo { font-weight: 600; font-size: 1.1rem; }
.campo-parametro { border-left: 3px solid #1351b4; padding-left: .5rem; }
.campo-especificacao { border-left: 3px solid #ffcd07; padding-left: .5rem; }
[hidden] { display: none !important; }
```

Substituir pelo conteúdo completo abaixo (mantém as 5 regras originais,
ajusta duas delas e acrescenta as novas classes de card):

```css
.header-logo { height: 40px; margin-right: 1rem; }
.header-titulo { font-weight: 600; font-size: 1.1rem; }
.campo-parametro { border-left: 3px solid #1351b4; padding-left: .5rem; margin-bottom: .75rem; }
.campo-especificacao { border-left: 3px solid #ffcd07; padding-left: .5rem; margin-bottom: .75rem; }
[hidden] { display: none !important; }

.card-secao {
  background: #ffffff;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, .08);
  padding: 16px;
  margin-bottom: 10px;
}
.card-secao h2 {
  margin-top: 0;
  color: #071D41;
}

.sumario-card {
  background: #ffffff;
  border-radius: 8px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, .1);
  padding: 8px;
  margin-bottom: 10px;
}
```

- [ ] **Step 2: Aplicar a classe `card-secao` em `renderizarSecoes`**

Em `js/render.js`, localizar a linha (dentro de `renderizarSecoes`):

```js
    bloco.className = "br-accordion";
```

Substituir por (mantém a classe do gov.br DS já usada e acrescenta a nova,
sem remover nada):

```js
    bloco.className = "br-accordion card-secao";
```

- [ ] **Step 3: Envolver o sumário num card**

Em `js/render.js`, a função `renderizarSumario` atualmente é:

```js
export function renderizarSumario(secoes, container) {
  container.innerHTML = "";
  const lista = document.createElement("ul");
  lista.className = "br-list";
  for (const secao of secoes) {
    const item = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#secao-${secao.id}`;
    link.textContent = `${secao.id}. ${secao.titulo}`;
    item.appendChild(link);
    lista.appendChild(item);
  }
  container.appendChild(lista);
}
```

Substituir por (a única mudança é envolver a `<ul>` existente num `<div class="sumario-card">`, sem alterar a lógica de montagem da lista):

```js
export function renderizarSumario(secoes, container) {
  container.innerHTML = "";
  const cartao = document.createElement("div");
  cartao.className = "sumario-card";
  const lista = document.createElement("ul");
  lista.className = "br-list";
  for (const secao of secoes) {
    const item = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#secao-${secao.id}`;
    link.textContent = `${secao.id}. ${secao.titulo}`;
    item.appendChild(link);
    lista.appendChild(item);
  }
  cartao.appendChild(lista);
  container.appendChild(cartao);
}
```

- [ ] **Step 4: Verificar sintaxe**

Run: `node --check js/render.js`
Expected: sem saída (sucesso silencioso).

- [ ] **Step 5: Rodar a suíte de testes completa (nenhuma quebra esperada)**

Run: `node --test tests/*.test.mjs`
Expected: `tests 31`, `pass 31`, `fail 0` — igual a antes, já que nenhuma
função pura mudou de assinatura ou comportamento.

- [ ] **Step 6: Verificação manual no navegador**

Servir o projeto por HTTP (ver `README.md` — `python -m http.server` a
partir da raiz do projeto), abrir `http://localhost:8000`, avançar da
triagem para o painel. Confirmar visualmente:
- cada uma das 20 seções aparece como um cartão branco com sombra sutil,
  separado das demais por um respiro visível;
- o sumário lateral também aparece como um cartão branco com sombra leve,
  não mais como uma lista solta sobre o fundo da página;
- os campos de parâmetro (borda azul) e especificação (borda amarela)
  continuam distinguíveis, agora com mais espaço entre um campo e outro.

- [ ] **Step 7: Commit**

```bash
git add css/estilo.css js/render.js
git commit -m "Add elevated card styling to sections and sumário"
```

---

### Task 2: Barra de progresso visível

**Files:**
- Modify: `index.html`
- Modify: `js/render.js` (`atualizarProgresso`)

**Interfaces:**
- Consumes: `progresso(secoes, respostas)` de `js/modelo.js` — já importado
  em `js/render.js`, sem mudança de assinatura.
- Produces: `atualizarProgresso(secoes, respostas, elementoBarra)` mantém
  exatamente a mesma assinatura — `js/app.js` não precisa de nenhuma
  alteração, porque `elementoBarra` continua sendo o elemento com
  `id="barra-progresso"`, só que agora é uma `<div>` (container) em vez de
  um `<progress>`.

Este achado motivou a tarefa: o `<progress>` HTML nativo só exibe seu
conteúdo textual como *fallback* quando o navegador não suporta o
elemento — nenhum navegador comum realmente renderiza esse texto. Por
isso a mensagem "X de Y campos de especificação preenchidos", escrita via
`elementoBarra.textContent = ...`, está invisível hoje. A troca por uma
estrutura de `<div>`s resolve isso de vez.

- [ ] **Step 1: Trocar o markup do `<progress>` em `index.html`**

Localizar, em `index.html`, dentro do template de `montarTelaPainel` em
`js/app.js` (não em `index.html` — a barra é criada dinamicamente pelo
JavaScript, não está no HTML estático). Abrir `js/app.js` e localizar a
linha:

```js
    <progress id="barra-progresso" value="0" max="100"></progress>
```

Substituir por:

```js
    <div id="barra-progresso" class="barra-progresso" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
      <div class="barra-progresso-trilho"><div class="barra-progresso-preenchimento"></div></div>
      <p class="barra-progresso-texto"></p>
    </div>
```

- [ ] **Step 2: Adicionar as classes ao CSS**

Em `css/estilo.css` (já editado na Task 1), acrescentar ao final:

```css
.barra-progresso-trilho {
  height: 8px;
  background: #E6E6E6;
  border-radius: 4px;
  overflow: hidden;
  margin-bottom: 4px;
}
.barra-progresso-preenchimento {
  height: 100%;
  background: #168821;
  transition: width .2s;
  width: 0%;
}
.barra-progresso-texto {
  font-size: .8125rem;
  color: #333333;
  margin: 0 0 10px;
}
```

- [ ] **Step 3: Reescrever `atualizarProgresso` para manipular a nova estrutura**

Em `js/render.js`, a função atual é:

```js
export function atualizarProgresso(secoes, respostas, elementoBarra) {
  const { totalEspecificacao, preenchidos } = progresso(secoes, respostas);
  const percentual = totalEspecificacao === 0 ? 100 : Math.round((preenchidos / totalEspecificacao) * 100);
  elementoBarra.setAttribute("value", String(percentual));
  elementoBarra.setAttribute("aria-valuenow", String(percentual));
  elementoBarra.textContent = `${preenchidos} de ${totalEspecificacao} campos de especificação preenchidos`;
}
```

Substituir por:

```js
export function atualizarProgresso(secoes, respostas, elementoBarra) {
  const { totalEspecificacao, preenchidos } = progresso(secoes, respostas);
  const percentual = totalEspecificacao === 0 ? 100 : Math.round((preenchidos / totalEspecificacao) * 100);
  const preenchimento = elementoBarra.querySelector(".barra-progresso-preenchimento");
  const texto = elementoBarra.querySelector(".barra-progresso-texto");
  preenchimento.style.width = `${percentual}%`;
  texto.textContent = `${preenchidos} de ${totalEspecificacao} campos de especificação preenchidos`;
  elementoBarra.setAttribute("aria-valuenow", String(percentual));
}
```

- [ ] **Step 4: Verificar sintaxe**

Run: `node --check js/render.js && node --check js/app.js`
Expected: sem saída (sucesso silencioso).

- [ ] **Step 5: Rodar a suíte de testes completa (nenhuma quebra esperada)**

Run: `node --test tests/*.test.mjs`
Expected: `tests 31`, `pass 31`, `fail 0`.

- [ ] **Step 6: Verificação manual no navegador**

Com o site servido por HTTP, avançar para o painel e digitar em um campo
de especificação (ex.: seção 01, "Descreva o material a ser adquirido").
Confirmar visualmente:
- a barra de progresso (trilho cinza + preenchimento verde) aparece logo
  acima do botão "Gerar Documento" e o preenchimento cresce ao digitar;
- o texto "X de Y campos de especificação preenchidos" agora **aparece de
  verdade** na tela, ao lado/abaixo da barra — isso é a correção do achado
  que estava invisível antes;
- inspecionar o elemento no DevTools e confirmar que `aria-valuenow` muda
  junto com o preenchimento.

- [ ] **Step 7: Commit**

```bash
git add index.html js/app.js js/render.js css/estilo.css
git commit -m "Replace invisible native progress fallback with a visible styled bar"
```

---

## Self-Review

**Cobertura da spec:** Task 1 cobre os cards de seção e o sumário-card do
design (`docs/superpowers/specs/2026-08-16-redesenho-visual-cards-design.md`);
Task 2 cobre a barra de progresso visível, incluindo a correção do achado
já ledgerado sobre o `<progress>` invisível. A distinção parâmetro/
especificação por borda colorida já existia no código antes deste plano
(`campo-parametro`/`campo-especificacao` em `css/estilo.css`) — Task 1 só
adiciona o espaçamento entre campos (`margin-bottom`), sem mudar as cores,
que já batem com a paleta oficial. A logo da Conab já foi trocada
separadamente pela usuária/controller antes deste plano (commit
`23b075a`), fora do escopo destas duas tarefas.

**Placeholders:** nenhum "TBD"/"implementar depois" — todo trecho de
código é o conteúdo final exato a escrever.

**Consistência de tipos:** `atualizarProgresso(secoes, respostas, elementoBarra)`
mantém a mesma assinatura em `js/render.js` e nas duas chamadas já
existentes em `js/app.js` — nenhuma delas precisa de edição. `renderizarSumario`
e `renderizarSecoes` também mantêm assinatura idêntica à já consumida por
`js/app.js`.
