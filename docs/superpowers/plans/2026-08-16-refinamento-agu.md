# Refinamento Visual Estilo AGU Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restructure the elevated-card redesign (already on this branch)
to match the AGU site's real structure — one strong shadow on an outer
card, flat bordered sections inside — and convert short-option `<select>`
fields into visible pill buttons, plus shrink oversized section/label
typography.

**Architecture:** Two CSS/markup-only tasks (no new logic) followed by one
task with real new interactive behavior (pill click handling in
`criarCampoInput`). No new pure functions, no data model change —
`js/gerador-docx.js` and `js/modelo.js` stay untouched.

**Tech Stack:** Vanilla CSS, vanilla DOM (unchanged from the rest of the
project). No build step, no new dependency.

## Global Constraints

- No servidor, sem build, sem npm no artefato final (unchanged project rule).
- Os 31 testes existentes (`node --test tests/*.test.mjs`) devem continuar
  passando sem alteração — nenhuma função pura muda de assinatura.
- Campos `select` com mais de 4 opções continuam usando `<select>` nativo
  — pills são só para grupos curtos (2–4 opções), conforme decisão de
  escopo do design.
- O valor gravado em `respostas[campo.id]` continua sendo a mesma string
  de opção de sempre (`"Sim"`, `"Não"`, `"Média"` etc.) — os pills mudam
  só a interface, não o contrato de dados consumido por
  `js/gerador-docx.js`.

---

### Task 1: CSS — cartão externo, subcartões planos, pills, tipografia

**Files:**
- Modify: `css/estilo.css`

**Interfaces:**
- Produces: as classes `.cartao-painel`, `.cabecalho-painel`,
  `.corpo-painel`, `.opcoes-pill`, `.pill`, `.pill.ativo`, usadas pelas
  Tasks 2 e 3. `.card-secao` e `.sumario-card` mudam de aparência (sombra
  → fundo+borda), mas mantêm os mesmos nomes de classe já consumidos por
  `js/render.js` — nenhuma mudança de nome de classe existente.

- [ ] **Step 1: Reescrever `css/estilo.css` por completo**

Conteúdo atual (para referência exata):

```css
body { background: #f8f8f8; }

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

Substituir pelo conteúdo completo abaixo:

```css
body { background: #f8f8f8; }

.header-logo { height: 40px; margin-right: 1rem; }
.header-titulo { font-weight: 600; font-size: 1.1rem; }
.campo-parametro { border-left: 3px solid #1351b4; padding-left: .5rem; margin-bottom: .75rem; }
.campo-especificacao { border-left: 3px solid #ffcd07; padding-left: .5rem; margin-bottom: .75rem; }
.campo-parametro label, .campo-especificacao label {
  display: block;
  font-size: .8125rem;
  margin-bottom: .25rem;
}
[hidden] { display: none !important; }

.cartao-painel {
  background: #ffffff;
  border-radius: 8px;
  box-shadow: 0 4px 16px rgba(0, 0, 0, .08);
  max-width: 900px;
  margin: 0 auto;
}
.cabecalho-painel {
  padding: 20px 28px;
  border-bottom: 1px solid #ededed;
  font-weight: 700;
  font-size: .9375rem;
  color: #333333;
}
.corpo-painel { padding: 20px 28px; }

.card-secao {
  background: #f8f8f8;
  border: 1px solid #ededed;
  border-radius: 8px;
  padding: 14px 16px;
  margin-bottom: 10px;
}
.card-secao h2 {
  margin: 0 0 10px;
  color: #071D41;
  font-size: .8125rem;
  font-weight: 700;
  text-transform: uppercase;
  letter-spacing: .02em;
}

.sumario-card {
  background: #f8f8f8;
  border: 1px solid #ededed;
  border-radius: 8px;
  padding: 8px;
  margin-bottom: 10px;
}

.opcoes-pill {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}
.opcoes-pill .pill {
  border: 1px solid #C5D4EB;
  border-radius: 6px;
  padding: 6px 14px;
  font-size: .8125rem;
  color: #1351b4;
  background: #ffffff;
  cursor: pointer;
}
.opcoes-pill .pill.ativo {
  background: #1351b4;
  color: #ffffff;
  border-color: #1351b4;
  font-weight: 600;
}

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

- [ ] **Step 2: Rodar a suíte de testes completa (nenhuma quebra esperada)**

Run: `node --test tests/*.test.mjs`
Expected: `tests 31`, `pass 31`, `fail 0` — CSS não afeta nenhum teste.

- [ ] **Step 3: Commit**

```bash
git add css/estilo.css
git commit -m "Restyle sections as flat sub-cards, add outer-card and pill CSS"
```

---

### Task 2: Cartão externo envolvendo o painel

**Files:**
- Modify: `js/app.js`

**Interfaces:**
- Consumes: as classes `.cartao-painel`, `.cabecalho-painel`,
  `.corpo-painel` da Task 1.
- Produces: nenhuma interface nova para outras tarefas — só reestrutura o
  HTML estático que `montarTelaPainel` escreve. `document.getElementById`
  em `#sumario`, `#secoes`, `#barra-progresso`, `#btn-gerar` continuam
  funcionando sem mudança, porque os ids não mudam, só o aninhamento.

- [ ] **Step 1: Localizar e substituir o template de `montarTelaPainel`**

Em `js/app.js`, dentro da função `montarTelaPainel`, o template atual é:

```js
  painel.innerHTML = `
    <nav id="sumario" class="br-list"></nav>
    <div id="secoes"></div>
    <div id="barra-progresso" class="barra-progresso" role="progressbar" aria-label="Progresso de preenchimento das especificações" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
      <div class="barra-progresso-trilho"><div class="barra-progresso-preenchimento"></div></div>
      <p class="barra-progresso-texto"></p>
    </div>
    <button id="btn-gerar" class="br-button primary" type="button">Gerar Documento</button>
  `;
```

Substituir por (envolve todo o conteúdo existente num `.cartao-painel`,
sem tirar nem adicionar nenhum elemento além do wrapper):

```js
  painel.innerHTML = `
    <div class="cartao-painel">
      <div class="cabecalho-painel">Termo de Referência — Aquisição de Bens</div>
      <div class="corpo-painel">
        <nav id="sumario" class="br-list"></nav>
        <div id="secoes"></div>
        <div id="barra-progresso" class="barra-progresso" role="progressbar" aria-label="Progresso de preenchimento das especificações" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
          <div class="barra-progresso-trilho"><div class="barra-progresso-preenchimento"></div></div>
          <p class="barra-progresso-texto"></p>
        </div>
        <button id="btn-gerar" class="br-button primary" type="button">Gerar Documento</button>
      </div>
    </div>
  `;
```

- [ ] **Step 2: Verificar sintaxe**

Run: `node --check js/app.js`
Expected: sem saída (sucesso silencioso).

- [ ] **Step 3: Rodar a suíte de testes completa (nenhuma quebra esperada)**

Run: `node --test tests/*.test.mjs`
Expected: `tests 31`, `pass 31`, `fail 0`.

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "Wrap painel content in a single elevated outer card"
```

---

### Task 3: Campos de poucas opções viram pills clicáveis

**Files:**
- Modify: `js/render.js` (`criarCampoInput`)

**Interfaces:**
- Consumes: `campo.input === "select"`, `campo.opcoes` (array já existente
  em `js/dados-tr-bens.js`, sem mudança nesta tarefa), `aoMudar(campoId, valor)`
  — mesmo callback já usado por todos os outros tipos de campo.
- Produces: nenhuma interface nova para outras tarefas. O valor gravado em
  `respostas[campo.id]` continua sendo a string da opção escolhida — igual
  ao que o `<select>` já produzia.

- [ ] **Step 1: Localizar e substituir `criarCampoInput`**

Em `js/render.js`, a função atual é:

```js
function criarCampoInput(campo, valorAtual, aoMudar) {
  const wrapper = document.createElement("div");
  wrapper.className = campo.tipo === "especificacao" ? "campo-especificacao" : "campo-parametro";

  const label = document.createElement("label");
  label.textContent = campo.pergunta;
  wrapper.appendChild(label);

  if (campo.ajuda) {
    const ajuda = document.createElement("span");
    ajuda.className = "br-tooltip";
    ajuda.setAttribute("data-text", campo.ajuda);
    ajuda.textContent = "?";
    wrapper.appendChild(ajuda);
  }

  let input;
  if (campo.input === "textarea") {
    input = document.createElement("textarea");
  } else if (campo.input === "select") {
    input = document.createElement("select");
    for (const opcao of campo.opcoes) {
      const optEl = document.createElement("option");
      optEl.value = opcao;
      optEl.textContent = opcao;
      input.appendChild(optEl);
    }
  } else {
    input = document.createElement("input");
    input.type = campo.input === "numero" ? "number" : "text";
  }
  input.id = campo.id;
  input.value = valorAtual ?? "";
  input.addEventListener("input", () => aoMudar(campo.id, input.value));
  wrapper.appendChild(input);
  return wrapper;
}
```

Substituir por (acrescenta um ramo antes da lógica de `<select>` — os
demais tipos de campo, `textarea`/`numero`/`texto`, ficam idênticos):

```js
function criarCampoInput(campo, valorAtual, aoMudar) {
  const wrapper = document.createElement("div");
  wrapper.className = campo.tipo === "especificacao" ? "campo-especificacao" : "campo-parametro";

  const label = document.createElement("label");
  label.textContent = campo.pergunta;
  wrapper.appendChild(label);

  if (campo.ajuda) {
    const ajuda = document.createElement("span");
    ajuda.className = "br-tooltip";
    ajuda.setAttribute("data-text", campo.ajuda);
    ajuda.textContent = "?";
    wrapper.appendChild(ajuda);
  }

  if (campo.input === "select" && campo.opcoes.length <= 4) {
    const grupo = document.createElement("div");
    grupo.className = "opcoes-pill";
    grupo.id = campo.id;
    for (const opcao of campo.opcoes) {
      const botao = document.createElement("button");
      botao.type = "button";
      botao.className = "pill" + (valorAtual === opcao ? " ativo" : "");
      botao.textContent = opcao;
      botao.addEventListener("click", () => {
        for (const irmao of grupo.querySelectorAll(".pill")) {
          irmao.classList.remove("ativo");
        }
        botao.classList.add("ativo");
        aoMudar(campo.id, opcao);
      });
      grupo.appendChild(botao);
    }
    wrapper.appendChild(grupo);
    return wrapper;
  }

  let input;
  if (campo.input === "textarea") {
    input = document.createElement("textarea");
  } else if (campo.input === "select") {
    input = document.createElement("select");
    for (const opcao of campo.opcoes) {
      const optEl = document.createElement("option");
      optEl.value = opcao;
      optEl.textContent = opcao;
      input.appendChild(optEl);
    }
  } else {
    input = document.createElement("input");
    input.type = campo.input === "numero" ? "number" : "text";
  }
  input.id = campo.id;
  input.value = valorAtual ?? "";
  input.addEventListener("input", () => aoMudar(campo.id, input.value));
  wrapper.appendChild(input);
  return wrapper;
}
```

- [ ] **Step 2: Verificar sintaxe**

Run: `node --check js/render.js`
Expected: sem saída (sucesso silencioso).

- [ ] **Step 3: Rodar a suíte de testes completa (nenhuma quebra esperada)**

Run: `node --test tests/*.test.mjs`
Expected: `tests 31`, `pass 31`, `fail 0` — `criarCampoInput` não é uma
função exportada nem coberta por teste automatizado (é interna a
`render.js`, chamada só por `renderizarSecoes`), então a mudança não
afeta nenhum teste existente.

- [ ] **Step 4: Verificação manual no navegador**

Servir o projeto por HTTP (`python -m http.server` a partir da raiz do
projeto — ver `README.md`), avançar até o painel. Confirmar visualmente:
- todo o conteúdo do painel aparece dentro de um único cartão branco com
  sombra forte, com um cabeçalho "Termo de Referência — Aquisição de Bens"
  no topo;
- cada seção e o sumário aparecem como blocos planos (fundo cinza claro,
  borda fina), sem sombra própria;
- os títulos de seção ("01. DO OBJETO" etc.) e os rótulos de campo estão
  visivelmente menores que antes;
- campos como "Metodologia da estimativa de preço" (seção 01) e "Será
  exigida garantia contratual?" (seção 07) aparecem como botões clicáveis
  visíveis, não como um menu suspenso — clicar num botão marca ele como
  selecionado (fundo azul) e desmarca os outros do mesmo grupo;
- confirmar, no painel de conferência (botão "Gerar Documento"), que uma
  resposta escolhida via pill é reconhecida corretamente como preenchida
  (ex.: responder "Sim" em "Será exigida garantia contratual?" faz o campo
  dependente "Percentual da garantia" aparecer, do mesmo jeito que já
  funcionava com o `<select>`).

- [ ] **Step 5: Commit**

```bash
git add js/render.js
git commit -m "Convert short-option select fields into visible pill buttons"
```

---

## Self-Review

**Cobertura da spec:** Task 1 cobre o cartão externo/subcartões planos e a
tipografia menor (`docs/superpowers/specs/2026-08-16-refinamento-agu-design.md`);
Task 2 cobre o cabeçalho interno do cartão; Task 3 cobre os pills para
campos de até 4 opções, mantendo `<select>` para listas maiores (nenhum
campo do TR-Bens hoje tem mais de 4 opções, então esse caminho não é
exercitado nesta rodada, mas o código já o suporta). O esclarecimento sobre
pills não pré-selecionarem nenhuma opção (diferente do `<select>` nativo)
está documentado no design e reflete o comportamento real do código.

**Placeholders:** nenhum "TBD"/"implementar depois" — todo trecho de
código é o conteúdo final exato a escrever.

**Consistência de tipos:** `criarCampoInput(campo, valorAtual, aoMudar)`
mantém a mesma assinatura; `aoMudar(campo.id, opcao)` usa exatamente o
mesmo contrato `(id, valor)` já consumido por `renderizarSecoes` e por
`js/app.js`. Nenhuma função exportada de `js/render.js` muda de nome ou
assinatura.
