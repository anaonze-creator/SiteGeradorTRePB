# Motor TR-Serviços Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add the TR-Serviços instrument to the existing gerador (TR-Bens is already
live), reusing the same declarative motor and extending it with two new
generalizations (multi-paragraph conditional blocks, N-way named variants) and
one new field type (`tabela-linhas`), grounded in the official minuta SEI
55379668.

**Architecture:** Same as TR-Bens — static site, `.docx` bookmark substitution
in the browser (`js/gerador-docx.js`), declarative section/field data
(`js/dados-tr-servicos.js`), a `.docx` template built once by a Python script
(`templates/build_template_tr_servicos.py`). `js/app.js` and `js/triagem.js`
become instrument-agnostic so both TR-Bens and TR-Serviços run from the same
app shell.

**Tech Stack:** Vanilla JS (ES modules), `python-docx` for template
generation, `node --test` for unit tests (no build step, no package.json —
run tests with `node --test tests/`).

## Global Constraints

- Nenhum texto do TR é inventado: todo trecho fixo vem transcrito da minuta
  `documentos de referencia/SEI_55379668_Conab___RLC__Termo_de_Referencia__TR____Servicos.pdf`
  (extraída com `pdftotext -layout -enc UTF-8`). Notas explicativas (destacadas
  em vermelho na minuta) são suprimidas do documento final — a própria minuta
  instrui isso no cabeçalho.
- Todo campo é `tipo: "parametro"` ou `tipo: "especificacao"` (nunca um
  terceiro valor) — a barra de progresso e `pendenciasParametro`/`pendencias`
  em `js/modelo.js` dependem disso. `input` (não `tipo`) é o que varia por
  widget: `"texto"`, `"numero"`, `"moeda"`, `"textarea"`, `"select"`,
  `"tabela-linhas"`.
- Todo campo com `blocoTexto: true` precisa também de `padrao` (o valor usado
  quando `respostas[campo.id]` está vazio) e de `opcoes` (a lista de variantes
  — cada uma vira um bookmark `${campoId}_BLOCO_${VARIANTE_NORMALIZADA}` no
  `.docx`).
- Regressão zero no TR-Bens: toda mudança em `js/modelo.js` e
  `js/gerador-docx.js` precisa manter os testes existentes
  (`tests/modelo.test.mjs`, `tests/gerador-docx.test.mjs`) passando sem
  edição, exceto a migração pontual do Task 2.

---

### Task 1: Bloco condicional multi-parágrafo em `gerador-docx.js`

**Files:**
- Modify: `js/gerador-docx.js:1-46` (`localizarParagrafoDoBookmark`)
- Test: `tests/gerador-docx.test.mjs` (adicionar casos, não remover os existentes)

**Interfaces:**
- Consumes: nada de tasks anteriores.
- Produces: `localizarParagrafoDoBookmark(xml, nomeBookmark)` continua
  retornando `{ inicioParagrafo, fimParagrafo }`, mas `fimParagrafo` agora é o
  fim do **último** `<w:p>` que contém o `bookmarkEnd` correspondente (mesmo
  `w:id` do `bookmarkStart` achado), não necessariamente o mesmo parágrafo do
  `bookmarkStart`. `substituirBookmark` e `removerBlocoBookmark` (que já
  consomem essa função) não mudam de assinatura.

- [ ] **Step 1: Escrever o teste que falha (bloco de 2 parágrafos)**

```js
// em tests/gerador-docx.test.mjs, dentro do describe/test existente:
test("removerBlocoBookmark remove um bloco que abrange MAIS de um parágrafo", () => {
  const xmlFalso =
    '<w:p><w:r><w:t>antes</w:t></w:r></w:p>' +
    '<w:p><w:bookmarkStart w:id="9" w:name="S08_TipoFiscalizacao_BLOCO_COMISSAO_DE_FISCALIZACAO"/>' +
    '<w:r><w:t>parágrafo 1 do bloco</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>parágrafo 2 do bloco</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>parágrafo 3 do bloco</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>' +
    '<w:p><w:r><w:t>depois</w:t></w:r></w:p>';
  const resultado = removerBlocoBookmark(xmlFalso, "S08_TipoFiscalizacao_BLOCO_COMISSAO_DE_FISCALIZACAO");
  assert.match(resultado, /antes/);
  assert.match(resultado, /depois/);
  assert.doesNotMatch(resultado, /parágrafo 1 do bloco/);
  assert.doesNotMatch(resultado, /parágrafo 2 do bloco/);
  assert.doesNotMatch(resultado, /parágrafo 3 do bloco/);
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: FAIL — hoje `fimParagrafo` para no primeiro `</w:p>` depois do
`bookmarkStart`, então "parágrafo 2" e "parágrafo 3" sobrevivem.

- [ ] **Step 3: Generalizar `localizarParagrafoDoBookmark`**

```js
function localizarParagrafoDoBookmark(xml, nomeBookmark) {
  const marcador = `w:name="${nomeBookmark}"`;
  const posNome = xml.indexOf(marcador);
  if (posNome === -1) return null;

  const inicioParagrafo = xml.lastIndexOf("<w:p>", posNome) !== -1
    ? xml.lastIndexOf("<w:p>", posNome)
    : xml.lastIndexOf("<w:p ", posNome);

  // Acha o w:id do bookmarkStart achado, para localizar o bookmarkEnd
  // correspondente (mesmo id) — que pode estar num parágrafo mais adiante,
  // quando o bloco abrange vários parágrafos.
  const inicioTagStart = xml.lastIndexOf("<w:bookmarkStart", posNome);
  const fimTagStart = xml.indexOf(">", inicioTagStart) + 1;
  const tagStart = xml.slice(inicioTagStart, fimTagStart);
  const idMatch = tagStart.match(/w:id="([^"]+)"/);
  if (!idMatch) return null;
  const bookmarkId = idMatch[1];

  const regexEnd = new RegExp(`<w:bookmarkEnd[^>]*w:id="${bookmarkId}"[^>]*/>`);
  const matchEnd = regexEnd.exec(xml.slice(posNome));
  if (!matchEnd) return null;
  const posEnd = posNome + matchEnd.index;

  const fimParagrafo = xml.indexOf("</w:p>", posEnd) + "</w:p>".length;
  return { inicioParagrafo, fimParagrafo };
}
```

- [ ] **Step 4: Rodar de novo e confirmar que passa, sem quebrar os testes antigos**

Run: `node --test tests/gerador-docx.test.mjs tests/modelo.test.mjs`
Expected: PASS em todos, incluindo os testes de blocos de 1 parágrafo já
existentes (`S07_ExigeGarantia_BLOCO_NAO` etc.) — o range de 1 parágrafo é o
caso onde `bookmarkStart` e `bookmarkEnd` batem no mesmo `<w:p>`, coberto pelo
mesmo código.

- [ ] **Step 5: Commit**

```bash
git add js/gerador-docx.js tests/gerador-docx.test.mjs
git commit -m "Generalize bookmark block lookup to span multiple paragraphs"
```

---

### Task 2: Variante nomeada (N-vias) substitui o mecanismo binário Sim/Não

**Files:**
- Modify: `js/modelo.js` (nova função `controladoresBlocoVariante`)
- Modify: `js/gerador-docx.js` (substitui `nomeBlocoSim`/`nomeBlocoNao` e o
  loop que usa `controladoresCondicionais`)
- Modify: `js/dados-tr-bens.js` (adiciona `blocoTexto: true, padrao: "Não"`
  aos 5 campos que já tinham bloco condicional)
- Modify: `js/app.js:6,62-65` (troca o import/uso de `controladoresCondicionais`
  onde ele hoje é reaproveitado para saber "que campo dispara rerender" — ver
  Step 5)
- Test: `tests/modelo.test.mjs`, `tests/gerador-docx.test.mjs`

**Interfaces:**
- Consumes: nada de tasks anteriores (Task 1 já mesclado, mas independente).
- Produces:
  - `controladoresBlocoVariante(secoes): Array<Campo>` em `js/modelo.js` —
    retorna os objetos de campo (não só IDs) que têm `blocoTexto: true`.
  - `normalizarVariante(valor: string): string` e
    `nomeBlocoVariante(campoId: string, opcao: string): string` em
    `js/gerador-docx.js` (não exportadas — uso interno de `gerarDocx`).
  - Campo de dados com o formato:
    `{ id, tipo: "parametro", input: "select", opcoes: [...], blocoTexto: true, padrao: "<uma das opcoes>" }`.

- [ ] **Step 1: Escrever o teste que falha em `modelo.js`**

```js
// em tests/modelo.test.mjs
test("controladoresBlocoVariante devolve os campos com blocoTexto", () => {
  const secoes = [{
    id: "05", titulo: "T", campos: [
      { id: "c1", tipo: "parametro", input: "select", opcoes: ["A", "B"], pergunta: "?" },
      { id: "c2", tipo: "parametro", input: "select", opcoes: ["Sim", "Não"],
        blocoTexto: true, padrao: "Não", pergunta: "?" },
    ],
  }];
  const resultado = controladoresBlocoVariante(secoes);
  assert.equal(resultado.length, 1);
  assert.equal(resultado[0].id, "c2");
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/modelo.test.mjs`
Expected: FAIL — `controladoresBlocoVariante is not a function`.

- [ ] **Step 3: Implementar em `js/modelo.js`**

```js
// Campos que controlam um bloco de texto alternativo inteiro no documento
// (Sim/Não ou N variantes nomeadas) — diferente de `mostrarSe`, que só
// esconde/mostra campos individuais no formulário. Um campo pode fazer as
// duas coisas ao mesmo tempo (ex.: controlar visibilidade de um campo E
// trocar um parágrafo inteiro do documento).
export function controladoresBlocoVariante(secoes) {
  const resultado = [];
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.blocoTexto) resultado.push(campo);
    }
  }
  return resultado;
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/modelo.test.mjs`
Expected: PASS.

- [ ] **Step 5: Migrar `js/gerador-docx.js` para o mecanismo de variante**

Substituir (remover as linhas 48-54 atuais — `nomeBlocoSim`/`nomeBlocoNao` —
e o loop das linhas 77-83 do arquivo original) por:

```js
import { campoVisivel, controladoresBlocoVariante, camposOcultos } from "./modelo.js";

function normalizarVariante(valor) {
  return valor
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, "_");
}

function nomeBlocoVariante(campoId, opcao) {
  return `${campoId}_BLOCO_${normalizarVariante(opcao)}`;
}
```

E, dentro de `gerarDocx`, trocar o loop:

```js
for (const campo of controladoresBlocoVariante(secoes)) {
  const respondida = respostas[campo.id] || campo.padrao;
  for (const opcao of campo.opcoes) {
    if (opcao !== respondida) {
      documentXml = removerBlocoBookmark(documentXml, nomeBlocoVariante(campo.id, opcao));
    }
  }
}
```

(`camposOcultos`/`campoVisivel` continuam exatamente como estavam — eles
tratam de `mostrarSe`, mecanismo separado e já genérico, sem relação com
`blocoTexto`.)

- [ ] **Step 6: Migrar os 5 campos existentes em `js/dados-tr-bens.js`**

Adicionar `blocoTexto: true, padrao: "Não"` a cada um destes (mantendo todo
o resto do campo igual): `S02_AplicaExcecaoMEEPP`, `S05_ExigeAmostra`,
`S07_ExigeGarantia`, `S12_Reajustavel`, `S13_PermiteSubcontratacao`. Exemplo
para o primeiro:

```js
{ id: "S02_AplicaExcecaoMEEPP", tipo: "parametro", input: "select",
  pergunta: "O valor estimado é até R$ 80.000,00 e a contratação se enquadra em exceção à exclusividade ME/EPP?",
  opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
```

- [ ] **Step 7: Rodar TODOS os testes existentes — devem passar sem edição**

Run: `node --test tests/`
Expected: PASS em `gerador-docx.test.mjs` (os testes de S07/S13/S05/S12/S02
usam exatamente os nomes de bookmark `_BLOCO_SIM`/`_BLOCO_NAO`, que
`normalizarVariante("Sim")`/`normalizarVariante("Não")` reproduzem
identicamente) e em `modelo.test.mjs`.

- [ ] **Step 8: Atualizar `js/app.js` para não depender mais de `controladoresCondicionais` para o mesmo fim que `controladoresBlocoVariante` cobre**

`controladoresCondicionais` (baseado em `mostrarSe`) continua correto para
"que campo, ao mudar, deve disparar um rerender do formulário" — isso não
muda. Nenhuma edição é necessária em `js/app.js` neste task; ele já usa
`controladoresCondicionais` para essa finalidade específica e continuará
funcionando sem alteração porque `mostrarSe` é ortogonal a `blocoTexto`.
(Este step existe só para deixar explícito, no plano, que a dupla
responsabilidade antiga foi separada e `app.js` já está no lado certo dela —
nenhuma mudança de código aqui.)

- [ ] **Step 9: Commit**

```bash
git add js/modelo.js js/gerador-docx.js js/dados-tr-bens.js tests/modelo.test.mjs
git commit -m "Replace binary Sim/Não block mechanism with named-variant mechanism"
```

---

### Task 3: Campo `input: "tabela-linhas"` (infrações extras da Seção 15)

**Files:**
- Modify: `js/render.js` (`criarCampoInput`)
- Modify: `js/gerador-docx.js` (clonagem de linha de tabela)
- Modify: `css/estilo.css` (ou arquivo CSS existente — checar nome real
  antes de editar; ver Step 0)
- Test: `tests/gerador-docx.test.mjs`, novo `tests/render.test.mjs` se ainda
  não existir cobertura de DOM (checar Step 0)

**Interfaces:**
- Consumes: nada das tasks anteriores diretamente (independente).
- Produces:
  - Formato de dado: `respostas[campo.id]` é um array de objetos
    `{ [colunaId]: valor }`, ex.: `[{ descricao: "...", grau: "3" }]`.
  - `inserirLinhasTabela(documentXml, campo, linhas): string` em
    `js/gerador-docx.js` (não exportada — chamada de dentro de `gerarDocx`).
  - Convenção de bookmark no template: uma linha-modelo de tabela marcada
    com bookmark `${campoId}_LINHA_MODELO`, com placeholders de texto
    `((NOMEDACOLUNA))` (maiúsculo) em cada célula — ex.: `((DESCRICAO))`,
    `((GRAU))`.

- [ ] **Step 0: Confirmar nome real do arquivo CSS e do padrão de teste de DOM**

Run: `ls "css/"` e `ls "tests/" | grep -i render`
Se não houver `tests/render.test.mjs`, este task não adiciona teste de DOM
(o projeto não tem `jsdom` nas dependências — não introduzir uma dependência
nova só para este campo). A cobertura funcional deste campo fica no teste de
`gerador-docx.js` do Step 4 adiante, que é o caminho de maior risco real
(geração do `.docx`).

- [ ] **Step 1: Adicionar a UI em `js/render.js`**

Em `criarCampoInput`, antes do bloco `if (campo.input === "select" && ...)`:

```js
if (campo.input === "tabela-linhas") {
  return criarCampoTabelaLinhas(campo, valorAtual, aoMudar);
}
```

E a função nova, no mesmo arquivo:

```js
function criarCampoTabelaLinhas(campo, valorAtual, aoMudar) {
  const wrapper = document.createElement("div");
  wrapper.className = "campo-parametro campo-tabela-linhas";

  const label = document.createElement("label");
  label.textContent = campo.pergunta;
  wrapper.appendChild(label);

  const linhas = Array.isArray(valorAtual) ? valorAtual.map((l) => ({ ...l })) : [];

  const tabela = document.createElement("table");
  tabela.className = "br-table tabela-linhas-editavel";
  const corpo = document.createElement("tbody");
  tabela.appendChild(corpo);
  wrapper.appendChild(tabela);

  function renderizarLinhas() {
    corpo.innerHTML = "";
    linhas.forEach((linha, indice) => {
      const tr = document.createElement("tr");
      for (const coluna of campo.colunas) {
        const td = document.createElement("td");
        let input;
        if (coluna.input === "select") {
          input = document.createElement("select");
          for (const opcao of coluna.opcoes) {
            const optEl = document.createElement("option");
            optEl.value = opcao;
            optEl.textContent = opcao;
            input.appendChild(optEl);
          }
          input.value = linha[coluna.id] ?? coluna.opcoes[0];
        } else {
          input = document.createElement("input");
          input.type = "text";
          input.value = linha[coluna.id] ?? "";
        }
        input.addEventListener("input", () => {
          linhas[indice] = { ...linhas[indice], [coluna.id]: input.value };
          aoMudar(campo.id, linhas.map((l) => ({ ...l })));
        });
        td.appendChild(input);
        tr.appendChild(td);
      }
      const tdRemover = document.createElement("td");
      const botaoRemover = document.createElement("button");
      botaoRemover.type = "button";
      botaoRemover.className = "br-button small";
      botaoRemover.textContent = "Remover";
      botaoRemover.addEventListener("click", () => {
        linhas.splice(indice, 1);
        aoMudar(campo.id, linhas.map((l) => ({ ...l })));
        renderizarLinhas();
      });
      tdRemover.appendChild(botaoRemover);
      tr.appendChild(tdRemover);
      corpo.appendChild(tr);
    });
  }
  renderizarLinhas();

  const botaoAdicionar = document.createElement("button");
  botaoAdicionar.type = "button";
  botaoAdicionar.className = "br-button secondary";
  botaoAdicionar.textContent = "Adicionar linha";
  botaoAdicionar.addEventListener("click", () => {
    const linhaVazia = {};
    for (const coluna of campo.colunas) {
      linhaVazia[coluna.id] = coluna.input === "select" ? coluna.opcoes[0] : "";
    }
    linhas.push(linhaVazia);
    aoMudar(campo.id, linhas.map((l) => ({ ...l })));
    renderizarLinhas();
  });
  wrapper.appendChild(botaoAdicionar);

  return wrapper;
}
```

- [ ] **Step 2: Adicionar CSS mínimo (arquivo confirmado no Step 0)**

```css
.tabela-linhas-editavel { width: 100%; border-collapse: collapse; margin: 0.5rem 0; }
.tabela-linhas-editavel td { padding: 0.25rem; }
.campo-tabela-linhas .br-button.small { padding: 0.25rem 0.5rem; font-size: 0.85rem; }
```

- [ ] **Step 3: Escrever o teste que falha em `gerador-docx.js` (clonagem de linha)**

```js
// em tests/gerador-docx.test.mjs
test("inserirLinhasTabela clona a linha-modelo uma vez por item e remove o modelo", async () => {
  const campoFalso = {
    id: "S15_InfracoesExtras",
    colunas: [{ id: "descricao" }, { id: "grau" }],
  };
  const xmlFalso =
    '<w:tbl>' +
    '<w:tr><w:tc><w:p><w:r><w:t>fixa</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:bookmarkStart w:id="7" w:name="S15_InfracoesExtras_LINHA_MODELO"/>' +
    '<w:r><w:t>((DESCRICAO))</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>((GRAU))</w:t></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl>';
  const linhas = [{ descricao: "Deixar de treinar equipe", grau: "1" }, { descricao: "Atraso reiterado", grau: "3" }];
  const resultado = inserirLinhasTabela(xmlFalso, campoFalso, linhas);
  assert.match(resultado, /fixa/);
  assert.match(resultado, /Deixar de treinar equipe/);
  assert.match(resultado, /Atraso reiterado/);
  assert.doesNotMatch(resultado, /\(\(DESCRICAO\)\)/);
  assert.doesNotMatch(resultado, /\(\(GRAU\)\)/);
  assert.doesNotMatch(resultado, /LINHA_MODELO/);
});

test("inserirLinhasTabela sem linhas extras remove a linha-modelo e não deixa placeholder", () => {
  const campoFalso = { id: "S15_InfracoesExtras", colunas: [{ id: "descricao" }, { id: "grau" }] };
  const xmlFalso =
    '<w:tr><w:tc><w:p><w:bookmarkStart w:id="1" w:name="S15_InfracoesExtras_LINHA_MODELO"/>' +
    '<w:r><w:t>((DESCRICAO))</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p></w:tc></w:tr>';
  const resultado = inserirLinhasTabela(xmlFalso, campoFalso, []);
  assert.doesNotMatch(resultado, /\(\(DESCRICAO\)\)/);
});
```

- [ ] **Step 4: Rodar e confirmar que falha**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: FAIL — `inserirLinhasTabela is not defined` (ainda não exportada).

- [ ] **Step 5: Implementar em `js/gerador-docx.js`**

```js
function localizarLinhaTabela(xml, nomeBookmark) {
  const marcador = `w:name="${nomeBookmark}"`;
  const posNome = xml.indexOf(marcador);
  if (posNome === -1) return null;
  const inicioLinha = xml.lastIndexOf("<w:tr>", posNome) !== -1
    ? xml.lastIndexOf("<w:tr>", posNome)
    : xml.lastIndexOf("<w:tr ", posNome);
  const fimLinha = xml.indexOf("</w:tr>", posNome) + "</w:tr>".length;
  return { inicioLinha, fimLinha };
}

function removerMarcacoesBookmark(trecho) {
  return trecho
    .replace(/<w:bookmarkStart[^>]*\/>/g, "")
    .replace(/<w:bookmarkEnd[^>]*\/>/g, "");
}

export function inserirLinhasTabela(documentXml, campo, linhas) {
  const posicao = localizarLinhaTabela(documentXml, `${campo.id}_LINHA_MODELO`);
  if (!posicao) return documentXml;
  const linhaModeloXml = documentXml.slice(posicao.inicioLinha, posicao.fimLinha);
  const linhasGeradas = linhas.map((linha) => {
    let linhaXml = removerMarcacoesBookmark(linhaModeloXml);
    for (const coluna of campo.colunas) {
      const marcador = `((${coluna.id.toUpperCase()}))`;
      linhaXml = linhaXml.split(marcador).join(escaparXml(String(linha[coluna.id] || "")));
    }
    return linhaXml;
  }).join("");
  return documentXml.slice(0, posicao.inicioLinha) + linhasGeradas + documentXml.slice(posicao.fimLinha);
}
```

E, dentro de `gerarDocx`, no loop principal que hoje faz
`substituirBookmark` para cada campo visível, pular (`continue`) os campos
com `campo.input === "tabela-linhas"` e, depois do loop, chamar:

```js
for (const secao of secoes) {
  for (const campo of secao.campos) {
    if (campo.input === "tabela-linhas") {
      const linhas = Array.isArray(respostas[campo.id]) ? respostas[campo.id] : [];
      documentXml = inserirLinhasTabela(documentXml, campo, linhas);
    }
  }
}
```

- [ ] **Step 6: Rodar e confirmar que passa**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: PASS, incluindo todos os testes anteriores (TR-Bens não tem campo
`tabela-linhas`, então `inserirLinhasTabela` nunca acha o bookmark e não
altera nada para ele).

- [ ] **Step 7: Commit**

```bash
git add js/render.js js/gerador-docx.js css/*.css tests/gerador-docx.test.mjs
git commit -m "Add tabela-linhas field type for dynamic table rows (Seção 15 infrações)"
```

---

### Task 4: `js/dados-tr-servicos.js` — modelo de dados das 20 seções

**Files:**
- Create: `js/dados-tr-servicos.js`
- Test: `tests/dados-tr-servicos.test.mjs`

**Interfaces:**
- Consumes: formato de campo definido nas Tasks 2 e 3 (`blocoTexto`/`padrao`,
  `input: "tabela-linhas"`/`colunas`) e o já existente (`mostrarSe`).
- Produces: `export const SECOES_TR_SERVICOS` — mesma forma de
  `SECOES_TR_BENS`, consumida por `js/app.js` (Task 6) e
  `js/gerador-docx.js`.

- [ ] **Step 1: Escrever o teste estrutural (falha por o arquivo não existir)**

```js
// tests/dados-tr-servicos.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";
import { controladoresBlocoVariante } from "../js/modelo.js";

test("SECOES_TR_SERVICOS tem 20 seções com ids únicos de 01 a 20", () => {
  assert.equal(SECOES_TR_SERVICOS.length, 20);
  const ids = SECOES_TR_SERVICOS.map((s) => s.id);
  assert.deepEqual(ids, [
    "01","02","03","04","05","06","07","08","09","10",
    "11","12","13","14","15","16","17","18","19","20",
  ]);
});

test("todo campo com blocoTexto tem padrao dentro de opcoes", () => {
  for (const campo of controladoresBlocoVariante(SECOES_TR_SERVICOS)) {
    assert.ok(campo.opcoes.includes(campo.padrao),
      `${campo.id}: padrao "${campo.padrao}" não está em opcoes`);
  }
});

test("todo campo tabela-linhas declara colunas", () => {
  for (const secao of SECOES_TR_SERVICOS) {
    for (const campo of secao.campos) {
      if (campo.input === "tabela-linhas") {
        assert.ok(Array.isArray(campo.colunas) && campo.colunas.length > 0, campo.id);
      }
    }
  }
});

test("ids de campo são únicos entre todas as seções", () => {
  const todosIds = SECOES_TR_SERVICOS.flatMap((s) => s.campos.map((c) => c.id));
  assert.equal(todosIds.length, new Set(todosIds).size);
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/dados-tr-servicos.test.mjs`
Expected: FAIL — módulo não existe.

- [ ] **Step 3: Criar `js/dados-tr-servicos.js`**

```js
export const SECOES_TR_SERVICOS = [
  { id: "01", titulo: "DO OBJETO", campos: [
    { id: "S01_DescricaoObjeto", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva os serviços a serem contratados",
      ajuda: "Ex.: limpeza e conservação, manutenção predial, desenvolvimento de sistema, consultoria técnica." },
    { id: "S01_RegimeExecucao", tipo: "parametro", input: "texto",
      pergunta: "Regime de execução indireta (art. 208, IV, RLC)" },
    { id: "S01_ValorEstimado", tipo: "parametro", input: "moeda",
      pergunta: "Custo total estimado da prestação dos serviços (R$)" },
    { id: "S01_MetodologiaEstimativa", tipo: "parametro", input: "select",
      pergunta: "Metodologia da estimativa de preço",
      opcoes: ["Média", "Mediana", "Menor Preço"] },
    { id: "S01_CriterioJulgamento", tipo: "parametro", input: "select",
      pergunta: "Critério de julgamento das propostas",
      opcoes: ["Menor preço", "Maior desconto"] },
    { id: "S01_SigiloValor", tipo: "parametro", input: "select",
      pergunta: "O valor estimado será", opcoes: ["Público", "Sigiloso"] },
    { id: "S01_ModoDisputa", tipo: "parametro", input: "select",
      pergunta: "Modo de disputa", opcoes: ["Aberto", "Aberto e fechado"] },
    { id: "S01_ND", tipo: "parametro", input: "texto", pergunta: "Natureza de Despesa (ND)" },
    { id: "S01_PTRES", tipo: "parametro", input: "texto", pergunta: "PTRES" },
    { id: "S01_Fonte", tipo: "parametro", input: "texto", pergunta: "Fonte de recurso" },
  ]},
  { id: "02", titulo: "DA JUSTIFICATIVA", campos: [
    { id: "S02_Motivacao", tipo: "especificacao", input: "textarea",
      pergunta: "Qual a motivação da contratação?",
      ajuda: "Baseie-se no Documento de Origem da Demanda (DOD) — art. 129, II, RLC." },
    { id: "S02_AplicaExcecaoMEEPP", tipo: "parametro", input: "select",
      pergunta: "O valor estimado é até R$ 80.000,00 e a contratação se enquadra em exceção à exclusividade ME/EPP?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "S02_MotivoExcecaoMEEPP", tipo: "especificacao", input: "textarea",
      pergunta: "Qual exceção do art. 49 da LC 123/2006 se aplica e por quê?",
      mostrarSe: { S02_AplicaExcecaoMEEPP: "Sim" } },
  ]},
  { id: "03", titulo: "DO OBJETIVO", campos: [
    { id: "S03_Objetivo", tipo: "especificacao", input: "textarea",
      pergunta: "Qual o objetivo a ser atingido com a contratação?" },
  ]},
  { id: "04", titulo: "DA VIGÊNCIA DO CONTRATO", campos: [
    { id: "S04_TipoVigencia", tipo: "parametro", input: "select",
      pergunta: "Regime de vigência do contrato",
      opcoes: ["Até 1 ano, prorrogável", "Plurianual (até 5 anos)"],
      blocoTexto: true, padrao: "Até 1 ano, prorrogável",
      ajuda: "Vigência de até 1 ano segue os arts. 497/498 do RLC (prorrogável); vigência plurianual segue o art. 488." },
    { id: "S04_PrazoVigenciaMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo de vigência inicial (meses)",
      mostrarSe: { S04_TipoVigencia: "Até 1 ano, prorrogável" } },
    { id: "S04_PrazoVigenciaAnos", tipo: "parametro", input: "numero",
      pergunta: "Prazo de vigência inicial (anos, máx. 5)",
      mostrarSe: { S04_TipoVigencia: "Plurianual (até 5 anos)" } },
  ]},
  { id: "05", titulo: "DA SELEÇÃO DO FORNECEDOR", campos: [
    { id: "S05_PrazoExperienciaMinimaAnos", tipo: "parametro", input: "numero",
      pergunta: "Prazo mínimo de experiência anterior exigido para comprovação de capacidade técnica (anos)" },
    { id: "S05_QuantitativoServicoPrincipal", tipo: "especificacao", input: "textarea",
      pergunta: "Quantitativo mínimo exigido e identificação do(s) serviço(s) principal(is), para fins de comprovação de capacidade técnica",
      ajuda: "Não fixar parâmetro acima de 50% dos quantitativos, por analogia ao Acórdão 361/2017-TCU." },
    { id: "S05_ExigeInstalacaoEscritorio", tipo: "parametro", input: "select",
      pergunta: "Será exigida declaração de instalação de escritório na cidade da Matriz/Sureg?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "S05_MunicipioInstalacaoEscritorio", tipo: "parametro", input: "texto",
      pergunta: "Cidade em que o escritório deverá ser instalado",
      mostrarSe: { S05_ExigeInstalacaoEscritorio: "Sim" } },
    { id: "S05_TipoVistoria", tipo: "parametro", input: "select",
      pergunta: "A vistoria técnica será", opcoes: ["Obrigatória", "Facultativa"],
      blocoTexto: true, padrao: "Facultativa",
      ajuda: "Vistoria obrigatória exige justificativa própria nos autos, à luz do objeto licitado." },
    { id: "S05_HorarioVistoriaInicio", tipo: "parametro", input: "texto",
      pergunta: "Horário de início para agendamento da vistoria" },
    { id: "S05_HorarioVistoriaFim", tipo: "parametro", input: "texto",
      pergunta: "Horário de fim para agendamento da vistoria" },
    { id: "S05_TelefoneAgendamentoVistoria", tipo: "parametro", input: "texto",
      pergunta: "Telefone para agendamento da vistoria" },
  ]},
  { id: "06", titulo: "DA DESCRIÇÃO DOS SERVIÇOS", campos: [
    { id: "S06_DetalhamentoServicos", tipo: "especificacao", input: "textarea",
      pergunta: "Detalhe as funcionalidades, características, especificações quantitativas e qualitativas, condições e requisitos do serviço" },
    { id: "S06_LocalExecucao", tipo: "parametro", input: "texto",
      pergunta: "Local de execução dos serviços (Matriz/Sureg e endereço)" },
    { id: "S06_HorarioExecucao", tipo: "parametro", input: "texto",
      pergunta: "Horário de execução dos serviços" },
    { id: "S06_PrazoExecucaoDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo de execução dos serviços, contado da assinatura do contrato (dias)" },
    { id: "S06_RotinaExecucao", tipo: "especificacao", input: "textarea",
      pergunta: "Detalhe as atividades a serem executadas pela Contratada e a metodologia de trabalho",
      ajuda: "Inciso IV, art. 129 do RLC." },
    { id: "S06_ExigeGarantiaServicos", tipo: "parametro", input: "select",
      pergunta: "Será exigido período de garantia dos serviços?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "S06_GarantiaServicosMeses", tipo: "parametro", input: "numero",
      pergunta: "Período mínimo de garantia dos serviços (meses)",
      mostrarSe: { S06_ExigeGarantiaServicos: "Sim" } },
    { id: "S06_PrazoRefazimentoDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo máximo para refazer o serviço em garantia, após comunicação (dias)",
      mostrarSe: { S06_ExigeGarantiaServicos: "Sim" } },
    { id: "S06_ExigeMateriaisDisponibilizados", tipo: "parametro", input: "select",
      pergunta: "A execução exige que a Contratada disponibilize materiais, equipamentos, ferramentas ou utensílios próprios?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "S06_MateriaisDisponibilizados", tipo: "especificacao", input: "textarea",
      pergunta: "Liste os materiais, equipamentos, ferramentas e utensílios e as quantidades/qualidades exigidas",
      mostrarSe: { S06_ExigeMateriaisDisponibilizados: "Sim" } },
  ]},
  { id: "07", titulo: "DA GARANTIA CONTRATUAL", campos: [
    { id: "S07_ExigeGarantia", tipo: "parametro", input: "select",
      pergunta: "Será exigida garantia contratual?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não" },
    { id: "S07_BaseCalculoGarantia", tipo: "parametro", input: "select",
      pergunta: "A garantia de 5% (art. 439 RLC) incide sobre",
      opcoes: ["valor do Contrato", "valor anual do Contrato"],
      mostrarSe: { S07_ExigeGarantia: "Sim" } },
  ]},
  { id: "08", titulo: "DO ACOMPANHAMENTO E DA FISCALIZAÇÃO", campos: [
    { id: "S08_TipoFiscalizacao", tipo: "parametro", input: "select",
      pergunta: "A fiscalização será exercida por",
      opcoes: ["Fiscal Funcional", "Comissão de Fiscalização"],
      blocoTexto: true, padrao: "Fiscal Funcional",
      ajuda: "Comissão de Fiscalização reúne Fiscal Funcional, Técnico e Setorial (arts. 537 e 543 do RLC) e muda a redação da Seção 11 (Pagamento)." },
    { id: "S08_RotinasFiscalizacaoEspecificas", tipo: "especificacao", input: "textarea",
      pergunta: "Há rotinas de fiscalização próprias do objeto, além das genéricas já previstas nesta seção?",
      ajuda: "Pode deixar em branco se não houver rotina adicional além das já previstas no corpo fixo desta seção." },
  ]},
  { id: "09", titulo: "DAS OBRIGAÇÕES DA CONTRATANTE", campos: [] },
  { id: "10", titulo: "DAS OBRIGAÇÕES DA CONTRATADA", campos: [
    { id: "S10_ExigeEscritorioMunicipio", tipo: "parametro", input: "select",
      pergunta: "Será exigida declaração de instalação de escritório em município(s) específico(s)?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "S10_MunicipiosEscritorio", tipo: "parametro", input: "texto",
      pergunta: "Município(s) em que o escritório deverá ser instalado",
      mostrarSe: { S10_ExigeEscritorioMunicipio: "Sim" } },
    { id: "S10_PrazoProgramaIntegridadeMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo para comprovação do Programa de Integridade (meses)",
      ajuda: "Exigido quando o valor anual estimado supera R$ 1.500.000,00 — Anexo XXXII do RLC. Se não se aplicar, registre 0." },
    { id: "S10_NaturezaIntelectual", tipo: "parametro", input: "select",
      pergunta: "O serviço é de natureza intelectual, estratégico, ou envolve desenvolvimento de produto/projeto?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Se Sim, o documento inclui cessão de direitos autorais/patrimoniais e transição de conhecimento e tecnologia (art. 17 e art. 525 do RLC)." },
  ]},
  { id: "11", titulo: "DO PAGAMENTO", campos: [] },
  { id: "12", titulo: "DO REAJUSTE", campos: [
    { id: "S12_Reajustavel", tipo: "parametro", input: "select",
      pergunta: "Os preços serão reajustáveis após um ano?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não" },
    { id: "S12_IndiceReajuste", tipo: "parametro", input: "texto",
      pergunta: "Índice de referência para o reajuste",
      mostrarSe: { S12_Reajustavel: "Sim" } },
  ]},
  { id: "13", titulo: "DA SUBCONTRATAÇÃO", campos: [
    { id: "S13_PermiteSubcontratacao", tipo: "parametro", input: "select",
      pergunta: "Será admitida subcontratação parcial do objeto?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não" },
    { id: "S13_PercentualMinSubcontratacao", tipo: "parametro", input: "numero",
      pergunta: "Percentual mínimo de subcontratação (%)",
      mostrarSe: { S13_PermiteSubcontratacao: "Sim" } },
    { id: "S13_PercentualMaxSubcontratacao", tipo: "parametro", input: "numero",
      pergunta: "Percentual máximo de subcontratação (%)",
      mostrarSe: { S13_PermiteSubcontratacao: "Sim" } },
    { id: "S13_DetalheSubcontratacao", tipo: "especificacao", input: "textarea",
      pergunta: "Detalhe os percentuais, condições e parâmetros da subcontratação parcial",
      mostrarSe: { S13_PermiteSubcontratacao: "Sim" } },
  ]},
  { id: "14", titulo: "DA ALTERAÇÃO SUBJETIVA", campos: [] },
  { id: "15", titulo: "DAS SANÇÕES ADMINISTRATIVAS", campos: [
    { id: "S15_MultaNaoRegularizacaoHabilitacao", tipo: "parametro", input: "numero",
      pergunta: "Multa por não regularização da documentação de habilitação (% sobre o valor estimado)" },
    { id: "S15_MultaCondutasArt576a580", tipo: "parametro", input: "numero",
      pergunta: "Multa pelas condutas dos art. 576 a 580 do RLC (% sobre o valor estimado)" },
    { id: "S15_MultaRecusaAssinarContrato", tipo: "parametro", input: "numero",
      pergunta: "Multa pela recusa em assinar o contrato (% sobre o valor homologado)" },
    { id: "S15_MultaCompensatoriaTotal", tipo: "parametro", input: "numero",
      pergunta: "Multa compensatória por inexecução total do contrato (%)" },
    { id: "S15_MultaRescisoria", tipo: "parametro", input: "numero",
      pergunta: "Multa rescisória por rescisão unilateral (% sobre o valor total do contrato)" },
    { id: "S15_BaseCalculoMulta", tipo: "parametro", input: "select",
      pergunta: "As multas moratória e compensatória por inexecução parcial incidem sobre",
      opcoes: ["valor do Contrato", "valor anual do Contrato", "valor mensal do Contrato"],
      ajuda: "Usado nas alíneas e/f/g da Seção 15 e na Tabela 2 (grau → percentual)." },
    { id: "S15_InfracoesExtras", tipo: "parametro", input: "tabela-linhas",
      colunas: [
        { id: "descricao", rotulo: "Descrição da conduta", input: "texto" },
        { id: "grau", rotulo: "Grau", input: "select", opcoes: ["1", "2", "3", "4", "5"] },
      ],
      pergunta: "Infrações adicionais à Tabela 1 (além das 9 já previstas na minuta)" },
  ]},
  { id: "16", titulo: "DA RESCISÃO CONTRATUAL", campos: [] },
  { id: "17", titulo: "DA VEDAÇÃO AO NEPOTISMO", campos: [] },
  { id: "18", titulo: "DA SUSTENTABILIDADE AMBIENTAL", campos: [
    { id: "S18_CriteriosEspecificos", tipo: "especificacao", input: "textarea",
      pergunta: "Há critérios de sustentabilidade específicos para este objeto, além dos genéricos do art. 10 do RLC?",
      ajuda: "Pode deixar em branco se não houver critério adicional além dos já previstos no corpo fixo desta seção." },
  ]},
  { id: "19", titulo: "DA MATRIZ DE RISCOS E RESPONSABILIDADES", campos: [] },
  { id: "20", titulo: "DAS DISPOSIÇÕES FINAIS", campos: [
    { id: "S20_ElabDemandanteNome", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Demandante: nome" },
    { id: "S20_ElabDemandanteSetor", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Demandante: setor/gerência" },
    { id: "S20_ElabDemandanteCargo", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Demandante: cargo" },
    { id: "S20_ElabTecnicaNome", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Técnica: nome" },
    { id: "S20_ElabTecnicaSetor", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Técnica: setor/gerência" },
    { id: "S20_ElabTecnicaCargo", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Técnica: cargo" },
    { id: "S20_AprovDemandanteNome", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Demandante: nome" },
    { id: "S20_AprovDemandanteSetor", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Demandante: setor/gerência" },
    { id: "S20_AprovTecnicaNome", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Técnica: nome" },
    { id: "S20_AprovTecnicaSetor", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Técnica: setor/gerência" },
    { id: "S20_SuperDemandanteNome", tipo: "parametro", input: "texto", pergunta: "Superintendente/Gerente — Área Demandante: nome" },
    { id: "S20_SuperTecnicaNome", tipo: "parametro", input: "texto", pergunta: "Superintendente/Gerente — Área Técnica: nome" },
  ]},
];
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/dados-tr-servicos.test.mjs`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add js/dados-tr-servicos.js tests/dados-tr-servicos.test.mjs
git commit -m "Add TR-Serviços data model (20 sections, transcribed from SEI 55379668)"
```

---

### Task 5: `templates/build_template_tr_servicos.py` e `Template_TR_Servicos.docx`

**Files:**
- Create: `templates/build_template_tr_servicos.py`
- Create (gerado pelo script): `templates/Template_TR_Servicos.docx`
- Test: `tests/gerador-docx.test.mjs` (novos casos end-to-end, ver Task 7)

**Interfaces:**
- Consumes: os IDs de campo definidos em `js/dados-tr-servicos.js` (Task 4)
  — todo bookmark do template precisa bater com um `campo.id` exato, mais os
  bookmarks sintéticos de bloco (`_BLOCO_<VARIANTE>`), os de duplicação
  (`_E`, `_F`, `_G`, `_TABELA2`, `_74`, ver Step 1) e o `_LINHA_MODELO` da
  Tabela 1 (Task 3).
- Produces: `templates/Template_TR_Servicos.docx`, o arquivo binário
  consumido por `js/app.js` (Task 6) via `fetch`.

- [ ] **Step 1: Escrever `templates/build_template_tr_servicos.py`**

Reaproveita a mesma base de `build_template_tr_bens.py` (helpers
`add_heading`, `add_paragraph_with_bookmarks`, `_wrap_paragraph_in_bookmark`)
e adiciona duas construções novas: bloco de variante multi-parágrafo e
linha-modelo de tabela.

```python
"""Gera Template_TR_Servicos.docx a partir do texto oficial da minuta
SEI 55379668, com bookmarks nos pontos de preenchimento identificados em
js/dados-tr-servicos.js. Rodar novamente sempre que o texto oficial mudar."""
import re
import unicodedata
import docx
from docx.shared import Pt

doc = docx.Document()
doc.styles["Normal"].font.name = "Arial"
doc.styles["Normal"].font.size = Pt(11)

_bookmark_id = [0]


def add_heading(texto):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.bold = True
    run.font.size = Pt(12)
    return p


def add_paragraph_with_bookmarks(partes):
    p = doc.add_paragraph()
    for parte in partes:
        if isinstance(parte, tuple):
            nome, texto_inicial = parte
            _bookmark_id[0] += 1
            bid = str(_bookmark_id[0])
            start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
            start.set(docx.oxml.ns.qn("w:id"), bid)
            start.set(docx.oxml.ns.qn("w:name"), nome)
            p._p.append(start)
            p.add_run(texto_inicial)
            end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
            end.set(docx.oxml.ns.qn("w:id"), bid)
            p._p.append(end)
        else:
            p.add_run(parte)
    return p


def normalizar_variante(valor):
    sem_acento = unicodedata.normalize("NFKD", valor).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9]+", "_", sem_acento).upper()


def add_bloco_variante(nome_base, variantes):
    """variantes: dict {valor_opcao: [paragrafo1, paragrafo2, ...]}. Cada
    opção vira um bookmark que abrange do primeiro ao último parágrafo dela
    (suporta bloco de 1 parágrafo — caso binário Sim/Não — e de vários, como
    a bifurcação da Seção 11)."""
    for opcao, paragrafos in variantes.items():
        nome_bookmark = f"{nome_base}_BLOCO_{normalizar_variante(opcao)}"
        _bookmark_id[0] += 1
        bid = str(_bookmark_id[0])
        primeiro = None
        ultimo = None
        for texto in paragrafos:
            p = doc.add_paragraph()
            p.add_run(texto)
            if primeiro is None:
                primeiro = p
            ultimo = p
        start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
        start.set(docx.oxml.ns.qn("w:id"), bid)
        start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
        primeiro._p.insert(0, start)
        end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
        end.set(docx.oxml.ns.qn("w:id"), bid)
        ultimo._p.append(end)


def add_linha_modelo_tabela(tabela, nome_bookmark, textos_por_coluna):
    linha = tabela.add_row()
    for celula, texto in zip(linha.cells, textos_por_coluna):
        celula.text = texto
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    paragrafo = linha.cells[0].paragraphs[0]
    start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start.set(docx.oxml.ns.qn("w:id"), bid)
    start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
    paragrafo._p.insert(0, start)
    end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end.set(docx.oxml.ns.qn("w:id"), bid)
    paragrafo._p.append(end)
    return linha


# Capa
doc.add_paragraph("COMPANHIA NACIONAL DE ABASTECIMENTO")
titulo = doc.add_paragraph()
titulo_run = titulo.add_run("TERMO DE REFERÊNCIA")
titulo_run.bold = True

# 1. DO OBJETO
add_heading("1. DO OBJETO")
add_paragraph_with_bookmarks([
    "1.1. O presente Termo de Referência tem por objeto a contratação de ",
    ("S01_DescricaoObjeto", "(descrição do serviço a ser contratado)"),
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento.",
])
add_paragraph_with_bookmarks([
    "1.2. Os serviços serão executados indiretamente no regime de ",
    ("S01_RegimeExecucao", "(regime de execução)"),
    ", conforme o inciso IV, art. 208 do RLC.",
])
add_paragraph_with_bookmarks([
    "1.3. O custo total estimado para a prestação dos serviços deste Termo de Referência é de R$ ",
    ("S01_ValorEstimado", "________"),
    ". O critério utilizado para a formação da estimativa foi de ",
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
])
add_paragraph_with_bookmarks([
    "1.4. O critério de julgamento das propostas será o de ",
    ("S01_CriterioJulgamento", "[menor preço ou maior desconto]"),
    " ofertado. Para fins de julgamento das propostas o valor estimado/valor de referência será ",
    ("S01_SigiloValor", "(público ou sigiloso)"),
    ". O modo de disputa adotado no certame será ",
    ("S01_ModoDisputa", "(aberto ou aberto e fechado)"),
    ".",
])
add_paragraph_with_bookmarks([
    "1.5. Os recursos orçamentários decorrentes da contratação do objeto deste Termo de Referência "
    "estão consignados no Orçamento da Conab para o ano 20XX, conforme segue: ND: (",
    ("S01_ND", "XXXX"), "), PTRES: (", ("S01_PTRES", "XXXX"), "), FONTE: (", ("S01_Fonte", "XXXX"), ").",
])

# 2. DA JUSTIFICATIVA
add_heading("2. DA JUSTIFICATIVA")
add_paragraph_with_bookmarks(["2.1. ", ("S02_Motivacao", "(motivação da contratação)")])
doc.add_paragraph(
    "2.2. O serviço em apreço será contratado na modalidade Pregão Eletrônico, por ser "
    "caracterizado como serviço comum, na forma do Inciso XIII do art. 3º do Regulamento de "
    "Licitações e Contratos da Conab - RLC.")
add_bloco_variante("S02_AplicaExcecaoMEEPP", {
    "Sim": [
        "2.3. Em que pese o valor estimado da contratação ser inferior ao limite de "
        "R$ 80.000,00 (oitenta mil reais), o certame não será realizado com exclusividade de "
        "participação para microempresas e empresas de pequeno porte, na forma da Lei "
        "Complementar 123/2006 e do Decreto 8.538/2015, em razão da contratação em apreço se "
        "enquadrar na exceção prevista no art. 49 da Lei nº 123/2006, isto é: ",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["", ("S02_MotivoExcecaoMEEPP", "")])

# 3. DO OBJETIVO
add_heading("3. DO OBJETIVO")
add_paragraph_with_bookmarks(["3.1. ", ("S03_Objetivo", "(objetivo da contratação)")])

# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_bloco_variante("S04_TipoVigencia", {
    "Até 1 ano, prorrogável": [
        "4.1. O prazo de vigência do Contrato será de ____ meses (duração do Contrato, "
        "conforme arts. 461 a 463 do RLC), contados da data de sua assinatura, prorrogável na "
        "forma dos arts. 497 e 498 do RLC.",
    ],
    "Plurianual (até 5 anos)": [
        "4.1. O prazo de vigência do Contrato será de ____ anos (duração do Contrato, "
        "conforme arts. 461 a 462 do RLC), contados da data de sua assinatura, podendo ser "
        "prorrogado até o limite de 05 (cinco) anos, desde que observados os requisitos "
        "previstos no art. 488 do RLC.",
    ],
})
add_paragraph_with_bookmarks(["Prazo de vigência (meses): ", ("S04_PrazoVigenciaMeses", "____")])
add_paragraph_with_bookmarks(["Prazo de vigência (anos): ", ("S04_PrazoVigenciaAnos", "____")])

# 5. DA SELEÇÃO DO FORNECEDOR
add_heading("5. DA SELEÇÃO DO FORNECEDOR")
doc.add_paragraph("5.1. DA QUALIFICAÇÃO TÉCNICA")
add_paragraph_with_bookmarks([
    "5.1.1. a) Comprovação de aptidão para a prestação dos serviços em características, "
    "quantidades e prazos compatíveis com o objeto desta licitação, por período não inferior a ",
    ("S05_PrazoExperienciaMinimaAnos", "____"), " anos, mediante atestados fornecidos por "
    "pessoas jurídicas de direito público ou privado.",
])
add_paragraph_with_bookmarks([
    "a.1) Para fins da compatibilidade mencionada na alínea \"a\": ",
    ("S05_QuantitativoServicoPrincipal", "(quantitativo mínimo e serviço(s) principal(is))"),
])
doc.add_paragraph("5.1.2. Declaração de instalação de escritório:")
add_bloco_variante("S05_ExigeInstalacaoEscritorio", {
    "Sim": [
        "Para a realização do objeto da licitação, a Contratada deverá entregar declaração de "
        "que instalará escritório na cidade em que está localizada a Matriz ou a Sureg da "
        "Conab, a ser comprovado no prazo máximo de 60 (sessenta) dias contados a partir da "
        "vigência do contrato.",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["Cidade do escritório: ", ("S05_MunicipioInstalacaoEscritorio", "")])
doc.add_paragraph("5.2. DA VISTORIA")
add_bloco_variante("S05_TipoVistoria", {
    "Obrigatória": [
        "5.2.1. Para o correto dimensionamento e elaboração de sua proposta, o licitante "
        "deverá realizar vistoria nas instalações do local de execução dos serviços, "
        "acompanhado por empregado designado para esse fim, devendo o agendamento ser "
        "efetuado previamente, podendo sua realização ser comprovada por declaração assinada "
        "pelo licitante de que realizou a vistoria e tem pleno conhecimento das condições e "
        "peculiaridades do local, na forma do Anexo III deste TR.",
    ],
    "Facultativa": [
        "5.2.1. Para o correto dimensionamento e elaboração de sua proposta, o licitante "
        "poderá realizar vistoria nas instalações do local de execução dos serviços, "
        "acompanhado por empregado designado para esse fim, devendo o agendamento ser "
        "efetuado previamente, podendo sua realização ser comprovada por declaração assinada "
        "pelo licitante de que realizou a vistoria e tem pleno conhecimento das condições e "
        "peculiaridades do local, na forma do Anexo III deste TR.",
        "O licitante que optar por não realizar a vistoria deverá apresentar declaração de que "
        "tem pleno conhecimento das condições e peculiaridades do local, assumindo total "
        "responsabilidade por sua opção, na forma do Anexo III deste TR.",
    ],
})
add_paragraph_with_bookmarks([
    "5.2.1.1. O agendamento da vistoria deverá ser feito de segunda a sexta-feira, das ",
    ("S05_HorarioVistoriaInicio", "____"), " às ", ("S05_HorarioVistoriaFim", "____"),
    " horas, pelo telefone (", ("S05_TelefoneAgendamentoVistoria", "____"), ").",
])
doc.add_paragraph(
    "5.2.2. O prazo para vistoria iniciar-se-á no dia útil seguinte ao da publicação do "
    "Edital, estendendo-se até o dia útil anterior à data prevista para a abertura da sessão "
    "pública.")

# 6. DA DESCRIÇÃO DOS SERVIÇOS
add_heading("6. DA DESCRIÇÃO DOS SERVIÇOS")
doc.add_paragraph(
    "6.1. A prestação dos serviços não gera vínculo empregatício entre os empregados da "
    "Contratada e a Conab, vedando-se qualquer relação que caracterize pessoalidade e "
    "subordinação direta.")
doc.add_paragraph("6.2. DO DETALHAMENTO DOS SERVIÇOS")
add_paragraph_with_bookmarks(["6.2.1. ", ("S06_DetalhamentoServicos", "")])
doc.add_paragraph("6.3. DO LOCAL E PRAZO DE EXECUÇÃO")
add_paragraph_with_bookmarks([
    "6.3.1. Os serviços objeto deste Termo de Referência deverão ser realizados em ",
    ("S06_LocalExecucao", "(local)"), ", de segunda a sexta-feira, no horário de ",
    ("S06_HorarioExecucao", "(horário)"), ".",
])
add_paragraph_with_bookmarks([
    "6.3.2. O prazo de execução dos serviços será de ", ("S06_PrazoExecucaoDias", "____"),
    " dias, contados a partir da assinatura do Contrato.",
])
doc.add_paragraph("6.4. DA ROTINA DE EXECUÇÃO DOS SERVIÇOS")
add_paragraph_with_bookmarks(["6.4.1. ", ("S06_RotinaExecucao", "")])
doc.add_paragraph("6.5. DA GARANTIA DOS SERVIÇOS")
add_bloco_variante("S06_ExigeGarantiaServicos", {
    "Sim": ["6.5.1. Os serviços deverão ter um período de garantia de, no mínimo, ____ meses, "
            "contados do recebimento definitivo pela Contratante."],
    "Não": ["6.5.1. Não será exigida garantia dos serviços."],
})
add_paragraph_with_bookmarks(["Período de garantia (meses): ", ("S06_GarantiaServicosMeses", "____")])
add_paragraph_with_bookmarks([
    "6.5.2. Durante o período de garantia, a Contratada deverá refazer o serviço, no prazo "
    "máximo de ", ("S06_PrazoRefazimentoDias", "____"), " dias, após ser comunicada.",
])
doc.add_paragraph("6.6. DOS MATERIAIS A SEREM DISPONIBILIZADOS")
add_bloco_variante("S06_ExigeMateriaisDisponibilizados", {
    "Sim": ["6.6.1. Para a perfeita execução dos serviços, a Contratada deverá disponibilizar "
            "os materiais, equipamentos, ferramentas e utensílios necessários, promovendo sua "
            "substituição quando necessário."],
    "Não": [""],
})
add_paragraph_with_bookmarks(["", ("S06_MateriaisDisponibilizados", "")])
doc.add_paragraph("6.7. DO INSTRUMENTO DE MEDIÇÃO DE RESULTADOS")
doc.add_paragraph(
    "6.7.1. Durante a vigência do Contrato, a Conab, para avaliar a qualidade na execução dos "
    "serviços, adotará o Instrumento de Medição de Resultados (IMR), constante no Anexo V "
    "deste Termo de Referência.")
doc.add_paragraph("6.8. DO RECEBIMENTO DOS SERVIÇOS")
doc.add_paragraph(
    "6.8.1. O recebimento provisório será realizado pela fiscalização; o recebimento "
    "definitivo, pelo empregado ou Comissão designada, nos termos do art. 526 e 527 do RLC.")

# 7. DA GARANTIA CONTRATUAL
add_heading("7. DA GARANTIA CONTRATUAL")
add_bloco_variante("S07_ExigeGarantia", {
    "Sim": [
        "7.1. A Contratada deverá apresentar, no prazo de 10 (dez) dias corridos após a "
        "assinatura do Contrato, prorrogáveis por igual período, comprovante de prestação de "
        "garantia correspondente a 5% (cinco por cento), conforme parâmetros do art. 439 do "
        "RLC, do ", ("S07_BaseCalculoGarantia", "(valor do Contrato ou valor anual do Contrato)"),
    ],
    "Não": ["7.1. Não será exigida garantia contratual."],
})
add_paragraph_with_bookmarks([
    "7.4. A inobservância do prazo fixado para apresentação da garantia acarretará a "
    "aplicação de multa de 0,08% (oito centésimos por cento) do ",
    ("S07_BaseCalculoGarantia_74", "(valor do Contrato ou valor anual do Contrato)"),
    " por dia de atraso, até o máximo de 2% (dois por cento).",
])

# 8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO
add_heading("8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO")
doc.add_paragraph(
    "8.1. Para a execução do ajuste, será adotado o método de trabalho baseado no conceito de "
    "delegação de responsabilidade. Nos termos dos arts. 543 e 544 do RLC, será designado "
    "Fiscal, seu substituto, ou Comissão de Fiscalização para acompanhar e fiscalizar a "
    "execução dos serviços.")
add_paragraph_with_bookmarks(["8.17. Rotinas de fiscalização específicas: ", ("S08_RotinasFiscalizacaoEspecificas", "")])

# 9. DAS OBRIGAÇÕES DA CONTRATANTE (texto fixo)
add_heading("9. DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "9.1. São obrigações da Contratante: exigir o cumprimento das obrigações da Contratada; "
    "receber o objeto no prazo e condições estabelecidas; verificar a conformidade dos "
    "serviços recebidos provisoriamente; comunicar à Contratada sobre imperfeições, falhas ou "
    "irregularidades; rejeitar serviço em desacordo; acompanhar e fiscalizar o cumprimento das "
    "obrigações da Contratada; e efetuar o pagamento no prazo e forma estabelecidos.")

# 10. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("10. DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph(
    "10.1. A Contratada deve cumprir todas as obrigações constantes neste Termo, seus anexos "
    "e sua proposta, assumindo como exclusivamente seus os riscos e as despesas decorrentes da "
    "boa e perfeita execução do objeto, incluindo: executar os serviços com a alocação dos "
    "empregados necessários; responsabilizar-se pelos vícios e danos decorrentes do objeto "
    "(CDC); substituir, reparar ou corrigir às suas expensas os serviços com vícios; comunicar "
    "atrasos no prazo máximo de 24 horas; utilizar empregados habilitados, uniformizados e "
    "identificados; apresentar relação nominal dos empregados; responsabilizar-se pelas "
    "obrigações trabalhistas, sociais, previdenciárias e tributárias; substituir empregados a "
    "pedido da fiscalização; instruir seus empregados quanto às normas internas e às "
    "atividades contratadas; relatar irregularidades; não permitir trabalho de menores; "
    "guardar sigilo; arcar com equívocos de dimensionamento de sua proposta; manter as "
    "condições de habilitação; prestar esclarecimentos; indicar preposto; e deter instalações, "
    "aparelhamento e pessoal técnico adequados.")
doc.add_paragraph("Declaração de instalação de escritório em município(s) específico(s):")
add_bloco_variante("S10_ExigeEscritorioMunicipio", {
    "Sim": [
        "A Contratada deverá entregar declaração de que instalará escritório no(s) "
        "município(s) indicado(s), a ser comprovado no prazo máximo de 60 (sessenta) dias "
        "contados a partir da vigência do Contrato, dispondo de capacidade operacional para "
        "receber e solucionar qualquer demanda da Contratante.",
    ],
    "Não": [""],
})
add_paragraph_with_bookmarks(["Município(s): ", ("S10_MunicipiosEscritorio", "")])
add_paragraph_with_bookmarks([
    "s) comprovar a implementação do Programa de Integridade, nos moldes do art. 603 e Anexo "
    "XXXII do RLC, no prazo de até ", ("S10_PrazoProgramaIntegridadeMeses", "____"),
    " meses após a assinatura do contrato.",
])
add_bloco_variante("S10_NaturezaIntelectual", {
    "Sim": [
        "u) ceder os direitos patrimoniais relativos ao projeto ou serviço técnico "
        "especializado, para que a Conab possa utilizá-lo, de acordo com o art. 525 do RLC. "
        "Quando o projeto referir-se a obra imaterial de caráter tecnológico, insuscetível de "
        "privilégio, a cessão dos direitos incluirá o fornecimento de todos os dados, "
        "documentos e elementos de informação pertinentes à tecnologia de concepção, "
        "desenvolvimento, fixação em suporte físico e aplicação da obra;",
        "v) assegurar à Contratante o direito de propriedade intelectual dos produtos "
        "desenvolvidos, inclusive sobre adequações e atualizações futuras, de forma "
        "permanente, e os direitos autorais da solução, do projeto, de suas especificações "
        "técnicas e da documentação produzida, inclusive os produzidos por terceiros "
        "subcontratados, ficando proibida sua utilização sem autorização expressa da "
        "Contratante, sob pena de multa, sem prejuízo das sanções civis e penais cabíveis;",
        "w) realizar a transição contratual com transferência de conhecimento, tecnologia e "
        "técnicas empregadas, sem perda de informações, podendo exigir, inclusive, a "
        "capacitação dos técnicos da Contratante ou da nova empresa que continuará a execução "
        "dos serviços, nos termos do art. 17 do RLC.",
    ],
    "Não": [""],
})

# 11. DO PAGAMENTO
add_heading("11. DO PAGAMENTO")
doc.add_paragraph(
    "11.1. O pagamento será efetuado pela Conab no prazo de até 10 (dez) dias corridos, "
    "contados do recebimento da Nota Fiscal/Fatura.")
doc.add_paragraph(
    "11.2. A emissão da Nota Fiscal/Fatura será precedida do recebimento provisório e "
    "definitivo do serviço, nos seguintes termos: a) no prazo de até 05 (cinco) dias corridos "
    "do adimplemento da parcela, a Contratada deverá entregar toda a documentação "
    "comprobatória do cumprimento da obrigação contratual;")
add_bloco_variante("S08_TipoFiscalizacao", {
    "Fiscal Funcional": [
        "b) no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada, o Fiscal Funcional deverá realizar a análise de toda a documentação "
        "apresentada, avaliar a execução por meio do Instrumento de Medição de Resultado (IMR) "
        "e emitir Termo de Recebimento Provisório, com detalhamento da execução contratual, "
        "encaminhando-o ao empregado ou comissão designada para o recebimento definitivo.",
    ],
    "Comissão de Fiscalização": [
        "b) no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada:",
        "b.1) o Fiscal Funcional deverá realizar a análise de toda a documentação apresentada "
        "e emitir relatório, com detalhamento da execução contratual;",
        "b.2) o Fiscal Técnico deverá realizar a análise de toda a documentação apresentada e "
        "emitir relatório, com detalhamento da execução contratual;",
        "b.3) o Fiscal Setorial deverá realizar a análise de toda a documentação apresentada e "
        "emitir relatório, com detalhamento da execução contratual;",
        "c) o Fiscal Funcional, no prazo de até 03 (três) dias úteis a partir do recebimento do "
        "último relatório, avaliará a execução por meio do Instrumento de Medição de Resultado "
        "(IMR) e emitirá Termo de Recebimento Provisório com base nos relatórios emitidos pelos "
        "fiscais, encaminhando-o, caso não haja irregularidades, ao empregado ou comissão "
        "designada para o recebimento definitivo.",
    ],
})
doc.add_paragraph(
    "d) constatadas impropriedades ou irregularidades que impeçam a liquidação e o pagamento "
    "da despesa, estas serão registradas no Termo de Recebimento Provisório, com prazo de até "
    "05 (cinco) dias úteis para regularização; e) sanadas as impropriedades, o Fiscal Funcional "
    "ou a Comissão de Fiscalização, em 05 (cinco) dias corridos, elaborará relatório detalhado "
    "e encaminhará o Termo de Recebimento Provisório ao empregado ou comissão designada para o "
    "recebimento definitivo; f) no prazo de até 05 (cinco) dias corridos, o empregado ou "
    "Comissão designada providenciará o recebimento definitivo, emitindo Termo Circunstanciado "
    "e comunicando a Contratada para emissão da Nota Fiscal ou Fatura, com o valor exato "
    "dimensionado pela fiscalização com base no IMR.")
doc.add_paragraph(
    "11.5. Antes de cada pagamento será realizada consulta ao SICAF; havendo resultado "
    "desfavorável, será concedido prazo de 05 (cinco) dias úteis para regularização ou defesa, "
    "prorrogável uma vez por igual período.")

# 12. DO REAJUSTE
add_heading("12. DO REAJUSTE")
add_bloco_variante("S12_Reajustavel", {
    "Sim": ["12.1. O preço consignado no Contrato será corrigido anualmente, observado o "
            "interregno mínimo de um ano, contado da data limite para apresentação da "
            "proposta, pela variação do índice de referência."],
    "Não": ["12.1. O preço é fixo e irreajustável."],
})
add_paragraph_with_bookmarks(["Índice de referência: ", ("S12_IndiceReajuste", "")])

# 13. DA SUBCONTRATAÇÃO
add_heading("13. DA SUBCONTRATAÇÃO")
add_bloco_variante("S13_PermiteSubcontratacao", {
    "Sim": ["13.1. É permitida a subcontratação parcial do objeto entre os limites mínimo e "
            "máximo, respectivamente, do valor total do Contrato."],
    "Não": ["13.1. Não será admitida a subcontratação do objeto licitatório."],
})
add_paragraph_with_bookmarks([
    "Limites: entre ", ("S13_PercentualMinSubcontratacao", "____"), "% e ",
    ("S13_PercentualMaxSubcontratacao", "____"), "% do valor total do Contrato.",
])
add_paragraph_with_bookmarks(["Detalhamento: ", ("S13_DetalheSubcontratacao", "")])
doc.add_paragraph("13.2. A subcontratação deverá observar o disposto no artigo 524 do RLC.")

# 14. DA ALTERAÇÃO SUBJETIVA (texto fixo)
add_heading("14. DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "14.1. É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa "
    "jurídica, desde que observados os requisitos de habilitação, mantidas as demais "
    "cláusulas e condições do Contrato, não haja prejuízo à execução do objeto pactuado e "
    "haja a anuência expressa da Conab.")

# 15. DAS SANÇÕES ADMINISTRATIVAS
add_heading("15. DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "15.1. A contratada, em caso de inadimplemento de suas obrigações, garantido o "
    "contraditório e ampla defesa, ficará sujeita às seguintes sanções previstas no RLC e na "
    "Lei nº 13.303, de 2016: advertência; multa moratória; multa compensatória; multa "
    "rescisória, para os casos de rescisão unilateral por descumprimento contratual; e "
    "suspensão do direito de participar de licitação e impedimento de contratar com a Conab, "
    "por até 02 (dois) anos.")
add_paragraph_with_bookmarks([
    "a) multa por não regularização da documentação de habilitação: ",
    ("S15_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado para a licitação.",
])
add_paragraph_with_bookmarks([
    "b) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("S15_MultaCondutasArt576a580", "____"), "% sobre o valor estimado para a licitação.",
])
add_paragraph_with_bookmarks([
    "c) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("S15_MultaRecusaAssinarContrato", "____"), "% sobre o valor homologado para a licitação.",
])
doc.add_paragraph("d) multa moratória por atraso injustificado na entrega da garantia contratual, conforme item 7.4.")
add_paragraph_with_bookmarks([
    "e) multa moratória de 0,2% (dois décimos por cento) sobre o ",
    ("S15_BaseCalculoMulta_E", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, até o limite de 15 (quinze) dias.",
])
add_paragraph_with_bookmarks([
    "f) multa moratória de 0,3% (três décimos por cento) sobre o ",
    ("S15_BaseCalculoMulta_F", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, por período superior ao previsto na alínea "
    "anterior, até o limite de 15 (quinze) dias. Esgotado o prazo, poderá ocorrer a não "
    "aceitação do objeto, sem prejuízo da rescisão unilateral da avença.",
])
add_paragraph_with_bookmarks([
    "g) multa compensatória no percentual de 10% (dez por cento) sobre o ",
    ("S15_BaseCalculoMulta_G", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", no caso de inexecução parcial do Contrato.",
])
add_paragraph_with_bookmarks([
    "h) multa compensatória de ", ("S15_MultaCompensatoriaTotal", "____"),
    "% sobre o valor total do Contrato, no caso de inexecução total do Contrato.",
])
add_paragraph_with_bookmarks([
    "i) multa rescisória de ", ("S15_MultaRescisoria", "____"),
    "% sobre o valor total do Contrato, no caso de rescisão contratual unilateral.",
])
add_paragraph_with_bookmarks([
    "j) multa de 0,2% a 3,2% por dia sobre o ",
    ("S15_BaseCalculoMulta_TABELA2", "(valor do Contrato, valor anual ou valor mensal)"),
    ", conforme grau da infração cometida (Tabelas 1 e 2 abaixo).",
])

tabela1 = doc.add_table(rows=1, cols=2)
tabela1.style = "Table Grid"
tabela1.rows[0].cells[0].text = "TABELA 1 — DESCRIÇÃO DA INFRAÇÃO"
tabela1.rows[0].cells[1].text = "GRAU"
infracoes_fixas = [
    ("Permitir situação que crie a possibilidade de causar dano físico, lesão corporal ou "
     "consequências letais, por ocorrência", "5"),
    ("Suspender ou interromper, salvo motivo de força maior ou caso fortuito, os serviços "
     "contratuais, por dia e por unidade de atendimento", "4"),
    ("Servir-se de funcionário sem qualificação para executar os serviços contratados, por "
     "empregado e por dia", "3"),
    ("Recusar-se a executar serviço determinado pela fiscalização, por serviço e por dia", "2"),
    ("Deixar de cumprir determinação formal ou instrução complementar da fiscalização, por "
     "ocorrência", "2"),
    ("Deixar de substituir empregado alocado que não atenda às necessidades do serviço, por "
     "funcionário e por dia", "1"),
    ("Deixar de cumprir quaisquer itens do Edital e seus Anexos não previstos nesta tabela, "
     "após reincidência formalmente notificada pela fiscalização, por item e por ocorrência", "3"),
    ("Deixar de indicar e manter durante a execução do Contrato os prepostos previstos no "
     "Edital/Contrato, por dia", "1"),
    ("Deixar de providenciar treinamento para seus funcionários conforme previsto na relação "
     "de obrigações da Contratada, por funcionário", "1"),
]
for descricao, grau in infracoes_fixas:
    linha = tabela1.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
add_linha_modelo_tabela(tabela1, "S15_InfracoesExtras_LINHA_MODELO", ["((DESCRICAO))", "((GRAU))"])

tabela2 = doc.add_table(rows=1, cols=2)
tabela2.style = "Table Grid"
tabela2.rows[0].cells[0].text = "GRAU"
tabela2.rows[0].cells[1].text = "CORRESPONDÊNCIA (percentual por dia)"
for grau, percentual in [("1", "0,2%"), ("2", "0,4%"), ("3", "0,8%"), ("4", "1,6%"), ("5", "3,2%")]:
    linha = tabela2.add_row()
    linha.cells[0].text = grau
    linha.cells[1].text = percentual
doc.add_paragraph(
    "15.8.2. As multas moratória, compensatória e rescisória possuem fatos geradores "
    "distintos; se aplicadas duas multas sobre o mesmo fato gerador, configura bis in idem.")
doc.add_paragraph(
    "15.9. Da sanção de suspensão: cabe em razão de ação ou omissão capaz de causar, ou que "
    "tenha causado, prejuízo à Conab, suas instalações, pessoas, imagem, meio ambiente, ou em "
    "decorrência de determinação legal, aplicada por até 02 (dois) anos, nos termos dos arts. "
    "579 a 580 do RLC.")

# 16. DA RESCISÃO CONTRATUAL (texto fixo)
add_heading("16. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "16.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução "
    "parcial poderá ensejar a sua rescisão, com as consequências cabíveis, conforme disposto "
    "nos arts. 568 a 572 do RLC. A rescisão poderá ser por ato unilateral e escrito da Conab, "
    "amigável, por acordo entre as partes, ou judicial.")

# 17. DA VEDAÇÃO AO NEPOTISMO (texto fixo)
add_heading("17. DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "17.1. Conforme disposto no parágrafo único do art. 12 do RLC e no art. 7º do Decreto nº "
    "7.203, de 2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa "
    "física, ou de pessoa/empresa com relação de parentesco até o terceiro grau civil com "
    "autoridade do Ministério da Agricultura, dirigente ou empregado da Conab com atribuições "
    "na área de licitação e contratação.")

# 18. DA SUSTENTABILIDADE AMBIENTAL
add_heading("18. DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "18.1. Compete à Contratada, no que couber, atender os critérios de sustentabilidade "
    "ambiental previstos no art. 10 do RLC.")
doc.add_paragraph(
    "18.2. A Contratada se responsabiliza administrativamente, civilmente e penalmente por "
    "qualquer dano causado pelo seu serviço ao meio ambiente.")
add_paragraph_with_bookmarks(["18.3. ", ("S18_CriteriosEspecificos", "")])

# 19. DA MATRIZ DE RISCOS E RESPONSABILIDADES (texto fixo)
add_heading("19. DA MATRIZ DE RISCOS E RESPONSABILIDADES")
doc.add_paragraph(
    "19.1. A Matriz de Riscos é a cláusula contratual definidora dos riscos e das "
    "responsabilidades entre o Contratante e a Contratada, conforme hipóteses elencadas na "
    "Matriz de Riscos - Anexo I do Termo de Referência.")

# 20. DAS DISPOSIÇÕES FINAIS
add_heading("20. DAS DISPOSIÇÕES FINAIS")
doc.add_paragraph(
    "20.1. Integram este Termo de Referência os seguintes Anexos: a) Anexo I - Matriz de "
    "Riscos; b) Anexo II - Modelo de Apresentação de Proposta; c) Anexo III - Declaração de "
    "Vistoria Técnica; d) Anexo IV - Modelo de Declaração de Contratos Firmados com a "
    "Iniciativa Privada e a Administração Pública; e e) Anexo V - Modelo de Instrumento de "
    "Medição de Resultados (IMR).")
add_paragraph_with_bookmarks(["Elaborador — Área Demandante: ", ("S20_ElabDemandanteNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S20_ElabDemandanteSetor", ""), " — Cargo: ", ("S20_ElabDemandanteCargo", "")])
add_paragraph_with_bookmarks(["Elaborador — Área Técnica: ", ("S20_ElabTecnicaNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S20_ElabTecnicaSetor", ""), " — Cargo: ", ("S20_ElabTecnicaCargo", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Demandante: ", ("S20_AprovDemandanteNome", ""), " — Setor: ", ("S20_AprovDemandanteSetor", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Técnica: ", ("S20_AprovTecnicaNome", ""), " — Setor: ", ("S20_AprovTecnicaSetor", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Demandante: ", ("S20_SuperDemandanteNome", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Técnica: ", ("S20_SuperTecnicaNome", "")])

# ANEXO I - MATRIZ DE RISCOS
doc.add_page_break()
add_heading("ANEXO I DO TERMO DE REFERÊNCIA")
add_heading("MATRIZ DE RISCOS")
add_paragraph_with_bookmarks(["Objeto: ", ("S01_DescricaoObjeto_AnexoI", "")])
tabela_riscos = doc.add_table(rows=1, cols=8)
tabela_riscos.style = "Table Grid"
for celula, texto in zip(tabela_riscos.rows[0].cells, [
    "Item", "Fases", "Evento de Risco", "Causas", "Consequências",
    "Probabilidade (P)", "Impacto (I)", "Nível de Risco (P x I)",
]):
    celula.text = texto
for _ in range(6):
    tabela_riscos.add_row()
doc.add_paragraph("Escala de Probabilidade: 1 Muito Baixa, 2 Baixa, 3 Média, 4 Alta, 5 Muito Alta.")
doc.add_paragraph("Escala de Impacto: 1 Muito Baixo, 2 Baixo, 3 Médio, 4 Alto, 5 Muito Alto.")
doc.add_paragraph("Nível de Risco: 1 a 2 Baixo; 3 a 6 Médio; 8 a 12 Elevado; 15 a 25 Extremo.")
doc.add_paragraph("Elaborado por: NOME     Supervisionado por: NOME     Autorizado por: NOME")

# ANEXO II - MODELO DE APRESENTAÇÃO DE PROPOSTA
doc.add_page_break()
add_heading("ANEXO II DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE APRESENTAÇÃO DE PROPOSTA")
tabela_proposta = doc.add_table(rows=1, cols=3)
tabela_proposta.style = "Table Grid"
for celula, texto in zip(tabela_proposta.rows[0].cells, ["Item", "Especificações", "Valor Total"]):
    celula.text = texto
for _ in range(5):
    tabela_proposta.add_row()
doc.add_paragraph("VALOR TOTAL DOS ITENS: R$")

# ANEXO III - MODELO DE DECLARAÇÃO DE VISTORIA
doc.add_page_break()
add_heading("ANEXO III DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE DECLARAÇÃO DE VISTORIA")
doc.add_paragraph("À Companhia Nacional de Abastecimento – CONAB")
doc.add_paragraph("Ref.: Pregão Eletrônico – nº ___/20XX")
doc.add_paragraph(
    "Declaramos que, em atendimento ao previsto no Pregão Eletrônico nº ___/_____, "
    "eu, ______________________________________, portador(a) da CI/RG nº______________ e do "
    "CPF nº _____________________, responsável da empresa "
    "___________________________________________________, como seu representante legal para "
    "os fins da presente declaração, compareci perante a CONAB/Sureg/XX, situada à "
    "_____________________________________________________________________________________, "
    "e vistoriei os locais onde serão executados os serviços objeto da licitação em apreço, "
    "tomando plena ciência das condições e graus de dificuldade existentes.")
doc.add_paragraph("Cidade-UF, ___/___/20XX.")
doc.add_paragraph("Assinatura e carimbo (Responsável Técnico da Empresa)")
doc.add_paragraph("Visto: Responsável pela CONAB")

# ANEXO IV - MODELO DE DECLARAÇÃO DE CONTRATOS FIRMADOS
doc.add_page_break()
add_heading("ANEXO IV DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE DECLARAÇÃO DE CONTRATOS FIRMADOS COM A INICIATIVA PRIVADA E A ADMINISTRAÇÃO PÚBLICA")
doc.add_paragraph(
    "Declaro que a empresa ________________________, inscrita no CNPJ (MF) nº ________, "
    "inscrição estadual nº _____________, estabelecida em _____________, possui os seguintes "
    "contratos firmados com a iniciativa privada e a Administração Pública:")
tabela_contratos = doc.add_table(rows=1, cols=3)
tabela_contratos.style = "Table Grid"
for celula, texto in zip(tabela_contratos.rows[0].cells, ["Nome do Órgão/Empresa", "Vigência do Contrato", "Valor total do Contrato*"]):
    celula.text = texto
for _ in range(5):
    tabela_contratos.add_row()
doc.add_paragraph("Valor total dos Contratos: R$")
doc.add_paragraph("*Considera-se o valor remanescente do contrato, excluindo o já executado.")
doc.add_paragraph(
    "Fórmula exemplificativa, para fins de atendimento ao disposto no art. 131, § 6º, inciso "
    "IV, alíneas a e b, do RLC:")
doc.add_paragraph(
    "1. A Declaração de Compromissos Assumidos deve informar que 1/12 (um doze avos) dos "
    "contratos firmados pela licitante não é superior ao Patrimônio Líquido da licitante. "
    "Fórmula: (Valor do Patrimônio Líquido / Valor total dos contratos*) x 12 > 1. Esse "
    "resultado deverá ser superior a 1 (um).")
doc.add_paragraph(
    "2. Caso a diferença entre a receita bruta discriminada na DRE e a declaração apresentada "
    "seja maior que 10% (dez por cento), positivo ou negativo, em relação à receita bruta, o "
    "licitante deverá apresentar justificativas. Fórmula: [(Valor da Receita Bruta - Valor "
    "total dos Contratos) / Valor da Receita Bruta] x 100.")

# ANEXO V - MODELO DE INSTRUMENTO DE MEDIÇÃO DE RESULTADO (IMR)
doc.add_page_break()
add_heading("ANEXO V DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE INSTRUMENTO DE MEDIÇÃO DE RESULTADO (IMR)")
tabela_imr = doc.add_table(rows=1, cols=2)
tabela_imr.style = "Table Grid"
tabela_imr.rows[0].cells[0].text = "Item"
tabela_imr.rows[0].cells[1].text = "Descrição"
for item in ["Finalidade", "Meta a cumprir", "Instrumento de Medição", "Forma de acompanhamento",
             "Periodicidade", "Mecanismo de cálculo", "Início de vigência",
             "Faixas de ajuste no pagamento", "Sanções", "Observações"]:
    linha = tabela_imr.add_row()
    linha.cells[0].text = item
doc.add_paragraph("Exemplo de indicador — Nº 01: Prazo de atendimento de demandas (OS)")
tabela_imr_exemplo = doc.add_table(rows=1, cols=2)
tabela_imr_exemplo.style = "Table Grid"
tabela_imr_exemplo.rows[0].cells[0].text = "Item"
tabela_imr_exemplo.rows[0].cells[1].text = "Descrição"
exemplo_imr = [
    ("Finalidade", "Garantir um atendimento célere às demandas da Conab."),
    ("Meta a cumprir", "24h"),
    ("Forma de acompanhamento", "Sistema informatizado de solicitação de serviços — Ordem de Serviço (OS) eletrônica."),
    ("Mecanismo de cálculo", "Cada OS é verificada e valorada individualmente: nº de horas no atendimento / 24 = X."),
    ("Faixas de ajuste no pagamento",
     "X até 1: 100% do valor da OS. De 1 a 1,5: 90%. De 1,5 a 2: 80%."),
    ("Sanções",
     "20% das OS acima de 2: multa. 30% das OS acima de 2: multa + rescisão contratual."),
]
for item, descricao in exemplo_imr:
    linha = tabela_imr_exemplo.add_row()
    linha.cells[0].text = item
    linha.cells[1].text = descricao

doc.save("Template_TR_Servicos.docx")
print("Template_TR_Servicos.docx gerado com sucesso.")
```

- [ ] **Step 2: Rodar o script e conferir que o arquivo é gerado**

Run: `cd "templates" && python build_template_tr_servicos.py`
Expected: imprime `Template_TR_Servicos.docx gerado com sucesso.` e o arquivo
aparece em `templates/Template_TR_Servicos.docx`.

- [ ] **Step 3: Commit**

```bash
git add templates/build_template_tr_servicos.py templates/Template_TR_Servicos.docx
git commit -m "Add TR-Serviços docx template build script and generated template"
```

---

### Task 6: Generalizar `js/app.js` e `js/triagem.js` para múltiplos instrumentos

**Files:**
- Modify: `js/app.js` (todo o arquivo)
- Modify: `js/triagem.js`
- Test: manual (ver "Testes" no rodapé do plano) — este task é plumbing de
  UI/DOM, sem `jsdom` no projeto; a cobertura automática já está nas tasks
  anteriores (motor) e na Task 7 (geração ponta a ponta).

**Interfaces:**
- Consumes: `SECOES_TR_BENS` (existente), `SECOES_TR_SERVICOS` (Task 4).
- Produces: nenhuma interface nova consumida por outro módulo — `app.js` é
  o topo da árvore de import.

- [ ] **Step 1: Reescrever `js/triagem.js` com dois caminhos**

```js
export function renderizarTriagem(container, aoConfirmar) {
  container.innerHTML = `
    <h1>Triagem inicial</h1>
    <p>Este protótipo cobre o Termo de Referência para aquisição de bens
    (TR-Bens, art. 127 do RLC) e para prestação de serviços (TR-Serviços,
    art. 129 do RLC). As demais variantes de TR (Dispensa) e o Projeto
    Básico entram em planos seguintes.</p>
    <button id="btn-tr-bens" class="br-button primary" type="button">
      Avançar para o TR-Bens
    </button>
    <button id="btn-tr-servicos" class="br-button primary" type="button">
      Avançar para o TR-Serviços
    </button>
  `;
  container.querySelector("#btn-tr-bens")
    .addEventListener("click", () => aoConfirmar("TR-Bens"));
  container.querySelector("#btn-tr-servicos")
    .addEventListener("click", () => aoConfirmar("TR-Serviços"));
}
```

- [ ] **Step 2: Reescrever `js/app.js` parametrizado por instrumento**

```js
import { SECOES_TR_BENS } from "./dados-tr-bens.js";
import { SECOES_TR_SERVICOS } from "./dados-tr-servicos.js";
import { renderizarSumario, renderizarSecoes, atualizarProgresso } from "./render.js";
import { renderizarTriagem } from "./triagem.js";
import { renderizarConferencia } from "./conferencia.js";
import { serializarRascunho, desserializarRascunho } from "./persistencia.js";
import { controladoresCondicionais } from "./modelo.js";
import { gerarDocx } from "./gerador-docx.js";

const INSTRUMENTOS = {
  "TR-Bens": {
    secoes: SECOES_TR_BENS,
    titulo: "Termo de Referência — Aquisição de Bens",
    templateUrl: "templates/Template_TR_Bens.docx",
    chaveAutosave: "gerador-tr-pb:tr-bens:rascunho",
    nomeArquivoRascunho: "rascunho-tr-bens.json",
    nomeArquivoGerado: "TR-Bens-gerado.docx",
  },
  "TR-Serviços": {
    secoes: SECOES_TR_SERVICOS,
    titulo: "Termo de Referência — Prestação de Serviços",
    templateUrl: "templates/Template_TR_Servicos.docx",
    chaveAutosave: "gerador-tr-pb:tr-servicos:rascunho",
    nomeArquivoRascunho: "rascunho-tr-servicos.json",
    nomeArquivoGerado: "TR-Servicos-gerado.docx",
  },
};

let instrumentoAtual = "TR-Bens";
let respostas = {};

function config() {
  return INSTRUMENTOS[instrumentoAtual];
}

function carregarAutosaveSeExistir() {
  let salvo;
  try {
    salvo = localStorage.getItem(config().chaveAutosave);
  } catch (erro) {
    console.warn("Armazenamento local indisponível; autosave desativado nesta sessão.", erro);
    return;
  }
  if (!salvo) return;
  try {
    const { respostas: respostasSalvas } = desserializarRascunho(salvo);
    respostas = respostasSalvas;
  } catch (erro) {
    alert("O rascunho salvo automaticamente estava corrompido e foi descartado: " + erro.message);
    try { localStorage.removeItem(config().chaveAutosave); } catch { /* já sem acesso ao storage */ }
  }
}

function persistirAutosave() {
  try {
    localStorage.setItem(config().chaveAutosave, serializarRascunho(instrumentoAtual, respostas));
  } catch (erro) {
    console.warn("Não foi possível salvar o rascunho automaticamente.", erro);
  }
}

function montarTelaPainel() {
  const { secoes, titulo, templateUrl, nomeArquivoGerado } = config();
  const painel = document.getElementById("tela-painel");
  painel.innerHTML = `
    <div class="cartao-painel">
      <div class="cabecalho-painel">${titulo}</div>
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
  const camposControladores = controladoresCondicionais(secoes);
  function algumCampoDependeDe(campoId) {
    return camposControladores.includes(campoId);
  }

  function aoMudar(id, valor) {
    respostas[id] = valor;
    persistirAutosave();
    atualizarProgresso(secoes, respostas, document.getElementById("barra-progresso"));
    if (algumCampoDependeDe(id)) {
      rerenderizar();
    }
  }

  function rerenderizar() {
    renderizarSecoes(secoes, respostas, document.getElementById("secoes"), aoMudar);
  }

  renderizarSumario(secoes, document.getElementById("sumario"));
  rerenderizar();
  atualizarProgresso(secoes, respostas, document.getElementById("barra-progresso"));

  document.getElementById("btn-gerar").addEventListener("click", () => {
    const overlay = document.getElementById("conferencia-overlay");
    renderizarConferencia(secoes, respostas, overlay, async () => {
      overlay.hidden = true;
      try {
        const resposta = await fetch(templateUrl);
        if (!resposta.ok) {
          throw new Error(`Não foi possível carregar o modelo (${resposta.status}).`);
        }
        const templateBuffer = await resposta.arrayBuffer();
        const docxGerado = await gerarDocx(templateBuffer, respostas, secoes);
        const blob = new Blob([docxGerado], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = nomeArquivoGerado;
        a.click();
        URL.revokeObjectURL(url);
      } catch (erro) {
        alert("Não foi possível gerar o documento: " + erro.message);
      }
    });
  });
}

function iniciar() {
  renderizarTriagem(document.getElementById("tela-triagem"), (instrumentoEscolhido) => {
    instrumentoAtual = instrumentoEscolhido;
    respostas = {};
    carregarAutosaveSeExistir();
    document.getElementById("tela-triagem").hidden = true;
    document.getElementById("tela-painel").hidden = false;
    montarTelaPainel();
  });

  document.getElementById("btn-salvar-rascunho").addEventListener("click", () => {
    const json = serializarRascunho(instrumentoAtual, respostas);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = config().nomeArquivoRascunho;
    a.click();
    URL.revokeObjectURL(url);
  });

  document.getElementById("btn-carregar-rascunho").addEventListener("click", () => {
    document.getElementById("input-carregar-rascunho").click();
  });

  document.getElementById("input-carregar-rascunho").addEventListener("change", async (evento) => {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;
    try {
      const texto = await arquivo.text();
      const { instrumento, respostas: respostasCarregadas } = desserializarRascunho(texto);
      if (!INSTRUMENTOS[instrumento]) {
        alert(`Rascunho de instrumento desconhecido: "${instrumento}".`);
        return;
      }
      if (instrumento !== instrumentoAtual) {
        if (!confirm(`Este rascunho é de "${instrumento}", diferente do instrumento atual (${instrumentoAtual}). Carregar mesmo assim?`)) return;
        instrumentoAtual = instrumento;
      }
      respostas = respostasCarregadas;
      persistirAutosave();
      document.getElementById("tela-triagem").hidden = true;
      document.getElementById("tela-painel").hidden = false;
      montarTelaPainel();
      alert("Rascunho carregado com sucesso.");
    } catch (erro) {
      alert("Não foi possível carregar o rascunho: " + erro.message);
    }
  });
}

iniciar();
```

Nota: o autosave deixa de ser carregado antes da triagem (o app antigo
carregava o autosave de TR-Bens assim que a página abria, mesmo antes de o
elaborador confirmar o instrumento — hoje isso não faz mais sentido porque
existe mais de uma chave de autosave possível). O autosave passa a ser
carregado **depois** de a triagem escolher o instrumento, na própria chave
daquele instrumento. Isso é uma mudança de comportamento pequena e
deliberada — vale confirmar com a usuária ao revisar o PR se ela prefere
que a triagem já detecte automaticamente "há um rascunho de TR-Bens salvo"
antes mesmo de perguntar; se preferir, isso vira um ajuste de follow-up, não
bloqueia este plano.

- [ ] **Step 3: Testar manualmente no navegador**

Run: `npx serve .` (ou qualquer servidor estático — o projeto já documenta
em commit anterior que precisa ser servido por HTTP, não aberto via
`file://`) e abrir no navegador:
1. Escolher "Avançar para o TR-Bens" → confirmar que o painel abre igual a
   antes (regressão visual/funcional zero).
2. Voltar, escolher "Avançar para o TR-Serviços" → confirmar que o sumário
   lista as 20 seções corretas, os campos aparecem, os pills Sim/Não/
   variante funcionam, a tabela de infrações extras (Seção 15) permite
   adicionar/remover linha, e "Gerar Documento" baixa
   `TR-Servicos-gerado.docx` sem erro no console.

- [ ] **Step 4: Commit**

```bash
git add js/app.js js/triagem.js
git commit -m "Generalize app shell to support multiple instruments (TR-Bens + TR-Serviços)"
```

---

### Task 7: Testes de geração ponta a ponta para o TR-Serviços

**Files:**
- Test: `tests/gerador-docx.test.mjs` (novos casos, mesmo arquivo das tasks
  anteriores — mantém tudo que testa `.docx` num só lugar, como já é o
  padrão do projeto)

**Interfaces:**
- Consumes: `SECOES_TR_SERVICOS` (Task 4), `Template_TR_Servicos.docx`
  (Task 5), `gerarDocx` (Tasks 1-3).
- Produces: nada consumido por outra task — é o teste de regressão que
  fecha o plano.

- [ ] **Step 1: Escrever os testes que fecham a matriz de condicionais**

```js
// em tests/gerador-docx.test.mjs
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";

function lerTemplateServicos() {
  const buffer = fs.readFileSync("templates/Template_TR_Servicos.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlServicos(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplateServicos(), respostas, SECOES_TR_SERVICOS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

test("TR-Serviços: vistoria obrigatória remove a declaração de quem não visitou", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S05_TipoVistoria: "Obrigatória" }));
  assert.ok(texto.includes("o licitante deverá realizar vistoria"));
  assert.ok(!texto.includes("O licitante que optar por não realizar a vistoria"));
});

test("TR-Serviços: vistoria facultativa mantém os dois lados (quem visita e quem não visita)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S05_TipoVistoria: "Facultativa" }));
  assert.ok(texto.includes("o licitante poderá realizar vistoria"));
  assert.ok(texto.includes("O licitante que optar por não realizar a vistoria"));
});

test("TR-Serviços: fiscalização por Fiscal Funcional usa só a alínea b única (Seção 11)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S08_TipoFiscalizacao: "Fiscal Funcional" }));
  assert.ok(texto.includes("o Fiscal Funcional deverá realizar a análise de toda a documentação apresentada, avaliar a execução"));
  assert.ok(!texto.includes("b.1)"));
  assert.ok(!texto.includes("b.2)"));
  assert.ok(!texto.includes("b.3)"));
});

test("TR-Serviços: Comissão de Fiscalização usa as alíneas b, b.1, b.2, b.3 e c (Seção 11)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S08_TipoFiscalizacao: "Comissão de Fiscalização" }));
  assert.ok(texto.includes("b.1) o Fiscal Funcional deverá realizar a análise"));
  assert.ok(texto.includes("b.2) o Fiscal Técnico deverá realizar a análise"));
  assert.ok(texto.includes("b.3) o Fiscal Setorial deverá realizar a análise"));
  assert.ok(!texto.includes("o Fiscal Funcional deverá realizar a análise de toda a documentação apresentada, avaliar a execução"));
});

test("TR-Serviços: natureza intelectual inclui o bloco de cessão de direitos autorais", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S10_NaturezaIntelectual: "Sim" }));
  assert.ok(texto.includes("ceder os direitos patrimoniais"));
  assert.ok(texto.includes("transição contratual com transferência de conhecimento"));
});

test("TR-Serviços: sem natureza intelectual, o bloco de cessão de direitos não aparece", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S10_NaturezaIntelectual: "Não" }));
  assert.ok(!texto.includes("ceder os direitos patrimoniais"));
});

test("TR-Serviços: infrações extras da Tabela 1 aparecem, as 9 fixas continuam presentes", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({
    S15_InfracoesExtras: [{ descricao: "Deixar de portar crachá de identificação", grau: "1" }],
  }));
  assert.ok(texto.includes("Deixar de portar crachá de identificação"));
  assert.ok(texto.includes("Permitir situação que crie a possibilidade de causar dano físico"));
});

test("TR-Serviços: sem infrações extras, a Tabela 1 tem só as 9 fixas", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({}));
  const ocorrenciasLinhaExtra = texto.split("((DESCRICAO))").length - 1;
  assert.equal(ocorrenciasLinhaExtra, 0);
});

test("TR-Serviços: base de cálculo da multa preenche os quatro pontos duplicados (e, f, g, Tabela 2)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S15_BaseCalculoMulta: "valor anual do Contrato" }));
  const ocorrencias = texto.split("valor anual do Contrato").length - 1;
  assert.ok(ocorrencias >= 4, `esperado valor em pelo menos 4 lugares, achou ${ocorrencias}`);
});

test("TR-Serviços: vigência plurianual usa a redação de até 5 anos (art. 488)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S04_TipoVigencia: "Plurianual (até 5 anos)" }));
  assert.ok(texto.includes("podendo ser prorrogado até o limite de 05 (cinco) anos"));
  assert.ok(!texto.includes("prorrogável na forma dos arts. 497 e 498"));
});
```

- [ ] **Step 2: Rodar e ajustar até passar**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: PASS. Se algum teste falhar por o texto real do template
(Task 5) não bater 100% com a string esperada aqui, ajustar a string do
teste para o texto real do template (não o contrário) — o template é a
fonte da verdade porque foi transcrito diretamente da minuta.

- [ ] **Step 3: Rodar a suíte inteira uma última vez**

Run: `node --test tests/`
Expected: PASS em todos os arquivos — nenhuma regressão no TR-Bens.

- [ ] **Step 4: Commit**

```bash
git add tests/gerador-docx.test.mjs
git commit -m "Add end-to-end generation tests for TR-Serviços conditionals"
```

---

## Verificação manual final (checklist, não automatizável neste porte)

Depois do Task 7 verde, gerar o `.docx` do TR-Serviços pelo navegador (Task
6, Step 3) com um caso preenchido por completo, cobrindo os dois lados de
cada condicional novo, e abrir no Word real conferindo que:
- os blocos multi-parágrafo (Seção 11, Seção 10) removem o lado certo por
  inteiro, sem parágrafo órfão nem numeração de alínea quebrada;
- as linhas extras da Tabela 1 (Seção 15) aparecem formatadas igual às 9
  fixas, sem a linha-modelo sobrando;
- os 5 anexos aparecem limpos, na ordem certa, com "Objeto" do Anexo I
  preenchido automaticamente;
- o TR-Bens continua gerando exatamente como antes (mesmo checklist da
  spec-mãe, seção "Testes").
