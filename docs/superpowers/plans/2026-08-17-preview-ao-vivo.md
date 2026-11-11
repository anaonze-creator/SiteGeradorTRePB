# Pré-visualização ao vivo do documento — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a live-updating side panel that renders the real document text (from the `.docx` template) with a single highlight color marking answer-derived text, for all four instruments.

**Architecture:** Reuse the existing `gerarDocx()` pipeline unchanged. A new module, `js/preview-docx.js`, converts the resulting `word/document.xml` into HTML via string/regex parsing (no DOM, no external dependency — same technique `js/gerador-docx.js` already uses), wrapping field-bookmark text in `<mark>`. `js/app.js` wires this into a debounced update on every form change.

**Tech Stack:** Vanilla JS (ES modules), JSZip (already vendored at `js/lib/jszip.min.js`), Node's built-in test runner.

## Global Constraints

- No new dependencies. `js/preview-docx.js` must be Node-testable with `node --test tests/*.test.mjs` exactly like the rest of the suite — this rules out `DOMParser` (confirmed absent from Node 24's globals) even though the approved design spec mentions it; use regex/string parsing consistent with `js/gerador-docx.js`'s existing style instead. This is a same-approach, different-mechanism correction — the spec's outcome (HTML from the real generated XML, one highlight color, reusing `gerarDocx()`) is unchanged.
- **Corrected understanding vs. the spec:** structural block-wrapper bookmarks (named `<campo>_BLOCO_<VARIANTE>`, including sentinela-wrapped ones — see `templates/build_template_pb.py`) are **not** stripped by `gerarDocx()` for the *surviving* variant — only the non-chosen variants get removed. So `word/document.xml` coming out of `gerarDocx()` still contains `_BLOCO_*`/`_LINHA_MODELO` bookmark tags for whatever content survived. `preview-docx.js` must explicitly recognize and skip these (never wrap them in `<mark>`, just drop the tags) rather than assuming "only field bookmarks remain."
- Field-bookmark text (including `duplicarPara` target bookmarks, e.g. `S01_DescricaoObjeto_AnexoI`) is always contained within a single paragraph — never spans multiple `<w:p>` — so per-paragraph processing is sufficient; no cross-paragraph bookmark tracking is needed.
- One highlight color only (`mark.preview-destaque`), no per-field/per-section colors.
- Portuguese for all code comments/commit messages, matching the existing codebase.

---

### Task 1: Paragraph/run renderer (text, bold, italic, line breaks)

**Files:**
- Create: `js/preview-docx.js`
- Test: `tests/preview-docx.test.mjs`

**Interfaces:**
- Produces: `renderizarParagrafo(paragrafoXml: string): string` — takes the XML of one `<w:p>...</w:p>` element and returns an HTML `<p>...</p>` string. Internal helper, not exported yet (exported in Task 2 once bookmarks are handled, since tests need it directly until then — export it now as a named export to keep it testable).

- [ ] **Step 1: Write the failing test**

```js
// tests/preview-docx.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { renderizarParagrafo } from "../js/preview-docx.js";

test("renderizarParagrafo: texto simples", () => {
  const xml = `<w:p><w:r><w:t>Texto simples</w:t></w:r></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>Texto simples</p>");
});

test("renderizarParagrafo: negrito e itálico", () => {
  const xml = `<w:p>` +
    `<w:r><w:rPr><w:b/></w:rPr><w:t>Negrito</w:t></w:r>` +
    `<w:r><w:t xml:space="preserve"> e </w:t></w:r>` +
    `<w:r><w:rPr><w:i/></w:rPr><w:t>itálico</w:t></w:r>` +
    `</w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p><strong>Negrito</strong> e <em>itálico</em></p>");
});

test("renderizarParagrafo: quebra de linha vira <br>", () => {
  const xml = `<w:p><w:r><w:t>Linha 1</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Linha 2</w:t></w:r></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>Linha 1<br>Linha 2</p>");
});

test("renderizarParagrafo: escapa caracteres especiais de HTML", () => {
  const xml = `<w:p><w:r><w:t>A &amp; B &lt; C</w:t></w:r></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>A &amp; B &lt; C</p>");
});

test("renderizarParagrafo: parágrafo vazio", () => {
  const xml = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p></p>");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test tests/preview-docx.test.mjs`
Expected: FAIL — `js/preview-docx.js` does not exist / `renderizarParagrafo` is not exported.

- [ ] **Step 3: Write minimal implementation**

```js
// js/preview-docx.js

function escaparHtml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Concatena o texto de um run (<w:r>), aplicando negrito/itálico do <w:rPr>
// e convertendo <w:br/> em <br>. XML de entrada de <w:t> já vem decodificado
// pelas entidades XML padrão (&amp; etc.), então escaparHtml reaproveita as
// mesmas entidades — não há dupla decodificação/escape a fazer aqui.
function renderizarRun(runXml) {
  const negrito = /<w:b\s*\/>|<w:b\/>|<w:b>/.test(runXml);
  const italico = /<w:i\s*\/>|<w:i\/>|<w:i>/.test(runXml);

  const partes = [];
  const regexConteudo = /<w:t[^>]*>([\s\S]*?)<\/w:t>|<w:br\s*\/>/g;
  let m;
  while ((m = regexConteudo.exec(runXml)) !== null) {
    if (m[0].startsWith("<w:br")) {
      partes.push("<br>");
    } else {
      partes.push(escaparHtml(m[1]));
    }
  }
  let texto = partes.join("");
  if (!texto) return "";
  if (negrito) texto = `<strong>${texto}</strong>`;
  if (italico) texto = `<em>${texto}</em>`;
  return texto;
}

export function renderizarParagrafo(paragrafoXml) {
  const regexRun = /<w:r\b[\s\S]*?<\/w:r>/g;
  const runs = paragrafoXml.match(regexRun) || [];
  const conteudo = runs.map(renderizarRun).join("");
  return `<p>${conteudo}</p>`;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test tests/preview-docx.test.mjs`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add js/preview-docx.js tests/preview-docx.test.mjs
git commit -m "Add preview-docx paragraph renderer (text, bold, italic, line breaks)"
```

---

### Task 2: Bookmark highlighting (`<mark>`), skipping structural bookmarks

**Files:**
- Modify: `js/preview-docx.js`
- Test: `tests/preview-docx.test.mjs`

**Interfaces:**
- Consumes: nothing new from Task 1 beyond the existing `renderizarRun`/`escaparHtml` internals (not exported).
- Produces: `renderizarParagrafo` now also wraps field-bookmark text in `<mark class="preview-destaque">`, and silently drops `_BLOCO_*`/`_LINHA_MODELO` structural bookmark tags without wrapping. Same exported signature as Task 1 — no signature change, only behavior.

- [ ] **Step 1: Write the failing test**

```js
// append to tests/preview-docx.test.mjs

test("renderizarParagrafo: bookmark de campo vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="1" w:name="S01_Objeto"/>` +
    `<w:r><w:t>Aquisição de bens</w:t></w:r>` +
    `<w:bookmarkEnd w:id="1"/>` +
    `</w:p>`;
  assert.equal(
    renderizarParagrafo(xml),
    `<p><mark class="preview-destaque">Aquisição de bens</mark></p>`
  );
});

test("renderizarParagrafo: bookmark estrutural (_BLOCO_) não vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="2" w:name="PB01_NaturezaObjeto_BLOCO_EXECUCAO_DE_OBRA"/>` +
    `<w:r><w:t>Texto fixo do bloco</w:t></w:r>` +
    `<w:bookmarkEnd w:id="2"/>` +
    `</w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>Texto fixo do bloco</p>");
});

test("renderizarParagrafo: bookmark de campo dentro de bloco estrutural sobrevivente ainda vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="3" w:name="PB28_BDI_BLOCO_EXECUCAO_DE_OBRA"/>` +
    `<w:r><w:t>BDI: </w:t></w:r>` +
    `<w:bookmarkStart w:id="4" w:name="PB28_PercentualBDI"/>` +
    `<w:r><w:t>25%</w:t></w:r>` +
    `<w:bookmarkEnd w:id="4"/>` +
    `<w:bookmarkEnd w:id="3"/>` +
    `</w:p>`;
  assert.equal(
    renderizarParagrafo(xml),
    `<p>BDI: <mark class="preview-destaque">25%</mark></p>`
  );
});

test("renderizarParagrafo: bookmark _LINHA_MODELO não vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="5" w:name="S15_InfracoesExtras_LINHA_MODELO"/>` +
    `<w:r><w:t>((DESCRICAO))</w:t></w:r>` +
    `<w:bookmarkEnd w:id="5"/>` +
    `</w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>((DESCRICAO))</p>");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test tests/preview-docx.test.mjs`
Expected: FAIL on the 4 new tests — no `<mark>` wrapping exists yet.

- [ ] **Step 3: Write minimal implementation**

```js
// js/preview-docx.js — replace renderizarParagrafo with a version that
// tokenizes bookmarkStart/bookmarkEnd/w:r in document order.

// Bookmarks de bloco/variante (`add_bloco_variante` e as duas sentinelas em
// templates/build_template_*.py) sempre terminam em "_BLOCO_<VARIANTE>", e
// linhas-modelo de tabela em "_LINHA_MODELO". gerarDocx() só remove as
// variantes NÃO escolhidas — a variante sobrevivente mantém suas próprias
// tags de bookmark no XML de saída. Nenhum ID de campo real usa esses
// sufixos, então o padrão é seguro para distinguir "estrutural" de "campo".
function ehBookmarkEstrutural(nome) {
  return /_BLOCO_[A-Z0-9_]+$/.test(nome) || nome.endsWith("_LINHA_MODELO");
}

export function renderizarParagrafo(paragrafoXml) {
  const regexToken = /<w:bookmarkStart[^>]*\/>|<w:bookmarkEnd[^>]*\/>|<w:r\b[\s\S]*?<\/w:r>/g;
  const idsDeCampoAbertos = new Set();
  let saida = "";
  let marcando = false;
  let m;
  while ((m = regexToken.exec(paragrafoXml)) !== null) {
    const token = m[0];
    if (token.startsWith("<w:bookmarkStart")) {
      const nomeMatch = token.match(/w:name="([^"]+)"/);
      const idMatch = token.match(/w:id="([^"]+)"/);
      if (nomeMatch && idMatch && !ehBookmarkEstrutural(nomeMatch[1])) {
        idsDeCampoAbertos.add(idMatch[1]);
        if (!marcando) {
          saida += `<mark class="preview-destaque">`;
          marcando = true;
        }
      }
    } else if (token.startsWith("<w:bookmarkEnd")) {
      const idMatch = token.match(/w:id="([^"]+)"/);
      if (idMatch && idsDeCampoAbertos.has(idMatch[1])) {
        idsDeCampoAbertos.delete(idMatch[1]);
        if (marcando && idsDeCampoAbertos.size === 0) {
          saida += `</mark>`;
          marcando = false;
        }
      }
    } else {
      saida += renderizarRun(token);
    }
  }
  if (marcando) saida += `</mark>`;
  return `<p>${saida}</p>`;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test tests/preview-docx.test.mjs`
Expected: PASS (9 tests total)

- [ ] **Step 5: Commit**

```bash
git add js/preview-docx.js tests/preview-docx.test.mjs
git commit -m "Highlight field-bookmark text in preview, skip structural block bookmarks"
```

---

### Task 3: Table support + top-level document walker (`renderizarPreviewHtml`)

**Files:**
- Modify: `js/preview-docx.js`
- Test: `tests/preview-docx.test.mjs`

**Interfaces:**
- Consumes: `renderizarParagrafo(paragrafoXml): string` from Task 2.
- Produces: `renderizarPreviewHtml(documentXml: string): string` — the module's main export, used by `js/app.js` in Task 4. Walks `<w:body>`, in document order, converting each top-level `<w:tbl>` into a `<table>` (one `<tr>`/`<td>` per `w:tr`/`w:tc`, cell content rendered via `renderizarParagrafo` on each `w:p` inside the cell) and each top-level `<w:p>` via `renderizarParagrafo`. Anything else at body level (e.g. `<w:sectPr>`) is ignored.

- [ ] **Step 1: Write the failing test**

```js
// append to tests/preview-docx.test.mjs
import { renderizarPreviewHtml } from "../js/preview-docx.js";

test("renderizarPreviewHtml: parágrafos em sequência", () => {
  const xml = `<w:document><w:body>` +
    `<w:p><w:r><w:t>Primeiro</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>Segundo</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>` +
    `</w:body></w:document>`;
  assert.equal(renderizarPreviewHtml(xml), "<p>Primeiro</p><p>Segundo</p>");
});

test("renderizarPreviewHtml: tabela simples", () => {
  const xml = `<w:document><w:body>` +
    `<w:tbl>` +
      `<w:tr><w:tc><w:p><w:r><w:t>Célula 1</w:t></w:r></w:p></w:tc>` +
              `<w:tc><w:p><w:r><w:t>Célula 2</w:t></w:r></w:p></w:tc></w:tr>` +
    `</w:tbl>` +
    `</w:body></w:document>`;
  assert.equal(
    renderizarPreviewHtml(xml),
    "<table><tr><td><p>Célula 1</p></td><td><p>Célula 2</p></td></tr></table>"
  );
});

test("renderizarPreviewHtml: parágrafo antes e depois de tabela, na ordem certa", () => {
  const xml = `<w:document><w:body>` +
    `<w:p><w:r><w:t>Antes</w:t></w:r></w:p>` +
    `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Meio</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:t>Depois</w:t></w:r></w:p>` +
    `</w:body></w:document>`;
  assert.equal(
    renderizarPreviewHtml(xml),
    "<p>Antes</p><table><tr><td><p>Meio</p></td></tr></table><p>Depois</p>"
  );
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test tests/preview-docx.test.mjs`
Expected: FAIL — `renderizarPreviewHtml` is not exported yet.

- [ ] **Step 3: Write minimal implementation**

```js
// js/preview-docx.js — add at the end

function renderizarCelula(celulaXml) {
  const paragrafos = celulaXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  return `<td>${paragrafos.map(renderizarParagrafo).join("")}</td>`;
}

function renderizarLinha(linhaXml) {
  const celulas = linhaXml.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || [];
  return `<tr>${celulas.map(renderizarCelula).join("")}</tr>`;
}

function renderizarTabela(tabelaXml) {
  const linhas = tabelaXml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || [];
  return `<table>${linhas.map(renderizarLinha).join("")}</table>`;
}

export function renderizarPreviewHtml(documentXml) {
  const corpoMatch = documentXml.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!corpoMatch) return "";
  const corpo = corpoMatch[1];
  const regexBloco = /<w:tbl>[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>/g;
  const blocos = corpo.match(regexBloco) || [];
  return blocos
    .map((bloco) => (bloco.startsWith("<w:tbl") ? renderizarTabela(bloco) : renderizarParagrafo(bloco)))
    .join("");
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node --test tests/preview-docx.test.mjs`
Expected: PASS (12 tests total)

- [ ] **Step 5: Commit**

```bash
git add js/preview-docx.js tests/preview-docx.test.mjs
git commit -m "Add table support and top-level document walker to preview-docx"
```

---

### Task 4: Integration tests against real templates (all 4 instruments)

**Files:**
- Test: `tests/preview-docx-integracao.test.mjs`

**Interfaces:**
- Consumes: `gerarDocx` from `js/gerador-docx.js`, `renderizarPreviewHtml` from `js/preview-docx.js`, and each instrument's `SECOES_*` export — all already public.

- [ ] **Step 1: Write the failing test**

```js
// tests/preview-docx-integracao.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { renderizarPreviewHtml } from "../js/preview-docx.js";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";
import { SECOES_TR_DISPENSA } from "../js/dados-tr-dispensa.js";
import { SECOES_PB } from "../js/dados-pb.js";

globalThis.JSZip = JSZip;

function lerTemplate(caminho) {
  const buffer = fs.readFileSync(caminho);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarPreviewHtml(caminhoTemplate, respostas, secoes) {
  const docxBuffer = await gerarDocx(lerTemplate(caminhoTemplate), respostas, secoes);
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  return renderizarPreviewHtml(documentXml);
}

test("preview TR-Bens: resposta preenchida aparece marcada", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_TR_Bens.docx",
    { S01_DescricaoObjeto: "Aquisição de computadores portáteis" },
    SECOES_TR_BENS
  );
  assert.match(html, /<mark class="preview-destaque">[^<]*Aquisição de computadores portáteis[^<]*<\/mark>/);
});

test("preview TR-Serviços: resposta preenchida aparece marcada", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_TR_Servicos.docx",
    { S01_DescricaoObjeto: "Prestação de serviços de limpeza" },
    SECOES_TR_SERVICOS
  );
  assert.match(html, /<mark class="preview-destaque">[^<]*Prestação de serviços de limpeza[^<]*<\/mark>/);
});

test("preview TR-Dispensa: resposta preenchida aparece marcada", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_TR_Dispensa.docx",
    { S01_DescricaoObjeto: "Contratação direta de manutenção predial" },
    SECOES_TR_DISPENSA
  );
  assert.match(html, /<mark class="preview-destaque">[^<]*Contratação direta de manutenção predial[^<]*<\/mark>/);
});

test("preview PB: resposta preenchida aparece marcada, sem tags de bookmark cruas", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_PB.docx",
    { PB01_DescricaoObjeto: "Construção de galpão de armazenagem" },
    SECOES_PB
  );
  assert.match(html, /<mark class="preview-destaque">[^<]*Construção de galpão de armazenagem[^<]*<\/mark>/);
  assert.ok(!html.includes("w:bookmarkStart"), "HTML não deve conter tags OOXML cruas");
  assert.ok(!html.includes("_BLOCO_"), "nomes de bookmark estruturais não devem vazar para o HTML");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node --test tests/preview-docx-integracao.test.mjs`
Expected: FAIL if any regression exists in Tasks 1–3; otherwise this step should already PASS since the implementation is complete — if so, still run it once to confirm and treat this as the verification step (no separate "make it pass" work needed unless a real gap is found; if it fails, fix `js/preview-docx.js` before proceeding).

- [ ] **Step 3: Run test to verify it passes**

Run: `node --test tests/preview-docx-integracao.test.mjs`
Expected: PASS (4 tests)

- [ ] **Step 4: Run the full suite to check for regressions**

Run: `node --test tests/*.test.mjs`
Expected: PASS (all existing + new tests)

- [ ] **Step 5: Commit**

```bash
git add tests/preview-docx-integracao.test.mjs
git commit -m "Add end-to-end preview tests against real templates for all 4 instruments"
```

---

### Task 5: Wire the live preview panel into the app

**Files:**
- Modify: `index.html`
- Modify: `js/app.js`
- Modify: `css/estilo.css`

**Interfaces:**
- Consumes: `renderizarPreviewHtml` from `js/preview-docx.js` (Task 3), `gerarDocx` from `js/gerador-docx.js` (already imported in `js/app.js`).

- [ ] **Step 1: Add the preview panel container to `index.html`**

In `index.html`, change the `tela-painel` section to hold two columns:

```html
    <section id="tela-painel" class="br-container" hidden>
      <div class="layout-painel-preview">
        <div id="tela-painel-formulario"></div>
        <aside id="painel-preview" class="painel-preview" aria-label="Pré-visualização do documento"></aside>
      </div>
    </section>
```

`js/app.js`'s `montarTelaPainel()` currently does
`document.getElementById("tela-painel").innerHTML = ...` — Step 2 changes it
to target `tela-painel-formulario` instead, so the `aside` survives.

- [ ] **Step 2: Update `montarTelaPainel()` in `js/app.js`**

Replace the target of the existing `innerHTML` assignment and the buttons/DOM
lookups that were querying inside `#tela-painel` to instead query inside
`#tela-painel-formulario` (the cartão-painel now lives one level deeper).
Concretely, change:

```js
const painel = document.getElementById("tela-painel");
```

to:

```js
const painel = document.getElementById("tela-painel-formulario");
```

No other line in `montarTelaPainel()` needs to change — every subsequent
`document.getElementById(...)` call in that function (`#sumario`, `#secoes`,
`#barra-progresso`, `#btn-gerar`) already looks up elements by their own ID
anywhere in the document, not scoped to `painel`.

- [ ] **Step 3: Add `templateBuffer` caching and the debounced preview updater**

In `js/app.js`, add the import and the new logic inside `montarTelaPainel()`:

```js
import { renderizarPreviewHtml } from "./preview-docx.js";
```

Inside `montarTelaPainel()`, after the existing `const { secoes, titulo, templateUrl, nomeArquivoGerado } = config();` line, add:

```js
  let templateBufferCache = null;
  async function obterTemplateBuffer() {
    if (!templateBufferCache) {
      const resposta = await fetch(templateUrl);
      if (!resposta.ok) throw new Error(`Não foi possível carregar o modelo (${resposta.status}).`);
      templateBufferCache = await resposta.arrayBuffer();
    }
    return templateBufferCache;
  }

  let timeoutPreview = null;
  function agendarAtualizacaoPreview() {
    if (timeoutPreview) clearTimeout(timeoutPreview);
    timeoutPreview = setTimeout(atualizarPreview, 400);
  }

  async function atualizarPreview() {
    const painelPreview = document.getElementById("painel-preview");
    try {
      const buffer = await obterTemplateBuffer();
      const docxGerado = await gerarDocx(buffer, respostas, secoes);
      const zip = await JSZip.loadAsync(docxGerado);
      const documentXml = await zip.file("word/document.xml").async("string");
      painelPreview.innerHTML = renderizarPreviewHtml(documentXml);
    } catch (erro) {
      // Estado intermediário inconsistente de `respostas` enquanto o
      // elaborador ainda digita não deve interromper a digitação nem gerar
      // alert — só o fluxo de geração final (btn-gerar) precisa disso.
      console.warn("Não foi possível atualizar a pré-visualização.", erro);
    }
  }
```

Then change the existing `aoMudar` function to also call
`agendarAtualizacaoPreview()`:

```js
  function aoMudar(id, valor) {
    respostas[id] = valor;
    persistirAutosave();
    atualizarProgresso(secoes, respostas, document.getElementById("barra-progresso"));
    if (algumCampoDependeDe(id)) {
      rerenderizar();
    }
    agendarAtualizacaoPreview();
  }
```

And call it once after the existing `rerenderizar(); atualizarProgresso(...)` calls near the bottom of `montarTelaPainel()`, so the preview shows the template even before any answer is typed:

```js
  renderizarSumario(secoes, document.getElementById("sumario"));
  rerenderizar();
  atualizarProgresso(secoes, respostas, document.getElementById("barra-progresso"));
  atualizarPreview();
```

`JSZip` here is the same global already used by `gerador-docx.js` in the
browser (loaded via the `<script src="js/lib/jszip.min.js">` tag in
`index.html`) — no new import needed in `app.js`.

- [ ] **Step 4: Add CSS for the two-column layout and the highlight**

Append to `css/estilo.css`:

```css
.layout-painel-preview {
  display: flex;
  gap: 24px;
  align-items: flex-start;
}
.layout-painel-preview > div:first-child { flex: 1 1 55%; min-width: 0; }
.painel-preview {
  flex: 1 1 45%;
  min-width: 0;
  max-height: calc(100vh - 40px);
  overflow-y: auto;
  position: sticky;
  top: 20px;
  background: #ffffff;
  border: 1px solid #ededed;
  border-radius: 8px;
  padding: 16px 20px;
  font-size: .8125rem;
  line-height: 1.5;
}
.painel-preview table { border-collapse: collapse; width: 100%; margin: 8px 0; }
.painel-preview td { border: 1px solid #ddd; padding: 4px 6px; }
mark.preview-destaque {
  background: #fff3b8;
  padding: 0 2px;
  border-radius: 2px;
}
@media (max-width: 900px) {
  .layout-painel-preview { flex-direction: column; }
  .painel-preview { position: static; max-height: none; width: 100%; }
}
```

- [ ] **Step 5: Manual verification in the browser**

Start a static server (e.g. `npx serve .` or any static file server) and open
the app. For each of the 4 instruments: confirm the preview panel appears
beside the form, shows the template text on load, and updates ~400ms after
typing into a field, with the typed value highlighted. Confirm the "Gerar
Documento" button still works unaffected (it now reuses the same
`obterTemplateBuffer()` cache — check the existing `btn-gerar` click handler
still calls `fetch(templateUrl)` directly; if so, leave it as is, since
sharing the cache with it is an optional optimization, not required by this
plan — the design's "reused in both situations" note is satisfied as long as
correctness holds, and premature sharing risk (stale cache across instrument
switches) is avoided by scoping `templateBufferCache` inside
`montarTelaPainel()`, which already re-runs fully on every instrument
switch).

- [ ] **Step 6: Run the full test suite**

Run: `node --test tests/*.test.mjs`
Expected: PASS (no regressions — this task touches no tested logic, only DOM wiring and CSS)

- [ ] **Step 7: Commit**

```bash
git add index.html js/app.js css/estilo.css
git commit -m "Wire live document preview panel into the app shell"
```

---

## Execution Handoff

After all 5 tasks are complete, run `node --test tests/*.test.mjs` once more
to confirm the whole suite is green, then follow
`superpowers:finishing-a-development-branch` to decide whether to merge
(this plan was written for inline execution directly on `master`, per the
project's established "step by step, budget-conscious" pattern this
session — no worktree branch was created for this feature; if a worktree
was in fact used, merge and clean it up per that skill instead).
