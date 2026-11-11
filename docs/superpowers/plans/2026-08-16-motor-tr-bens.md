# Motor do Gerador + TR-Bens Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a working, end-to-end static site that lets a GEPEA analyst
fill in the TR-Bens (goods) template through a two-screen guided form and
download a completed `.docx`, proving the shared engine (data model,
rendering, persistence, docx generation) that later plans reuse for
TR-Serviços, TR-Dispensa, and PB.

**Architecture:** Client-side only, no backend, no build step. The 20
sections of the official TR-Bens minuta (SEI 55379667) are declared as data
(`js/dados-tr-bens.js`), rendered by a generic engine (`js/motor.js`) that
knows nothing section-specific, and turned into a `.docx` by substituting
Word bookmarks inside a pre-built template (`templates/Template_TR_Bens.docx`)
using JSZip. All state lives in one in-memory object (`respostas`), mirrored
to `localStorage` and exportable as `.json`.

**Tech Stack:** Vanilla HTML/CSS/JS, `@govbr-ds/core` (gov.br Design
System, CDN), JSZip (vendored locally), Python 3.12 + `python-docx` (used
once, offline, to author the `.docx` template — not a runtime dependency of
the site), Node.js (used only to run tests against the pure-logic modules
and the docx generator — not shipped to the site).

## Global Constraints

- No servidor, sem build, sem npm no artefato final — `index.html` deve
  abrir direto no navegador (`file://` ou qualquer host estático).
- Nenhum dado do elaborador sai do navegador — a geração do `.docx`
  acontece inteiramente no cliente.
- Nenhum arquivo do projeto `Projeto-Basico-App` (`Gerador_PB.html`,
  `Template_PB.docx`, etc.) é lido ou modificado por este projeto.
- O texto oficial da minuta (SEI 55379667) é copiado verbatim; texto de
  ligação autoral segue a convenção de voz da spec (decidir antes de
  justificar, não citar o RLC em toda frase, sem enumeração simétrica
  forçada).
- Todas as 20 seções ficam sempre visíveis (sem toggle aplicável/não
  aplicável) — variação é conteúdo condicional dentro da seção, nunca
  visibilidade da seção inteira.
- Todo campo é tipado `parametro` ou `especificacao`; a barra de progresso
  conta só `especificacao`.

---

## Contexto de conteúdo real (não inventar campos fora desta lista)

O texto integral do TR-Bens (SEI 55379667) foi extraído e lido por
completo. As 20 seções e os pontos de preenchimento reais são:

1. DO OBJETO — descrição do material, valor estimado, metodologia da
   estimativa (média/mediana/menor preço), critério de julgamento, sigilo
   do valor, modo de disputa, dados orçamentários (ND/PTRES/Fonte).
2. DA JUSTIFICATIVA — motivação (texto livre), exceção de exclusividade
   ME/EPP (condicional).
3. DO OBJETIVO — objetivo (texto livre).
4. DA VIGÊNCIA DO CONTRATO — prazo em meses.
5. DA SELEÇÃO DO FORNECEDOR — parcela relevante para qualificação técnica
   (texto livre), exigência de amostra (condicional: prazo, local, horário).
6. DA DESCRIÇÃO DA AQUISIÇÃO — forma de fornecimento, local e horário de
   entrega, prazo de entrega, tipo de remessa, garantia dos bens (meses),
   prazo de substituição, prazos de recebimento provisório e definitivo.
7. DA GARANTIA CONTRATUAL — exigida ou não (condicional: percentual).
8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO — texto fixo, sem campos.
9. DAS OBRIGAÇÕES DA CONTRATANTE — texto fixo, sem campos.
10. DAS OBRIGAÇÕES DA CONTRATADA — prazo para comprovação do Programa de
    Integridade (meses).
11. DO PAGAMENTO — prazo em dias úteis.
12. DO REAJUSTE — reajustável ou não (condicional: índice de referência).
13. DA SUBCONTRATAÇÃO — permitida ou não (condicional: percentuais
    mínimo/máximo, detalhamento).
14. DA ALTERAÇÃO SUBJETIVA — texto fixo, sem campos.
15. DAS SANÇÕES ADMINISTRATIVAS — 7 percentuais de multa (alíneas a, b, c,
    e, f, g, h — a alínea d remete ao item 7.4, sem campo próprio) mais o
    limite de dias da alínea e, 8 campos ao todo (todos parâmetros
    numéricos).
16. DA RESCISÃO CONTRATUAL — texto fixo, sem campos.
17. DA VEDAÇÃO AO NEPOTISMO — texto fixo, sem campos.
18. DA SUSTENTABILIDADE AMBIENTAL — critérios específicos (texto livre).
19. DA MATRIZ DE RISCOS E RESPONSABILIDADES — texto fixo, remete ao Anexo I.
20. DAS DISPOSIÇÕES FINAIS — lista fixa de anexos (Anexo I — Matriz de
    Riscos; Anexo II — Modelo de Apresentação de Proposta) + 12 campos de
    identificação/assinatura (nome, setor, cargo dos responsáveis pela
    elaboração e aprovação, nas áreas demandante e técnica).

Anexos aplicáveis ao TR-Bens (confirmado no corpo da própria minuta, não
o conjunto mais amplo do RLC especulado na spec): **Anexo I — Matriz de
Riscos** (RLC Anexo II) e **Anexo II — Modelo de Apresentação de Proposta**.
Nenhum outro anexo (IMR, Planilha de Custos) aparece na minuta de Bens —
esses são de Serviços com mão de obra exclusiva, fora do escopo deste
plano.

---

## Estrutura de arquivos

```
Site Gerador de TR e PB/
├── index.html
├── css/
│   └── estilo.css
├── js/
│   ├── lib/
│   │   └── jszip.min.js          (vendorizado, Task 1)
│   ├── modelo.js                  (funções puras do motor — Task 2)
│   ├── dados-tr-bens.js           (as 20 seções, dado declarativo — Task 3)
│   ├── render.js                  (Tela 2 — Task 4)
│   ├── triagem.js                 (Tela 1 — Task 5)
│   ├── persistencia.js            (autosave + exportar/importar — Task 6)
│   ├── conferencia.js             (painel de pendências — Task 7)
│   ├── gerador-docx.js            (substituição de bookmarks — Task 9)
│   └── app.js                     (bootstrap, liga tudo — Task 11)
├── templates/
│   ├── build_template_tr_bens.py  (gera o .docx-modelo — Task 8)
│   └── Template_TR_Bens.docx      (artefato gerado pelo script acima)
├── assets/
│   └── logo-conab.svg             (placeholder — Task 1)
└── tests/
    ├── modelo.test.mjs            (Task 2)
    ├── dados-tr-bens.test.mjs     (Task 3)
    ├── persistencia.test.mjs      (Task 6)
    ├── conferencia.test.mjs       (Task 7)
    └── gerador-docx.test.mjs      (Task 9)
```

---

### Task 1: Scaffold do site + gov.br Design System + logo

**Files:**
- Create: `index.html`
- Create: `css/estilo.css`
- Create: `assets/logo-conab.svg`
- Create: `js/lib/jszip.min.js`
- Create: `js/app.js` (vazio por enquanto, só para o `<script>` não quebrar)

**Interfaces:**
- Produces: `index.html` com dois contêineres `<section id="tela-triagem">` e
  `<section id="tela-painel" hidden>`, um `<header>` com logo e um
  `<div id="conferencia-overlay" hidden>` — usados por todas as tarefas
  seguintes.

- [ ] **Step 1: Vendorizar o JSZip**

```bash
curl -sL https://unpkg.com/jszip@3.10.1/dist/jszip.min.js -o "js/lib/jszip.min.js"
```

- [ ] **Step 2: Verificar o download**

Run: `node -e "require('./js/lib/jszip.min.js'); console.log(typeof JSZip)"`
Expected: `function`

- [ ] **Step 3: Criar `assets/logo-conab.svg`**

Placeholder simples (a logo oficial em vetor entra depois, quando a
usuária fornecer o arquivo definitivo):

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 60" width="240" height="60">
  <rect width="240" height="60" fill="#1351b4"/>
  <text x="12" y="38" font-family="Rawline, sans-serif" font-size="24" fill="#ffffff">CONAB</text>
</svg>
```

- [ ] **Step 4: Criar `index.html`**

```html
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Gerador de TR e PB — Conab</title>
  <link rel="stylesheet" href="https://cdngovbr-ds.estaleiro.serpro.gov.br/design-system/fonts/rawline/css/rawline.css">
  <link rel="stylesheet" href="https://unpkg.com/@govbr-ds/core@latest/dist/core.min.css">
  <link rel="stylesheet" href="css/estilo.css">
</head>
<body class="br-body">
  <header class="br-header" id="cabecalho-site">
    <div class="br-header-top">
      <img src="assets/logo-conab.svg" alt="Conab" class="header-logo">
      <span class="header-titulo">Gerador de Termo de Referência e Projeto Básico</span>
    </div>
  </header>

  <main>
    <section id="tela-triagem" class="br-container"></section>
    <section id="tela-painel" class="br-container" hidden></section>
  </main>

  <div id="conferencia-overlay" class="br-modal" hidden></div>

  <script src="js/lib/jszip.min.js"></script>
  <script src="js/app.js" type="module"></script>
</body>
</html>
```

- [ ] **Step 5: Criar `css/estilo.css`**

```css
.header-logo { height: 40px; margin-right: 1rem; }
.header-titulo { font-weight: 600; font-size: 1.1rem; }
.campo-parametro { border-left: 3px solid #1351b4; padding-left: .5rem; }
.campo-especificacao { border-left: 3px solid #ffcd07; padding-left: .5rem; }
[hidden] { display: none !important; }
```

- [ ] **Step 6: Criar `js/app.js` vazio (placeholder de bootstrap)**

```js
// Bootstrap da aplicação — preenchido na Task 11.
console.log("Gerador de TR e PB carregado.");
```

- [ ] **Step 7: Verificar no navegador**

Abrir `index.html` diretamente (duplo clique ou `file:///.../index.html`).
Confirmar no console: sem erros, mensagem "Gerador de TR e PB carregado."
aparece, cabeçalho com logo placeholder é exibido.

- [ ] **Step 8: Commit**

```bash
git add index.html css/estilo.css assets/logo-conab.svg js/lib/jszip.min.js js/app.js
git commit -m "Scaffold static site with gov.br DS shell and vendored JSZip"
```

---

### Task 2: Motor de dados — funções puras (visibilidade, progresso, pendências)

**Files:**
- Create: `js/modelo.js`
- Test: `tests/modelo.test.mjs`

**Interfaces:**
- Consumes: nada (módulo raiz, sem dependências).
- Produces:
  - `campoVisivel(campo, respostas) -> boolean`
  - `progresso(secoes, respostas) -> { totalEspecificacao: number, preenchidos: number }`
  - `pendencias(secoes, respostas) -> Array<{ secaoId: string, campoId: string, pergunta: string }>`
  - Formato esperado de `campo`: `{ id, tipo: "parametro"|"especificacao", pergunta, input, opcoes?, mostrarSe? }`
  - Formato esperado de `secoes`: `Array<{ id, titulo, campos: Array<campo> }>`
  - Formato de `mostrarSe`: `{ campoId: valorEsperado }` (só um par por
    campo, suficiente para todos os condicionais do TR-Bens).

- [ ] **Step 1: Escrever o teste que falha**

```js
// tests/modelo.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { campoVisivel, progresso, pendencias } from "../js/modelo.js";

test("campo sem mostrarSe é sempre visível", () => {
  const campo = { id: "x", tipo: "parametro", pergunta: "?" };
  assert.equal(campoVisivel(campo, {}), true);
});

test("campo com mostrarSe só aparece quando a condição bate", () => {
  const campo = { id: "y", tipo: "parametro", pergunta: "?", mostrarSe: { permite: "Sim" } };
  assert.equal(campoVisivel(campo, { permite: "Não" }), false);
  assert.equal(campoVisivel(campo, { permite: "Sim" }), true);
});

test("progresso conta só campos de especificação visíveis", () => {
  const secoes = [{
    id: "01", titulo: "T", campos: [
      { id: "a", tipo: "parametro", pergunta: "?" },
      { id: "b", tipo: "especificacao", pergunta: "?" },
      { id: "c", tipo: "especificacao", pergunta: "?", mostrarSe: { a: "X" } },
    ],
  }];
  const r1 = progresso(secoes, { a: "outro", b: "preenchido" });
  assert.equal(r1.totalEspecificacao, 1); // "c" não visível, não conta
  assert.equal(r1.preenchidos, 1);

  const r2 = progresso(secoes, { a: "X", b: "", c: "preenchido" });
  assert.equal(r2.totalEspecificacao, 2); // "c" agora visível
  assert.equal(r2.preenchidos, 1); // "b" vazio não conta
});

test("pendencias lista especificações visíveis e vazias, com a pergunta", () => {
  const secoes = [{
    id: "01", titulo: "T", campos: [
      { id: "a", tipo: "especificacao", pergunta: "Descreva A" },
    ],
  }];
  const r = pendencias(secoes, { a: "" });
  assert.deepEqual(r, [{ secaoId: "01", campoId: "a", pergunta: "Descreva A" }]);
  assert.deepEqual(pendencias(secoes, { a: "ok" }), []);
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/modelo.test.mjs`
Expected: FAIL — `Cannot find module '../js/modelo.js'`

- [ ] **Step 3: Implementar `js/modelo.js`**

```js
export function campoVisivel(campo, respostas) {
  if (!campo.mostrarSe) return true;
  const [campoId, valorEsperado] = Object.entries(campo.mostrarSe)[0];
  return respostas[campoId] === valorEsperado;
}

export function progresso(secoes, respostas) {
  let totalEspecificacao = 0;
  let preenchidos = 0;
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.tipo !== "especificacao") continue;
      if (!campoVisivel(campo, respostas)) continue;
      totalEspecificacao++;
      const valor = respostas[campo.id];
      if (valor && String(valor).trim() !== "") preenchidos++;
    }
  }
  return { totalEspecificacao, preenchidos };
}

export function pendencias(secoes, respostas) {
  const resultado = [];
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.tipo !== "especificacao") continue;
      if (!campoVisivel(campo, respostas)) continue;
      const valor = respostas[campo.id];
      if (!valor || String(valor).trim() === "") {
        resultado.push({ secaoId: secao.id, campoId: campo.id, pergunta: campo.pergunta });
      }
    }
  }
  return resultado;
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/modelo.test.mjs`
Expected: PASS — 4 testes.

- [ ] **Step 5: Commit**

```bash
git add js/modelo.js tests/modelo.test.mjs
git commit -m "Add pure engine functions: visibility, progress, pending fields"
```

---

### Task 3: Dados completos das 20 seções do TR-Bens

**Files:**
- Create: `js/dados-tr-bens.js`
- Test: `tests/dados-tr-bens.test.mjs`

**Interfaces:**
- Consumes: nada (dado puro).
- Produces: `SECOES_TR_BENS` — array de 20 seções no formato consumido pela
  Task 2 (`{ id, titulo, campos }`). Seções sem campo (8, 9, 14, 16, 17, 19)
  têm `campos: []`; o texto fixo delas não vive em `js/dados-tr-bens.js` —
  já está escrito diretamente no `.docx`-modelo pela Task 8, não duplicado
  aqui como dado à parte.

- [ ] **Step 1: Escrever o teste que falha**

```js
// tests/dados-tr-bens.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";

test("existem exatamente 20 seções, na ordem 01 a 20", () => {
  assert.equal(SECOES_TR_BENS.length, 20);
  assert.deepEqual(SECOES_TR_BENS.map(s => s.id),
    Array.from({ length: 20 }, (_, i) => String(i + 1).padStart(2, "0")));
});

test("seção 13 (Subcontratação) tem o campo condicional esperado", () => {
  const s13 = SECOES_TR_BENS.find(s => s.id === "13");
  const permite = s13.campos.find(c => c.id === "S13_PermiteSubcontratacao");
  assert.equal(permite.tipo, "parametro");
  assert.deepEqual(permite.opcoes, ["Sim", "Não"]);
  const percentualMin = s13.campos.find(c => c.id === "S13_PercentualMinSubcontratacao");
  assert.deepEqual(percentualMin.mostrarSe, { S13_PermiteSubcontratacao: "Sim" });
});

test("seção 15 (Sanções) tem os 7 percentuais de multa mais o limite de dias, 8 campos ao todo", () => {
  const s15 = SECOES_TR_BENS.find(s => s.id === "15");
  assert.equal(s15.campos.length, 8);
  assert.ok(s15.campos.every(c => c.tipo === "parametro"));
});

test("seções sem campo (texto fixo) têm array de campos vazio", () => {
  for (const id of ["08", "09", "14", "16", "17", "19"]) {
    const secao = SECOES_TR_BENS.find(s => s.id === id);
    assert.deepEqual(secao.campos, []);
  }
});

test("todo campo de especificação tem pergunta não vazia", () => {
  for (const secao of SECOES_TR_BENS) {
    for (const campo of secao.campos) {
      assert.ok(campo.pergunta && campo.pergunta.trim().length > 0,
        `campo ${campo.id} sem pergunta`);
    }
  }
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/dados-tr-bens.test.mjs`
Expected: FAIL — módulo não existe.

- [ ] **Step 3: Implementar `js/dados-tr-bens.js`**

```js
export const SECOES_TR_BENS = [
  { id: "01", titulo: "DO OBJETO", campos: [
    { id: "S01_DescricaoObjeto", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva o material a ser adquirido",
      ajuda: "Ex.: notebooks corporativos, mobiliário de escritório, transportadores helicoidais." },
    { id: "S01_ValorEstimado", tipo: "parametro", input: "moeda",
      pergunta: "Valor total estimado da contratação (R$)" },
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
      opcoes: ["Sim", "Não"] },
    { id: "S02_MotivoExcecaoMEEPP", tipo: "especificacao", input: "textarea",
      pergunta: "Qual exceção do art. 49 da LC 123/2006 se aplica e por quê?",
      mostrarSe: { S02_AplicaExcecaoMEEPP: "Sim" } },
  ]},
  { id: "03", titulo: "DO OBJETIVO", campos: [
    { id: "S03_Objetivo", tipo: "especificacao", input: "textarea",
      pergunta: "Qual o objetivo a ser atingido com a aquisição?" },
  ]},
  { id: "04", titulo: "DA VIGÊNCIA DO CONTRATO", campos: [
    { id: "S04_PrazoVigenciaMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo de vigência do contrato (meses)" },
  ]},
  { id: "05", titulo: "DA SELEÇÃO DO FORNECEDOR", campos: [
    { id: "S05_ParcelaRelevante", tipo: "especificacao", input: "textarea",
      pergunta: "Qual a parcela do objeto considerada relevante para fins de qualificação técnica, e o quantitativo mínimo exigido?",
      ajuda: "Não fixar parâmetro acima de 50% dos quantitativos, por analogia ao Acórdão 361/2017-TCU." },
    { id: "S05_ExigeAmostra", tipo: "parametro", input: "select",
      pergunta: "Será exigida amostra dos bens?", opcoes: ["Sim", "Não"] },
    { id: "S05_PrazoAmostraDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para apresentação da amostra (dias úteis)",
      mostrarSe: { S05_ExigeAmostra: "Sim" } },
    { id: "S05_LocalAmostra", tipo: "parametro", input: "texto",
      pergunta: "Local (gerência/setor e endereço) de entrega da amostra",
      mostrarSe: { S05_ExigeAmostra: "Sim" } },
    { id: "S05_HorarioAmostra", tipo: "parametro", input: "texto",
      pergunta: "Horário de recebimento da amostra",
      mostrarSe: { S05_ExigeAmostra: "Sim" } },
  ]},
  { id: "06", titulo: "DA DESCRIÇÃO DA AQUISIÇÃO", campos: [
    { id: "S06_FormaFornecimento", tipo: "parametro", input: "select",
      pergunta: "Forma de fornecimento", opcoes: ["Integral", "Parcelada"] },
    { id: "S06_LocalEntrega", tipo: "parametro", input: "texto",
      pergunta: "Local de entrega (gerência/setor e endereço)" },
    { id: "S06_HorarioEntrega", tipo: "parametro", input: "texto", pergunta: "Horário de entrega" },
    { id: "S06_PrazoEntregaDias", tipo: "parametro", input: "numero", pergunta: "Prazo de entrega (dias)" },
    { id: "S06_TipoRemessa", tipo: "parametro", input: "select",
      pergunta: "Tipo de remessa", opcoes: ["Única", "Parcelada"] },
    { id: "S06_GarantiaBensMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo mínimo de garantia dos bens (meses)" },
    { id: "S06_PrazoSubstituicaoDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para substituição do bem em garantia (dias)" },
    { id: "S06_PrazoRecebimentoProvisorioDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para recebimento provisório (dias)" },
    { id: "S06_PrazoRecebimentoDefinitivoDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para recebimento definitivo, contado do provisório (dias)" },
  ]},
  { id: "07", titulo: "DA GARANTIA CONTRATUAL", campos: [
    { id: "S07_ExigeGarantia", tipo: "parametro", input: "select",
      pergunta: "Será exigida garantia contratual?", opcoes: ["Sim", "Não"] },
    { id: "S07_PercentualGarantia", tipo: "parametro", input: "numero",
      pergunta: "Percentual da garantia contratual (máx. 5%, art. 439 RLC)",
      mostrarSe: { S07_ExigeGarantia: "Sim" } },
  ]},
  { id: "08", titulo: "DO ACOMPANHAMENTO E DA FISCALIZAÇÃO", campos: [] },
  { id: "09", titulo: "DAS OBRIGAÇÕES DA CONTRATANTE", campos: [] },
  { id: "10", titulo: "DAS OBRIGAÇÕES DA CONTRATADA", campos: [
    { id: "S10_PrazoProgramaIntegridadeMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo para comprovação do Programa de Integridade (meses)",
      ajuda: "Exigido quando o valor anual estimado supera R$ 1.500.000,00 — Anexo XXXII do RLC. Se não se aplicar, registre 0." },
  ]},
  { id: "11", titulo: "DO PAGAMENTO", campos: [
    { id: "S11_PrazoPagamentoDiasUteis", tipo: "parametro", input: "numero",
      pergunta: "Prazo de pagamento após recebimento definitivo (dias úteis, máx. 30 — art. 562 RLC)" },
  ]},
  { id: "12", titulo: "DO REAJUSTE", campos: [
    { id: "S12_Reajustavel", tipo: "parametro", input: "select",
      pergunta: "Os preços serão reajustáveis após um ano?", opcoes: ["Sim", "Não"] },
    { id: "S12_IndiceReajuste", tipo: "parametro", input: "texto",
      pergunta: "Índice de referência para o reajuste",
      mostrarSe: { S12_Reajustavel: "Sim" } },
  ]},
  { id: "13", titulo: "DA SUBCONTRATAÇÃO", campos: [
    { id: "S13_PermiteSubcontratacao", tipo: "parametro", input: "select",
      pergunta: "Será admitida subcontratação parcial do objeto?", opcoes: ["Sim", "Não"] },
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
    { id: "S15_MultaMoratoriaPercentualDia", tipo: "parametro", input: "numero",
      pergunta: "Multa moratória por dia de atraso (% sobre o valor da parcela não executada)" },
    { id: "S15_MultaMoratoriaLimiteDias", tipo: "parametro", input: "numero",
      pergunta: "Limite de dias de atraso para a multa moratória" },
    { id: "S15_MultaCompensatoriaParcial", tipo: "parametro", input: "numero",
      pergunta: "Multa compensatória por inexecução parcial (% sobre o valor da parcela)" },
    { id: "S15_MultaCompensatoriaTotal", tipo: "parametro", input: "numero",
      pergunta: "Multa compensatória por inexecução total (% sobre o valor total do contrato)" },
    { id: "S15_MultaRescisoria", tipo: "parametro", input: "numero",
      pergunta: "Multa rescisória por rescisão unilateral (% sobre o valor total do contrato)" },
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

Run: `node --test tests/dados-tr-bens.test.mjs`
Expected: PASS — 5 testes.

- [ ] **Step 5: Rodar também os testes da Task 2 com este dado real**

Run: `node --test tests/modelo.test.mjs tests/dados-tr-bens.test.mjs`
Expected: PASS — 9 testes ao todo.

- [ ] **Step 6: Commit**

```bash
git add js/dados-tr-bens.js tests/dados-tr-bens.test.mjs
git commit -m "Add full TR-Bens data model: 20 sections, all fields grounded in SEI 55379667"
```

---

### Task 4: Renderização da Tela 2 (painel principal)

**Files:**
- Create: `js/render.js`

**Interfaces:**
- Consumes: `SECOES_TR_BENS` (Task 3), `campoVisivel`, `progresso` (Task 2).
- Produces:
  - `renderizarSumario(secoes, container)` — preenche a navegação lateral.
  - `renderizarSecoes(secoes, respostas, container, aoMudar)` — preenche o
    acordeão de seções/campos; `aoMudar(campoId, novoValor)` é chamado a
    cada alteração de campo.
  - `atualizarProgresso(secoes, respostas, elementoBarra)` — atualiza a
    barra de progresso visível.

- [ ] **Step 1: Implementar `js/render.js`**

```js
import { campoVisivel, progresso } from "./modelo.js";

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

export function renderizarSecoes(secoes, respostas, container, aoMudar) {
  container.innerHTML = "";
  for (const secao of secoes) {
    const bloco = document.createElement("section");
    bloco.id = `secao-${secao.id}`;
    bloco.className = "br-accordion";

    const titulo = document.createElement("h2");
    titulo.textContent = `${secao.id}. ${secao.titulo}`;
    bloco.appendChild(titulo);

    for (const campo of secao.campos) {
      if (!campoVisivel(campo, respostas)) continue;
      bloco.appendChild(criarCampoInput(campo, respostas[campo.id], aoMudar));
    }
    container.appendChild(bloco);
  }
}

export function atualizarProgresso(secoes, respostas, elementoBarra) {
  const { totalEspecificacao, preenchidos } = progresso(secoes, respostas);
  const percentual = totalEspecificacao === 0 ? 100 : Math.round((preenchidos / totalEspecificacao) * 100);
  elementoBarra.setAttribute("value", String(percentual));
  elementoBarra.setAttribute("aria-valuenow", String(percentual));
  elementoBarra.textContent = `${preenchidos} de ${totalEspecificacao} campos de especificação preenchidos`;
}
```

- [ ] **Step 2: Verificação manual no navegador**

No `js/app.js`, temporariamente importar e chamar:

```js
import { SECOES_TR_BENS } from "./dados-tr-bens.js";
import { renderizarSumario, renderizarSecoes, atualizarProgresso } from "./render.js";

const respostas = {};
const painel = document.getElementById("tela-painel");
painel.hidden = false;
painel.innerHTML = '<nav id="sumario"></nav><div id="secoes"></div><progress id="barra"></progress>';
renderizarSumario(SECOES_TR_BENS, document.getElementById("sumario"));
renderizarSecoes(SECOES_TR_BENS, respostas, document.getElementById("secoes"),
  (id, valor) => { respostas[id] = valor; atualizarProgresso(SECOES_TR_BENS, respostas, document.getElementById("barra")); });
atualizarProgresso(SECOES_TR_BENS, respostas, document.getElementById("barra"));
```

Abrir `index.html` no navegador. Confirmar: 20 seções aparecem no sumário e
no corpo, os campos condicionais (ex. seção 13) ficam ocultos até a
resposta correta ser dada, e a barra de progresso muda ao digitar num
campo de especificação. Remover esse trecho temporário de `app.js` depois
de confirmado (ele será substituído pela integração real na Task 11).

- [ ] **Step 3: Commit**

```bash
git add js/render.js
git commit -m "Add Tela 2 renderer: sumário, seções condicionais, progress bar"
```

---

### Task 5: Tela 1 — Triagem (versão TR-Bens)

**Files:**
- Create: `js/triagem.js`

**Interfaces:**
- Consumes: nada além do DOM.
- Produces: `renderizarTriagem(container, aoConfirmar)` — `aoConfirmar()` é
  chamado quando o elaborador confirma a triagem; neste plano, o único
  caminho é sempre TR-Bens (a decisão real entre instrumentos/variantes é
  tarefa dos planos seguintes, quando TR-Serviços, TR-Dispensa e PB
  existirem). A tela ainda mostra as perguntas e a fundamentação, para não
  exigir retrabalho de interface depois.

- [ ] **Step 1: Implementar `js/triagem.js`**

```js
export function renderizarTriagem(container, aoConfirmar) {
  container.innerHTML = `
    <h1>Triagem inicial</h1>
    <p>Este protótipo cobre, por enquanto, apenas o Termo de Referência
    para aquisição de bens (TR-Bens), fundamentado no art. 127 do RLC
    10.901 — bem comum. As demais variantes de TR e o Projeto Básico
    entram em planos seguintes.</p>
    <button id="btn-confirmar-triagem" class="br-button primary" type="button">
      Avançar para o TR-Bens
    </button>
  `;
  container.querySelector("#btn-confirmar-triagem")
    .addEventListener("click", aoConfirmar);
}
```

- [ ] **Step 2: Verificação manual**

Chamar `renderizarTriagem(document.getElementById("tela-triagem"), () => alert("avançou"))`
a partir do console do navegador (com `index.html` aberto) e confirmar que
o botão aparece e dispara o alerta.

- [ ] **Step 3: Commit**

```bash
git add js/triagem.js
git commit -m "Add Tela 1 (triagem) scoped to TR-Bens for this plan"
```

---

### Task 6: Persistência — autosave e exportar/importar rascunho

**Files:**
- Create: `js/persistencia.js`
- Test: `tests/persistencia.test.mjs`

**Interfaces:**
- Consumes: nada.
- Produces:
  - `serializarRascunho(instrumento, respostas) -> string` (JSON)
  - `desserializarRascunho(json) -> { instrumento: string, respostas: object }` (lança erro se o JSON for inválido ou faltar `instrumento`/`respostas`)
  - `salvarAutosave(chave, instrumento, respostas)` — grava em `localStorage`
  - `carregarAutosave(chave) -> { instrumento, respostas } | null`

- [ ] **Step 1: Escrever o teste que falha**

```js
// tests/persistencia.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { serializarRascunho, desserializarRascunho } from "../js/persistencia.js";

test("serializar e desserializar preserva instrumento e respostas", () => {
  const json = serializarRascunho("TR-Bens", { S01_ValorEstimado: "1000" });
  const de_volta = desserializarRascunho(json);
  assert.equal(de_volta.instrumento, "TR-Bens");
  assert.deepEqual(de_volta.respostas, { S01_ValorEstimado: "1000" });
});

test("desserializar rejeita JSON sem instrumento ou respostas", () => {
  assert.throws(() => desserializarRascunho(JSON.stringify({ respostas: {} })), /instrumento/);
  assert.throws(() => desserializarRascunho(JSON.stringify({ instrumento: "TR-Bens" })), /respostas/);
});

test("desserializar rejeita JSON malformado", () => {
  assert.throws(() => desserializarRascunho("{não é json"));
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/persistencia.test.mjs`
Expected: FAIL — módulo não existe.

- [ ] **Step 3: Implementar `js/persistencia.js`**

```js
export function serializarRascunho(instrumento, respostas) {
  return JSON.stringify({ instrumento, respostas, salvoEm: new Date().toISOString() }, null, 2);
}

export function desserializarRascunho(json) {
  let dado;
  try {
    dado = JSON.parse(json);
  } catch {
    throw new Error("Arquivo de rascunho inválido: não é um JSON válido.");
  }
  if (!dado.instrumento) throw new Error("Arquivo de rascunho inválido: falta o campo 'instrumento'.");
  if (!dado.respostas) throw new Error("Arquivo de rascunho inválido: falta o campo 'respostas'.");
  return { instrumento: dado.instrumento, respostas: dado.respostas };
}

export function salvarAutosave(chave, instrumento, respostas) {
  localStorage.setItem(chave, serializarRascunho(instrumento, respostas));
}

export function carregarAutosave(chave) {
  const json = localStorage.getItem(chave);
  if (!json) return null;
  return desserializarRascunho(json);
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/persistencia.test.mjs`
Expected: PASS — 3 testes. (`salvarAutosave`/`carregarAutosave` dependem
de `localStorage`, indisponível no Node puro — por isso não têm teste
automatizado aqui; são verificados manualmente no Step 6.)

- [ ] **Step 5: Ligar os botões de rascunho na interface**

Adicionar em `index.html`, dentro do `<header>`:

```html
<button id="btn-salvar-rascunho" class="br-button secondary" type="button">Salvar Rascunho</button>
<input type="file" id="input-carregar-rascunho" accept="application/json" hidden>
<button id="btn-carregar-rascunho" class="br-button secondary" type="button">Carregar Rascunho</button>
```

Em `js/app.js` (trecho a integrar de vez na Task 11, mas já testável agora):

```js
import { serializarRascunho, desserializarRascunho, salvarAutosave, carregarAutosave } from "./persistencia.js";

const CHAVE_AUTOSAVE = "gerador-tr-pb:tr-bens:rascunho";

document.getElementById("btn-salvar-rascunho").addEventListener("click", () => {
  const json = serializarRascunho("TR-Bens", window.__respostas ?? {});
  const blob = new Blob([json], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "rascunho-tr-bens.json";
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
    const { instrumento, respostas } = desserializarRascunho(texto);
    if (instrumento !== "TR-Bens") {
      if (!confirm(`Este rascunho é de "${instrumento}", diferente do instrumento atual (TR-Bens). Carregar mesmo assim?`)) return;
    }
    window.__respostas = respostas;
    alert("Rascunho carregado. (Repopulação visual da tela será ligada na Task 11.)");
  } catch (erro) {
    alert("Não foi possível carregar o rascunho: " + erro.message);
  }
});
```

- [ ] **Step 6: Verificação manual no navegador**

Abrir `index.html`, definir `window.__respostas = { S01_ValorEstimado: "500" }`
no console, clicar "Salvar Rascunho" e confirmar que um `.json` é baixado
com esse conteúdo. Clicar "Carregar Rascunho", selecionar o arquivo baixado,
confirmar o alerta de sucesso. Depois, testar o caminho de erro: editar o
`.json` baixado trocando `"instrumento": "TR-Bens"` por `"instrumento": "PB"`,
carregar de novo e confirmar que aparece a pergunta de confirmação antes de
sobrescrever.

- [ ] **Step 7: Commit**

```bash
git add js/persistencia.js tests/persistencia.test.mjs index.html js/app.js
git commit -m "Add draft persistence: export/import JSON with instrument mismatch warning"
```

---

### Task 7: Painel de conferência (pendências antes de gerar)

**Files:**
- Create: `js/conferencia.js`
- Test: `tests/conferencia.test.mjs`

**Interfaces:**
- Consumes: `pendencias` (Task 2), `SECOES_TR_BENS` (Task 3).
- Produces: `renderizarConferencia(secoes, respostas, overlay, aoConfirmarGeracao)`

- [ ] **Step 1: Escrever o teste que falha**

```js
// tests/conferencia.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { montarMensagemConferencia } from "../js/conferencia.js";

test("sem pendências, mensagem confirma que está tudo preenchido", () => {
  const msg = montarMensagemConferencia([]);
  assert.match(msg, /Nenhum campo de especificação pendente/);
});

test("com pendências, lista cada uma por seção e pergunta", () => {
  const msg = montarMensagemConferencia([
    { secaoId: "01", campoId: "S01_DescricaoObjeto", pergunta: "Descreva o material" },
    { secaoId: "18", campoId: "S18_CriteriosEspecificos", pergunta: "Critérios específicos?" },
  ]);
  assert.match(msg, /2 campo\(s\) de especificação ainda não preenchidos/);
  assert.match(msg, /Seção 01: Descreva o material/);
  assert.match(msg, /Seção 18: Critérios específicos\?/);
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/conferencia.test.mjs`
Expected: FAIL — módulo não existe.

- [ ] **Step 3: Implementar `js/conferencia.js`**

```js
import { pendencias } from "./modelo.js";

export function montarMensagemConferencia(listaPendencias) {
  if (listaPendencias.length === 0) {
    return "Nenhum campo de especificação pendente. O documento está pronto para gerar.";
  }
  const linhas = listaPendencias.map(p => `- Seção ${p.secaoId}: ${p.pergunta}`);
  return `${listaPendencias.length} campo(s) de especificação ainda não preenchidos:\n\n${linhas.join("\n")}`;
}

export function renderizarConferencia(secoes, respostas, overlay, aoConfirmarGeracao) {
  const lista = pendencias(secoes, respostas);
  overlay.innerHTML = `
    <div class="br-modal-body">
      <pre>${montarMensagemConferencia(lista)}</pre>
      <button id="btn-confirmar-geracao" class="br-button primary" type="button">Gerar Documento</button>
      <button id="btn-fechar-conferencia" class="br-button secondary" type="button">Voltar e Preencher</button>
    </div>
  `;
  overlay.hidden = false;
  overlay.querySelector("#btn-confirmar-geracao").addEventListener("click", aoConfirmarGeracao);
  overlay.querySelector("#btn-fechar-conferencia").addEventListener("click", () => { overlay.hidden = true; });
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/conferencia.test.mjs`
Expected: PASS — 2 testes.

- [ ] **Step 5: Commit**

```bash
git add js/conferencia.js tests/conferencia.test.mjs
git commit -m "Add pre-generation review panel listing unfilled specification fields"
```

---

### Task 8: Construir o `.docx`-modelo com bookmarks (TR-Bens)

Este é o único passo que não segue o ciclo de teste JS: é um script Python
que roda uma vez (offline) para produzir o artefato `Template_TR_Bens.docx`
a partir do texto oficial já lido na íntegra (ver "Contexto de conteúdo
real"). O script fica no repositório para que o template possa ser
regenerado se o texto oficial mudar.

Convenção de bookmark: todo campo do `js/dados-tr-bens.js` vira um bookmark
de mesmo nome (`id`) no ponto exato onde o valor deve entrar. Para os 5
blocos condicionais de parágrafo inteiro (exceção ME/EPP na seção 2,
amostra na seção 5, garantia na seção 7, reajuste na seção 12,
subcontratação na seção 13), **os dois parágrafos alternativos** (o usado
quando a resposta é "Sim" e o usado quando é "Não") ficam cada um dentro do
seu próprio bookmark — `<ID_DA_PERGUNTA>_BLOCO_SIM` e
`<ID_DA_PERGUNTA>_BLOCO_NAO` — e o gerador (Task 9) sempre remove o
parágrafo do lado que **não** foi respondido, mantendo só o aplicável.
(Marcar apenas o lado "Não", como uma primeira versão deste plano chegou a
propor, deixa os dois parágrafos contraditórios no documento sempre que a
resposta for "Não" ou ficar em branco — o lado "Sim" nunca seria removido.
Os dois lados precisam de bookmark próprio.)

**Files:**
- Create: `templates/build_template_tr_bens.py`
- Create: `templates/Template_TR_Bens.docx` (artefato gerado pelo script)

**Interfaces:**
- Produces: `Template_TR_Bens.docx`, um arquivo Word real com 20 seções
  numeradas e os bookmarks descritos acima — consumido pela Task 9.

- [ ] **Step 1: Escrever `templates/build_template_tr_bens.py`**

```python
"""Gera Template_TR_Bens.docx a partir do texto oficial da minuta
SEI 55379667, com bookmarks nos pontos de preenchimento identificados
em js/dados-tr-bens.js. Rodar novamente sempre que o texto oficial mudar."""
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
    """partes: lista de strings (texto fixo) e tuplas (bookmark_name, texto_inicial)."""
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


def _wrap_paragraph_in_bookmark(paragrafo, nome_bookmark):
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start.set(docx.oxml.ns.qn("w:id"), bid)
    start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
    paragrafo._p.insert(0, start)
    end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end.set(docx.oxml.ns.qn("w:id"), bid)
    paragrafo._p.append(end)


def add_conditional_block(nome_base, texto_sim, texto_nao):
    """Dois parágrafos, cada um dentro do seu próprio bookmark
    (<nome_base>_BLOCO_SIM e <nome_base>_BLOCO_NAO). O gerador (Task 9)
    sempre remove o parágrafo do lado não escolhido pela resposta —
    nunca deixa os dois lados contraditórios no documento final."""
    p_sim = doc.add_paragraph()
    p_sim.add_run(texto_sim)
    _wrap_paragraph_in_bookmark(p_sim, f"{nome_base}_BLOCO_SIM")

    p_nao = doc.add_paragraph()
    p_nao.add_run(texto_nao)
    _wrap_paragraph_in_bookmark(p_nao, f"{nome_base}_BLOCO_NAO")


# Capa
doc.add_paragraph("COMPANHIA NACIONAL DE ABASTECIMENTO")
titulo = doc.add_paragraph()
titulo_run = titulo.add_run("TERMO DE REFERÊNCIA")
titulo_run.bold = True

# 1. DO OBJETO
add_heading("1. DO OBJETO")
add_paragraph_with_bookmarks([
    "1.1. O presente Termo de Referência tem por objeto a aquisição de ",
    ("S01_DescricaoObjeto", "(descrição do material a ser adquirido)"),
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento.",
])
add_paragraph_with_bookmarks([
    "1.2. O custo total estimado para aquisição dos produtos objetos deste Termo de Referência é de R$ ",
    ("S01_ValorEstimado", "________"),
    ". O critério utilizado para a formação da estimativa foi de ",
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
])
add_paragraph_with_bookmarks([
    "1.3. O critério de julgamento das propostas será o de ",
    ("S01_CriterioJulgamento", "[menor preço ou maior desconto]"),
    " ofertado. Para fins de julgamento das propostas o valor estimado/valor de referência será ",
    ("S01_SigiloValor", "(público ou sigiloso)"),
    ". O modo de disputa adotado no certame será ",
    ("S01_ModoDisputa", "(aberto ou aberto e fechado)"),
    ".",
])
add_paragraph_with_bookmarks([
    "1.4. Os recursos orçamentários decorrentes da aquisição do objeto deste Termo de Referência "
    "estão consignados no Orçamento da Conab para o ano 20XX, conforme segue: ND: (",
    ("S01_ND", "XXXX"),
    "), PTRES: (",
    ("S01_PTRES", "XXXX"),
    "), FONTE: (",
    ("S01_Fonte", "XXXX"),
    ").",
])

# 2. DA JUSTIFICATIVA
add_heading("2. DA JUSTIFICATIVA")
add_paragraph_with_bookmarks(["2.1. ", ("S02_Motivacao", "(motivação da contratação)")])
doc.add_paragraph(
    "2.2. O material em apreço será adquirido na modalidade Pregão Eletrônico, por ser "
    "caracterizado como bem comum, na forma do Inciso XIII do art. 3º do Regulamento de "
    "Licitações e Contratos da Conab - RLC e, portanto, possui padrões de desempenho e "
    "qualidade que podem ser objetivamente definidos no Edital, por meio de especificações "
    "usuais de mercado.")
add_conditional_block(
    "S02_AplicaExcecaoMEEPP",
    texto_sim=(
        "2.3. Em que pese o valor estimado da contratação ser inferior ao limite de "
        "R$ 80.000,00 (oitenta mil reais), o certame não será realizado com exclusividade "
        "de participação para microempresas e empresas de pequeno porte, na forma da Lei "
        "Complementar 123/2006 e do Decreto 8.538/2015, em razão da contratação em apreço "
        "se enquadrar na exceção prevista no art. 49 da Lei nº 123/2006, isto é: "
    ),
    texto_nao="",
)
add_paragraph_with_bookmarks(["", ("S02_MotivoExcecaoMEEPP", "")])

# 3. DO OBJETIVO
add_heading("3. DO OBJETIVO")
add_paragraph_with_bookmarks(["3.1. ", ("S03_Objetivo", "(objetivo da aquisição)")])

# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_paragraph_with_bookmarks([
    "4.1. O prazo de vigência do contrato será de ",
    ("S04_PrazoVigenciaMeses", "____"),
    " meses (duração do contrato, conforme art. 461 a 463 do RLC), contados da data de sua "
    "assinatura, prorrogável na forma do art. 497 e 498 do RLC.",
])

# 5. DA SELEÇÃO DO FORNECEDOR
add_heading("5. DA SELEÇÃO DO FORNECEDOR")
doc.add_paragraph("5.1. DA QUALIFICAÇÃO TÉCNICA")
add_paragraph_with_bookmarks([
    "5.1.1. a.1) Para efeito de caracterização de compatibilidade é definida como relevante "
    "a seguinte parcela do objeto: ",
    ("S05_ParcelaRelevante", "(descrever a parte mais relevante do objeto e o seu quantitativo)"),
])
doc.add_paragraph("5.2. DA AMOSTRA DOS BENS")
add_conditional_block(
    "S05_ExigeAmostra",
    texto_sim="5.2.1. Poderá ser solicitada à licitante provisoriamente classificada em primeiro lugar a apresentação de amostra dos bens.",
    texto_nao="5.2.1. Não será exigida amostra dos bens objeto desta licitação.",
)
add_paragraph_with_bookmarks([
    "5.2.1.1. Prazo para apresentação: ", ("S05_PrazoAmostraDias", "____"), " dias úteis.",
])
add_paragraph_with_bookmarks([
    "5.2.2.1. A amostra deverá ser apresentada em ", ("S05_LocalAmostra", "(local)"),
    ", no horário de ", ("S05_HorarioAmostra", "(horário)"), ".",
])

# 6. DA DESCRIÇÃO DA AQUISIÇÃO
add_heading("6. DA DESCRIÇÃO DA AQUISIÇÃO")
doc.add_paragraph("6.1. DA FORMA DE FORNECIMENTO, DO PRAZO, DO LOCAL DE ENTREGA E DA GARANTIA DOS BENS")
add_paragraph_with_bookmarks([
    "6.1.1. Os bens objetos deste Termo de Referência deverão ser entregues de forma ",
    ("S06_FormaFornecimento", "(integral/parcelada)"),
    ", devidamente embalados e lacrados, em ", ("S06_LocalEntrega", "(local)"),
    ", no horário de ", ("S06_HorarioEntrega", "(horário)"), ".",
])
add_paragraph_with_bookmarks([
    "6.1.2. O prazo de entrega dos bens é de ", ("S06_PrazoEntregaDias", "____"),
    " dias, contados a partir da assinatura do contrato, em remessa ",
    ("S06_TipoRemessa", "(única ou parcelada)"), ".",
])
add_paragraph_with_bookmarks([
    "6.1.4. Os bens deverão ter um período de garantia de, no mínimo, ",
    ("S06_GarantiaBensMeses", "____"), " meses, contados do recebimento definitivo pela Contratante.",
])
add_paragraph_with_bookmarks([
    "6.1.5. Durante o período de garantia, a Contratada deverá substituir o bem, no prazo "
    "máximo de ", ("S06_PrazoSubstituicaoDias", "____"), " dias, após ser comunicado.",
])
doc.add_paragraph("6.2. DO RECEBIMENTO DOS BENS")
add_paragraph_with_bookmarks([
    "6.2.1. Os bens serão recebidos provisoriamente no prazo de ",
    ("S06_PrazoRecebimentoProvisorioDias", "____"), " dias, pela fiscalização.",
])
add_paragraph_with_bookmarks([
    "6.2.3. Os bens serão recebidos definitivamente no prazo de ",
    ("S06_PrazoRecebimentoDefinitivoDias", "____"), " dias contados do recebimento provisório.",
])

# 7. DA GARANTIA CONTRATUAL
add_heading("7. DA GARANTIA CONTRATUAL")
add_conditional_block(
    "S07_ExigeGarantia",
    texto_sim="7.1. A Contratada deverá apresentar comprovante de prestação de garantia.",
    texto_nao="7.1. Não será exigida garantia contratual.",
)
add_paragraph_with_bookmarks([
    "Percentual da garantia: ", ("S07_PercentualGarantia", "____"), "% do valor do Contrato, "
    "conforme parâmetros do art. 439 do RLC.",
])

# 8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO (texto fixo)
add_heading("8. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO")
doc.add_paragraph(
    "8.1. Para a execução do ajuste, será adotado o método de trabalho baseado no conceito "
    "de delegação de responsabilidade. Nos termos dos art. 543 e 544 do RLC será designado "
    "fiscal, seu substituto, ou comissão de fiscalização para acompanhar e fiscalizar a "
    "entrega dos bens.")

# 9. DAS OBRIGAÇÕES DA CONTRATANTE (texto fixo)
add_heading("9. DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "9.1. São obrigações da Contratante: receber o objeto no prazo e condições estabelecidas; "
    "verificar a conformidade dos bens recebidos; comunicar à Contratada sobre imperfeições; "
    "rejeitar fornecimento em desacordo; acompanhar e fiscalizar o cumprimento das obrigações "
    "da Contratada; efetuar o pagamento no prazo e forma estabelecidos.")

# 10. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("10. DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph(
    "10.1. A Contratada deve cumprir todas as obrigações constantes neste Termo, seus anexos "
    "e sua proposta, incluindo efetuar a entrega em perfeitas condições, responsabilizar-se "
    "pelos vícios e danos do objeto, substituir ou reparar às suas expensas, comunicar "
    "atrasos, manter as condições de habilitação, prestar esclarecimentos e indicar preposto.")
add_paragraph_with_bookmarks([
    "h) comprovar a implementação do Programa de Integridade, nos moldes do art. 603 e "
    "Anexo XXXII do RLC, no prazo de até ", ("S10_PrazoProgramaIntegridadeMeses", "____"),
    " meses após a assinatura do contrato.",
])

# 11. DO PAGAMENTO
add_heading("11. DO PAGAMENTO")
add_paragraph_with_bookmarks([
    "11.1. O pagamento será efetuado em até ", ("S11_PrazoPagamentoDiasUteis", "____"),
    " dias úteis após o recebimento definitivo do material, de acordo com a apresentação da "
    "Nota Fiscal/Fatura.",
])

# 12. DO REAJUSTE
add_heading("12. DO REAJUSTE")
add_conditional_block(
    "S12_Reajustavel",
    texto_sim="12.1. Dentro do prazo de vigência do contrato, os preços contratados poderão sofrer reajuste após o interregno de um ano.",
    texto_nao="12.1. Os preços são fixos e irreajustáveis no prazo de um ano, contado da data limite para a apresentação das propostas.",
)
add_paragraph_with_bookmarks(["Índice de referência: ", ("S12_IndiceReajuste", "")])

# 13. DA SUBCONTRATAÇÃO
add_heading("13. DA SUBCONTRATAÇÃO")
add_conditional_block(
    "S13_PermiteSubcontratacao",
    texto_sim="13.1. É permitida a subcontratação parcial do objeto.",
    texto_nao="13.1. Não será admitida a subcontratação do objeto licitatório.",
)
add_paragraph_with_bookmarks([
    "Limites: entre ", ("S13_PercentualMinSubcontratacao", "____"), "% e ",
    ("S13_PercentualMaxSubcontratacao", "____"), "% do valor total do contrato.",
])
add_paragraph_with_bookmarks(["Detalhamento: ", ("S13_DetalheSubcontratacao", "")])
doc.add_paragraph("13.2. A subcontratação deverá observar o disposto no artigo 524 do RLC.")

# 14. DA ALTERAÇÃO SUBJETIVA (texto fixo)
add_heading("14. DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "14.1. É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa "
    "jurídica, desde que sejam observados pela nova pessoa jurídica todos os requisitos de "
    "habilitação exigidos na licitação, sejam mantidas as demais cláusulas e condições do "
    "Contrato, não haja prejuízo à execução do objeto pactuado e haja a anuência expressa "
    "da Conab à continuidade do Contrato.")

# 15. DAS SANÇÕES ADMINISTRATIVAS
add_heading("15. DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "15.1. A contratada, em caso de inadimplemento de suas obrigações, garantido o "
    "contraditório e ampla defesa, ficará sujeita às seguintes sanções previstas no RLC e "
    "na Lei nº 13.303, de 2016: advertência; multa moratória; multa compensatória; multa "
    "rescisória, para os casos de rescisão unilateral; suspensão do direito de participar "
    "de licitação e impedimento de contratar com a Conab, por até 02 (dois) anos.")
add_paragraph_with_bookmarks([
    "a) multa por não regularização da documentação de habilitação: ",
    ("S15_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado.",
])
add_paragraph_with_bookmarks([
    "b) multa pelas condutas dos art. 576 a 580 do RLC: ",
    ("S15_MultaCondutasArt576a580", "____"), "% sobre o valor estimado.",
])
add_paragraph_with_bookmarks([
    "c) multa pela recusa em assinar o Contrato: ",
    ("S15_MultaRecusaAssinarContrato", "____"), "% sobre o valor homologado.",
])
add_paragraph_with_bookmarks([
    "e) multa moratória de ", ("S15_MultaMoratoriaPercentualDia", "____"),
    "% por dia de atraso injustificado, até o limite de ",
    ("S15_MultaMoratoriaLimiteDias", "____"), " dias.",
])
add_paragraph_with_bookmarks([
    "f) multa compensatória de ", ("S15_MultaCompensatoriaParcial", "____"),
    "% sobre o valor da parcela não executada, em caso de inexecução parcial.",
])
add_paragraph_with_bookmarks([
    "g) multa compensatória de ", ("S15_MultaCompensatoriaTotal", "____"),
    "% sobre o valor total do contrato, em caso de inexecução total.",
])
add_paragraph_with_bookmarks([
    "h) multa rescisória de ", ("S15_MultaRescisoria", "____"),
    "% sobre o valor total do contrato, em caso de rescisão unilateral.",
])

# 16. DA RESCISÃO CONTRATUAL (texto fixo)
add_heading("16. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "16.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução "
    "parcial poderá ensejar a sua rescisão, com as consequências cabíveis, conforme "
    "disposto nos Arts. 568 a 572 do RLC. A rescisão poderá ser por ato unilateral e "
    "escrito da Conab, amigável, por acordo entre as partes, ou judicial.")

# 17. DA VEDAÇÃO AO NEPOTISMO (texto fixo)
add_heading("17. DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "17.1. Conforme disposto no parágrafo único do art. 12 do RLC e do artigo 7º do "
    "Decreto nº 7.203, de 2010, fica vedada a contratação de empregado ou dirigente da "
    "Conab como pessoa física, ou de pessoa com relação de parentesco até o terceiro grau "
    "civil com autoridade do Ministério da Agricultura, dirigente ou empregado da Conab "
    "com atribuições na área de licitação e contratação.")

# 18. DA SUSTENTABILIDADE AMBIENTAL
add_heading("18. DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "18.1. Compete à Contratada, no que couber, atender os critérios de sustentabilidade "
    "ambiental previstos no art. 10 do RLC.")
doc.add_paragraph(
    "18.2. A Contratada se responsabiliza administrativamente, civilmente e penalmente por "
    "qualquer dano causado pelo seu produto ao meio ambiente.")
add_paragraph_with_bookmarks(["18.3. ", ("S18_CriteriosEspecificos", "")])

# 19. DA MATRIZ DE RISCOS E RESPONSABILIDADES (texto fixo)
add_heading("19. DA MATRIZ DE RISCOS E RESPONSABILIDADES")
doc.add_paragraph(
    "19.1. A Matriz de Riscos é a cláusula contratual definidora dos riscos e das "
    "responsabilidades entre o CONTRATANTE e a CONTRATADA, conforme hipóteses elencadas na "
    "Matriz de Riscos - Anexo I do Termo de Referência.")

# 20. DAS DISPOSIÇÕES FINAIS
add_heading("20. DAS DISPOSIÇÕES FINAIS")
doc.add_paragraph(
    "20.1. Integram este Termo de Referência os seguintes Anexos: a) Anexo I - Matriz de "
    "Riscos; e b) Anexo II - Modelo de Apresentação de Proposta.")
add_paragraph_with_bookmarks(["Elaborador — Área Demandante: ", ("S20_ElabDemandanteNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S20_ElabDemandanteSetor", ""), " — Cargo: ", ("S20_ElabDemandanteCargo", "")])
add_paragraph_with_bookmarks(["Elaborador — Área Técnica: ", ("S20_ElabTecnicaNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S20_ElabTecnicaSetor", ""), " — Cargo: ", ("S20_ElabTecnicaCargo", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Demandante: ", ("S20_AprovDemandanteNome", ""), " — Setor: ", ("S20_AprovDemandanteSetor", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Técnica: ", ("S20_AprovTecnicaNome", ""), " — Setor: ", ("S20_AprovTecnicaSetor", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Demandante: ", ("S20_SuperDemandanteNome", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Técnica: ", ("S20_SuperTecnicaNome", "")])

# ANEXO I - MATRIZ DE RISCOS (limpo, oficial, sem geração de conteúdo)
doc.add_page_break()
add_heading("ANEXO I DO TERMO DE REFERÊNCIA")
add_heading("MATRIZ DE RISCOS")
add_paragraph_with_bookmarks(["Objeto: ", ("S01_DescricaoObjeto_AnexoI", "")])
tabela = doc.add_table(rows=1, cols=8)
tabela.style = "Table Grid"
cabecalhos = ["Item", "Fases", "Evento de Risco", "Causas", "Consequências",
              "Probabilidade (P)", "Impacto (I)", "Nível de Risco (P x I)"]
for celula, texto in zip(tabela.rows[0].cells, cabecalhos):
    celula.text = texto
for _ in range(5):
    tabela.add_row()
doc.add_paragraph(
    "Escala de Probabilidade: 1 Muito Baixa, 2 Baixa, 3 Média, 4 Alta, 5 Muito Alta.")
doc.add_paragraph(
    "Escala de Impacto: 1 Muito Baixo, 2 Baixo, 3 Médio, 4 Alto, 5 Muito Alto.")
doc.add_paragraph(
    "Nível de Risco: 1 a 2 Baixo; 3 a 6 Médio; 8 a 12 Elevado; 15 a 25 Extremo.")

# ANEXO II - MODELO DE APRESENTAÇÃO DE PROPOSTA (limpo, oficial)
doc.add_page_break()
add_heading("ANEXO II DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE APRESENTAÇÃO DE PROPOSTA")
tabela2 = doc.add_table(rows=1, cols=5)
tabela2.style = "Table Grid"
for celula, texto in zip(tabela2.rows[0].cells, ["Item", "Especificações", "CATMAT", "Quantidade", "Valor Total"]):
    celula.text = texto
for _ in range(5):
    tabela2.add_row()

doc.save("Template_TR_Bens.docx")
print("Template_TR_Bens.docx gerado com sucesso.")
```

- [ ] **Step 2: Rodar o script**

```bash
PY="C:/Users/Ana Paula Ribeiro/AppData/Local/Programs/Python/Python312/python.exe"
cd "templates" && "$PY" build_template_tr_bens.py
```

Expected: `Template_TR_Bens.docx gerado com sucesso.`

- [ ] **Step 3: Verificar os bookmarks gerados**

```bash
PY="C:/Users/Ana Paula Ribeiro/AppData/Local/Programs/Python/Python312/python.exe"
"$PY" -c "
import docx
d = docx.Document('templates/Template_TR_Bens.docx')
xml = d.element.xml
import re
nomes = re.findall(r'w:name=\"([^\"]+)\"', xml)
print(len(nomes), 'bookmarks encontrados')
faltando = set()
import json, re as re2
with open('js/dados-tr-bens.js', encoding='utf-8') as f:
    conteudo = f.read()
ids_esperados = re2.findall(r'id: \"(S\d\d_[A-Za-z0-9]+)\"', conteudo)
CONDICIONAIS = {'S02_AplicaExcecaoMEEPP','S05_ExigeAmostra','S07_ExigeGarantia','S12_Reajustavel','S13_PermiteSubcontratacao'}
for esperado in ids_esperados:
    if esperado in CONDICIONAIS:
        continue  # controladores de condicional não têm bookmark próprio, só _BLOCO_SIM/_BLOCO_NAO
    if esperado not in nomes:
        faltando.add(esperado)
print('faltando:', faltando if faltando else 'nenhum')
"
```

Expected: 65 bookmarks encontrados (59 ids declarados em
`js/dados-tr-bens.js`, menos os 5 controladores de condicional que não têm
bookmark próprio, mais os 10 bookmarks `_BLOCO_SIM`/`_BLOCO_NAO` dos 5
blocos condicionais, mais o bookmark auxiliar `S01_DescricaoObjeto_AnexoI`
= 54 + 10 + 1 = 65); `faltando: nenhum`.

- [ ] **Step 4: Abrir no Word e conferir visualmente**

Abrir `templates/Template_TR_Bens.docx` no Word real. Confirmar: as 20
seções aparecem na ordem certa com os títulos certos, os dois anexos
aparecem no final com as tabelas limpas, e nenhum texto parece corrompido
ou fora de ordem.

- [ ] **Step 5: Commit**

```bash
git add templates/build_template_tr_bens.py templates/Template_TR_Bens.docx
git commit -m "Generate TR-Bens Word template with bookmarks at every fill point"
```

---

### Task 9: Geração do `.docx` no navegador (substituição de bookmarks)

**Files:**
- Create: `js/gerador-docx.js`
- Test: `tests/gerador-docx.test.mjs`

**Interfaces:**
- Consumes: `JSZip` (global, Task 1), `Template_TR_Bens.docx` como
  `ArrayBuffer` (via `fetch` no navegador; via `fs.readFileSync` no teste
  Node), `SECOES_TR_BENS` (Task 3).
- Produces:
  - `substituirBookmark(documentXml, nomeBookmark, novoTexto) -> string`
  - `removerBlocoBookmark(documentXml, nomeBookmark) -> string` (remove o
    parágrafo inteiro que contém o bookmark)
  - `gerarDocx(templateArrayBuffer, respostas, secoes) -> Promise<ArrayBuffer>`
    — aplica as substituições de todos os campos com resposta e, para cada
    um dos 5 condicionais, remove sempre o parágrafo do lado **não**
    escolhido: remove `_BLOCO_NAO` quando a resposta for "Sim", remove
    `_BLOCO_SIM` em qualquer outro caso (resposta "Não" ou campo em
    branco) — nunca deixa os dois lados no documento final.

- [ ] **Step 1: Escrever o teste que falha**

```js
// tests/gerador-docx.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { substituirBookmark, removerBlocoBookmark, gerarDocx } from "../js/gerador-docx.js";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";

test("substituirBookmark troca o texto dentro do bookmark, preservando o resto do XML", () => {
  const xmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S04_PrazoVigenciaMeses"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';
  const resultado = substituirBookmark(xmlFalso, "S04_PrazoVigenciaMeses", "36");
  assert.match(resultado, /<w:t>36<\/w:t>/);
  assert.doesNotMatch(resultado, /____/);
});

test("removerBlocoBookmark remove o parágrafo inteiro que contém o bookmark", () => {
  const xmlFalso =
    '<w:p><w:r><w:t>antes</w:t></w:r></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S07_ExigeGarantia_BLOCO_NAO"/>' +
    '<w:r><w:t>Não será exigida garantia contratual.</w:t></w:r>' +
    '<w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:r><w:t>depois</w:t></w:r></w:p>';
  const resultado = removerBlocoBookmark(xmlFalso, "S07_ExigeGarantia_BLOCO_NAO");
  assert.match(resultado, /antes/);
  assert.match(resultado, /depois/);
  assert.doesNotMatch(resultado, /Não será exigida garantia contratual/);
});

test("gerarDocx produz um .docx válido com os valores substituídos", async () => {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  const respostas = {
    S01_DescricaoObjeto: "notebooks corporativos",
    S01_ValorEstimado: "150.000,00",
    S04_PrazoVigenciaMeses: "24",
    S07_ExigeGarantia: "Não",
    S13_PermiteSubcontratacao: "Não",
  };
  const resultadoBuffer = await gerarDocx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  assert.match(documentXml, /notebooks corporativos/);
  assert.match(documentXml, /150\.000,00/);
  assert.match(documentXml, /<w:t>24<\/w:t>/);
  // Como S07_ExigeGarantia é "Não", o bloco negativo (BLOCO_NAO) permanece...
  assert.match(documentXml, /Não será exigida garantia contratual/);
  // ...e o bloco positivo (BLOCO_SIM) precisa ter sido removido, senão os
  // dois parágrafos contraditórios ficariam juntos no documento final.
  assert.doesNotMatch(documentXml, /A Contratada deverá apresentar comprovante de prestação de garantia/);
  // Mesma checagem no lado inverso: S13 respondeu "Não", então o texto de
  // "é permitida a subcontratação" (BLOCO_SIM) não pode sobreviver.
  assert.doesNotMatch(documentXml, /É permitida a subcontratação parcial do objeto/);
  assert.match(documentXml, /Não será admitida a subcontratação do objeto licitatório/);
});

test("gerarDocx remove o bloco negativo quando a resposta é Sim", async () => {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  const respostas = { S07_ExigeGarantia: "Sim", S07_PercentualGarantia: "3" };
  const resultadoBuffer = await gerarDocx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  assert.match(documentXml, /A Contratada deverá apresentar comprovante de prestação de garantia/);
  assert.doesNotMatch(documentXml, /Não será exigida garantia contratual/);
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: FAIL — módulo `js/gerador-docx.js` não existe.

- [ ] **Step 3: Implementar `js/gerador-docx.js`**

```js
function localizarParagrafoDoBookmark(xml, nomeBookmark) {
  const marcador = `w:name="${nomeBookmark}"`;
  const posNome = xml.indexOf(marcador);
  if (posNome === -1) return null;
  const inicioParagrafo = xml.lastIndexOf("<w:p>", posNome) !== -1
    ? xml.lastIndexOf("<w:p>", posNome)
    : xml.lastIndexOf("<w:p ", posNome);
  const fimParagrafo = xml.indexOf("</w:p>", posNome) + "</w:p>".length;
  return { inicioParagrafo, fimParagrafo };
}

export function substituirBookmark(documentXml, nomeBookmark, novoTexto) {
  const posicao = localizarParagrafoDoBookmark(documentXml, nomeBookmark);
  if (!posicao) return documentXml;
  const trechoParagrafo = documentXml.slice(posicao.inicioParagrafo, posicao.fimParagrafo);
  const trechoSubstituido = trechoParagrafo.replace(
    /(<w:bookmarkStart[^>]*w:name="[^"]*"[^>]*\/>)([\s\S]*?)(<w:bookmarkEnd)/,
    (match, inicio, meio, fim) => {
      const meioComTextoTrocado = meio.replace(/<w:t([^>]*)>[^<]*<\/w:t>/, `<w:t$1>${novoTexto}</w:t>`);
      return inicio + meioComTextoTrocado + fim;
    }
  );
  return documentXml.slice(0, posicao.inicioParagrafo) + trechoSubstituido + documentXml.slice(posicao.fimParagrafo);
}

export function removerBlocoBookmark(documentXml, nomeBookmark) {
  const posicao = localizarParagrafoDoBookmark(documentXml, nomeBookmark);
  if (!posicao) return documentXml;
  return documentXml.slice(0, posicao.inicioParagrafo) + documentXml.slice(posicao.fimParagrafo);
}

const CONDICIONAIS_TR_BENS = [
  { campoRespostaSim: "S02_AplicaExcecaoMEEPP", blocoSim: "S02_AplicaExcecaoMEEPP_BLOCO_SIM", blocoNao: "S02_AplicaExcecaoMEEPP_BLOCO_NAO" },
  { campoRespostaSim: "S05_ExigeAmostra", blocoSim: "S05_ExigeAmostra_BLOCO_SIM", blocoNao: "S05_ExigeAmostra_BLOCO_NAO" },
  { campoRespostaSim: "S07_ExigeGarantia", blocoSim: "S07_ExigeGarantia_BLOCO_SIM", blocoNao: "S07_ExigeGarantia_BLOCO_NAO" },
  { campoRespostaSim: "S12_Reajustavel", blocoSim: "S12_Reajustavel_BLOCO_SIM", blocoNao: "S12_Reajustavel_BLOCO_NAO" },
  { campoRespostaSim: "S13_PermiteSubcontratacao", blocoSim: "S13_PermiteSubcontratacao_BLOCO_SIM", blocoNao: "S13_PermiteSubcontratacao_BLOCO_NAO" },
];

export async function gerarDocx(templateArrayBuffer, respostas, secoes) {
  const zip = await JSZip.loadAsync(templateArrayBuffer);
  let documentXml = await zip.file("word/document.xml").async("string");

  for (const secao of secoes) {
    for (const campo of secao.campos) {
      const valor = respostas[campo.id];
      if (valor !== undefined && valor !== null && valor !== "") {
        documentXml = substituirBookmark(documentXml, campo.id, escaparXml(String(valor)));
      }
    }
  }

  for (const condicional of CONDICIONAIS_TR_BENS) {
    const respondeuSim = respostas[condicional.campoRespostaSim] === "Sim";
    // Sempre remove o lado NÃO escolhido — nunca deixa os dois parágrafos
    // contraditórios juntos no documento. Resposta ausente ou "Não" cai
    // no mesmo ramo do "Não" (mantém o texto negativo, remove o positivo).
    documentXml = removerBlocoBookmark(documentXml, respondeuSim ? condicional.blocoNao : condicional.blocoSim);
  }

  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "arraybuffer" });
}

function escaparXml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: PASS — 3 testes.

- [ ] **Step 5: Verificação manual — abrir o resultado no Word**

```bash
node -e "
import('./js/gerador-docx.js').then(async ({ gerarDocx }) => {
  const fs = await import('node:fs');
  const { SECOES_TR_BENS } = await import('./js/dados-tr-bens.js');
  const buffer = fs.readFileSync('templates/Template_TR_Bens.docx');
  const respostas = {
    S01_DescricaoObjeto: 'notebooks corporativos', S01_ValorEstimado: '150.000,00',
    S04_PrazoVigenciaMeses: '24', S07_ExigeGarantia: 'Sim', S07_PercentualGarantia: '3',
    S13_PermiteSubcontratacao: 'Não',
  };
  const resultado = await gerarDocx(buffer.buffer, respostas, SECOES_TR_BENS);
  fs.writeFileSync('tests/saida-manual.docx', Buffer.from(resultado));
  console.log('gerado');
});
"
```

Abrir `tests/saida-manual.docx` no Word. Confirmar: valores substituídos
aparecem corretamente, o parágrafo "Não será exigida garantia contratual."
sumiu (porque a resposta foi "Sim"), e "Não será admitida a subcontratação"
permanece (porque a resposta foi "Não"). Apagar `tests/saida-manual.docx`
depois de conferir (é artefato de verificação manual, não deve ser
commitado).

- [ ] **Step 6: Commit**

```bash
git add js/gerador-docx.js tests/gerador-docx.test.mjs
echo "tests/saida-manual.docx" >> .gitignore
git add .gitignore
git commit -m "Add browser-side docx generator: bookmark substitution and conditional block removal"
```

---

### Task 10: Anexos limpos anexados automaticamente

Já resolvido na Task 8: os dois anexos do TR-Bens (Matriz de Riscos e
Modelo de Apresentação de Proposta) já fazem parte do `Template_TR_Bens.docx`
gerado, na forma limpa e oficial, com o único campo automático (`Objeto`
no cabeçalho da Matriz de Riscos, via bookmark `S01_DescricaoObjeto_AnexoI`)
preenchido a partir da resposta já dada na Seção 1.

**Files:**
- Modify: `js/gerador-docx.js`

**Interfaces:**
- Consumes: `respostas.S01_DescricaoObjeto` (já existente).
- Produces: preenchimento automático do bookmark `S01_DescricaoObjeto_AnexoI`
  a partir do mesmo valor de `S01_DescricaoObjeto`, sem exigir nova
  pergunta ao elaborador.

- [ ] **Step 1: Escrever o teste que falha**

Adicionar a `tests/gerador-docx.test.mjs`:

```js
test("o campo Objeto do Anexo I é preenchido automaticamente a partir da Seção 1", async () => {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  const respostas = { S01_DescricaoObjeto: "cadeiras de escritório" };
  const resultadoBuffer = await gerarDocx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const ocorrencias = documentXml.split("cadeiras de escritório").length - 1;
  assert.ok(ocorrencias >= 2, "esperado o valor tanto na Seção 1 quanto no Anexo I");
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: FAIL — o bookmark do anexo não recebe o valor ainda.

- [ ] **Step 3: Implementar o preenchimento automático em `gerarDocx`**

Adicionar, logo após o laço de campos de `secoes` em `js/gerador-docx.js`:

```js
  if (respostas.S01_DescricaoObjeto) {
    documentXml = substituirBookmark(documentXml, "S01_DescricaoObjeto_AnexoI", escaparXml(String(respostas.S01_DescricaoObjeto)));
  }
```

(inserir antes do laço `for (const condicional of CONDICIONAIS_TR_BENS)`)

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: PASS — 4 testes.

- [ ] **Step 5: Commit**

```bash
git add js/gerador-docx.js tests/gerador-docx.test.mjs
git commit -m "Auto-fill Anexo I object field from Seção 1 answer"
```

---

### Task 11: Integração end-to-end + verificação manual completa

**Files:**
- Modify: `js/app.js`

**Interfaces:**
- Consumes: todos os módulos das Tasks 2 a 10.
- Produces: fluxo completo funcionando — triagem → painel → conferência →
  download do `.docx`.

- [ ] **Step 1: Escrever `js/app.js` final**

```js
import { SECOES_TR_BENS } from "./dados-tr-bens.js";
import { renderizarSumario, renderizarSecoes, atualizarProgresso } from "./render.js";
import { renderizarTriagem } from "./triagem.js";
import { renderizarConferencia } from "./conferencia.js";
import { serializarRascunho, desserializarRascunho } from "./persistencia.js";
import { gerarDocx } from "./gerador-docx.js";

const INSTRUMENTO = "TR-Bens";
const CHAVE_AUTOSAVE = "gerador-tr-pb:tr-bens:rascunho";

let respostas = {};

function carregarAutosaveSeExistir() {
  const salvo = localStorage.getItem(CHAVE_AUTOSAVE);
  if (!salvo) return;
  try {
    const { respostas: respostasSalvas } = desserializarRascunho(salvo);
    respostas = respostasSalvas;
  } catch {
    // rascunho corrompido no localStorage: ignora e começa do zero
  }
}

function persistirAutosave() {
  localStorage.setItem(CHAVE_AUTOSAVE, serializarRascunho(INSTRUMENTO, respostas));
}

function montarTelaPainel() {
  const painel = document.getElementById("tela-painel");
  painel.innerHTML = `
    <nav id="sumario" class="br-list"></nav>
    <div id="secoes"></div>
    <progress id="barra-progresso" value="0" max="100"></progress>
    <button id="btn-gerar" class="br-button primary" type="button">Gerar Documento</button>
  `;
  function algumCampoDependeDe(campoId) {
    return SECOES_TR_BENS.some((secao) =>
      secao.campos.some((campo) => campo.mostrarSe && Object.keys(campo.mostrarSe)[0] === campoId)
    );
  }

  function aoMudar(id, valor) {
    respostas[id] = valor;
    persistirAutosave();
    atualizarProgresso(SECOES_TR_BENS, respostas, document.getElementById("barra-progresso"));
    // Só reconstrói o DOM das seções quando o campo alterado controla a
    // visibilidade de outro campo (os 5 seletores condicionais). Refazer o
    // DOM a cada tecla digitada nos campos de texto perderia o foco/cursor
    // do elaborador no meio da digitação.
    if (algumCampoDependeDe(id)) {
      rerenderizar();
    }
  }

  function rerenderizar() {
    renderizarSecoes(SECOES_TR_BENS, respostas, document.getElementById("secoes"), aoMudar);
  }

  renderizarSumario(SECOES_TR_BENS, document.getElementById("sumario"));
  rerenderizar();
  atualizarProgresso(SECOES_TR_BENS, respostas, document.getElementById("barra-progresso"));

  document.getElementById("btn-gerar").addEventListener("click", () => {
    const overlay = document.getElementById("conferencia-overlay");
    renderizarConferencia(SECOES_TR_BENS, respostas, overlay, async () => {
      overlay.hidden = true;
      try {
        const resposta = await fetch("templates/Template_TR_Bens.docx");
        if (!resposta.ok) {
          throw new Error(`Não foi possível carregar o modelo (${resposta.status}).`);
        }
        const templateBuffer = await resposta.arrayBuffer();
        const docxGerado = await gerarDocx(templateBuffer, respostas, SECOES_TR_BENS);
        const blob = new Blob([docxGerado], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "TR-Bens-gerado.docx";
        a.click();
        URL.revokeObjectURL(url);
      } catch (erro) {
        alert("Não foi possível gerar o documento: " + erro.message);
      }
    });
  });
}

function iniciar() {
  carregarAutosaveSeExistir();
  renderizarTriagem(document.getElementById("tela-triagem"), () => {
    document.getElementById("tela-triagem").hidden = true;
    document.getElementById("tela-painel").hidden = false;
    montarTelaPainel();
  });

  document.getElementById("btn-salvar-rascunho").addEventListener("click", () => {
    const json = serializarRascunho(INSTRUMENTO, respostas);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "rascunho-tr-bens.json";
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
      if (instrumento !== INSTRUMENTO) {
        if (!confirm(`Este rascunho é de "${instrumento}", diferente do instrumento atual (${INSTRUMENTO}). Carregar mesmo assim?`)) return;
      }
      respostas = respostasCarregadas;
      persistirAutosave();
      montarTelaPainel();
    } catch (erro) {
      alert("Não foi possível carregar o rascunho: " + erro.message);
    }
  });
}

iniciar();
```

- [ ] **Step 2: Verificação manual completa (checklist)**

Abrir `index.html` no navegador e, sem usar o console, percorrer o fluxo
real de ponta a ponta:

1. Tela de triagem aparece; clicar "Avançar para o TR-Bens".
2. Painel principal aparece com sumário de 20 seções e barra de progresso
   em 0%.
3. Preencher `S01_DescricaoObjeto`; confirmar que a barra de progresso
   sobe.
4. Ir até a seção 13 (Subcontratação); confirmar que os campos de
   percentual e detalhamento só aparecem depois de responder "Sim" em
   "Admite subcontratação parcial?".
5. Clicar "Salvar Rascunho"; confirmar que baixa um `.json`.
6. Recarregar a página (F5); confirmar que os dados voltam sozinhos
   (autosave via `localStorage`).
7. Clicar "Gerar Documento" com campos de especificação ainda vazios;
   confirmar que o painel de conferência lista as pendências.
8. Clicar "Voltar e Preencher", preencher os campos pendentes, clicar
   "Gerar Documento" de novo; confirmar que a mensagem agora diz "Nenhum
   campo de especificação pendente."
9. Confirmar a geração; um arquivo `TR-Bens-gerado.docx` é baixado.
10. Abrir o `.docx` baixado no Word: nenhum bookmark aparece sem
    substituir (nenhum `____` residual nos campos preenchidos), os blocos
    condicionais escolhidos corretamente aparecem, e os dois anexos vêm
    no final com a Matriz de Riscos já com o campo Objeto preenchido.

- [ ] **Step 3: Rodar a suíte de testes completa**

Run: `node --test tests/`
Expected: PASS — todos os testes das Tasks 2, 3, 6, 7 e 9/10.

- [ ] **Step 4: Commit**

```bash
git add js/app.js
git commit -m "Wire full TR-Bens flow end-to-end: triagem, painel, conferência, docx download"
```

---

## Self-Review

**Cobertura da spec:** motor de dados declarativo (Task 2/3), duas telas
com painel de conferência sobreposto em vez de terceira tela (Task 4/5/7),
persistência dupla localStorage + JSON (Task 6), geração de `.docx` no
navegador via bookmarks (Task 8/9), anexos limpos com auto-preenchimento
de campo trivial (Task 8/10), logo da Conab no cabeçalho (Task 1), gov.br
Design System (Task 1). A única peça da spec não coberta neste plano é a
lista completa de anexos por variante de TR e o PB inteiro — ambos
explicitamente adiados para os Planos 2 e 3, conforme combinado.

**Placeholders:** nenhum "TBD"/"implementar depois" no corpo das tarefas;
todo texto de seção vem do documento oficial já lido na íntegra.

**Consistência de tipos:** `campo.id` usado como nome de bookmark em todos
os lugares (Task 3, Task 8, Task 9) sem variação de nome; `secoes` sempre
`Array<{id, titulo, campos}>` do início ao fim; `respostas` sempre um
objeto plano `{ [campoId]: valor }`.
