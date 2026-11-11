# Motor TR-Dispensa — Fase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add the TR-Serviços por Dispensa instrument (Fase 1: TR funcional
com anexos limpos — sem pré-preenchimento cruzado, que é Fase 2) ao gerador
já existente (TR-Bens e TR-Serviços já mesclados).

**Architecture:** Mesmo motor declarativo já usado pelos outros dois
instrumentos. Uma pequena extensão (`blocosAdicionais`) permite que um
único campo controle mais de um bloco/seção/anexo com a mesma resposta —
necessária porque a escolha "Contrato Simplificado vs. Contrato" do
TR-Dispensa decide, simultaneamente, se 5 seções inteiras do corpo do TR
existem e qual de dois anexos grandes (com tabelas) sobrevive.

**Tech Stack:** Vanilla JS (ES modules), `python-docx`, `node --test`
(`node --test tests/*.test.mjs` — o glob explícito é necessário neste
ambiente Windows/Git Bash, `node --test tests/` sozinho falha).

## Global Constraints

- Nenhum texto é inventado: todo trecho fixo vem transcrito da minuta
  `documentos de referencia/SEI_55379661_Conab___RLC__Termo_Referencia__TR____Serv_Dispensa.pdf`
  (extraída com `pdftotext -layout -enc UTF-8`, lida por inteiro — 2.350
  linhas). Notas explicativas são suprimidas, como já orienta a própria
  minuta.
- Fase 1 não inclui pré-preenchimento cruzado do Anexo VI (Minuta de
  Contrato) — ele entra limpo, mesmo tratamento que os outros anexos hoje.
  Isso é Fase 2, plano futuro separado.
- Todo campo é `tipo: "parametro"` ou `tipo: "especificacao"`. `input`
  varia: `"texto"`, `"numero"`, `"moeda"`, `"textarea"`, `"select"`,
  `"tabela-linhas"`.
- Regressão zero em TR-Bens e TR-Serviços: os testes existentes
  (`tests/modelo.test.mjs`, `tests/gerador-docx.test.mjs`,
  `tests/dados-tr-bens.test.mjs`, `tests/dados-tr-servicos.test.mjs`)
  precisam continuar passando sem edição.

---

### Task 1: `blocosAdicionais` e `duplicarPara` no schema de campo

**Files:**
- Modify: `js/gerador-docx.js` (loop de remoção em `gerarDocx`, e
  generalização do preenchimento duplicado)
- Modify: `js/dados-tr-servicos.js` (migra `S07_BaseCalculoGarantia` e
  `S15_BaseCalculoMulta` para `duplicarPara`, removendo os `if` fixos)
- Test: `tests/gerador-docx.test.mjs`

**Interfaces:**
- Consumes: `controladoresBlocoVariante(secoes)` (já existe, sem mudança —
  continua devolvendo campos com `blocoTexto: true`).
- Produces: campos podem opcionalmente declarar
  `blocosAdicionais: string[]` (nomes de bookmark extras, além do próprio
  `campo.id`, que seguem a mesma lógica de variante da resposta) e/ou
  `duplicarPara: string[]` (nomes de bookmark extras que recebem o MESMO
  valor do campo via substituição simples — usado quando um valor de
  parâmetro aparece em mais de um lugar do documento, como a base de
  cálculo de multa repetida nas alíneas e/f/g e na Tabela 2).

**Nota sobre o gap encontrado:** hoje `gerarDocx` tem dois blocos com `if`
fixos nos nomes `S07_BaseCalculoGarantia` e `S15_BaseCalculoMulta`
(específicos do TR-Serviços) para preencher os bookmarks duplicados
`_74`/`_E`/`_F`/`_G`/`_TABELA2`. Isso não generaliza: o TR-Dispensa usa
`S16_BaseCalculoMulta` (nome de campo diferente, mesmo padrão), que nunca
seria preenchido nesses pontos. Este task substitui os `if` fixos por um
mecanismo genérico dirigido por dado (`duplicarPara`), usado tanto pelo
TR-Serviços quanto pelo TR-Dispensa.

- [ ] **Step 1: Escrever o teste que falha**

```js
// em tests/gerador-docx.test.mjs
test("gerarDocx remove blocosAdicionais junto com o bloco principal do campo", async () => {
  const secoesFalsas = [{
    id: "07", titulo: "T", campos: [
      { id: "S07_Tipo", tipo: "parametro", input: "select",
        opcoes: ["A", "B"], blocoTexto: true, padrao: "A",
        blocosAdicionais: ["S07_Tipo_EXTRA"] },
    ],
  }];
  const xmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S07_Tipo_BLOCO_B"/>' +
    '<w:r><w:t>texto principal B</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S07_Tipo_EXTRA_BLOCO_B"/>' +
    '<w:r><w:t>texto extra B</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:r><w:t>fica</w:t></w:r></w:p>';
  // resposta ausente -> cai no padrão "A", então tudo marcado "_BLOCO_B" deve sumir
  const zipFalso = { file: () => ({ async: async () => xmlFalso }), generateAsync: async () => xmlFalso };
  // Testamos a função de baixo nível diretamente para não depender de JSZip/zip real:
  const documentXmlFalso = xmlFalso;
  const { removerBlocoBookmark } = await import("../js/gerador-docx.js");
  let resultado = removerBlocoBookmark(documentXmlFalso, "S07_Tipo_BLOCO_B");
  resultado = removerBlocoBookmark(resultado, "S07_Tipo_EXTRA_BLOCO_B");
  assert.doesNotMatch(resultado, /texto principal B/);
  assert.doesNotMatch(resultado, /texto extra B/);
  assert.match(resultado, /fica/);
});
```

- [ ] **Step 2: Rodar e confirmar que este teste específico já passa (ele testa `removerBlocoBookmark` diretamente, não `gerarDocx` ainda) e depois escrever o teste de verdade via `gerarDocx`**

```js
// substituir o teste acima por este, que exercita gerarDocx de ponta a ponta
// contra um documento .docx real fabricado em memória com JSZip:
test("gerarDocx aplica blocosAdicionais junto com o bloco principal do campo", async () => {
  const secoesFalsas = [{
    id: "07", titulo: "T", campos: [
      { id: "S07_Tipo", tipo: "parametro", input: "select",
        opcoes: ["A", "B"], blocoTexto: true, padrao: "A",
        blocosAdicionais: ["S07_Tipo_EXTRA"] },
    ],
  }];
  const documentXmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S07_Tipo_BLOCO_B"/>' +
    '<w:r><w:t>texto principal B</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S07_Tipo_EXTRA_BLOCO_B"/>' +
    '<w:r><w:t>texto extra B</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:r><w:t>fica</w:t></w:r></w:p>';
  const zip = new JSZip();
  zip.file("word/document.xml", documentXmlFalso);
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  const resultadoBuffer = await gerarDocx(buffer, {}, secoesFalsas);
  const zipResultado = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zipResultado.file("word/document.xml").async("string");
  assert.doesNotMatch(documentXml, /texto principal B/);
  assert.doesNotMatch(documentXml, /texto extra B/);
  assert.match(documentXml, /fica/);
});
```

- [ ] **Step 3: Rodar e confirmar que falha**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: FAIL — `gerarDocx` hoje só remove `nomeBlocoVariante(campo.id, opcao)`,
nunca olha `blocosAdicionais`, então `S07_Tipo_EXTRA_BLOCO_B` sobrevive.

- [ ] **Step 4: Implementar em `js/gerador-docx.js`**

Trocar o loop atual (dentro de `gerarDocx`):

```js
for (const campo of controladoresBlocoVariante(secoes)) {
  const respondida = respostas[campo.id] || campo.padrao;
  const basesDeBookmark = [campo.id, ...(campo.blocosAdicionais || [])];
  for (const opcao of campo.opcoes) {
    if (opcao !== respondida) {
      for (const base of basesDeBookmark) {
        documentXml = removerBlocoBookmark(documentXml, nomeBlocoVariante(base, opcao));
      }
    }
  }
}
```

- [ ] **Step 5: Rodar e confirmar que passa, sem regressão**

Run: `node --test tests/*.test.mjs`
Expected: PASS em tudo — campos sem `blocosAdicionais` continuam
funcionando idênticos (o array extra fica vazio, o loop interno roda uma
vez só com `campo.id`, igual ao comportamento anterior).

- [ ] **Step 6: Escrever o teste que falha para `duplicarPara`**

```js
// em tests/gerador-docx.test.mjs
test("gerarDocx preenche duplicarPara com o mesmo valor do campo", async () => {
  const secoesFalsas = [{
    id: "16", titulo: "T", campos: [
      { id: "S16_BaseCalculoMulta", tipo: "parametro", input: "select",
        opcoes: ["valor do Contrato", "valor anual do Contrato"],
        pergunta: "?", duplicarPara: ["S16_BaseCalculoMulta_E", "S16_BaseCalculoMulta_F"] },
    ],
  }];
  const documentXmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S16_BaseCalculoMulta"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S16_BaseCalculoMulta_E"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="3" w:name="S16_BaseCalculoMulta_F"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>';
  const zip = new JSZip();
  zip.file("word/document.xml", documentXmlFalso);
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  const resultadoBuffer = await gerarDocx(buffer, { S16_BaseCalculoMulta: "valor anual do Contrato" }, secoesFalsas);
  const zipResultado = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zipResultado.file("word/document.xml").async("string");
  const ocorrencias = documentXml.split("valor anual do Contrato").length - 1;
  assert.equal(ocorrencias, 3, "esperado o valor no campo principal + 2 pontos de duplicarPara");
});
```

- [ ] **Step 7: Rodar e confirmar que falha**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: FAIL — `duplicarPara` ainda não é lido em lugar nenhum.

- [ ] **Step 8: Implementar `duplicarPara` em `js/gerador-docx.js`**

Logo depois do loop principal que já faz `substituirBookmark` para cada
campo visível (dentro de `gerarDocx`), adicionar:

```js
for (const secao of secoes) {
  for (const campo of secao.campos) {
    if (!campo.duplicarPara || !campo.duplicarPara.length) continue;
    if (!campoVisivel(campo, respostas)) continue;
    const valor = respostas[campo.id];
    if (valor === undefined || valor === null || valor === "") continue;
    for (const nomeExtra of campo.duplicarPara) {
      documentXml = substituirBookmark(documentXml, nomeExtra, escaparXml(String(valor)));
    }
  }
}
```

E remover os dois blocos fixos hoje existentes (`S07_BaseCalculoGarantia`/
`S07_BaseCalculoGarantia_74` e `S15_BaseCalculoMulta`/`_E`/`_F`/`_G`/
`_TABELA2`) — eles são substituídos pelo mecanismo genérico acima. O bloco
fixo de `S01_DescricaoObjeto`/`S01_DescricaoObjeto_AnexoI` **não** muda
neste task (é um caso à parte: duplica um campo de `tipo: "especificacao"`
para um anexo, não um campo de parâmetro reaproveitado dentro do próprio
corpo — generalizá-lo fica de fora do escopo deste plano, não há
necessidade concreta agora).

- [ ] **Step 9: Migrar os dois campos existentes em `js/dados-tr-servicos.js`**

```js
// S07_BaseCalculoGarantia (Seção 07):
{ id: "S07_BaseCalculoGarantia", tipo: "parametro", input: "select",
  pergunta: "A garantia de 5% (art. 439 RLC) incide sobre",
  opcoes: ["valor do Contrato", "valor anual do Contrato"],
  mostrarSe: { S07_ExigeGarantia: "Sim" },
  duplicarPara: ["S07_BaseCalculoGarantia_74"] },

// S15_BaseCalculoMulta (Seção 15):
{ id: "S15_BaseCalculoMulta", tipo: "parametro", input: "select",
  pergunta: "As multas moratória e compensatória por inexecução parcial incidem sobre",
  opcoes: ["valor do Contrato", "valor anual do Contrato", "valor mensal do Contrato"],
  ajuda: "Usado nas alíneas e/f/g da Seção 15 e na Tabela 2 (grau → percentual).",
  duplicarPara: ["S15_BaseCalculoMulta_E", "S15_BaseCalculoMulta_F", "S15_BaseCalculoMulta_G", "S15_BaseCalculoMulta_TABELA2"] },
```

- [ ] **Step 10: Rodar TODOS os testes**

Run: `node --test tests/*.test.mjs`
Expected: PASS em tudo, incluindo os testes de ponta a ponta do
TR-Serviços que já cobrem `S15_BaseCalculoMulta` (eles continuam passando
porque o comportamento observável não muda, só a forma como é produzido).

- [ ] **Step 11: Commit**

```bash
git add js/gerador-docx.js js/dados-tr-servicos.js tests/gerador-docx.test.mjs
git commit -m "Generalize duplicate-fill bookmarks (duplicarPara) and add blocosAdicionais"
```

---

### Task 2: `js/dados-tr-dispensa.js` — modelo de dados das 26 seções

**Files:**
- Create: `js/dados-tr-dispensa.js`
- Test: `tests/dados-tr-dispensa.test.mjs`

**Interfaces:**
- Consumes: formato de campo das Tasks anteriores (`blocoTexto`, `padrao`,
  `blocosAdicionais`, `mostrarSe`, `input: "tabela-linhas"`/`colunas`).
- Produces: `export const SECOES_TR_DISPENSA` — consumida por
  `js/app.js` (Task 4) e `js/gerador-docx.js`.

- [ ] **Step 1: Escrever o teste estrutural**

```js
// tests/dados-tr-dispensa.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_TR_DISPENSA } from "../js/dados-tr-dispensa.js";
import { controladoresBlocoVariante } from "../js/modelo.js";

test("SECOES_TR_DISPENSA tem 26 seções com ids únicos de 01 a 26", () => {
  assert.equal(SECOES_TR_DISPENSA.length, 26);
  const ids = SECOES_TR_DISPENSA.map((s) => s.id);
  const esperado = Array.from({ length: 26 }, (_, i) => String(i + 1).padStart(2, "0"));
  assert.deepEqual(ids, esperado);
});

test("todo campo com blocoTexto tem padrao dentro de opcoes", () => {
  for (const campo of controladoresBlocoVariante(SECOES_TR_DISPENSA)) {
    assert.ok(campo.opcoes.includes(campo.padrao),
      `${campo.id}: padrao "${campo.padrao}" não está em opcoes`);
  }
});

test("o campo de instrumento contratual declara os dois blocosAdicionais", () => {
  const campo = SECOES_TR_DISPENSA
    .flatMap((s) => s.campos)
    .find((c) => c.id === "S07_TipoInstrumentoContratual");
  assert.ok(campo, "S07_TipoInstrumentoContratual não encontrado");
  assert.deepEqual(campo.blocosAdicionais, [
    "S07_TipoInstrumentoContratual_SECOES21A25",
    "S07_TipoInstrumentoContratual_ANEXO",
  ]);
});

test("o campo de matriz de riscos declara o blocoAdicional do Anexo I", () => {
  const campo = SECOES_TR_DISPENSA
    .flatMap((s) => s.campos)
    .find((c) => c.id === "S20_ExigeMatrizRiscos");
  assert.ok(campo);
  assert.deepEqual(campo.blocosAdicionais, ["S20_ExigeMatrizRiscos_ANEXO_I"]);
});

test("ids de campo são únicos entre todas as seções", () => {
  const todosIds = SECOES_TR_DISPENSA.flatMap((s) => s.campos.map((c) => c.id));
  assert.equal(todosIds.length, new Set(todosIds).size);
});

test("todo campo tabela-linhas declara colunas", () => {
  for (const secao of SECOES_TR_DISPENSA) {
    for (const campo of secao.campos) {
      if (campo.input === "tabela-linhas") {
        assert.ok(Array.isArray(campo.colunas) && campo.colunas.length > 0, campo.id);
      }
    }
  }
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/dados-tr-dispensa.test.mjs`
Expected: FAIL — módulo não existe.

- [ ] **Step 3: Criar `js/dados-tr-dispensa.js`**

```js
export const SECOES_TR_DISPENSA = [
  { id: "01", titulo: "DO OBJETO", campos: [
    { id: "S01_DescricaoObjeto", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva os serviços a serem contratados",
      ajuda: "Ex.: manutenção pontual, serviço técnico de baixo valor, consultoria eventual." },
    { id: "S01_ValorEstimado", tipo: "parametro", input: "moeda",
      pergunta: "Custo total estimado da prestação dos serviços (R$)" },
    { id: "S01_MetodologiaEstimativa", tipo: "parametro", input: "select",
      pergunta: "Metodologia da estimativa de preço",
      opcoes: ["Média", "Mediana", "Menor Preço"] },
    { id: "S01_CriterioJulgamento", tipo: "parametro", input: "select",
      pergunta: "Critério de julgamento das propostas",
      opcoes: ["Menor preço", "Maior desconto"],
      ajuda: "Art. 264, incisos I e II, do RLC." },
    { id: "S01_ND", tipo: "parametro", input: "texto", pergunta: "Natureza de Despesa (ND)" },
    { id: "S01_PTRES", tipo: "parametro", input: "texto", pergunta: "PTRES" },
    { id: "S01_Fonte", tipo: "parametro", input: "texto", pergunta: "Fonte de recurso" },
  ]},
  { id: "02", titulo: "DA JUSTIFICATIVA", campos: [
    { id: "S02_Motivacao", tipo: "especificacao", input: "textarea",
      pergunta: "Qual a motivação da contratação?",
      ajuda: "Deve demonstrar que a contratação direta é a melhor (ou única) solução — DOD e inciso II, art. 129 do RLC; inciso VI, art. 413 do RLC." },
    { id: "S02_ExcecaoExclusividadeMEEPP", tipo: "parametro", input: "select",
      pergunta: "A dispensa (valor até R$ 80.000,00) mantém a exclusividade padrão para ME/EPP?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Sim",
      ajuda: "Ao contrário do TR-Serviços, aqui o padrão JÁ é exclusivo para ME/EPP (Decreto 8.538/2015); \"Não\" abre a disputa por se enquadrar numa exceção do art. 10 do Decreto." },
    { id: "S02_MotivoExcecaoExclusividade", tipo: "especificacao", input: "textarea",
      pergunta: "Qual exceção do art. 10 do Decreto 8.538/2015 se aplica e por quê?",
      mostrarSe: { S02_ExcecaoExclusividadeMEEPP: "Não" } },
  ]},
  { id: "03", titulo: "DO OBJETIVO", campos: [
    { id: "S03_Objetivo", tipo: "especificacao", input: "textarea",
      pergunta: "Qual o objetivo a ser atingido com a contratação?" },
  ]},
  { id: "04", titulo: "DA VIGÊNCIA DO CONTRATO", campos: [
    { id: "S04_TipoVigencia", tipo: "parametro", input: "select",
      pergunta: "Regime de vigência do contrato",
      opcoes: ["Até 1 ano, prorrogável", "Plurianual (até 5 anos)"],
      blocoTexto: true, padrao: "Até 1 ano, prorrogável" },
    { id: "S04_PrazoVigenciaMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo de vigência inicial (meses)",
      mostrarSe: { S04_TipoVigencia: "Até 1 ano, prorrogável" } },
    { id: "S04_PrazoVigenciaAnos", tipo: "parametro", input: "numero",
      pergunta: "Prazo de vigência inicial (anos, máx. 5)",
      mostrarSe: { S04_TipoVigencia: "Plurianual (até 5 anos)" } },
  ]},
  { id: "05", titulo: "DA SELEÇÃO DO FORNECEDOR", campos: [
    { id: "S05_PrazoLimiteEnvioProposta", tipo: "parametro", input: "texto",
      pergunta: "Data e horário limite para envio da proposta" },
    { id: "S05_PrazoExperienciaMinimaAnos", tipo: "parametro", input: "numero",
      pergunta: "Prazo mínimo de experiência anterior exigido para comprovação de capacidade técnica (anos)" },
    { id: "S05_QuantitativoServicoPrincipal", tipo: "especificacao", input: "textarea",
      pergunta: "Quantitativo mínimo exigido e identificação do(s) serviço(s) principal(is), para fins de comprovação de capacidade técnica",
      ajuda: "Não fixar parâmetro acima de 50% dos quantitativos, por analogia ao Acórdão 361/2017-TCU." },
    { id: "S05_RegistroEspecialAtividade", tipo: "especificacao", input: "textarea",
      pergunta: "A atividade exige registro/autorização especial de órgão regulador (ANVISA, ANP, Ministério da Agricultura etc.)? Se sim, especifique o documento, o órgão e o fundamento legal",
      ajuda: "Deixe em branco se o objeto não exigir registro ou autorização especial de funcionamento." },
    { id: "S05_TipoVistoria", tipo: "parametro", input: "select",
      pergunta: "A vistoria técnica será", opcoes: ["Obrigatória", "Facultativa"],
      blocoTexto: true, padrao: "Facultativa" },
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
    { id: "S06_ExigeMateriaisDisponibilizados", tipo: "parametro", input: "select",
      pergunta: "A execução exige que a Contratada disponibilize materiais, equipamentos, ferramentas ou utensílios próprios?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "S06_MateriaisDisponibilizados", tipo: "especificacao", input: "textarea",
      pergunta: "Liste os materiais, equipamentos, ferramentas e utensílios e as quantidades/qualidades exigidas",
      mostrarSe: { S06_ExigeMateriaisDisponibilizados: "Sim" } },
  ]},
  { id: "07", titulo: "DO INSTRUMENTO CONTRATUAL", campos: [
    { id: "S07_TipoInstrumentoContratual", tipo: "parametro", input: "select",
      pergunta: "O instrumento contratual será",
      opcoes: ["Contrato Simplificado", "Contrato"],
      blocoTexto: true, padrao: "Contrato Simplificado",
      ajuda: "Art. 435 do RLC. Contrato Simplificado é mais enxuto e traz as cláusulas de vinculação, vedações, legislação aplicável e foro dentro do próprio Termo de Referência (Seções 21 a 25); o Contrato completo (Anexo VI) já as contém, então essas seções somem do TR.",
      blocosAdicionais: [
        "S07_TipoInstrumentoContratual_SECOES21A25",
        "S07_TipoInstrumentoContratual_ANEXO",
      ] },
  ]},
  { id: "08", titulo: "DA GARANTIA CONTRATUAL", campos: [
    { id: "S08_ExigeGarantia", tipo: "parametro", input: "select",
      pergunta: "Será exigida garantia contratual?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não" },
    { id: "S08_PercentualGarantia", tipo: "parametro", input: "numero",
      pergunta: "Percentual da garantia contratual (máx. 5%, art. 439 RLC)",
      mostrarSe: { S08_ExigeGarantia: "Sim" } },
    { id: "S08_BaseCalculoGarantia", tipo: "parametro", input: "select",
      pergunta: "A garantia incide sobre",
      opcoes: ["valor do Contrato", "valor anual do Contrato"],
      mostrarSe: { S08_ExigeGarantia: "Sim" } },
  ]},
  { id: "09", titulo: "DO ACOMPANHAMENTO E DA FISCALIZAÇÃO", campos: [
    { id: "S09_TipoFiscalizacao", tipo: "parametro", input: "select",
      pergunta: "A fiscalização será exercida por",
      opcoes: ["Fiscal Funcional", "Comissão de Fiscalização"],
      blocoTexto: true, padrao: "Fiscal Funcional",
      ajuda: "Comissão de Fiscalização reúne Fiscal Funcional, Técnico e Setorial (arts. 537 e 543 do RLC) e muda a redação da Seção 12 (Pagamento)." },
    { id: "S09_RotinasFiscalizacaoEspecificas", tipo: "especificacao", input: "textarea",
      pergunta: "Há rotinas de fiscalização próprias do objeto, além das genéricas já previstas nesta seção?",
      ajuda: "Pode deixar em branco se não houver rotina adicional." },
  ]},
  { id: "10", titulo: "DAS OBRIGAÇÕES DA CONTRATANTE", campos: [] },
  { id: "11", titulo: "DAS OBRIGAÇÕES DA CONTRATADA", campos: [
    { id: "S11_NaturezaIntelectual", tipo: "parametro", input: "select",
      pergunta: "O serviço é de natureza intelectual, estratégico, ou envolve desenvolvimento de produto/projeto?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Se Sim, o documento inclui cessão de direitos autorais/patrimoniais e transição de conhecimento e tecnologia (art. 17 e art. 525 do RLC)." },
  ]},
  { id: "12", titulo: "DO PAGAMENTO", campos: [] },
  { id: "13", titulo: "DO REAJUSTE", campos: [
    { id: "S13_Reajustavel", tipo: "parametro", input: "select",
      pergunta: "Os preços serão reajustáveis após um ano?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não" },
    { id: "S13_IndiceReajuste", tipo: "parametro", input: "texto",
      pergunta: "Índice de referência para o reajuste",
      mostrarSe: { S13_Reajustavel: "Sim" } },
  ]},
  { id: "14", titulo: "DA SUBCONTRATAÇÃO", campos: [
    { id: "S14_PermiteSubcontratacao", tipo: "parametro", input: "select",
      pergunta: "Será admitida subcontratação parcial do objeto?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não" },
    { id: "S14_PercentualMinSubcontratacao", tipo: "parametro", input: "numero",
      pergunta: "Percentual mínimo de subcontratação (%)",
      mostrarSe: { S14_PermiteSubcontratacao: "Sim" } },
    { id: "S14_PercentualMaxSubcontratacao", tipo: "parametro", input: "numero",
      pergunta: "Percentual máximo de subcontratação (%)",
      mostrarSe: { S14_PermiteSubcontratacao: "Sim" } },
    { id: "S14_DetalheSubcontratacao", tipo: "especificacao", input: "textarea",
      pergunta: "Detalhe os percentuais, condições e parâmetros da subcontratação parcial",
      mostrarSe: { S14_PermiteSubcontratacao: "Sim" } },
  ]},
  { id: "15", titulo: "DA ALTERAÇÃO SUBJETIVA", campos: [] },
  { id: "16", titulo: "DAS SANÇÕES ADMINISTRATIVAS", campos: [
    { id: "S16_MultaNaoRegularizacaoHabilitacao", tipo: "parametro", input: "numero",
      pergunta: "Multa por não regularização da documentação de habilitação (% sobre o valor estimado)" },
    { id: "S16_MultaCondutasArt576a580", tipo: "parametro", input: "numero",
      pergunta: "Multa pelas condutas dos art. 576 a 580 do RLC (% sobre o valor estimado)" },
    { id: "S16_MultaRecusaAssinarContrato", tipo: "parametro", input: "numero",
      pergunta: "Multa pela recusa em assinar o contrato (% sobre o valor da contratação)" },
    { id: "S16_MultaCompensatoriaTotal", tipo: "parametro", input: "numero",
      pergunta: "Multa compensatória por inexecução total do contrato (%)" },
    { id: "S16_MultaRescisoria", tipo: "parametro", input: "numero",
      pergunta: "Multa rescisória por rescisão unilateral (% sobre o valor total do contrato)" },
    { id: "S16_BaseCalculoMulta", tipo: "parametro", input: "select",
      pergunta: "As multas moratória e compensatória por inexecução parcial incidem sobre",
      opcoes: ["valor do Contrato", "valor anual do Contrato", "valor mensal do Contrato"],
      ajuda: "Usado nas alíneas e/f/g e na Tabela 2 (grau → percentual).",
      duplicarPara: ["S16_BaseCalculoMulta_E", "S16_BaseCalculoMulta_F", "S16_BaseCalculoMulta_G", "S16_BaseCalculoMulta_TABELA2"] },
    { id: "S16_InfracoesExtras", tipo: "parametro", input: "tabela-linhas",
      colunas: [
        { id: "descricao", rotulo: "Descrição da conduta", input: "texto" },
        { id: "grau", rotulo: "Grau", input: "select", opcoes: ["1", "2", "3", "4", "5"] },
      ],
      pergunta: "Infrações adicionais à Tabela 1 (além das 9 já previstas na minuta)" },
  ]},
  { id: "17", titulo: "DA RESCISÃO CONTRATUAL", campos: [] },
  { id: "18", titulo: "DA VEDAÇÃO AO NEPOTISMO", campos: [] },
  { id: "19", titulo: "DA SUSTENTABILIDADE AMBIENTAL", campos: [
    { id: "S19_CriteriosEspecificos", tipo: "especificacao", input: "textarea",
      pergunta: "Há critérios de sustentabilidade específicos para este objeto, além dos genéricos do art. 10 do RLC?",
      ajuda: "Pode deixar em branco se não houver critério adicional." },
  ]},
  { id: "20", titulo: "DA MATRIZ DE RISCOS E RESPONSABILIDADES", campos: [
    { id: "S20_ExigeMatrizRiscos", tipo: "parametro", input: "select",
      pergunta: "A Matriz de Riscos se aplica a esta contratação?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Sim",
      ajuda: "Marque \"Não\" apenas para serviços de pronto pagamento, sem risco nas fases de planejamento, seleção e gestão do contrato — dispensa a apresentação do Anexo I (Matriz de Riscos).",
      blocosAdicionais: ["S20_ExigeMatrizRiscos_ANEXO_I"] },
  ]},
  { id: "21", titulo: "DA ALTERAÇÃO DO CONTRATO", campos: [] },
  { id: "22", titulo: "DAS VEDAÇÕES", campos: [] },
  { id: "23", titulo: "DA VINCULAÇÃO E DA MANUTENÇÃO DA QUALIFICAÇÃO E HABILITAÇÃO", campos: [] },
  { id: "24", titulo: "DA LEGISLAÇÃO APLICÁVEL", campos: [] },
  { id: "25", titulo: "DO FORO", campos: [] },
  { id: "26", titulo: "DAS DISPOSIÇÕES FINAIS", campos: [
    { id: "S26_ElabDemandanteNome", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Demandante: nome" },
    { id: "S26_ElabDemandanteSetor", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Demandante: setor/gerência" },
    { id: "S26_ElabDemandanteCargo", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Demandante: cargo" },
    { id: "S26_ElabTecnicaNome", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Técnica: nome" },
    { id: "S26_ElabTecnicaSetor", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Técnica: setor/gerência" },
    { id: "S26_ElabTecnicaCargo", tipo: "parametro", input: "texto", pergunta: "Elaborador — Área Técnica: cargo" },
    { id: "S26_AprovDemandanteNome", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Demandante: nome" },
    { id: "S26_AprovDemandanteSetor", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Demandante: setor/gerência" },
    { id: "S26_AprovTecnicaNome", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Técnica: nome" },
    { id: "S26_AprovTecnicaSetor", tipo: "parametro", input: "texto", pergunta: "Aprovador — Área Técnica: setor/gerência" },
    { id: "S26_SuperDemandanteNome", tipo: "parametro", input: "texto", pergunta: "Superintendente/Gerente — Área Demandante: nome" },
    { id: "S26_SuperTecnicaNome", tipo: "parametro", input: "texto", pergunta: "Superintendente/Gerente — Área Técnica: nome" },
  ]},
];
```

- [ ] **Step 4: Rodar e confirmar que passa**

Run: `node --test tests/dados-tr-dispensa.test.mjs`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add js/dados-tr-dispensa.js tests/dados-tr-dispensa.test.mjs
git commit -m "Add TR-Dispensa data model (26 sections, transcribed from SEI 55379661)"
```

---

### Task 3: `templates/build_template_tr_dispensa.py` e `Template_TR_Dispensa.docx`

**Files:**
- Create: `templates/build_template_tr_dispensa.py`
- Create (gerado pelo script): `templates/Template_TR_Dispensa.docx`

**Interfaces:**
- Consumes: os IDs de campo de `js/dados-tr-dispensa.js` (Task 2), mais os
  bookmarks sintéticos `_BLOCO_<VARIANTE>` (incluindo os de
  `blocosAdicionais`), o `_LINHA_MODELO` da Tabela 1 (mesma convenção da
  Task 3 do plano do TR-Serviços), e os pontos de duplicação
  (`S08_BaseCalculoGarantia_74`, `S16_BaseCalculoMulta_E/_F/_G/_TABELA2`,
  seguindo o mesmo padrão já usado no TR-Serviços).
- Produces: `templates/Template_TR_Dispensa.docx`.

- [ ] **Step 1: Escrever `templates/build_template_tr_dispensa.py`**

Reaproveita os helpers de `build_template_tr_servicos.py`
(`add_heading`, `add_paragraph_with_bookmarks`, `normalizar_variante`,
`add_bloco_variante`, `add_linha_modelo_tabela`) e adiciona dois helpers
novos para blocos com tabela no meio (Seções 21–25 combinadas e os dois
anexos de contrato), que `add_bloco_variante` não cobre (ela só lida com
listas de parágrafos de texto puro):

```python
"""Gera Template_TR_Dispensa.docx a partir do texto oficial da minuta
SEI 55379661, com bookmarks nos pontos de preenchimento identificados em
js/dados-tr-dispensa.js. Rodar novamente sempre que o texto oficial mudar."""
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
    """variantes: dict {valor_opcao: [paragrafo1, paragrafo2, ...]}."""
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


def add_bloco_variante_sentinela_inicio(nome_bookmark):
    """Para blocos que misturam parágrafos e tabelas (Seções 21-25
    combinadas, e cada um dos dois anexos de contrato): marca só o INÍCIO
    do bloco num parágrafo vazio. Combinar com
    add_bloco_variante_sentinela_fim(mesmo nome_bookmark) depois de todo o
    conteúdo (parágrafos, tabelas, o que for) já ter sido adicionado ao
    documento. O JS remove tudo entre as duas sentinelas por recorte de
    texto bruto — não importa o que tem no meio."""
    p = doc.add_paragraph()
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start.set(docx.oxml.ns.qn("w:id"), bid)
    start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
    p._p.append(start)
    return bid


def add_bloco_variante_sentinela_fim(bid):
    p = doc.add_paragraph()
    end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end.set(docx.oxml.ns.qn("w:id"), bid)
    p._p.append(end)


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
    ", conforme especificações, condições, quantidades e exigências estabelecidas neste instrumento, por meio de contratação direta por dispensa de licitação, com fundamento legal no art. 416, inciso II, do RLC.",
])
add_paragraph_with_bookmarks([
    "1.2. O custo total estimado para a prestação dos serviços deste Termo de Referência é de R$ ",
    ("S01_ValorEstimado", "________"),
    ". O critério utilizado para a formação da estimativa foi de ",
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
])
add_paragraph_with_bookmarks([
    "1.3. O critério de julgamento das propostas será o de ",
    ("S01_CriterioJulgamento", "[menor preço ou maior desconto]"), " ofertado.",
])
add_paragraph_with_bookmarks([
    "1.4. Os recursos orçamentários decorrentes da contratação do objeto deste Termo de Referência "
    "estão consignados no Orçamento da Conab para o ano 20XX, conforme segue: ND: (",
    ("S01_ND", "XXXX"), "), PTRES: (", ("S01_PTRES", "XXXX"), "), FONTE: (", ("S01_Fonte", "XXXX"), ").",
])

# 2. DA JUSTIFICATIVA
add_heading("2. DA JUSTIFICATIVA")
add_paragraph_with_bookmarks(["2.1. ", ("S02_Motivacao", "(motivação da contratação)")])
doc.add_paragraph(
    "2.2. O serviço em apreço será contratado por meio de dispensa de licitação, com fundamento "
    "legal no art. 416, inciso II, do Regulamento de Licitações e Contratos da Conab.")
add_bloco_variante("S02_ExcecaoExclusividadeMEEPP", {
    "Sim": [
        "2.3. Em atendimento à Lei Complementar 123/2016 e ao Decreto 8.538/2015, que regulamentam "
        "o tratamento favorecido, diferenciado e simplificado para as microempresas e empresas de "
        "pequeno porte, este processo será destinado exclusivamente à participação de "
        "microempresas e empresas de pequeno porte, tendo em vista o valor estimado da contratação "
        "ser de até R$ 80.000,00 (oitenta mil reais), estabelecido em lei.",
    ],
    "Não": [
        "2.3. Em que pese o valor estimado da contratação seja até R$ 80.000,00 (oitenta mil "
        "reais), a dispensa de licitação não será destinada exclusivamente à participação de "
        "microempresas e empresas de pequeno porte, na forma da Lei Complementar 123/2016 e do "
        "inciso III do art. 10 do Decreto 8.538/2015, em razão da contratação em apreço se "
        "enquadrar na exceção prevista no art. 10 do Decreto 8.538/2015, isto é: ",
    ],
})
add_paragraph_with_bookmarks(["", ("S02_MotivoExcecaoExclusividade", "")])

# 3. DO OBJETIVO
add_heading("3. DO OBJETIVO")
add_paragraph_with_bookmarks(["3.1. ", ("S03_Objetivo", "(objetivo da contratação)")])

# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_bloco_variante("S04_TipoVigencia", {
    "Até 1 ano, prorrogável": [
        "4.1. O prazo de vigência do Contrato será de ____ meses (duração do Contrato, conforme "
        "arts. 461 a 463 do RLC), contados da data de sua assinatura, prorrogável na forma dos "
        "arts. 497 e 498 do RLC.",
    ],
    "Plurianual (até 5 anos)": [
        "4.1. O prazo de vigência do Contrato será de ____ anos (duração do Contrato, conforme "
        "arts. 461 a 462 do RLC), contados da data de sua assinatura, podendo ser prorrogado até "
        "o limite de 05 (cinco) anos, desde que observados os requisitos previstos no art. 488 "
        "do RLC.",
    ],
})
add_paragraph_with_bookmarks(["Prazo de vigência (meses): ", ("S04_PrazoVigenciaMeses", "____")])
add_paragraph_with_bookmarks(["Prazo de vigência (anos): ", ("S04_PrazoVigenciaAnos", "____")])

# 5. DA SELEÇÃO DO FORNECEDOR
add_heading("5. DA SELEÇÃO DO FORNECEDOR")
doc.add_paragraph("5.1. DA APRESENTAÇÃO DA PROPOSTA")
add_paragraph_with_bookmarks([
    "5.1.1. A proponente deverá enviar a sua proposta à Conab até ",
    ("S05_PrazoLimiteEnvioProposta", "(data e horário)"),
    " contendo: descrição detalhada do objeto; preço unitário e global, em algarismos e por "
    "extenso; prazo de validade da proposta, não inferior a 60 (sessenta) dias corridos; "
    "informação de que os preços cotados incluem todos os custos necessários (impostos, tributos, "
    "encargos trabalhistas, previdenciários, comerciais e fiscais); e os dados da proponente "
    "(razão social, endereço, telefone, CNPJ, dados bancários).",
])
doc.add_paragraph(
    "5.1.2. A proposta de preços deverá ser redigida em língua portuguesa, sem emendas, rasuras "
    "ou entrelinhas, devidamente datada e assinada pelo representante legal da proponente. Todas "
    "as especificações do objeto contidas na proposta vinculam a Contratada.")
doc.add_paragraph("5.2. DA HABILITAÇÃO E QUALIFICAÇÃO TÉCNICA")
doc.add_paragraph(
    "5.2.1. Para a habilitação, a proponente detentora da melhor oferta deverá apresentar toda a "
    "documentação relacionada a seguir. A proponente já regularmente cadastrada e habilitada "
    "parcialmente no Sistema de Cadastramento Unificado de Fornecedores (SICAF) ficará dispensada "
    "de apresentar os documentos correlatos, quando os registros no SICAF estiverem válidos, "
    "completos e regulares.")
doc.add_paragraph(
    "5.2.2. Relativos à Habilitação Jurídica: os documentos constitutivos compatíveis com a "
    "natureza jurídica da proponente (inscrição no Registro Público de Empresas Mercantis para o "
    "empresário individual; Certificado da Condição de Microempreendedor Individual para o MEI; "
    "ato constitutivo, estatuto ou contrato social em vigor, devidamente registrado, para "
    "sociedade empresária, EIRELI ou sociedade por ações; inscrição do ato constitutivo para "
    "sociedade simples; ata de fundação e estatuto social para cooperativa; decreto de "
    "autorização para empresa estrangeira em funcionamento no País), acompanhados de todas as "
    "alterações ou da consolidação respectiva.")
add_paragraph_with_bookmarks([
    "5.2.2.1. Quando a atividade relativa ao objeto exigir registro ou autorização de "
    "funcionamento de órgão regulador (ANVISA, ANP, Ministério da Agricultura, Polícia Federal, "
    "IBAMA etc.): ", ("S05_RegistroEspecialAtividade", ""),
])
doc.add_paragraph(
    "5.2.3. Relativos à Regularidade Fiscal e Trabalhista: prova de inscrição no CNPJ; prova de "
    "regularidade fiscal perante a Fazenda Nacional (RFB e PGFN); prova de inscrição no cadastro "
    "de contribuintes estadual ou municipal, quando houver; prova de regularidade perante a "
    "Fazenda Estadual e Municipal; prova de regularidade perante o INSS e o FGTS; e prova de "
    "inexistência de débitos inadimplidos perante a Justiça do Trabalho (CNDT). Havendo alguma "
    "restrição na comprovação da regularidade fiscal de microempresa, empresa de pequeno porte ou "
    "cooperativa, será assegurado prazo de 5 (cinco) dias úteis, prorrogável por igual período, "
    "para regularização, sob pena de decadência do direito à contratação.")
doc.add_paragraph(
    "5.2.4. Relativos à Qualificação Econômico-Financeira: certidão negativa de feitos sobre "
    "falência, expedida pelo distribuidor da sede da proponente, com prazo máximo de 60 "
    "(sessenta) dias quando não mencionar validade explícita; e balanço patrimonial e "
    "demonstrações contábeis do último exercício social, vedada a substituição por balancetes ou "
    "balanços provisórios. Será considerada boa a situação financeira da proponente quando os "
    "Índices de Liquidez Geral (LG), Solvência Geral (SG) e Liquidez Corrente (LC) forem maiores "
    "que 1 (um), apurados pelas fórmulas: LG = (Ativo Circulante + Realizável a Longo Prazo) / "
    "(Passivo Circulante + Exigível a Longo Prazo); SG = Ativo Total / (Passivo Circulante + "
    "Exigível a Longo Prazo); LC = Ativo Circulante / Passivo Circulante. A proponente com "
    "qualquer desses índices igual ou inferior a 1 (um) deverá comprovar patrimônio líquido "
    "mínimo de 10% (dez por cento) do valor total estimado da contratação.")
add_paragraph_with_bookmarks([
    "5.2.5. Relativos à Qualificação Técnica: atestado de capacidade técnica que comprove a "
    "execução de serviços com características, quantidades e prazos compatíveis com o objeto, "
    "por período não inferior a ", ("S05_PrazoExperienciaMinimaAnos", "____"),
    " anos, fornecido por pessoas jurídicas de direito público ou privado. Para fins de "
    "comprovação de compatibilidade: ",
    ("S05_QuantitativoServicoPrincipal", "(quantitativo mínimo e serviço(s) principal(is))"),
])
doc.add_paragraph(
    "5.2.6. Será aceito o somatório de atestados para fins de comprovação de capacidade técnica. "
    "A proponente vencedora, apenas por ocasião da assinatura do Contrato, deverá comprovar "
    "inscrição na entidade profissional competente, quando for o caso, e apresentar declaração de "
    "que é Microempresa, Empresa de Pequeno Porte ou Cooperativa, se for o caso.")
doc.add_paragraph(
    "5.2.7. Como condição prévia ao exame da documentação de habilitação, será verificado o "
    "eventual descumprimento das condições de participação mediante consulta ao SICAF, ao "
    "Cadastro Nacional de Empresas Inidôneas e Suspensas (CEIS), ao Cadastro Nacional de "
    "Condenações Cíveis por Atos de Improbidade Administrativa e à Lista de Inidôneos do TCU, em "
    "nome da proponente e de seu sócio majoritário.")
doc.add_paragraph(
    "5.2.8. Será inabilitada a proponente que não comprovar sua habilitação, ressalvado o "
    "saneamento de erros ou falhas que não alterem a substância das propostas e dos documentos, "
    "mediante despacho fundamentado. Será considerada vencedora a proponente que oferecer o menor "
    "preço ou maior desconto e atender às exigências de habilitação deste Termo de Referência.")
doc.add_paragraph("5.3. DA VISTORIA")
add_bloco_variante("S05_TipoVistoria", {
    "Obrigatória": [
        "5.3.1. Para o correto dimensionamento e elaboração de sua proposta, o proponente deverá "
        "realizar vistoria nas instalações do local de execução dos serviços, acompanhado por "
        "empregado designado para esse fim, devendo o agendamento ser efetuado previamente, "
        "podendo sua realização ser comprovada por declaração assinada pelo proponente de que "
        "realizou a vistoria e tem pleno conhecimento das condições e peculiaridades do local, na "
        "forma do Anexo III deste TR.",
    ],
    "Facultativa": [
        "5.3.1. Para o correto dimensionamento e elaboração de sua proposta, o proponente poderá "
        "realizar vistoria nas instalações do local de execução dos serviços, acompanhado por "
        "empregado designado para esse fim, devendo o agendamento ser efetuado previamente, "
        "podendo sua realização ser comprovada por declaração assinada pelo proponente de que "
        "realizou a vistoria e tem pleno conhecimento das condições e peculiaridades do local, na "
        "forma do Anexo III deste TR.",
        "O proponente que optar por não realizar a vistoria deverá apresentar declaração de que "
        "tem pleno conhecimento das condições e peculiaridades do local, assumindo total "
        "responsabilidade por sua opção, na forma do Anexo III deste TR.",
    ],
})
add_paragraph_with_bookmarks([
    "5.3.1.1. O agendamento da vistoria deverá ser feito de segunda a sexta-feira, das ",
    ("S05_HorarioVistoriaInicio", "____"), " às ", ("S05_HorarioVistoriaFim", "____"),
    " horas, pelo telefone (", ("S05_TelefoneAgendamentoVistoria", "____"), ").",
])
doc.add_paragraph("5.3.2. O prazo para vistoria será igual ao prazo estabelecido para envio das propostas.")

# 6. DA DESCRIÇÃO DOS SERVIÇOS
add_heading("6. DA DESCRIÇÃO DOS SERVIÇOS")
doc.add_paragraph(
    "6.1. A prestação dos serviços não gera vínculo empregatício entre os empregados da "
    "Contratada e a Conab, vedando-se qualquer relação que caracterize pessoalidade e "
    "subordinação direta.")
add_paragraph_with_bookmarks(["6.2. ", ("S06_DetalhamentoServicos", "")])
add_paragraph_with_bookmarks([
    "6.3. Os serviços objeto deste Termo de Referência deverão ser realizados em ",
    ("S06_LocalExecucao", "(local)"), ", de segunda a sexta-feira, no horário de ",
    ("S06_HorarioExecucao", "(horário)"), ".",
])
add_paragraph_with_bookmarks([
    "6.4. O prazo de execução dos serviços será de ", ("S06_PrazoExecucaoDias", "____"),
    " dias, contados a partir da assinatura do Contrato.",
])
add_paragraph_with_bookmarks(["6.5. Rotina de execução: ", ("S06_RotinaExecucao", "")])
add_bloco_variante("S06_ExigeGarantiaServicos", {
    "Sim": ["6.6. Os serviços deverão ter um período de garantia de, no mínimo, ____ meses, "
            "contados do recebimento definitivo pela Contratante."],
    "Não": ["6.6. Não será exigida garantia dos serviços na contratação em apreço."],
})
add_paragraph_with_bookmarks(["Período de garantia (meses): ", ("S06_GarantiaServicosMeses", "____")])
add_bloco_variante("S06_ExigeMateriaisDisponibilizados", {
    "Sim": ["6.7. Para a perfeita execução dos serviços, a Contratada deverá disponibilizar os "
            "materiais, equipamentos, ferramentas e utensílios necessários, promovendo sua "
            "substituição quando necessário."],
    "Não": [""],
})
add_paragraph_with_bookmarks(["", ("S06_MateriaisDisponibilizados", "")])
doc.add_paragraph(
    "6.8. Durante a vigência do Contrato, a Conab, para avaliar a qualidade na execução dos "
    "serviços, adotará o Instrumento de Medição de Resultados (IMR), constante no Anexo IV deste "
    "Termo de Referência.")
doc.add_paragraph(
    "6.9. O recebimento provisório será realizado pela fiscalização; o recebimento definitivo, "
    "pelo empregado ou Comissão designada, nos termos dos arts. 526 e 527 do RLC.")

# 7. DO INSTRUMENTO CONTRATUAL
add_heading("7. DO INSTRUMENTO CONTRATUAL")
doc.add_paragraph(
    "7.1. As obrigações decorrentes da presente dispensa de licitação serão formalizadas por meio "
    "de contrato simplificado ou contrato, conforme escolha consignada neste Termo de Referência, "
    "celebrado entre a Conab e a proponente vencedora, observando os termos do RLC.")
doc.add_paragraph(
    "7.2. A proponente vencedora será convocada para, no prazo de até 05 (cinco) dias úteis, "
    "celebrar o instrumento contratual, mediante prévia comprovação das condições de habilitação "
    "consignadas neste TR, mantidas durante a vigência contratual. Antes da contratação será "
    "feita consulta ao CADIN, ao SICAF, ao CEIS, à CNDT, ao Cadastro Nacional de Condenações "
    "Cíveis por Atos de Improbidade Administrativa e à Lista de Inidôneos do TCU.")

# 8. DA GARANTIA CONTRATUAL
add_heading("8. DA GARANTIA CONTRATUAL")
add_bloco_variante("S08_ExigeGarantia", {
    "Sim": [
        "8.1. A Contratada deverá apresentar, no prazo de 10 (dez) dias corridos após a "
        "assinatura do Contrato, prorrogáveis por igual período, comprovante de prestação de "
        "garantia correspondente a ____% (____ por cento), conforme parâmetros do art. 439 do "
        "RLC, do ", ("S08_BaseCalculoGarantia", "(valor do Contrato ou valor anual do Contrato)"),
        ", mediante caução em dinheiro, seguro-garantia ou fiança bancária.",
    ],
    "Não": ["8.1. Não será exigida garantia contratual, para a prestação dos serviços em apreço."],
})
add_paragraph_with_bookmarks(["Percentual da garantia: ", ("S08_PercentualGarantia", "____"), "%"])
doc.add_paragraph(
    "8.2. A inobservância do prazo fixado para apresentação da garantia acarretará multa de "
    "0,08% (oito centésimos por cento) do valor do Contrato por dia de atraso, até o máximo de 2% "
    "(dois por cento). O atraso superior a 25 (vinte e cinco) dias autoriza a rescisão do "
    "Contrato, conforme o inciso I do art. 569 do RLC. A garantia deverá ter validade de 90 "
    "(noventa) dias após o término da vigência do Contrato.")

# 9. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO
add_heading("9. DO ACOMPANHAMENTO E DA FISCALIZAÇÃO")
doc.add_paragraph(
    "9.1. Para a execução do ajuste, será adotado o método de trabalho baseado no conceito de "
    "delegação de responsabilidade. Nos termos dos arts. 543 e 544 do RLC, será designado Fiscal, "
    "seu substituto, ou Comissão de Fiscalização para acompanhar e fiscalizar a execução dos "
    "serviços, avaliando-a constantemente por meio do Instrumento de Medição de Resultado (IMR), "
    "nos termos dos arts. 545 a 548 do RLC.")
add_paragraph_with_bookmarks(["9.2. Rotinas de fiscalização específicas: ", ("S09_RotinasFiscalizacaoEspecificas", "")])

# 10. DAS OBRIGAÇÕES DA CONTRATANTE (texto fixo)
add_heading("10. DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "10.1. São obrigações da Contratante: exigir o cumprimento das obrigações da Contratada; "
    "receber o objeto no prazo e condições estabelecidas; verificar a conformidade dos serviços "
    "recebidos provisoriamente; comunicar à Contratada sobre imperfeições, falhas ou "
    "irregularidades; rejeitar serviço em desacordo; acompanhar e fiscalizar o cumprimento das "
    "obrigações da Contratada; e efetuar o pagamento no prazo e forma estabelecidos. A Conab não "
    "responderá por compromissos assumidos pela Contratada com terceiros.")

# 11. DAS OBRIGAÇÕES DA CONTRATADA
add_heading("11. DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph(
    "11.1. A Contratada deve cumprir todas as obrigações constantes neste Termo, seus anexos e "
    "sua proposta, assumindo como exclusivamente seus os riscos e as despesas decorrentes da boa "
    "e perfeita execução do objeto, incluindo: executar os serviços com a alocação dos "
    "empregados necessários; responsabilizar-se pelos vícios e danos do objeto (CDC); substituir, "
    "reparar ou corrigir às suas expensas os serviços com vícios; comunicar atrasos no prazo "
    "máximo de 24 horas; utilizar empregados habilitados, uniformizados e identificados; "
    "responsabilizar-se pelas obrigações trabalhistas, sociais, previdenciárias e tributárias; "
    "substituir empregados a pedido da fiscalização; instruir seus empregados quanto às normas "
    "internas e às atividades contratadas; relatar irregularidades; não permitir trabalho de "
    "menores; guardar sigilo; arcar com equívocos de dimensionamento de sua proposta; manter as "
    "condições de habilitação; prestar esclarecimentos; indicar preposto; e deter instalações, "
    "aparelhamento e pessoal técnico adequados.")
add_bloco_variante("S11_NaturezaIntelectual", {
    "Sim": [
        "11.2. Ceder os direitos patrimoniais relativos ao projeto ou serviço técnico "
        "especializado, para que a Conab possa utilizá-lo, de acordo com o art. 525 do RLC. "
        "Quando o projeto referir-se a obra imaterial de caráter tecnológico, insuscetível de "
        "privilégio, a cessão dos direitos incluirá o fornecimento de todos os dados, documentos "
        "e elementos de informação pertinentes à tecnologia de concepção, desenvolvimento, "
        "fixação em suporte físico e aplicação da obra.",
        "11.3. Assegurar à Contratante o direito de propriedade intelectual dos produtos "
        "desenvolvidos, inclusive sobre adequações e atualizações futuras, de forma permanente, "
        "e os direitos autorais da solução, do projeto, de suas especificações técnicas e da "
        "documentação produzida, inclusive os produzidos por terceiros subcontratados, ficando "
        "proibida sua utilização sem autorização expressa da Contratante, sob pena de multa, sem "
        "prejuízo das sanções civis e penais cabíveis.",
        "11.4. Realizar a transição contratual com transferência de conhecimento, tecnologia e "
        "técnicas empregadas, sem perda de informações, podendo exigir, inclusive, a capacitação "
        "dos técnicos da Contratante ou da nova empresa que continuará a execução dos serviços, "
        "nos termos do art. 17 do RLC.",
    ],
    "Não": [""],
})

# 12. DO PAGAMENTO
add_heading("12. DO PAGAMENTO")
doc.add_paragraph(
    "12.1. O pagamento será efetuado pela Conab no prazo de até 10 (dez) dias corridos, contados "
    "do recebimento da Nota Fiscal/Fatura.")
doc.add_paragraph(
    "12.2. A emissão da Nota Fiscal/Fatura será precedida do recebimento provisório e definitivo "
    "do serviço: no prazo de até 05 (cinco) dias corridos do adimplemento da parcela, a "
    "Contratada deverá entregar a documentação comprobatória do cumprimento da obrigação "
    "contratual;")
add_bloco_variante("S09_TipoFiscalizacao", {
    "Fiscal Funcional": [
        "no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada, o Fiscal Funcional deverá realizar a análise de toda a documentação "
        "apresentada, avaliar a execução por meio do Instrumento de Medição de Resultado (IMR) e "
        "emitir Termo de Recebimento Provisório, encaminhando-o ao empregado ou comissão "
        "designada para o recebimento definitivo.",
    ],
    "Comissão de Fiscalização": [
        "no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da "
        "Contratada: o Fiscal Funcional deverá realizar a análise de toda a documentação "
        "apresentada e emitir relatório; o Fiscal Técnico deverá realizar a análise de toda a "
        "documentação apresentada e emitir relatório; o Fiscal Setorial deverá realizar a análise "
        "de toda a documentação apresentada e emitir relatório;",
        "o Fiscal Funcional, no prazo de até 03 (três) dias úteis a partir do recebimento do "
        "último relatório, avaliará a execução por meio do Instrumento de Medição de Resultado "
        "(IMR) e emitirá Termo de Recebimento Provisório com base nos relatórios emitidos pelos "
        "fiscais, encaminhando-o, caso não haja irregularidades, ao empregado ou comissão "
        "designada para o recebimento definitivo.",
    ],
})
doc.add_paragraph(
    "12.3. Constatadas impropriedades que impeçam a liquidação e o pagamento da despesa, estas "
    "serão registradas no Termo de Recebimento Provisório, com prazo de até 05 (cinco) dias úteis "
    "para regularização; sanadas as impropriedades, será elaborado relatório detalhado e "
    "encaminhado o Termo de Recebimento Provisório ao empregado ou Comissão designada, que "
    "providenciará o recebimento definitivo em até 05 (cinco) dias corridos, emitindo Termo "
    "Circunstanciado e comunicando a Contratada para emissão da Nota Fiscal ou Fatura, com o "
    "valor exato dimensionado pela fiscalização com base no IMR.")
doc.add_paragraph(
    "12.4. Antes de cada pagamento será realizada consulta ao SICAF; havendo resultado "
    "desfavorável, será concedido prazo de 05 (cinco) dias úteis para regularização ou defesa, "
    "prorrogável uma vez por igual período.")

# 13. DO REAJUSTE
add_heading("13. DO REAJUSTE")
add_bloco_variante("S13_Reajustavel", {
    "Sim": ["13.1. O preço consignado no Contrato será corrigido anualmente, observado o "
            "interregno mínimo de um ano, contado da data limite para apresentação da proposta, "
            "pela variação do índice de referência."],
    "Não": ["13.1. O preço é fixo e irreajustável."],
})
add_paragraph_with_bookmarks(["Índice de referência: ", ("S13_IndiceReajuste", "")])

# 14. DA SUBCONTRATAÇÃO
add_heading("14. DA SUBCONTRATAÇÃO")
add_bloco_variante("S14_PermiteSubcontratacao", {
    "Sim": ["14.1. É permitida a subcontratação parcial do objeto entre os limites mínimo e "
            "máximo, respectivamente, do valor total do Contrato."],
    "Não": ["14.1. Não será admitida a subcontratação do objeto desta dispensa de licitação."],
})
add_paragraph_with_bookmarks([
    "Limites: entre ", ("S14_PercentualMinSubcontratacao", "____"), "% e ",
    ("S14_PercentualMaxSubcontratacao", "____"), "% do valor total do Contrato.",
])
add_paragraph_with_bookmarks(["Detalhamento: ", ("S14_DetalheSubcontratacao", "")])
doc.add_paragraph("14.2. A subcontratação deverá observar o disposto no art. 524 do RLC.")

# 15. DA ALTERAÇÃO SUBJETIVA (texto fixo)
add_heading("15. DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "15.1. É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa "
    "jurídica, desde que observados os requisitos de habilitação, mantidas as demais cláusulas e "
    "condições do Contrato, não haja prejuízo à execução do objeto pactuado e haja a anuência "
    "expressa da Conab.")

# 16. DAS SANÇÕES ADMINISTRATIVAS
add_heading("16. DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "16.1. A contratada, em caso de inadimplemento de suas obrigações, garantido o contraditório "
    "e ampla defesa, ficará sujeita às seguintes sanções previstas no RLC e na Lei nº 13.303, de "
    "2016: advertência; multa moratória; multa compensatória; multa rescisória, para os casos de "
    "rescisão unilateral por descumprimento contratual; e suspensão do direito de participar de "
    "licitação e impedimento de contratar com a Conab, por até 02 (dois) anos.")
add_paragraph_with_bookmarks([
    "a) multa por não regularização da documentação de habilitação: ",
    ("S16_MultaNaoRegularizacaoHabilitacao", "____"), "% sobre o valor estimado para a contratação.",
])
add_paragraph_with_bookmarks([
    "b) multa pela prática das condutas elencadas nos arts. 576 a 580 do RLC: ",
    ("S16_MultaCondutasArt576a580", "____"), "% sobre o valor estimado para a contratação.",
])
add_paragraph_with_bookmarks([
    "c) multa pela recusa em assinar o Contrato dentro do prazo estabelecido: ",
    ("S16_MultaRecusaAssinarContrato", "____"), "% sobre o valor da contratação.",
])
doc.add_paragraph("d) multa moratória por atraso injustificado na entrega da garantia contratual, conforme item 8.2.")
add_paragraph_with_bookmarks([
    "e) multa moratória de 0,2% (dois décimos por cento) sobre o ",
    ("S16_BaseCalculoMulta_E", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, até o limite de 15 (quinze) dias.",
])
add_paragraph_with_bookmarks([
    "f) multa moratória de 0,3% (três décimos por cento) sobre o ",
    ("S16_BaseCalculoMulta_F", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", por dia de atraso na execução dos serviços, por período superior ao previsto na alínea "
    "anterior, até o limite de 15 (quinze) dias. Esgotado o prazo, poderá ocorrer a não "
    "aceitação do objeto, sem prejuízo da rescisão unilateral da avença.",
])
add_paragraph_with_bookmarks([
    "g) multa compensatória no percentual de 10% (dez por cento) sobre o ",
    ("S16_BaseCalculoMulta_G", "(valor do Contrato ou valor anual/mensal do Contrato)"),
    ", no caso de inexecução parcial do Contrato.",
])
add_paragraph_with_bookmarks([
    "h) multa compensatória de ", ("S16_MultaCompensatoriaTotal", "____"),
    "% sobre o valor total do Contrato, no caso de inexecução total do Contrato.",
])
add_paragraph_with_bookmarks([
    "i) multa rescisória de ", ("S16_MultaRescisoria", "____"),
    "% sobre o valor total do Contrato, no caso de rescisão contratual unilateral.",
])
add_paragraph_with_bookmarks([
    "j) multa de 0,2% a 3,2% por dia sobre o ",
    ("S16_BaseCalculoMulta_TABELA2", "(valor do Contrato, valor anual ou valor mensal)"),
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
    ("Deixar de cumprir quaisquer itens do Termo de Referência e seus Anexos não previstos nesta "
     "tabela, após reincidência formalmente notificada pela fiscalização, por item e por "
     "ocorrência", "3"),
    ("Deixar de indicar e manter durante a execução do Contrato os prepostos previstos no "
     "Contrato, por dia", "1"),
    ("Deixar de providenciar treinamento para seus funcionários conforme previsto na relação de "
     "obrigações da Contratada, por funcionário", "1"),
]
for descricao, grau in infracoes_fixas:
    linha = tabela1.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
add_linha_modelo_tabela(tabela1, "S16_InfracoesExtras_LINHA_MODELO", ["((DESCRICAO))", "((GRAU))"])

tabela2 = doc.add_table(rows=1, cols=2)
tabela2.style = "Table Grid"
tabela2.rows[0].cells[0].text = "GRAU"
tabela2.rows[0].cells[1].text = "CORRESPONDÊNCIA (percentual por dia)"
for grau, percentual in [("1", "0,2%"), ("2", "0,4%"), ("3", "0,8%"), ("4", "1,6%"), ("5", "3,2%")]:
    linha = tabela2.add_row()
    linha.cells[0].text = grau
    linha.cells[1].text = percentual
doc.add_paragraph(
    "16.2. As multas moratória, compensatória e rescisória possuem fatos geradores distintos; se "
    "aplicadas duas multas sobre o mesmo fato gerador, configura bis in idem.")
doc.add_paragraph(
    "16.3. Da sanção de suspensão: cabe em razão de ação ou omissão capaz de causar, ou que "
    "tenha causado, prejuízo à Conab, suas instalações, pessoas, imagem, meio ambiente, ou em "
    "decorrência de determinação legal, aplicada por até 02 (dois) anos, nos termos dos arts. "
    "579 a 580 do RLC.")

# 17. DA RESCISÃO CONTRATUAL (texto fixo)
add_heading("17. DA RESCISÃO CONTRATUAL")
doc.add_paragraph(
    "17.1. A inexecução total do Contrato ensejará a sua rescisão, enquanto a inexecução parcial "
    "poderá ensejar a sua rescisão, com as consequências cabíveis, conforme disposto nos arts. "
    "568 a 572 do RLC. A rescisão poderá ser por ato unilateral e escrito da Conab, amigável, por "
    "acordo entre as partes, ou judicial.")

# 18. DA VEDAÇÃO AO NEPOTISMO (texto fixo)
add_heading("18. DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "18.1. Conforme disposto no parágrafo único do art. 12 do RLC e no art. 7º do Decreto nº "
    "7.203, de 2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa "
    "física, ou de pessoa/empresa com relação de parentesco até o terceiro grau civil com "
    "autoridade do Ministério da Agricultura, dirigente ou empregado da Conab com atribuições na "
    "área de licitação e contratação.")

# 19. DA SUSTENTABILIDADE AMBIENTAL
add_heading("19. DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "19.1. Compete à Contratada, no que couber, atender os critérios de sustentabilidade "
    "ambiental previstos no art. 10 do RLC.")
doc.add_paragraph(
    "19.2. A Contratada se responsabiliza administrativamente, civilmente e penalmente por "
    "qualquer dano causado pelo seu serviço ao meio ambiente.")
add_paragraph_with_bookmarks(["19.3. ", ("S19_CriteriosEspecificos", "")])

# 20. DA MATRIZ DE RISCOS E RESPONSABILIDADES
add_heading("20. DA MATRIZ DE RISCOS E RESPONSABILIDADES")
add_bloco_variante("S20_ExigeMatrizRiscos", {
    "Sim": [
        "20.1. A Matriz de Riscos é a cláusula contratual definidora dos riscos e das "
        "responsabilidades entre o Contratante e a Contratada e caracterizadora do equilíbrio "
        "econômico financeiro na execução do Contrato, em termos de ônus financeiros decorrentes "
        "de eventos supervenientes à contratação.",
        "20.2. A Contratada é integral e exclusivamente responsável por todos os riscos e "
        "responsabilidades relacionados ao objeto do ajuste, conforme hipóteses não-exaustivas "
        "elencadas na Matriz de Riscos — Anexo I do Termo de Referência, e não é responsável "
        "pelos riscos e responsabilidades quando estes competirem à Contratante, conforme "
        "estabelecido no mesmo Anexo.",
    ],
    "Não": [
        "20.1. A presente contratação é composta por serviços de pronto pagamento, os quais não "
        "apresentam nenhum tipo de risco nas fases de planejamento, seleção e gestão do contrato, "
        "ficando assim dispensada a apresentação de matriz de riscos para este Termo de "
        "Referência.",
    ],
})

# 21. DA ALTERAÇÃO DO CONTRATO / 22. DAS VEDAÇÕES / 23. DA VINCULAÇÃO / 24. LEGISLAÇÃO / 25. FORO
# Bloco inteiro (5 seções, com tabelas nenhuma no meio) só existe quando o
# instrumento contratual é "Contrato Simplificado" — ver Seção 7.
_bid_secoes_21a25 = add_bloco_variante_sentinela_inicio(
    "S07_TipoInstrumentoContratual_SECOES21A25_BLOCO_CONTRATO_SIMPLIFICADO")

add_heading("21. DA ALTERAÇÃO DO CONTRATO")
doc.add_paragraph(
    "21.1. O Contrato Simplificado decorrente deste Termo de Referência poderá ser alterado nas "
    "hipóteses previstas no art. 510 do RLC, vedando-se alterações que resultem em violação ao "
    "dever de licitar. A Contratada poderá aceitar, nas mesmas condições contratuais, os "
    "acréscimos ou supressões que se fizerem necessários, até 25% (vinte e cinco por cento) do "
    "valor inicial atualizado do contrato. Fica vedada a celebração de termos aditivos "
    "decorrentes de eventos supervenientes alocados, na Matriz de Riscos, como de "
    "responsabilidade da Contratada, que somente poderá solicitar reequilíbrio "
    "econômico-financeiro nas hipóteses excluídas de suas responsabilidades.")

add_heading("22. DAS VEDAÇÕES")
doc.add_paragraph(
    "22.1. É vedado à Contratada: caucionar ou utilizar o Contrato Simplificado decorrente deste "
    "Termo de Referência para qualquer operação financeira; interromper a execução contratual "
    "sob alegação de inadimplemento por parte da Contratante, salvo nos casos previstos em lei; "
    "empregar menor de 18 anos em trabalho noturno, perigoso ou insalubre e não empregar menor de "
    "16 anos, salvo na condição de aprendiz a partir de 14 anos; e subcontratar, na íntegra ou "
    "parcialmente, o objeto da contratação.")

add_heading("23. DA VINCULAÇÃO E DA MANUTENÇÃO DA QUALIFICAÇÃO E HABILITAÇÃO")
doc.add_paragraph(
    "23.1. Consideram-se integrantes do Contrato Simplificado o Termo de Referência da Dispensa "
    "de Licitação e seus Anexos, a Proposta da Contratada, no que couber, e demais documentos "
    "pertinentes, independentemente de transcrição. A Contratada se obriga a manter, durante toda "
    "a execução do objeto, as condições de habilitação e qualificação exigidas neste Termo de "
    "Referência.")

add_heading("24. DA LEGISLAÇÃO APLICÁVEL")
doc.add_paragraph(
    "24.1. Os casos omissos serão decididos pela Contratante, segundo as disposições contidas no "
    "RLC, na Lei nº 13.303, de 2016, e demais normas federais de licitações e contratos "
    "administrativos aplicáveis às empresas públicas e, subsidiariamente, segundo as disposições "
    "do Código de Defesa do Consumidor e normas e princípios gerais dos contratos.")

add_heading("25. DO FORO")
doc.add_paragraph(
    "25.1. As partes elegem o foro da Justiça Federal, Seção Judiciária do Estado da Sureg ou de "
    "Brasília-DF, no caso da Matriz, competente para dirimir quaisquer dúvidas suscitadas em "
    "razão do Contrato Simplificado decorrente deste Termo de Referência, que não puderem ser "
    "resolvidas de comum acordo.")

add_bloco_variante_sentinela_fim(_bid_secoes_21a25)

# 26. DAS DISPOSIÇÕES FINAIS
add_heading("26. DAS DISPOSIÇÕES FINAIS")
doc.add_paragraph(
    "26.1. Integram este Termo de Referência os seguintes Anexos: a) Anexo I - Matriz de Riscos "
    "(quando aplicável); b) Anexo II - Modelo de Apresentação de Proposta; c) Anexo III - "
    "Declaração de Vistoria Técnica; d) Anexo IV - Modelo de Instrumento de Medição de Resultados "
    "(IMR); e e) o instrumento contratual (Contrato Simplificado ou Minuta de Contrato, conforme "
    "escolhido na Seção 7).")
add_paragraph_with_bookmarks(["Elaborador — Área Demandante: ", ("S26_ElabDemandanteNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S26_ElabDemandanteSetor", ""), " — Cargo: ", ("S26_ElabDemandanteCargo", "")])
add_paragraph_with_bookmarks(["Elaborador — Área Técnica: ", ("S26_ElabTecnicaNome", "")])
add_paragraph_with_bookmarks(["Setor: ", ("S26_ElabTecnicaSetor", ""), " — Cargo: ", ("S26_ElabTecnicaCargo", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Demandante: ", ("S26_AprovDemandanteNome", ""), " — Setor: ", ("S26_AprovDemandanteSetor", "")])
add_paragraph_with_bookmarks(["Aprovador — Área Técnica: ", ("S26_AprovTecnicaNome", ""), " — Setor: ", ("S26_AprovTecnicaSetor", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Demandante: ", ("S26_SuperDemandanteNome", "")])
add_paragraph_with_bookmarks(["Superintendente/Gerente — Área Técnica: ", ("S26_SuperTecnicaNome", "")])

# ANEXO I - MATRIZ DE RISCOS (só existe se S20_ExigeMatrizRiscos = "Sim")
doc.add_page_break()
_bid_anexo_i = add_bloco_variante_sentinela_inicio("S20_ExigeMatrizRiscos_ANEXO_I_BLOCO_SIM")
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
for _ in range(5):
    tabela_riscos.add_row()
doc.add_paragraph("Escala de Probabilidade: 1 Muito Baixa, 2 Baixa, 3 Média, 4 Alta, 5 Muito Alta.")
doc.add_paragraph("Escala de Impacto: 1 Muito Baixo, 2 Baixo, 3 Médio, 4 Alto, 5 Muito Alto.")
doc.add_paragraph("Nível de Risco: 1 a 2 Baixo; 3 a 6 Médio; 8 a 12 Elevado; 15 a 25 Extremo.")
doc.add_paragraph("Elaborado por: NOME     Supervisionado por: NOME     Autorizado por: NOME")
add_bloco_variante_sentinela_fim(_bid_anexo_i)

# ANEXO II - MODELO DE APRESENTAÇÃO DE PROPOSTA
doc.add_page_break()
add_heading("ANEXO II DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE APRESENTAÇÃO DE PROPOSTA")
tabela_proposta = doc.add_table(rows=1, cols=2)
tabela_proposta.style = "Table Grid"
for celula, texto in zip(tabela_proposta.rows[0].cells, ["Item/Especificações", "Valor Total"]):
    celula.text = texto
for _ in range(4):
    tabela_proposta.add_row()
doc.add_paragraph("VALOR TOTAL DOS ITENS: R$")

# ANEXO III - MODELO DE DECLARAÇÃO DE VISTORIA
doc.add_page_break()
add_heading("ANEXO III DO TERMO DE REFERÊNCIA")
add_heading("MODELO DE DECLARAÇÃO DE VISTORIA")
doc.add_paragraph("À Companhia Nacional de Abastecimento – CONAB")
doc.add_paragraph("Ref.: Dispensa de Licitação – nº ___/20XX")
doc.add_paragraph(
    "Declaramos que, em atendimento ao previsto na Dispensa de Licitação nº ___/_____, eu, "
    "______________________________________, portador(a) da CI/RG nº______________ e do CPF nº "
    "_____________________, responsável da empresa "
    "___________________________________________________, como seu representante legal para os "
    "fins da presente declaração, compareci perante a CONAB/Sureg/XX, situada à "
    "_____________________________________________________________________________________, e "
    "vistoriei os locais onde serão executados os serviços objeto da Dispensa de Licitação em "
    "apreço, tomando plena ciência das condições e graus de dificuldade existentes.")
doc.add_paragraph("Cidade-UF, ___/___/20XX.")
doc.add_paragraph("Assinatura e carimbo (Responsável Técnico da Empresa)")
doc.add_paragraph("Visto: Responsável pela CONAB")

# ANEXO IV - MODELO DE IMR
doc.add_page_break()
add_heading("ANEXO IV DO TERMO DE REFERÊNCIA")
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

# ANEXO V - CONTRATO SIMPLIFICADO (só existe se Contrato Simplificado)
doc.add_page_break()
_bid_anexo_v = add_bloco_variante_sentinela_inicio(
    "S07_TipoInstrumentoContratual_ANEXO_BLOCO_CONTRATO_SIMPLIFICADO")
add_heading("ANEXO V DO TERMO DE REFERÊNCIA")
add_heading("MINUTA DE CONTRATO SIMPLIFICADO")
doc.add_paragraph("O presente CONTRATO SIMPLIFICADO obedecerá aos seguintes termos e condições:")
doc.add_paragraph(
    "1. O Fornecedor deverá acusar, prontamente, o recebimento deste, comprometendo-se a aceitar "
    "os termos e condições do presente Contrato Simplificado e do respectivo Termo de "
    "Referência.")
doc.add_paragraph(
    "2. Concorda o Fornecedor que todos os serviços especificados no presente Contrato "
    "Simplificado são de primeira qualidade, por ele garantidos como tais, responsabilizando-se "
    "pelos vícios e defeitos, visíveis ou ocultos, obrigando-se a refazer os serviços que não "
    "satisfaçam plenamente as especificações apresentadas pela Conab, livre de qualquer despesa "
    "extra, no prazo que lhe for designado.")
doc.add_paragraph(
    "3. O Fornecedor assume toda a responsabilidade por qualquer interpelação envolvendo "
    "patentes ou direitos autorais, correndo por sua conta quaisquer pagamentos ou multas "
    "exigidas por titulares desses direitos.")
doc.add_paragraph(
    "4. A responsabilidade da Conab na execução do presente limita-se e extingue-se ao preço e "
    "seu pagamento, sendo da exclusiva responsabilidade do Fornecedor todas e quaisquer despesas "
    "ou ônus, inclusive trabalhistas e de previdência social, não se compreendendo a operação "
    "como subcontratação ou subempreitada.")
doc.add_paragraph(
    "5. Os serviços serão prestados diretamente à Conab nos locais por ela indicados, "
    "considerando-se realizada a prestação somente após a verificação de que satisfazem os "
    "requisitos e especificações da Conab.")
doc.add_paragraph(
    "6. O prazo de entrega deverá ser rigorosamente observado, sob pena de aplicação das sanções "
    "administrativas previstas no Termo de Referência respectivo.")
doc.add_paragraph("7. O pagamento do presente Contrato Simplificado far-se-á via Ordem Bancária ou Ordem de Pagamento.")
add_bloco_variante_sentinela_fim(_bid_anexo_v)

# ANEXO VI - MINUTA DE CONTRATO (só existe se Contrato completo)
doc.add_page_break()
_bid_anexo_vi = add_bloco_variante_sentinela_inicio(
    "S07_TipoInstrumentoContratual_ANEXO_BLOCO_CONTRATO")
add_heading("ANEXO VI DO TERMO DE REFERÊNCIA")
add_heading("MINUTA DE CONTRATO")
doc.add_paragraph(
    "CONTRATO QUE ENTRE SI CELEBRAM A COMPANHIA NACIONAL DE ABASTECIMENTO - CONAB E A EMPRESA "
    "___________________ PARA A PRESTAÇÃO DE SERVIÇOS DE ___________________, de conformidade "
    "com o Processo Administrativo referente à Dispensa de Licitação, regendo-se pelo Termo de "
    "Referência e seus anexos e pela proposta da Contratada, pelo RLC, pela Lei nº 13.303, de "
    "2016, e demais legislações pertinentes.")
doc.add_paragraph("CLÁUSULA PRIMEIRA - DO OBJETO")
doc.add_paragraph(
    "O objeto do presente instrumento é a contratação de serviços, que serão prestados nas "
    "condições estabelecidas no Termo de Referência, ao qual este Termo de Contrato se vincula, "
    "independentemente de transcrição, juntamente com a proposta vencedora.")
doc.add_paragraph("CLÁUSULA SEGUNDA - DA VIGÊNCIA")
doc.add_paragraph(
    "O prazo de vigência do Contrato é o estabelecido no Termo de Referência, contado da data de "
    "sua assinatura, observadas as condições de prorrogação nele previstas.")
doc.add_paragraph("CLÁUSULA TERCEIRA - DO REGIME DE EXECUÇÃO, DOS PRAZOS DE EXECUÇÃO E DA GARANTIA DOS SERVIÇOS")
doc.add_paragraph(
    "Os serviços serão executados conforme o regime, os prazos de execução e as condições de "
    "garantia estabelecidos no Termo de Referência, admitida prorrogação do prazo de execução "
    "observado o disposto nos arts. 497 e 498 do RLC.")
doc.add_paragraph("CLÁUSULA QUARTA - DO RECEBIMENTO DOS SERVIÇOS")
doc.add_paragraph(
    "O recebimento provisório e definitivo dos serviços observará o disposto no Termo de "
    "Referência e nos arts. 526 e 527 do RLC, não excluindo a responsabilidade da Contratada "
    "pelos prejuízos resultantes da incorreta execução do Contrato.")
doc.add_paragraph("CLÁUSULA QUINTA - DO VALOR DO CONTRATO")
doc.add_paragraph(
    "O valor do Contrato corresponde ao valor proposto pela Contratada vencedora, incluídas todas "
    "as despesas ordinárias diretas e indiretas decorrentes da execução do objeto, inclusive "
    "tributos, encargos sociais, trabalhistas e previdenciários.")
doc.add_paragraph("CLÁUSULA SEXTA - DA GARANTIA CONTRATUAL")
doc.add_paragraph(
    "As condições de garantia contratual, quando exigida, observarão o disposto no Termo de "
    "Referência e no art. 439 do RLC.")
doc.add_paragraph("CLÁUSULA SÉTIMA - DA DOTAÇÃO ORÇAMENTÁRIA")
doc.add_paragraph(
    "A despesa orçamentária da execução deste Contrato correrá à conta da Natureza da Despesa, "
    "PTRES e Fonte indicados no Termo de Referência, conforme Nota de Empenho a ser emitida.")
doc.add_paragraph("CLÁUSULA OITAVA - DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "São obrigações da Contratante as estabelecidas no Termo de Referência, não respondendo a "
    "Conab por compromissos assumidos pela Contratada com terceiros.")
doc.add_paragraph("CLÁUSULA NONA - DAS OBRIGAÇÕES DA CONTRATADA")
doc.add_paragraph("São obrigações da Contratada as estabelecidas no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA - DA SUSTENTABILIDADE AMBIENTAL")
doc.add_paragraph(
    "A Contratada observará os critérios de sustentabilidade ambiental previstos no Termo de "
    "Referência e no art. 10 do RLC, respondendo administrativa, civil e penalmente por danos "
    "causados ao meio ambiente.")
doc.add_paragraph("CLÁUSULA DÉCIMA PRIMEIRA - DA MANUTENÇÃO DA QUALIFICAÇÃO E HABILITAÇÃO")
doc.add_paragraph(
    "A Contratada se obriga a manter, durante toda a execução do objeto, as condições de "
    "habilitação e qualificação exigidas para esta contratação.")
doc.add_paragraph("CLÁUSULA DÉCIMA SEGUNDA - DO ACOMPANHAMENTO E DA FISCALIZAÇÃO DO CONTRATO")
doc.add_paragraph("O acompanhamento e a fiscalização do Contrato observarão o previsto no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA TERCEIRA - DO PAGAMENTO")
doc.add_paragraph("O prazo para pagamento e demais condições a ele referentes se encontram definidos no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA QUARTA - DO REAJUSTE")
doc.add_paragraph("As condições de reajuste, quando aplicáveis, observam o disposto no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA QUINTA - DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "A Contratada, em caso de inadimplemento de suas obrigações, garantido o contraditório e "
    "ampla defesa, ficará sujeita às sanções previstas no RLC, na Lei nº 13.303, de 2016, e no "
    "Termo de Referência: advertência; multa moratória; multa compensatória; multa rescisória; e "
    "suspensão do direito de participar de licitação e impedimento de contratar com a Conab, por "
    "até 02 (dois) anos, aplicadas conforme os percentuais e a tabela de infrações estabelecidos "
    "no Termo de Referência.")
tabela1_anexo_vi = doc.add_table(rows=1, cols=2)
tabela1_anexo_vi.style = "Table Grid"
tabela1_anexo_vi.rows[0].cells[0].text = "TABELA 1 — DESCRIÇÃO DA INFRAÇÃO"
tabela1_anexo_vi.rows[0].cells[1].text = "GRAU"
for descricao, grau in infracoes_fixas:
    linha = tabela1_anexo_vi.add_row()
    linha.cells[0].text = descricao
    linha.cells[1].text = grau
add_linha_modelo_tabela(tabela1_anexo_vi, "S16_InfracoesExtras_LINHA_MODELO", ["((DESCRICAO))", "((GRAU))"])
doc.add_paragraph("CLÁUSULA DÉCIMA SEXTA - DA RESCISÃO")
doc.add_paragraph(
    "A rescisão observará o disposto no Termo de Referência e nos arts. 568 a 572 e 582 a 593 do "
    "RLC.")
doc.add_paragraph("CLÁUSULA DÉCIMA SÉTIMA - DA SUBCONTRATAÇÃO")
doc.add_paragraph("As condições de subcontratação, quando admitidas, observam o disposto no Termo de Referência.")
doc.add_paragraph("CLÁUSULA DÉCIMA OITAVA - DA MATRIZ DE RISCOS E RESPONSABILIDADES")
doc.add_paragraph(
    "A Matriz de Riscos, quando aplicável — Anexo I do Termo de Referência —, constitui peça "
    "integrante deste Contrato, independentemente de transcrição, definindo os riscos e "
    "responsabilidades entre Contratante e Contratada.")
doc.add_paragraph("CLÁUSULA DÉCIMA NONA - DA ALTERAÇÃO DO CONTRATO")
doc.add_paragraph(
    "O presente Contrato poderá ser alterado nas hipóteses do art. 510 do RLC, vedando-se "
    "alterações que resultem em violação ao dever de licitar. A Contratada poderá aceitar "
    "acréscimos ou supressões de até 25% (vinte e cinco por cento) do valor inicial atualizado do "
    "contrato.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA - DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa jurídica, "
    "observados os requisitos de habilitação e a anuência expressa da Conab.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA PRIMEIRA - DAS VEDAÇÕES")
doc.add_paragraph(
    "É vedado à Contratada caucionar ou utilizar este Contrato para qualquer operação financeira; "
    "interromper a execução contratual sob alegação de inadimplemento da Contratante, salvo nos "
    "casos previstos em lei; empregar menor de 18 anos em trabalho noturno, perigoso ou insalubre "
    "e não empregar menor de 16 anos, salvo na condição de aprendiz a partir de 14 anos; e "
    "subcontratar, na íntegra ou parcialmente, o objeto da contratação, salvo quando a "
    "subcontratação parcial for admitida no Termo de Referência.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA SEGUNDA - DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "Conforme disposto no parágrafo único do art. 12 do RLC e no art. 7º do Decreto nº 7.203, de "
    "2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa física, ou de "
    "pessoa/empresa com relação de parentesco até o terceiro grau civil com autoridade do "
    "Ministério da Agricultura, dirigente ou empregado da Conab com atribuições na área de "
    "licitação e contratação.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA TERCEIRA - DA VINCULAÇÃO")
doc.add_paragraph(
    "Consideram-se integrantes do presente Instrumento Contratual o Termo de Referência da "
    "Dispensa de Licitação e seus Anexos, a Proposta da Contratada, no que couber, e demais "
    "documentos pertinentes, independentemente de transcrição.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA QUARTA - DA LEGISLAÇÃO APLICÁVEL")
doc.add_paragraph(
    "Os casos omissos serão decididos pela Contratante, segundo as disposições do RLC, da Lei nº "
    "13.303, de 2016, e demais normas aplicáveis, e, subsidiariamente, do Código de Defesa do "
    "Consumidor.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA QUINTA - DA PUBLICAÇÃO")
doc.add_paragraph(
    "A publicação do extrato do presente Contrato deverá ser providenciada pela Contratante até o "
    "5º (quinto) dia útil do mês subsequente ao da sua assinatura no Diário Oficial da União, "
    "conforme o art. 480 do RLC.")
doc.add_paragraph("CLÁUSULA VIGÉSIMA SEXTA - DO FORO")
doc.add_paragraph(
    "As partes elegem o foro da Justiça Federal, Seção Judiciária do Estado da Sureg ou de "
    "Brasília-DF, no caso da Matriz, competente para dirimir quaisquer dúvidas suscitadas em "
    "razão deste Contrato, que não puderem ser resolvidas de comum acordo.")
doc.add_paragraph(
    "Por estarem justas e acordadas, firmam o presente instrumento, em 03 (três) vias, de igual "
    "teor e forma, diante das testemunhas abaixo identificadas.")
doc.add_paragraph("Pela Contratante: ___________________     Pela Contratada: ___________________")
doc.add_paragraph("Testemunha 1: ___________________     Testemunha 2: ___________________")
add_bloco_variante_sentinela_fim(_bid_anexo_vi)

doc.save("Template_TR_Dispensa.docx")
print("Template_TR_Dispensa.docx gerado com sucesso.")
```

- [ ] **Step 2: Rodar o script e conferir que o arquivo é gerado**

Run: `cd "templates" && python build_template_tr_dispensa.py`
Expected: imprime `Template_TR_Dispensa.docx gerado com sucesso.` e o
arquivo aparece em `templates/Template_TR_Dispensa.docx`.

- [ ] **Step 3: Commit**

```bash
git add templates/build_template_tr_dispensa.py templates/Template_TR_Dispensa.docx
git commit -m "Add TR-Dispensa docx template build script and generated template"
```

---

### Task 4: Registrar TR-Dispensa em `js/app.js` e `js/triagem.js`

**Files:**
- Modify: `js/app.js` (objeto `INSTRUMENTOS`)
- Modify: `js/triagem.js` (terceiro botão)

**Interfaces:**
- Consumes: `SECOES_TR_DISPENSA` (Task 2), `Template_TR_Dispensa.docx`
  (Task 3).

- [ ] **Step 1: Adicionar a entrada em `INSTRUMENTOS` (`js/app.js`)**

```js
import { SECOES_TR_DISPENSA } from "./dados-tr-dispensa.js";
```

E dentro do objeto `INSTRUMENTOS`:

```js
"TR-Dispensa": {
  secoes: SECOES_TR_DISPENSA,
  titulo: "Termo de Referência — Serviços por Dispensa de Licitação",
  templateUrl: "templates/Template_TR_Dispensa.docx",
  chaveAutosave: "gerador-tr-pb:tr-dispensa:rascunho",
  nomeArquivoRascunho: "rascunho-tr-dispensa.json",
  nomeArquivoGerado: "TR-Dispensa-gerado.docx",
},
```

- [ ] **Step 2: Terceiro botão em `js/triagem.js`**

```js
container.innerHTML = `
  <h1>Triagem inicial</h1>
  <p>Este protótipo cobre o Termo de Referência para Aquisição de Bens
  (art. 127 do RLC), o Termo de Referência para Prestação de Serviços
  (art. 129 do RLC) e o Termo de Referência para Serviços por Dispensa de
  Licitação (art. 416, II, do RLC). O Projeto Básico entra em plano
  seguinte.</p>
  <button id="btn-tr-bens" class="br-button primary" type="button">
    Avançar para o Termo de Referência — Aquisição de Bens
  </button>
  <button id="btn-tr-servicos" class="br-button primary" type="button">
    Avançar para o Termo de Referência — Prestação de Serviços
  </button>
  <button id="btn-tr-dispensa" class="br-button primary" type="button">
    Avançar para o Termo de Referência — Serviços por Dispensa
  </button>
`;
container.querySelector("#btn-tr-bens")
  .addEventListener("click", () => aoConfirmar("TR-Bens"));
container.querySelector("#btn-tr-servicos")
  .addEventListener("click", () => aoConfirmar("TR-Serviços"));
container.querySelector("#btn-tr-dispensa")
  .addEventListener("click", () => aoConfirmar("TR-Dispensa"));
```

- [ ] **Step 3: Testar manualmente no navegador**

Run: servir o diretório por HTTP (ex.: `python -m http.server`, nunca
`file://`). Escolher "Avançar para o Termo de Referência — Serviços por
Dispensa" e conferir:
1. As 26 seções aparecem no sumário.
2. Escolher "Contrato" em vez de "Contrato Simplificado" (Seção 7) e gerar
   o documento — abrir o `.docx` baixado e confirmar que as Seções 21 a 25
   não aparecem, que o Anexo V não aparece e o Anexo VI aparece.
3. Escolher "Contrato Simplificado" — confirmar o oposto (Seções 21-25
   presentes, Anexo V presente, Anexo VI ausente).
4. Marcar "Não" em "A Matriz de Riscos se aplica" (Seção 20) — confirmar
   que o Anexo I não aparece no documento gerado.
5. TR-Bens e TR-Serviços continuam funcionando exatamente como antes
   (regressão zero).

- [ ] **Step 4: Commit**

```bash
git add js/app.js js/triagem.js
git commit -m "Register TR-Dispensa as a third instrument in the app shell"
```

---

### Task 5: Testes de geração ponta a ponta para o TR-Dispensa

**Files:**
- Test: `tests/gerador-docx.test.mjs`

**Interfaces:**
- Consumes: `SECOES_TR_DISPENSA` (Task 2), `Template_TR_Dispensa.docx`
  (Task 3), `gerarDocx` com `blocosAdicionais` (Task 1).

- [ ] **Step 1: Escrever os testes que fecham a matriz de condicionais novos**

```js
// em tests/gerador-docx.test.mjs
import { SECOES_TR_DISPENSA } from "../js/dados-tr-dispensa.js";

function lerTemplateDispensa() {
  const buffer = fs.readFileSync("templates/Template_TR_Dispensa.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlDispensa(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplateDispensa(), respostas, SECOES_TR_DISPENSA);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

test("TR-Dispensa: Contrato Simplificado mantém Seções 21-25 e o Anexo V, remove o Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S07_TipoInstrumentoContratual: "Contrato Simplificado" }));
  assert.ok(texto.includes("21. DA ALTERAÇÃO DO CONTRATO"));
  assert.ok(texto.includes("25. DO FORO"));
  assert.ok(texto.includes("MINUTA DE CONTRATO SIMPLIFICADO"));
  assert.ok(!texto.includes("CLÁUSULA PRIMEIRA - DO OBJETO"));
});

test("TR-Dispensa: Contrato completo remove as Seções 21-25 e o Anexo V, mantém o Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S07_TipoInstrumentoContratual: "Contrato" }));
  assert.ok(!texto.includes("21. DA ALTERAÇÃO DO CONTRATO"));
  assert.ok(!texto.includes("25. DO FORO"));
  assert.ok(!texto.includes("MINUTA DE CONTRATO SIMPLIFICADO"));
  assert.ok(texto.includes("CLÁUSULA PRIMEIRA - DO OBJETO"));
  assert.ok(texto.includes("CLÁUSULA VIGÉSIMA SEXTA - DO FORO"));
});

test("TR-Dispensa: exclusividade ME/EPP no padrão (Sim) usa o texto de exclusividade", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S02_ExcecaoExclusividadeMEEPP: "Sim" }));
  assert.ok(texto.includes("destinado exclusivamente à participação de"));
  assert.ok(!texto.includes("não será destinada exclusivamente"));
});

test("TR-Dispensa: exceção à exclusividade ME/EPP (Não) abre a disputa", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S02_ExcecaoExclusividadeMEEPP: "Não" }));
  assert.ok(texto.includes("não será destinada exclusivamente"));
  assert.ok(!texto.includes("destinado exclusivamente à participação de"));
});

test("TR-Dispensa: Matriz de Riscos aplicável mantém o Anexo I", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S20_ExigeMatrizRiscos: "Sim" }));
  assert.ok(texto.includes("MATRIZ DE RISCOS"));
  assert.ok(texto.includes("Escala de Probabilidade"));
});

test("TR-Dispensa: Matriz de Riscos dispensada remove o Anexo I", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S20_ExigeMatrizRiscos: "Não" }));
  assert.ok(texto.includes("serviços de pronto pagamento"));
  assert.ok(!texto.includes("Escala de Probabilidade"));
});

test("TR-Dispensa: fiscalização por Comissão usa as alíneas b.1, b.2, b.3 na Seção 12", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S09_TipoFiscalizacao: "Comissão de Fiscalização" }));
  assert.ok(texto.includes("o Fiscal Técnico deverá realizar a análise"));
  assert.ok(texto.includes("o Fiscal Setorial deverá realizar a análise"));
});

test("TR-Dispensa: infrações extras da Tabela 1 aparecem junto às 9 fixas", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({
    S16_InfracoesExtras: [{ descricao: "Deixar de portar crachá de identificação", grau: "1" }],
  }));
  assert.ok(texto.includes("Deixar de portar crachá de identificação"));
  assert.ok(texto.includes("Permitir situação que crie a possibilidade de causar dano físico"));
});
```

- [ ] **Step 2: Rodar e ajustar até passar**

Run: `node --test tests/gerador-docx.test.mjs`
Expected: PASS. Se alguma string esperada não bater com o texto real do
template (Task 3), ajustar a string do teste — o template é a fonte da
verdade, transcrito da minuta.

- [ ] **Step 3: Rodar a suíte inteira**

Run: `node --test tests/*.test.mjs`
Expected: PASS em todos os arquivos — nenhuma regressão em TR-Bens ou
TR-Serviços.

- [ ] **Step 4: Commit**

```bash
git add tests/gerador-docx.test.mjs
git commit -m "Add end-to-end generation tests for TR-Dispensa conditionals"
```

---

## Verificação manual final (checklist)

Depois do Task 5 verde, gerar o `.docx` do TR-Dispensa pelo navegador
(Task 4, Step 3) cobrindo os dois lados de cada condicional novo, e abrir
no Word real conferindo que:
- a troca de instrumento contratual remove/mantém as Seções 21-25 e o
  anexo certo sem parágrafo órfão nem numeração de seção quebrada;
- a Seção 3 (Habilitação) sai formatada corretamente, sem nota explicativa
  residual;
- TR-Bens e TR-Serviços continuam gerando exatamente como antes.
