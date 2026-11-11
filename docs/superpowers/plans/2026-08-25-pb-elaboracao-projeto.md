# PB: Corrigir Lacunas da Natureza "Elaboração de Projeto" — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Corrigir, no instrumento PB, as três lacunas de conteúdo e o
problema de fundamentação que o teste real do processo SUREG/SC expôs —
todos condicionados ao campo já existente `PB01_NaturezaObjeto = "Elaboração
de projeto, laudo ou anteprojeto por terceiros"` — sem alterar o
comportamento hoje validado para `"Execução de obra ou serviço já
projetado"`.

**Architecture:** `js/dados-pb.js` (novos campos) + `templates/build_template_pb.py`
(novo texto condicional no `.docx`-modelo, gerando `templates/Template_PB.docx`)
+ `js/gerador-docx.js` (mecanismo já existente, não muda). Toda a variação
usa o mecanismo de blocos condicionais já estabelecido no arquivo
(`add_bloco_variante` para texto com bookmark de campo embutido,
`add_bloco_variante_sentinela_inicio/fim` para regiões com vários
parágrafos ou tabelas), com o mesmo padrão de "ausência por omissão" já
usado no Anexo VIII (bookmark só existe para uma das duas naturezas — a
outra nunca é criada, então a remoção correspondente é sempre um no-op
inofensivo).

**Tech Stack:** JavaScript puro (motor do site), Python 3.12 +
python-docx (gerador do `.docx`-modelo, offline), Node.js `--test` (testes
contra `gerador-docx.js`).

## Global Constraints

- Nenhum campo novo aparece fora de uma condição de visibilidade correta
  (`mostrarSe` no formulário, `blocoTexto`/`blocosAdicionais` no
  documento) — nunca altera o comportamento hoje validado para
  `PB01_NaturezaObjeto = "Execução de obra ou serviço já projetado"`.
- Toda citação legal usa o texto exato do RLC 10.901, conferido em
  `documentos de referencia/rlc.txt` — nenhuma paráfrase de artigo, inciso
  ou alínea.
- `templates/Template_PB.docx` é sempre regenerado (`python
  build_template_pb.py`, rodado de dentro de `templates/`) depois de
  qualquer mudança em `build_template_pb.py`, antes de rodar os testes —
  os testes carregam o `.docx` do disco, não o script Python.
- Cada tarefa termina com `node --test tests/*.test.mjs` passando (suíte
  inteira, não só o arquivo novo) — regressão de qualquer outro
  instrumento pararia a tarefa.

---

## Contexto de mecanismo (não repetir por tarefa — vale para todas)

`js/gerador-docx.js:198-212` processa todo campo com `blocoTexto: true`
(a lista vem de `controladoresBlocoVariante(secoes)`, `js/modelo.js`):
para cada opção do campo que **não** é a resposta atual (ou o `padrao`,
se não respondido), remove o bookmark `{base}_BLOCO_{OPÇÃO_NORMALIZADA}`
para `base` = o próprio `campo.id` **e** cada nome em
`campo.blocosAdicionais`. Se esse bookmark nunca existir no `.docx`-modelo
(porque só foi criado texto para uma das opções), a remoção é um no-op —
é assim que o Anexo VIII (`templates/build_template_pb.py:1217-1232`) já
fica ausente quando a natureza é "Elaboração de projeto...": só existe
bookmark `PB_AnexoVIII_BLOCO_EXECUCAO_DE_OBRA_OU_SERVICO_JA_PROJETADO`,
nunca o inverso. Esta mesma técnica ("ausência por omissão") é reutilizada
nas Tarefas 1, 3 e 4.

`normalizar_variante` (Python, `build_template_pb.py:62-64`) e
`normalizarVariante` (JS, `gerador-docx.js:87-91`) fazem a mesma
transformação: NFKD, remove acentos, maiúsculas, troca tudo que não é
alfanumérico por `_`. `"Elaboração de projeto, laudo ou anteprojeto por
terceiros"` normaliza para
`ELABORACAO_DE_PROJETO_LAUDO_OU_ANTEPROJETO_POR_TERCEIROS`; `"Não"` para
`NAO`. Os nomes de bookmark abaixo já vêm pré-calculados com esse
resultado — não precisa rodar a função para conferir.

`removerBlocoBookmark` remove o **parágrafo inteiro** que contém o
bookmark (do início do parágrafo do `bookmarkStart` até o fim do
parágrafo do `bookmarkEnd` correspondente, mesmo que sejam parágrafos
diferentes — é assim que blocos de vários parágrafos, como o Anexo II-A,
já funcionam). **Duas variantes de texto nunca podem dividir o mesmo
parágrafo** — cada variante precisa do próprio parágrafo dedicado, senão
remover uma apaga a outra junto. Todo o texto novo abaixo já respeita essa
regra.

---

### Task 1: Seção 13 (Qualificação Técnica) condicionada à natureza do objeto

**Files:**
- Modify: `js/dados-pb.js:16` (array `blocosAdicionais` de `PB01_NaturezaObjeto`)
- Modify: `templates/build_template_pb.py:422-444` (parágrafos 13.1.1-13.1.3)
- Modify: `templates/Template_PB.docx` (regenerado pelo script acima)
- Test: `tests/pb-elaboracao-projeto.test.mjs` (novo arquivo)

**Interfaces:**
- Consumes: `PB01_NaturezaObjeto` (já existe, `js/dados-pb.js:11-16`,
  `opcoes: ["Execução de obra ou serviço já projetado", "Elaboração de
  projeto, laudo ou anteprojeto por terceiros"]`).
- Produces: bookmark `PB13_QualificacaoTecnica_BLOCO_<OPÇÃO>` no
  `.docx`-modelo — não introduz nenhum campo novo de resposta (o texto de
  13.1 não tem parâmetro preenchível em nenhuma das duas variantes, só
  muda o texto fixo).

- [ ] **Step 1: Escrever o teste que falha**

Criar `tests/pb-elaboracao-projeto.test.mjs`:

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { SECOES_PB } from "../js/dados-pb.js";

globalThis.JSZip = JSZip;

function lerTemplate() {
  const buffer = fs.readFileSync("templates/Template_PB.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function textoDoDocumento(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate(), respostas, SECOES_PB);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const xml = await zip.file("word/document.xml").async("string");
  return xml.replace(/<[^>]+>/g, "");
}

test("PB Seção 13: natureza Elaboração de projeto usa Anexo VIII-B, item 1, alínea r, e art. 133, §6º, por analogia", async () => {
  const texto = await textoDoDocumento({
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
  });
  assert.ok(texto.includes("item 1, alínea “r”, do Anexo VIII-B do RLC"));
  assert.ok(texto.includes("art. 133, §6º, do RLC, por analogia"));
  assert.ok(!texto.includes("parcelas de maior relevância técnica e econômica do objeto"));
  assert.ok(!texto.includes("Admite-se o somatório de atestados."));
  // 13.1.1 (registro CREA/CAU) e as vedações de 13.1.4 continuam comuns às duas naturezas.
  assert.ok(texto.includes("Certidão de registro de pessoa jurídica no CREA/CAU"));
  assert.ok(texto.includes("percentuais mínimos acima de 50% dos itens de maior relevância"));
});

test("PB Seção 13: natureza Execução de obra mantém o texto de hoje (não regressão)", async () => {
  const texto = await textoDoDocumento({
    PB01_NaturezaObjeto: "Execução de obra ou serviço já projetado",
  });
  assert.ok(texto.includes("parcelas de maior relevância técnica e econômica do objeto"));
  assert.ok(texto.includes("Admite-se o somatório de atestados."));
  assert.ok(!texto.includes("Anexo VIII-B"));
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: FAIL — o texto do Anexo VIII-B ainda não existe no `.docx`-modelo
(a variante "Elaboração de projeto" hoje sai igual à de obra, porque
13.1.1-13.1.3 ainda não são condicionais).

- [ ] **Step 3: Adicionar `"PB13_QualificacaoTecnica"` a `blocosAdicionais`**

Em `js/dados-pb.js`, campo `PB01_NaturezaObjeto` (linha 16):

Find:
```js
      blocosAdicionais: ["PB14_CriteriosObra", "PB21_GarantiaObra", "PB_AnexoVIII", "PB28_MetodologiaCusto", "PB_CessaoDireitos"] },
```

Replace:
```js
      blocosAdicionais: ["PB13_QualificacaoTecnica", "PB14_CriteriosObra", "PB21_GarantiaObra", "PB_AnexoVIII", "PB28_MetodologiaCusto", "PB_CessaoDireitos"] },
```

- [ ] **Step 4: Tornar 13.1.1-13.1.3 condicionais em `build_template_pb.py`**

Find (linhas 424-437):
```python
doc.add_paragraph("13.1. Qualificação técnica")
doc.add_paragraph(
    "13.1.1. Certidão de registro de pessoa jurídica no CREA/CAU, em nome da licitante, com "
    "validade na data de recebimento dos documentos de habilitação.")
add_paragraph_with_bookmarks([
    "13.1.2. Certidão de Acervo Técnico (CAT) em nome do(s) responsável(is) técnico(s), "
    "comprovando responsabilidade por serviços de características semelhantes às parcelas "
    "de maior relevância técnica e econômica do objeto: ", ("PB13_ParcelasRelevantes", ""), ".",
])
add_paragraph_with_bookmarks([
    "13.1.3. Atestado(s) de capacidade técnica emitido(s) por pessoa jurídica de direito "
    "público ou privado, comprovando execução de serviços semelhantes: ", ("PB13_QuantitativoMinimo", ""),
    " Admite-se o somatório de atestados.",
])
```

Replace:
```python
doc.add_paragraph("13.1. Qualificação técnica")
doc.add_paragraph(
    "13.1.1. Certidão de registro de pessoa jurídica no CREA/CAU, em nome da licitante, com "
    "validade na data de recebimento dos documentos de habilitação.")
add_bloco_variante("PB13_QualificacaoTecnica", {
    "Execução de obra ou serviço já projetado": [
        [
            "13.1.2. Certidão de Acervo Técnico (CAT) em nome do(s) responsável(is) "
            "técnico(s), comprovando responsabilidade por serviços de características "
            "semelhantes às parcelas de maior relevância técnica e econômica do objeto: ",
            ("PB13_ParcelasRelevantes", ""), ".",
        ],
        [
            "13.1.3. Atestado(s) de capacidade técnica emitido(s) por pessoa jurídica de "
            "direito público ou privado, comprovando execução de serviços semelhantes: ",
            ("PB13_QuantitativoMinimo", ""), " Admite-se o somatório de atestados.",
        ],
    ],
    "Elaboração de projeto, laudo ou anteprojeto por terceiros": [
        [
            "13.1.2. Nos termos do item 1, alínea “r”, do Anexo VIII-B do RLC, "
            "constará dos autos certificação, acompanhada de comprovação, de que a "
            "contratada preenche os requisitos de habilitação e de qualificação técnica "
            "mínima necessários à elaboração do projeto, mediante Certidão de Acervo "
            "Técnico (CAT) ou instrumento equivalente em nome do(s) responsável(is) "
            "técnico(s), comprovando responsabilidade por serviço de características "
            "semelhantes ao objeto: ", ("PB13_ParcelasRelevantes", ""),
            ", bem como as razões da escolha da contratada.",
        ],
    ],
})
```

`13.1.4` (linhas 438-444, vedações do art. 157, XII) e `13.2` inteira
(linhas 445-462, qualificação econômico-financeira) **não mudam de
posição nem de condição** — continuam comuns às duas naturezas.

- [ ] **Step 5: Citar o art. 133, §6º, por analogia, em 13.2**

Find (linha 445):
```python
doc.add_paragraph("13.2. Qualificação econômico-financeira")
```

Replace:
```python
doc.add_paragraph(
    "13.2. Qualificação econômico-financeira (art. 133, §6º, do RLC, por analogia — o "
    "Regulamento não prevê regra própria de habilitação econômico-financeira para serviços "
    "técnicos de natureza intelectual; aplicam-se, por analogia, os mesmos parâmetros já "
    "usados neste Projeto Básico para execução de obra ou serviço de engenharia)")
```

Esse texto novo é comum às duas naturezas (o parágrafo 13.2 inteiro não
está dentro de nenhum `add_bloco_variante`) — a citação aparece sempre,
não só quando a natureza é elaboração de projeto. Isso está correto: o
achado da spec é que o §6º já fundamenta os mesmos números hoje usados
sem citação nenhuma; a correção vale para os dois casos.

- [ ] **Step 6: Regenerar o `.docx`-modelo**

```bash
cd templates
python build_template_pb.py
cd ..
```

Expected: `Template_PB.docx gerado com sucesso.`

- [ ] **Step 7: Rodar e confirmar que passa**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: PASS — 2 testes.

- [ ] **Step 8: Rodar a suíte inteira (checar não regressão)**

Run: `node --test tests/*.test.mjs`
Expected: PASS — todos os testes, incluindo `tests/pb-dispensa.test.mjs` e
`tests/dados-pb.test.mjs`.

- [ ] **Step 9: Commit**

```bash
git add js/dados-pb.js templates/build_template_pb.py templates/Template_PB.docx tests/pb-elaboracao-projeto.test.mjs
git commit -m "PB: Seção 13 condicionada à natureza — Anexo VIII-B 'r' + art. 133 §6º por analogia"
```

---

### Task 2: Seção 16 (Recebimento) — variante simplificada do art. 534

**Files:**
- Modify: `js/dados-pb.js:205-213` (nova pergunta na Seção 16)
- Modify: `templates/build_template_pb.py:499-511`
- Modify: `templates/Template_PB.docx` (regenerado)
- Test: `tests/pb-elaboracao-projeto.test.mjs` (adicionar)

**Interfaces:**
- Consumes: `PB01_NaturezaObjeto` (só controla a visibilidade do campo no
  formulário via `mostrarSe`; o mecanismo de remoção de bloco no `.docx`
  **não** depende de natureza — ver decisão abaixo).
- Produces: campo novo `PB16_ElegivelRecebimentoSimplificado` (parâmetro,
  `blocoTexto: true`), bookmarks
  `PB16_ElegivelRecebimentoSimplificado_BLOCO_SIM` e
  `PB16_ElegivelRecebimentoSimplificado_BLOCO_NAO`.

**Decisão de mecanismo:** `padrao: "Não"`. O laço de remoção em
`gerador-docx.js:198-212` roda para **todo** campo `blocoTexto`, mesmo
quando ele está escondido no formulário por `mostrarSe` (a visibilidade
só afeta a UI e a remoção do bookmark do próprio campo — não afeta o laço
de blocos-variante). Como o campo só é visível quando a natureza é
"Elaboração de projeto...", ele fica sem resposta (`respostas[id] ===
undefined`) sempre que a natureza é "Execução de obra..." — nesse caso o
laço usa `campo.padrao`. Com `padrao: "Não"`, o resultado para natureza
"obra" é sempre manter o texto de dois estágios de hoje (nenhuma mudança
de comportamento), sem precisar aninhar esta variante dentro de um
segundo bloco condicional por natureza.

- [ ] **Step 1: Escrever o teste que falha**

Adicionar a `tests/pb-elaboracao-projeto.test.mjs`:

```js
test("PB Seção 16: elegível ao recebimento simplificado usa o rito do art. 534", async () => {
  const texto = await textoDoDocumento({
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
    PB16_ElegivelRecebimentoSimplificado: "Sim",
  });
  assert.ok(texto.includes("art. 534 do RLC"));
  assert.ok(texto.includes("Recibo Definitivo Simplificado"));
  assert.ok(!texto.includes("Termo de Recebimento Provisório detalhado"));
});

test("PB Seção 16: sem resposta (natureza obra ou natureza projeto sem confirmar) mantém o rito de dois estágios", async () => {
  const textoObra = await textoDoDocumento({
    PB01_NaturezaObjeto: "Execução de obra ou serviço já projetado",
  });
  assert.ok(textoObra.includes("Termo de Recebimento Provisório detalhado"));
  assert.ok(!textoObra.includes("art. 534 do RLC"));

  const textoProjetoSemConfirmar = await textoDoDocumento({
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
  });
  assert.ok(textoProjetoSemConfirmar.includes("Termo de Recebimento Provisório detalhado"));
  assert.ok(!textoProjetoSemConfirmar.includes("art. 534 do RLC"));
});

test("PB Seção 16: natureza projeto com Não explícito mantém o rito de dois estágios", async () => {
  const texto = await textoDoDocumento({
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
    PB16_ElegivelRecebimentoSimplificado: "Não",
  });
  assert.ok(texto.includes("Termo de Recebimento Provisório detalhado"));
  assert.ok(!texto.includes("art. 534 do RLC"));
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: FAIL nos 3 testes novos — o campo e o texto do art. 534 ainda
não existem.

- [ ] **Step 3: Adicionar o campo em `js/dados-pb.js`**

Find (Seção 16, linhas 205-213):
```js
  { id: "16", titulo: "DO RECEBIMENTO DO OBJETO CONTRATUAL", campos: [
    { id: "PB16_PrazoRecebimentoProvisorioDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para o recebimento provisório, contado do adimplemento da parcela/entrega (dias corridos)",
      ajuda: "Art. 533, I, \"a\", do RLC: até 15 dias para obras/serviços de engenharia.",
      duplicarPara: ["PB16_PrazoRecebimentoProvisorioDias_AnexoContrato"] },
    { id: "PB16_PrazoRecebimentoDefinitivoDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para o recebimento definitivo, contado do recebimento provisório (dias corridos)",
      duplicarPara: ["PB16_PrazoRecebimentoDefinitivoDias_AnexoContrato"] },
  ]},
```

Replace:
```js
  { id: "16", titulo: "DO RECEBIMENTO DO OBJETO CONTRATUAL", campos: [
    { id: "PB16_ElegivelRecebimentoSimplificado", tipo: "parametro", input: "select",
      pergunta: "O objeto não se compõe de aparelhos, equipamentos ou instalações sujeitos à verificação de funcionamento e produtividade (art. 534, caput, RLC)?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" },
      ajuda: "Se \"Sim\", o recebimento segue o rito simplificado do art. 534 do RLC (uma vistoria, Recibo Definitivo Simplificado) em vez dos dois estágios do art. 533. Responda \"Não\" se a entrega envolver instalação de equipamento sujeito a comissionamento." },
    { id: "PB16_PrazoRecebimentoProvisorioDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para o recebimento provisório, contado do adimplemento da parcela/entrega (dias corridos)",
      ajuda: "Art. 533, I, \"a\", do RLC: até 15 dias para obras/serviços de engenharia.",
      duplicarPara: ["PB16_PrazoRecebimentoProvisorioDias_AnexoContrato"] },
    { id: "PB16_PrazoRecebimentoDefinitivoDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para o recebimento definitivo, contado do recebimento provisório (dias corridos)",
      duplicarPara: ["PB16_PrazoRecebimentoDefinitivoDias_AnexoContrato"] },
  ]},
```

- [ ] **Step 4: Envolver o texto de hoje em um bloco "Não" e acrescentar o "Sim" em `build_template_pb.py`**

Find (linhas 500-511):
```python
add_heading("16. DO RECEBIMENTO DO OBJETO CONTRATUAL (ART. 157, XV, RLC)")
doc.add_paragraph("16.1. O recebimento provisório ou definitivo não exclui a responsabilidade da Contratada pelos prejuízos resultantes da incorreta execução do Contrato.")
add_paragraph_with_bookmarks([
    "16.2. O recebimento provisório será realizado pela Fiscalização em até ", ("PB16_PrazoRecebimentoProvisorioDias", "____"),
    " dias corridos do adimplemento da parcela/entrega, formalizado em Termo de Recebimento "
    "Provisório detalhado, encaminhado ao responsável pelo recebimento definitivo.",
])
add_paragraph_with_bookmarks([
    "16.3. O recebimento definitivo, que concretiza o ateste da execução, será realizado "
    "pelo empregado ou comissão designada em até ", ("PB16_PrazoRecebimentoDefinitivoDias", "____"),
    " dias corridos do recebimento provisório, após verificação de toda a documentação.",
])
```

Replace:
```python
add_heading("16. DO RECEBIMENTO DO OBJETO CONTRATUAL (ART. 157, XV, RLC)")
doc.add_paragraph("16.1. O recebimento provisório ou definitivo não exclui a responsabilidade da Contratada pelos prejuízos resultantes da incorreta execução do Contrato.")

_bid_16_dois_estagios = add_bloco_variante_sentinela_inicio(
    "PB16_ElegivelRecebimentoSimplificado_BLOCO_NAO")
add_paragraph_with_bookmarks([
    "16.2. O recebimento provisório será realizado pela Fiscalização em até ", ("PB16_PrazoRecebimentoProvisorioDias", "____"),
    " dias corridos do adimplemento da parcela/entrega, formalizado em Termo de Recebimento "
    "Provisório detalhado, encaminhado ao responsável pelo recebimento definitivo.",
])
add_paragraph_with_bookmarks([
    "16.3. O recebimento definitivo, que concretiza o ateste da execução, será realizado "
    "pelo empregado ou comissão designada em até ", ("PB16_PrazoRecebimentoDefinitivoDias", "____"),
    " dias corridos do recebimento provisório, após verificação de toda a documentação.",
])
add_bloco_variante_sentinela_fim(_bid_16_dois_estagios)

add_bloco_variante("PB16_ElegivelRecebimentoSimplificado", {
    "Sim": [
        "16.2. Nos termos do art. 534 do RLC, por se tratar de objeto sem aparelhos, "
        "equipamentos ou instalações sujeitos à verificação de funcionamento e "
        "produtividade, o recebimento seguirá o rito simplificado: executado o serviço, a "
        "Contratada comunicará por escrito ao Fiscal do Contrato ou à Comissão de "
        "Fiscalização, dentro do prazo contratual, para a realização de vistoria.",
        "16.3. No prazo máximo de 10 (dez) dias contados da comunicação, será realizada "
        "vistoria com vistas à emissão do Recibo Definitivo Simplificado; havendo "
        "pendências, será concedido prazo de até 20 (vinte) dias para as correções, "
        "seguido de nova vistoria e, sanada a pendência, emissão do Recibo Definitivo "
        "Simplificado em até 10 (dez) dias.",
    ],
})
```

`16.4` e `16.5` (linhas 512+) não mudam — continuam depois deste trecho,
comuns às duas variantes.

- [ ] **Step 5: Regenerar o `.docx`-modelo**

```bash
cd templates
python build_template_pb.py
cd ..
```

- [ ] **Step 6: Rodar e confirmar que passa**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: PASS — 5 testes (2 da Task 1 + 3 desta tarefa).

- [ ] **Step 7: Rodar a suíte inteira**

Run: `node --test tests/*.test.mjs`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add js/dados-pb.js templates/build_template_pb.py templates/Template_PB.docx tests/pb-elaboracao-projeto.test.mjs
git commit -m "PB: Seção 16 ganha variante de recebimento simplificado (art. 534 RLC)"
```

---

### Task 3: Anexo X — Fator K/TRDE estruturado — PAUSADA

**Não executar ainda.** Aguardando a Ana confirmar se a GEPEA já tem uma
referência própria para os rótulos de K1-K4 (ver "Dúvida de escopo" na
spec `2026-08-25-pb-elaboracao-projeto-design.md`). As Tasks 4 e 5 abaixo
já foram ajustadas para rodar corretamente **sem** esta tarefa ter
passado — quando ela for retomada, quem a executar precisa também voltar
em `js/dados-pb.js` e no Step 3 da Task 4 para inserir `"PB_AnexoX"` de
volta no array `blocosAdicionais` (removido de lá propositalmente por
enquanto), e reativar as asserções de Achado 3 comentadas no teste de
fixture da Task 5.

**Files:**
- Modify: `js/dados-pb.js:277-299` (Seção 27/28 — substituir
  `PB28_MetodologiaCustoIntelectual` único por 6 campos + 1 mantido)
- Modify: `templates/build_template_pb.py:1217-1232` (inserir Anexo X logo
  depois do Anexo VIII, antes da Minuta de Contrato)
- Modify: `templates/Template_PB.docx` (regenerado)
- Test: `tests/pb-elaboracao-projeto.test.mjs` (adicionar)

**Interfaces:**
- Consumes: `PB01_NaturezaObjeto` (`blocosAdicionais` ganha `"PB_AnexoX"`).
- Produces: campos novos `PB28_FatorK1`, `PB28_FatorK2`, `PB28_FatorK3`,
  `PB28_FatorK4`, `PB28_FatorKResultante`, `PB28_TRDE` (todos `parametro`,
  `input: "numero"`, `mostrarSe: { PB01_NaturezaObjeto: "Elaboração de
  projeto, laudo ou anteprojeto por terceiros" }`); mantém
  `PB28_MetodologiaCustoIntelectual` (mesmo id, mesma visibilidade) como
  campo de narrativa. Anexo X inteiro só existe quando a natureza é
  "Elaboração de projeto...".

- [ ] **Step 1: Escrever o teste que falha**

Adicionar a `tests/pb-elaboracao-projeto.test.mjs`:

```js
test("PB Anexo X: aparece com os 6 coeficientes preenchidos quando a natureza é Elaboração de projeto", async () => {
  const texto = await textoDoDocumento({
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
    PB28_FatorK1: "3,5", PB28_FatorK2: "2,0", PB28_FatorK3: "1,2", PB28_FatorK4: "8,0",
    PB28_FatorKResultante: "14,7", PB28_TRDE: "10",
    PB28_MetodologiaCustoIntelectual: "Correlação entre os coeficientes K1-K4 e a TRDE, própria da engenharia consultiva.",
  });
  assert.ok(texto.includes("ANEXO X DO PROJETO BÁSICO"));
  assert.ok(texto.includes("MEMÓRIA DE CÁLCULO"));
  assert.ok(texto.includes("3,5"));
  assert.ok(texto.includes("14,7"));
  assert.ok(texto.includes("Correlação entre os coeficientes K1-K4"));
});

test("PB Anexo X: não aparece quando a natureza é Execução de obra", async () => {
  const texto = await textoDoDocumento({
    PB01_NaturezaObjeto: "Execução de obra ou serviço já projetado",
  });
  assert.ok(!texto.includes("ANEXO X DO PROJETO BÁSICO"));
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: FAIL nos 2 testes novos — "ANEXO X" ainda não existe.

- [ ] **Step 3: Adicionar `"PB_AnexoX"` a `blocosAdicionais` e trocar os campos da Seção 28**

Em `js/dados-pb.js`, campo `PB01_NaturezaObjeto` (mesma linha já editada
na Task 1 — acrescentar mais um item):

Find:
```js
      blocosAdicionais: ["PB13_QualificacaoTecnica", "PB14_CriteriosObra", "PB21_GarantiaObra", "PB_AnexoVIII", "PB28_MetodologiaCusto", "PB_CessaoDireitos"] },
```

Replace:
```js
      blocosAdicionais: ["PB13_QualificacaoTecnica", "PB14_CriteriosObra", "PB21_GarantiaObra", "PB_AnexoVIII", "PB28_MetodologiaCusto", "PB_AnexoX", "PB_CessaoDireitos"] },
```

Find (Seção 28, linhas 295-298):
```js
    { id: "PB28_MetodologiaCustoIntelectual", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva a metodologia de estimativa de custo adotada (Fator K/TRDE), com a correlação entre os componentes explicitada",
      ajuda: "Metodologia própria da engenharia consultiva para serviços de natureza intelectual — substitui o BDI de obras/serviços de execução.",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
```

Replace:
```js
    { id: "PB28_FatorK1", tipo: "parametro", input: "numero",
      pergunta: "Coeficiente K1 — despesas administrativas (%)",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
    { id: "PB28_FatorK2", tipo: "parametro", input: "numero",
      pergunta: "Coeficiente K2 — benefícios e despesas indiretas (equivalente ao BDI) (%)",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
    { id: "PB28_FatorK3", tipo: "parametro", input: "numero",
      pergunta: "Coeficiente K3 — despesas financeiras (%)",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
    { id: "PB28_FatorK4", tipo: "parametro", input: "numero",
      pergunta: "Coeficiente K4 — lucro (%)",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
    { id: "PB28_FatorKResultante", tipo: "parametro", input: "numero",
      pergunta: "Fator K resultante (soma dos coeficientes K1 a K4) (%)",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
    { id: "PB28_TRDE", tipo: "parametro", input: "numero",
      pergunta: "Taxa de Retribuição pela Despesa de Estrutura (TRDE), se aplicável (%)",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
    { id: "PB28_MetodologiaCustoIntelectual", tipo: "especificacao", input: "textarea",
      pergunta: "Correlação entre os componentes do Fator K/TRDE e a composição do preço",
      ajuda: "Metodologia própria da engenharia consultiva para serviços de natureza intelectual — substitui o BDI de obras/serviços de execução (art. 165 do RLC). O detalhamento numérico de cada coeficiente vai no Anexo X.",
      mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros" } },
```

O texto de introdução da Seção 28 (`28.1.`, linhas 782-787 de
`build_template_pb.py`, variante `_bid_28_elaboracao`) não muda — continua
citando a mesma substituição do art. 165. Só o `28.2.` some do corpo da
Seção 28 (a narrativa migra, ressignificada, para o Anexo X) — atualizar:

Find (`build_template_pb.py:782-789`):
```python
_bid_28_elaboracao = add_bloco_variante_sentinela_inicio(
    f"PB28_MetodologiaCusto_BLOCO_{normalizar_variante('Elaboração de projeto, laudo ou anteprojeto por terceiros')}")
doc.add_paragraph(
    "28.1. Por se tratar de serviço técnico de natureza intelectual, a orçamentação não "
    "segue a equação de BDI do art. 165 do RLC, adotando metodologia de Fator K e TRDE, "
    "própria da engenharia consultiva.")
add_paragraph_with_bookmarks(["28.2. ", ("PB28_MetodologiaCustoIntelectual", "")])
add_bloco_variante_sentinela_fim(_bid_28_elaboracao)
```

Replace:
```python
_bid_28_elaboracao = add_bloco_variante_sentinela_inicio(
    f"PB28_MetodologiaCusto_BLOCO_{normalizar_variante('Elaboração de projeto, laudo ou anteprojeto por terceiros')}")
doc.add_paragraph(
    "28.1. Por se tratar de serviço técnico de natureza intelectual, a orçamentação não "
    "segue a equação de BDI do art. 165 do RLC, adotando metodologia de Fator K e TRDE, "
    "própria da engenharia consultiva, detalhada no Anexo X deste Projeto Básico.")
add_bloco_variante_sentinela_fim(_bid_28_elaboracao)
```

- [ ] **Step 4: Inserir o Anexo X em `build_template_pb.py`, logo após o Anexo VIII**

Find (linhas 1222-1234, fim do Anexo VIII, início do comentário da Minuta
de Contrato):
```python
_bid_anexo_viii = add_bloco_variante_sentinela_inicio(
    f"PB_AnexoVIII_BLOCO_{normalizar_variante('Execução de obra ou serviço já projetado')}")
add_heading("ANEXO VIII DO PROJETO BÁSICO")
add_heading("MODELO DE PLACA DE OBRA/SERVIÇO DE ENGENHARIA")
doc.add_paragraph(
    "A placa de obra/serviço de engenharia deverá seguir o modelo institucional vigente à "
    "época da contratação, observando as dimensões mínimas, materiais, identidade visual, "
    "logomarcas, informações do contrato, identificação do responsável técnico e número da "
    "ART/RRT, conforme os padrões da Conab, a ABNT NBR 12284 e demais normas aplicáveis "
    "(inclusive eventuais restrições da legislação eleitoral vigentes à época).")
add_bloco_variante_sentinela_fim(_bid_anexo_viii)

# ANEXO — MINUTA DE CONTRATO (só existe se Dispensa de Licitação; art.
```

Replace:
```python
_bid_anexo_viii = add_bloco_variante_sentinela_inicio(
    f"PB_AnexoVIII_BLOCO_{normalizar_variante('Execução de obra ou serviço já projetado')}")
add_heading("ANEXO VIII DO PROJETO BÁSICO")
add_heading("MODELO DE PLACA DE OBRA/SERVIÇO DE ENGENHARIA")
doc.add_paragraph(
    "A placa de obra/serviço de engenharia deverá seguir o modelo institucional vigente à "
    "época da contratação, observando as dimensões mínimas, materiais, identidade visual, "
    "logomarcas, informações do contrato, identificação do responsável técnico e número da "
    "ART/RRT, conforme os padrões da Conab, a ABNT NBR 12284 e demais normas aplicáveis "
    "(inclusive eventuais restrições da legislação eleitoral vigentes à época).")
add_bloco_variante_sentinela_fim(_bid_anexo_viii)

# ANEXO X - MEMÓRIA DE CÁLCULO (FATOR K/TRDE). Só existe para
# "Elaboração de projeto..." — o RLC não define Fator K/TRDE (metodologia
# de mercado da engenharia consultiva, fora do Regulamento); o gerador
# não calcula nem valida a fórmula, só estrutura os coeficientes que o
# elaborador já apurou, no mesmo espírito dos três campos de BDI da
# Seção 28 (variante "Execução de obra").
doc.add_page_break()
_bid_anexo_x = add_bloco_variante_sentinela_inicio(
    f"PB_AnexoX_BLOCO_{normalizar_variante('Elaboração de projeto, laudo ou anteprojeto por terceiros')}")
add_heading("ANEXO X DO PROJETO BÁSICO")
add_heading("MEMÓRIA DE CÁLCULO — FATOR K E TRDE")
doc.add_paragraph(
    "A metodologia de Fator K e Taxa de Retribuição pela Despesa de Estrutura (TRDE) "
    "substitui a equação de BDI (art. 165 do RLC) para serviços técnicos de natureza "
    "intelectual (elaboração de projeto, laudo ou anteprojeto), sendo o cálculo de "
    "responsabilidade da área técnica.")
tabela_fator_k = add_tabela(["Componente", "Valor (%)"])
for rotulo, campo_id in [
    ("K1 — Despesas administrativas", "PB28_FatorK1"),
    ("K2 — Benefícios e despesas indiretas", "PB28_FatorK2"),
    ("K3 — Despesas financeiras", "PB28_FatorK3"),
    ("K4 — Lucro", "PB28_FatorK4"),
    ("Fator K resultante", "PB28_FatorKResultante"),
    ("TRDE (quando aplicável)", "PB28_TRDE"),
]:
    linha = tabela_fator_k.add_row()
    linha.cells[0].text = rotulo
    add_paragraph_with_bookmarks_em_celula = linha.cells[1].paragraphs[0]
    _bookmark_id[0] += 1
    bid_celula = str(_bookmark_id[0])
    start_celula = docx.oxml.shared.OxmlElement("w:bookmarkStart")
    start_celula.set(docx.oxml.ns.qn("w:id"), bid_celula)
    start_celula.set(docx.oxml.ns.qn("w:name"), campo_id)
    add_paragraph_with_bookmarks_em_celula._p.append(start_celula)
    add_paragraph_with_bookmarks_em_celula.add_run("____")
    end_celula = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
    end_celula.set(docx.oxml.ns.qn("w:id"), bid_celula)
    add_paragraph_with_bookmarks_em_celula._p.append(end_celula)
add_paragraph_with_bookmarks([
    "Correlação entre os componentes: ", ("PB28_MetodologiaCustoIntelectual", ""),
])
add_bloco_variante_sentinela_fim(_bid_anexo_x)

# ANEXO — MINUTA DE CONTRATO (só existe se Dispensa de Licitação; art.
```

**Nota de implementação:** os bookmarks dentro de célula de tabela usam o
mesmo padrão de `add_paragraph_with_bookmarks`, mas escrito manualmente
porque essa função só sabe adicionar a um novo parágrafo de `doc`, não a
um parágrafo já existente dentro de uma célula (`tabela.add_table` já
cria um parágrafo vazio por célula). Se a implementação achar mais limpo
extrair isso para uma função auxiliar `add_bookmark_em_paragrafo(p,
campo_id, texto_inicial)` reutilizável, é uma refatoração de baixo risco
— o teste do Step 6 não muda.

- [ ] **Step 5: Regenerar o `.docx`-modelo**

```bash
cd templates
python build_template_pb.py
cd ..
```

- [ ] **Step 6: Rodar e confirmar que passa**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: PASS — 7 testes (5 das Tasks 1-2 + 2 desta tarefa).

- [ ] **Step 7: Rodar a suíte inteira**

Run: `node --test tests/*.test.mjs`
Expected: PASS.

- [ ] **Step 8: Commit**

```bash
git add js/dados-pb.js templates/build_template_pb.py templates/Template_PB.docx tests/pb-elaboracao-projeto.test.mjs
git commit -m "PB: Anexo X estruturado para Fator K/TRDE (coeficientes K1-K4), condicionado à natureza"
```

---

### Task 4: Fundamentação da dispensa condicionada à natureza (Seções 2.5 e 3)

**Files:**
- Modify: `js/dados-pb.js:16` (mais um item em `blocosAdicionais`)
- Modify: `templates/build_template_pb.py:189-199` (Seção 2.5)
- Modify: `templates/build_template_pb.py:245-251` (Seção 3, linha
  "Fundamento da dispensa")
- Modify: `templates/Template_PB.docx` (regenerado)
- Test: `tests/pb-elaboracao-projeto.test.mjs` (adicionar)

**Interfaces:**
- Consumes: `PB01_RegimeContratacao` (controla se a Seção 2.5/3 fala de
  dispensa) e `PB01_NaturezaObjeto` (controla qual texto de dispensa,
  aninhado dentro do primeiro).
- Produces: bases novas `PB02_FundamentoDispensaNatureza` e
  `PB03_FundamentoDispensaNatureza`, ambas em
  `PB01_NaturezaObjeto.blocosAdicionais`; os bookmarks
  `PB02_ModalidadeOuFundamento_BLOCO_DISPENSA_DE_LICITACAO` e
  `PB03_ModalidadeOuFundamento_BLOCO_DISPENSA_DE_LICITACAO` (hoje
  parágrafos simples de `add_bloco_variante`) passam a ser pares de
  sentinela envolvendo o `add_bloco_variante` aninhado — mesma técnica já
  usada em vários anexos do arquivo, aplicada aqui a um trecho de corpo.

- [ ] **Step 1: Escrever o teste que falha**

Adicionar a `tests/pb-elaboracao-projeto.test.mjs`:

```js
test("PB Seção 2.5/3: Dispensa + Elaboração de projeto cita art. 154 e a modalidade correta", async () => {
  const texto = await textoDoDocumento({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
  });
  assert.ok(texto.includes("por se tratar de elaboração de projeto, laudo ou anteprojeto por terceiros"));
  assert.ok(texto.includes("art. 154 do RLC"));
  assert.ok(texto.includes("art. 416, I, do RLC"));
  assert.ok(!texto.includes("por se tratar de obra ou serviço de engenharia de valor"));
});

test("PB Seção 2.5/3: Dispensa + Execução de obra mantém o texto de hoje (não regressão)", async () => {
  const texto = await textoDoDocumento({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB01_NaturezaObjeto: "Execução de obra ou serviço já projetado",
  });
  assert.ok(texto.includes("por se tratar de obra ou serviço de engenharia de valor"));
  assert.ok(!texto.includes("elaboração de projeto, laudo ou anteprojeto por terceiros"));
});

test("PB Seção 2.5: regime Licitação não fala de dispensa, com nenhuma das duas naturezas", async () => {
  const texto = await textoDoDocumento({
    PB01_RegimeContratacao: "Licitação",
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
    PB02_ModalidadeLicitatoria: "Pregão Eletrônico",
  });
  assert.ok(!texto.includes("dispensa de licitação, com fundamento"));
  assert.ok(texto.includes("O serviço em apreço será contratado na modalidade"));
});
```

- [ ] **Step 2: Rodar e confirmar que falha**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: FAIL nos 2 primeiros testes novos (o terceiro já passa hoje —
serve de proteção de não regressão explícita).

- [ ] **Step 3: Adicionar as duas bases novas a `blocosAdicionais`**

Em `js/dados-pb.js`, campo `PB01_NaturezaObjeto` (mesma linha das Tasks 1
e 3):

**Nota:** a Task 3 (Anexo X) está pausada — o array abaixo ainda não tem
`"PB_AnexoX"`. Este Find/Replace parte diretamente do estado deixado pela
Task 1. Se a Task 3 já tiver rodado quando esta tarefa for executada,
ajuste o Find para incluir `"PB_AnexoX"` também (ver nota na Task 3).

Find:
```js
      blocosAdicionais: ["PB13_QualificacaoTecnica", "PB14_CriteriosObra", "PB21_GarantiaObra", "PB_AnexoVIII", "PB28_MetodologiaCusto", "PB_CessaoDireitos"] },
```

Replace:
```js
      blocosAdicionais: ["PB13_QualificacaoTecnica", "PB14_CriteriosObra", "PB21_GarantiaObra", "PB_AnexoVIII", "PB28_MetodologiaCusto", "PB_CessaoDireitos", "PB02_FundamentoDispensaNatureza", "PB03_FundamentoDispensaNatureza"] },
```

- [ ] **Step 4: Aninhar a Seção 2.5 por natureza, dentro do bloco de Dispensa**

Find (`build_template_pb.py:189-199`):
```python
add_bloco_variante("PB02_ModalidadeOuFundamento", {
    "Licitação": [[
        "2.5. O serviço em apreço será contratado na modalidade ",
        ("PB02_ModalidadeLicitatoria", "(modalidade)"), ", nos termos do art. 3º do RLC.",
    ]],
    "Dispensa de Licitação": [
        "2.5. A contratação será realizada por dispensa de licitação, com fundamento no "
        "art. 416, I, do RLC, por se tratar de obra ou serviço de engenharia de valor "
        "global até R$ 140.000,00 (cento e quarenta mil reais).",
    ],
})
```

Replace:
```python
add_bloco_variante("PB02_ModalidadeOuFundamento", {
    "Licitação": [[
        "2.5. O serviço em apreço será contratado na modalidade ",
        ("PB02_ModalidadeLicitatoria", "(modalidade)"), ", nos termos do art. 3º do RLC.",
    ]],
})
_bid_25_dispensa = add_bloco_variante_sentinela_inicio(
    "PB02_ModalidadeOuFundamento_BLOCO_DISPENSA_DE_LICITACAO")
add_bloco_variante("PB02_FundamentoDispensaNatureza", {
    "Execução de obra ou serviço já projetado": [
        "2.5. A contratação será realizada por dispensa de licitação, com fundamento no "
        "art. 416, I, do RLC, por se tratar de obra ou serviço de engenharia de valor "
        "global até R$ 140.000,00 (cento e quarenta mil reais).",
    ],
    "Elaboração de projeto, laudo ou anteprojeto por terceiros": [
        "2.5. A contratação será realizada por dispensa de licitação, com fundamento no "
        "art. 416, I, do RLC, por se tratar de elaboração de projeto, laudo ou "
        "anteprojeto por terceiros — modalidade de serviço de engenharia sujeita à "
        "elaboração de Projeto Básico nos termos do art. 154 do RLC —, de valor "
        "global até R$ 140.000,00 (cento e quarenta mil reais).",
    ],
})
add_bloco_variante_sentinela_fim(_bid_25_dispensa)
```

**Por que isso funciona sem duplicar o parágrafo:** o bookmark de
sentinela `PB02_ModalidadeOuFundamento_BLOCO_DISPENSA_DE_LICITACAO` agora
não é mais criado automaticamente por `add_bloco_variante` (a chave
"Dispensa de Licitação" saiu do dicionário) — é criado manualmente pelo
par `sentinela_inicio/fim`, com o **mesmo nome** que
`gerador-docx.js:198-212` já procura (via `PB01_RegimeContratacao.
blocosAdicionais`, que já inclui `"PB02_ModalidadeOuFundamento"` — ver
`js/dados-pb.js:22-25`, não precisa mexer). Quando o regime é "Licitação",
esse bookmark é removido inteiro (junto com o `add_bloco_variante`
aninhado por dentro, não importa a natureza). Quando o regime é
"Dispensa de Licitação", o bookmark nunca é alvo de remoção pelo laço do
regime — sobra o `add_bloco_variante` aninhado, que aí sim reage à
natureza normalmente.

- [ ] **Step 5: Mesmo aninhamento na Seção 3**

Find (`build_template_pb.py:245-251`):
```python
add_bloco_variante("PB03_ModalidadeOuFundamento", {
    "Licitação": [["Modalidade licitatória: ", ("PB03_ModalidadeLicitatoria", "")]],
    "Dispensa de Licitação": [
        "Fundamento da dispensa: art. 416, I, do RLC (obras e serviços de engenharia de "
        "valor global até R$ 140.000,00).",
    ],
})
```

Replace:
```python
add_bloco_variante("PB03_ModalidadeOuFundamento", {
    "Licitação": [["Modalidade licitatória: ", ("PB03_ModalidadeLicitatoria", "")]],
})
_bid_3_dispensa = add_bloco_variante_sentinela_inicio(
    "PB03_ModalidadeOuFundamento_BLOCO_DISPENSA_DE_LICITACAO")
add_bloco_variante("PB03_FundamentoDispensaNatureza", {
    "Execução de obra ou serviço já projetado": [
        "Fundamento da dispensa: art. 416, I, do RLC (obras e serviços de engenharia de "
        "valor global até R$ 140.000,00).",
    ],
    "Elaboração de projeto, laudo ou anteprojeto por terceiros": [
        "Fundamento da dispensa: art. 416, I, do RLC (elaboração de projeto, laudo ou "
        "anteprojeto por terceiros, modalidade de serviço de engenharia, de valor global "
        "até R$ 140.000,00).",
    ],
})
add_bloco_variante_sentinela_fim(_bid_3_dispensa)
```

- [ ] **Step 6: Regenerar o `.docx`-modelo**

```bash
cd templates
python build_template_pb.py
cd ..
```

- [ ] **Step 7: Rodar e confirmar que passa**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: PASS — 8 testes (5 das Tasks 1-2 + 3 desta tarefa; a Task 3
está pausada, ver nota no início desta tarefa).

- [ ] **Step 8: Rodar a suíte inteira, com atenção especial a `pb-dispensa.test.mjs`**

Run: `node --test tests/*.test.mjs`
Expected: PASS — em particular, os 6 testes de
`tests/pb-dispensa.test.mjs` continuam passando (eles usam
`PB01_NaturezaObjeto` implicitamente ausente das `respostas`, o que cai
no `padrao: "Execução de obra ou serviço já projetado"` declarado em
`js/dados-pb.js:14` — o texto de fundamentação da dispensa esperado por
esses testes é exatamente o de "obra", que não mudou).

- [ ] **Step 9: Commit**

```bash
git add js/dados-pb.js templates/build_template_pb.py templates/Template_PB.docx tests/pb-elaboracao-projeto.test.mjs
git commit -m "PB: fundamentação da dispensa (Seções 2.5 e 3) condicionada à natureza do objeto"
```

---

### Task 5: Fixture de regressão SUREG/SC — os quatro achados juntos

**Files:**
- Test: `tests/pb-elaboracao-projeto.test.mjs` (adicionar teste final)

**Interfaces:**
- Consumes: `gerarDocx` (Tasks 1, 2, 4 — a Task 3/Anexo X está pausada,
  ver nota abaixo), `SECOES_PB`.
- Produces: nada de novo — é o teste de integração que prova que os
  achados já implementados funcionam juntos, com os dados reais do
  processo SUREG/SC usados no teste manual do site.

**Nota — Task 3 pausada:** este teste cobre só os Achados 1, 2 e 4 por
enquanto (a fixture não inclui os campos `PB28_Fator*`, que ainda não
existem em `js/dados-pb.js`, e não faz nenhuma asserção sobre o Anexo X).
Quando a Task 3 for retomada, ela deve **acrescentar** ao objeto de
respostas e às asserções abaixo — não recriar o teste do zero — os itens
marcados com "// Achado 3 (Task 3, pausada)".

- [ ] **Step 1: Escrever o teste de integração com a fixture completa**

Adicionar a `tests/pb-elaboracao-projeto.test.mjs`:

```js
test("PB: fixture real SUREG/SC (dispensa + elaboração de projeto) gera os achados corretamente", async () => {
  const respostasSuregSC = {
    PB01_DescricaoObjeto: "Elaboração de levantamento métrico, diagnóstico institucional, estudo de layout e projetos executivos de arquitetura, estrutura, instalações hidrossanitárias, infraestrutura elétrica/lógica/TI, sistema audiovisual e ampliação de área externa, para as edificações administrativas e área externa da SUREG/SC.",
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB01_ValorEstimado: "35.500,00",
    PB01_ValorPorExtenso: "trinta e cinco mil e quinhentos reais",
    PB04_LocaisExecucao: "SUREG/SC — Superintendência Regional de Santa Catarina, situada na Rua Francisco Pedro Machado, s/n, Barreiros, São José/SC, CEP 88.117-402 (edificações administrativas e área externa).",
    PB07_PrazoTotalExecucao: "150 dias corridos",
    PB08_PrazoVigencia: "240 dias corridos",
    PB12_PrazoARTDias: "15",
    PB12_PrazoBaixaARTDias: "15",
    PB16_ElegivelRecebimentoSimplificado: "Sim",
    // Achado 3 (Task 3, pausada): PB28_FatorK1: "3,5", PB28_FatorK2: "2,0", PB28_FatorK3: "1,2", PB28_FatorK4: "8,0",
    // Achado 3 (Task 3, pausada): PB28_FatorKResultante: "14,7", PB28_TRDE: "10",
    // Achado 3 (Task 3, pausada): PB28_MetodologiaCustoIntelectual: "Correlação entre os coeficientes K1-K4 e a TRDE, própria da engenharia consultiva.",
  };
  const texto = await textoDoDocumento(respostasSuregSC);

  // Achado 1 — Seção 13
  assert.ok(texto.includes("Anexo VIII-B"));
  assert.ok(!texto.includes("parcelas de maior relevância técnica e econômica do objeto"));

  // Achado 2 — Seção 16
  assert.ok(texto.includes("art. 534 do RLC"));
  assert.ok(texto.includes("Recibo Definitivo Simplificado"));

  // Achado 3 (Task 3, pausada):
  // assert.ok(texto.includes("ANEXO X DO PROJETO BÁSICO"));
  // assert.ok(texto.includes("14,7"));

  // Achado 4 — fundamentação da dispensa
  assert.ok(texto.includes("elaboração de projeto, laudo ou anteprojeto por terceiros"));
  assert.ok(texto.includes("art. 154 do RLC"));

  // Dados do próprio caso continuam presentes (sem regressão dos campos existentes)
  assert.ok(texto.includes("SUREG/SC"));
  assert.ok(texto.includes("35.500,00") || texto.includes("35500"));
});
```

- [ ] **Step 2: Rodar e confirmar que passa de primeira**

Run: `node --test tests/pb-elaboracao-projeto.test.mjs`
Expected: PASS — 9 testes (8 das Tasks 1, 2 e 4, mais este).
Se algum assert falhar aqui apesar das Tasks 1, 2 e 4 passarem
isoladamente, é sinal de uma interação entre achados não prevista (ex.:
um bloco condicional comendo o parágrafo de outro) — depurar antes de
prosseguir, não ajustar o teste para "passar de qualquer jeito".

- [ ] **Step 3: Rodar a suíte inteira do projeto uma última vez**

Run: `node --test tests/*.test.mjs`
Expected: PASS — todos os arquivos de teste, todos os instrumentos.

- [ ] **Step 4: Commit**

```bash
git add tests/pb-elaboracao-projeto.test.mjs
git commit -m "PB: teste de integração com a fixture real SUREG/SC (dispensa + elaboração de projeto)"
```

---

## Self-Review (contra a spec `2026-08-25-pb-elaboracao-projeto-design.md`)

1. **Cobertura da spec:** Achado 1 → Task 1. Achado 2 → Task 2. Achado 3
   → Task 3. Achado 4 → Task 4. Fixture SUREG/SC → Task 5. Sem lacunas.
2. **Placeholders:** nenhum "TBD"/"implementar depois" — todo texto legal
   e todo trecho de código está escrito por extenso em cada Step.
3. **Consistência de nomes:** `PB13_QualificacaoTecnica`,
   `PB16_ElegivelRecebimentoSimplificado`, `PB_AnexoX`,
   `PB02_FundamentoDispensaNatureza`/`PB03_FundamentoDispensaNatureza`
   usados de forma idêntica entre `js/dados-pb.js`,
   `templates/build_template_pb.py` e os testes em todas as tarefas onde
   aparecem.
