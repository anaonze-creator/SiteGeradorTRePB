# Projeto Básico por Dispensa Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend the existing PB instrument (not a new instrument) to cover contratação by dispensa de licitação (art. 416, I, do RLC), adding a `PB01_RegimeContratacao` field that swaps a handful of body clauses and gates a new Minuta de Contrato annex.

**Architecture:** Same declarative pattern already in use — new fields in `js/dados-pb.js`, matching changes in `templates/build_template_pb.py`. Reuses `blocoTexto`/`blocosAdicionais` (structural block removal), `duplicarPara` (cross-filling already-answered fields into the new annex), and sentinela wrapping (the annex spans many paragraphs). One small, backward-compatible generalization: `add_bloco_variante` in `build_template_pb.py` gains support for bookmark tuples inside variant paragraphs (today it only accepts plain strings) — every existing call site keeps working unchanged.

**Tech Stack:** Same as the rest of the project — vanilla JS, python-docx, Node's built-in test runner.

## Global Constraints

- No source minuta exists for this variant — content is derived directly from RLC articles 416, I; 435, II; 438; 496 and cross-checked in `docs/superpowers/specs/2026-08-18-pb-dispensa-design.md`. Do not add content beyond what that spec grounds.
- PB por Dispensa never offers Contrato Simplificado (art. 435, II only covers services ≤ R$14.000) — the annex is a single unconditional Minuta de Contrato, no internal Sim/Não variant like TR-Dispensa's Anexo V/VI split.
- Reuse TR-Dispensa's Anexo VI wording verbatim for RLC-generic clauses (legislação aplicável, foro, vedações, nepotismo, alteração de contrato, alteração subjetiva) — these aren't TR-specific, just swap "Termo de Referência" for "Projeto Básico" and "Termo de Contrato" wording as needed.
- Portuguese for all code comments/commit messages, matching the existing codebase.

---

### Task 1: Generalize `add_bloco_variante` and add `PB01_RegimeContratacao`

**Files:**
- Modify: `templates/build_template_pb.py`
- Modify: `js/dados-pb.js`

**Interfaces:**
- Produces: `add_bloco_variante(nome_base, variantes)` in `build_template_pb.py` now accepts `str` or `(nome_bookmark, texto_inicial)` tuples inside each variant's paragraph list — same signature, backward-compatible.
- Produces: `PB01_RegimeContratacao` field (`js/dados-pb.js`, Seção "01"), values `"Licitação"` / `"Dispensa de Licitação"`, `blocoTexto: true`, `padrao: "Licitação"`.

- [ ] **Step 1: Generalize `add_bloco_variante` in `build_template_pb.py`**

Find the existing function (near the top of the file, after `normalizar_variante`):

```python
def add_bloco_variante(nome_base, variantes):
    """variantes: dict {valor_opcao: [paragrafo1, paragrafo2, ...]}. Cada
    variante fica dentro do seu próprio bookmark
    (<nome_base>_BLOCO_<VARIANTE_NORMALIZADA>) — o gerador remove todas
    as variantes não escolhidas."""
    for opcao, paragrafos in variantes.items():
        nome_bookmark = f"{nome_base}_BLOCO_{normalizar_variante(opcao)}"
        _bookmark_id[0] += 1
        bid = str(_bookmark_id[0])
        for i, texto in enumerate(paragrafos):
            p = doc.add_paragraph()
            if i == 0:
                start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
                start.set(docx.oxml.ns.qn("w:id"), bid)
                start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
                p._p.insert(0, start)
            p.add_run(texto)
            if i == len(paragrafos) - 1:
                end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
                end.set(docx.oxml.ns.qn("w:id"), bid)
                p._p.append(end)
```

Replace the paragraph body so each item in a variant's list can be either
a plain string or a `(nome_campo, texto_inicial)` tuple wrapped in its own
inner bookmark, mirroring `add_paragraph_with_bookmarks`:

```python
def add_bloco_variante(nome_base, variantes):
    """variantes: dict {valor_opcao: [paragrafo1, paragrafo2, ...]}. Cada
    item de paragrafos pode ser uma string (texto fixo) ou uma lista de
    partes no mesmo formato de add_paragraph_with_bookmarks (strings e
    tuplas (nome_campo, texto_inicial)), para permitir um bookmark de
    campo dentro de uma variante. Cada variante fica dentro do seu
    próprio bookmark (<nome_base>_BLOCO_<VARIANTE_NORMALIZADA>) — o
    gerador remove todas as variantes não escolhidas."""
    for opcao, paragrafos in variantes.items():
        nome_bookmark = f"{nome_base}_BLOCO_{normalizar_variante(opcao)}"
        _bookmark_id[0] += 1
        bid = str(_bookmark_id[0])
        for i, item in enumerate(paragrafos):
            partes = item if isinstance(item, list) else [item]
            p = doc.add_paragraph()
            if i == 0:
                start = docx.oxml.shared.OxmlElement("w:bookmarkStart")
                start.set(docx.oxml.ns.qn("w:id"), bid)
                start.set(docx.oxml.ns.qn("w:name"), nome_bookmark)
                p._p.insert(0, start)
            for parte in partes:
                if isinstance(parte, tuple):
                    nome_campo, texto_inicial = parte
                    _bookmark_id[0] += 1
                    bid_campo = str(_bookmark_id[0])
                    start_campo = docx.oxml.shared.OxmlElement("w:bookmarkStart")
                    start_campo.set(docx.oxml.ns.qn("w:id"), bid_campo)
                    start_campo.set(docx.oxml.ns.qn("w:name"), nome_campo)
                    p._p.append(start_campo)
                    p.add_run(texto_inicial)
                    end_campo = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
                    end_campo.set(docx.oxml.ns.qn("w:id"), bid_campo)
                    p._p.append(end_campo)
                else:
                    p.add_run(parte)
            if i == len(paragrafos) - 1:
                end = docx.oxml.shared.OxmlElement("w:bookmarkEnd")
                end.set(docx.oxml.ns.qn("w:id"), bid)
                p._p.append(end)
```

- [ ] **Step 2: Add the field to `js/dados-pb.js`**

In Seção "01", right after the `PB01_NaturezaObjeto` field (which already
uses `blocosAdicionais` — follow the same pattern), add:

```js
    { id: "PB01_RegimeContratacao", tipo: "parametro", input: "select",
      pergunta: "O regime de contratação é",
      opcoes: ["Licitação", "Dispensa de Licitação"],
      blocoTexto: true, padrao: "Licitação",
      ajuda: "Art. 416, I, do RLC: obras e serviços de engenharia de valor global até R$ 140.000,00 podem ser dispensados de licitação (desde que não se refiram a parcelas de uma mesma obra/serviço que pudessem ser executadas em conjunto). Art. 435, II, do RLC: essa dispensa NÃO admite Contrato Simplificado (restrito a serviços de pronto pagamento até R$ 14.000) — segue sempre com o Contrato completo.",
      blocosAdicionais: [
        "PB02_ModalidadeOuFundamento", "PB03_ModalidadeOuFundamento",
        "PB09_VistoriaPrazo", "PB34_RevogacaoOuAnulacao", "PB_AnexoContrato",
      ] },
```

Then find the existing `PB02_ModalidadeLicitatoria` field (Seção "02")
and add `mostrarSe`:

```js
    { id: "PB02_ModalidadeLicitatoria", tipo: "parametro", input: "select",
      pergunta: "Modalidade licitatória", opcoes: ["Pregão Eletrônico", "Competição Pública", "RDC"],
      mostrarSe: { PB01_RegimeContratacao: "Licitação" },
      duplicarPara: ["PB03_ModalidadeLicitatoria"] },
```

(Only the `mostrarSe` line is new — `duplicarPara` already existed.)

- [ ] **Step 3: Verify the Python script still runs**

Run: `cd templates && python build_template_pb.py`
Expected: `Template_PB.docx gerado com sucesso.` (the new field has no
template usage yet — this step just confirms the generalized
`add_bloco_variante` didn't break existing calls)

- [ ] **Step 4: Run the full test suite**

Run: `node --test tests/*.test.mjs`
Expected: PASS (no regressions — `PB01_RegimeContratacao` isn't
referenced by the template yet, and `mostrarSe` only hides a field in
the UI, it doesn't change generation)

- [ ] **Step 5: Commit**

```bash
git add templates/build_template_pb.py js/dados-pb.js
git commit -m "Generalize add_bloco_variante and add PB01_RegimeContratacao field"
```

---

### Task 2: Wire the 4 body-clause variants

**Files:**
- Modify: `templates/build_template_pb.py`

**Interfaces:**
- Consumes: `PB01_RegimeContratacao`, `PB02_ModalidadeLicitatoria`, `PB03_ModalidadeLicitatoria` (existing bookmark).

- [ ] **Step 1: Seção 2.5 (modalidade/fundamento) and Seção 3 (linha "Modalidade licitatória")**

Find:

```python
add_paragraph_with_bookmarks([
    "2.5. O serviço em apreço será contratado na modalidade ", ("PB02_ModalidadeLicitatoria", "(modalidade)"),
    ", nos termos do art. 3º do RLC.",
])
```

Replace with:

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

Find:

```python
add_paragraph_with_bookmarks(["Modalidade licitatória: ", ("PB03_ModalidadeLicitatoria", "")])
```

Replace with:

```python
add_bloco_variante("PB03_ModalidadeOuFundamento", {
    "Licitação": [["Modalidade licitatória: ", ("PB03_ModalidadeLicitatoria", "")]],
    "Dispensa de Licitação": [
        "Fundamento da dispensa: art. 416, I, do RLC (obras e serviços de engenharia de "
        "valor global até R$ 140.000,00).",
    ],
})
```

- [ ] **Step 2: Seção 9.3 (vistoria)**

Find:

```python
doc.add_paragraph(
    "9.3. O prazo para vistoria iniciar-se-á no dia útil seguinte ao da publicação do Edital "
    "e estender-se-á até o dia útil anterior à data prevista para a abertura da sessão "
    "pública.")
```

Replace with:

```python
add_bloco_variante("PB09_VistoriaPrazo", {
    "Licitação": [
        "9.3. O prazo para vistoria iniciar-se-á no dia útil seguinte ao da publicação do "
        "Edital e estender-se-á até o dia útil anterior à data prevista para a abertura da "
        "sessão pública.",
    ],
    "Dispensa de Licitação": [
        "9.3. O prazo para a vistoria será definido no processo de dispensa de licitação, "
        "conforme cronograma estabelecido pela área demandante.",
    ],
})
```

- [ ] **Step 3: Seção 34.2 (revogação/anulação)**

Find:

```python
doc.add_paragraph("34.2. A Conab poderá revogar a licitação, por interesse público, ou anulá-la por ilegalidade, mediante despacho fundamentado.")
```

Replace with:

```python
add_bloco_variante("PB34_RevogacaoOuAnulacao", {
    "Licitação": [
        "34.2. A Conab poderá revogar a licitação, por interesse público, ou anulá-la por "
        "ilegalidade, mediante despacho fundamentado.",
    ],
    "Dispensa de Licitação": [
        "34.2. A Conab poderá anular o processo de dispensa por ilegalidade, mediante "
        "despacho fundamentado, assegurados o contraditório e a ampla defesa.",
    ],
})
```

- [ ] **Step 4: Run the build script and full test suite**

Run: `cd templates && python build_template_pb.py`
Expected: `Template_PB.docx gerado com sucesso.`

Run: `node --test tests/*.test.mjs`
Expected: PASS. If any PB test fails because it searched for the old
literal text (e.g. `"O serviço em apreço será contratado na modalidade"`
without a following value), that's expected — those tests still pass
because the "Licitação" variant reproduces the same wording; only fix a
test if it breaks for a reason unrelated to this refactor.

- [ ] **Step 5: Commit**

```bash
git add templates/build_template_pb.py
git commit -m "Add Licitação/Dispensa variants for modalidade, vistoria and revogação clauses"
```

---

### Task 3: Add annex-only fields and `duplicarPara` cross-fills to `js/dados-pb.js`

**Files:**
- Modify: `js/dados-pb.js`

**Interfaces:**
- Produces: 3 new fields (`PB_ND_AnexoContrato`, `PB_PTRES_AnexoContrato`, `PB_Fonte_AnexoContrato`) and `duplicarPara` additions on 14 existing fields, all targeting bookmark names Task 4's Python script will define.

- [ ] **Step 1: Add the 3 new dotação orçamentária fields**

Right after `PB01_RegimeContratacao` (added in Task 1), add:

```js
    { id: "PB_ND_AnexoContrato", tipo: "parametro", input: "texto",
      pergunta: "Natureza de Despesa (ND) — para a Minuta de Contrato",
      mostrarSe: { PB01_RegimeContratacao: "Dispensa de Licitação" } },
    { id: "PB_PTRES_AnexoContrato", tipo: "parametro", input: "texto",
      pergunta: "PTRES — para a Minuta de Contrato",
      mostrarSe: { PB01_RegimeContratacao: "Dispensa de Licitação" } },
    { id: "PB_Fonte_AnexoContrato", tipo: "parametro", input: "texto",
      pergunta: "Fonte de recurso — para a Minuta de Contrato",
      mostrarSe: { PB01_RegimeContratacao: "Dispensa de Licitação" } },
```

- [ ] **Step 2: Add `duplicarPara` to the 14 existing fields the annex cross-fills**

Each bullet below is a one-line change — add (or extend, if `duplicarPara`
already exists) the array on that exact field. Field locations (by
`pergunta` text, to find them unambiguously — every id below already
exists in the file from earlier sessions):

- `PB01_DescricaoObjeto` (Seção 01) — extend the existing `duplicarPara`
  array to also include `"PB01_DescricaoObjeto_AnexoContrato"`.
- `PB01_RegimeExecucao` (Seção 01) — extend `duplicarPara: ["PB03_RegimeExecucao"]`
  to `duplicarPara: ["PB03_RegimeExecucao", "PB01_RegimeExecucao_AnexoContrato"]`.
- `PB01_TipoObjeto` (Seção 01) — extend `duplicarPara: ["PB03_TipoObjeto"]`
  to `duplicarPara: ["PB03_TipoObjeto", "PB01_TipoObjeto_AnexoContrato"]`.
- `PB08_PrazoVigencia` (Seção 08) — add `duplicarPara: ["PB08_PrazoVigencia_AnexoContrato"]`.
- `PB12_PrazoARTDias` (Seção 12) — add `duplicarPara: ["PB12_PrazoARTDias_AnexoContrato"]`.
- `PB12_PrazoBaixaARTDias` (Seção 12) — add `duplicarPara: ["PB12_PrazoBaixaARTDias_AnexoContrato"]`.
- `PB16_PrazoRecebimentoProvisorioDias` (Seção 16) — add `duplicarPara: ["PB16_PrazoRecebimentoProvisorioDias_AnexoContrato"]`.
- `PB16_PrazoRecebimentoDefinitivoDias` (Seção 16) — add `duplicarPara: ["PB16_PrazoRecebimentoDefinitivoDias_AnexoContrato"]`.
- `PB18_PrazoARTContratoDias` (Seção 18) — add `duplicarPara: ["PB18_PrazoARTContratoDias_AnexoContrato"]`.
- `PB19_PrazoPagamentoDiasUteis` (Seção 19) — add `duplicarPara: ["PB19_PrazoPagamentoDiasUteis_AnexoContrato"]`.
- `PB20_IndiceReajuste` (Seção 20) — add `duplicarPara: ["PB20_IndiceReajuste_AnexoContrato"]`.
- `PB21_PrazoGarantiaServico` (Seção 21) — add `duplicarPara: ["PB21_PrazoGarantiaServico_AnexoContrato"]`.
- `PB22_PercentualGarantia` (Seção 22) — add `duplicarPara: ["PB22_PercentualGarantia_AnexoContrato"]`.
- `PB22_PrazoGarantiaDias` (Seção 22) — add `duplicarPara: ["PB22_PrazoGarantiaDias_AnexoContrato"]`.
- `PB25_MultaMoratoriaPercentualDia` (Seção 25) — extend
  `duplicarPara: ["PB25_MultaMoratoriaPercentualDia_ANEXO_VII"]` to also
  include `"PB25_MultaMoratoriaPercentualDia_AnexoContrato"`.
- `PB25_MultaMoratoriaLimite` (Seção 25) — add `duplicarPara: ["PB25_MultaMoratoriaLimite_AnexoContrato"]`.
- `PB25_MultaCompensatoria` (Seção 25) — add `duplicarPara: ["PB25_MultaCompensatoria_AnexoContrato"]`.
- `PB25_MultaRescisoria` (Seção 25) — add `duplicarPara: ["PB25_MultaRescisoria_AnexoContrato"]`.

- [ ] **Step 2: Syntax-check**

Run: `node --check js/dados-pb.js`
Expected: no output (valid syntax)

- [ ] **Step 3: Commit**

```bash
git add js/dados-pb.js
git commit -m "Add dotação orçamentária fields and duplicarPara targets for the Minuta de Contrato annex"
```

---

### Task 4: Build the Minuta de Contrato annex in `build_template_pb.py`

**Files:**
- Modify: `templates/build_template_pb.py`

**Interfaces:**
- Consumes: every `_AnexoContrato`-suffixed bookmark name from Task 3, plus `PB_AnexoContrato` from Task 1's `blocosAdicionais`, plus `add_bloco_variante_sentinela_inicio`/`_fim` (already defined in this file from the `PB01_NaturezaObjeto` work).

- [ ] **Step 1: Add the annex at the end of the file, before `doc.save(...)`**

Find the end of the file (the last anexo — Anexo VIII placa de obra —
followed by `doc.save("Template_PB.docx")`). Insert this new annex right
before the `doc.save(...)` line:

```python
# ANEXO — MINUTA DE CONTRATO (só existe se Dispensa de Licitação; art.
# 438 do RLC — 30 partes essenciais da minuta de Contrato — e art. 435,
# II, que impede o Contrato Simplificado para obras/serviços de
# engenharia). Sem minuta SEI de referência para esta variante — texto
# derivado diretamente do RLC, reaproveitando a redação já validada do
# Anexo VI do TR-Dispensa para as cláusulas RLC-genéricas. Ver
# docs/superpowers/specs/2026-08-18-pb-dispensa-design.md.
_bid_anexo_contrato = add_bloco_variante_sentinela_inicio("PB_AnexoContrato_BLOCO_DISPENSA_DE_LICITACAO")
doc.add_page_break()
add_heading("ANEXO DO PROJETO BÁSICO")
add_heading("MINUTA DE CONTRATO")
add_paragraph_with_bookmarks([
    "CONTRATO QUE ENTRE SI CELEBRAM A COMPANHIA NACIONAL DE ABASTECIMENTO - CONAB E A EMPRESA "
    "___________________ PARA A CONTRATAÇÃO DE ", ("PB01_TipoObjeto_AnexoContrato", "(obra/serviço)"),
    ", de conformidade com o Processo Administrativo referente à Dispensa de Licitação, "
    "regendo-se pelo Projeto Básico e seus anexos e pela proposta da Contratada, pelo RLC, "
    "pela Lei nº 13.303, de 2016, e demais legislações pertinentes.",
])

doc.add_paragraph("CLÁUSULA PRIMEIRA - DO OBJETO")
add_paragraph_with_bookmarks([
    "O objeto do presente instrumento é a contratação de ",
    ("PB01_DescricaoObjeto_AnexoContrato", ""),
    ", nas condições estabelecidas no Projeto Básico, ao qual este Termo de Contrato se "
    "vincula, independentemente de transcrição, juntamente com a proposta vencedora.",
])

doc.add_paragraph("CLÁUSULA SEGUNDA - DA VIGÊNCIA")
add_paragraph_with_bookmarks([
    "O prazo de vigência do Contrato será de ", ("PB08_PrazoVigencia_AnexoContrato", "____"),
    ", contado da data de sua assinatura.",
])

doc.add_paragraph("CLÁUSULA TERCEIRA - DO REGIME DE EXECUÇÃO, DOS PRAZOS E DO RECEBIMENTO")
add_paragraph_with_bookmarks([
    "Os serviços/obras serão executados no regime de ", ("PB01_RegimeExecucao_AnexoContrato", ""),
    ", com recebimento provisório em até ", ("PB16_PrazoRecebimentoProvisorioDias_AnexoContrato", "____"),
    " dias corridos do adimplemento da parcela/entrega, e recebimento definitivo em até ",
    ("PB16_PrazoRecebimentoDefinitivoDias_AnexoContrato", "____"),
    " dias corridos do recebimento provisório, observado o Instrumento de Medição de "
    "Resultado (IMR), Anexo VII do Projeto Básico, quando for o caso.",
])

doc.add_paragraph("CLÁUSULA QUARTA - DO VALOR DO CONTRATO")
doc.add_paragraph(
    "O valor do Contrato corresponde ao valor proposto pela Contratada vencedora, incluídas "
    "todas as despesas ordinárias diretas e indiretas decorrentes da execução do objeto, "
    "inclusive tributos, encargos sociais, trabalhistas e previdenciários.")

doc.add_paragraph("CLÁUSULA QUINTA - DA GARANTIA CONTRATUAL")
add_paragraph_with_bookmarks([
    "A Contratada deverá apresentar, no prazo de ", ("PB22_PrazoGarantiaDias_AnexoContrato", "____"),
    " dias corridos após a assinatura do Contrato, comprovante de prestação de garantia "
    "correspondente a ", ("PB22_PercentualGarantia_AnexoContrato", "____"),
    "% (percentual), conforme parâmetros do art. 439 do RLC, mediante caução em dinheiro, "
    "seguro-garantia ou fiança bancária.",
])

doc.add_paragraph("CLÁUSULA SEXTA - DO PRAZO DE GARANTIA DA OBRA OU DO SERVIÇO DE ENGENHARIA")
add_paragraph_with_bookmarks([
    "O prazo de garantia dos serviços/obra, contado do recebimento definitivo, é de ",
    ("PB21_PrazoGarantiaServico_AnexoContrato", "____"),
    ". Para obras de engenharia, o Código Civil já garante 5 (cinco) anos contra vícios "
    "ocultos (art. 618), independentemente deste prazo.",
])

doc.add_paragraph("CLÁUSULA SÉTIMA - DA DOTAÇÃO ORÇAMENTÁRIA")
add_paragraph_with_bookmarks([
    "A despesa orçamentária da execução deste Contrato correrá à conta da Natureza da "
    "Despesa (", ("PB_ND_AnexoContrato", "XXXX"), "), PTRES (", ("PB_PTRES_AnexoContrato", "XXXX"),
    ") e Fonte (", ("PB_Fonte_AnexoContrato", "XXXX"), "), conforme Nota de Empenho a ser emitida.",
])

doc.add_paragraph("CLÁUSULA OITAVA - DAS OBRIGAÇÕES DA CONTRATANTE")
doc.add_paragraph(
    "São obrigações da Contratante as estabelecidas no Projeto Básico, não respondendo a "
    "Conab por compromissos assumidos pela Contratada com terceiros.")

doc.add_paragraph("CLÁUSULA NONA - DAS OBRIGAÇÕES DA CONTRATADA")
add_paragraph_with_bookmarks([
    "São obrigações da Contratada as estabelecidas no Projeto Básico, incluindo a "
    "apresentação das ART/RRT registradas no CREA/CAU em até ",
    ("PB18_PrazoARTContratoDias_AnexoContrato", "____"), " dias corridos da assinatura.",
])

doc.add_paragraph("CLÁUSULA DÉCIMA - DA ANOTAÇÃO DE RESPONSABILIDADE TÉCNICA (ART/RRT)")
add_paragraph_with_bookmarks([
    "A Contratada apresentará a ART/RRT referente ao objeto em até ",
    ("PB12_PrazoARTDias_AnexoContrato", "____"),
    " dias corridos da assinatura do Contrato, e a respectiva baixa em até ",
    ("PB12_PrazoBaixaARTDias_AnexoContrato", "____"), " dias úteis após a finalização.",
])

doc.add_paragraph("CLÁUSULA DÉCIMA PRIMEIRA - DO PAGAMENTO")
add_paragraph_with_bookmarks([
    "O pagamento será efetuado em até ", ("PB19_PrazoPagamentoDiasUteis_AnexoContrato", "____"),
    " dias úteis após o recebimento definitivo e atesto da Nota Fiscal/Fatura.",
])

doc.add_paragraph("CLÁUSULA DÉCIMA SEGUNDA - DO REAJUSTAMENTO / REPACTUAÇÃO")
add_paragraph_with_bookmarks([
    "Observado o interregno mínimo de 12 (doze) meses, os preços serão reajustados/"
    "repactuados por solicitação da Contratada, pela variação do índice ",
    ("PB20_IndiceReajuste_AnexoContrato", ""), ".",
])

doc.add_paragraph("CLÁUSULA DÉCIMA TERCEIRA - DA FISCALIZAÇÃO E GERENCIAMENTO DO CONTRATO")
doc.add_paragraph(
    "O acompanhamento e a fiscalização do Contrato observarão o previsto no Projeto Básico.")

doc.add_paragraph("CLÁUSULA DÉCIMA QUARTA - DA MATRIZ DE RISCOS E QUADRO DE INOVAÇÕES")
doc.add_paragraph(
    "A Matriz de Riscos — Anexo II do Projeto Básico — e, quando aplicável, o Quadro de "
    "Inovações — Anexo II-A —, constituem peças integrantes deste Contrato, "
    "independentemente de transcrição, definindo os riscos e responsabilidades entre "
    "Contratante e Contratada.")

doc.add_paragraph("CLÁUSULA DÉCIMA QUINTA - DAS SANÇÕES ADMINISTRATIVAS")
doc.add_paragraph(
    "A Contratada, em caso de inadimplemento de suas obrigações, garantido o contraditório e "
    "ampla defesa, ficará sujeita às sanções previstas no RLC, na Lei nº 13.303, de 2016, e no "
    "Projeto Básico: advertência; multa moratória; multa compensatória; multa rescisória; e "
    "suspensão do direito de participar de licitação e impedimento de contratar com a Conab, "
    "por até 02 (dois) anos.")
add_paragraph_with_bookmarks([
    "a) multa moratória por dia de atraso: ",
    ("PB25_MultaMoratoriaPercentualDia_AnexoContrato", "____"), "% sobre o valor inadimplido, "
    "até o limite acumulado de ", ("PB25_MultaMoratoriaLimite_AnexoContrato", "____"), "%.",
])
add_paragraph_with_bookmarks([
    "b) multa compensatória: ", ("PB25_MultaCompensatoria_AnexoContrato", "____"),
    "% do valor total do Contrato, pelo descumprimento grave ou rescisão por culpa da "
    "Contratada.",
])
add_paragraph_with_bookmarks([
    "c) multa rescisória: ", ("PB25_MultaRescisoria_AnexoContrato", "____"),
    "% nos casos de rescisão unilateral por descumprimento contratual.",
])

doc.add_paragraph("CLÁUSULA DÉCIMA SEXTA - DA RESCISÃO")
doc.add_paragraph(
    "A rescisão observará o disposto no Projeto Básico e nos arts. 568 a 572 e 582 a 593 do "
    "RLC.")

doc.add_paragraph("CLÁUSULA DÉCIMA SÉTIMA - DA SUBCONTRATAÇÃO")
doc.add_paragraph("As condições de subcontratação, quando admitidas, observam o disposto no Projeto Básico.")

doc.add_paragraph("CLÁUSULA DÉCIMA OITAVA - DA ALTERAÇÃO DO CONTRATO")
doc.add_paragraph(
    "O presente Contrato poderá ser alterado nas hipóteses do art. 510 do RLC, vedando-se "
    "alterações que resultem em violação ao dever de licitar. A Contratada poderá aceitar "
    "acréscimos ou supressões de até 25% (vinte e cinco por cento) do valor inicial "
    "atualizado do contrato.")

doc.add_paragraph("CLÁUSULA DÉCIMA NONA - DA ALTERAÇÃO SUBJETIVA")
doc.add_paragraph(
    "É admissível a fusão, cisão ou incorporação da Contratada com/em outra pessoa jurídica, "
    "observados os requisitos de habilitação e a anuência expressa da Conab.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA - DAS VEDAÇÕES")
doc.add_paragraph(
    "É vedado à Contratada caucionar ou utilizar este Contrato para qualquer operação "
    "financeira; interromper a execução contratual sob alegação de inadimplemento da "
    "Contratante, salvo nos casos previstos em lei; empregar menor de 18 anos em trabalho "
    "noturno, perigoso ou insalubre e não empregar menor de 16 anos, salvo na condição de "
    "aprendiz a partir de 14 anos; e subcontratar, na íntegra, o objeto da contratação, "
    "salvo quando a subcontratação parcial for admitida no Projeto Básico.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA PRIMEIRA - DA VEDAÇÃO AO NEPOTISMO")
doc.add_paragraph(
    "Conforme disposto no parágrafo único do art. 12 do RLC e no art. 7º do Decreto nº 7.203, "
    "de 2010, fica vedada a contratação de empregado ou dirigente da Conab como pessoa "
    "física, ou de pessoa/empresa com relação de parentesco até o terceiro grau civil com "
    "autoridade do Ministério da Agricultura, dirigente ou empregado da Conab com "
    "atribuições na área de licitação e contratação.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA SEGUNDA - DA VINCULAÇÃO")
doc.add_paragraph(
    "Consideram-se integrantes do presente Instrumento Contratual o Projeto Básico da "
    "Dispensa de Licitação e seus Anexos, a Proposta da Contratada, no que couber, e demais "
    "documentos pertinentes, independentemente de transcrição.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA TERCEIRA - DA LEGISLAÇÃO APLICÁVEL")
doc.add_paragraph(
    "Os casos omissos serão decididos pela Contratante, segundo as disposições do RLC, da "
    "Lei nº 13.303, de 2016, e demais normas aplicáveis, e, subsidiariamente, do Código de "
    "Defesa do Consumidor.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA QUARTA - DA PROTEÇÃO DE DADOS PESSOAIS")
doc.add_paragraph(
    "As Partes obrigam-se a cumprir a legislação vigente sobre proteção de dados pessoais, "
    "em especial a Lei nº 13.709, de 2018 (Lei Geral de Proteção de Dados Pessoais), no que "
    "for aplicável à execução deste Contrato.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA QUINTA - DA PUBLICAÇÃO")
doc.add_paragraph(
    "A publicação do extrato do presente Contrato deverá ser providenciada pela Contratante "
    "até o 5º (quinto) dia útil do mês subsequente ao da sua assinatura no Diário Oficial da "
    "União, conforme o art. 480 do RLC.")

doc.add_paragraph("CLÁUSULA VIGÉSIMA SEXTA - DO FORO")
doc.add_paragraph(
    "As partes elegem o foro da Justiça Federal, Seção Judiciária do Estado da Sureg ou de "
    "Brasília-DF, no caso da Matriz, competente para dirimir quaisquer dúvidas suscitadas em "
    "razão deste Contrato, que não puderem ser resolvidas de comum acordo.")

doc.add_paragraph(
    "Por estarem justas e acordadas, firmam o presente instrumento, em 03 (três) vias, de "
    "igual teor e forma, diante das testemunhas abaixo identificadas.")
doc.add_paragraph("Pela Contratante: ___________________     Pela Contratada: ___________________")
doc.add_paragraph("Testemunha 1: ___________________     Testemunha 2: ___________________")
add_bloco_variante_sentinela_fim(_bid_anexo_contrato)
```

- [ ] **Step 2: Run the build script**

Run: `cd templates && python build_template_pb.py`
Expected: `Template_PB.docx gerado com sucesso.`

- [ ] **Step 3: Run the full test suite**

Run: `node --test tests/*.test.mjs`
Expected: PASS (no existing test exercises `PB01_RegimeContratacao`, so
nothing should break — the annex only appears when that field is set)

- [ ] **Step 4: Commit**

```bash
git add templates/build_template_pb.py templates/Template_PB.docx
git commit -m "Add Minuta de Contrato annex for PB por Dispensa"
```

---

### Task 5: End-to-end tests

**Files:**
- Create: `tests/pb-dispensa.test.mjs`

**Interfaces:**
- Consumes: `gerarDocx` from `js/gerador-docx.js`, `SECOES_PB` from `js/dados-pb.js`.

- [ ] **Step 1: Write the failing test**

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

async function gerarXmlPB(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate(), respostas, SECOES_PB);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function textoDoDocumento(xml) {
  return xml.replace(/<[^>]+>/g, "");
}

test("PB: regime Licitação (padrão) não altera o comportamento existente", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Licitação",
    PB02_ModalidadeLicitatoria: "Pregão Eletrônico",
  }));
  assert.ok(texto.includes("O serviço em apreço será contratado na modalidade"));
  assert.ok(texto.includes("Pregão Eletrônico"));
  assert.ok(texto.includes("Modalidade licitatória:"));
  assert.ok(texto.includes("A Conab poderá revogar a licitação"));
  assert.ok(!texto.includes("MINUTA DE CONTRATO"));
});

test("PB: regime Dispensa troca modalidade, vistoria e revogação, e adiciona o anexo de contrato", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB01_DescricaoObjeto: "reforma da SUREG/RO",
    PB08_PrazoVigencia: "12 meses",
    PB22_PercentualGarantia: "5",
  }));
  assert.ok(texto.includes("A contratação será realizada por dispensa de licitação, com fundamento no art. 416, I, do RLC"));
  assert.ok(texto.includes("Fundamento da dispensa: art. 416, I, do RLC"));
  assert.ok(texto.includes("O prazo para a vistoria será definido no processo de dispensa de licitação"));
  assert.ok(texto.includes("A Conab poderá anular o processo de dispensa por ilegalidade"));
  assert.ok(!texto.includes("A Conab poderá revogar a licitação"));
  assert.ok(texto.includes("MINUTA DE CONTRATO"));
  assert.ok(texto.includes("reforma da SUREG/RO"));
  assert.ok(texto.includes("CLÁUSULA VIGÉSIMA QUARTA - DA PROTEÇÃO DE DADOS PESSOAIS"));
  assert.ok(texto.includes("Lei nº 13.709, de 2018"));
});

test("PB: campo de modalidade licitatória some quando o regime é Dispensa", async () => {
  const respostas = { PB01_RegimeContratacao: "Dispensa de Licitação", PB02_ModalidadeLicitatoria: "Pregão Eletrônico" };
  const { campoVisivel } = await import("../js/modelo.js");
  const campoModalidade = SECOES_PB.flatMap((s) => s.campos).find((c) => c.id === "PB02_ModalidadeLicitatoria");
  assert.equal(campoVisivel(campoModalidade, respostas), false);
});
```

- [ ] **Step 2: Run the tests**

Run: `node --test tests/pb-dispensa.test.mjs`
Expected: PASS (3 tests). If any fail, fix `js/dados-pb.js` or
`templates/build_template_pb.py` — whichever bookmark/field mismatches —
before proceeding.

- [ ] **Step 3: Run the full suite to check for regressions**

Run: `node --test tests/*.test.mjs`
Expected: PASS (all existing + new tests, 113 total)

- [ ] **Step 4: Commit**

```bash
git add tests/pb-dispensa.test.mjs
git commit -m "Add end-to-end tests for PB por Dispensa"
```

---

## Execution Handoff

After all 5 tasks are complete, run `node --test tests/*.test.mjs` once
more to confirm the whole suite is green, then follow
`superpowers:finishing-a-development-branch` to merge (this plan follows
the project's worktree-per-feature convention — create
`.worktrees/pb-dispensa` off `master` before starting Task 1).
