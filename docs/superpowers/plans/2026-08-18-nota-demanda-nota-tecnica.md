# Nota de Demanda e Nota Técnica Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add two new instruments — Nota de Demanda and Nota Técnica — to the generator, sharing a 13-section skeleton (art. 108/115 do RLC) but with genuinely different question wording (functional vs "tecnicamente", per art. 117) and a Nota-Técnica-only conditional 4-section block for obras/serviços de engenharia (art. 123, V-VIII).

**Architecture:** Same declarative pattern as the 4 existing instruments — a `SECOES_*` data model in `js/`, a matching `templates/build_template_*.py` script producing a bookmarked `.docx`, and an entry in `js/app.js`'s `INSTRUMENTOS` registry plus a card in `js/triagem.js`. No new engine mechanism: reuses `blocoTexto` (Sim/Não and Viável/Inviável natural-sentence swaps), `mostrarSe` (detail fields gated on Sim), and `blocosAdicionais` (the 4-section engineering block).

**Tech Stack:** Same as the rest of the project — vanilla JS, python-docx for template generation, Node's built-in test runner.

## Global Constraints

- All RLC citations in `ajuda` text must match exactly what was verified against `documentos de referencia/10.901_Regulamento_RLC_30.12.2024_versao4.pdf` and recorded in `docs/superpowers/specs/2026-08-18-nota-demanda-nota-tecnica-design.md` — do not invent or approximate article numbers.
- Art. 115 is cited only in Seções 1, 4, 5, 9, 13 (the 5 sections where it has a real corresponding item); Seção 7 cites `art. 115, IV` (not V — V belongs to Seção 4).
- Nota Técnica's per-section question wording must use "tecnicamente"/technical framing per the spec's table — never reuse Nota de Demanda's wording verbatim for the 13 base sections.
- Portuguese for all code comments/commit messages, matching the existing codebase.

---

### Task 1: `js/dados-nota-demanda.js` — Nota de Demanda data model

**Files:**
- Create: `js/dados-nota-demanda.js`

**Interfaces:**
- Produces: `export const SECOES_NOTA_DEMANDA` — array of `{id, titulo, campos}`, same shape as `SECOES_TR_BENS` etc. Field ids all prefixed `ND`.

- [ ] **Step 1: Write the file**

```js
export const SECOES_NOTA_DEMANDA = [
  { id: "01", titulo: "NECESSIDADE DE CONTRATAÇÃO", campos: [
    { id: "ND01_Necessidade", tipo: "especificacao", input: "textarea",
      pergunta: "Qual é a situação-problema que originou a necessidade desta contratação, e quais os impactos para a Conab?",
      ajuda: "Art. 108, §1º, I, e art. 115, I, do RLC: a justificativa da necessidade deve ser fornecida pela área interessada na contratação, por ser quem melhor conhece o problema. Explique a situação-problema e seus impactos para a Conab." },
  ]},
  { id: "02", titulo: "REFERÊNCIA A OUTROS INSTRUMENTOS DE PLANEJAMENTO DA COMPANHIA", campos: [
    { id: "ND02_ReferenciaPlanejamento", tipo: "especificacao", input: "textarea",
      pergunta: "A contratação está alinhada a algum instrumento de planejamento da Conab (Plano de Desenvolvimento Institucional, Planejamento Estratégico) ou política pública?",
      ajuda: "Art. 108, §1º, II, do RLC: indique se a contratação está alinhada aos planos instituídos, e a política pública a que esteja vinculada, quando couber." },
  ]},
  { id: "03", titulo: "REQUISITOS DA CONTRATAÇÃO", campos: [
    { id: "ND03_Requisitos", tipo: "especificacao", input: "textarea",
      pergunta: "Quais os requisitos necessários ao atendimento da necessidade (incluindo critérios de sustentabilidade e vigência pretendida do Contrato)?",
      ajuda: "Art. 108, §1º, III, do RLC: elenque os requisitos necessários, critérios de sustentabilidade quando couber, e a duração pretendida do Contrato (limite de 5 anos para prestação de serviços), justificando a escolha." },
  ]},
  { id: "04", titulo: "ESTIMATIVA DAS QUANTIDADES, ACOMPANHADAS DAS MEMÓRIAS DE CÁLCULO E DOS DOCUMENTOS QUE LHE DÃO SUPORTE", campos: [
    { id: "ND04_EstimativaQuantidades", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva o método usado para estimar as quantidades do objeto, com a data prevista de execução.",
      ajuda: "Art. 108, §1º, IV, e art. 115, V, do RLC: defina e documente o método para a estimativa das quantidades, utilizando informações de contratações anteriores quando houver, e inclua a data prevista de execução do objeto." },
  ]},
  { id: "05", titulo: "LEVANTAMENTO DE MERCADO E JUSTIFICATIVA DE ESCOLHA DO TIPO DE SOLUÇÃO A CONTRATAR", campos: [
    { id: "ND05_LevantamentoMercado", tipo: "especificacao", input: "textarea",
      pergunta: "Qual solução foi identificada como a que melhor atende à contratação, e por quê?",
      ajuda: "Art. 108, §1º, V, e art. 115, II, do RLC: considere diferentes fontes (contratações similares de outros órgãos, pesquisas e estudos do setor) para justificar a escolha da solução." },
  ]},
  { id: "06", titulo: "ESTIMATIVAS DE PREÇOS OU PREÇOS REFERENCIAIS", campos: [
    { id: "ND06_EstimativaPrecos", tipo: "especificacao", input: "textarea",
      pergunta: "Qual metodologia foi usada para estimar o preço de referência?",
      ajuda: "Art. 108, §1º, VI, do RLC: utilize preferencialmente preços referenciais divulgados por órgãos públicos, justificando a escolha do método, e inclua nos autos a memória de cálculo." },
  ]},
  { id: "07", titulo: "DESCRIÇÃO DA SOLUÇÃO COMO UM TODO", campos: [
    { id: "ND07_DescricaoSolucao", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva a solução como um todo: os elementos que devem ser produzidos, contratados ou executados.",
      ajuda: "Art. 108, §1º, VII, e art. 115, IV, do RLC: apresente todos os elementos que devem ser produzidos, contratados ou executados para que a contratação produza os resultados esperados, com definição precisa e suficiente do objeto." },
  ]},
  { id: "08", titulo: "JUSTIFICATIVAS PARA O PARCELAMENTO OU NÃO DA SOLUÇÃO", campos: [
    { id: "ND08_Parcelamento", tipo: "parametro", input: "select",
      pergunta: "A solução será parcelada (licitada por item)?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Art. 108, §1º, VIII, do RLC: o parcelamento é a regra quando o objeto é divisível, desde que não haja prejuízo ao conjunto da solução ou perda de economia de escala." },
    { id: "ND08_Justificativa", tipo: "especificacao", input: "textarea",
      pergunta: "Justifique a decisão de parcelar ou não a solução" },
  ]},
  { id: "09", titulo: "DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS", campos: [
    { id: "ND09_ResultadosPretendidos", tipo: "especificacao", input: "textarea",
      pergunta: "Quais os benefícios/resultados pretendidos com a contratação, e a quem se destinam?",
      ajuda: "Art. 108, §1º, IX, e art. 115, VI, do RLC: declare os benefícios diretos e indiretos em termos de economicidade, eficácia, eficiência e melhor aproveitamento dos recursos disponíveis." },
  ]},
  { id: "10", titulo: "PROVIDÊNCIAS PARA ADEQUAÇÃO DO AMBIENTE DA CONAB", campos: [
    { id: "ND10_AdequacaoAmbiente", tipo: "parametro", input: "select",
      pergunta: "Há necessidade de adequação do ambiente da Conab para a execução da contratação?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Art. 108, §1º, X, do RLC: elabore cronograma das atividades necessárias e considere a capacitação de empregados, quando aplicável." },
    { id: "ND10_DetalheAdequacao", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva as adequações necessárias",
      mostrarSe: { ND10_AdequacaoAmbiente: "Sim" } },
  ]},
  { id: "11", titulo: "CONTRATAÇÕES CORRELATAS OU INTERDEPENDENTES", campos: [
    { id: "ND11_ContratacoesCorrelatas", tipo: "parametro", input: "select",
      pergunta: "Há contratos correlatos ou interdependentes a esta contratação?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "ND11_DetalheCorrelatas", tipo: "especificacao", input: "textarea",
      pergunta: "Liste os contratos correlatos ou interdependentes",
      mostrarSe: { ND11_ContratacoesCorrelatas: "Sim" } },
  ]},
  { id: "12", titulo: "DECLARAÇÃO DA VIABILIDADE OU NÃO DA CONTRATAÇÃO", campos: [
    { id: "ND12_Viabilidade", tipo: "parametro", input: "select",
      pergunta: "A contratação é viável?",
      opcoes: ["Viável", "Inviável"], blocoTexto: true, padrao: "Viável",
      ajuda: "Art. 108, §1º, XII, do RLC: a declaração deve ser justificada com base nos elementos averiguados nos itens anteriores dos Estudos Preliminares." },
    { id: "ND12_Justificativa", tipo: "especificacao", input: "textarea",
      pergunta: "Justifique a declaração de viabilidade" },
  ]},
  { id: "13", titulo: "IDENTIFICAÇÃO DA MODALIDADE DE LICITAÇÃO OU DE CONTRATAÇÃO DIRETA", campos: [
    { id: "ND13_Modalidade", tipo: "parametro", input: "select",
      pergunta: "Qual modalidade de licitação (ou contratação direta) será utilizada?",
      opcoes: ["Pregão Eletrônico", "Competição Pública", "RDC", "Dispensa de Licitação"],
      ajuda: "Art. 115, III, do RLC: identifique a modalidade e justifique sua adoção." },
    { id: "ND13_Justificativa", tipo: "especificacao", input: "textarea",
      pergunta: "Justifique a escolha da modalidade" },
  ]},
  { id: "14", titulo: "DISPOSIÇÕES FINAIS", campos: [
    { id: "ND14_ResponsavelNome", tipo: "parametro", input: "texto",
      pergunta: "Responsável pelo estudo — nome" },
    { id: "ND14_ResponsavelCargo", tipo: "parametro", input: "texto",
      pergunta: "Responsável pelo estudo — cargo/função" },
  ]},
];
```

- [ ] **Step 2: Commit**

```bash
git add js/dados-nota-demanda.js
git commit -m "Add Nota de Demanda data model (13 sections, art. 108/115 do RLC)"
```

---

### Task 2: `templates/build_template_nota_demanda.py` — Nota de Demanda template

**Files:**
- Create: `templates/build_template_nota_demanda.py`
- Generate: `templates/Template_Nota_Demanda.docx`

**Interfaces:**
- Consumes: field ids from Task 1 (`ND01_Necessidade` … `ND14_ResponsavelCargo`) — every bookmark name in this script must match one of those ids exactly, or `substituirBookmark` silently no-ops on it.
- Produces: `templates/Template_Nota_Demanda.docx`, loaded by `js/app.js` via `templateUrl`.

- [ ] **Step 1: Write the build script**

```python
"""Gera Template_Nota_Demanda.docx com bookmarks nos pontos de
preenchimento identificados em js/dados-nota-demanda.js. Estrutura de
13 seções (art. 108, §1º, I-XII, e art. 115, III, do RLC), redigida com
viés funcional/administrativo (área demandante), conforme
docs/superpowers/specs/2026-08-18-nota-demanda-nota-tecnica-design.md."""
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


add_heading("COMPANHIA NACIONAL DE ABASTECIMENTO")
add_heading("NOTA DE DEMANDA")

add_heading("1. NECESSIDADE DE CONTRATAÇÃO (ART. 108, §1º, I, E ART. 115, I, DO RLC)")
add_paragraph_with_bookmarks([("ND01_Necessidade", "")])

add_heading("2. REFERÊNCIA A OUTROS INSTRUMENTOS DE PLANEJAMENTO DA COMPANHIA (ART. 108, §1º, II, DO RLC)")
add_paragraph_with_bookmarks([("ND02_ReferenciaPlanejamento", "")])

add_heading("3. REQUISITOS DA CONTRATAÇÃO (ART. 108, §1º, III, DO RLC)")
add_paragraph_with_bookmarks([("ND03_Requisitos", "")])

add_heading("4. ESTIMATIVA DAS QUANTIDADES, ACOMPANHADAS DAS MEMÓRIAS DE CÁLCULO E DOS DOCUMENTOS QUE LHE DÃO SUPORTE (ART. 108, §1º, IV, E ART. 115, V, DO RLC)")
add_paragraph_with_bookmarks([("ND04_EstimativaQuantidades", "")])

add_heading("5. LEVANTAMENTO DE MERCADO E JUSTIFICATIVA DE ESCOLHA DO TIPO DE SOLUÇÃO A CONTRATAR (ART. 108, §1º, V, E ART. 115, II, DO RLC)")
add_paragraph_with_bookmarks([("ND05_LevantamentoMercado", "")])

add_heading("6. ESTIMATIVAS DE PREÇOS OU PREÇOS REFERENCIAIS (ART. 108, §1º, VI, DO RLC)")
add_paragraph_with_bookmarks([("ND06_EstimativaPrecos", "")])

add_heading("7. DESCRIÇÃO DA SOLUÇÃO COMO UM TODO (ART. 108, §1º, VII, E ART. 115, IV, DO RLC)")
add_paragraph_with_bookmarks([("ND07_DescricaoSolucao", "")])

add_heading("8. JUSTIFICATIVAS PARA O PARCELAMENTO OU NÃO DA SOLUÇÃO, QUANDO NECESSÁRIA PARA INDIVIDUALIZAÇÃO DO OBJETO (ART. 108, §1º, VIII, DO RLC)")
add_bloco_variante("ND08_Parcelamento", {
    "Sim": ["8.1. A solução será parcelada, sendo a licitação realizada por item."],
    "Não": ["8.1. A solução não será parcelada, sendo a licitação realizada em lote único."],
})
add_paragraph_with_bookmarks([("ND08_Justificativa", "")])

add_heading("9. DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS (ART. 108, §1º, IX, E ART. 115, VI, DO RLC)")
add_paragraph_with_bookmarks([("ND09_ResultadosPretendidos", "")])

add_heading("10. PROVIDÊNCIAS PARA ADEQUAÇÃO DO AMBIENTE DA CONAB (ART. 108, §1º, X, DO RLC)")
add_bloco_variante("ND10_AdequacaoAmbiente", {
    "Sim": ["10.1. Será necessária a adequação do ambiente da Conab, conforme a seguir:"],
    "Não": ["10.1. Não há necessidade de adequação do ambiente da Conab para esta contratação."],
})
add_paragraph_with_bookmarks([("ND10_DetalheAdequacao", "")])

add_heading("11. CONTRATAÇÕES CORRELATAS OU INTERDEPENDENTES (ART. 108, §1º, XI, DO RLC)")
add_bloco_variante("ND11_ContratacoesCorrelatas", {
    "Sim": ["11.1. São contratos correlatos ou interdependentes a esta contratação:"],
    "Não": ["11.1. Para este objeto, não há contratos correlatos ou interdependentes."],
})
add_paragraph_with_bookmarks([("ND11_DetalheCorrelatas", "")])

add_heading("12. DECLARAÇÃO DA VIABILIDADE OU NÃO DA CONTRATAÇÃO (ART. 108, §1º, XII, DO RLC)")
add_bloco_variante("ND12_Viabilidade", {
    "Viável": ["12.1. A contratação mostra-se viável."],
    "Inviável": ["12.1. A contratação mostra-se inviável."],
})
add_paragraph_with_bookmarks([("ND12_Justificativa", "")])

add_heading("13. IDENTIFICAÇÃO DA MODALIDADE DE LICITAÇÃO OU DE CONTRATAÇÃO DIRETA (ART. 115, III, DO RLC)")
add_paragraph_with_bookmarks([
    "13.1. Considerando-se tudo o quanto foi exposto, a contratação será realizada na forma de ",
    ("ND13_Modalidade", ""), ".",
])
add_paragraph_with_bookmarks([("ND13_Justificativa", "")])

add_heading("Responsável(is) pelo estudo — Área Demandante")
add_paragraph_with_bookmarks([("ND14_ResponsavelNome", "")])
add_paragraph_with_bookmarks([("ND14_ResponsavelCargo", "")])

doc.save("Template_Nota_Demanda.docx")
print("Template_Nota_Demanda.docx gerado com sucesso.")
```

- [ ] **Step 2: Run the script from `templates/`**

Run: `cd templates && python build_template_nota_demanda.py`
Expected: `Template_Nota_Demanda.docx gerado com sucesso.`

- [ ] **Step 3: Commit**

```bash
git add templates/build_template_nota_demanda.py templates/Template_Nota_Demanda.docx
git commit -m "Add Nota de Demanda docx template build script and generated template"
```

---

### Task 3: `js/dados-nota-tecnica.js` — Nota Técnica data model

**Files:**
- Create: `js/dados-nota-tecnica.js`

**Interfaces:**
- Produces: `export const SECOES_NOTA_TECNICA` — same shape, field ids prefixed `NT`. Seções 1-13 mirror Task 1's structure but with the "tecnicamente" question wording from the spec table. JS section "14" is a UI-only gate ("obra/serviço de engenharia?") with no heading of its own in the generated document — matching the real example, which jumps straight from Seção 13 to Seção 14 "Adequação Técnica" with no separate "natureza do objeto" heading. JS sections 15-18 hold the 4 conditional fields but render as **document headings "14." through "17."** (the doc-visible numbering the Python script in Task 4 actually uses — JS section id and doc heading number are offset by one here, intentionally, since the gate itself isn't a document section). The closing "Responsável(is) pelo estudo" block (JS section "19") is unnumbered in the document, same as the real example's signature block.

- [ ] **Step 1: Write the file**

```js
export const SECOES_NOTA_TECNICA = [
  { id: "01", titulo: "NECESSIDADE DE CONTRATAÇÃO", campos: [
    { id: "NT01_Necessidade", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva tecnicamente a situação-problema (condições físicas, funcionais ou normativas constatadas, de preferência com registro de vistoria) que origina a necessidade desta contratação, e seus impactos para a Conab.",
      ajuda: "Art. 108, §1º, I, e art. 115, I, c/c art. 117, do RLC: a Nota Técnica deve apresentar tecnicamente a situação-problema, com base em vistoria ou levantamento técnico, não apenas no relato funcional já constante da Nota de Demanda." },
  ]},
  { id: "02", titulo: "REFERÊNCIA A OUTROS INSTRUMENTOS DE PLANEJAMENTO DA COMPANHIA", campos: [
    { id: "NT02_ReferenciaPlanejamento", tipo: "especificacao", input: "textarea",
      pergunta: "Do ponto de vista técnico, a contratação está alinhada a algum instrumento de planejamento da Conab ou política pública? Indique eventuais diretrizes técnicas internas aplicáveis.",
      ajuda: "Art. 108, §1º, II, do RLC: indique se a contratação está alinhada aos planos instituídos, e a política pública a que esteja vinculada, quando couber." },
  ]},
  { id: "03", titulo: "REQUISITOS DA CONTRATAÇÃO", campos: [
    { id: "NT03_Requisitos", tipo: "especificacao", input: "textarea",
      pergunta: "Quais os requisitos técnicos necessários ao atendimento da necessidade (especificações técnicas, critérios de sustentabilidade aplicáveis e vigência tecnicamente adequada ao objeto)?",
      ajuda: "Art. 108, §1º, III, do RLC: elenque os requisitos técnicos necessários, critérios de sustentabilidade quando couber, e a duração pretendida do Contrato (limite de 5 anos para prestação de serviços), justificando tecnicamente a escolha." },
  ]},
  { id: "04", titulo: "ESTIMATIVA DAS QUANTIDADES, ACOMPANHADAS DAS MEMÓRIAS DE CÁLCULO E DOS DOCUMENTOS QUE LHE DÃO SUPORTE", campos: [
    { id: "NT04_EstimativaQuantidades", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva tecnicamente o método e a memória de cálculo usados para estimar as quantidades do objeto, com a data prevista de execução.",
      ajuda: "Art. 108, §1º, IV, e art. 115, V, do RLC: defina e documente o método técnico para a estimativa das quantidades, incluindo a memória de cálculo e a data prevista de execução do objeto." },
  ]},
  { id: "05", titulo: "LEVANTAMENTO DE MERCADO E JUSTIFICATIVA DE ESCOLHA DO TIPO DE SOLUÇÃO A CONTRATAR", campos: [
    { id: "NT05_LevantamentoMercado", tipo: "especificacao", input: "textarea",
      pergunta: "Tecnicamente, qual solução foi identificada como a que melhor atende à contratação, com base em quais referências técnicas, normas ou contratações similares?",
      ajuda: "Art. 108, §1º, V, e art. 115, II, do RLC: considere diferentes fontes técnicas (contratações similares, normas técnicas, pesquisas e estudos do setor) para justificar tecnicamente a escolha da solução." },
  ]},
  { id: "06", titulo: "ESTIMATIVAS DE PREÇOS OU PREÇOS REFERENCIAIS", campos: [
    { id: "NT06_EstimativaPrecos", tipo: "especificacao", input: "textarea",
      pergunta: "Qual metodologia técnica foi usada para estimar o preço de referência (fontes, memória de cálculo)?",
      ajuda: "Art. 108, §1º, VI, do RLC: utilize preferencialmente preços referenciais divulgados por órgãos públicos, justificando a escolha do método, e inclua nos autos a memória de cálculo." },
  ]},
  { id: "07", titulo: "DESCRIÇÃO DA SOLUÇÃO COMO UM TODO", campos: [
    { id: "NT07_DescricaoSolucao", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva tecnicamente a solução como um todo, incluindo as normas técnicas aplicáveis (ABNT e correlatas) e os elementos que devem ser produzidos, contratados ou executados.",
      ajuda: "Art. 108, §1º, VII, e art. 115, IV, do RLC: apresente todos os elementos técnicos que devem ser produzidos, contratados ou executados, com definição precisa do objeto e as normas técnicas aplicáveis." },
  ]},
  { id: "08", titulo: "JUSTIFICATIVAS PARA O PARCELAMENTO OU NÃO DA SOLUÇÃO", campos: [
    { id: "NT08_Parcelamento", tipo: "parametro", input: "select",
      pergunta: "A solução será parcelada (licitada por item)?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Art. 108, §1º, VIII, do RLC: o parcelamento é a regra quando o objeto é divisível, desde que não haja prejuízo ao conjunto da solução ou perda de economia de escala." },
    { id: "NT08_Justificativa", tipo: "especificacao", input: "textarea",
      pergunta: "Justifique tecnicamente a decisão de parcelar ou não a solução (viabilidade técnica do parcelamento, risco de perda de escala ou de compatibilização entre partes)" },
  ]},
  { id: "09", titulo: "DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS", campos: [
    { id: "NT09_ResultadosPretendidos", tipo: "especificacao", input: "textarea",
      pergunta: "Quais os resultados técnicos pretendidos com a contratação (ganhos de desempenho, conformidade normativa, durabilidade), e a quem se destinam?",
      ajuda: "Art. 108, §1º, IX, e art. 115, VI, do RLC: declare os benefícios técnicos diretos e indiretos em termos de economicidade, eficácia, eficiência e melhor aproveitamento dos recursos disponíveis." },
  ]},
  { id: "10", titulo: "PROVIDÊNCIAS PARA ADEQUAÇÃO DO AMBIENTE DA CONAB", campos: [
    { id: "NT10_AdequacaoAmbiente", tipo: "parametro", input: "select",
      pergunta: "Há necessidade técnica de adequação do ambiente da Conab (acesso, infraestrutura, capacitação) para a execução da contratação?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Art. 108, §1º, X, do RLC: elabore cronograma das atividades técnicas necessárias e considere a capacitação de empregados, quando aplicável." },
    { id: "NT10_DetalheAdequacao", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva tecnicamente as adequações necessárias",
      mostrarSe: { NT10_AdequacaoAmbiente: "Sim" } },
  ]},
  { id: "11", titulo: "CONTRATAÇÕES CORRELATAS OU INTERDEPENDENTES", campos: [
    { id: "NT11_ContratacoesCorrelatas", tipo: "parametro", input: "select",
      pergunta: "Há contratos ou projetos técnicos correlatos ou interdependentes a esta contratação?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não" },
    { id: "NT11_DetalheCorrelatas", tipo: "especificacao", input: "textarea",
      pergunta: "Liste os contratos ou projetos técnicos correlatos ou interdependentes",
      mostrarSe: { NT11_ContratacoesCorrelatas: "Sim" } },
  ]},
  { id: "12", titulo: "DECLARAÇÃO DA VIABILIDADE OU NÃO DA CONTRATAÇÃO", campos: [
    { id: "NT12_Viabilidade", tipo: "parametro", input: "select",
      pergunta: "A contratação é tecnicamente viável?",
      opcoes: ["Viável", "Inviável"], blocoTexto: true, padrao: "Viável",
      ajuda: "Art. 108, §1º, XII, do RLC: a declaração deve ser justificada com base nos elementos técnicos averiguados nos itens anteriores." },
    { id: "NT12_Justificativa", tipo: "especificacao", input: "textarea",
      pergunta: "Justifique tecnicamente a declaração de viabilidade" },
  ]},
  { id: "13", titulo: "IDENTIFICAÇÃO DA MODALIDADE DE LICITAÇÃO OU DE CONTRATAÇÃO DIRETA", campos: [
    { id: "NT13_Modalidade", tipo: "parametro", input: "select",
      pergunta: "Qual modalidade de licitação (ou contratação direta) será utilizada?",
      opcoes: ["Pregão Eletrônico", "Competição Pública", "RDC", "Dispensa de Licitação"],
      ajuda: "Art. 115, III, do RLC: identifique a modalidade e justifique sua adoção." },
    { id: "NT13_Justificativa", tipo: "especificacao", input: "textarea",
      pergunta: "Justifique a escolha da modalidade" },
  ]},
  { id: "14", titulo: "DA NATUREZA DO OBJETO PARA FINS DO ART. 123 DO RLC", campos: [
    { id: "NT_ObraServicoEngenharia", tipo: "parametro", input: "select",
      pergunta: "O objeto desta contratação é obra ou serviço de engenharia?",
      opcoes: ["Sim", "Não"], blocoTexto: true, padrao: "Não",
      ajuda: "Art. 121 a 124 do RLC (Subseção II — Da Elaboração da Nota Técnica para Obras e Serviços de Engenharia): quando o objeto é obra ou serviço de engenharia, a Nota Técnica deve conter, adicionalmente, as seções de Adequação Técnica, Funcionalidade, Requisitos Ambientais e Adequação às Normas Vigentes (art. 123, V a VIII).",
      blocosAdicionais: ["NT_SecoesEngenharia"] },
  ]},
  { id: "15", titulo: "ADEQUAÇÃO TÉCNICA", campos: [
    { id: "NT15_AdequacaoTecnica", tipo: "especificacao", input: "textarea",
      pergunta: "Demonstre a adequação técnica da solução escolhida",
      ajuda: "Art. 123, V, do RLC.",
      mostrarSe: { NT_ObraServicoEngenharia: "Sim" } },
  ]},
  { id: "16", titulo: "FUNCIONALIDADE", campos: [
    { id: "NT16_Funcionalidade", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva a funcionalidade da contratação pretendida",
      ajuda: "Art. 123, VI, do RLC.",
      mostrarSe: { NT_ObraServicoEngenharia: "Sim" } },
  ]},
  { id: "17", titulo: "REQUISITOS AMBIENTAIS", campos: [
    { id: "NT17_RequisitosAmbientais", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva os requisitos ambientais aplicáveis à execução do objeto",
      ajuda: "Art. 123, VII, do RLC.",
      mostrarSe: { NT_ObraServicoEngenharia: "Sim" } },
  ]},
  { id: "18", titulo: "ADEQUAÇÃO ÀS NORMAS VIGENTES", campos: [
    { id: "NT18_AdequacaoNormasVigentes", tipo: "especificacao", input: "textarea",
      pergunta: "Descreva a adequação da solução às normas técnicas vigentes aplicáveis",
      ajuda: "Art. 123, VIII, do RLC.",
      mostrarSe: { NT_ObraServicoEngenharia: "Sim" } },
  ]},
  { id: "19", titulo: "DISPOSIÇÕES FINAIS", campos: [
    { id: "NT19_ResponsavelNome", tipo: "parametro", input: "texto",
      pergunta: "Responsável pelo estudo — nome" },
    { id: "NT19_ResponsavelCargo", tipo: "parametro", input: "texto",
      pergunta: "Responsável pelo estudo — cargo/função (Área Técnica)" },
  ]},
];
```

- [ ] **Step 2: Commit**

```bash
git add js/dados-nota-tecnica.js
git commit -m "Add Nota Técnica data model (13 base sections + art. 123 engineering block)"
```

---

### Task 4: `templates/build_template_nota_tecnica.py` — Nota Técnica template

**Files:**
- Create: `templates/build_template_nota_tecnica.py`
- Generate: `templates/Template_Nota_Tecnica.docx`

**Interfaces:**
- Consumes: field ids from Task 3.
- Produces: `templates/Template_Nota_Tecnica.docx`.

- [ ] **Step 1: Write the build script**

```python
"""Gera Template_Nota_Tecnica.docx com bookmarks nos pontos de
preenchimento identificados em js/dados-nota-tecnica.js. Mesmo esqueleto
de 13 seções da Nota de Demanda (art. 108/115 do RLC), com redação
tecnicamente orientada (art. 117), mais um bloco condicional de 4
seções extras (art. 123, V-VIII) para obras e serviços de engenharia.
Ver docs/superpowers/specs/2026-08-18-nota-demanda-nota-tecnica-design.md."""
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


def add_bloco_variante_sentinela_inicio(nome_bookmark):
    """Abre um bookmark que vai envolver vários parágrafos/seções
    inteiras (não cabe num único parágrafo). Combinar com
    add_bloco_variante_sentinela_fim(bid)."""
    _bookmark_id[0] += 1
    bid = str(_bookmark_id[0])
    p = doc.add_paragraph()
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


add_heading("COMPANHIA NACIONAL DE ABASTECIMENTO")
add_heading("NOTA TÉCNICA")

add_heading("1. NECESSIDADE DE CONTRATAÇÃO (ART. 108, §1º, I, E ART. 115, I, DO RLC)")
add_paragraph_with_bookmarks([("NT01_Necessidade", "")])

add_heading("2. REFERÊNCIA A OUTROS INSTRUMENTOS DE PLANEJAMENTO DA COMPANHIA (ART. 108, §1º, II, DO RLC)")
add_paragraph_with_bookmarks([("NT02_ReferenciaPlanejamento", "")])

add_heading("3. REQUISITOS DA CONTRATAÇÃO (ART. 108, §1º, III, DO RLC)")
add_paragraph_with_bookmarks([("NT03_Requisitos", "")])

add_heading("4. ESTIMATIVA DAS QUANTIDADES, ACOMPANHADAS DAS MEMÓRIAS DE CÁLCULO E DOS DOCUMENTOS QUE LHE DÃO SUPORTE (ART. 108, §1º, IV, E ART. 115, V, DO RLC)")
add_paragraph_with_bookmarks([("NT04_EstimativaQuantidades", "")])

add_heading("5. LEVANTAMENTO DE MERCADO E JUSTIFICATIVA DE ESCOLHA DO TIPO DE SOLUÇÃO A CONTRATAR (ART. 108, §1º, V, E ART. 115, II, DO RLC)")
add_paragraph_with_bookmarks([("NT05_LevantamentoMercado", "")])

add_heading("6. ESTIMATIVAS DE PREÇOS OU PREÇOS REFERENCIAIS (ART. 108, §1º, VI, DO RLC)")
add_paragraph_with_bookmarks([("NT06_EstimativaPrecos", "")])

add_heading("7. DESCRIÇÃO DA SOLUÇÃO COMO UM TODO (ART. 108, §1º, VII, E ART. 115, IV, DO RLC)")
add_paragraph_with_bookmarks([("NT07_DescricaoSolucao", "")])

add_heading("8. JUSTIFICATIVAS PARA O PARCELAMENTO OU NÃO DA SOLUÇÃO, QUANDO NECESSÁRIA PARA INDIVIDUALIZAÇÃO DO OBJETO (ART. 108, §1º, VIII, DO RLC)")
add_bloco_variante("NT08_Parcelamento", {
    "Sim": ["8.1. A solução será parcelada, sendo a licitação realizada por item."],
    "Não": ["8.1. A solução não será parcelada, sendo a licitação realizada em lote único."],
})
add_paragraph_with_bookmarks([("NT08_Justificativa", "")])

add_heading("9. DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS (ART. 108, §1º, IX, E ART. 115, VI, DO RLC)")
add_paragraph_with_bookmarks([("NT09_ResultadosPretendidos", "")])

add_heading("10. PROVIDÊNCIAS PARA ADEQUAÇÃO DO AMBIENTE DA CONAB (ART. 108, §1º, X, DO RLC)")
add_bloco_variante("NT10_AdequacaoAmbiente", {
    "Sim": ["10.1. Será necessária a adequação do ambiente da Conab, conforme a seguir:"],
    "Não": ["10.1. Não há necessidade de adequação do ambiente da Conab para esta contratação."],
})
add_paragraph_with_bookmarks([("NT10_DetalheAdequacao", "")])

add_heading("11. CONTRATAÇÕES CORRELATAS OU INTERDEPENDENTES (ART. 108, §1º, XI, DO RLC)")
add_bloco_variante("NT11_ContratacoesCorrelatas", {
    "Sim": ["11.1. São contratos ou projetos técnicos correlatos ou interdependentes a esta contratação:"],
    "Não": ["11.1. Para este objeto, não há contratos ou projetos técnicos correlatos ou interdependentes."],
})
add_paragraph_with_bookmarks([("NT11_DetalheCorrelatas", "")])

add_heading("12. DECLARAÇÃO DA VIABILIDADE OU NÃO DA CONTRATAÇÃO (ART. 108, §1º, XII, DO RLC)")
add_bloco_variante("NT12_Viabilidade", {
    "Viável": ["12.1. A contratação mostra-se tecnicamente viável."],
    "Inviável": ["12.1. A contratação mostra-se tecnicamente inviável."],
})
add_paragraph_with_bookmarks([("NT12_Justificativa", "")])

add_heading("13. IDENTIFICAÇÃO DA MODALIDADE DE LICITAÇÃO OU DE CONTRATAÇÃO DIRETA (ART. 115, III, DO RLC)")
add_paragraph_with_bookmarks([
    "13.1. Considerando-se tudo o quanto foi exposto, a contratação será realizada na forma de ",
    ("NT13_Modalidade", ""), ".",
])
add_paragraph_with_bookmarks([("NT13_Justificativa", "")])

# Bloco condicional (art. 123, V-VIII do RLC) — só para obras e
# serviços de engenharia. Abrange 4 seções inteiras, por isso usa as
# sentinelas em vez de add_bloco_variante (que só envolve parágrafos
# soltos, não títulos com heading próprio).
_bid_engenharia = add_bloco_variante_sentinela_inicio("NT_SecoesEngenharia_BLOCO_SIM")

add_heading("14. ADEQUAÇÃO TÉCNICA (ART. 123, V, DO RLC)")
add_paragraph_with_bookmarks([("NT15_AdequacaoTecnica", "")])

add_heading("15. FUNCIONALIDADE (ART. 123, VI, DO RLC)")
add_paragraph_with_bookmarks([("NT16_Funcionalidade", "")])

add_heading("16. REQUISITOS AMBIENTAIS (ART. 123, VII, DO RLC)")
add_paragraph_with_bookmarks([("NT17_RequisitosAmbientais", "")])

add_heading("17. ADEQUAÇÃO ÀS NORMAS VIGENTES (ART. 123, VIII, DO RLC)")
add_paragraph_with_bookmarks([("NT18_AdequacaoNormasVigentes", "")])

add_bloco_variante_sentinela_fim(_bid_engenharia)

add_heading("Responsável(is) pelo estudo — Área Técnica")
add_paragraph_with_bookmarks([("NT19_ResponsavelNome", "")])
add_paragraph_with_bookmarks([("NT19_ResponsavelCargo", "")])

doc.save("Template_Nota_Tecnica.docx")
print("Template_Nota_Tecnica.docx gerado com sucesso.")
```

- [ ] **Step 2: Run the script from `templates/`**

Run: `cd templates && python build_template_nota_tecnica.py`
Expected: `Template_Nota_Tecnica.docx gerado com sucesso.`

- [ ] **Step 3: Commit**

```bash
git add templates/build_template_nota_tecnica.py templates/Template_Nota_Tecnica.docx
git commit -m "Add Nota Técnica docx template build script and generated template"
```

---

### Task 5: Register both instruments in the app shell

**Files:**
- Modify: `js/app.js`
- Modify: `js/triagem.js`

**Interfaces:**
- Consumes: `SECOES_NOTA_DEMANDA` from Task 1, `SECOES_NOTA_TECNICA` from Task 3.

- [ ] **Step 1: Add imports and registry entries to `js/app.js`**

Add near the top, with the other imports:

```js
import { SECOES_NOTA_DEMANDA } from "./dados-nota-demanda.js";
import { SECOES_NOTA_TECNICA } from "./dados-nota-tecnica.js";
```

Add to the `INSTRUMENTOS` object, after the `"PB"` entry:

```js
  "Nota-Demanda": {
    secoes: SECOES_NOTA_DEMANDA,
    titulo: "Nota de Demanda",
    templateUrl: "templates/Template_Nota_Demanda.docx",
    chaveAutosave: "gerador-tr-pb:nota-demanda:rascunho",
    nomeArquivoRascunho: "rascunho-nota-demanda.json",
    nomeArquivoGerado: "Nota-de-Demanda-gerada.docx",
  },
  "Nota-Tecnica": {
    secoes: SECOES_NOTA_TECNICA,
    titulo: "Nota Técnica",
    templateUrl: "templates/Template_Nota_Tecnica.docx",
    chaveAutosave: "gerador-tr-pb:nota-tecnica:rascunho",
    nomeArquivoRascunho: "rascunho-nota-tecnica.json",
    nomeArquivoGerado: "Nota-Tecnica-gerada.docx",
  },
```

- [ ] **Step 2: Add two cards to `js/triagem.js`**

Read the current file first — it has an `ICONE_*` constant per card (SVG strings) and a `renderizarTriagem(container, aoConfirmar)` function building `.card-instrumento` buttons inside `.grid-instrumentos`. Add two new icon constants and two new cards, following the exact same pattern as the 4 existing ones:

```js
const ICONE_NOTA_DEMANDA = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 3h9l5 5v13a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M15 3v5h5"/><path d="M8 13h8M8 17h5"/></svg>`;
const ICONE_NOTA_TECNICA = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 3h9l5 5v13a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M15 3v5h5"/><path d="M9 13l1.5 1.5L14 11"/></svg>`;
```

Add two `<button>` cards inside `.grid-instrumentos`, after the PB card:

```html
      <button id="btn-nota-demanda" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_NOTA_DEMANDA}</span>
        <span class="card-instrumento-titulo">Estudo Preliminar</span>
        <span class="card-instrumento-subtitulo">Nota de Demanda</span>
        <span class="card-instrumento-desc">Art. 108 e 115 do RLC</span>
      </button>
      <button id="btn-nota-tecnica" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_NOTA_TECNICA}</span>
        <span class="card-instrumento-titulo">Estudo Preliminar</span>
        <span class="card-instrumento-subtitulo">Nota Técnica</span>
        <span class="card-instrumento-desc">Art. 108, 115 e 123 do RLC</span>
      </button>
```

And wire the two new listeners next to the existing 4:

```js
  container.querySelector("#btn-nota-demanda")
    .addEventListener("click", () => aoConfirmar("Nota-Demanda"));
  container.querySelector("#btn-nota-tecnica")
    .addEventListener("click", () => aoConfirmar("Nota-Tecnica"));
```

- [ ] **Step 3: Manual verification in the browser**

Start the static server, open the triagem screen, confirm 6 cards now render (4 existing + 2 new), click into each new one, confirm the painel renders all sections, confirm the live preview panel shows the template text, type into `ND08_Parcelamento`/`NT_ObraServicoEngenharia` and confirm the conditional blocks add/remove correctly in the preview.

- [ ] **Step 4: Run the full test suite**

Run: `node --test tests/*.test.mjs`
Expected: PASS (no regressions — this task only adds new registry entries and DOM wiring)

- [ ] **Step 5: Commit**

```bash
git add js/app.js js/triagem.js
git commit -m "Register Nota de Demanda and Nota Técnica as instruments"
```

---

### Task 6: End-to-end generation tests for both instruments

**Files:**
- Create: `tests/nota-demanda-nota-tecnica.test.mjs`

**Interfaces:**
- Consumes: `gerarDocx` from `js/gerador-docx.js`, `renderizarPreviewHtml` from `js/preview-docx.js`, `SECOES_NOTA_DEMANDA`/`SECOES_NOTA_TECNICA`.

- [ ] **Step 1: Write the failing test**

```js
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { renderizarPreviewHtml } from "../js/preview-docx.js";
import { SECOES_NOTA_DEMANDA } from "../js/dados-nota-demanda.js";
import { SECOES_NOTA_TECNICA } from "../js/dados-nota-tecnica.js";

globalThis.JSZip = JSZip;

function lerTemplate(caminho) {
  const buffer = fs.readFileSync(caminho);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlNotaDemanda(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate("templates/Template_Nota_Demanda.docx"), respostas, SECOES_NOTA_DEMANDA);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

async function gerarXmlNotaTecnica(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate("templates/Template_Nota_Tecnica.docx"), respostas, SECOES_NOTA_TECNICA);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function textoDoDocumento(xml) {
  return xml.replace(/<[^>]+>/g, "");
}

test("Nota de Demanda: geração básica preenche a necessidade", async () => {
  const xml = await gerarXmlNotaDemanda({ ND01_Necessidade: "Reforma da unidade de Rondônia" });
  const texto = textoDoDocumento(xml);
  assert.ok(texto.includes("Reforma da unidade de Rondônia"));
  assert.ok(texto.includes("NOTA DE DEMANDA"));
});

test("Nota de Demanda: parcelamento Sim/Não troca a redação da Seção 8", async () => {
  const textoSim = textoDoDocumento(await gerarXmlNotaDemanda({ ND08_Parcelamento: "Sim" }));
  assert.ok(textoSim.includes("A solução será parcelada, sendo a licitação realizada por item."));
  const textoNao = textoDoDocumento(await gerarXmlNotaDemanda({ ND08_Parcelamento: "Não" }));
  assert.ok(textoNao.includes("A solução não será parcelada, sendo a licitação realizada em lote único."));
});

test("Nota de Demanda: adequação do ambiente Sim mostra detalhe; Não remove", async () => {
  const textoSim = textoDoDocumento(await gerarXmlNotaDemanda({
    ND10_AdequacaoAmbiente: "Sim",
    ND10_DetalheAdequacao: "Necessário treinar a equipe de fiscalização",
  }));
  assert.ok(textoSim.includes("Necessário treinar a equipe de fiscalização"));
  const textoNao = textoDoDocumento(await gerarXmlNotaDemanda({
    ND10_AdequacaoAmbiente: "Não",
    ND10_DetalheAdequacao: "Não deveria aparecer",
  }));
  assert.ok(!textoNao.includes("Não deveria aparecer"));
  assert.ok(textoNao.includes("Não há necessidade de adequação do ambiente da Conab"));
});

test("Nota Técnica: geração básica usa redação tecnicamente orientada", async () => {
  const xml = await gerarXmlNotaTecnica({ NT01_Necessidade: "Reforma técnica da SUREG/RO" });
  const texto = textoDoDocumento(xml);
  assert.ok(texto.includes("Reforma técnica da SUREG/RO"));
  assert.ok(texto.includes("NOTA TÉCNICA"));
  assert.ok(texto.includes("Descreva tecnicamente a situação-problema"));
});

test("Nota Técnica: obra/serviço de engenharia Sim mantém as 4 seções extras; Não remove", async () => {
  const textoSim = textoDoDocumento(await gerarXmlNotaTecnica({
    NT_ObraServicoEngenharia: "Sim",
    NT15_AdequacaoTecnica: "Solução compatível com as normas ABNT aplicáveis",
    NT16_Funcionalidade: "Atende à necessidade de reforma predial",
    NT17_RequisitosAmbientais: "Reaproveitamento de materiais sempre que viável",
    NT18_AdequacaoNormasVigentes: "Conforme NBR 5410 e NBR 5419",
  }));
  assert.ok(textoSim.includes("ADEQUAÇÃO TÉCNICA (ART. 123, V, DO RLC)"));
  assert.ok(textoSim.includes("Solução compatível com as normas ABNT aplicáveis"));
  assert.ok(textoSim.includes("Atende à necessidade de reforma predial"));
  assert.ok(textoSim.includes("Reaproveitamento de materiais sempre que viável"));
  assert.ok(textoSim.includes("Conforme NBR 5410 e NBR 5419"));

  const textoNao = textoDoDocumento(await gerarXmlNotaTecnica({ NT_ObraServicoEngenharia: "Não" }));
  assert.ok(!textoNao.includes("ADEQUAÇÃO TÉCNICA (ART. 123, V, DO RLC)"));
  assert.ok(!textoNao.includes("FUNCIONALIDADE (ART. 123, VI, DO RLC)"));
  assert.ok(!textoNao.includes("REQUISITOS AMBIENTAIS (ART. 123, VII, DO RLC)"));
  assert.ok(!textoNao.includes("ADEQUAÇÃO ÀS NORMAS VIGENTES (ART. 123, VIII, DO RLC)"));
});

test("Nota Técnica: preview ao vivo marca o texto preenchido, sem tags cruas", async () => {
  const docxBuffer = await gerarDocx(
    lerTemplate("templates/Template_Nota_Tecnica.docx"),
    { NT01_Necessidade: "Reforma elétrica da SUREG/RO" },
    SECOES_NOTA_TECNICA
  );
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const html = renderizarPreviewHtml(documentXml);
  assert.match(html, /<mark class="preview-destaque"[^>]*>[^<]*Reforma elétrica da SUREG\/RO[^<]*<\/mark>/);
  assert.ok(!html.includes("w:bookmarkStart"));
  assert.ok(!html.includes("_BLOCO_"));
});
```

- [ ] **Step 2: Run the tests**

Run: `node --test tests/nota-demanda-nota-tecnica.test.mjs`
Expected: PASS (7 tests). If any fail, fix `js/dados-nota-demanda.js`, `js/dados-nota-tecnica.js`, or the two `templates/build_template_*.py` scripts (whichever bookmark name mismatches) before proceeding — do not adjust the test to match broken output.

- [ ] **Step 3: Run the full suite to check for regressions**

Run: `node --test tests/*.test.mjs`
Expected: PASS (all existing + new tests, 110 total)

- [ ] **Step 4: Commit**

```bash
git add tests/nota-demanda-nota-tecnica.test.mjs
git commit -m "Add end-to-end generation tests for Nota de Demanda and Nota Técnica"
```

---

## Execution Handoff

After all 6 tasks are complete, run `node --test tests/*.test.mjs` once more to
confirm the whole suite is green, then follow
`superpowers:finishing-a-development-branch` to merge (this plan follows the
project's worktree-per-feature convention — create `.worktrees/nota-demanda-nota-tecnica`
off `master` before starting Task 1, per `superpowers:using-git-worktrees`).
