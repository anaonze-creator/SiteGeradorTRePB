# Melhorias diversas (amostra, similaridade, prorrogação, matriz de riscos etc.) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Resolve o lote de 12 pontos levantados pela usuária ao revisar o site: um bug de pontuação, dois campos sem ajuda/exemplo, um campo de texto livre faltando, uma cláusula de prorrogação ausente, uma automação de aviso por valor, valores sugeridos de multa/pagamento por analogia, e a reconstrução estrutural da tabela da Matriz de Riscos.

**Architecture:** Mesmo padrão de field-schema + bookmark do projeto inteiro. Nenhum mecanismo novo do motor — usa `ajuda`, `opcoes` com item extra, `mostrarSe`, e uma função Python nova para o cabeçalho de dois níveis da Matriz de Riscos (`python-docx` permite mesclar células de cabeçalho via `merge`).

**Tech Stack:** Vanilla JS (schemas + `js/render.js` para o aviso de Programa de Integridade), Python + `python-docx` (templates), Node test runner.

## Global Constraints

- Nunca inventar conteúdo. Toda citação abaixo foi verificada no PDF do RLC ou nos `.odt` oficiais durante a investigação.
- Regenerar o `.docx` correspondente após cada edição de template Python, sempre rodando `python build_template_X.py` de dentro de `templates/`.
- Rodar `node --test tests/*.test.mjs` após cada task antes de seguir para a próxima.

---

### Task 1: Corrigir pontuação faltando na Seção 13 do PB (Habilitação)

**Files:**
- Modify: `templates/build_template_pb.py` (parágrafo `13.1.2.`, perto da linha 354-358)

**Interfaces:** nenhuma nova — só ajuste de texto fixo.

- [ ] **Step 1: Adicionar o ponto final faltando**

Localizar:
```python
add_paragraph_with_bookmarks([
    "13.1.2. Certidão de Acervo Técnico (CAT) em nome do(s) responsável(is) técnico(s), "
    "comprovando responsabilidade por serviços de características semelhantes às parcelas "
    "de maior relevância técnica e econômica do objeto: ", ("PB13_ParcelasRelevantes", ""),
])
```

Trocar por:
```python
add_paragraph_with_bookmarks([
    "13.1.2. Certidão de Acervo Técnico (CAT) em nome do(s) responsável(is) técnico(s), "
    "comprovando responsabilidade por serviços de características semelhantes às parcelas "
    "de maior relevância técnica e econômica do objeto: ", ("PB13_ParcelasRelevantes", ""), ".",
])
```

- [ ] **Step 2: Regenerar e verificar**

```bash
cd templates && python build_template_pb.py && cd ..
python -c "
import zipfile
z = zipfile.ZipFile('templates/Template_PB.docx')
xml = z.read('word/document.xml').decode('utf-8')
print('do objeto: .13.1.3' not in xml.replace(chr(10),''))
"
```

- [ ] **Step 3: Commit**

```bash
git add templates/build_template_pb.py templates/Template_PB.docx
git commit -m "Fix missing period after PB13_ParcelasRelevantes in Seção 13"
```

---

### Task 2: Adicionar ajuda/exemplos em amostra (TR-Bens) e similaridade (PB)

**Files:**
- Modify: `js/dados-tr-bens.js:49-57` (Seção 05, campos `S05_PrazoAmostraDias`, `S05_LocalAmostra`, `S05_HorarioAmostra`)
- Modify: `js/dados-pb.js:136-143` (Seção 11, campos `PB11_ExigeSimilaridade`, `PB11_DetalheSimilaridade`)

**Interfaces:** nenhuma nova — só `ajuda` em campos já existentes.

- [ ] **Step 1: `js/dados-tr-bens.js`**

```javascript
    { id: "S05_PrazoAmostraDias", tipo: "parametro", input: "numero",
      pergunta: "Prazo para apresentação da amostra (dias úteis)",
      mostrarSe: { S05_ExigeAmostra: "Sim" },
      ajuda: "Prazo contado da convocação do licitante provisoriamente classificado em primeiro lugar. Ex.: 5 dias úteis para bens de pronta entrega; até 10 dias úteis para itens que exijam fabricação/montagem." },
    { id: "S05_LocalAmostra", tipo: "parametro", input: "texto",
      pergunta: "Local (gerência/setor e endereço) de entrega da amostra",
      mostrarSe: { S05_ExigeAmostra: "Sim" },
      ajuda: "Ex.: Gerência de Logística — SGAN 601, Módulo J, Brasília/DF." },
    { id: "S05_HorarioAmostra", tipo: "parametro", input: "texto",
      pergunta: "Horário de recebimento da amostra",
      mostrarSe: { S05_ExigeAmostra: "Sim" },
      ajuda: "Ex.: das 9h às 12h e das 14h às 17h, em dias úteis." },
```

- [ ] **Step 2: `js/dados-pb.js`**

```javascript
  { id: "11", titulo: "DOS CRITÉRIOS DE SIMILARIDADE", campos: [
    { id: "PB11_ExigeSimilaridade", tipo: "parametro", input: "select",
      pergunta: "Haverá referências comerciais/marcas para fixação de padrão de qualidade?", opcoes: ["Sim", "Não"],
      blocoTexto: true, padrao: "Não",
      ajuda: "Art. 157, X, do RLC: use referências comerciais apenas para fixar um padrão objetivo de qualidade — nunca para direcionar a licitação a uma marca específica." },
    { id: "PB11_DetalheSimilaridade", tipo: "especificacao", input: "textarea",
      pergunta: "Indique as referências comerciais/marcas (acrescente \"ou equivalente/similar aprovado pela Fiscalização\")",
      mostrarSe: { PB11_ExigeSimilaridade: "Sim" },
      ajuda: "Ex.: piso cerâmico padrão Portinari linha XYZ ou equivalente/similar aprovado pela Fiscalização; grupo motor-gerador padrão Cummins ou equivalente/similar aprovado pela Fiscalização." },
  ]},
```

- [ ] **Step 3: Rodar a suíte**

```bash
node --test tests/*.test.mjs
```

Nenhuma mudança de template Python — `ajuda` é só UI, não entra no `.docx`.

- [ ] **Step 4: Commit**

```bash
git add js/dados-tr-bens.js js/dados-pb.js
git commit -m "Add ajuda/exemplos to amostra (TR-Bens) and similaridade (PB) fields"
```

---

### Task 3: Campo "Outro" em Metodologia da Estimativa (TR-Bens e TR-Serviços)

**Files:**
- Modify: `js/dados-tr-bens.js:8-10` (`S01_MetodologiaEstimativa`)
- Modify: `js/dados-tr-servicos.js:11-13` (`S01_MetodologiaEstimativa`)
- Modify: `templates/build_template_tr_bens.py` (parágrafo `1.2.`, perto da linha 164-170)
- Modify: `templates/build_template_tr_servicos.py` (parágrafo `1.3.`, perto da linha 192-198)

**Interfaces:**
- Produces: campo novo `S01_MetodologiaEstimativaOutro` em cada arquivo, `mostrarSe: { S01_MetodologiaEstimativa: "Outro" }`.

- [ ] **Step 1: Adicionar a opção "Outro" e o campo de texto livre em `js/dados-tr-bens.js`**

```javascript
    { id: "S01_MetodologiaEstimativa", tipo: "parametro", input: "select",
      pergunta: "Metodologia da estimativa de preço",
      opcoes: ["Média", "Mediana", "Menor Preço", "Outro"] },
    { id: "S01_MetodologiaEstimativaOutro", tipo: "parametro", input: "texto",
      pergunta: "Descreva a metodologia de composição do preço de referência utilizada",
      mostrarSe: { S01_MetodologiaEstimativa: "Outro" } },
```

Repita o mesmo em `js/dados-tr-servicos.js` (campo já tem `ajuda`; preserve-a, só adicione "Outro" às `opcoes` e o novo campo condicional logo abaixo).

- [ ] **Step 2: Wire no template — TR-Bens**

Localizar o parágrafo `1.2.` (embute `S01_MetodologiaEstimativa` diretamente). Trocar o trecho:
```python
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
```
por:
```python
    ("S01_MetodologiaEstimativa", "(MÉDIA, MEDIANA OU MENOR PREÇO)"),
    " dos preços praticados no mercado.",
```
(mantém igual) e adicionar, logo depois do `add_paragraph_with_bookmarks` desse parágrafo, um novo bloco condicional simples:
```python
add_bloco_variante("S01_MetodologiaEstimativaOutro_BLOCO", {
    "Outro": [["Metodologia detalhada: ", ("S01_MetodologiaEstimativaOutro", "____"), "."]],
    "Média": [""], "Mediana": [""], "Menor Preço": [""],
})
```
Como TR-Bens não tem `add_bloco_variante` gated por `S01_MetodologiaEstimativa` diretamente (é campo simples, não `blocoTexto`), a forma mais simples e consistente com o resto do arquivo é usar `mostrarSe` no campo (Step 1, já feito) e deixar o parágrafo condicional como um `add_paragraph_with_bookmarks` de campo escondido — que já é removido automaticamente pelo motor quando oculto (mesmo mecanismo do `S05_PrazoAmostraDias`). Ou seja, substitua o bloco `add_bloco_variante` acima por, simplesmente:
```python
add_paragraph_with_bookmarks(["Metodologia detalhada: ", ("S01_MetodologiaEstimativaOutro", "____")])
```
logo após o parágrafo `1.2.`. O motor remove esse parágrafo inteiro quando `S01_MetodologiaEstimativaOutro` estiver oculto (resposta ausente por `mostrarSe` não bater), pelo mesmo mecanismo documentado em `js/gerador-docx.js` (`camposOcultos`).

- [ ] **Step 3: Wire no template — TR-Serviços**

Mesma lógica, após o parágrafo `1.3.` (que embute `S01_MetodologiaEstimativa`):
```python
add_paragraph_with_bookmarks(["Metodologia detalhada: ", ("S01_MetodologiaEstimativaOutro", "____")])
```

- [ ] **Step 4: Regenerar e testar**

```bash
cd templates && python build_template_tr_bens.py && python build_template_tr_servicos.py && cd ..
node --test tests/*.test.mjs
```

- [ ] **Step 5: Commit**

```bash
git add js/dados-tr-bens.js js/dados-tr-servicos.js templates/build_template_tr_bens.py templates/build_template_tr_servicos.py templates/Template_TR_Bens.docx templates/Template_TR_Servicos.docx
git commit -m "Add free-text 'Outro' option to Metodologia da Estimativa (TR-Bens, TR-Serviços)"
```

---

### Task 4: Cláusula de prorrogação com campo de prazo — TR-Bens

**Files:**
- Modify: `js/dados-tr-bens.js:37-40` (Seção 04)
- Modify: `templates/build_template_tr_bens.py:217-224` (Seção 4)

**Contexto:** A minuta-fonte de TR-Bens (SEI 55379667) já cita "prorrogável na forma do art. 497 e 498 do RLC" no texto fixo — mantemos essa citação (é fiel à fonte, mesmo sendo a mesma citação encontrada questionável em outro instrumento). O que falta é um campo para a área demandante registrar por quanto tempo a prorrogação pode se estender, já que hoje só existe o prazo inicial.

**Interfaces:**
- Produces: campo `S04_PrazoProrrogacaoMeses` (opcional, texto livre — a prorrogação de compras não tem um teto legal fixo equivalente ao art. 488, então não force um `input: "numero"` com limite).

- [ ] **Step 1: Adicionar o campo em `js/dados-tr-bens.js`**

```javascript
  { id: "04", titulo: "DA VIGÊNCIA DO CONTRATO", campos: [
    { id: "S04_PrazoVigenciaMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo de vigência do contrato (meses)" },
    { id: "S04_PrazoProrrogacaoMeses", tipo: "parametro", input: "texto",
      pergunta: "Por quanto tempo o contrato poderá ser prorrogado, se necessário (ex.: até mais 12 meses)?",
      ajuda: "Arts. 497 e 498 do RLC: a prorrogação de prazos de execução/vigência depende do motivo (alteração de projeto, fato imprevisível ou atraso da própria Conab — art. 497 — ou culpa da contratada — art. 498) e não tem um teto legal fixo para compras; defina um limite razoável para o objeto." },
  ]},
```

- [ ] **Step 2: Wire no template**

```python
# 4. DA VIGÊNCIA DO CONTRATO
add_heading("4. DA VIGÊNCIA DO CONTRATO")
add_paragraph_with_bookmarks([
    "4.1. O prazo de vigência do contrato será de ",
    ("S04_PrazoVigenciaMeses", "____"),
    " meses (duração do contrato, conforme art. 461 a 463 do RLC), contados da data de sua "
    "assinatura, prorrogável na forma do art. 497 e 498 do RLC.",
])
add_paragraph_with_bookmarks([
    "4.2. Quando cabível, a prorrogação poderá se estender por ",
    ("S04_PrazoProrrogacaoMeses", "____"), ".",
])
```

- [ ] **Step 3: Regenerar e testar**

```bash
cd templates && python build_template_tr_bens.py && cd ..
node --test tests/*.test.mjs
```

- [ ] **Step 4: Commit**

```bash
git add js/dados-tr-bens.js templates/build_template_tr_bens.py templates/Template_TR_Bens.docx
git commit -m "Add prorrogação duration field to TR-Bens vigência clause"
```

---

### Task 5: Aviso automático de Programa de Integridade por valor estimado

**Files:**
- Modify: `js/render.js` (função `criarCampoInput`, ou novo helper específico)
- Modify: `js/dados-tr-bens.js:89-93` (campo `S10_PrazoProgramaIntegridadeMeses`)
- Modify: `js/dados-tr-servicos.js` (campo equivalente)

**Contexto:** Hoje a `ajuda` do campo já explica o limiar (R$ 5.000.000,00, Anexo XXXII), mas a usuária precisa lembrar de checar manualmente contra o valor estimado (Seção 1). Pedido: um aviso automático, calculado a partir do valor já preenchido em `S01_ValorEstimado`.

**Design:** Sem criar um mecanismo novo de "campo computado" no motor (que afetaria `gerador-docx.js` e todos os outros campos) — implementação mais simples e isolada: adicionar um pequeno aviso visual (`<div class="aviso-limiar">`) logo abaixo do campo `S10_PrazoProgramaIntegridadeMeses` no formulário, montado em `criarCampoInput`, que lê `respostas["S01_ValorEstimado"]` no momento da renderização e mostra/esconde uma mensagem — sem gerar bookmark novo no `.docx` (o aviso é 100% UI, não vai para o documento).

**Interfaces:**
- Produces: propriedade nova no schema do campo, `avisoLimiar: { campoValor: "S01_ValorEstimado", limite: 5000000, mensagem: "..." }`, lida por `criarCampoInput`.

- [ ] **Step 1: Adicionar a propriedade `avisoLimiar` aos dois campos**

`js/dados-tr-bens.js`:
```javascript
    { id: "S10_PrazoProgramaIntegridadeMeses", tipo: "parametro", input: "numero",
      pergunta: "Prazo para comprovação do Programa de Integridade (meses)",
      ajuda: "Anexo XXXII do RLC: exigido quando o valor anual estimado supera R$ 5.000.000,00, com prazo máximo de 50% da vigência do contrato — limitado a 9 meses quando a vigência superar 12 meses. Se não se aplicar, registre 0.",
      avisoLimiar: { campoValor: "S01_ValorEstimado", limite: 5000000,
        mensagemAcima: "O valor estimado desta contratação supera R$ 5.000.000,00 — o Programa de Integridade É EXIGIDO (Anexo XXXII do RLC).",
        mensagemAbaixo: "O valor estimado desta contratação está abaixo de R$ 5.000.000,00 — o Programa de Integridade NÃO é exigido; registre 0 neste campo." } },
```

Mesma propriedade no campo equivalente de `js/dados-tr-servicos.js`.

- [ ] **Step 2: Ler `avisoLimiar` em `js/render.js`**

Em `criarCampoInput`, depois do bloco que renderiza a `ajuda` (`.ajuda-gatilho`), adicionar:

```javascript
  if (campo.avisoLimiar) {
    const aviso = document.createElement("div");
    aviso.className = "aviso-limiar";
    wrapper.appendChild(aviso);
    const atualizarAviso = () => {
      const valor = Number(String(valorAtual ?? "").replace(/[^\d,]/g, "").replace(",", "."));
      // valorAtual é o valor do PRÓPRIO campo; o aviso depende de outro campo,
      // então lê direto do DOM do campo de referência.
      const campoValorEl = document.getElementById(campo.avisoLimiar.campoValor);
      const valorRef = campoValorEl ? Number(String(campoValorEl.value || "").replace(/[^\d,]/g, "").replace(",", ".")) : NaN;
      if (Number.isNaN(valorRef) || valorRef === 0) { aviso.textContent = ""; aviso.style.display = "none"; return; }
      aviso.style.display = "block";
      aviso.textContent = valorRef > campo.avisoLimiar.limite
        ? campo.avisoLimiar.mensagemAcima
        : campo.avisoLimiar.mensagemAbaixo;
    };
    // Atualiza no carregamento e sempre que o campo de referência mudar.
    setTimeout(atualizarAviso, 0);
    document.addEventListener("input", (e) => {
      if (e.target && e.target.id === campo.avisoLimiar.campoValor) atualizarAviso();
    });
  }
```

Nota: usar `document.addEventListener` delegado (não um listener direto no elemento, que pode não existir ainda no momento do `criarCampoInput` do campo do Programa de Integridade, já que ele vem depois na Seção 10 e o campo de valor está na Seção 1). O `setTimeout(atualizarAviso, 0)` cobre o caso de reabrir um rascunho salvo com o valor já preenchido.

- [ ] **Step 3: CSS mínimo para `.aviso-limiar`**

Em `css/estilo.css`, adicionar:
```css
.aviso-limiar {
  display: none;
  margin-top: 4px;
  padding: 8px 12px;
  font-size: 13px;
  border-radius: 4px;
  background: #fff3cd;
  border: 1px solid #ffe69c;
  color: #664d03;
}
```

- [ ] **Step 4: Verificar manualmente no navegador**

Abrir o TR-Bens, preencher `S01_ValorEstimado` com um valor > 5.000.000, navegar até a Seção 10 e confirmar que o aviso aparece; testar também com valor abaixo do limiar.

- [ ] **Step 5: Commit**

```bash
git add js/dados-tr-bens.js js/dados-tr-servicos.js js/render.js css/estilo.css
git commit -m "Add automatic Programa de Integridade threshold warning based on valor estimado"
```

---

### Task 6: Valores sugeridos de multa e prazo de pagamento — TR-Bens

**Contexto:** A minuta-fonte de TR-Bens não tem uma tabela própria de percentuais sugeridos de multa (diferente de TR-Serviços/TR-Dispensa, que têm). **Decisão de design, sinalizada para a usuária confirmar**: reaproveitar a MESMA tabela já usada em TR-Serviços/TR-Dispensa (15%/7,5%/5%/3,75%/3% compensatória; 20%/10%/6,6%/5%/4% rescisória, por faixa de duração do contrato) — o princípio é genérico do RLC (razoabilidade e proporcionalidade, art. 578, §§1º-2º), não específico de serviços. Para o prazo de pagamento, sem nenhum valor sugerido em nenhuma fonte, adicionamos apenas uma nota orientativa (sem número fixo).

**Files:**
- Modify: `js/dados-tr-bens.js:135-142` (multas) e `:96-97` (pagamento)

- [ ] **Step 1: Ajuda de multa compensatória/rescisória**

```javascript
    { id: "S15_MultaCompensatoriaTotal", tipo: "parametro", input: "numero",
      pergunta: "Multa compensatória por inexecução total (% sobre o valor total do contrato)",
      ajuda: "Nota explicativa da minuta de TR-Serviços, reaproveitada por analogia (o princípio do art. 578, §§1º-2º, do RLC é genérico, não específico de serviços) — parâmetros sugeridos, podendo ser reduzidos conforme o valor do contrato: duração até 1 ano — 15%; 2 anos — 7,5%; 3 anos — 5%; 4 anos — 3,75%; 5 anos — 3% sobre o valor total do contrato." },
    { id: "S15_MultaRescisoria", tipo: "parametro", input: "numero",
      pergunta: "Multa rescisória por rescisão unilateral (% sobre o valor total do contrato)",
      ajuda: "Nota explicativa da minuta de TR-Serviços, reaproveitada por analogia — parâmetros sugeridos, podendo ser reduzidos conforme o valor do contrato: duração até 1 ano — 20%; 2 anos — 10%; 3 anos — 6,6%; 4 anos — 5%; 5 anos — 4% sobre o valor total do contrato." },
```

- [ ] **Step 2: Ajuda de prazo de pagamento**

```javascript
    { id: "S11_PrazoPagamentoDiasUteis", tipo: "parametro", input: "numero",
      pergunta: "Prazo de pagamento após recebimento definitivo (dias úteis, máx. 30 — art. 562 RLC)",
      ajuda: "Art. 562 do RLC: o teto é 30 dias úteis. Não há um valor sugerido pela minuta-fonte — compete a cada Regional/Matriz estipular, conforme a própria rotina administrativa, um prazo razoável dentro desse limite." },
```

- [ ] **Step 3: Rodar a suíte** (só `ajuda`, sem impacto no `.docx`)

```bash
node --test tests/*.test.mjs
```

- [ ] **Step 4: Commit**

```bash
git add js/dados-tr-bens.js
git commit -m "Add suggested multa values (by analogy with TR-Serviços) and prazo de pagamento guidance to TR-Bens"
```

---

### Task 7: Reconstruir a tabela da Matriz de Riscos (cabeçalho de dois níveis + coluna Objeto + coluna numerada)

**Files:**
- Modify: `templates/build_template_tr_bens.py` (função `add_matriz_de_riscos_rlc`, linhas ~72-134)
- Modify: `templates/build_template_tr_servicos.py` (mesma função duplicada)
- Modify: `templates/build_template_tr_dispensa.py` (mesma função duplicada)

**Contexto:** Comparação direta com `anexos RLC/anexo_ii_minuta_de_matriz_de_riscosodt.odt` confirmou que a estrutura oficial tem: (a) uma coluna "1. Objeto" que falta na nossa tabela; (b) um cabeçalho de DOIS NÍVEIS — grupos "Identificação" (abrangendo Objeto/Evento de Risco/Causas/Consequências), "5. Avaliação" (Probabilidade/Impacto), "Tratamento ao Risco" (Resposta/Responsável) acima dos rótulos de coluna individuais; (c) uma coluna de índice numérico (1, 2, 3, 4, 5) como primeira coluna, separada da numeração dos itens textuais.

**Interfaces:**
- Nenhuma mudança de campo JS — só a estrutura visual da tabela no Python. O bookmark `nome_bookmark_objeto` (parâmetro existente da função) agora vai na nova coluna "Objeto".

- [ ] **Step 1: Reescrever `add_matriz_de_riscos_rlc` (aplicar a mesma mudança nos 3 arquivos)**

A tabela precisa de 10 colunas: Nº, Objeto, Evento de Risco, Causas, Consequências, Probabilidade, Impacto, Nível de Risco, Resposta, Responsável. O cabeçalho de grupo (linha 1, mesclada) cobre: "Identificação" sobre as colunas 1-4 (Objeto/Evento de Risco/Causas/Consequências); "5. Avaliação" sobre as colunas 5-6 (Probabilidade/Impacto) — a coluna "Nível de Risco" fica fora desse grupo, como no oficial; "Tratamento ao Risco" sobre as colunas 8-9 (Resposta/Responsável).

```python
def add_matriz_de_riscos_rlc(nome_bookmark_objeto):
    """Anexo II do RLC — Minuta de Matriz de Riscos, transcrita
    integralmente (cabeçalho de dois níveis, coluna de índice numérico e
    coluna de Objeto, conforme a estrutura real do anexo oficial) a
    partir de anexos RLC/anexo_ii_minuta_de_matriz_de_riscosodt.odt."""
    add_paragraph_with_bookmarks(["Objeto: ", (nome_bookmark_objeto, "")])
    tabela = doc.add_table(rows=2, cols=10)
    tabela.style = "Table Grid"

    tabela.cell(0, 1).merge(tabela.cell(0, 4))
    tabela.cell(0, 1).text = "Identificação"
    tabela.cell(0, 5).merge(tabela.cell(0, 6))
    tabela.cell(0, 5).text = "5. Avaliação"
    tabela.cell(0, 8).merge(tabela.cell(0, 9))
    tabela.cell(0, 8).text = "Tratamento ao Risco"

    cabecalhos = ["Nº", "1. Objeto", "2. Evento de Risco", "3. Causas",
                  "4. Consequências", "Probabilidade (P)", "Impacto (I)",
                  "6. Nível de Risco (P) x (I)", "7. Resposta ao evento do Risco",
                  "8. Responsável"]
    for i, texto in enumerate(cabecalhos):
        tabela.cell(1, i).text = texto

    for indice in ["1", "2", "3", "4", "5"]:
        linha = tabela.add_row()
        linha.cells[0].text = indice

    doc.add_paragraph(
        "1. Descrição do objeto previsto para contratação. 2. O risco é um evento incerto "
        "que, se ocorrer, afeta a realização dos objetivos da contratação. 3. Condições que "
        "viabilizam a concretização de um evento de risco. 4. Identificação de quais são as "
        "consequências no caso da ocorrência do risco. 5. A avaliação da probabilidade e do "
        "impacto deverá ser analisada em uma escala de 1 a 5, conforme definida nas tabelas "
        "abaixo:")
    add_tabela(
        ["Escala de Probabilidade (P) – Associada às chances do risco ocorrer", "", ""],
    )
    add_tabela(
        ["Descritor", "Descrição", "Nível"],
        linhas=[
            ("Muito Baixa", "Evento extraordinário, sem histórico de ocorrência.", "1"),
            ("Baixa", "Evento casual e inesperado, muito embora raro, há histórico de sua ocorrência.", "2"),
            ("Média", "Evento esperado, de frequência reduzida, e com histórico de ocorrência parcialmente conhecido.", "3"),
            ("Alta", "Evento usual, com histórico de ocorrência amplamente conhecido.", "4"),
            ("Muito Alta", "Evento repetitivo e constante.", "5"),
        ],
    )
    add_tabela(
        ["Escala de Impacto (I) – Associada às consequências no caso do risco ocorrer", "", ""],
    )
    add_tabela(
        ["Descritor", "Descrição", "Nível"],
        linhas=[
            ("Muito Baixo", "Impacto insignificante nos objetivos.", "1"),
            ("Baixo", "Impacto mínimo nos objetivos.", "2"),
            ("Médio", "Impacto mediano nos objetivos, com possibilidade de recuperação.", "3"),
            ("Alto", "Impacto significante nos objetivos, com possibilidade remota de recuperação.", "4"),
            ("Muito Alto", "Impacto máximo nos objetivos, sem possibilidade de recuperação.", "5"),
        ],
    )
    doc.add_paragraph(
        "6. Após o resultado do cálculo da probabilidade x impacto será obtido o nível do "
        "risco, que poderá ser classificado como baixo, médio, elevado e extremo, conforme "
        "tabela abaixo:")
    add_tabela(
        ["Faixa (P x I)", "Nível de Risco"],
        linhas=[
            ("1 a 2", "Baixo"),
            ("3 a 6", "Médio"),
            ("8 a 12", "Elevado"),
            ("15 a 25", "Extremo"),
        ],
    )
    doc.add_paragraph(
        "7. Tratar o risco consiste em propor ações para prevenir, transferir, mitigar ou "
        "aceitar o risco. Neste campo, deve-se descrever a ação/resposta mais adequada para "
        "tratamento do risco identificado. 8. Identificar o responsável/os responsáveis pela "
        "ação proposta.")
    add_tabela(["Elaborado por: _________________________________________________ Nome e Assinatura"])
```

- [ ] **Step 2: Confirmar que `merge()` está disponível na versão do `python-docx` instalada**

```bash
python -c "import docx; print(docx.__version__)"
python -c "
import docx
doc = docx.Document()
t = doc.add_table(rows=2, cols=4)
t.cell(0,0).merge(t.cell(0,1))
print('merge ok')
"
```

Se `merge` não existir ou lançar erro na versão instalada, ajuste a estratégia: em vez de mesclar, repita o texto do grupo como um parágrafo `doc.add_paragraph()` logo acima da tabela em vez de tentar uma célula mesclada (mais simples, ligeiramente menos fiel à formatação visual, mas mantém a informação).

- [ ] **Step 3: Aplicar a mesma reescrita nos 3 arquivos, regenerar e testar**

```bash
cd templates
python build_template_tr_bens.py
python build_template_tr_servicos.py
python build_template_tr_dispensa.py
cd ..
node --test tests/*.test.mjs
```

Observação: os testes existentes que verificam a Matriz de Riscos (ex.: presença das colunas "Resposta ao evento do Risco", "Responsável") devem continuar passando, já que os rótulos de texto se mantêm — só a estrutura de linhas/colunas muda. Se algum teste falhar por causa da nova coluna "Objeto" ou dos cabeçalhos de grupo, ajuste o teste para refletir a estrutura correta (não a antiga).

- [ ] **Step 4: Commit**

```bash
git add templates/build_template_tr_bens.py templates/build_template_tr_servicos.py templates/build_template_tr_dispensa.py templates/Template_TR_Bens.docx templates/Template_TR_Servicos.docx templates/Template_TR_Dispensa.docx
git commit -m "Rebuild Matriz de Riscos table to match official anexo_ii structure (2-level header, Objeto and index columns)"
```

---

## Fora do escopo deste plano

- **Reescrita completa das Obrigações da Contratante/Contratada** (incluindo o "h)" órfão do item 10 do TR-Bens) — achado anterior, maior, ainda pendente como item separado.
- **Bug de numeração da Nota Técnica** (Seção "14" sem correspondência no documento) — achado anterior, ainda pendente.
- **SRP/TIC** — spec e plano próprios já escritos, aguardando execução.
- **Matriz de Riscos do PB** (que usa "Fase" em vez de "Objeto" — estrutura diferente, ligada ao Anexo III) — não investigada nesta rodada; PB usa uma tabela própria, não a função compartilhada `add_matriz_de_riscos_rlc`.
- Os ~19 pontos do checklist de cláusulas (habilitação, vinculação, cessão de direitos, garantias, prescrição de reajuste por duração, etc.) estão sendo verificados separadamente por um workflow de revisão — entram num plano próprio depois que os achados chegarem.

## Execution Handoff

Após as 7 tasks, rodar `node --test tests/*.test.mjs` mais uma vez para confirmar a suíte inteira verde, depois seguir para `superpowers:finishing-a-development-branch` se este trabalho estiver numa branch separada (ou seguir commitando direto em `master`, como o restante desta sessão).
