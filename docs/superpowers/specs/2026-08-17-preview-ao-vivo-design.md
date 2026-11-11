# Design — Pré-visualização ao vivo do documento

## Contexto e motivação

Os quatro instrumentos (TR-Bens, TR-Serviços, TR-Dispensa, PB) estão
completos e mesclados no `master`, com o motor de geração de `.docx`
totalmente compartilhado (`js/gerador-docx.js`, `js/modelo.js`). O
elaborador hoje só vê o resultado final ao clicar em "Gerar Documento" —
não há como conferir, durante o preenchimento, como o texto vai ficar
nem quais trechos do documento vieram de uma resposta sua.

O editor de edital da AGU (`cgu.agu.gov.br/edital`) resolve isso com uma
segunda tela que renderiza o documento em HTML e destaca em cor os
trechos preenchidos pelo usuário (investigado nesta mesma sessão: cada
valor de campo é envolvido em `<mark><span id="fieldId">`, com uma cor
genérica para "texto alterado" e cores extras por tópico condicional
específico do edital deles). Este spec adapta essa ideia ao nosso site,
com as adaptações já decididas em conversa com o usuário:

- **Painel lateral ao vivo**, não uma tela extra sob demanda — atualiza
  conforme o elaborador digita, sempre visível ao lado do formulário.
- **Uma cor só** para todo texto que veio de resposta — sem tentar
  replicar cores por tópico como a AGU, porque nossos instrumentos têm
  bem mais campos condicionais (5–10+ por documento) do que o edital
  deles (~4), e cores por campo/tópico ficariam poluídas.
- **Nos 4 instrumentos ao mesmo tempo** — o motor já é 100% compartilhado
  entre eles, então não há razão estrutural para pilotar em só um antes.
- **Mostra o texto real do documento** (parágrafos extraídos do modelo
  `.docx`), não um resumo estruturado das respostas — mais fiel ao efeito
  "ver o documento crescendo" da AGU.

## Abordagem técnica

Três caminhos foram considerados:

**A. Reaproveitar `gerarDocx()` já existente (escolhida).** A cada
mudança relevante, chamamos a mesma função que já roda no botão "Gerar
Documento" sobre o template carregado, extraímos o `word/document.xml`
do `.docx` resultante, e um conversor novo (`js/preview-docx.js`)
transforma esse XML em HTML para exibição, marcando com `<mark>` o texto
que ficou dentro de bookmarks de campo. Zero duplicação da lógica de
resolução (blocos condicionais multi-parágrafo, `blocosAdicionais`,
`duplicarPara`, `tabela-linhas`) — o único código novo é o conversor de
exibição.

**B. Implementação paralela em HTML**, replicando a lógica de
substituição/remoção do `gerador-docx.js` mas emitindo HTML em vez de
mutar XML. Rejeitada: duplicaria toda a complexidade acumulada do motor,
com risco real de as duas implementações divergirem ao longo do tempo.

**C. Refatoração para um modelo intermediário genérico**, consumido por
dois renderizadores (`.docx` e HTML). Rejeitada para este momento: é uma
reestruturação grande, desproporcional ao ganho, e fora do orçamento
desta sessão. Pode ser revisitada no futuro se o conversor da abordagem A
crescer demais.

## Arquitetura

```
respostas muda (aoMudar)
  → persistirAutosave() [já existe, sem mudança]
  → agendarAtualizacaoPreview() [novo: debounce ~400ms]
      → gerarDocx(templateBuffer, respostas, secoes)  [já existe]
      → extrair "word/document.xml" do zip resultante
      → renderizarPreviewHtml(xml)  [novo: js/preview-docx.js]
      → substituir innerHTML do painel de preview
```

O `templateBuffer` já precisa ser buscado (`fetch`) para a geração final;
para o preview ele passa a ser buscado uma vez ao montar `tela-painel`
(não a cada clique em "Gerar Documento"), e reaproveitado nas duas
situações.

## Componentes

**`js/preview-docx.js`** (novo módulo). Exporta
`renderizarPreviewHtml(documentXml: string): string`. Usa `DOMParser`
nativo do navegador (sem biblioteca externa) para percorrer `w:body`:

- `w:p` → `<p>`, concatenando o texto de `w:r/w:t` na ordem, aplicando
  `<strong>`/`<em>` quando `w:rPr` indicar negrito/itálico.
- `w:tbl` → `<table>`, uma linha HTML por `w:tr`, uma célula por `w:tc`
  (célula processada como uma sequência de parágrafos, igual ao corpo).
- Bookmarks (`w:bookmarkStart`/`w:bookmarkEnd`) cujo `w:name` bate com um
  ID de campo (mesma convenção já usada por `substituirBookmark`) fazem o
  texto entre elas ser envolvido em `<mark class="preview-destaque">`.
- Bookmarks estruturais (`_BLOCO_*`, `_LINHA_MODELO`, sentinelas) não
  precisam de tratamento especial aqui: `gerarDocx()` já as resolveu
  (substituiu ou removeu) antes do XML chegar neste conversor — o que
  sobra no XML de saída são só bookmarks de campo (preenchidas ou vazias)
  e o texto final.
- Bookmarks de campo sem valor preenchido (texto ainda vazio ou igual ao
  placeholder) não recebem `<mark>` — só o que tem conteúdo de resposta.

**`js/app.js`** (modificado). `montarTelaPainel()` passa a:
- Buscar e guardar `templateBuffer` uma vez (reaproveitado por
  `btn-gerar` e pelo preview).
- Adicionar um elemento `<aside id="painel-preview">` no HTML da tela.
- Em `aoMudar`, além do já existente, chamar `agendarAtualizacaoPreview()`.
- `agendarAtualizacaoPreview()`: limpa um `setTimeout` pendente, agenda um
  novo para ~400ms; ao disparar, chama `gerarDocx` + `renderizarPreviewHtml`
  e escreve o resultado no `#painel-preview`.
- Chamar `agendarAtualizacaoPreview()` uma vez também ao montar a tela
  (preview inicial, mesmo sem nenhuma resposta ainda — mostra o modelo
  "cru").

**CSS** (arquivo de estilos existente do projeto). Regra
`mark.preview-destaque { background: <cor suave do tema gov.br>; }` e
layout de duas colunas para `tela-painel` (formulário à esquerda, preview
à direita) em telas largas, empilhando em telas estreitas.

## Performance

Debounce de ~400ms após a última mudança antes de rodar `gerarDocx` +
conversão — evita rodar o JSZip inteiro a cada tecla digitada. Essa é a
mesma operação que hoje já roda ao clicar "Gerar Documento" (tipicamente
< 1s para os templates atuais), então não há necessidade de indicador de
carregamento; se ficar perceptível em algum instrumento maior (PB, 34
seções), o debounce pode subir sem mudar a arquitetura.

## Erros

Se `gerarDocx` ou `renderizarPreviewHtml` lançarem exceção durante uma
atualização de preview (ex.: estado intermediário inconsistente de
`respostas` enquanto o elaborador ainda está digitando), o painel mantém
o último HTML válido renderizado — sem interromper a digitação e sem
`alert`. Só um `console.warn` para depuração. Isso é diferente do fluxo
de geração final (`btn-gerar`), que continua mostrando `alert` em erro,
porque ali o elaborador espera uma ação conclusiva.

## Testes

- **Unitários de `preview-docx.js`**: dado um XML de exemplo fixo
  (string curta simulando `word/document.xml`), verifica que parágrafos
  simples, tabelas, negrito/itálico e bookmarks de campo preenchidas
  geram o HTML esperado, incluindo o `<mark>` correto.
- **Integração por instrumento** (4 testes, um por instrumento): carrega
  o template real (mesmo helper `lerTemplate*` já usado em
  `tests/gerador-docx.test.mjs`), gera o preview a partir de um conjunto
  de respostas de exemplo, e confere que o texto esperado aparece no HTML
  resultante e que está marcado com `preview-destaque`.

## Fora do escopo

- Réplica de qualquer formatação avançada do Word além de
  negrito/itálico (não há necessidade — os templates atuais não usam
  outras formatações de destaque).
- Sincronizar scroll do painel de preview com a seção do formulário em
  foco (pode ser um refinamento futuro, não faz parte desta entrega).
- Cores diferentes por campo/seção (decisão explícita: uma cor só).
