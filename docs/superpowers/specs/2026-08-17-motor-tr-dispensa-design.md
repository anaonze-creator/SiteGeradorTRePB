# Design — Motor TR-Serviços por Dispensa

## Contexto e motivação

TR-Bens e TR-Serviços estão completos e mesclados no `master`. O próximo
instrumento é o TR-Serviços por Dispensa, minuta oficial fornecida em
`documentos de referencia/SEI_55379661_Conab___RLC__Termo_Referencia__TR____Serv_Dispensa.pdf`.
O texto foi extraído com `pdftotext -layout -enc UTF-8` e lido **por
inteiro**, do início ao fim (2.350 linhas) — nada neste documento foi
inventado.

Diferente do que a nomenclatura sugere, o TR-Dispensa não é "TR-Serviços com
sub-seções extras": é um instrumento de outra ordem de tamanho, porque a
contratação direta por dispensa não tem edital separado. Tudo que
normalmente estaria no edital — habilitação jurídica, fiscal/trabalhista,
econômico-financeira, qualificação técnica — está dentro do próprio TR. O
documento também embute, como anexo, a própria minuta de contrato (em duas
variantes possíveis), algo que TR-Bens e TR-Serviços não têm.

## Achado que orienta a arquitetura

Do meio para o fim do documento (Seção 6 "Descrição dos Serviços" em
diante: Garantia, Fiscalização, Obrigações, Pagamento, Reajuste,
Subcontratação, Sanções, Rescisão, Nepotismo, Sustentabilidade, Disposições
Finais), o texto é **praticamente idêntico** ao TR-Serviços — mesmos
mecanismos de condicional, só renumerado e trocando "licitante"/"licitação"
por "proponente"/"dispensa de licitação". Nenhum mecanismo novo é
necessário para essa metade do documento; ela reaproveita tudo o que já
existe (bloco multi-parágrafo, variante nomeada, `tabela-linhas`).

O que é genuinamente novo está concentrado em três pontos: as duas seções
iniciais exclusivas da dispensa (Apresentação da Proposta, Habilitação e
Qualificação Técnica), a escolha do instrumento contratual e sua
propagação por várias seções e anexos, e duas inversões de polaridade
(ME/EPP, Matriz de Riscos).

## Escopo

Este documento cobre a arquitetura completa (necessária para tomar as
decisões de mecanismo corretamente), mas a **implementação é dividida em
duas fases/planos**, cada uma com seu próprio plano em
`docs/superpowers/plans/` e execução em worktree isolado:

**Fase 1 — TR-Dispensa funcional (este é o próximo plano a escrever):**
- As 26 seções do corpo do TR, incluindo a nova Seção 3 (Habilitação e
  Qualificação Técnica) e Seção 2 (Apresentação da Proposta).
- A escolha de instrumento contratual (Contrato Simplificado vs. Contrato)
  e sua propagação: quais Seções 21–25 do corpo do TR sobrevivem, e qual
  dos dois anexos (V ou VI) é incluído.
- Anexo V (Contrato Simplificado) e Anexo VI (Minuta de Contrato) entram
  **limpos** — mesmo tratamento que os anexos do TR-Serviços hoje: estrutura
  oficial, sem preenchimento cruzado automático, mesmo que dupliquem
  conteúdo já respondido no corpo do TR.
- A pequena extensão de engenharia (`blocosAdicionais` no schema de campo)
  necessária para um único campo controlar mais de um bloco de bookmark.

**Fase 2 — Pré-preenchimento cruzado do Anexo VI (plano futuro, incremento
sobre a Fase 1):**
- Reaproveitar os valores já respondidos no corpo do TR para preencher
  automaticamente as ~10 cláusulas do Anexo VI que duplicam conteúdo do
  corpo (Vigência, Regime/Prazos de Execução, Garantia, Dotação
  Orçamentária, Natureza Intelectual, Sustentabilidade, Reajuste, Sanções —
  incluindo a Tabela 1/2 dupla, que aparece de novo dentro do Anexo VI).
- Generalização de `inserirLinhasTabela` para processar todas as
  ocorrências de um bookmark `_LINHA_MODELO` no documento (hoje só a
  primeira), necessária porque a tabela de infrações existe duas vezes
  quando o Contrato completo é escolhido.

**Fora do escopo de ambas as fases:**
- Dados que só existem depois da adjudicação (valor final da proposta
  vencedora, número de empenho, data de assinatura, tabela de itens da
  proposta vencedora) — ficam em branco no Anexo V/VI, como já acontece
  hoje com campos de assinatura nos outros anexos. Não há como o TR, no
  momento em que é elaborado, saber esse dado.
- Projeto Básico — spec própria futura.

## Mecanismos

### 1. Reaproveitados sem alteração

Bloco multi-parágrafo, variante nomeada (`blocoTexto`/`padrao`/`opcoes`) e
campo `tabela-linhas` — descritos em
`2026-08-17-motor-tr-servicos-design.md` — cobrem sozinhos toda a segunda
metade do corpo do TR-Dispensa (Seções 6 a 20, 26): Vistoria
(obrigatória/facultativa), Garantia dos Serviços (Sim/Não), Materiais a
Disponibilizar (Sim/Não), Fiscalização (Fiscal Funcional/Comissão, com a
mesma bifurcação de redação na Seção de Pagamento), Natureza Intelectual
(bloco de cessão de direitos), Reajuste (Sim/Não), Subcontratação (Sim/Não),
Sanções (percentuais simples + `S_BaseCalculoMulta` com bookmarks
duplicados + `tabela-linhas` para infrações extras).

### 2. Novo: `blocosAdicionais` no schema de campo

**Problema:** a escolha de instrumento contratual precisa controlar, com a
mesma resposta, **três coisas ao mesmo tempo**: (a) se as Seções 21 a 25 do
corpo do TR existem, (b) se o Anexo V sobrevive, (c) se o Anexo VI
sobrevive. O mecanismo de variante atual (`controladoresBlocoVariante` +
`nomeBlocoVariante(campo.id, opcao)`) só resolve **um** bookmark por campo
por opção — suficiente quando o campo controla um único parágrafo ou bloco,
insuficiente aqui.

**Decisão:** o campo controlador declara `blocosAdicionais: string[]` —
nomes de bookmark extras (além do próprio `campo.id`) que seguem a mesma
lógica de variante. Em `js/modelo.js`, `controladoresBlocoVariante`
continua devolvendo os campos com `blocoTexto: true`; em
`js/gerador-docx.js`, o loop de remoção passa a iterar sobre
`[campo.id, ...(campo.blocosAdicionais || [])]` para cada opção não
escolhida, em vez de só `campo.id`. Retrocompatível: campos sem
`blocosAdicionais` continuam funcionando exatamente como hoje.

```js
{ id: "S07_TipoInstrumentoContratual", tipo: "parametro", input: "select",
  pergunta: "O instrumento contratual será",
  opcoes: ["Contrato Simplificado", "Contrato"],
  blocoTexto: true, padrao: "Contrato Simplificado",
  blocosAdicionais: [
    "S07_TipoInstrumentoContratual_SECOES21A25",
    "S07_TipoInstrumentoContratual_ANEXO",
  ] }
```

No template, `S07_TipoInstrumentoContratual_SECOES21A25_BLOCO_CONTRATO_SIMPLIFICADO`
envolve as Seções 21–25 inteiras (texto presente só nessa variante;
`_BLOCO_CONTRATO` fica vazio). `S07_TipoInstrumentoContratual_ANEXO_BLOCO_CONTRATO_SIMPLIFICADO`
envolve o Anexo V inteiro e `..._BLOCO_CONTRATO` envolve o Anexo VI inteiro
— cada um demarcado por um parágrafo-sentinela vazio antes do título do
anexo e outro depois do fim do seu conteúdo (inclusive tabelas: a remoção
já opera por recorte de texto bruto entre os dois bookmarks, não por árvore
XML, então tabelas no meio são removidas de brinde).

### 3. Fase 2 apenas: `inserirLinhasTabela` processa todas as ocorrências

Hoje a função localiza a **primeira** ocorrência de
`${campo.id}_LINHA_MODELO` e para. Isso é suficiente na Fase 1 (Anexo VI
limpo não tem bookmark de linha-modelo nenhum) mas quebra na Fase 2, quando
o Contrato completo é escolhido e a Tabela 1 de infrações existe duas
vezes (corpo do TR + Anexo VI, ambas com bookmark `S15_InfracoesExtras_LINHA_MODELO`).
Na Fase 2, a função passa a rodar em loop, processando cada ocorrência
encontrada até não sobrar nenhuma — mesmo array de `linhas` clonado em cada
uma. Não há mudança de assinatura nem de comportamento observável na Fase
1 (não regressão).

## Seções novas do corpo (exclusivas da Dispensa)

- **Seção 2 — Apresentação da Proposta**: texto quase todo fixo; um único
  ponto de preenchimento real (prazo/data-limite de envio da proposta).
- **Seção 3 — Habilitação e Qualificação Técnica**: a maior seção do
  documento, ~95% texto fixo e obrigatório por lei (documentação jurídica,
  fiscal/trabalhista, qualificação econômico-financeira com as fórmulas de
  Liquidez Geral/Solvência Geral/Liquidez Corrente). Pontos de
  preenchimento reais: prazo mínimo de experiência anterior e
  quantitativo/serviço principal para comprovação técnica (mesmo padrão já
  usado em `S05_PrazoExperienciaMinimaAnos`/`S05_QuantitativoServicoPrincipal`
  do TR-Serviços), mais um campo opcional de texto livre para registro/
  autorização especial de atividade regulada (ANVISA, ANP etc.), deixado em
  branco quando não se aplica.
- **Seção "Instrumento Contratual"**: o campo `blocoTexto` com
  `blocosAdicionais` descrito acima.

## Duas inversões de polaridade

- **Exclusividade ME/EPP**: no TR-Serviços, o padrão é *sem* exclusividade
  e o campo pergunta se se aplica uma *exceção que restringe* a ME/EPP. No
  TR-Dispensa, o padrão (contratações até R$ 80.000) já É exclusivo para
  ME/EPP, e o campo pergunta se se aplica uma *exceção que abre* a disputa.
  Campo novo e próprio (não reaproveita `S02_AplicaExcecaoMEEPP` — mesmo
  nome com sentido oposto seria confuso na tela e no dado salvo):
  `S02_ExcecaoExclusividadeMEEPP`, variantes `["Não — exclusiva para ME/EPP", "Sim — exceção aplicada, sem exclusividade"]`.
- **Matriz de Riscos**: ganha uma segunda variante que o TR-Serviços não
  tem — "serviços de pronto pagamento, sem risco nas fases de
  planejamento, seleção e gestão do contrato", que dispensa a apresentação
  da Matriz de Riscos inteira (o Anexo I também some, pelo mesmo mecanismo
  `blocosAdicionais`). Campo `S19_ExigeMatrizRiscos` (variantes
  `["Sim", "Não — serviço de pronto pagamento"]`), com `blocosAdicionais`
  incluindo o Anexo I.

## Anexo V e Anexo VI (Fase 1: conteúdo limpo)

- **Anexo V — Contrato Simplificado**: texto 100% fixo, zero blanks. Mesmo
  tratamento que Matriz de Riscos/Modelo de Proposta hoje: conteúdo
  estático no script de build, sem bookmark nenhum de campo.
- **Anexo VI — Minuta de Contrato**: 26 cláusulas. Na Fase 1 entra também
  limpo (mesmo padrão), apesar de duplicar bastante conteúdo do corpo do
  TR — esse reaproveitamento automático é justamente o que a Fase 2
  resolve. A maioria das cláusulas do Anexo VI, aliás, já é só uma
  remissão ("conforme previsto no Termo de Referência"), sem duplicar
  texto — só ~10 cláusulas (Vigência, Regime/Prazos de Execução, Garantia,
  Dotação Orçamentária, Natureza Intelectual, Sustentabilidade, Reajuste,
  Sanções) de fato repetem conteúdo, e são o alvo da Fase 2.

## Testes

Mesmo critério das specs anteriores: gerar o documento cobrindo os dois
lados de cada condicional novo (Contrato Simplificado e Contrato completo —
conferindo que as Seções 21–25 e o anexo certo aparecem/somem em espelho;
exclusividade ME/EPP nos dois sentidos; Matriz de Riscos presente e
dispensada) e abrir no Word real, além dos testes automatizados de geração
já no padrão estabelecido (`tests/gerador-docx.test.mjs`).

## Decisões em aberto (não bloqueiam o início da Fase 1)

- Texto exato do campo de registro/autorização especial da Seção 3
  (ANVISA, ANP, Ministério da Agricultura etc.) — fica como campo de texto
  livre opcional; se a demanda real pedir mais estrutura (ex.: órgão +
  número + fundamento legal como campos separados), revisar na extração
  fina do plano de implementação.
- Se o valor estimado da contratação (ainda que só estimado, não o valor
  final adjudicado) deve ou não ser reaproveitado na Cláusula Quinta do
  Anexo VI ("Do Valor do Contrato") mesmo sabendo que não é o valor real —
  decisão da Fase 2, não da Fase 1.
