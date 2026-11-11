# Design — Projeto Básico por Dispensa

## Contexto e motivação

Extensão do PB existente (não um instrumento novo) para cobrir a
contratação de obras/serviços de engenharia por dispensa de licitação
(art. 416, I, do RLC — valor global até R$ 140.000,00). **Sem minuta SEI
de referência** para esta variante — a usuária confirmou não ter uma.
Diferente de todos os outros 5 instrumentos deste projeto, o conteúdo
específico da dispensa (a Minuta de Contrato) é derivado diretamente do
texto do RLC, não transcrito de um documento real da Conab. Decisão
explícita da usuária, ciente do risco.

## Achado 1 — não há Contrato Simplificado para obras (art. 435, II, do RLC)

O Contrato Simplificado (Anexo XI do RLC) só vale para "serviços de
pronto pagamento de valor não superior a 20% do limite estabelecido no
inciso II do Artigo 416" — ou seja, até R$ 14.000, e só para *serviços*,
não obras. A dispensa de obras/serviços de engenharia (art. 416, I) nunca
se enquadra nessa hipótese. Logo, **PB por Dispensa não tem a escolha
binária Contrato Simplificado x Contrato** que o TR-Dispensa tem — é
sempre Contrato completo. Isso simplifica o mecanismo: um único anexo
condicional, sem variante interna.

## Achado 2 — o que muda no corpo do PB é mínimo

Localizei, por grep no `build_template_pb.py` atual, todos os pontos que
pressupõem um processo licitatório (menção a "Edital", "licitação"):
Seção 2.5 (modalidade), Seção 3 (linha "Modalidade licitatória"), Seção 9.3
(prazo de vistoria "a partir da publicação do Edital"), Seção 34.2
("revogar a licitação"). **Todo o resto do PB (garantias, sanções,
fiscalização, ART/RRT, anexos técnicos, Matriz de Riscos) permanece
idêntico** — o art. 108/157 do RLC exige o mesmo conteúdo
independentemente do regime de contratação.

## Achado 3 — a Minuta de Contrato do PB deriva do art. 438 do RLC

O art. 438 lista as 30 partes essenciais de qualquer minuta de Contrato
regida pelo RLC — é a mesma base legal que já sustentou, implicitamente,
a Minuta de Contrato do TR-Dispensa (Anexo VI), construída nesta mesma
sessão a partir de uma minuta real. Reaproveito a redação já validada do
TR-Dispensa para as cláusulas genéricas (legislação aplicável, foro,
vedações, nepotismo, alteração de contrato — todas RLC-genéricas, não
específicas de TR), e escrevo as cláusulas de objeto cruzando os campos
que o PB já coleta (regime de execução, ART/RRT, garantia de obra, matriz
de riscos, sanções), em vez de reinventar conteúdo.

**Lacuna encontrada no TR-Dispensa:** falta a cláusula de Proteção de
Dados Pessoais (art. 438, XXX, do RLC). Incluída na nova minuta do PB;
correção do TR-Dispensa fica como item separado, fora deste escopo.

**Lacuna pré-existente do PB (não introduzida por esta mudança):** o PB
nunca coletou ND/PTRES/Fonte (dotação orçamentária) — nem a variante de
licitação tinha esses campos. A cláusula de dotação orçamentária da nova
minuta usa 3 campos novos, preenchidos só no anexo (mesmo padrão que o
TR usa no corpo, mas aqui isolado ao anexo por não haver equivalente no
corpo do PB).

## Decisão estrutural

Novo campo `PB01_RegimeContratacao` (Licitação / Dispensa de Licitação),
`blocoTexto`, padrão "Licitação", com `blocosAdicionais` controlando:
1. A variante da Seção 2.5 (modalidade x fundamento da dispensa).
2. A variante da linha "Modalidade licitatória" da Seção 3.
3. A variante da Seção 9.3 (vistoria).
4. A variante da Seção 34.2 (revogação x anulação).
5. O anexo inteiro da Minuta de Contrato (sentinela, só existe se
   "Dispensa de Licitação").

O campo `PB02_ModalidadeLicitatoria` ganha `mostrarSe: { PB01_RegimeContratacao: "Licitação" }`
— não faz sentido perguntar a modalidade quando é dispensa.

**Mecanismo:** `add_bloco_variante` do `build_template_pb.py`, hoje, só
aceita parágrafos de texto puro (sem bookmark aninhado) — precisa ser
generalizada para aceitar também tuplas `(nome_bookmark, texto_inicial)`
dentro da lista de parágrafos de cada variante, igual ao que
`add_paragraph_with_bookmarks` já faz. Sem isso não dá para manter o
bookmark `PB03_ModalidadeLicitatoria` (já usado por `duplicarPara`)
dentro da variante "Licitação". Generalização pequena, sem mudar
nenhuma chamada existente (todas continuam passando só strings).

## Duas frases sem citação exata do RLC (assinaladas para revisão)

Duas conexões textuais não têm uma citação de artigo específica por trás
— são frases de transição razoáveis, não deduzidas de um dispositivo
preciso do RLC:
- Seção 9.3 (dispensa): "O prazo para a vistoria será definido no
  processo de dispensa de licitação, conforme cronograma estabelecido
  pela área demandante." (mais genérica que a versão de licitação, de
  propósito, para não inventar um prazo específico que o RLC não fixa
  para dispensa).
- Seção 34.2 (dispensa): "A Conab poderá anular o processo de dispensa
  por ilegalidade, mediante despacho fundamentado, assegurados o
  contraditório e a ampla defesa." (usa vocabulário já presente em
  outras cláusulas RLC-padrão do próprio documento, mas não é uma
  transcrição de um artigo específico).

## Fora do escopo

- Corrigir a lacuna de Proteção de Dados Pessoais no TR-Dispensa (tarefa
  separada, se a usuária quiser).
- Coletar ND/PTRES/Fonte no corpo do PB (ficam só na Minuta de Contrato,
  como campos novos e isolados).

## Testes

Mesma convenção dos outros instrumentos: testes de geração cobrindo
`PB01_RegimeContratacao` = "Licitação" (comportamento idêntico ao PB
atual — nenhuma regressão) e = "Dispensa de Licitação" (anexo de
Contrato presente, cláusulas de modalidade/vistoria/revogação trocadas,
campos de dotação orçamentária e ART/RRT/garantia/sanções corretamente
cruzados do corpo do PB para o anexo).
