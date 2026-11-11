# Design — Campos SRP e TIC nos TRs

## Contexto e motivação

A usuária levantou a ideia de criar um instrumento de Ata de Registro de
Preços (ARP). Investigação no RLC (arts. 327-360) mostrou que a minuta da
ARP final é elaborada pela área de contratos (Matriz ou Sureg), não pela
área demandante — a competência da GEPEA se encerra na formulação do TR/PB
que antecede a licitação por SRP (art. 330). A usuária decidiu não criar o
instrumento da Ata e, em vez disso, cobrir o que realmente é trabalho da
área demandante: um campo SRP nos TRs, sinalizando essa modalidade e
ajustando o que o RLC efetivamente exige do TR quando ela é usada.

Uma revisão adversarial com três agentes independentes (verificação
jurídica, conflito de schema, completude comparada à ferramenta da AGU)
identificou que o escopo mínimo original (só um Sim/Não) deixava de fora
exigências textuais do próprio RLC. Também surgiu, por analogia, um campo
TIC com base legal própria (arts. 599-600), sem relação com SRP.

## Escopo: SRP só em TR-Bens e TR-Serviços

RLC art. 327: *"A licitação para registro de preços será realizada
conforme diretrizes..."* — não há, em nenhuma parte do RLC, previsão de
uso de SRP fora de licitação (busca full-text por "dispensa" + "SRP"/
"registro de preços" não retornou nenhuma hipótese). Logo, SRP não se
aplica a TR-Dispensa.

A usuária confirmou explicitamente que o PB (obras/serviços de
engenharia) também fica de fora — os artigos do RLC sobre PB (154-155)
nunca mencionam SRP, e na prática obras/engenharia raramente usam
registro de preços.

## Campo SRP (Sim/Não) — TR-Bens e TR-Servços, Seção 01

`S01_SRP`, `blocoTexto`, padrão "Não". Quando "Sim":

1. **Nota sobre dotação orçamentária** (`mostrarSe: {S01_SRP: "Sim"}`,
   próxima aos campos ND/PTRES/Fonte já existentes): pelo art. 331 do
   RLC, dotação orçamentária não precisa ser indicada nesta fase — só na
   formalização do Contrato (art. 332).
2. **Critério de julgamento perde a opção "Maior desconto"**: art.
   327/328 exige que SRP use exclusivamente "menor preço". Hoje o campo
   `S01_CriterioJulgamento` é embutido diretamente como bookmark numa
   frase fixa do template (`"1.3. O critério de julgamento das propostas
   será o de [S01_CriterioJulgamento] ofertado."`), sem `blocoTexto` — só
   insere o texto da resposta. Solução sem tocar no motor: esconder o
   campo `S01_CriterioJulgamento` quando SRP=Sim (`mostrarSe: {S01_SRP:
   "Não"}`, com "Não" sendo o padrão) e envolver essa frase inteira num
   `add_bloco_variante` chaveado em `S01_SRP` — variante "Não" mantém a
   frase de hoje (bookmark com a resposta do usuário); variante "Sim" usa
   a mesma frase com "menor preço" fixo no texto, citando o art. 328.
   Assim nunca existe a possibilidade de o usuário escolher "Maior
   desconto" com SRP=Sim — a pergunta simplesmente não aparece nesse
   caso.
3. **Campo de justificativa do motivo do SRP** (`S01_JustificativaSRP`,
   textarea, `mostrarSe: {S01_SRP: "Sim"}`): arts. 127-129 e art. 330,
   parágrafo único, do RLC exigem que a área demandante fundamente, na
   Nota de Demanda ou no TR, qual das hipóteses do art. 327 (I-IV) ampara
   o uso do SRP — não basta a flag Sim/Não.
4. **Exclusivo de TR-Bens**: campo de quantidade mínima a ser cotada por
   item (`S05_QuantidadeMinimaCotadaPorItem`, textarea, `mostrarSe:
   {S01_SRP: "Sim"}`, Seção 05 — DA SELEÇÃO DO FORNECEDOR). Art. 335, III,
   do RLC exige essa informação especificamente "no caso de bens" — não
   se aplica a TR-Serviços.

## Campo TIC (Sim/Não) — TR-Bens, TR-Serviços e PB, Seção 01

`S01_TIC` / `PB01_TIC` (nomes de acordo com a convenção de cada arquivo),
`blocoTexto`, padrão "Não", independente do SRP. Arts. 599-600 do RLC:
contratações de Tecnologia da Informação devem, obrigatoriamente, estar
previstas no PDTI/PDTIC ou ser previamente aprovadas pelo Comitê
Estratégico de Tecnologia da Informação. Quando "Sim", nota explicativa
citando essa exigência (sem inventar o conteúdo do PDTIC em si — é um
documento externo ao TR).

## Fora do escopo desta rodada

- **Instrumento de Ata de Registro de Preços**: descartado — não é
  competência da área demandante (GEPEA).
- **Dois Estudos Preliminares para SRP** (art. 108, §§5º-6º): quando a
  Conab gerencia um SRP, o RLC exige um Estudo Preliminar completo e
  outro específico para a Ata. Isso afeta Nota de Demanda/Nota Técnica,
  não TR/PB — registrado aqui para uma eventual rodada futura, não
  implementado agora.
- **Responsabilidades administrativas pós-licitação** (arts. 351-353):
  gerenciar a ARP assinada, renegociar preços etc. são atos operacionais
  da área demandante depois da licitação — não geram campo de formulário,
  já que não fazem parte do conteúdo do TR em si.

## Testes

Mesma convenção dos demais instrumentos: geração básica de TR-Bens/
TR-Serviços com SRP="Não" (comportamento idêntico ao atual, sem
regressão); com SRP="Sim" (nota de dotação orçamentária aparece,
critério de julgamento força "menor preço", campo de justificativa
aparece); TR-Bens especificamente com o campo de quantidade mínima por
item; TIC="Sim" nos três instrumentos (TR-Bens, TR-Serviços, PB) mostra a
nota do PDTIC/Comitê.
