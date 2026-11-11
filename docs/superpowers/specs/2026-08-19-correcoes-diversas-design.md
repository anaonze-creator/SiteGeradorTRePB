# Design — Correções diversas (RDC, recebimento, reajuste/repactuação, reequilíbrio, IMR)

## Contexto

Seis achados pontuais levantados pela usuária ao revisar o site, cada um
verificado contra o RLC ou a fonte oficial correspondente antes de entrar
neste desenho.

## 1. Remover "RDC" das opções de modalidade

RDC (Regime Diferenciado de Contratações) some das opções em três
campos — `Nota de Demanda` (`ND13_Modalidade`), `Nota Técnica`
(`NT13_Modalidade`) e `PB` (`PB02_ModalidadeLicitatoria`). Motivo:
decisão da usuária, antecipando uma retirada do RDC do RLC numa próxima
atualização normativa — não é um achado de fidelidade, é uma escolha
editorial dela, registrada aqui para rastreabilidade.

## 2. Prazo de recebimento provisório/definitivo — TR-Serviços e TR-Dispensa

TR-Bens e PB já têm os campos (`S06_PrazoRecebimentoProvisorioDias`/
`DefinitivoDias` e `PB16_...`, respectivamente, ambos com ajuda citando o
art. 526 do RLC — limites de 15 e 90 dias). TR-Serviços e TR-Dispensa não
têm nada equivalente. Adiciona os mesmos dois campos, mesma ajuda, na
Seção 06 (DA DESCRIÇÃO DOS SERVIÇOS) de cada um. Em TR-Dispensa, seguindo
o padrão já usado nesta seção, os campos ganham `duplicarPara` para o
Anexo VI (Minuta de Contrato), na Cláusula que trata do recebimento.

## 3. Diferenciar reajuste de repactuação

RLC, art. 3º (definições), inciso LXXXIII: *"Repactuação: espécie de
reajustamento contratual utilizada para serviços com dedicação exclusiva
da mão de obra, por meio da análise da variação dos custos contratuais
ou por índice, se for o caso..."* — e art. 501: exige interregno mínimo
de um ano, vinculado a Acordo/Convenção/Dissídio Coletivo de Trabalho
para os custos de mão de obra.

Hoje TR-Serviços e TR-Dispensa só perguntam "reajustável? (Sim/Não) +
índice" — sem distinguir os dois institutos. Novo campo
`S12_DedicacaoExclusivaMaoObra` (TR-Serviços) / `S13_...` (TR-Dispensa),
Sim/Não, controlando qual dos dois textos aparece:

- **Sim (repactuação)**: cláusula com base no art. 501 — interregno
  mínimo de 1 ano contado do Acordo/Convenção/Dissídio Coletivo para
  custos de mão de obra, e da data-base da proposta para os demais
  custos (art. 501, §§1º-4º).
- **Não (reajuste)**: mantém o texto atual (correção anual pela variação
  de índice, contado da proposta).

Nos dois casos, o campo de índice/instrumento de referência
(`S12_IndiceReajuste`) continua existindo, mas seu rótulo passa a
refletir qual dos dois institutos está em jogo.

## 4. Reequilíbrio econômico-financeiro

Ausente em todos os TRs e no PB. RLC, art. 511: *"Reequilíbrio
econômico-financeiro em sentido estrito é decorrência da teoria da
imprevisão, tendo lugar quando a interferência causadora do
desequilíbrio econômico-financeiro consistir em um fato imprevisível ou
previsível de consequências incalculáveis, anormal ou extraordinário."*
— distinto de reajuste/repactuação (que são periódicos e previsíveis).

Adiciona um parágrafo fixo (sem campo de formulário — é uma cláusula de
direito, não um dado que a área demandante preenche) na Seção "DO
REAJUSTE" de TR-Bens, TR-Serviços, TR-Dispensa e PB, logo após a
cláusula de reajuste/repactuação, informando que o contrato pode ser
objeto de reequilíbrio econômico-financeiro em sentido estrito nas
condições do art. 511 do RLC, independentemente de previsão contratual
expressa.

## 5-6. IMR do PB — substituir por modelo oficial em branco

Comparação direta com `anexos RLC/anexo_xv_instrumento_de_medicao_resultadoodt.odt`
confirma que a tabela atual em `build_template_pb.py` (colunas "Item /
Critério de avaliação / Adequado / Inadequado", com 5 linhas de critérios
inventados como "Uso de materiais com qualidade") não corresponde à
estrutura oficial. O anexo real é um **modelo em branco de
indicador**, com estas colunas fixas: Item, Descrição, Finalidade, Meta
a cumprir, Instrumento de medição, Forma de acompanhamento,
Periodicidade, Mecanismo de cálculo, Início de vigência, Faixas de
ajuste no pagamento, Sanções, Observações — mais **um exemplo
ilustrativo preenchido** ("Indicador Nº 01 – Prazo de atendimento de
demandas (OS)", com a meta, mecanismo de cálculo e faixas de ajuste
reais do documento oficial).

Troca: remove a tabela de 5 critérios inventados; usa a estrutura de
colunas oficial acima, com uma linha de exemplo preenchida com o
indicador real do anexo oficial, e o restante em branco para a Conab
preencher fora do site (Word ou SEI), conforme pedido explícito da
usuária — o próprio anexo oficial já é pensado assim (uma tabela-modelo
+ um exemplo).

## Fora do escopo

- Nenhuma mudança na Ata de Registro de Preços, no design de SRP/TIC (já
  tem spec e plano próprios, aguardando execução), nem na revisão de
  Obrigações da Contratante/Contratada e no bug de numeração da Nota
  Técnica (achados anteriores, ainda não priorizados para plano).

## Testes

Mesma convenção do projeto: geração básica de cada instrumento afetado
sem regressão; TR-Serviços/TR-Dispensa com recebimento provisório/
definitivo preenchidos; dedicação exclusiva de mão de obra Sim/Não
trocando reajuste↔repactuação; presença da cláusula de reequilíbrio nos
4 instrumentos; tabela do IMR do PB usando as colunas oficiais e o
exemplo real, sem os critérios inventados antigos.
