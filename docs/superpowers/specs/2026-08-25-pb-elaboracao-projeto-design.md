# Design — PB: Lacunas da Natureza "Elaboração de Projeto" (caso real SUREG/SC)

## Contexto e motivação

O teste real do PB-Dispensa rodado com os dados do processo SUREG/SC
(elaboração de levantamento métrico, diagnóstico institucional, estudo de
layout e projetos executivos multidisciplinares, por dispensa de licitação
— art. 416, I, do RLC, teto R$140.000,00) expôs três lacunas de conteúdo e
um problema de fundamentação, todos já mapeados pela NT/PB manual desse
mesmo processo:

1. A Seção 13 (Qualificação Técnica) sempre gera o texto de obra física
   (registro CREA/CAU + CAT de execução + parcela relevante), mesmo quando
   `PB01_NaturezaObjeto = "Elaboração de projeto, laudo ou anteprojeto por
   terceiros"` — natureza que já existe no motor do PB desde a spec
   `2026-08-17-motor-pb-design.md` e já troca a Seção 28 (custo) para
   Fator K/TRDE, mas nunca foi propagada à Seção 13.
2. A Seção 16 (Recebimento) sempre gera o rito de dois estágios (art. 533),
   mesmo quando o objeto — serviço técnico intelectual sem instalação de
   equipamento — se qualifica para o recebimento simplificado do art. 534.
3. Não existe anexo estruturado para a metodologia de custo por Fator
   K/TRDE: a Seção 28 hoje só tem um campo de texto livre
   (`PB28_MetodologiaCustoIntelectual`), sem os coeficientes K1-K4
   individualizados.
4. A cláusula de fundamentação da dispensa (art. 416, I) é texto fixo —
   "por se tratar de obra ou serviço de engenharia" — em vez de se adaptar
   à natureza do objeto quando esta é "Elaboração de projeto...".

Este plano corrige as quatro lacunas, todas condicionadas ao campo já
existente `PB01_NaturezaObjeto` (não introduz nenhum campo novo de
triagem). O caso SUREG/SC testado (ver `respostas` abaixo, "Fixture de
teste") serve de fixture de regressão para os quatro itens.

## Achado 1 — Seção 13 (Qualificação Técnica): duas bases legais distintas, não uma reescrita

**Localização atual:** `templates/build_template_pb.py:422-462`
(`13. DA HABILITAÇÃO — QUALIFICAÇÃO TÉCNICA E OPERACIONAL`), campos
`PB13_ParcelasRelevantes`, `PB13_QuantitativoMinimo`,
`PB13_PrazoValidadeCertidaoFalencia`, `PB13_CapitalCirculanteMinimo`,
`PB13_PatrimonioLiquidoMinimo` em `js/dados-pb.js:179-195`.

A seção tem duas subseções com fundamentos jurídicos diferentes — só a
primeira precisa mudar:

**13.1 (Qualificação técnica) — hoje texto de obra, deve virar
condicional.** O texto atual (registro CREA/CAU + CAT comprovando
"responsabilidade por serviços de características semelhantes às parcelas
de maior relevância técnica e econômica do objeto", com quantitativo
mínimo proporcional) pressupõe um processo licitatório de porte, com
comprovação técnico-operacional plena. Para dispensa de valor (art. 416, I)
cujo objeto é elaboração de projeto por terceiros, o item correto é o do
**Anexo VIII-B, item 1, alínea "r", do RLC**, achado por leitura integral
do anexo (`documentos de referencia/rlc.txt:17374-17376`):

> r) Consta dos autos certificação acompanhada de comprovação de que o
> contratado preenche os requisitos de habilitação e de qualificação
> mínima necessários e as razões da escolha do contratado?

Esse item é o checklist de verificação da própria contratação por
dispensa — mais compatível com o porte e o rito simplificado do que a
exigência técnico-operacional plena de qualificação por CAT/parcela
relevante desenhada para licitação de obra. Texto proposto para a
variante "Elaboração de projeto, laudo ou anteprojeto por terceiros":

> 13.1.1. Certidão de registro de pessoa jurídica no CREA/CAU, em nome da
> licitante, com validade na data de recebimento dos documentos de
> habilitação.
>
> 13.1.2. Nos termos do item 1, alínea "r", do Anexo VIII-B do RLC,
> constará dos autos certificação, acompanhada de comprovação, de que a
> contratada preenche os requisitos de habilitação e de qualificação
> técnica mínima necessários à elaboração do projeto, mediante Certidão de
> Acervo Técnico (CAT) ou instrumento equivalente em nome do(s)
> responsável(is) técnico(s), comprovando responsabilidade por serviço de
> características semelhantes ao objeto, bem como as razões da escolha da
> contratada.

O registro CREA/CAU (13.1.1) permanece — é exigência do próprio exercício
profissional regulamentado (Lei 5.194/1966 e 12.378/2010), não uma
exigência de porte licitatório, e continua necessária para qualquer
elaboração de projeto de engenharia/arquitetura. O que muda é a moldura:
sai a exigência técnico-operacional plena (parcela relevante, quantitativo
mínimo proporcional — desenhada para o rito competitivo de licitação de
obra) e entra a certificação simplificada do checklist de dispensa.

**13.2 (Qualificação econômico-financeira) — mantém os números, ganha
citação correta.** Investiguei a origem dos parâmetros hoje sem citação no
código (`PB13_CapitalCirculanteMinimo` "Referência do modelo institucional:
16,66%", `PB13_PatrimonioLiquidoMinimo` "10%") e encontrei a fonte exata:
**art. 133, §6º, do RLC** (`documentos de referencia/rlc.txt:3794-3809`):

> §6º Para efeito de habilitação econômico-financeira, a Conab deverá
> exigir: I - Balanço patrimonial e demonstrações contábeis referentes ao
> último exercício social, comprovando índices de Liquidez Geral (LG),
> Liquidez Corrente (LC), e Solvência Geral (SG) superiores a 1 (um); II -
> Capital Circulante Líquido ou Capital de Giro (Ativo Circulante –
> Passivo Circulante) de, no mínimo, 16,66% (dezesseis inteiros e sessenta
> e seis centésimos por cento) do valor estimado da contratação [...]

O art. 133 rege especificamente "diretrizes específicas para elaboração do
Termo de Referência para a contratação de serviços de mão de obra
exclusiva" (`rlc.txt:3360-3361`) — **não** elaboração de projeto. A
aplicação aqui é **por analogia**: o RLC não tem uma regra de habilitação
econômico-financeira própria para serviços técnicos intelectuais, e a
mesma fórmula do §6º já é o parâmetro que o modelo institucional do PB usa
hoje para qualquer natureza de objeto. A mudança não é de conteúdo (os
percentuais continuam os mesmos, para as duas naturezas), é de citação: a
variante "Elaboração de projeto" declara explicitamente a base legal por
analogia; a variante "Execução de obra" pode manter a redação atual
(sem prejuízo de também ganhar a citação, se a implementação achar mais
simples aplicá-la às duas).

**Decisão:** `add_bloco_variante` em `PB01_NaturezaObjeto` para o
parágrafo 13.1 (texto de obra completo vs. texto do Anexo VIII-B "r"),
mantendo 13.2, 13.3 e 13.4 comuns às duas variantes (com o acréscimo da
citação ao art. 133, §6º, por analogia, no texto de introdução de 13.2).

## Achado 2 — Seção 16 (Recebimento): variante do art. 534, condicionada à natureza

**Localização atual:** `templates/build_template_pb.py:499-512` (`16. DO
RECEBIMENTO DO OBJETO CONTRATUAL`), campos
`PB16_PrazoRecebimentoProvisorioDias`, `PB16_PrazoRecebimentoDefinitivoDias`
em `js/dados-pb.js:205-213`.

Texto integral do **art. 534 do RLC**, obtido por leitura direta
(`documentos de referencia/rlc.txt:13286-13367`, únicas ressalvas: a
extração do PDF embaralhou a ordem de I-VIII/§1º-§8º nas linhas 13315-13327
— o texto abaixo já está na ordem lógica correta, conferida contra a
numeração):

> Art. 534 O recebimento simplificado de obras e serviços de engenharia de
> valor igual ou inferior ao limite previsto no Artigo 29, inciso I, da Lei
> nº 13.303, de 2016, **desde que não se componham de aparelhos,
> equipamentos e instalações sujeitos à verificação de funcionamento e
> produtividade**, será realizado da seguinte forma:
>
> I - executado o serviço, estando o mesmo em condições de ser recebido, a
> contratada deverá comunicar ao Fiscal do Contrato ou à Comissão de
> Fiscalização, por escrito e dentro do prazo contratual fixado, a fim de
> que seja realizada vistoria para fins de recebimento;
>
> II - a emissão da comunicação referida no inciso I fora do prazo
> contratual fixado caracterizará atraso, sujeitando a contratada às
> penalidades cabíveis;
>
> III - no prazo máximo de 10 (dez) dias contados da comunicação do
> término da obra ou do serviço de engenharia, será efetuada vistoria pelo
> Fiscal do Contrato ou Comissão de Fiscalização, com vistas à emissão do
> Recibo Definitivo Simplificado;
>
> IV - havendo indicação de pendências, será concedido prazo, limitado a
> 20 (vinte) dias contados da vistoria, a fim de efetuarem-se as correções
> necessárias;
>
> V - concluídos os trabalhos relativos às pendências indicadas, a
> contratada efetuará, dentro do prazo fixado no inciso anterior, por
> escrito, comunicação ao Fiscal do Contrato ou à Comissão de Fiscalização
> solicitando a realização de nova vistoria;
>
> VI - sanada a pendência, após a comunicação escrita da contratada a que
> se refere o inciso V, será efetuada vistoria final e, verificada a
> perfeita adequação da obra ou serviço aos termos do Projeto Básico, será
> emitido o Recibo Definitivo Simplificado, em até 10 (dez) dias após
> aquela comunicação;
>
> VII - se, porventura, durante a vistoria final, verificar-se que a
> pendência apontada pelo Fiscal do Contrato ou Comissão de Fiscalização
> não foi sanada, caracterizar-se-á atraso a partir da data da vistoria
> final;
>
> VIII - em caso de constatação local da não finalização da obra ou do
> serviço e da existência de parcelas ainda não executadas ou fornecidas,
> não será reconhecido efeito à comunicação referida no inciso I, o que
> implicará no não recebimento da obra ou do serviço de engenharia e na
> caracterização de atraso caso ultrapassado o prazo contratual.
>
> A soma dos valores dos pagamentos das faturas emitidas até a última
> medição não poderá ser superior a 90% (noventa por cento) do valor
> global do Contrato. O saldo restante só poderá ser liberado após a
> emissão do Recibo Definitivo Simplificado, não podendo seu valor ser
> inferior a 10% (dez por cento) do valor global do Contrato.

**A condição de aplicabilidade do próprio caput já é exatamente a que a
usuária descreveu** ("aplicável quando o objeto não envolver verificação
de funcionamento de equipamento") — não é uma simplificação minha, é
literal do art. 534. Isso significa que a variante simplificada não pode
ser automática por natureza do objeto sozinha: mesmo dentro de
"Elaboração de projeto", teoricamente poderia haver uma entrega que
envolvesse instalação/comissionamento de equipamento (o que tornaria o
art. 534 inaplicável). Para o v0 desta correção, a variante fica
condicionada a `PB01_NaturezaObjeto = "Elaboração de projeto..."` (o caso
real SUREG/SC — puramente documental, sem instalação) mais um novo campo
de confirmação explícito, para não presumir silenciosamente a
inaplicabilidade do art. 534 quando o objeto por acaso envolver
equipamento:

- `PB16_ElegivelRecebimentoSimplificado` (parâmetro, select Sim/Não,
  `mostrarSe: { PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou
  anteprojeto por terceiros" }`, `blocoTexto: true`, `padrao: "Sim"`):
  "O objeto não se compõe de aparelhos, equipamentos ou instalações
  sujeitos à verificação de funcionamento e produtividade (art. 534,
  caput, RLC)?"

Texto da variante simplificada (16.2/16.3 substituídos, 16.1/16.4/16.5
mantidos — a responsabilidade residual da Contratada e a exigência
documental de 16.5 não são incompatíveis com o rito simplificado):

> 16.2. Nos termos do art. 534 do RLC, por se tratar de objeto sem
> aparelhos, equipamentos ou instalações sujeitos à verificação de
> funcionamento e produtividade, o recebimento seguirá o rito
> simplificado: executado o serviço, a Contratada comunicará por escrito
> ao Fiscal do Contrato ou à Comissão de Fiscalização, dentro do prazo
> contratual, para a realização de vistoria.
>
> 16.3. No prazo máximo de 10 (dez) dias contados da comunicação, será
> realizada vistoria com vistas à emissão do Recibo Definitivo
> Simplificado; havendo pendências, será concedido prazo de até 20 (vinte)
> dias para as correções, seguido de nova vistoria e, sanada a pendência,
> emissão do Recibo Definitivo Simplificado em até 10 (dez) dias.

**Decisão:** `add_bloco_variante` em `PB16_ElegivelRecebimentoSimplificado`
(campo novo, visível só quando `PB01_NaturezaObjeto = "Elaboração de
projeto..."`) para os parágrafos 16.2/16.3; quando o campo não aparece
(natureza = obra) ou é respondido "Não", mantém-se o rito de dois estágios
do art. 533 já existente.

## Achado 3 — Anexo X (Fator K/TRDE): estruturar sem inventar a fórmula

**Localização atual:** `templates/build_template_pb.py:766-789` (Seção 28,
variante "Elaboração de projeto" já existe: `_bid_28_elaboracao`), campo
`PB28_MetodologiaCustoIntelectual` (textarea livre) em
`js/dados-pb.js:295-298`.

**Achado importante:** o RLC **não define** Fator K nem TRDE em nenhum
artigo (busquei "Fator K", "TRDE", "K1"-"K4" no texto integral do
Regulamento — zero ocorrências). É uma metodologia de mercado da
engenharia consultiva, externa ao RLC. Isso significa que este projeto não
pode "transcrever verbatim" uma fórmula oficial da forma como faz com o
texto das minutas — não existe fórassim mula RLC para transcrever. A
única fonte de conteúdo hoje é a própria usuária/GEPEA.

**Como isso já é resolvido no BDI, o par direto do Fator K/TRDE**: a
Seção 28, variante "Execução de obra" (`_bid_28_execucao`,
`build_template_pb.py:769-780`), **também não codifica a equação de BDI**
— ela só coleta três valores brutos que o elaborador já calculou em
planilha externa: `PB28_CustoSemBDI`, `PB28_PercentualBDI`,
`PB28_ValorTotalComBDI`. O gerador nunca calcula nem valida a matemática;
só formata o texto de entrega. **Decisão: aplicar exatamente o mesmo
princípio ao Fator K/TRDE** — não inventar a fórmula nem os pesos de K1-K4,
apenas estruturar o que hoje é um único campo de texto livre em campos
individualizados, do mesmo jeito que o par de BDI já é individualizado,
deixando o valor de cada coeficiente inteiramente a cargo de quem
preenche.

**Campos novos** (`js/dados-pb.js`, dentro da Seção 28, variante
"Elaboração de projeto..."), substituindo
`PB28_MetodologiaCustoIntelectual` por:

- `PB28_FatorK1` (parâmetro, numero) — "Coeficiente K1 (despesas
  administrativas)"
- `PB28_FatorK2` (parâmetro, numero) — "Coeficiente K2 (custo de
  benefícios e despesas indiretas — BDI equivalente)"
- `PB28_FatorK3` (parâmetro, numero) — "Coeficiente K3 (despesas
  financeiras)"
- `PB28_FatorK4` (parâmetro, numero) — "Coeficiente K4 (lucro)"
- `PB28_FatorKResultante` (parâmetro, numero) — "Fator K resultante (soma
  dos coeficientes)"
- `PB28_TRDE` (parâmetro, numero) — "Taxa de Retribuição pela Despesa de
  Estrutura (TRDE), se aplicável"
- `PB28_MetodologiaCustoIntelectual` (especificação, textarea, mantido) —
  "Correlação entre os componentes do Fator K/TRDE e a composição do
  preço" (mesmo campo de hoje, ressignificado como a narrativa que
  acompanha os números, não mais o único conteúdo)

Os nomes e a definição de K1-K4 acima (despesas administrativas, BDI
equivalente, despesas financeiras, lucro) seguem a decomposição mais
comum descrita pela literatura de precificação de serviços de engenharia
consultiva no mercado brasileiro — **não há dispositivo do RLC que defina
esses rótulos**. Sinalizo isso como uma decisão que precisa de validação
da usuária antes da implementação (ver "Dúvida de escopo" abaixo) — os
nomes dos campos e o texto de introdução do Anexo podem mudar sem afetar
a arquitetura da correção.

**Anexo X — estrutura, análoga em nível de detalhe ao Anexo VII (IMR,
`build_template_pb.py:1173+`) e ao Anexo V do RLC (Planilha de Custos):**

```
ANEXO X DO PROJETO BÁSICO
MEMÓRIA DE CÁLCULO — FATOR K E TRDE

Nota: A metodologia de Fator K e Taxa de Retribuição pela Despesa de
Estrutura (TRDE) substitui a equação de BDI (art. 165 do RLC) para
serviços técnicos de natureza intelectual (elaboração de projeto, laudo
ou anteprojeto), sendo o cálculo de responsabilidade da área técnica.

| Componente                              | Valor |
|------------------------------------------|-------|
| K1 — Despesas administrativas            | ____  |
| K2 — Benefícios e despesas indiretas     | ____  |
| K3 — Despesas financeiras                | ____  |
| K4 — Lucro                                | ____  |
| Fator K resultante                        | ____  |
| TRDE (quando aplicável)                   | ____  |

Correlação entre os componentes: [PB28_MetodologiaCustoIntelectual]
```

Este anexo só é anexado quando `PB01_NaturezaObjeto = "Elaboração de
projeto..."` — mesmo mecanismo de sentinela condicional já usado para o
Anexo II-A (Quadro de Inovações, `_bid_anexo_iia`) e para o anexo da
Minuta de Contrato (`_bid_anexo_contrato`).

**Decisão:** `add_bloco_variante_sentinela_inicio/fim` em
`PB01_NaturezaObjeto` envolvendo o novo Anexo X inteiro, seguindo o padrão
já estabelecido no arquivo.

## Achado 4 — Fundamentação da dispensa: dois pontos no corpo, não um

**Localização real** (mais precisa que o relato do teste anterior — a
frase não vive dentro da própria Minuta de Contrato, vive no corpo
principal do PB, mas o efeito prático relatado — texto genérico de
"obra ou serviço de engenharia" sempre que a dispensa se aplica — é o
mesmo):

1. `templates/build_template_pb.py:189-199` — Seção 2.5
   (`PB02_ModalidadeOuFundamento`, variante "Dispensa de Licitação").
2. `templates/build_template_pb.py:245-251` — Seção 3, linha "Fundamento
   da dispensa" (`PB03_ModalidadeOuFundamento`, variante "Dispensa de
   Licitação").

Ambos hoje fixam: *"[...] com fundamento no art. 416, I, do RLC, por se
tratar de obra ou serviço de engenharia de valor global até R$
140.000,00"*. A base legal (art. 416, I, mesmo teto) não muda — só a
descrição do objeto precisa se adaptar, porque "elaboração de projeto,
laudo ou anteprojeto por terceiros" é uma modalidade de serviço de
engenharia (é por isso que o Projeto Básico — art. 154 do RLC, "Serão
obrigatoriamente precedidas de elaboração de Projeto Básico [...] as
licitações para a execução de obras e de serviços de engenharia" — é o
instrumento correto para este objeto, e não um Termo de Referência de
serviço comum), mas o texto hoje não deixa isso explícito.

Texto proposto para a combinação Dispensa + Elaboração de projeto:

> 2.5. A contratação será realizada por dispensa de licitação, com
> fundamento no art. 416, I, do RLC, por se tratar de elaboração de
> projeto, laudo ou anteprojeto por terceiros — modalidade de serviço de
> engenharia sujeita à elaboração de Projeto Básico nos termos do art.
> 154 do RLC —, de valor global até R$ 140.000,00 (cento e quarenta mil
> reais).

E, na Seção 3: *"Fundamento da dispensa: art. 416, I, do RLC (elaboração
de projeto, laudo ou anteprojeto por terceiros, modalidade de serviço de
engenharia, de valor global até R$ 140.000,00)."*

**Mecanismo — duas variáveis, uma variante de texto:** hoje
`add_bloco_variante` só chaveia por um único campo (`respostas[campoId]`).
A combinação "Dispensa de Licitação" + natureza precisa de uma condição
composta. Duas opções, a decidir na implementação:

- (a) aninhar `add_bloco_variante_sentinela_inicio/fim` por
  `PB01_NaturezaObjeto` **dentro** da lista de parágrafos da variante
  "Dispensa de Licitação" de `PB02_ModalidadeOuFundamento` — mesmo padrão
  que a Seção 28 já usa para aninhar uma condição dentro de outra; ou
- (b) generalizar `add_bloco_variante` para aceitar uma tupla de campos
  `(campoId1, campoId2)` e casar contra a combinação de valores.

(a) reaproveita um mecanismo já testado no arquivo (menor risco); (b) é
mais reutilizável se uma quarta combinação aparecer depois. Recomendo (a)
pelo princípo YAGNI já seguido no resto do projeto — mas deixo para o
plano decidir com mais detalhe de código.

## Fixture de teste — caso real SUREG/SC

As mesmas respostas usadas no teste manual do site, servindo de fixture
para os quatro itens (reproduzível também via
`tests/gerador-docx.test.mjs`, seguindo o padrão de
`tests/pb-dispensa.test.mjs`):

```js
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
  PB28_FatorK1: "3,5", PB28_FatorK2: "2,0", PB28_FatorK3: "1,2", PB28_FatorK4: "8,0",
  PB28_FatorKResultante: "14,7", PB28_TRDE: "10",
  PB28_MetodologiaCustoIntelectual: "Correlação entre os coeficientes K1-K4 e a TRDE, própria da engenharia consultiva.",
};
```

## Testes

Cada achado ganha teste de geração real (molde de
`tests/pb-dispensa.test.mjs`, sobre `templates/Template_PB.docx`
regenerado):

1. Com a fixture acima: Seção 13 contém "Anexo VIII-B" e "alínea \"r\"" e
   **não** contém "parcela relevante" nem "quantitativo mínimo
   proporcional"; com `PB01_NaturezaObjeto = "Execução de obra..."`,
   Seção 13 mantém o texto de obra inalterado (teste de não regressão).
2. Com `PB16_ElegivelRecebimentoSimplificado = "Sim"`: Seção 16 contém
   "art. 534" e "Recibo Definitivo Simplificado", não contém "recebimento
   definitivo... em até" (o texto de dois estágios). Com natureza = obra
   (campo nem aparece) ou `"Não"`: Seção 16 mantém o rito de dois estágios
   atual.
3. Com natureza = elaboração de projeto: documento contém "ANEXO X" e os
   seis campos de coeficiente/TRDE preenchidos; com natureza = obra: "ANEXO
   X" **não** aparece no documento (mesmo padrão de ausência condicional já
   testado para o Anexo II-A).
4. Com Dispensa + elaboração de projeto: Seção 2.5 e a linha "Fundamento
   da dispensa" contêm "elaboração de projeto, laudo ou anteprojeto por
   terceiros" e "art. 154"; com Dispensa + obra: mantém "por se tratar de
   obra ou serviço de engenharia" (não regressão do texto já validado por
   `tests/pb-dispensa.test.mjs`).

## Fora do escopo

- Qualquer alteração na Seção 21 (garantia da obra/serviço), Seção 22
  (garantia contratual) ou Seções de sanções — não foram apontadas como
  lacuna pelo teste real; ficam como estão.
- Preenchimento automático dos 15 itens da Matriz de Riscos ou dos 4
  indicadores do IMR do processo SUREG/SC — decisão de escopo já tomada no
  design original do projeto (anexos técnicos sempre "limpos", nunca
  gerados automaticamente); não reaberta aqui.
- TR-Dispensa e as demais variantes de TR não são tocados — o caso real
  testado é exclusivamente PB-Dispensa.

## Dúvida de escopo (não bloqueia o início do plano para os itens 1, 2 e 4)

Item 3 (Anexo X) depende de uma decisão que só a Ana pode validar: os
rótulos de K1-K4 que propus acima (despesas administrativas, benefícios e
despesas indiretas, despesas financeiras, lucro) são a decomposição mais
comum na literatura de precificação de engenharia consultiva no Brasil,
mas **não vêm de nenhum documento deste projeto nem do RLC** — o
Regulamento não define Fator K/TRDE. Preciso saber: a GEPEA já usa uma
planilha ou nota técnica própria com uma definição diferente de K1-K4
(ou de quantos coeficientes existem — algumas metodologias usam só K1-K3,
outras adicionam um K5 de ISS)? Se houver uma referência real, prefiro
seguir ela à decomposição genérica que propus. Se não houver, sigo com a
proposta acima (que já está estruturalmente pronta para o plano), deixando
claro que os rótulos são uma convenção de mercado, não uma exigência
regulamentar.
