# Design — Gerador de TR e PB (Conab / RLC 10.901)

## Contexto e motivação

A Conab se rege pela Lei 13.303/2016 e pelo Regulamento de Licitações e
Contratos (RLC 10.901), não pela Lei 14.133/2021. O site de elaboração de
edital da AGU (`cgu.agu.gov.br/edital`) demonstrou que um formulário guiado,
com interface amigável e textos de ajuda contextuais, reduz drasticamente o
tempo gasto em preenchimento repetitivo de um instrumento licitatório.

O objetivo deste projeto é replicar essa experiência para os instrumentos de
planejamento da contratação da Conab — Termo de Referência (TR) e Projeto
Básico (PB) — para que o elaborador da GEPEA concentre energia na
especificação técnica do objeto, não nos campos administrativos repetitivos.

Este é um projeto novo e isolado, sem relação de código com o gerador de PB
já existente em `Projeto-Basico-App` (`Gerador_PB.html`). A razão não é
apenas o TR não ter tido gerador antes: é que a interface daquele projeto
não é considerada amigável o suficiente pela usuária, e o padrão visual que
se quer replicar aqui é o do site da AGU (gov.br Design System, fluxo
guiado, textos de ajuda contextuais) — para os **dois** instrumentos, não só
para o TR. Nenhum arquivo daquele projeto é reaproveitado; apenas a técnica
de geração de `.docx` via manipulação de bookmarks no navegador, e o
conteúdo normativo do modelo institucional de PB (ver Escopo), são
reutilizados como referência.

## Escopo da v0

**Dentro do escopo:**
- Termo de Referência (TR), nas três variantes já usadas pela GEPEA e
  fornecidas como modelo real do SEI:
  - TR — Bens (SEI 55379667)
  - TR — Serviços (SEI 55379668)
  - TR — Serviços por Dispensa (SEI 55379661)
- Projeto Básico (PB), reaproveitando como fonte normativa o modelo
  institucional de 35 seções já validado pela GEPEA (Art. 157 do RLC,
  `assets/modelo_institucional_projeto_basico.txt` da skill
  `gepea-instrumentos`). O conteúdo normativo (títulos de seção, incisos,
  texto fixo) é portado para o modelo de dados deste projeto novo; nenhum
  arquivo do projeto `Projeto-Basico-App` (`Gerador_PB.html`) é lido em
  tempo de execução nem modificado — a relação é só de referência de
  conteúdo, não de código.
- Anexos padrão do RLC aplicáveis a cada variante (TR e PB), inseridos
  **limpos** (estrutura oficial, sem conteúdo gerado), exceto os campos
  triviais que já foram coletados em outra parte do formulário (ex.:
  endereço, objeto, nome da unidade demandante) — esses são preenchidos
  automaticamente no cabeçalho do anexo.
- Botões de salvar/carregar rascunho.

**Fora do escopo da v0** (mas a arquitetura não deve impedir extensão
futura):
- Nota Técnica (NT) — explicitamente excluída pela usuária.
- Geração automática de conteúdo técnico de anexos (ex.: preenchimento de
  linhas da Matriz de Riscos, indicadores de IMR). A usuária prefere pedir
  isso a uma IA separadamente ao final, ou preencher manualmente sobre o
  anexo limpo.
- Backend, autenticação, ou qualquer persistência fora da máquina do
  elaborador.
- Hospedagem definitiva: o artefato é entregue pronto e neutro (HTML/JS
  estático); a decisão de onde hospedar (subdomínio próprio, intranet,
  Nextcloud institucional) fica com a TI da Conab, a ser tratada depois de
  pronto o protótipo.

## Achado que orienta a arquitetura

As três minutas de TR fornecidas pela GEPEA têm exatamente as mesmas 20
seções, na mesma ordem e com os mesmos títulos (a única variação é o nome da
seção 6: "DA AQUISIÇÃO" em Bens vs. "DOS SERVIÇOS" em Serviços). Entre 93% e
96% do texto de cada minuta já é fixo e oficial; o que varia é um conjunto
de pontos de preenchimento (~65 em Bens/Serviços, ~140 em Dispensa, que tem
sub-seções adicionais de habilitação e vistoria).

Conclusão de escopo: o sistema **não redige** o TR. Ele **coleta respostas
estruturadas e preenche a minuta oficial já existente**, substituindo os
pontos de preenchimento (hoje marcados como `______` ou `(XX)`) pelo dado
informado ou pelo trecho de texto condicional correspondente.

Confirmado também que as minutas reais de TR **não usam** a lógica de seção
"aplicável / não aplicável". Nas três minutas, as 20 seções estão sempre
presentes; o que varia dentro de cada seção é o conteúdo condicional (ex.:
seção 13, Subcontratação, sempre existe, mas o texto muda entre "admite-se
subcontratação parcial, limitada a X%" e "não se admite subcontratação do
objeto", dependendo da resposta do elaborador).

**Decisão: o PB segue a mesma regra — todas as 35 seções sempre
presentes, sem toggle de aplicabilidade na interface.** No modelo
institucional de 35 seções hoje existente (`Gerador_PB.html`), apenas 3
seções em 35 variam de fato por tipo de objeto (Amostra/Demonstração
Prévia, Critérios de Similaridade, Quadro de Inovações para Obras e
Serviços de Engenharia) — as outras 32 são aplicáveis em praticamente todo
caso real. Isso torna um toggle de seção inteira um custo de interface para
cobrir uma exceção pequena.

O Art. 157, §1º, do RLC continua sendo respeitado, só que sem widget
dedicado: nessas 3 seções, a "aplicabilidade" vira uma pergunta de
parâmetro comum (ex.: "este objeto exige amostra ou demonstração prévia?
Sim/Não"), e quando a resposta é "Não", o próprio texto gerado da seção já
nasce redigido como a justificativa de não aplicação exigida pelo
Regulamento — a seção nunca desaparece da tela nem do documento, só muda de
conteúdo. É o mesmo mecanismo do TR (seção 13, Subcontratação), aplicado
aqui às 3 seções do PB que realmente variam.

Consequência de arquitetura: TR e PB usam o **mesmo modo de seção** —
sempre presente, variação resolvida por parâmetro dentro da seção, nunca
por toggle de seção inteira. Não há necessidade de um modo de interface
diferente por instrumento.

## Arquitetura

Site estático, sem servidor, sem processo de build:

- **HTML + CSS + JavaScript puro**, usando o **gov.br Design System**
  (`@govbr-ds/core`), o mesmo usado pelo site da AGU. Fornece pronto:
  inputs, seletores, tooltips (`br-tooltip`), acordeão, mensagens de alerta,
  acessibilidade (eMAG/WCAG) já resolvida por ser o padrão federal.
- **Execução inteiramente no navegador do elaborador.** Nenhum dado
  digitado, nem o documento final, sai da máquina do usuário. Resolve
  sigilo de valor e dado sensível sem precisar de política de custódia de
  dado.
- **Geração do `.docx` no navegador**: cada variante de TR e o PB são
  arquivos `.docx`-modelo com marcadores (bookmarks) nos pontos hoje
  preenchidos manualmente. O JavaScript abre o `.docx` como arquivo zip,
  substitui cada marcador pelo texto resolvido (resposta direta ou bloco
  condicional) e gera um novo `.docx` para download. Os anexos aplicáveis
  são anexados ao final do mesmo documento.
- **Modelo de dados declarativo**: as seções de cada instrumento (20 no TR,
  35 no PB) e seus campos são descritos numa estrutura de dados (não em
  código imperativo por seção), por exemplo:

  ```js
  {
    id: "13",
    titulo: "DA SUBCONTRATAÇÃO",
    instrumento: "TR",
    campos: [
      { id: "permiteSubcontratacao", tipo: "parametro",
        pergunta: "Admite subcontratação parcial?",
        opcoes: ["Sim", "Não"] },
      { id: "percentualSubcontratacao", tipo: "parametro",
        pergunta: "Percentual máximo (%)",
        mostrarSe: { permiteSubcontratacao: "Sim" } }
    ]
  }
  ```

  No PB, as 3 seções que variam por tipo de objeto (Amostra/Demonstração
  Prévia, Critérios de Similaridade, Quadro de Inovações) seguem o mesmo
  padrão: um campo de parâmetro Sim/Não decide, por `mostrarSe`, qual bloco
  de texto entra — incluindo o texto de justificativa de não aplicação
  exigido pelo Art. 157, §1º, quando a resposta é "Não". Nenhuma seção do
  PB tem tratamento especial de visibilidade.

  Isso permite que a tela se construa automaticamente a partir da lista, e
  que ajustar um campo ou adicionar uma quarta variante no futuro seja
  edição de dado, não reescrita de lógica.

- Todo campo é tipado como **parâmetro** (decisão fechada, resolvida em
  segundos: dropdown, seleção, valor numérico) ou **especificação** (texto
  livre que exige elaboração técnica do analista). A distinção é visível na
  tela e a barra de progresso conta apenas campos de especificação — é o
  mecanismo central que concentra a energia do elaborador onde ela importa.

## Fluxo de telas

Duas telas, priorizando simplicidade e minimalismo:

**Tela 1 — Triagem.** Perguntas fechadas e curtas (natureza do objeto: bem,
serviço comum, serviço com mão de obra exclusiva, obra/serviço de
engenharia, elaboração de projeto por terceiros; se é licitação ou
dispensa). Com base nisso, o sistema decide **instrumento** (TR ou PB,
citando o dispositivo do RLC — Art. 127/154 — que fundamenta a escolha) e,
se for TR, a variante aplicável (Bens / Serviços / Dispensa); se for PB, a
variante de objeto entre as três já usadas no modelo institucional
("Manutenção predial sob demanda", "Execução de obra ou serviço específico
já projetado", "Contratação de elaboração de projeto por terceiros"), mais
uma quarta opção **"Outro"**, que libera um campo de texto livre para o
elaborador descrever a variante quando nenhuma das três se encaixa. O
elaborador confirma ou corrige antes de avançar.

**Tela 2 — Painel principal.** Duas colunas: sumário fixo e clicável à
esquerda (as 20 seções do TR ou as 35 do PB, conforme o instrumento
identificado na triagem), seção ativa em acordeão à direita, com campos
diferenciados visualmente entre parâmetro e especificação, e tooltip de
ajuda contextual ao lado de cada campo (mecanismo nativo `br-tooltip`).
Todas as seções permanecem sempre visíveis, em TR e em PB — nenhuma seção
some ou é marcada como não aplicável (ver "Achado que orienta a
arquitetura").

Um botão "Gerar Documento" abre um **painel de conferência sobreposto**
(não uma terceira tela de navegação): resumo do que ainda está pendente de
preenchimento, antes de confirmar a geração e o download do `.docx`. Isso
preserva uma checagem de qualidade final sem adicionar uma página extra ao
fluxo.

Botões "Salvar Rascunho" e "Carregar Rascunho" ficam fixos e acessíveis a
partir de qualquer tela.

## Persistência (sem backend)

Duas camadas complementares:

1. **Auto-save silencioso via `localStorage`**: o progresso é salvo
   automaticamente no navegador enquanto o elaborador preenche, protegendo
   contra fechamento acidental de aba, queda de energia ou travamento.
   Limitação: preso ao navegador/computador específico.
2. **Exportar/Importar rascunho (`.json`)**: botão que baixa um arquivo com
   todas as respostas já digitadas, e outro que carrega esse arquivo de
   volta, repopulando o formulário. Portátil entre computadores e entre
   colegas (equivalente ao "Recuperar Edital pelo código" da AGU, mas sem
   precisar de servidor — o "código" é o próprio arquivo, guardado junto ao
   processo SEI).

## Anexos do RLC

Os anexos aplicáveis a cada instrumento e variante são inseridos no
documento final na forma **limpa e oficial** do RLC — estrutura de tabela,
cabeçalhos, escalas de Probabilidade/Impacto (no caso da Matriz de Riscos)
preservadas — sem geração automática de conteúdo técnico. Apenas os campos
triviais já coletados em outra parte do formulário (endereço, objeto,
unidade demandante) são preenchidos automaticamente no cabeçalho do anexo.

- **TR**: conjunto reduzido (ex.: Matriz de Riscos — Anexo II do RLC —, IMR
  — Anexo XV —, Planilha de Custos e Formação de Preços — Anexo V — quando
  houver mão de obra exclusiva).
- **PB**: os 13 anexos já previstos no modelo institucional de 35 seções
  (Matriz de Riscos, Quadro de Inovações, IMR, declarações diversas etc.),
  na mesma forma limpa. A lista exata é herdada do modelo institucional
  existente, não redesenhada do zero.

## Identidade visual

Cabeçalho do site com a logo da Conab (fixa, padrão institucional). Campo
adicional para logo/identificação da unidade (Matriz ou SUREG específica),
quando aplicável ao documento gerado.

## Tratamento de erro

- Campo de especificação vazio ao gerar → não bloqueia; aparece listado no
  painel de conferência antes da geração final.
- Rascunho carregado que não corresponde ao instrumento/variante
  atualmente selecionado (TR Bens/Serviços/Dispensa, ou PB) → aviso
  explícito antes de sobrescrever dados, nunca substituição silenciosa.
- Falha ao montar o `.docx` (modelo corrompido, navegador incompatível) →
  mensagem de erro explicando o problema, nunca tela em branco.

## Testes

Dado o porte do projeto (site estático, sem backend), o risco real está na
geração do `.docx`, não em lógica de negócio complexa. Verificação
principal: gerar o documento para um caso preenchido por completo em cada
uma das três variantes de TR e no PB, e abrir no Word real, conferindo que
todos os marcadores foram substituídos, que as 3 seções condicionais do PB
geram a justificativa de não aplicação corretamente quando a resposta é
"Não", e que a formatação (negrito, numeração, estilos) não quebrou. Não se
justifica suíte de testes
automatizados elaborada neste porte; um checklist manual de geração cobre o
risco real.

## Convenção de redação do texto de ligação

A maior parte do texto de cada minuta (TR) e do modelo institucional (PB) é
texto oficial, copiado verbatim — não é gerado, é transcrito. Mas os
trechos de ligação que este projeto escreve (frases que costuram a resposta
do elaborador ao texto fixo, e os blocos condicionais das seções que variam
por parâmetro) precisam seguir a mesma diretriz de voz já validada pela
usuária com a skill `gepea-instrumentos`, para não terem "cara de texto
gerado por IA":

- decidir a frase primeiro, justificar depois, em frases separadas — não
  embrulhar a decisão dentro de duas ou três orações subordinadas na mesma
  frase;
- citar o dispositivo do RLC só no ponto que depende dele, não em toda
  oração do bloco;
- evitar enumeração mecânica e simétrica ("(i)... (ii)... (iii)...") onde a
  realidade do campo não é simétrica;
- variar o comprimento dos blocos entre seções, em vez de produzir todas do
  mesmo tamanho e grau de polimento.

Isso vale para todo texto de ligação autoral deste projeto — tanto no TR
quanto no PB — não é regra exclusiva da elaboração de instrumentos pela
skill.

## Decisões em aberto (não bloqueiam o início da implementação)

- Lista final de quais anexos do RLC se aplicam a cada uma das três
  variantes de TR (Bens / Serviços / Dispensa) — a confirmar lendo o corpo
  completo de cada minuta e cruzando com o Título VIII do RLC.
- Onde a TI da Conab vai hospedar o site (fora do escopo técnico do
  protótipo).
