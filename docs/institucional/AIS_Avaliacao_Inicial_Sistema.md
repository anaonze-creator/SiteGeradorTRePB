# AVALIAÇÃO INICIAL DE SISTEMA (AIS)

> **Nota sobre este documento**: a Norma 60.215 (Processo de Software) descreve o
> propósito e o conteúdo esperado da AIS (Capítulo IV), mas não traz um modelo/layout
> oficial no seu próprio texto — ela remete a uma Norma Interpretativa da Conab (NIC)
> específica para o detalhamento do processo, que não estava disponível para consulta
> ao escrever este documento. A estrutura abaixo segue o que a 60.215 descreve como
> conteúdo mínimo da AIS; confirme com a Sutin se existe um modelo/formulário
> específico (via SEI) a ser usado no lugar deste rascunho.
>
> Este documento também é retroativo: o sistema já está construído e em uso informal;
> a AIS aqui serve para apoiar a deliberação do CETI sobre seu registro/homologação
> formal, não para autorizar um desenvolvimento ainda não iniciado.

## 1. Origem da demanda

Ver `RDN_Registro_Necessidade_TI.md`.

## 2. Descrição da solução

Site estático (HTML/CSS/JavaScript vanilla, sem framework, sem build step) que gera
arquivos `.docx` de seis instrumentos de contratação da Conab, através de um
formulário guiado com pré-visualização do documento em tempo real:

1. Termo de Referência — Aquisição de Bens
2. Termo de Referência — Prestação de Serviços
3. Termo de Referência — Serviços por Dispensa de Licitação
4. Projeto Básico — Obras e Serviços de Engenharia (com variante por dispensa)
5. Nota de Demanda (Estudo Preliminar — Área Demandante)
6. Nota Técnica (Estudo Preliminar — Área Técnica)

## 3. Tecnologia empregada

- **Frontend**: HTML/CSS/JavaScript puro, seguindo o padrão visual do gov.br Design
  System. Sem dependências de build (sem npm/webpack em produção — o site roda
  diretamente os arquivos `.js`).
- **Geração de `.docx`**: manipulação direta de XML OOXML (`word/document.xml`) dentro
  do `.zip` que compõe um arquivo `.docx`, usando a biblioteca JSZip no navegador.
  Templates-base (`.docx` com bookmarks nos pontos de preenchimento) são construídos
  previamente com Python + `python-docx`, a partir do texto das minutas oficiais.
- **Testes**: suíte automatizada com o test runner nativo do Node.js (`node --test`),
  116 testes cobrindo geração de documento, lógica de exibição condicional de campos e
  a pré-visualização em tempo real.
- **Hospedagem**: arquivos estáticos — pode ser hospedado em qualquer servidor de
  arquivos estáticos ou aberto localmente, sem necessidade de backend.

## 4. Escopo funcional (resumo — detalhado na DRS)

- Formulário dinâmico por seções, com campos condicionais (`mostrarSe`), blocos de
  texto alternativos conforme a resposta (`blocoTexto`), e propagação de uma resposta
  para múltiplos pontos do documento (`duplicarPara`).
- Pré-visualização ao vivo do texto final, com destaque do campo em edição.
- Tabelas de linhas repetíveis para conteúdo de tamanho variável (ex.: endereços
  adicionais de execução, infrações administrativas extras).
- Anexos técnicos completos: Matriz de Riscos, Instrumento de Medição de Resultado
  (IMR), Planilha de Custos, Declaração de Contratos Firmados, Modelo de Apresentação
  de Proposta, entre outros — todos transcritos dos anexos oficiais do RLC 10.901.

## 5. Estimativa preliminar

Não aplicável no sentido prospectivo da norma (o sistema já foi desenvolvido). Como
referência retroativa: o desenvolvimento ocorreu ao longo de múltiplas sessões de
trabalho assistidas por IA (Claude Code), com commits incrementais registrados no
histórico do Git do repositório.

## 6. Riscos identificados

- **Desatualização normativa**: o conteúdo é fiel ao RLC 10.901 na versão vigente no
  momento da construção de cada instrumento. Mudanças no RLC exigem atualização manual
  do código — não há mecanismo de sincronização automática.
- **Dependência de revisão humana**: o sistema não substitui a revisão jurídica/técnica
  da Conab antes do uso do documento gerado; ele reduz erros de transcrição e
  padronização, não elimina a necessidade de validação final pela área competente.
- **Ausência de autenticação/controle de versão de quem preenche**: por ser client-side
  e sem backend, não há registro de quem gerou qual documento nem controle de acesso —
  qualquer pessoa com o link pode gerar os instrumentos.

## 7. Benefícios esperados

- Redução de erros de citação de artigos do RLC e de omissão de cláusulas obrigatórias
  (achados concretos já documentados nas revisões de conteúdo desta iniciativa — ver
  histórico de commits do repositório).
- Padronização de conteúdo entre instrumentos que tratam da mesma regra.
- Redução do tempo de elaboração de TR/PB pela área demandante.

## 8. Recomendação

Submeter este documento e a DRS (`DRS_Definicao_Requisitos_Sistema.md`) à deliberação
do CETI para registro formal do sistema no Portfólio de Soluções de TI da Conab, e
encaminhamento à Geasi para avaliação dos aspectos de segurança da informação
(conforme Capítulo VI, item 2, da Norma 60.215).
