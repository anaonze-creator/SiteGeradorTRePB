# DEFINIÇÃO DE REQUISITOS DE SISTEMA (DRS)

> **Nota sobre este documento**: assim como a AIS, a 60.215 não traz um modelo oficial
> de DRS no seu próprio texto (remete a uma NIC específica). A estrutura abaixo é uma
> catalogação razoável dos requisitos funcionais tal como o sistema existe hoje,
> escrita retroativamente a partir do código-fonte e das especificações internas do
> projeto (`docs/superpowers/specs/`). Confirme com a Sutin se há um formulário
> específico a preencher no lugar deste documento.

## 1. Requisitos funcionais — Instrumentos gerados

| ID | Requisito | Status |
|---|---|---|
| RF-01 | Gerar Termo de Referência para Aquisição de Bens (art. 127 do RLC) | Implementado |
| RF-02 | Gerar Termo de Referência para Prestação de Serviços (art. 129 do RLC) | Implementado |
| RF-03 | Gerar Termo de Referência para Serviços por Dispensa de Licitação (art. 416, II, do RLC) | Implementado |
| RF-04 | Gerar Projeto Básico para obras e serviços de engenharia (art. 157 do RLC), incluindo a variante por dispensa de licitação (art. 416, I) com Minuta de Contrato própria | Implementado |
| RF-05 | Gerar Nota de Demanda (Estudo Preliminar — Área Demandante, arts. 108 e 115 do RLC) | Implementado |
| RF-06 | Gerar Nota Técnica (Estudo Preliminar — Área Técnica, arts. 108, 115 e 123 do RLC), incluindo seções adicionais quando o objeto é obra/serviço de engenharia | Implementado |

## 2. Requisitos funcionais — Motor de formulário

| ID | Requisito | Status |
|---|---|---|
| RF-07 | Exibir/ocultar um campo do formulário condicionalmente à resposta de outro campo (`mostrarSe`) | Implementado |
| RF-08 | Trocar o texto de uma cláusula inteira do documento conforme a resposta de um campo Sim/Não ou de múltipla escolha (`blocoTexto`), removendo do `.docx` final qualquer variante não escolhida | Implementado |
| RF-09 | Propagar a resposta de um único campo para múltiplos pontos do documento (corpo e anexos) sem exigir nova digitação (`duplicarPara`) | Implementado |
| RF-10 | Controlar mais de um bloco/seção/anexo a partir de uma única resposta (`blocosAdicionais`) | Implementado |
| RF-11 | Permitir tabelas de linhas de tamanho variável, adicionadas/removidas livremente pelo usuário (ex.: endereços adicionais de execução, infrações administrativas extras) | Implementado |
| RF-12 | Calcular e exibir o valor de uma quantia monetária por extenso automaticamente a partir do valor numérico informado | Implementado |
| RF-13 | Exibir aviso automático ao usuário quando um valor preenchido ultrapassa um limiar legal relevante (ex.: exigência do Programa de Integridade acima de R$ 5.000.000,00) | Implementado |

## 3. Requisitos funcionais — Pré-visualização

| ID | Requisito | Status |
|---|---|---|
| RF-14 | Exibir, em painel lateral, o texto real do documento sendo elaborado, atualizado a cada alteração de resposta (debounce de ~400ms) | Implementado |
| RF-15 | Destacar visualmente, na pré-visualização, o trecho correspondente ao campo em edição, e rolar a visualização até esse trecho automaticamente | Implementado |
| RF-16 | Exibir tabelas do documento (ex.: Matriz de Riscos, IMR) corretamente formatadas na pré-visualização, não apenas em texto corrido | Implementado |

## 4. Requisitos funcionais — Fidelidade normativa

| ID | Requisito | Status |
|---|---|---|
| RF-17 | Todo texto de cláusula fixa do documento deve ser transcrito ou derivado diretamente do RLC 10.901 ou de uma minuta/anexo oficial da Conab — nenhum conteúdo normativo criado por suposição | Implementado, verificado por revisão manual e por múltiplas rodadas de auditoria (ver histórico de commits) |
| RF-18 | Todo anexo técnico obrigatório (Matriz de Riscos, IMR, Planilha de Custos, Declaração de Contratos Firmados, Contrato Simplificado, Mapa Comparativo de Propostas, Termo de Cooperação Técnica) deve seguir a estrutura (colunas, cabeçalhos, escalas) do anexo oficial correspondente do RLC | Implementado, verificado por comparação estrutural direta com os arquivos `.odt` oficiais |

## 5. Requisitos não funcionais

| ID | Requisito | Status |
|---|---|---|
| RNF-01 | O sistema deve funcionar inteiramente no navegador, sem servidor de aplicação nem banco de dados | Implementado |
| RNF-02 | Nenhum dado preenchido pelo usuário deve ser transmitido a servidores externos | Implementado (não há chamadas de rede além do carregamento dos arquivos estáticos do próprio site) |
| RNF-03 | O visual do sistema deve seguir o padrão gov.br Design System | Implementado |
| RNF-04 | Formulários devem seguir boas práticas de acessibilidade (eMag), incluindo associação programática entre rótulo e campo | Implementado (auditoria de acessibilidade realizada; ver commit "Associate form labels with their inputs") |
| RNF-05 | O sistema deve possuir suíte de testes automatizados cobrindo a geração de cada instrumento e os mecanismos condicionais do motor de formulário | Implementado — 116 testes automatizados (`node --test`) |

## 6. Fora de escopo (decisão registrada)

- Integração com SEI, SICAF ou qualquer sistema corporativo da Conab.
- Autenticação de usuários ou controle de quem gerou qual documento.
- Geração de PDF nativa (o formato de saída é exclusivamente `.docx`).
- Atualização automática de conteúdo normativo quando o RLC é alterado — cada
  atualização precisa ser aplicada manualmente ao código-fonte.

## 7. Rastreabilidade

As decisões de design e o histórico de correções de conteúdo (achados de fidelidade
normativa, correções de citação, restauração de cláusulas comprimidas) estão
registrados em `docs/superpowers/specs/` e `docs/superpowers/plans/`, e no histórico
de commits do Git deste repositório.
