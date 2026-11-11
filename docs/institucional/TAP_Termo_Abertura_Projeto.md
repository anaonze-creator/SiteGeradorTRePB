# TERMO DE ABERTURA DE PROJETO (TAP)

> Preenchido conforme o modelo oficial do Anexo III da Norma 60.114. Campos marcados
> com `[PREENCHER]` dependem de decisão institucional (equipe, aprovadores) que só a
> área requisitante/Sutin pode definir.

**1. Superintendência:** `[PREENCHER — Superintendência da área demandante]`

**2. Área Requisitante:** `[PREENCHER]`

**3. Código/Sigla:** Gerador TR/PB (sugestão — sujeito a padronização da Sutin)

**4. Nome do Projeto:** Gerador de Termo de Referência e Projeto Básico da Conab

**5. Resumo da Demanda (RDN – Objetivos Esperados):**
Ver RDN (`RDN_Registro_Necessidade_TI.md`) — item 9. Em síntese: padronizar e agilizar
a elaboração de TR e PB, com fidelidade ao RLC 10.901 e às minutas oficiais da Conab.

**6. Premissas:**
- O sistema é 100% estático e roda inteiramente no navegador (client-side), sem
  servidor de aplicação, sem banco de dados e sem envio de dados a terceiros — a
  geração do `.docx` acontece localmente, no computador de quem preenche.
- Todo conteúdo textual gerado é transcrito ou derivado diretamente do RLC 10.901 ou
  das minutas/anexos oficiais da Conab — nenhum conteúdo normativo é criado por
  suposição.
- O sistema reaproveita a técnica de manipulação de bookmarks OOXML já validada no
  projeto correlato `Gerador_PB.html` (Projeto-Basico-App).

**7. Restrições:**
- Não há integração com SEI, SICAF ou qualquer sistema corporativo da Conab — o
  usuário preenche o formulário e baixa o `.docx` gerado manualmente.
- O sistema depende de o Microsoft Word (ou compatível) para abrir o `.docx` final;
  não há geração de PDF nativa.
- Alterações no RLC ou nas minutas oficiais exigem atualização manual do código-fonte
  (schemas de campos e templates) — não há mecanismo de atualização automática de
  conteúdo normativo.

**8. Demandante (RDN – Interlocutor Responsável):**

| Nome | Contato |
|---|---|
| `[PREENCHER]` | `[PREENCHER]` |

**9. Equipe Inicial:**

| Nome | Papel no Projeto | Contato | Cargo/Função |
|---|---|---|---|
| `[PREENCHER]` | Área Demandante / Product Owner | `[PREENCHER]` | `[PREENCHER]` |
| Claude Code (Anthropic) | Ferramenta de desenvolvimento assistido por IA, usada para implementação, testes e documentação do código | — | — |

**10. Alinhamento ao Planejamento Estratégico:**
`[PREENCHER — mesmo campo do RDN, item 10]`

**11. Aprovação (assinatura):**

| Gestor Imediato Área Executora | Gestor Imediato Área Demandante | Superintendente/Diretor Área Demandante (patrocinador) |
|---|---|---|
| Nome: `[PREENCHER]` | Nome: `[PREENCHER]` | Nome: `[PREENCHER]` |
| Cargo/Função: | Cargo/Função: | Cargo/Função: |
| Data: | Data: | Data: |
