import { pendencias, pendenciasParametro } from "./modelo.js";

export function montarMensagemConferencia(listaPendencias, listaPendenciasParametro = []) {
  const partes = [];
  if (listaPendencias.length === 0) {
    // Só promete "pronto para gerar" quando também não há parâmetro vazio —
    // parâmetro vazio deixa placeholder cru ("____") no documento.
    partes.push(listaPendenciasParametro.length === 0
      ? "Nenhum campo de especificação pendente. O documento está pronto para gerar."
      : "Nenhum campo de especificação pendente.");
  } else {
    const linhas = listaPendencias.map(p => `- Seção ${p.secaoId}: ${p.pergunta}`);
    // Deixar explícito que a lista é um AVISO, não um bloqueio — o botão
    // "Gerar Documento" logo abaixo continua funcionando mesmo com itens
    // pendentes (o documento sai com o placeholder "____" nesses trechos,
    // para o elaborador completar depois).
    partes.push(
      `${listaPendencias.length} campo(s) de especificação ainda não preenchidos — isto é só um aviso, ` +
      `você ainda pode clicar em "Gerar Documento" abaixo, o placeholder "____" fica no lugar para completar depois:\n\n${linhas.join("\n")}`
    );
  }

  if (listaPendenciasParametro.length > 0) {
    const linhas = listaPendenciasParametro.map(p => `- Seção ${p.secaoId}: ${p.pergunta}`);
    partes.push(
      `Parâmetros ainda não preenchidos (mesmo aviso — não impede a geração) — ${listaPendenciasParametro.length} campo(s):\n\n${linhas.join("\n")}`
    );
  }
  return partes.join("\n\n");
}

export function renderizarConferencia(secoes, respostas, overlay, aoConfirmarGeracao) {
  const lista = pendencias(secoes, respostas);
  const listaParametro = pendenciasParametro(secoes, respostas);
  overlay.innerHTML = `
    <div class="br-modal-body">
      <pre>${montarMensagemConferencia(lista, listaParametro)}</pre>
      <button id="btn-confirmar-geracao" class="br-button primary" type="button">Gerar Documento</button>
      <button id="btn-fechar-conferencia" class="br-button secondary" type="button">Voltar e Preencher</button>
    </div>
  `;
  overlay.hidden = false;
  overlay.querySelector("#btn-confirmar-geracao").addEventListener("click", aoConfirmarGeracao);
  overlay.querySelector("#btn-fechar-conferencia").addEventListener("click", () => { overlay.hidden = true; });
}
