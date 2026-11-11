function valorEfetivo(campoId, respostas, secoes) {
  if (respostas[campoId] !== undefined) return respostas[campoId];
  if (!secoes) return undefined;
  for (const secao of secoes) {
    const controlador = secao.campos.find((c) => c.id === campoId);
    if (controlador) return controlador.padrao;
  }
  return undefined;
}

// `secoes` é opcional, mas necessário para que um controlador Sim/Não com
// `padrao` (ex.: S01_SRP, padrao "Não") seja tratado corretamente antes de o
// usuário tocar no campo — sem ele, um campo `mostrarSe: {S01_SRP: "Não"}`
// ficaria oculto (e seu valor nunca escrito no documento) enquanto o
// controlador não fosse respondido explicitamente.
export function campoVisivel(campo, respostas, secoes) {
  if (!campo.mostrarSe) return true;
  const [campoId, valorEsperado] = Object.entries(campo.mostrarSe)[0];
  return valorEfetivo(campoId, respostas, secoes) === valorEsperado;
}

export function progresso(secoes, respostas) {
  let totalEspecificacao = 0;
  let preenchidos = 0;
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.tipo !== "especificacao") continue;
      if (!campoVisivel(campo, respostas, secoes)) continue;
      totalEspecificacao++;
      const valor = respostas[campo.id];
      if (valor && String(valor).trim() !== "") preenchidos++;
    }
  }
  return { totalEspecificacao, preenchidos };
}

// Fonte única de verdade sobre QUAIS campos controlam a visibilidade de
// outros (os "condicionais Sim/Não"). Derivado de `mostrarSe`, para que um
// 6º condicional adicionado aos dados seja reconhecido tanto pelo formulário
// (js/app.js) quanto pelo gerador de .docx (js/gerador-docx.js) sem edição.
export function controladoresCondicionais(secoes) {
  const ids = new Set();
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.mostrarSe) ids.add(Object.keys(campo.mostrarSe)[0]);
    }
  }
  return [...ids];
}

// Campos que controlam um bloco de texto alternativo inteiro no documento
// (Sim/Não ou N variantes nomeadas) — diferente de `mostrarSe`, que só
// esconde/mostra campos individuais no formulário. Um campo pode fazer as
// duas coisas ao mesmo tempo (ex.: controlar visibilidade de um campo E
// trocar um parágrafo inteiro do documento).
export function controladoresBlocoVariante(secoes) {
  const resultado = [];
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.blocoTexto) resultado.push(campo);
    }
  }
  return resultado;
}

// Campos que estão escondidos no formulário por causa da resposta atual do
// seu controlador. O gerador usa esta lista para apagar do documento os
// parágrafos que só fazem sentido quando o controlador é "Sim".
export function camposOcultos(secoes, respostas) {
  const resultado = [];
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.mostrarSe && !campoVisivel(campo, respostas, secoes)) resultado.push(campo);
    }
  }
  return resultado;
}

export function pendencias(secoes, respostas) {
  const resultado = [];
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.tipo !== "especificacao") continue;
      if (!campoVisivel(campo, respostas, secoes)) continue;
      const valor = respostas[campo.id];
      if (!valor || String(valor).trim() === "") {
        resultado.push({ secaoId: secao.id, campoId: campo.id, pergunta: campo.pergunta });
      }
    }
  }
  return resultado;
}

// Espelha `pendencias`, mas para os campos de parâmetro. Não entra no cálculo
// de progresso (que, por decisão de projeto, só mede especificação) — serve
// apenas de aviso no painel de conferência, já que parâmetro não preenchido
// deixa placeholder cru ("____") no documento gerado.
export function pendenciasParametro(secoes, respostas) {
  const resultado = [];
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.tipo !== "parametro") continue;
      if (!campoVisivel(campo, respostas, secoes)) continue;
      const valor = respostas[campo.id];
      if (!valor || String(valor).trim() === "") {
        resultado.push({ secaoId: secao.id, campoId: campo.id, pergunta: campo.pergunta });
      }
    }
  }
  return resultado;
}
