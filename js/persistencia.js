export function serializarRascunho(instrumento, respostas) {
  return JSON.stringify({ instrumento, respostas, salvoEm: new Date().toISOString() }, null, 2);
}

export function desserializarRascunho(json) {
  let dado;
  try {
    dado = JSON.parse(json);
  } catch {
    throw new Error("Arquivo de rascunho inválido: não é um JSON válido.");
  }
  if (!dado.instrumento) throw new Error("Arquivo de rascunho inválido: falta o campo 'instrumento'.");
  if (!dado.respostas) throw new Error("Arquivo de rascunho inválido: falta o campo 'respostas'.");
  return { instrumento: dado.instrumento, respostas: dado.respostas };
}

export function salvarAutosave(chave, instrumento, respostas) {
  localStorage.setItem(chave, serializarRascunho(instrumento, respostas));
}

export function carregarAutosave(chave) {
  const json = localStorage.getItem(chave);
  if (!json) return null;
  return desserializarRascunho(json);
}
