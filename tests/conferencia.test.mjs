import { test } from "node:test";
import assert from "node:assert/strict";
import { montarMensagemConferencia } from "../js/conferencia.js";

test("sem pendências, mensagem confirma que está tudo preenchido", () => {
  const msg = montarMensagemConferencia([]);
  assert.match(msg, /Nenhum campo de especificação pendente/);
});

test("com pendências, lista cada uma por seção e pergunta", () => {
  const msg = montarMensagemConferencia([
    { secaoId: "01", campoId: "S01_DescricaoObjeto", pergunta: "Descreva o material" },
    { secaoId: "18", campoId: "S18_CriteriosEspecificos", pergunta: "Critérios específicos?" },
  ]);
  assert.match(msg, /2 campo\(s\) de especificação ainda não preenchidos/);
  assert.match(msg, /Seção 01: Descreva o material/);
  assert.match(msg, /Seção 18: Critérios específicos\?/);
});

test("parâmetros vazios são listados à parte e retiram a promessa de 'pronto para gerar'", () => {
  const msg = montarMensagemConferencia([], [
    { secaoId: "15", campoId: "S15_MultaRescisoria", pergunta: "Multa rescisória (%)" },
  ]);
  assert.match(msg, /Nenhum campo de especificação pendente\./);
  assert.doesNotMatch(msg, /pronto para gerar/);
  assert.match(msg, /Parâmetros ainda não preenchidos \(mesmo aviso — não impede a geração\)/);
  assert.match(msg, /Seção 15: Multa rescisória \(%\)/);
});

test("sem especificação e sem parâmetro pendente, a mensagem confirma que está pronto", () => {
  assert.match(montarMensagemConferencia([], []), /pronto para gerar/);
});

test("mensagem de pendência deixa claro que a geração não é bloqueada", () => {
  const msg = montarMensagemConferencia([
    { secaoId: "01", campoId: "S01_X", pergunta: "teste" },
  ]);
  assert.match(msg, /Gerar Documento/);
  assert.doesNotMatch(msg, /impedid|bloquead/);
});
