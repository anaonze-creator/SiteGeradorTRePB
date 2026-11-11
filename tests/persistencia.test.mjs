import { test } from "node:test";
import assert from "node:assert/strict";
import { serializarRascunho, desserializarRascunho } from "../js/persistencia.js";

test("serializar e desserializar preserva instrumento e respostas", () => {
  const json = serializarRascunho("TR-Bens", { S01_ValorEstimado: "1000" });
  const de_volta = desserializarRascunho(json);
  assert.equal(de_volta.instrumento, "TR-Bens");
  assert.deepEqual(de_volta.respostas, { S01_ValorEstimado: "1000" });
});

test("desserializar rejeita JSON sem instrumento ou respostas", () => {
  assert.throws(() => desserializarRascunho(JSON.stringify({ respostas: {} })), /instrumento/);
  assert.throws(() => desserializarRascunho(JSON.stringify({ instrumento: "TR-Bens" })), /respostas/);
});

test("desserializar rejeita JSON malformado", () => {
  assert.throws(() => desserializarRascunho("{não é json"));
});
