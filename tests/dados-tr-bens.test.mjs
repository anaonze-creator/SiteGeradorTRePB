// tests/dados-tr-bens.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";

test("existem exatamente 21 seções, na ordem 01 a 21", () => {
  assert.equal(SECOES_TR_BENS.length, 21);
  assert.deepEqual(SECOES_TR_BENS.map(s => s.id),
    Array.from({ length: 21 }, (_, i) => String(i + 1).padStart(2, "0")));
});

test("seção 13 (Subcontratação) tem o campo condicional esperado", () => {
  const s13 = SECOES_TR_BENS.find(s => s.id === "13");
  const permite = s13.campos.find(c => c.id === "S13_PermiteSubcontratacao");
  assert.equal(permite.tipo, "parametro");
  assert.deepEqual(permite.opcoes, ["Sim", "Não"]);
  const percentualMin = s13.campos.find(c => c.id === "S13_PercentualMinSubcontratacao");
  assert.deepEqual(percentualMin.mostrarSe, { S13_PermiteSubcontratacao: "Sim" });
});

test("seção 16 (Sanções) tem os 7 percentuais de multa mais o limite de dias, 8 campos ao todo", () => {
  const s16 = SECOES_TR_BENS.find(s => s.id === "16");
  assert.equal(s16.campos.length, 8);
  assert.ok(s16.campos.every(c => c.tipo === "parametro"));
});

test("seções sem campo (texto fixo) têm array de campos vazio", () => {
  for (const id of ["08", "09", "14", "15", "17", "18", "20"]) {
    const secao = SECOES_TR_BENS.find(s => s.id === id);
    assert.deepEqual(secao.campos, []);
  }
});

test("todo campo de especificação tem pergunta não vazia", () => {
  for (const secao of SECOES_TR_BENS) {
    for (const campo of secao.campos) {
      assert.ok(campo.pergunta && campo.pergunta.trim().length > 0,
        `campo ${campo.id} sem pergunta`);
    }
  }
});
