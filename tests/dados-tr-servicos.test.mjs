import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";
import { controladoresBlocoVariante } from "../js/modelo.js";

test("SECOES_TR_SERVICOS tem 21 seções com ids únicos de 01 a 21", () => {
  assert.equal(SECOES_TR_SERVICOS.length, 21);
  const ids = SECOES_TR_SERVICOS.map((s) => s.id);
  assert.deepEqual(ids, [
    "01","02","03","04","05","06","07","08","09","10",
    "11","12","13","14","15","16","17","18","19","20","21",
  ]);
});

test("todo campo com blocoTexto tem padrao dentro de opcoes", () => {
  for (const campo of controladoresBlocoVariante(SECOES_TR_SERVICOS)) {
    assert.ok(campo.opcoes.includes(campo.padrao),
      `${campo.id}: padrao "${campo.padrao}" não está em opcoes`);
  }
});

test("todo campo tabela-linhas declara colunas", () => {
  for (const secao of SECOES_TR_SERVICOS) {
    for (const campo of secao.campos) {
      if (campo.input === "tabela-linhas") {
        assert.ok(Array.isArray(campo.colunas) && campo.colunas.length > 0, campo.id);
      }
    }
  }
});

test("ids de campo são únicos entre todas as seções", () => {
  const todosIds = SECOES_TR_SERVICOS.flatMap((s) => s.campos.map((c) => c.id));
  assert.equal(todosIds.length, new Set(todosIds).size);
});

test("todo campo é tipo parametro ou especificacao", () => {
  for (const secao of SECOES_TR_SERVICOS) {
    for (const campo of secao.campos) {
      assert.ok(campo.tipo === "parametro" || campo.tipo === "especificacao", campo.id);
    }
  }
});
