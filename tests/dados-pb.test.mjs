import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_PB } from "../js/dados-pb.js";
import { controladoresBlocoVariante } from "../js/modelo.js";

test("SECOES_PB tem 35 seções com ids únicos de 01 a 35", () => {
  assert.equal(SECOES_PB.length, 35);
  const ids = SECOES_PB.map((s) => s.id);
  const esperado = Array.from({ length: 35 }, (_, i) => String(i + 1).padStart(2, "0"));
  assert.deepEqual(ids, esperado);
});

test("todo campo com blocoTexto tem padrao dentro de opcoes", () => {
  for (const campo of controladoresBlocoVariante(SECOES_PB)) {
    assert.ok(campo.opcoes.includes(campo.padrao),
      `${campo.id}: padrao "${campo.padrao}" não está em opcoes`);
  }
});

test("ids de campo são únicos entre todas as seções", () => {
  const todosIds = SECOES_PB.flatMap((s) => s.campos.map((c) => c.id));
  assert.equal(todosIds.length, new Set(todosIds).size);
});

test("todo campo é tipo parametro ou especificacao", () => {
  for (const secao of SECOES_PB) {
    for (const campo of secao.campos) {
      assert.ok(campo.tipo === "parametro" || campo.tipo === "especificacao", campo.id);
    }
  }
});

test("PB01_TipoObjeto e PB26_ExigeQuadroInovacoes declaram blocosAdicionais", () => {
  const tipoObjeto = SECOES_PB.flatMap((s) => s.campos).find((c) => c.id === "PB01_TipoObjeto");
  assert.ok(tipoObjeto);
  const quadroInovacoes = SECOES_PB.flatMap((s) => s.campos).find((c) => c.id === "PB26_ExigeQuadroInovacoes");
  assert.deepEqual(quadroInovacoes.blocosAdicionais, ["PB26_ExigeQuadroInovacoes_ANEXO_IIA"]);
});
