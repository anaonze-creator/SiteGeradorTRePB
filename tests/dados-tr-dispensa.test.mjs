import { test } from "node:test";
import assert from "node:assert/strict";
import { SECOES_TR_DISPENSA } from "../js/dados-tr-dispensa.js";
import { controladoresBlocoVariante } from "../js/modelo.js";

test("SECOES_TR_DISPENSA tem 26 seções com ids únicos de 01 a 26", () => {
  assert.equal(SECOES_TR_DISPENSA.length, 26);
  const ids = SECOES_TR_DISPENSA.map((s) => s.id);
  const esperado = Array.from({ length: 26 }, (_, i) => String(i + 1).padStart(2, "0"));
  assert.deepEqual(ids, esperado);
});

test("todo campo com blocoTexto tem padrao dentro de opcoes", () => {
  for (const campo of controladoresBlocoVariante(SECOES_TR_DISPENSA)) {
    assert.ok(campo.opcoes.includes(campo.padrao),
      `${campo.id}: padrao "${campo.padrao}" não está em opcoes`);
  }
});

test("o campo de instrumento contratual declara os dois blocosAdicionais", () => {
  const campo = SECOES_TR_DISPENSA
    .flatMap((s) => s.campos)
    .find((c) => c.id === "S07_TipoInstrumentoContratual");
  assert.ok(campo, "S07_TipoInstrumentoContratual não encontrado");
  assert.deepEqual(campo.blocosAdicionais, [
    "S07_TipoInstrumentoContratual_SECOES21A25",
    "S07_TipoInstrumentoContratual_ANEXO",
  ]);
});

test("o campo de matriz de riscos declara o blocoAdicional do Anexo I", () => {
  const campo = SECOES_TR_DISPENSA
    .flatMap((s) => s.campos)
    .find((c) => c.id === "S20_ExigeMatrizRiscos");
  assert.ok(campo);
  assert.deepEqual(campo.blocosAdicionais, ["S20_ExigeMatrizRiscos_ANEXO_I"]);
});

test("o campo de base de cálculo de multa declara duplicarPara", () => {
  const campo = SECOES_TR_DISPENSA
    .flatMap((s) => s.campos)
    .find((c) => c.id === "S16_BaseCalculoMulta");
  assert.ok(campo);
  assert.deepEqual(campo.duplicarPara, [
    "S16_BaseCalculoMulta_E", "S16_BaseCalculoMulta_F",
    "S16_BaseCalculoMulta_G", "S16_BaseCalculoMulta_TABELA2",
    "S16_BaseCalculoMulta_ANEXO_VI_E", "S16_BaseCalculoMulta_ANEXO_VI_F",
    "S16_BaseCalculoMulta_ANEXO_VI_G", "S16_BaseCalculoMulta_ANEXO_VI_TABELA2",
  ]);
});

test("ids de campo são únicos entre todas as seções", () => {
  const todosIds = SECOES_TR_DISPENSA.flatMap((s) => s.campos.map((c) => c.id));
  assert.equal(todosIds.length, new Set(todosIds).size);
});

test("todo campo tabela-linhas declara colunas", () => {
  for (const secao of SECOES_TR_DISPENSA) {
    for (const campo of secao.campos) {
      if (campo.input === "tabela-linhas") {
        assert.ok(Array.isArray(campo.colunas) && campo.colunas.length > 0, campo.id);
      }
    }
  }
});

test("todo campo é tipo parametro ou especificacao", () => {
  for (const secao of SECOES_TR_DISPENSA) {
    for (const campo of secao.campos) {
      assert.ok(campo.tipo === "parametro" || campo.tipo === "especificacao", campo.id);
    }
  }
});
