// tests/modelo.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import { campoVisivel, progresso, pendencias, pendenciasParametro, controladoresCondicionais, camposOcultos, controladoresBlocoVariante } from "../js/modelo.js";

test("campo sem mostrarSe é sempre visível", () => {
  const campo = { id: "x", tipo: "parametro", pergunta: "?" };
  assert.equal(campoVisivel(campo, {}), true);
});

test("campo com mostrarSe só aparece quando a condição bate", () => {
  const campo = { id: "y", tipo: "parametro", pergunta: "?", mostrarSe: { permite: "Sim" } };
  assert.equal(campoVisivel(campo, { permite: "Não" }), false);
  assert.equal(campoVisivel(campo, { permite: "Sim" }), true);
});

test("progresso conta só campos de especificação visíveis", () => {
  const secoes = [{
    id: "01", titulo: "T", campos: [
      { id: "a", tipo: "parametro", pergunta: "?" },
      { id: "b", tipo: "especificacao", pergunta: "?" },
      { id: "c", tipo: "especificacao", pergunta: "?", mostrarSe: { a: "X" } },
    ],
  }];
  const r1 = progresso(secoes, { a: "outro", b: "preenchido" });
  assert.equal(r1.totalEspecificacao, 1); // "c" não visível, não conta
  assert.equal(r1.preenchidos, 1);

  const r2 = progresso(secoes, { a: "X", b: "", c: "preenchido" });
  assert.equal(r2.totalEspecificacao, 2); // "c" agora visível
  assert.equal(r2.preenchidos, 1); // "b" vazio não conta
});

test("pendencias lista especificações visíveis e vazias, com a pergunta", () => {
  const secoes = [{
    id: "01", titulo: "T", campos: [
      { id: "a", tipo: "especificacao", pergunta: "Descreva A" },
    ],
  }];
  const r = pendencias(secoes, { a: "" });
  assert.deepEqual(r, [{ secaoId: "01", campoId: "a", pergunta: "Descreva A" }]);
  assert.deepEqual(pendencias(secoes, { a: "ok" }), []);
});

test("pendenciasParametro lista só parâmetros visíveis e vazios", () => {
  const secoes = [{
    id: "07", titulo: "T", campos: [
      { id: "esp", tipo: "especificacao", pergunta: "Descreva" },
      { id: "ctrl", tipo: "parametro", pergunta: "Exige?" },
      { id: "dep", tipo: "parametro", pergunta: "Percentual?", mostrarSe: { ctrl: "Sim" } },
    ],
  }];
  // "dep" está oculto (ctrl = Não), então não é pendência; "esp" é especificação.
  assert.deepEqual(pendenciasParametro(secoes, { ctrl: "Não" }), []);
  assert.deepEqual(pendenciasParametro(secoes, { ctrl: "Sim" }),
    [{ secaoId: "07", campoId: "dep", pergunta: "Percentual?" }]);
  assert.deepEqual(pendenciasParametro(secoes, {}),
    [{ secaoId: "07", campoId: "ctrl", pergunta: "Exige?" }]);
});

test("controladoresCondicionais deriva os campos que controlam visibilidade", () => {
  const secoes = [
    { id: "05", titulo: "T", campos: [
      { id: "c1", tipo: "parametro", pergunta: "?" },
      { id: "d1", tipo: "parametro", pergunta: "?", mostrarSe: { c1: "Sim" } },
      { id: "d2", tipo: "parametro", pergunta: "?", mostrarSe: { c1: "Sim" } },
    ]},
    { id: "07", titulo: "T", campos: [
      { id: "d3", tipo: "parametro", pergunta: "?", mostrarSe: { c2: "Sim" } },
    ]},
  ];
  assert.deepEqual(controladoresCondicionais(secoes), ["c1", "c2"]);
});

test("controladoresBlocoVariante devolve os campos com blocoTexto", () => {
  const secoes = [{
    id: "05", titulo: "T", campos: [
      { id: "c1", tipo: "parametro", input: "select", opcoes: ["A", "B"], pergunta: "?" },
      { id: "c2", tipo: "parametro", input: "select", opcoes: ["Sim", "Não"],
        blocoTexto: true, padrao: "Não", pergunta: "?" },
    ],
  }];
  const resultado = controladoresBlocoVariante(secoes);
  assert.equal(resultado.length, 1);
  assert.equal(resultado[0].id, "c2");
});

test("camposOcultos devolve os campos escondidos pela resposta atual", () => {
  const secoes = [{
    id: "05", titulo: "T", campos: [
      { id: "c1", tipo: "parametro", pergunta: "?" },
      { id: "d1", tipo: "parametro", pergunta: "?", mostrarSe: { c1: "Sim" } },
    ],
  }];
  assert.deepEqual(camposOcultos(secoes, { c1: "Sim" }).map(c => c.id), []);
  assert.deepEqual(camposOcultos(secoes, { c1: "Não" }).map(c => c.id), ["d1"]);
  assert.deepEqual(camposOcultos(secoes, {}).map(c => c.id), ["d1"]);
});
