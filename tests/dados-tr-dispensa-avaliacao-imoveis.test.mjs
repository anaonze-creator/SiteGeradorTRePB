// tests/dados-tr-dispensa-avaliacao-imoveis.test.mjs
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS } from "../js/dados-tr-dispensa-avaliacao-imoveis.js";
import { controladoresBlocoVariante } from "../js/modelo.js";

globalThis.JSZip = JSZip;

test("SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS tem 21 seções com ids únicos de 01 a 21", () => {
  assert.equal(SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS.length, 21);
  const ids = SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS.map((s) => s.id);
  const esperado = Array.from({ length: 21 }, (_, i) => String(i + 1).padStart(2, "0"));
  assert.deepEqual(ids, esperado);
});

test("seções fixas (Garantia Contratual e Subcontratação) não têm campos condicionais", () => {
  const s08 = SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS.find((s) => s.id === "08");
  const s14 = SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS.find((s) => s.id === "14");
  assert.deepEqual(s08.campos, []);
  assert.deepEqual(s14.campos, []);
});

test("Seção 20 (Matriz de Riscos) não tem campo de ativação — a matriz sempre se aplica", () => {
  const s20 = SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS.find((s) => s.id === "20");
  assert.deepEqual(s20.campos, []);
});

test("todo campo com blocoTexto tem padrao dentro de opcoes", () => {
  for (const campo of controladoresBlocoVariante(SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS)) {
    assert.ok(campo.opcoes.includes(campo.padrao),
      `campo ${campo.id} tem padrao "${campo.padrao}" fora de opcoes ${JSON.stringify(campo.opcoes)}`);
  }
});

test("ids de campo são únicos entre todas as seções", () => {
  const ids = [];
  for (const secao of SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS) {
    for (const campo of secao.campos) ids.push(campo.id);
  }
  assert.equal(new Set(ids).size, ids.length);
});

function lerTemplate() {
  const buffer = fs.readFileSync("templates/Template_TR_Dispensa_Avaliacao_Imoveis.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXml(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate(), respostas, SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function textoDoDocumento(documentXml) {
  return documentXml
    .replace(/<w:p\b[^>]*>/g, "\n")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, "$1")
    .replace(/<[^>]+>/g, "");
}

test("gerarDocx produz um .docx válido com a checklist do laudo e a matriz de riscos pré-preenchida", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S01_LocalizacaoImovel: "Rua das Palmeiras, 123, Brasília-DF",
    S01_ValorEstimado: "50.000,00",
    S05_AreaMinimaAvaliadaM2: "500",
  }));
  assert.ok(texto.includes("Rua das Palmeiras, 123, Brasília-DF"));
  assert.ok(texto.includes("Laudo de Avaliação Imobiliária"));
  assert.ok(texto.includes("ABNT NBR 14653-1"));
  assert.ok(texto.includes("NOC 60.208"));
  assert.ok(texto.includes("500") && texto.includes("m²"));
  assert.ok(texto.includes("Impossibilidade de vistoria interna do imóvel"));
  assert.ok(texto.includes("Não será exigida garantia contratual"));
  assert.ok(texto.includes("Não será admitida a subcontratação do objeto"));
});

test("Anexo II reaproveita a mesma checklist da Seção 1", async () => {
  const texto = await gerarXml({});
  const ocorrencias = (textoDoDocumento(texto).match(/ABNT NBR 14653-1/g) || []).length;
  assert.ok(ocorrencias >= 4, `esperado a citação repetida entre Seção 1 e Anexo II, achou ${ocorrencias}`);
});

test("não há Contrato Simplificado neste instrumento: sempre usa a Minuta de Contrato completa", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S_AnexoContrato_NumeroProcesso: "21447.001437/2026-39",
  }));
  assert.ok(!texto.includes("21. DA ALTERAÇÃO DO CONTRATO"));
  assert.ok(!texto.includes("MINUTA DE CONTRATO SIMPLIFICADO"));
  assert.ok((texto.match(/21447\.001437\/2026-39/g) || []).length >= 2);
  assert.ok(texto.includes("CLÁUSULA VIGÉSIMA SÉTIMA - DO FORO"));
  assert.ok(texto.includes("MINUTA DE CONTRATO"));
});

test("qualificação da Contratada na Minuta de Contrato é branco fixo, não campo preenchível", async () => {
  const texto = textoDoDocumento(await gerarXml({}));
  assert.ok(texto.includes("CONTRATADA: (razão social), (CNPJ, endereço e representante legal com nome e CPF)"));
  assert.ok(texto.includes("CONTRATANTE: (unidade da Conab"));
  const ids = [];
  for (const secao of SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS) {
    for (const campo of secao.campos) ids.push(campo.id);
  }
  assert.ok(!ids.includes("S_AnexoContrato_QualificacaoConab"));
  assert.ok(!ids.includes("S_AnexoContrato_NomeContratada"));
  assert.ok(!ids.includes("S_AnexoContrato_QualificacaoContratada"));
});

test("vigência: prazo sempre em meses, prorrogação Sim exige vantajosidade (art. 488) no corpo e no Anexo IV", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S04_PrazoVigenciaMeses: "12",
    S04_HaveraProrrogacao: "Sim",
    S04_PrazoProrrogacaoMeses: "até mais 12 meses",
  }));
  assert.ok((texto.match(/12 meses/g) || []).length >= 1);
  assert.ok((texto.match(/vantajosidade na manutenção do ajuste/g) || []).length >= 2);
  assert.ok((texto.match(/art\. 488 do RLC/g) || []).length >= 2);
});

test("vigência: prorrogação Não remove a exigência de vantajosidade e informa que não haverá prorrogação", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S04_PrazoVigenciaMeses: "6",
    S04_HaveraProrrogacao: "Não",
  }));
  assert.ok((texto.match(/Não haverá prorrogação do prazo de vigência do Contrato\./g) || []).length >= 2);
  assert.ok(!texto.includes("vantajosidade na manutenção do ajuste"));
});
