import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";
import { SECOES_PB } from "../js/dados-pb.js";

globalThis.JSZip = JSZip;

function lerTemplate(nomeArquivo) {
  const buffer = fs.readFileSync(`templates/${nomeArquivo}`);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXml(nomeArquivo, respostas, secoes) {
  const resultadoBuffer = await gerarDocx(lerTemplate(nomeArquivo), respostas, secoes);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function textoDoDocumento(xml) {
  return xml.replace(/<[^>]+>/g, "");
}

test("TR-Bens: SRP Não (padrão) mantém critério de julgamento editável, sem nota de SRP/TIC", async () => {
  const texto = textoDoDocumento(await gerarXml("Template_TR_Bens.docx", {
    S01_CriterioJulgamento: "Maior desconto",
  }, SECOES_TR_BENS));
  assert.ok(texto.includes("Maior desconto"));
  assert.ok(!texto.includes("por se tratar de licitação para Sistema de Registro de Preços"));
  assert.ok(!texto.includes("PDTIC"));
});

test("TR-Bens: SRP Sim força menor preço, dispensa dotação, exige justificativa e quantidade mínima", async () => {
  const texto = textoDoDocumento(await gerarXml("Template_TR_Bens.docx", {
    S01_SRP: "Sim",
    S01_JustificativaSRP: "hipótese do inciso II do art. 327",
    S05_QuantidadeMinimaCotadaPorItem: "10 unidades por item",
  }, SECOES_TR_BENS));
  assert.ok(texto.includes("menor preço ofertado, por se tratar de licitação para Sistema de Registro de Preços"));
  assert.ok(!texto.includes("[menor preço ou maior desconto]"));
  assert.ok(texto.includes("não é necessária a indicação de dotação orçamentária"));
  assert.ok(texto.includes("hipótese do inciso II do art. 327"));
  assert.ok(texto.includes("10 unidades por item"));
});

test("TR-Bens: campo de critério de julgamento some quando SRP é Sim", async () => {
  const { campoVisivel } = await import("../js/modelo.js");
  const respostas = { S01_SRP: "Sim" };
  const campo = SECOES_TR_BENS.flatMap((s) => s.campos).find((c) => c.id === "S01_CriterioJulgamento");
  assert.equal(campoVisivel(campo, respostas), false);
});

test("TR-Bens: TIC Sim mostra a nota do PDTIC, independente do SRP", async () => {
  const texto = textoDoDocumento(await gerarXml("Template_TR_Bens.docx", {
    S01_SRP: "Não", S01_TIC: "Sim",
  }, SECOES_TR_BENS));
  assert.ok(texto.includes("PDTIC"));
  assert.ok(!texto.includes("por se tratar de licitação para Sistema de Registro de Preços"));
});

test("TR-Serviços: SRP Sim força menor preço e exige justificativa (sem campo de quantidade mínima)", async () => {
  const texto = textoDoDocumento(await gerarXml("Template_TR_Servicos.docx", {
    S01_SRP: "Sim",
    S01_JustificativaSRP: "hipótese do inciso I do art. 327",
  }, SECOES_TR_SERVICOS));
  assert.ok(texto.includes("menor preço ofertado, por se tratar de licitação para Sistema de Registro de Preços"));
  assert.ok(texto.includes("hipótese do inciso I do art. 327"));
  assert.ok(!texto.includes("Quantidade mínima de unidades a ser cotada"));
});

test("TR-Serviços: TIC Sim mostra a nota do PDTIC", async () => {
  const texto = textoDoDocumento(await gerarXml("Template_TR_Servicos.docx", {
    S01_TIC: "Sim",
  }, SECOES_TR_SERVICOS));
  assert.ok(texto.includes("PDTIC"));
});

test("PB: TIC Sim mostra a nota do PDTIC (PB não tem campo de SRP)", async () => {
  const texto = textoDoDocumento(await gerarXml("Template_PB.docx", {
    PB01_TIC: "Sim",
  }, SECOES_PB));
  assert.ok(texto.includes("PDTIC"));
  const campoSRP = SECOES_PB.flatMap((s) => s.campos).find((c) => c.id === "PB01_SRP");
  assert.equal(campoSRP, undefined);
});
