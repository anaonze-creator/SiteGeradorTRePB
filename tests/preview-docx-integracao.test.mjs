import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { renderizarPreviewHtml } from "../js/preview-docx.js";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";
import { SECOES_TR_DISPENSA } from "../js/dados-tr-dispensa.js";
import { SECOES_PB } from "../js/dados-pb.js";

globalThis.JSZip = JSZip;

function lerTemplate(caminho) {
  const buffer = fs.readFileSync(caminho);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarPreviewHtml(caminhoTemplate, respostas, secoes) {
  const docxBuffer = await gerarDocx(lerTemplate(caminhoTemplate), respostas, secoes);
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  return renderizarPreviewHtml(documentXml);
}

test("preview TR-Bens: resposta preenchida aparece marcada", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_TR_Bens.docx",
    { S01_DescricaoObjeto: "Aquisição de computadores portáteis" },
    SECOES_TR_BENS
  );
  assert.match(html, /<mark class="preview-destaque"[^>]*>[^<]*Aquisição de computadores portáteis[^<]*<\/mark>/);
  assert.match(html, /data-campo="[^"]*S01_DescricaoObjeto[^"]*"/);
});

test("preview TR-Serviços: resposta preenchida aparece marcada", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_TR_Servicos.docx",
    { S01_DescricaoObjeto: "Prestação de serviços de limpeza" },
    SECOES_TR_SERVICOS
  );
  assert.match(html, /<mark class="preview-destaque"[^>]*>[^<]*Prestação de serviços de limpeza[^<]*<\/mark>/);
});

test("preview TR-Dispensa: resposta preenchida aparece marcada", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_TR_Dispensa.docx",
    { S01_DescricaoObjeto: "Contratação direta de manutenção predial" },
    SECOES_TR_DISPENSA
  );
  assert.match(html, /<mark class="preview-destaque"[^>]*>[^<]*Contratação direta de manutenção predial[^<]*<\/mark>/);
});

test("preview PB: resposta preenchida aparece marcada, sem tags de bookmark cruas", async () => {
  const html = await gerarPreviewHtml(
    "templates/Template_PB.docx",
    { PB01_DescricaoObjeto: "Construção de galpão de armazenagem" },
    SECOES_PB
  );
  assert.match(html, /<mark class="preview-destaque"[^>]*>[^<]*Construção de galpão de armazenagem[^<]*<\/mark>/);
  assert.ok(!html.includes("w:bookmarkStart"), "HTML não deve conter tags OOXML cruas");
  assert.ok(!html.includes("_BLOCO_"), "nomes de bookmark estruturais não devem vazar para o HTML");
});
