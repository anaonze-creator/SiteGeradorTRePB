import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { renderizarPreviewHtml } from "../js/preview-docx.js";
import { SECOES_NOTA_DEMANDA } from "../js/dados-nota-demanda.js";
import { SECOES_NOTA_TECNICA } from "../js/dados-nota-tecnica.js";

globalThis.JSZip = JSZip;

function lerTemplate(caminho) {
  const buffer = fs.readFileSync(caminho);
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlNotaDemanda(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate("templates/Template_Nota_Demanda.docx"), respostas, SECOES_NOTA_DEMANDA);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

async function gerarXmlNotaTecnica(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate("templates/Template_Nota_Tecnica.docx"), respostas, SECOES_NOTA_TECNICA);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function textoDoDocumento(xml) {
  return xml.replace(/<[^>]+>/g, "");
}

test("Nota de Demanda: geração básica preenche a necessidade", async () => {
  const xml = await gerarXmlNotaDemanda({ ND01_Necessidade: "Reforma da unidade de Rondônia" });
  const texto = textoDoDocumento(xml);
  assert.ok(texto.includes("Reforma da unidade de Rondônia"));
  assert.ok(texto.includes("NOTA DE DEMANDA"));
});

test("Nota de Demanda: parcelamento Sim/Não troca a redação da Seção 8", async () => {
  const textoSim = textoDoDocumento(await gerarXmlNotaDemanda({ ND08_Parcelamento: "Sim" }));
  assert.ok(textoSim.includes("A solução será parcelada, sendo a licitação realizada por item."));
  const textoNao = textoDoDocumento(await gerarXmlNotaDemanda({ ND08_Parcelamento: "Não" }));
  assert.ok(textoNao.includes("A solução não será parcelada, sendo a licitação realizada em lote único."));
});

test("Nota de Demanda: adequação do ambiente Sim mostra detalhe; Não remove", async () => {
  const textoSim = textoDoDocumento(await gerarXmlNotaDemanda({
    ND10_AdequacaoAmbiente: "Sim",
    ND10_DetalheAdequacao: "Necessário treinar a equipe de fiscalização",
  }));
  assert.ok(textoSim.includes("Necessário treinar a equipe de fiscalização"));
  const textoNao = textoDoDocumento(await gerarXmlNotaDemanda({
    ND10_AdequacaoAmbiente: "Não",
    ND10_DetalheAdequacao: "Não deveria aparecer",
  }));
  assert.ok(!textoNao.includes("Não deveria aparecer"));
  assert.ok(textoNao.includes("Não há necessidade de adequação do ambiente da Conab"));
});

test("Nota Técnica: geração básica preenche a necessidade e usa o cabeçalho e assinatura próprios", async () => {
  const xml = await gerarXmlNotaTecnica({
    NT01_Necessidade: "Reforma técnica da SUREG/RO",
    NT19_Responsaveis: "Analista Engenheiro Civil A\nAnalista Engenheiro Civil B",
  });
  const texto = textoDoDocumento(xml);
  assert.ok(texto.includes("Reforma técnica da SUREG/RO"));
  assert.ok(texto.includes("NOTA TÉCNICA"));
  assert.ok(texto.includes("Responsável(is) pelo estudo — Área Técnica"));
  assert.ok(texto.includes("Analista Engenheiro Civil A"));
  assert.ok(texto.includes("Analista Engenheiro Civil B"));
});

test("Nota Técnica x Nota de Demanda: a pergunta da Seção 1 tem redação diferente entre os dois instrumentos", () => {
  const perguntaND = SECOES_NOTA_DEMANDA[0].campos[0].pergunta;
  const perguntaNT = SECOES_NOTA_TECNICA[0].campos[0].pergunta;
  assert.notEqual(perguntaND, perguntaNT);
  assert.ok(perguntaNT.toLowerCase().includes("tecnicamente"));
  assert.ok(!perguntaND.toLowerCase().includes("tecnicamente"));
});

test("Nota Técnica: obra/serviço de engenharia Sim mantém as 4 seções extras; Não remove", async () => {
  const textoSim = textoDoDocumento(await gerarXmlNotaTecnica({
    NT_ObraServicoEngenharia: "Sim",
    NT15_AdequacaoTecnica: "Solução compatível com as normas ABNT aplicáveis",
    NT16_Funcionalidade: "Atende à necessidade de reforma predial",
    NT17_RequisitosAmbientais: "Reaproveitamento de materiais sempre que viável",
    NT18_AdequacaoNormasVigentes: "Conforme NBR 5410 e NBR 5419",
  }));
  assert.ok(textoSim.includes("ADEQUAÇÃO TÉCNICA (ART. 123, V, DO RLC)"));
  assert.ok(textoSim.includes("Solução compatível com as normas ABNT aplicáveis"));
  assert.ok(textoSim.includes("Atende à necessidade de reforma predial"));
  assert.ok(textoSim.includes("Reaproveitamento de materiais sempre que viável"));
  assert.ok(textoSim.includes("Conforme NBR 5410 e NBR 5419"));

  const textoNao = textoDoDocumento(await gerarXmlNotaTecnica({ NT_ObraServicoEngenharia: "Não" }));
  assert.ok(!textoNao.includes("ADEQUAÇÃO TÉCNICA (ART. 123, V, DO RLC)"));
  assert.ok(!textoNao.includes("FUNCIONALIDADE (ART. 123, VI, DO RLC)"));
  assert.ok(!textoNao.includes("REQUISITOS AMBIENTAIS (ART. 123, VII, DO RLC)"));
  assert.ok(!textoNao.includes("ADEQUAÇÃO ÀS NORMAS VIGENTES (ART. 123, VIII, DO RLC)"));
});

test("Nota Técnica: preview ao vivo marca o texto preenchido, sem tags cruas", async () => {
  const docxBuffer = await gerarDocx(
    lerTemplate("templates/Template_Nota_Tecnica.docx"),
    { NT01_Necessidade: "Reforma elétrica da SUREG/RO" },
    SECOES_NOTA_TECNICA
  );
  const zip = await JSZip.loadAsync(docxBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const html = renderizarPreviewHtml(documentXml);
  assert.match(html, /<mark class="preview-destaque"[^>]*>[^<]*Reforma elétrica da SUREG\/RO[^<]*<\/mark>/);
  assert.ok(!html.includes("w:bookmarkStart"));
  assert.ok(!html.includes("_BLOCO_"));
});
