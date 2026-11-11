import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { gerarDocx } from "../js/gerador-docx.js";
import { SECOES_PB } from "../js/dados-pb.js";

globalThis.JSZip = JSZip;

function lerTemplate() {
  const buffer = fs.readFileSync("templates/Template_PB.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlPB(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate(), respostas, SECOES_PB);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function textoDoDocumento(xml) {
  return xml.replace(/<[^>]+>/g, "");
}

test("PB: regime Licitação (padrão) não altera o comportamento existente", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Licitação",
    PB02_ModalidadeLicitatoria: "Pregão Eletrônico",
  }));
  assert.ok(texto.includes("O serviço em apreço será contratado na modalidade"));
  assert.ok(texto.includes("Pregão Eletrônico"));
  assert.ok(texto.includes("Modalidade licitatória:"));
  assert.ok(texto.includes("A Conab poderá revogar a licitação"));
  assert.ok(!texto.includes("MINUTA DE CONTRATO"));
});

test("PB: regime Dispensa troca modalidade, vistoria e revogação, e adiciona o anexo de contrato", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB01_DescricaoObjeto: "reforma da SUREG/RO",
    PB08_PrazoVigencia: "12 meses",
    PB22_PercentualGarantia: "5",
  }));
  assert.ok(texto.includes("A contratação será realizada por dispensa de licitação, com fundamento no art. 416, I, do RLC"));
  assert.ok(texto.includes("Fundamento da dispensa: art. 416, I, do RLC"));
  assert.ok(texto.includes("O prazo para a vistoria será definido no processo de dispensa de licitação"));
  assert.ok(texto.includes("A Conab poderá anular o processo de dispensa por ilegalidade"));
  assert.ok(!texto.includes("A Conab poderá revogar a licitação"));
  assert.ok(texto.includes("MINUTA DE CONTRATO"));
  assert.ok(texto.includes("reforma da SUREG/RO"));
  assert.ok(texto.includes("CLÁUSULA VIGÉSIMA QUARTA - DA PROTEÇÃO DE DADOS PESSOAIS"));
  assert.ok(texto.includes("Lei nº 13.709, de 2018"));
});

test("PB: campo de modalidade licitatória some quando o regime é Dispensa", async () => {
  const respostas = { PB01_RegimeContratacao: "Dispensa de Licitação", PB02_ModalidadeLicitatoria: "Pregão Eletrônico" };
  const { campoVisivel } = await import("../js/modelo.js");
  const campoModalidade = SECOES_PB.flatMap((s) => s.campos).find((c) => c.id === "PB02_ModalidadeLicitatoria");
  assert.equal(campoVisivel(campoModalidade, respostas), false);
});

test("PB: regime Licitação não traz os controles de dispensa (somatório/aviso/cartão)", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Licitação",
  }));
  assert.ok(!texto.includes("foi observado o limite de valor"));
  assert.ok(!texto.includes("Instrução Normativa SEGES/ME nº 67"));
  assert.ok(!texto.includes("cartão de pagamento"));
});

test("PB: regime Dispensa traz somatório, aviso de 3 dias úteis e pagamento por cartão (Sim)", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB02_SomatorioContratacoesMesmaNatureza: "não há outras contratações de mesma natureza no exercício",
    PB02_PagamentoCartao: "Sim",
  }));
  assert.ok(texto.includes("não há outras contratações de mesma natureza no exercício"));
  assert.ok(texto.includes("Instrução Normativa SEGES/ME nº 67, de 2021"));
  assert.ok(texto.includes("paga por meio de cartão de pagamento, com divulgação do respectivo extrato no Portal Nacional de Contratações Públicas (PNCP)"));
});

test("PB: regime Dispensa com pagamento por cartão Não exige justificativa", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB02_PagamentoCartao: "Não",
    PB02_JustificativaNaoCartao: "fornecedor não aceita cartão de pagamento",
  }));
  assert.ok(texto.includes("não será paga por meio de cartão de pagamento"));
  assert.ok(texto.includes("fornecedor não aceita cartão de pagamento"));
});

test("PB: Seção 20 (Reajuste) com \"Sim\" (padrão) traz o texto de reajuste/repactuação", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB20_IndiceReajuste: "INCC-DI/FGV",
  }));
  assert.ok(texto.includes("os preços serão reajustados"));
  assert.ok(texto.includes("INCC-DI/FGV"));
  assert.ok(!texto.includes("Não haverá reajuste ou repactuação de preços"));
});

test("PB: Seção 20 (Reajuste) com \"Não\" remove o texto de reajuste sem deixar placeholder cru", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB20_TipoCorrecao: "Não",
  }));
  assert.ok(texto.includes("Não haverá reajuste ou repactuação de preços neste Contrato."));
  assert.ok(!texto.includes("os preços serão reajustados"));
  assert.ok(!texto.includes("índice de reajuste será o"));
});

test("PB: Seção 20 (Reajuste) \"Não\" também troca a Cláusula Décima Segunda da Minuta de Contrato", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB20_TipoCorrecao: "Não",
  }));
  assert.ok(texto.includes("CLÁUSULA DÉCIMA SEGUNDA - DO REAJUSTAMENTO / REPACTUAÇÃO"));
  const ocorrencias = (texto.match(/Não haverá reajuste ou repactuação de preços neste Contrato\./g) || []).length;
  assert.ok(ocorrencias >= 2, `esperado no corpo e na Minuta de Contrato, achou ${ocorrencias}`);
});
