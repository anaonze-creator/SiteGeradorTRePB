import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import JSZip from "../js/lib/jszip.min.js";
import { substituirBookmark, removerBlocoBookmark, gerarDocx, inserirLinhasTabela } from "../js/gerador-docx.js";
import { SECOES_TR_BENS } from "../js/dados-tr-bens.js";
import { SECOES_TR_SERVICOS } from "../js/dados-tr-servicos.js";
import { SECOES_TR_DISPENSA } from "../js/dados-tr-dispensa.js";
import { SECOES_PB } from "../js/dados-pb.js";

// No navegador, js/lib/jszip.min.js é carregado via <script> clássico e cria
// o global `window.JSZip`, que js/gerador-docx.js consome como `JSZip`
// (identificador global, sem import). No Node, o import ESM acima só cria um
// binding local — não popula um global — então replicamos aqui o mesmo
// efeito que o <script> tag produz no navegador.
globalThis.JSZip = JSZip;

function lerTemplate() {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXml(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplate(), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function lerTemplateServicos() {
  const buffer = fs.readFileSync("templates/Template_TR_Servicos.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlServicos(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplateServicos(), respostas, SECOES_TR_SERVICOS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function lerTemplateDispensa() {
  const buffer = fs.readFileSync("templates/Template_TR_Dispensa.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlDispensa(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplateDispensa(), respostas, SECOES_TR_DISPENSA);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

function lerTemplatePB() {
  const buffer = fs.readFileSync("templates/Template_PB.docx");
  return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

async function gerarXmlPB(respostas) {
  const resultadoBuffer = await gerarDocx(lerTemplatePB(), respostas, SECOES_PB);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  return zip.file("word/document.xml").async("string");
}

// Texto corrido do documento (concatena os <w:t>), para conferir que cada
// valor caiu no lugar CERTO da frase — só procurar o valor solto no XML não
// pegaria a troca entre bookmarks de um mesmo parágrafo.
function textoDoDocumento(documentXml) {
  return documentXml
    .replace(/<w:p\b[^>]*>/g, "\n")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<w:t[^>]*>([^<]*)<\/w:t>/g, "$1")
    .replace(/<[^>]+>/g, "");
}

test("substituirBookmark troca o texto dentro do bookmark, preservando o resto do XML", () => {
  const xmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S04_PrazoVigenciaMeses"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>';
  const resultado = substituirBookmark(xmlFalso, "S04_PrazoVigenciaMeses", "36");
  assert.match(resultado, /<w:t[^>]*>36<\/w:t>/);
  assert.doesNotMatch(resultado, /____/);
});

test("removerBlocoBookmark remove o parágrafo inteiro que contém o bookmark", () => {
  const xmlFalso =
    '<w:p><w:r><w:t>antes</w:t></w:r></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S07_ExigeGarantia_BLOCO_NAO"/>' +
    '<w:r><w:t>Não será exigida garantia contratual.</w:t></w:r>' +
    '<w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:r><w:t>depois</w:t></w:r></w:p>';
  const resultado = removerBlocoBookmark(xmlFalso, "S07_ExigeGarantia_BLOCO_NAO");
  assert.match(resultado, /antes/);
  assert.match(resultado, /depois/);
  assert.doesNotMatch(resultado, /Não será exigida garantia contratual/);
});

test("removerBlocoBookmark remove um bloco que abrange MAIS de um parágrafo", () => {
  const xmlFalso =
    '<w:p><w:r><w:t>antes</w:t></w:r></w:p>' +
    '<w:p><w:bookmarkStart w:id="9" w:name="S08_TipoFiscalizacao_BLOCO_COMISSAO_DE_FISCALIZACAO"/>' +
    '<w:r><w:t>parágrafo 1 do bloco</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>parágrafo 2 do bloco</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>parágrafo 3 do bloco</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p>' +
    '<w:p><w:r><w:t>depois</w:t></w:r></w:p>';
  const resultado = removerBlocoBookmark(xmlFalso, "S08_TipoFiscalizacao_BLOCO_COMISSAO_DE_FISCALIZACAO");
  assert.match(resultado, /antes/);
  assert.match(resultado, /depois/);
  assert.doesNotMatch(resultado, /parágrafo 1 do bloco/);
  assert.doesNotMatch(resultado, /parágrafo 2 do bloco/);
  assert.doesNotMatch(resultado, /parágrafo 3 do bloco/);
});

test("inserirLinhasTabela clona a linha-modelo uma vez por item e remove o modelo", () => {
  const campoFalso = {
    id: "S15_InfracoesExtras",
    colunas: [{ id: "descricao" }, { id: "grau" }],
  };
  const xmlFalso =
    '<w:tbl>' +
    '<w:tr><w:tc><w:p><w:r><w:t>fixa</w:t></w:r></w:p></w:tc></w:tr>' +
    '<w:tr><w:tc><w:p><w:bookmarkStart w:id="7" w:name="S15_InfracoesExtras_LINHA_MODELO"/>' +
    '<w:r><w:t>((DESCRICAO))</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>((GRAU))</w:t></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl>';
  const linhas = [{ descricao: "Deixar de treinar equipe", grau: "1" }, { descricao: "Atraso reiterado", grau: "3" }];
  const resultado = inserirLinhasTabela(xmlFalso, campoFalso, linhas);
  assert.match(resultado, /fixa/);
  assert.match(resultado, /Deixar de treinar equipe/);
  assert.match(resultado, /Atraso reiterado/);
  assert.doesNotMatch(resultado, /\(\(DESCRICAO\)\)/);
  assert.doesNotMatch(resultado, /\(\(GRAU\)\)/);
  assert.doesNotMatch(resultado, /LINHA_MODELO/);
});

test("inserirLinhasTabela sem linhas extras remove a linha-modelo e não deixa placeholder", () => {
  const campoFalso = { id: "S15_InfracoesExtras", colunas: [{ id: "descricao" }, { id: "grau" }] };
  const xmlFalso =
    '<w:tr><w:tc><w:p><w:bookmarkStart w:id="1" w:name="S15_InfracoesExtras_LINHA_MODELO"/>' +
    '<w:r><w:t>((DESCRICAO))</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p></w:tc></w:tr>';
  const resultado = inserirLinhasTabela(xmlFalso, campoFalso, []);
  assert.doesNotMatch(resultado, /\(\(DESCRICAO\)\)/);
});

test("gerarDocx produz um .docx válido com os valores substituídos", async () => {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  const respostas = {
    S01_DescricaoObjeto: "notebooks corporativos",
    S01_ValorEstimado: "150.000,00",
    S04_PrazoVigenciaMeses: "24",
    S07_ExigeGarantia: "Não",
    S13_PermiteSubcontratacao: "Não",
  };
  const resultadoBuffer = await gerarDocx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  assert.match(documentXml, /notebooks corporativos/);
  assert.match(documentXml, /150\.000,00/);
  assert.match(documentXml, /<w:t[^>]*>24<\/w:t>/);
  // Como S07_ExigeGarantia é "Não", o bloco negativo (BLOCO_NAO) permanece...
  assert.match(documentXml, /Não será exigida garantia contratual/);
  // ...e o bloco positivo (BLOCO_SIM) precisa ter sido removido, senão os
  // dois parágrafos contraditórios ficariam juntos no documento final.
  assert.doesNotMatch(documentXml, /A Contratada deverá apresentar, no prazo de 10/);
  // Mesma checagem no lado inverso: S13 respondeu "Não", então o texto de
  // "é permitida a subcontratação" (BLOCO_SIM) não pode sobreviver.
  assert.doesNotMatch(documentXml, /É permitida a subcontratação parcial do objeto/);
  assert.match(documentXml, /Não será admitida a subcontratação do objeto licitatório/);
});

test("gerarDocx remove o bloco negativo quando a resposta é Sim", async () => {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  const respostas = { S07_ExigeGarantia: "Sim", S07_PercentualGarantia: "3" };
  const resultadoBuffer = await gerarDocx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  assert.match(documentXml, /A Contratada deverá apresentar, no prazo de 10/);
  assert.doesNotMatch(documentXml, /Não será exigida garantia contratual/);
});

test("o campo Objeto do Anexo I é preenchido automaticamente a partir da Seção 1", async () => {
  const buffer = fs.readFileSync("templates/Template_TR_Bens.docx");
  const respostas = { S01_DescricaoObjeto: "cadeiras de escritório" };
  const resultadoBuffer = await gerarDocx(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), respostas, SECOES_TR_BENS);
  const zip = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zip.file("word/document.xml").async("string");
  const ocorrencias = documentXml.split("cadeiras de escritório").length - 1;
  assert.ok(ocorrencias >= 2, "esperado o valor tanto na Seção 1 quanto no Anexo I");
});

test("cada campo cai no seu próprio bookmark em parágrafos com vários (1.4: ND/PTRES/Fonte)", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S01_ND: "339030", S01_PTRES: "170175", S01_Fonte: "0100000000",
  }));
  assert.ok(texto.includes("ND: (339030)"), "ND fora do lugar: " + texto.slice(texto.indexOf("1.4."), texto.indexOf("1.4.") + 260));
  assert.ok(texto.includes("PTRES: (170175)"), "PTRES fora do lugar");
  assert.ok(texto.includes("FONTE: (0100000000)"), "Fonte fora do lugar");
  assert.ok(!texto.includes("XXXX"), "sobrou placeholder não substituído no parágrafo 1.4");
});

test("cada campo cai no seu próprio bookmark no parágrafo 6.1.1", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S06_FormaFornecimento: "Parcelada",
    S06_LocalEntrega: "Armazém de Brasília/DF",
    S06_HorarioEntrega: "8h às 17h",
  }));
  assert.ok(texto.includes("entregues de forma Parcelada"), "forma de fornecimento fora do lugar");
  assert.ok(texto.includes("lacrados, em Armazém de Brasília/DF"), "local de entrega fora do lugar");
  assert.ok(texto.includes("no horário de 8h às 17h"), "horário de entrega fora do lugar");
});

test("quebras de linha da textarea viram <w:br/> no documento", async () => {
  const documentXml = await gerarXml({ S01_DescricaoObjeto: "primeira linha\nsegunda linha" });
  assert.match(documentXml, /<w:br\/>/);
  assert.match(documentXml, /<w:t[^>]*>primeira linha<\/w:t>/);
  assert.match(documentXml, /<w:t[^>]*>segunda linha<\/w:t>/);
});

test("mudar S07 de Sim para Não apaga a resposta antiga e o parágrafo dependente", async () => {
  const respostas = { S07_ExigeGarantia: "Sim", S07_PercentualGarantia: "3" };
  const textoSim = textoDoDocumento(await gerarXml(respostas));
  assert.ok(textoSim.includes("Percentual da garantia: 3%"), "com Sim, o percentual deve aparecer");

  respostas.S07_ExigeGarantia = "Não";
  const textoNao = textoDoDocumento(await gerarXml(respostas));
  assert.ok(!textoNao.includes("Percentual da garantia"),
    "parágrafo dependente sobreviveu à mudança para Não");
  assert.ok(!textoNao.includes("3%"), "resposta obsoleta vazou para o documento");
  assert.ok(textoNao.includes("Não será exigida garantia contratual"));
});

test("mudar S13 de Sim para Não apaga limites e detalhamento da subcontratação", async () => {
  const respostas = {
    S13_PermiteSubcontratacao: "Sim",
    S13_PercentualMinSubcontratacao: "10",
    S13_PercentualMaxSubcontratacao: "30",
    S13_DetalheSubcontratacao: "Apenas o serviço de instalação.",
  };
  const textoSim = textoDoDocumento(await gerarXml(respostas));
  assert.ok(textoSim.includes("Limites: entre 10% e 30%"));
  assert.ok(textoSim.includes("Detalhamento: Apenas o serviço de instalação."));

  respostas.S13_PermiteSubcontratacao = "Não";
  const textoNao = textoDoDocumento(await gerarXml(respostas));
  assert.ok(!textoNao.includes("Limites: entre"), "parágrafo de limites sobreviveu");
  assert.ok(!textoNao.includes("Detalhamento:"), "parágrafo de detalhamento sobreviveu");
  assert.ok(!textoNao.includes("Apenas o serviço de instalação"), "resposta obsoleta vazou");
  assert.ok(textoNao.includes("Não será admitida a subcontratação"));
});

test("mudar S05 de Sim para Não apaga prazo, local e horário da amostra", async () => {
  const respostas = {
    S05_ExigeAmostra: "Sim",
    S05_PrazoAmostraDias: "5",
    S05_LocalAmostra: "Gerência de Logística",
    S05_HorarioAmostra: "9h às 12h",
  };
  const textoSim = textoDoDocumento(await gerarXml(respostas));
  assert.ok(textoSim.includes("Prazo para apresentação: 5 dias úteis."));
  assert.ok(textoSim.includes("apresentada em Gerência de Logística, no horário de 9h às 12h"));

  respostas.S05_ExigeAmostra = "Não";
  const textoNao = textoDoDocumento(await gerarXml(respostas));
  assert.ok(!textoNao.includes("Prazo para apresentação"), "parágrafo 5.2.1.1 sobreviveu");
  assert.ok(!textoNao.includes("A amostra deverá ser apresentada"), "parágrafo 5.2.2.1 sobreviveu");
  assert.ok(!textoNao.includes("Gerência de Logística"), "resposta obsoleta vazou");
  assert.ok(textoNao.includes("Não será exigida amostra dos bens"));
});

test("com S12 e S02 em Não, os parágrafos dependentes não deixam rótulo órfão", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S12_Reajustavel: "Não", S12_IndiceReajuste: "IPCA",
    S02_AplicaExcecaoMEEPP: "Não", S02_MotivoExcecaoMEEPP: "motivo antigo",
  }));
  assert.ok(!texto.includes("Índice de referência"), "rótulo do índice sobreviveu com Não");
  assert.ok(!texto.includes("IPCA"), "resposta obsoleta do índice vazou");
  assert.ok(!texto.includes("motivo antigo"), "resposta obsoleta da exceção ME/EPP vazou");
  assert.ok(texto.includes("Os preços são fixos e irreajustáveis"));
});

test("TR-Serviços: vistoria obrigatória remove a declaração de quem não visitou", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S05_TipoVistoria: "Obrigatória" }));
  assert.ok(texto.includes("o licitante deverá realizar vistoria"));
  assert.ok(!texto.includes("O licitante que optar por não realizar a vistoria"));
});

test("TR-Serviços: vistoria facultativa mantém os dois lados (quem visita e quem não visita)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S05_TipoVistoria: "Facultativa" }));
  assert.ok(texto.includes("recomenda-se que o licitante realize vistoria"));
  assert.ok(texto.includes("O licitante que optar por não realizar a vistoria"));
});

test("TR-Serviços: fiscalização por Fiscal Funcional usa as alíneas b) e c) (Seção 11)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S08_TipoFiscalizacao: "Fiscal Funcional" }));
  assert.ok(texto.includes("b) no prazo de até 07 (sete) dias corridos a partir do recebimento dos documentos da Contratada, o Fiscal Funcional deverá realizar a análise de toda a documentação apresentada;"));
  assert.ok(texto.includes("c) no prazo de até 03 (três) dias úteis a partir da análise da documentação, o Fiscal Funcional avaliará a execução por meio do Instrumento de Medição de Resultado (IMR)"));
  assert.ok(!texto.includes("b.1)"));
  assert.ok(!texto.includes("b.2)"));
  assert.ok(!texto.includes("b.3)"));
});

test("TR-Serviços: Comissão de Fiscalização usa as alíneas b, b.1, b.2, b.3 e c (Seção 11)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S08_TipoFiscalizacao: "Comissão de Fiscalização" }));
  assert.ok(texto.includes("b.1) o Fiscal Funcional deverá realizar a análise"));
  assert.ok(texto.includes("b.2) o Fiscal Técnico deverá realizar a análise"));
  assert.ok(texto.includes("b.3) o Fiscal Setorial deverá realizar a análise"));
  assert.ok(!texto.includes("o Fiscal Funcional deverá realizar a análise de toda a documentação apresentada, avaliar a execução"));
});

test("TR-Serviços: natureza intelectual inclui o bloco de cessão de direitos autorais", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S10_NaturezaIntelectual: "Sim" }));
  assert.ok(texto.includes("ceder os direitos patrimoniais"));
  assert.ok(texto.includes("transição contratual com transferência de conhecimento"));
});

test("TR-Serviços: sem natureza intelectual, o bloco de cessão de direitos não aparece", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S10_NaturezaIntelectual: "Não" }));
  assert.ok(!texto.includes("ceder os direitos patrimoniais"));
});

test("TR-Serviços: infrações extras da Tabela 1 aparecem, as 9 fixas continuam presentes", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({
    S15_InfracoesExtras: [{ descricao: "Deixar de portar crachá de identificação", grau: "1" }],
  }));
  assert.ok(texto.includes("Deixar de portar crachá de identificação"));
  assert.ok(texto.includes("Permitir situação que crie a possibilidade de causar dano físico"));
});

test("TR-Serviços: sem infrações extras, a Tabela 1 tem só as 9 fixas", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({}));
  const ocorrenciasLinhaExtra = texto.split("((DESCRICAO))").length - 1;
  assert.equal(ocorrenciasLinhaExtra, 0);
});

test("TR-Serviços: base de cálculo da multa preenche os quatro pontos duplicados (e, f, g, Tabela 2)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S15_BaseCalculoMulta: "valor anual do Contrato" }));
  const ocorrencias = texto.split("valor anual do Contrato").length - 1;
  assert.ok(ocorrencias >= 4, `esperado valor em pelo menos 4 lugares, achou ${ocorrencias}`);
});

test("TR-Serviços: vigência plurianual usa a redação de até 5 anos (art. 488)", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({ S04_TipoVigencia: "Plurianual (até 5 anos)" }));
  assert.ok(texto.includes("podendo ser prorrogado até o limite de 05 (cinco) anos"));
  assert.ok(!texto.includes("prorrogável na forma dos arts. 497 e 498"));
});

test("gerarDocx remove blocosAdicionais junto com o bloco principal do campo", async () => {
  const secoesFalsas = [{
    id: "07", titulo: "T", campos: [
      { id: "S07_Tipo", tipo: "parametro", input: "select",
        opcoes: ["A", "B"], blocoTexto: true, padrao: "A",
        blocosAdicionais: ["S07_Tipo_EXTRA"] },
    ],
  }];
  const documentXmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S07_Tipo_BLOCO_B"/>' +
    '<w:r><w:t>texto principal B</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S07_Tipo_EXTRA_BLOCO_B"/>' +
    '<w:r><w:t>texto extra B</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:r><w:t>fica</w:t></w:r></w:p>';
  const zip = new JSZip();
  zip.file("word/document.xml", documentXmlFalso);
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  const resultadoBuffer = await gerarDocx(buffer, {}, secoesFalsas);
  const zipResultado = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zipResultado.file("word/document.xml").async("string");
  assert.doesNotMatch(documentXml, /texto principal B/);
  assert.doesNotMatch(documentXml, /texto extra B/);
  assert.match(documentXml, /fica/);
});

test("gerarDocx preenche duplicarPara com o mesmo valor do campo", async () => {
  const secoesFalsas = [{
    id: "16", titulo: "T", campos: [
      { id: "S16_BaseCalculoMulta", tipo: "parametro", input: "select",
        opcoes: ["valor do Contrato", "valor anual do Contrato"],
        pergunta: "?", duplicarPara: ["S16_BaseCalculoMulta_E", "S16_BaseCalculoMulta_F"] },
    ],
  }];
  const documentXmlFalso =
    '<w:p><w:bookmarkStart w:id="1" w:name="S16_BaseCalculoMulta"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="2" w:name="S16_BaseCalculoMulta_E"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
    '<w:p><w:bookmarkStart w:id="3" w:name="S16_BaseCalculoMulta_F"/>' +
    '<w:r><w:t>____</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>';
  const zip = new JSZip();
  zip.file("word/document.xml", documentXmlFalso);
  const buffer = await zip.generateAsync({ type: "arraybuffer" });
  const resultadoBuffer = await gerarDocx(buffer, { S16_BaseCalculoMulta: "valor anual do Contrato" }, secoesFalsas);
  const zipResultado = await JSZip.loadAsync(resultadoBuffer);
  const documentXml = await zipResultado.file("word/document.xml").async("string");
  const ocorrencias = documentXml.split("valor anual do Contrato").length - 1;
  assert.equal(ocorrencias, 3, "esperado o valor no campo principal + 2 pontos de duplicarPara");
});

test("TR-Dispensa: Contrato Simplificado mantém Seções 21-25 e o Anexo V, remove o Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S07_TipoInstrumentoContratual: "Contrato Simplificado" }));
  assert.ok(texto.includes("21. DA ALTERAÇÃO DO CONTRATO"));
  assert.ok(texto.includes("25. DO FORO"));
  assert.ok(texto.includes("MINUTA DE CONTRATO SIMPLIFICADO"));
  assert.ok(!texto.includes("CLÁUSULA PRIMEIRA - DO OBJETO"));
});

test("TR-Dispensa: Contrato completo remove as Seções 21-25 e o Anexo V, mantém o Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S07_TipoInstrumentoContratual: "Contrato" }));
  assert.ok(!texto.includes("21. DA ALTERAÇÃO DO CONTRATO"));
  assert.ok(!texto.includes("25. DO FORO"));
  assert.ok(!texto.includes("MINUTA DE CONTRATO SIMPLIFICADO"));
  assert.ok(texto.includes("CLÁUSULA PRIMEIRA - DO OBJETO"));
  assert.ok(texto.includes("CLÁUSULA VIGÉSIMA QUINTA - DA PROTEÇÃO DE DADOS PESSOAIS"));
  assert.ok(texto.includes("Lei nº 13.709, de 2018"));
  assert.ok(texto.includes("CLÁUSULA VIGÉSIMA SÉTIMA - DO FORO"));
});

test("TR-Dispensa: exclusividade ME/EPP no padrão (Sim) usa o texto de exclusividade", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S02_ExcecaoExclusividadeMEEPP: "Sim" }));
  assert.ok(texto.includes("destinado exclusivamente à participação de"));
  assert.ok(!texto.includes("não será destinada exclusivamente"));
});

test("TR-Dispensa: exceção à exclusividade ME/EPP (Não) abre a disputa", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S02_ExcecaoExclusividadeMEEPP: "Não" }));
  assert.ok(texto.includes("não será destinada exclusivamente"));
  assert.ok(!texto.includes("destinado exclusivamente à participação de"));
});

test("TR-Dispensa: Matriz de Riscos aplicável mantém o Anexo I", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S20_ExigeMatrizRiscos: "Sim" }));
  assert.ok(texto.includes("MATRIZ DE RISCOS"));
  assert.ok(texto.includes("Escala de Probabilidade"));
});

test("TR-Dispensa: Matriz de Riscos dispensada remove o Anexo I", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S20_ExigeMatrizRiscos: "Não" }));
  assert.ok(texto.includes("serviços de pronto pagamento"));
  assert.ok(!texto.includes("Escala de Probabilidade"));
});

test("TR-Dispensa: fiscalização por Comissão usa as alíneas b.1, b.2, b.3 na Seção 12", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({ S09_TipoFiscalizacao: "Comissão de Fiscalização" }));
  assert.ok(texto.includes("o Fiscal Técnico deverá realizar a análise"));
  assert.ok(texto.includes("o Fiscal Setorial deverá realizar a análise"));
});

test("TR-Dispensa: infrações extras da Tabela 1 aparecem junto às 9 fixas", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({
    S16_InfracoesExtras: [{ descricao: "Deixar de portar crachá de identificação", grau: "1" }],
  }));
  assert.ok(texto.includes("Deixar de portar crachá de identificação"));
  assert.ok(texto.includes("Permitir situação que crie a possibilidade de causar dano físico"));
});

test("inserirLinhasTabela processa TODAS as ocorrências do bookmark _LINHA_MODELO, não só a primeira", () => {
  const campoFalso = { id: "S16_InfracoesExtras", colunas: [{ id: "descricao" }, { id: "grau" }] };
  const linhaModelo =
    '<w:tr><w:tc><w:p><w:bookmarkStart w:id="ID" w:name="S16_InfracoesExtras_LINHA_MODELO"/>' +
    '<w:r><w:t>((DESCRICAO))</w:t></w:r><w:bookmarkEnd w:id="ID"/></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>((GRAU))</w:t></w:r></w:p></w:tc></w:tr>';
  const xmlFalso =
    '<w:tbl>' + linhaModelo.replace(/ID/g, "1") + '</w:tbl>' +
    '<w:tbl>' + linhaModelo.replace(/ID/g, "2") + '</w:tbl>';
  const linhas = [{ descricao: "Infração extra", grau: "2" }];
  const resultado = inserirLinhasTabela(xmlFalso, campoFalso, linhas);
  const ocorrencias = resultado.split("Infração extra").length - 1;
  assert.equal(ocorrencias, 2, "esperado a linha extra clonada nas DUAS tabelas, não só na primeira");
  assert.doesNotMatch(resultado, /LINHA_MODELO/);
});

test("TR-Dispensa: Contrato completo replica vigência, dotação, natureza intelectual, sustentabilidade e reajuste no Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({
    S07_TipoInstrumentoContratual: "Contrato",
    S04_TipoVigencia: "Plurianual (até 5 anos)",
    S04_PrazoVigenciaAnos: "3",
    S01_ND: "339039", S01_PTRES: "170175", S01_Fonte: "0100000000",
    S11_NaturezaIntelectual: "Sim",
    S19_CriteriosEspecificos: "Uso de materiais reciclados",
    S13_TipoCorrecao: "Reajuste (sem dedicação exclusiva de mão de obra)", S13_IndiceReajuste: "IPCA",
  }));
  assert.ok(texto.includes("podendo ser prorrogado até o limite de 05 (cinco) anos"),
    "Cláusula Segunda deveria usar a redação plurianual");
  assert.ok(texto.includes("Prazo de vigência (anos): 3"));
  assert.ok(texto.includes("Natureza da Despesa (339039)"));
  assert.ok(texto.includes("PTRES (170175)"));
  const ocorrenciasCessao = texto.split("Ceder os direitos patrimoniais").length - 1;
  assert.equal(ocorrenciasCessao, 2, "esperado o bloco de cessão de direitos no corpo do TR E no Anexo VI");
  assert.ok(texto.includes("Critérios específicos: Uso de materiais reciclados"));
  assert.ok(texto.includes("variação do índice de referência: IPCA"));
  const ocorrenciasIpca = texto.split("IPCA").length - 1;
  assert.equal(ocorrenciasIpca, 2, "esperado o índice no corpo do TR E no Anexo VI");
});

test("TR-Dispensa: Contrato completo replica a garantia contratual no Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({
    S07_TipoInstrumentoContratual: "Contrato",
    S08_ExigeGarantia: "Sim", S08_PercentualGarantia: "3", S08_BaseCalculoGarantia: "valor anual do Contrato",
  }));
  const ocorrenciasComprovante = texto.split("comprovante de prestação de garantia").length - 1;
  assert.equal(ocorrenciasComprovante, 2, "esperado o texto de garantia no corpo do TR E no Anexo VI");
  const ocorrenciasPercentual = texto.split("Percentual da garantia: 3%").length - 1;
  assert.equal(ocorrenciasPercentual, 2);
  const ocorrenciasBase = texto.split("valor anual do Contrato").length - 1;
  assert.ok(ocorrenciasBase >= 2, "esperado a base de cálculo no corpo do TR E no Anexo VI");
});

test("TR-Dispensa: Contrato completo replica sanções (percentuais, base de cálculo e infrações extras) no Anexo VI", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({
    S07_TipoInstrumentoContratual: "Contrato",
    S16_MultaRescisoria: "10",
    S16_BaseCalculoMulta: "valor mensal do Contrato",
    S16_InfracoesExtras: [{ descricao: "Não portar crachá de identificação", grau: "1" }],
  }));
  const ocorrenciasMultaRescisoria = texto.split("multa rescisória de 10%").length - 1;
  assert.equal(ocorrenciasMultaRescisoria, 2, "esperado a multa rescisória no corpo do TR E no Anexo VI");
  const ocorrenciasBaseCalculo = texto.split("valor mensal do Contrato").length - 1;
  assert.ok(ocorrenciasBaseCalculo >= 8, `esperado a base de cálculo em pelo menos 8 lugares (4 no corpo + 4 no Anexo VI), achou ${ocorrenciasBaseCalculo}`);
  const ocorrenciasInfracaoExtra = texto.split("Não portar crachá de identificação").length - 1;
  assert.equal(ocorrenciasInfracaoExtra, 2, "esperado a infração extra na Tabela 1 do corpo do TR E na do Anexo VI");
});

test("PB: tipo de objeto Obra usa a redação de obra (Seção 1) e propaga para o resumo (Seção 3)", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({ PB01_TipoObjeto: "Obra" }));
  assert.ok(texto.includes("enquadram-se como obra, cujos padrões"));
  assert.ok(!texto.includes("enquadram-se como serviço comum de engenharia"));
  assert.ok(texto.includes("Tipo de objeto: Obra"));
});

test("PB: BIM Sim inclui o módulo na Seção 4 e a obrigação na Seção 18", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({ PB04_AdotaBIM: "Sim" }));
  const ocorrenciasBIM = texto.split("metodologia BIM").length - 1;
  assert.ok(ocorrenciasBIM >= 1, "esperado texto de BIM na Seção 4");
  assert.ok(texto.includes("Adotar a metodologia BIM na execução dos serviços"));
});

test("PB: BIM Não remove o módulo e a obrigação", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({ PB04_AdotaBIM: "Não" }));
  assert.ok(!texto.includes("Adotar a metodologia BIM na execução dos serviços"));
});

test("PB: vistoria obrigatória usa 'deverá realizar vistoria'", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({ PB09_TipoVistoria: "Obrigatória" }));
  assert.ok(texto.includes("o licitante deverá realizar vistoria"));
  assert.ok(!texto.includes("o licitante poderá realizar vistoria"));
});

test("PB: Quadro de Inovações Sim mantém o Anexo II-A; Não remove", async () => {
  const textoSim = textoDoDocumento(await gerarXmlPB({ PB26_ExigeQuadroInovacoes: "Sim" }));
  assert.ok(textoSim.includes("ANEXO II-A DO PROJETO BÁSICO"));
  assert.ok(textoSim.includes("Frações do objeto (Obras e serviços de engenharia)"));
  assert.ok(textoSim.includes("Terraplanagem"));
  const textoNao = textoDoDocumento(await gerarXmlPB({ PB26_ExigeQuadroInovacoes: "Não" }));
  assert.ok(!textoNao.includes("ANEXO II-A DO PROJETO BÁSICO"));
  assert.ok(!textoNao.includes("Terraplanagem"));
  // A menção fixa no corpo (Seção 26.3) some junto, mas o título da Seção 26
  // (que cita "Quadro de Inovações" no próprio nome do artigo) permanece.
  assert.ok(textoNao.includes("26. DA MATRIZ DE RISCOS E QUADRO DE INOVAÇÕES"));
});

test("PB: subcontratação Sim/Não troca a redação da Seção 31", async () => {
  const textoSim = textoDoDocumento(await gerarXmlPB({ PB31_PermiteSubcontratacao: "Sim" }));
  assert.ok(textoSim.includes("É permitida a subcontratação parcial do objeto"));
  const textoNao = textoDoDocumento(await gerarXmlPB({ PB31_PermiteSubcontratacao: "Não" }));
  assert.ok(textoNao.includes("Não será admitida a subcontratação do objeto"));
});

test("PB: multas de habilitação/condutas/recusa preenchem corpo e Minuta de Contrato, sem ____% órfão", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_RegimeContratacao: "Dispensa de Licitação",
    PB25_MultaNaoRegularizacaoHabilitacao: "5",
    PB25_MultaCondutasArt576a580: "10",
    PB25_MultaRecusaAssinarContrato: "15",
  }));
  assert.equal((texto.match(/5% sobre o valor estimado/g) || []).length, 2);
  assert.equal((texto.match(/10% sobre o valor estimado/g) || []).length, 2);
  assert.equal((texto.match(/15% sobre o valor da contrata/g) || []).length, 2);
});

test("PB: descrição do objeto se propaga para os Anexos I, II, IV e V", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({ PB01_DescricaoObjeto: "reforma do prédio sede" }));
  const ocorrencias = texto.split("reforma do prédio sede").length - 1;
  assert.ok(ocorrencias >= 5, `esperado o objeto na Seção 1 + 4 anexos, achou ${ocorrencias}`);
});

test("PB: Anexo VII (IMR) usa a estrutura oficial em branco + o exemplo real, sem critérios inventados", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({}));
  assert.ok(texto.includes("Indicador Nº ____"));
  assert.ok(texto.includes("Faixas de ajuste no pagamento"));
  assert.ok(texto.includes("Exemplo de indicador"));
  assert.ok(texto.includes("Prazo de atendimento de demandas (OS)"));
  assert.ok(texto.includes("Sistema informatizado de solicitação de serviços"));
  assert.ok(!texto.includes("Uso de materiais com qualidade"));
  assert.ok(!texto.includes("Critério de avaliação"));
});

test("PB: natureza 'Elaboração de projeto' remove agregados reciclados/PGRCC, garantia de obra, Anexo VIII e BDI", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_NaturezaObjeto: "Elaboração de projeto, laudo ou anteprojeto por terceiros",
    PB28_MetodologiaCustoIntelectual: "Fator K de 1,15, aplicado sobre o custo técnico direto.",
  }));
  assert.ok(!texto.includes("uso obrigatório de agregados reciclados"));
  assert.ok(!texto.includes("é de 5 (cinco) anos (art. 618 do Código Civil)"));
  assert.ok(!texto.includes("MODELO DE PLACA DE OBRA/SERVIÇO DE ENGENHARIA"));
  assert.ok(!texto.includes("Os custos diretos sem BDI totalizam"));
  assert.ok(texto.includes("adotando metodologia de Fator K e TRDE"));
  assert.ok(texto.includes("Fator K de 1,15, aplicado sobre o custo técnico direto."));
});

test("PB: natureza 'Execução de obra' padrão mantém agregados reciclados/PGRCC, garantia de obra, Anexo VIII e BDI", async () => {
  const texto = textoDoDocumento(await gerarXmlPB({
    PB01_NaturezaObjeto: "Execução de obra ou serviço já projetado",
    PB28_CriterioEstimativa: "Média", PB28_FonteReferencia: "Tabela SINAPI",
    PB28_CustoSemBDI: "100.000,00", PB28_PercentualBDI: "25", PB28_ValorTotalComBDI: "125.000,00",
  }));
  assert.ok(texto.includes("uso obrigatório de agregados reciclados"));
  assert.ok(texto.includes("é de 5 (cinco) anos (art. 618 do Código Civil)"));
  assert.ok(texto.includes("MODELO DE PLACA DE OBRA/SERVIÇO DE ENGENHARIA"));
  assert.ok(texto.includes("Os custos diretos sem BDI totalizam R$ 100.000,00"));
  assert.ok(!texto.includes("adotando metodologia de Fator K e TRDE"));
});

test("TR-Serviços: tabela de endereços adicionais preenche as linhas informadas", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({
    S06_EnderecosAdicionais: [
      { endereco: "Rua A, 123", observacao: "Unidade Norte" },
      { endereco: "Rua B, 456", observacao: "Unidade Sul" },
    ],
  }));
  assert.ok(texto.includes("Rua A, 123") && texto.includes("Unidade Norte"));
  assert.ok(texto.includes("Rua B, 456") && texto.includes("Unidade Sul"));
});

test("TR-Bens: tabela de endereços adicionais preenche as linhas informadas", async () => {
  const texto = textoDoDocumento(await gerarXml({
    S06_EnderecosAdicionais: [{ endereco: "Av. C, 789", observacao: "Depósito" }],
  }));
  assert.ok(texto.includes("Av. C, 789") && texto.includes("Depósito"));
});

test("TR-Dispensa: tabela de endereços adicionais preenche as linhas informadas", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({
    S06_EnderecosAdicionais: [{ endereco: "Rua D, 321", observacao: "Anexo" }],
  }));
  assert.ok(texto.includes("Rua D, 321") && texto.includes("Anexo"));
});

test("TR-Bens: sem endereços adicionais, a frase de introdução e a tabela vazia somem", async () => {
  const texto = textoDoDocumento(await gerarXml({}));
  assert.ok(!texto.includes("locais adicionais são os seguintes"));
  assert.ok(!texto.includes("Endereço\nObservação") && !texto.includes("EndereçoObservação"));
});

test("TR-Serviços: sem endereços adicionais, a frase de introdução e a tabela vazia somem", async () => {
  const texto = textoDoDocumento(await gerarXmlServicos({}));
  assert.ok(!texto.includes("locais adicionais são os seguintes"));
});

test("TR-Dispensa: sem endereços adicionais, a frase de introdução e a tabela vazia somem", async () => {
  const texto = textoDoDocumento(await gerarXmlDispensa({}));
  assert.ok(!texto.includes("locais adicionais são os seguintes"));
});
