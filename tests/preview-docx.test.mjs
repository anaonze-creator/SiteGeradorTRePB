import { test } from "node:test";
import assert from "node:assert/strict";
import { renderizarParagrafo, renderizarPreviewHtml } from "../js/preview-docx.js";

test("renderizarParagrafo: texto simples", () => {
  const xml = `<w:p><w:r><w:t>Texto simples</w:t></w:r></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>Texto simples</p>");
});

test("renderizarParagrafo: negrito e itálico", () => {
  const xml = `<w:p>` +
    `<w:r><w:rPr><w:b/></w:rPr><w:t>Negrito</w:t></w:r>` +
    `<w:r><w:t xml:space="preserve"> e </w:t></w:r>` +
    `<w:r><w:rPr><w:i/></w:rPr><w:t>itálico</w:t></w:r>` +
    `</w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p><strong>Negrito</strong> e <em>itálico</em></p>");
});

test("renderizarParagrafo: quebra de linha vira <br>", () => {
  const xml = `<w:p><w:r><w:t>Linha 1</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>Linha 2</w:t></w:r></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>Linha 1<br>Linha 2</p>");
});

test("renderizarParagrafo: escapa caracteres especiais de HTML", () => {
  const xml = `<w:p><w:r><w:t>A &amp; B &lt; C</w:t></w:r></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>A &amp; B &lt; C</p>");
});

test("renderizarParagrafo: parágrafo vazio", () => {
  const xml = `<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p></p>");
});

test("renderizarParagrafo: bookmark de campo vira <mark> com data-campo", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="1" w:name="S01_Objeto"/>` +
    `<w:r><w:t>Aquisição de bens</w:t></w:r>` +
    `<w:bookmarkEnd w:id="1"/>` +
    `</w:p>`;
  assert.equal(
    renderizarParagrafo(xml),
    `<p><mark class="preview-destaque" data-campo="S01_Objeto">Aquisição de bens</mark></p>`
  );
});

test("renderizarParagrafo: bookmark estrutural (_BLOCO_) não vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="2" w:name="PB01_NaturezaObjeto_BLOCO_EXECUCAO_DE_OBRA"/>` +
    `<w:r><w:t>Texto fixo do bloco</w:t></w:r>` +
    `<w:bookmarkEnd w:id="2"/>` +
    `</w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>Texto fixo do bloco</p>");
});

test("renderizarParagrafo: bookmark de campo dentro de bloco estrutural sobrevivente ainda vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="3" w:name="PB28_BDI_BLOCO_EXECUCAO_DE_OBRA"/>` +
    `<w:r><w:t>BDI: </w:t></w:r>` +
    `<w:bookmarkStart w:id="4" w:name="PB28_PercentualBDI"/>` +
    `<w:r><w:t>25%</w:t></w:r>` +
    `<w:bookmarkEnd w:id="4"/>` +
    `<w:bookmarkEnd w:id="3"/>` +
    `</w:p>`;
  assert.equal(
    renderizarParagrafo(xml),
    `<p>BDI: <mark class="preview-destaque" data-campo="PB28_PercentualBDI">25%</mark></p>`
  );
});

test("renderizarParagrafo: bookmark _LINHA_MODELO não vira <mark>", () => {
  const xml = `<w:p>` +
    `<w:bookmarkStart w:id="5" w:name="S15_InfracoesExtras_LINHA_MODELO"/>` +
    `<w:r><w:t>((DESCRICAO))</w:t></w:r>` +
    `<w:bookmarkEnd w:id="5"/>` +
    `</w:p>`;
  assert.equal(renderizarParagrafo(xml), "<p>((DESCRICAO))</p>");
});

test("renderizarPreviewHtml: parágrafos em sequência", () => {
  const xml = `<w:document><w:body>` +
    `<w:p><w:r><w:t>Primeiro</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t>Segundo</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240"/></w:sectPr>` +
    `</w:body></w:document>`;
  assert.equal(renderizarPreviewHtml(xml), "<p>Primeiro</p><p>Segundo</p>");
});

test("renderizarPreviewHtml: tabela simples", () => {
  const xml = `<w:document><w:body>` +
    `<w:tbl>` +
      `<w:tr><w:tc><w:p><w:r><w:t>Célula 1</w:t></w:r></w:p></w:tc>` +
              `<w:tc><w:p><w:r><w:t>Célula 2</w:t></w:r></w:p></w:tc></w:tr>` +
    `</w:tbl>` +
    `</w:body></w:document>`;
  assert.equal(
    renderizarPreviewHtml(xml),
    "<table><tr><td><p>Célula 1</p></td><td><p>Célula 2</p></td></tr></table>"
  );
});

test("renderizarPreviewHtml: parágrafo antes e depois de tabela, na ordem certa", () => {
  const xml = `<w:document><w:body>` +
    `<w:p><w:r><w:t>Antes</w:t></w:r></w:p>` +
    `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Meio</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:t>Depois</w:t></w:r></w:p>` +
    `</w:body></w:document>`;
  assert.equal(
    renderizarPreviewHtml(xml),
    "<p>Antes</p><table><tr><td><p>Meio</p></td></tr></table><p>Depois</p>"
  );
});
