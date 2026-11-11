import { campoVisivel, controladoresBlocoVariante, camposOcultos } from "./modelo.js";

function localizarParagrafoDoBookmark(xml, nomeBookmark) {
  const marcador = `w:name="${nomeBookmark}"`;
  const posNome = xml.indexOf(marcador);
  if (posNome === -1) return null;

  const inicioParagrafo = xml.lastIndexOf("<w:p>", posNome) !== -1
    ? xml.lastIndexOf("<w:p>", posNome)
    : xml.lastIndexOf("<w:p ", posNome);

  // O bookmarkEnd correspondente (mesmo w:id do bookmarkStart achado) pode
  // estar num parágrafo mais adiante — blocos que abrangem vários parágrafos
  // (ex.: a bifurcação de fiscalização da Seção 11 do TR-Serviços) têm o
  // bookmarkStart no primeiro parágrafo e o bookmarkEnd só no último.
  const inicioTagStart = xml.lastIndexOf("<w:bookmarkStart", posNome);
  const fimTagStart = xml.indexOf(">", inicioTagStart) + 1;
  const tagStart = xml.slice(inicioTagStart, fimTagStart);
  const idMatch = tagStart.match(/w:id="([^"]+)"/);
  if (!idMatch) return null;
  const bookmarkId = idMatch[1];

  const regexEnd = new RegExp(`<w:bookmarkEnd[^>]*w:id="${bookmarkId}"[^>]*/>`);
  const matchEnd = regexEnd.exec(xml.slice(posNome));
  if (!matchEnd) return null;
  const posEnd = posNome + matchEnd.index;

  const fimParagrafo = xml.indexOf("</w:p>", posEnd) + "</w:p>".length;
  return { inicioParagrafo, fimParagrafo };
}

// Quebras de linha digitadas nas <textarea> precisam virar <w:br/> entre runs:
// um "\n" cru dentro de <w:t> o Word simplesmente ignora, colapsando a resposta
// em um parágrafo corrido.
function construirRunsComQuebras(texto) {
  return texto
    .split("\n")
    .map((linha) => `<w:r><w:t xml:space="preserve">${linha}</w:t></w:r>`)
    .join("<w:r><w:br/></w:r>");
}

export function substituirBookmark(documentXml, nomeBookmark, novoTexto) {
  const posicao = localizarParagrafoDoBookmark(documentXml, nomeBookmark);
  if (!posicao) return documentXml;
  const trechoParagrafo = documentXml.slice(posicao.inicioParagrafo, posicao.fimParagrafo);
  // O regex precisa estar ancorado NO NOME PEDIDO: parágrafos com mais de um
  // bookmark (1.4 tem ND, PTRES e Fonte juntos) casariam com o primeiro
  // bookmark do parágrafo, escrevendo o valor no campo errado.
  const nomeEscapado = nomeBookmark.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const regexBookmarkEspecifico = new RegExp(
    `(<w:bookmarkStart[^>]*w:name="${nomeEscapado}"[^>]*\\/>)([\\s\\S]*?)(<w:bookmarkEnd)`
  );
  const trechoSubstituido = trechoParagrafo.replace(
    regexBookmarkEspecifico,
    (match, inicio, meio, fim) => inicio + construirRunsComQuebras(novoTexto) + fim
  );
  return documentXml.slice(0, posicao.inicioParagrafo) + trechoSubstituido + documentXml.slice(posicao.fimParagrafo);
}

export function removerBlocoBookmark(documentXml, nomeBookmark) {
  const posicao = localizarParagrafoDoBookmark(documentXml, nomeBookmark);
  if (!posicao) return documentXml;
  return documentXml.slice(0, posicao.inicioParagrafo) + documentXml.slice(posicao.fimParagrafo);
}

// Remove TODAS as ocorrências de um bookmark de bloco, não só a primeira —
// diferente de removerBlocoBookmark (usado pelos controladores Sim/Não, cujo
// bookmark só existe uma vez por opção). Necessário para o container de
// tabela-linhas, que pode se repetir (corpo + Anexo, quando o instrumento
// contratual replica a seção inteira, ex.: TR-Dispensa com Contrato completo).
export function removerTodosBlocosBookmark(documentXml, nomeBookmark) {
  let resultado = documentXml;
  while (true) {
    const antes = resultado;
    resultado = removerBlocoBookmark(resultado, nomeBookmark);
    if (resultado === antes) break;
  }
  return resultado;
}

// Convenção de nomes dos bookmarks que envolvem cada variante de um bloco de
// texto condicional no template (ver templates/build_template_tr_bens.py e
// templates/build_template_tr_servicos.py, função add_bloco_variante). O
// CONJUNTO de campos que controlam bloco NÃO é repetido aqui: vem de
// modelo.js/controladoresBlocoVariante, derivado de `blocoTexto` nos dados —
// fonte única compartilhada com js/app.js.
function normalizarVariante(valor) {
  return valor
    .normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/[^A-Z0-9]+/g, "_");
}

function nomeBlocoVariante(campoId, opcao) {
  return `${campoId}_BLOCO_${normalizarVariante(opcao)}`;
}

function localizarLinhaTabela(xml, nomeBookmark) {
  const marcador = `w:name="${nomeBookmark}"`;
  const posNome = xml.indexOf(marcador);
  if (posNome === -1) return null;
  const inicioLinha = xml.lastIndexOf("<w:tr>", posNome) !== -1
    ? xml.lastIndexOf("<w:tr>", posNome)
    : xml.lastIndexOf("<w:tr ", posNome);
  const fimLinha = xml.indexOf("</w:tr>", posNome) + "</w:tr>".length;
  return { inicioLinha, fimLinha };
}

function removerMarcacoesBookmark(trecho) {
  return trecho
    .replace(/<w:bookmarkStart[^>]*\/>/g, "")
    .replace(/<w:bookmarkEnd[^>]*\/>/g, "");
}

// Clona a linha-modelo de uma tabela (marcada com bookmark
// `${campo.id}_LINHA_MODELO`, contendo placeholders `((NOMEDACOLUNA))` em
// cada célula) uma vez por item de `linhas`, e remove a linha-modelo do
// resultado final. Usado pelo campo tipo tabela-linhas (ex.: infrações
// extras da Seção 15 do TR-Serviços). Processa TODAS as ocorrências do
// bookmark no documento — a mesma tabela pode existir mais de uma vez
// (ex.: no corpo do TR-Dispensa e de novo dentro do Anexo VI, quando o
// instrumento contratual é o Contrato completo).
export function inserirLinhasTabela(documentXml, campo, linhas) {
  const nomeBookmark = `${campo.id}_LINHA_MODELO`;
  let resultado = documentXml;
  while (true) {
    const posicao = localizarLinhaTabela(resultado, nomeBookmark);
    if (!posicao) break;
    const linhaModeloXml = resultado.slice(posicao.inicioLinha, posicao.fimLinha);
    const linhasGeradas = linhas.map((linha) => {
      let linhaXml = removerMarcacoesBookmark(linhaModeloXml);
      for (const coluna of campo.colunas) {
        const marcadorColuna = `((${coluna.id.toUpperCase()}))`;
        linhaXml = linhaXml.split(marcadorColuna).join(escaparXml(String(linha[coluna.id] || "")));
      }
      return linhaXml;
    }).join("");
    resultado = resultado.slice(0, posicao.inicioLinha) + linhasGeradas + resultado.slice(posicao.fimLinha);
  }
  return resultado;
}

export async function gerarDocx(templateArrayBuffer, respostas, secoes) {
  const zip = await JSZip.loadAsync(templateArrayBuffer);
  let documentXml = await zip.file("word/document.xml").async("string");

  for (const secao of secoes) {
    for (const campo of secao.campos) {
      // Campos tabela-linhas não escrevem num bookmark de valor único — são
      // resolvidos à parte, clonando linha de tabela (ver loop abaixo).
      if (campo.input === "tabela-linhas") continue;
      // Campo escondido pela resposta atual não pode escrever no documento:
      // o elaborador pode ter respondido "Sim", preenchido, e depois mudado
      // para "Não" — a resposta antiga fica no rascunho, mas não vale mais.
      if (!campoVisivel(campo, respostas, secoes)) continue;
      const valor = respostas[campo.id];
      if (valor !== undefined && valor !== null && valor !== "") {
        documentXml = substituirBookmark(documentXml, campo.id, escaparXml(String(valor)));
      }
    }
  }

  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (campo.input !== "tabela-linhas") continue;
      const linhas = Array.isArray(respostas[campo.id]) ? respostas[campo.id] : [];
      documentXml = inserirLinhasTabela(documentXml, campo, linhas);
      // Sem nenhuma linha extra, a frase de introdução + a tabela (só com o
      // cabeçalho, sem nenhum dado) ficariam penduradas no documento sem
      // servir pra nada. O par de sentinelas `${campo.id}_CONTAINER`
      // (ver add_bloco_variante_sentinela_inicio/fim nos templates Python)
      // envolve a introdução + a tabela inteira para remoção conjunta.
      if (linhas.length === 0) {
        documentXml = removerTodosBlocosBookmark(documentXml, `${campo.id}_CONTAINER`);
      }
    }
  }

  if (respostas.S01_DescricaoObjeto) {
    documentXml = substituirBookmark(documentXml, "S01_DescricaoObjeto_AnexoI", escaparXml(String(respostas.S01_DescricaoObjeto)));
  }

  // Alguns valores se repetem em mais de um ponto do documento (o mesmo
  // bookmark só existe uma vez no template). Um campo declara
  // `duplicarPara: [...]` com os nomes de bookmark extras que recebem o
  // mesmo valor, via substituição simples.
  for (const secao of secoes) {
    for (const campo of secao.campos) {
      if (!campo.duplicarPara || !campo.duplicarPara.length) continue;
      if (!campoVisivel(campo, respostas, secoes)) continue;
      const valor = respostas[campo.id];
      if (valor === undefined || valor === null || valor === "") continue;
      for (const nomeExtra of campo.duplicarPara) {
        documentXml = substituirBookmark(documentXml, nomeExtra, escaparXml(String(valor)));
      }
    }
  }

  for (const campo of controladoresBlocoVariante(secoes)) {
    const respondida = respostas[campo.id] || campo.padrao;
    const basesDeBookmark = [campo.id, ...(campo.blocosAdicionais || [])];
    // Sempre remove todas as variantes NÃO escolhidas — nunca deixa dois
    // blocos contraditórios juntos no documento. Resposta ausente cai no
    // padrão declarado pelo campo. `blocosAdicionais` permite que a mesma
    // resposta controle mais de um bloco/seção/anexo ao mesmo tempo.
    for (const opcao of campo.opcoes) {
      if (opcao !== respondida) {
        for (const base of basesDeBookmark) {
          documentXml = removerBlocoBookmark(documentXml, nomeBlocoVariante(base, opcao));
        }
      }
    }
  }

  // Parágrafos dependentes: cada campo escondido tem, no modelo, um parágrafo
  // próprio que só existe por causa dele ("Percentual da garantia: ____%",
  // "Limites: entre ____% e ____%", "Índice de referência:", ...). Sem esta
  // remoção o documento sairia com o placeholder cru mesmo quando o
  // condicional é "Não". Remover pelo bookmark do próprio campo dispensa
  // bookmarks extras no modelo; parágrafos com dois campos ocultos são
  // removidos na primeira passagem e a segunda vira no-op.
  for (const campo of camposOcultos(secoes, respostas)) {
    documentXml = removerBlocoBookmark(documentXml, campo.id);
  }

  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "arraybuffer" });
}

function escaparXml(texto) {
  return texto
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
