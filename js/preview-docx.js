// Concatena o texto de um run (<w:r>), aplicando negrito/itálico do <w:rPr>
// e convertendo <w:br/> em <br>. O texto dentro de <w:t> já vem com
// entidades XML (&amp;, &lt;, &gt;) escapadas pelo próprio arquivo .docx —
// as mesmas entidades são válidas em HTML, então o texto é reaproveitado
// como está, sem re-escapar (evita "&amp;" virar "&amp;amp;").
function renderizarRun(runXml) {
  const negrito = /<w:b\s*\/>|<w:b\/>|<w:b>/.test(runXml);
  const italico = /<w:i\s*\/>|<w:i\/>|<w:i>/.test(runXml);

  const partes = [];
  const regexConteudo = /<w:t[^>]*>([\s\S]*?)<\/w:t>|<w:br\s*\/>/g;
  let m;
  while ((m = regexConteudo.exec(runXml)) !== null) {
    if (m[0].startsWith("<w:br")) {
      partes.push("<br>");
    } else {
      partes.push(m[1]);
    }
  }
  let texto = partes.join("");
  if (!texto) return "";
  if (negrito) texto = `<strong>${texto}</strong>`;
  if (italico) texto = `<em>${texto}</em>`;
  return texto;
}

// Bookmarks de bloco/variante (add_bloco_variante e as duas sentinelas em
// templates/build_template_*.py) sempre terminam em "_BLOCO_<VARIANTE>", e
// linhas-modelo de tabela em "_LINHA_MODELO". gerarDocx() só remove as
// variantes NÃO escolhidas — a variante sobrevivente mantém suas próprias
// tags de bookmark no XML de saída. Nenhum ID de campo real usa esses
// sufixos, então o padrão é seguro para distinguir "estrutural" de "campo".
function ehBookmarkEstrutural(nome) {
  return /_BLOCO_[A-Z0-9_]+$/.test(nome) || nome.endsWith("_LINHA_MODELO");
}

// O nome do bookmark de campo vira o atributo data-campo do <mark>, para
// que js/app.js consiga rolar o painel de preview até o trecho
// correspondente ao campo que o elaborador está digitando (scrollIntoView
// por CSS.escape(id) via seletor [data-campo~="..."]). Como o texto do
// bloco marcado só é conhecido por inteiro quando o bookmark fecha, ele é
// acumulado num buffer em vez de escrito direto em `saida`.
export function renderizarParagrafo(paragrafoXml) {
  const regexToken = /<w:bookmarkStart[^>]*\/>|<w:bookmarkEnd[^>]*\/>|<w:r\b[\s\S]*?<\/w:r>/g;
  const idsDeCampoAbertos = new Set();
  let nomesDoBlocoAtual = [];
  let saida = "";
  let bufferMarcado = "";
  let marcando = false;
  let m;
  while ((m = regexToken.exec(paragrafoXml)) !== null) {
    const token = m[0];
    if (token.startsWith("<w:bookmarkStart")) {
      const nomeMatch = token.match(/w:name="([^"]+)"/);
      const idMatch = token.match(/w:id="([^"]+)"/);
      if (nomeMatch && idMatch && !ehBookmarkEstrutural(nomeMatch[1])) {
        idsDeCampoAbertos.add(idMatch[1]);
        nomesDoBlocoAtual.push(nomeMatch[1]);
        marcando = true;
      }
    } else if (token.startsWith("<w:bookmarkEnd")) {
      const idMatch = token.match(/w:id="([^"]+)"/);
      if (idMatch && idsDeCampoAbertos.has(idMatch[1])) {
        idsDeCampoAbertos.delete(idMatch[1]);
        if (marcando && idsDeCampoAbertos.size === 0) {
          const nomes = [...new Set(nomesDoBlocoAtual)].join(" ");
          saida += `<mark class="preview-destaque" data-campo="${nomes}">${bufferMarcado}</mark>`;
          bufferMarcado = "";
          nomesDoBlocoAtual = [];
          marcando = false;
        }
      }
    } else {
      const texto = renderizarRun(token);
      if (marcando) {
        bufferMarcado += texto;
      } else {
        saida += texto;
      }
    }
  }
  if (marcando) {
    const nomes = [...new Set(nomesDoBlocoAtual)].join(" ");
    saida += `<mark class="preview-destaque" data-campo="${nomes}">${bufferMarcado}</mark>`;
  }
  return `<p>${saida}</p>`;
}

function renderizarCelula(celulaXml) {
  const paragrafos = celulaXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  return `<td>${paragrafos.map(renderizarParagrafo).join("")}</td>`;
}

function renderizarLinha(linhaXml) {
  const celulas = linhaXml.match(/<w:tc\b[\s\S]*?<\/w:tc>/g) || [];
  return `<tr>${celulas.map(renderizarCelula).join("")}</tr>`;
}

function renderizarTabela(tabelaXml) {
  const linhas = tabelaXml.match(/<w:tr\b[\s\S]*?<\/w:tr>/g) || [];
  return `<table>${linhas.map(renderizarLinha).join("")}</table>`;
}

export function renderizarPreviewHtml(documentXml) {
  const corpoMatch = documentXml.match(/<w:body>([\s\S]*)<\/w:body>/);
  if (!corpoMatch) return "";
  const corpo = corpoMatch[1];
  const regexBloco = /<w:tbl>[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>/g;
  const blocos = corpo.match(regexBloco) || [];
  return blocos
    .map((bloco) => (bloco.startsWith("<w:tbl") ? renderizarTabela(bloco) : renderizarParagrafo(bloco)))
    .join("");
}
