import { SECOES_TR_BENS } from "./dados-tr-bens.js";
import { SECOES_TR_SERVICOS } from "./dados-tr-servicos.js";
import { SECOES_TR_DISPENSA } from "./dados-tr-dispensa.js";
import { SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS } from "./dados-tr-dispensa-avaliacao-imoveis.js";
import { SECOES_PB } from "./dados-pb.js";
import { SECOES_NOTA_DEMANDA } from "./dados-nota-demanda.js";
import { SECOES_NOTA_TECNICA } from "./dados-nota-tecnica.js";
import { renderizarSumario, renderizarSecoes, atualizarProgresso } from "./render.js";
import { renderizarTriagem } from "./triagem.js";
import { renderizarConferencia } from "./conferencia.js";
import { serializarRascunho, desserializarRascunho } from "./persistencia.js";
import { controladoresCondicionais } from "./modelo.js";
import { gerarDocx } from "./gerador-docx.js";
import { renderizarPreviewHtml } from "./preview-docx.js";
import { valorPorExtenso } from "./valor-extenso.js";

const INSTRUMENTOS = {
  "TR-Bens": {
    secoes: SECOES_TR_BENS,
    titulo: "Termo de Referência — Aquisição de Bens",
    templateUrl: "templates/Template_TR_Bens.docx",
    chaveAutosave: "gerador-tr-pb:tr-bens:rascunho",
    nomeArquivoRascunho: "rascunho-tr-bens.json",
    nomeArquivoGerado: "TR-Bens-gerado.docx",
  },
  "TR-Serviços": {
    secoes: SECOES_TR_SERVICOS,
    titulo: "Termo de Referência — Prestação de Serviços",
    templateUrl: "templates/Template_TR_Servicos.docx",
    chaveAutosave: "gerador-tr-pb:tr-servicos:rascunho",
    nomeArquivoRascunho: "rascunho-tr-servicos.json",
    nomeArquivoGerado: "TR-Servicos-gerado.docx",
  },
  "TR-Dispensa": {
    secoes: SECOES_TR_DISPENSA,
    titulo: "Termo de Referência — Serviços por Dispensa de Licitação",
    templateUrl: "templates/Template_TR_Dispensa.docx",
    chaveAutosave: "gerador-tr-pb:tr-dispensa:rascunho",
    nomeArquivoRascunho: "rascunho-tr-dispensa.json",
    nomeArquivoGerado: "TR-Dispensa-gerado.docx",
  },
  "TR-Dispensa-Avaliacao-Imoveis": {
    secoes: SECOES_TR_DISPENSA_AVALIACAO_IMOVEIS,
    titulo: "Termo de Referência — Avaliação de Imóveis (Dispensa de Licitação)",
    templateUrl: "templates/Template_TR_Dispensa_Avaliacao_Imoveis.docx",
    chaveAutosave: "gerador-tr-pb:tr-dispensa-avaliacao-imoveis:rascunho",
    nomeArquivoRascunho: "rascunho-tr-dispensa-avaliacao-imoveis.json",
    nomeArquivoGerado: "TR-Dispensa-Avaliacao-Imoveis-gerado.docx",
  },
  "PB": {
    secoes: SECOES_PB,
    titulo: "Projeto Básico",
    templateUrl: "templates/Template_PB.docx",
    chaveAutosave: "gerador-tr-pb:pb:rascunho",
    nomeArquivoRascunho: "rascunho-pb.json",
    nomeArquivoGerado: "PB-gerado.docx",
  },
  "Nota-Demanda": {
    secoes: SECOES_NOTA_DEMANDA,
    titulo: "Nota de Demanda",
    templateUrl: "templates/Template_Nota_Demanda.docx",
    chaveAutosave: "gerador-tr-pb:nota-demanda:rascunho",
    nomeArquivoRascunho: "rascunho-nota-demanda.json",
    nomeArquivoGerado: "Nota-de-Demanda-gerada.docx",
  },
  "Nota-Tecnica": {
    secoes: SECOES_NOTA_TECNICA,
    titulo: "Nota Técnica",
    templateUrl: "templates/Template_Nota_Tecnica.docx",
    chaveAutosave: "gerador-tr-pb:nota-tecnica:rascunho",
    nomeArquivoRascunho: "rascunho-nota-tecnica.json",
    nomeArquivoGerado: "Nota-Tecnica-gerada.docx",
  },
};

let instrumentoAtual = "TR-Bens";
let respostas = {};

function config() {
  return INSTRUMENTOS[instrumentoAtual];
}

function carregarAutosaveSeExistir() {
  let salvo;
  try {
    // Acesso ao storage pode falhar antes mesmo de ler conteúdo (navegação
    // privada, política corporativa, cota estourada). Sem este try/catch a
    // exceção abortaria iniciar() e a tela de triagem nunca apareceria.
    salvo = localStorage.getItem(config().chaveAutosave);
  } catch (erro) {
    console.warn("Armazenamento local indisponível; autosave desativado nesta sessão.", erro);
    return;
  }
  if (!salvo) return;
  try {
    const { respostas: respostasSalvas } = desserializarRascunho(salvo);
    respostas = respostasSalvas;
  } catch (erro) {
    alert("O rascunho salvo automaticamente estava corrompido e foi descartado: " + erro.message);
    try { localStorage.removeItem(config().chaveAutosave); } catch { /* já sem acesso ao storage */ }
  }
}

function persistirAutosave() {
  try {
    localStorage.setItem(config().chaveAutosave, serializarRascunho(instrumentoAtual, respostas));
  } catch (erro) {
    // Chamado a cada tecla digitada: avisa no console, sem alert repetitivo.
    console.warn("Não foi possível salvar o rascunho automaticamente.", erro);
  }
}

function mostrarPainel() {
  document.getElementById("tela-triagem").hidden = true;
  document.getElementById("tela-painel").hidden = false;
  document.getElementById("btn-voltar-triagem").hidden = false;
}

function mostrarTriagem() {
  document.getElementById("tela-painel").hidden = true;
  document.getElementById("tela-triagem").hidden = false;
  document.getElementById("btn-voltar-triagem").hidden = true;
}

function montarTelaPainel() {
  const { secoes, titulo, templateUrl, nomeArquivoGerado } = config();
  const painel = document.getElementById("tela-painel-formulario");
  painel.innerHTML = `
    <div class="cartao-painel">
      <div class="cabecalho-painel">${titulo}</div>
      <div class="corpo-painel">
        <nav id="sumario" class="br-list"></nav>
        <div id="secoes"></div>
        <div id="barra-progresso" class="barra-progresso" role="progressbar" aria-label="Progresso de preenchimento das especificações" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
          <div class="barra-progresso-trilho"><div class="barra-progresso-preenchimento"></div></div>
          <p class="barra-progresso-texto"></p>
        </div>
        <button id="btn-gerar" class="br-button primary" type="button">Gerar Documento</button>
      </div>
    </div>
  `;
  // Mesma fonte de verdade usada por js/gerador-docx.js para saber quais
  // campos controlam blocos condicionais no documento.
  const camposControladores = controladoresCondicionais(secoes);
  const todosOsCampos = secoes.flatMap((s) => s.campos);
  // Campos que o elaborador já digitou diretamente — usado só para saber
  // quando NÃO sobrescrever um valor por extenso editado à mão (ver
  // `autoPreencheExtenso` abaixo).
  const camposEditadosManualmente = new Set();
  function algumCampoDependeDe(campoId) {
    return camposControladores.includes(campoId);
  }

  let templateBufferCache = null;
  async function obterTemplateBuffer() {
    if (!templateBufferCache) {
      const resposta = await fetch(templateUrl);
      if (!resposta.ok) throw new Error(`Não foi possível carregar o modelo (${resposta.status}).`);
      templateBufferCache = await resposta.arrayBuffer();
    }
    return templateBufferCache;
  }

  let timeoutPreview = null;
  let ultimoCampoEditado = null;
  function agendarAtualizacaoPreview(campoId) {
    ultimoCampoEditado = campoId;
    if (timeoutPreview) clearTimeout(timeoutPreview);
    timeoutPreview = setTimeout(atualizarPreview, 400);
  }

  // Rola o painel de preview até o trecho marcado com o campo editado por
  // último, usando o atributo data-campo que renderizarPreviewHtml grava em
  // cada <mark> (mesmo nome de bookmark usado na geração do .docx).
  //
  // Usa scrollBy() no próprio painel em vez de mark.scrollIntoView(): o
  // scrollIntoView() nativo rola TODOS os ancestrais roláveis necessários
  // para exibir o elemento — inclusive a janela/página inteira — não só o
  // painel de preview, fazendo a tela pular pra cima a cada tecla digitada.
  function rolarPreviewParaCampo(painelPreview, campoId) {
    if (!campoId) return;
    const alvo = painelPreview.querySelector(`mark[data-campo~="${CSS.escape(campoId)}"]`);
    if (!alvo) return;
    const retPainel = painelPreview.getBoundingClientRect();
    const retAlvo = alvo.getBoundingClientRect();
    const deltaParaCentralizar = (retAlvo.top + retAlvo.height / 2) - (retPainel.top + retPainel.height / 2);
    painelPreview.scrollBy({ top: deltaParaCentralizar, behavior: "smooth" });
  }

  async function atualizarPreview() {
    const painelPreview = document.getElementById("painel-preview");
    try {
      const buffer = await obterTemplateBuffer();
      const docxGerado = await gerarDocx(buffer, respostas, secoes);
      const zip = await JSZip.loadAsync(docxGerado);
      const documentXml = await zip.file("word/document.xml").async("string");
      painelPreview.innerHTML = renderizarPreviewHtml(documentXml);
      rolarPreviewParaCampo(painelPreview, ultimoCampoEditado);
    } catch (erro) {
      // Estado intermediário inconsistente de `respostas` enquanto o
      // elaborador ainda digita não deve interromper a digitação nem gerar
      // alert — só o fluxo de geração final (btn-gerar) precisa disso.
      console.warn("Não foi possível atualizar a pré-visualização.", erro);
    }
  }

  function aoMudar(id, valor) {
    respostas[id] = valor;
    camposEditadosManualmente.add(id);
    // Preenche automaticamente o valor por extenso a partir do campo
    // monetário, salvo se o elaborador já tiver digitado algo à mão nesse
    // campo — atualiza só o <input> em tela (sem `rerenderizar()`, que
    // refaria todo o DOM das seções e derrubaria o foco/cursor de quem
    // ainda está digitando o valor numérico).
    const campo = todosOsCampos.find((c) => c.id === id);
    if (campo?.autoPreencheExtenso && !camposEditadosManualmente.has(campo.autoPreencheExtenso)) {
      const valorExtenso = valor === "" ? "" : valorPorExtenso(valor);
      respostas[campo.autoPreencheExtenso] = valorExtenso;
      const inputExtenso = document.getElementById(campo.autoPreencheExtenso);
      if (inputExtenso) inputExtenso.value = valorExtenso;
    }
    persistirAutosave();
    atualizarProgresso(secoes, respostas, document.getElementById("barra-progresso"));
    // Só reconstrói o DOM das seções quando o campo alterado controla a
    // visibilidade de outro campo. Refazer o DOM a cada tecla digitada nos
    // campos de texto perderia o foco/cursor do elaborador no meio da digitação.
    if (algumCampoDependeDe(id)) {
      rerenderizar();
    }
    agendarAtualizacaoPreview(id);
  }

  function rerenderizar() {
    renderizarSecoes(secoes, respostas, document.getElementById("secoes"), aoMudar);
  }

  renderizarSumario(secoes, document.getElementById("sumario"));
  rerenderizar();
  atualizarProgresso(secoes, respostas, document.getElementById("barra-progresso"));
  atualizarPreview();

  document.getElementById("btn-gerar").addEventListener("click", () => {
    const overlay = document.getElementById("conferencia-overlay");
    renderizarConferencia(secoes, respostas, overlay, async () => {
      overlay.hidden = true;
      try {
        const templateBuffer = await obterTemplateBuffer();
        const docxGerado = await gerarDocx(templateBuffer, respostas, secoes);
        const blob = new Blob([docxGerado], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = nomeArquivoGerado;
        a.click();
        URL.revokeObjectURL(url);
      } catch (erro) {
        alert("Não foi possível gerar o documento: " + erro.message);
      }
    });
  });
}

function iniciar() {
  renderizarTriagem(document.getElementById("tela-triagem"), (instrumentoEscolhido) => {
    instrumentoAtual = instrumentoEscolhido;
    respostas = {};
    carregarAutosaveSeExistir();
    mostrarPainel();
    montarTelaPainel();
  });

  document.getElementById("btn-voltar-triagem").addEventListener("click", () => {
    // Não apaga `respostas`: o autosave já salvou o progresso deste
    // instrumento a cada mudança, então escolher o mesmo instrumento de
    // novo na triagem recarrega o rascunho sozinho (carregarAutosaveSeExistir).
    mostrarTriagem();
  });

  document.getElementById("btn-salvar-rascunho").addEventListener("click", () => {
    const json = serializarRascunho(instrumentoAtual, respostas);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = config().nomeArquivoRascunho;
    a.click();
    URL.revokeObjectURL(url);
  });

  document.getElementById("btn-carregar-rascunho").addEventListener("click", () => {
    document.getElementById("input-carregar-rascunho").click();
  });

  document.getElementById("input-carregar-rascunho").addEventListener("change", async (evento) => {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;
    try {
      const texto = await arquivo.text();
      const { instrumento, respostas: respostasCarregadas } = desserializarRascunho(texto);
      if (!INSTRUMENTOS[instrumento]) {
        alert(`Rascunho de instrumento desconhecido: "${instrumento}".`);
        return;
      }
      if (instrumento !== instrumentoAtual) {
        if (!confirm(`Este rascunho é de "${INSTRUMENTOS[instrumento].titulo}", diferente do instrumento atual (${config().titulo}). Carregar mesmo assim?`)) return;
        instrumentoAtual = instrumento;
      }
      respostas = respostasCarregadas;
      persistirAutosave();
      // O botão "Carregar Rascunho" fica no cabeçalho fixo, então o rascunho
      // pode ser carregado ainda na triagem: é preciso trocar de tela, senão
      // a ação parece não ter surtido efeito nenhum.
      mostrarPainel();
      montarTelaPainel();
      alert("Rascunho carregado com sucesso.");
    } catch (erro) {
      alert("Não foi possível carregar o rascunho: " + erro.message);
    }
  });
}

iniciar();
