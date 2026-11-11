const ICONE_CAIXA = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 8l9-4 9 4-9 4-9-4Z"/><path d="M3 8v8l9 4 9-4V8"/><path d="M12 12v8"/></svg>`;
const ICONE_ENGRENAGEM = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="3"/><path d="M19.4 13.5a1.7 1.7 0 0 0 .34 1.87l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06a1.7 1.7 0 0 0-1.87-.34 1.7 1.7 0 0 0-1 1.56V19.5a2 2 0 1 1-4 0v-.09a1.7 1.7 0 0 0-1.11-1.56 1.7 1.7 0 0 0-1.87.34l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06a1.7 1.7 0 0 0 .34-1.87 1.7 1.7 0 0 0-1.56-1H3.5a2 2 0 1 1 0-4h.09a1.7 1.7 0 0 0 1.56-1.11 1.7 1.7 0 0 0-.34-1.87l-.06-.06a2 2 0 1 1 2.83-2.83l.06.06a1.7 1.7 0 0 0 1.87.34H9a1.7 1.7 0 0 0 1-1.56V3.5a2 2 0 1 1 4 0v.09a1.7 1.7 0 0 0 1 1.56 1.7 1.7 0 0 0 1.87-.34l.06-.06a2 2 0 1 1 2.83 2.83l-.06.06a1.7 1.7 0 0 0-.34 1.87V9a1.7 1.7 0 0 0 1.56 1h.09a2 2 0 1 1 0 4h-.09a1.7 1.7 0 0 0-1.56 1Z"/></svg>`;
const ICONE_RAIO = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M13 2 4 14h7l-1 8 9-12h-7l1-8Z"/></svg>`;
const ICONE_PREDIO = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 21V6l8-3 8 3v15"/><path d="M4 21h16"/><path d="M9 21v-5h6v5"/><path d="M9 9h.01M9 13h.01M15 9h.01M15 13h.01"/></svg>`;
const ICONE_NOTA_DEMANDA = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V3a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1"/><path d="M8 10h8M8 14h8M8 18h5"/></svg>`;
const ICONE_NOTA_TECNICA = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.8 2.8-2-2 2.8-2.8Z"/></svg>`;
const ICONE_AVALIACAO_IMOVEIS = `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M3 11l7-6 7 6"/><path d="M5 10v9a1 1 0 0 0 1 1h4v-5h2v5h1"/><circle cx="17.5" cy="16.5" r="2.8"/><path d="M19.6 18.6 22 21"/></svg>`;

export function renderizarTriagem(container, aoConfirmar) {
  container.innerHTML = `
    <div class="hero-triagem">
      <h1>Gerador de Termo de Referência e Projeto Básico</h1>
      <p>Preenchimento guiado, seção por seção, com pré-visualização do
      documento final em tempo real. Escolha o instrumento que deseja
      elaborar: Termos de Referência, Projeto Básico, ou os estudos
      preliminares que os precedem (Nota de Demanda e Nota Técnica).</p>
    </div>
    <div class="grid-instrumentos">
      <button id="btn-tr-bens" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_CAIXA}</span>
        <span class="card-instrumento-titulo">Termo de Referência</span>
        <span class="card-instrumento-subtitulo">Aquisição de Bens</span>
        <span class="card-instrumento-desc">Art. 127 do RLC</span>
      </button>
      <button id="btn-tr-servicos" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_ENGRENAGEM}</span>
        <span class="card-instrumento-titulo">Termo de Referência</span>
        <span class="card-instrumento-subtitulo">Prestação de Serviços</span>
        <span class="card-instrumento-desc">Art. 129 do RLC</span>
      </button>
      <button id="btn-tr-dispensa" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_RAIO}</span>
        <span class="card-instrumento-titulo">Termo de Referência</span>
        <span class="card-instrumento-subtitulo">Serviços por Dispensa</span>
        <span class="card-instrumento-desc">Art. 416, II, do RLC</span>
      </button>
      <button id="btn-tr-dispensa-avaliacao-imoveis" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_AVALIACAO_IMOVEIS}</span>
        <span class="card-instrumento-titulo">Termo de Referência</span>
        <span class="card-instrumento-subtitulo">Avaliação de Imóveis (Dispensa)</span>
        <span class="card-instrumento-desc">Art. 416, II, do RLC</span>
      </button>
      <button id="btn-pb" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_PREDIO}</span>
        <span class="card-instrumento-titulo">Projeto Básico</span>
        <span class="card-instrumento-subtitulo">Obras e Serviços de Engenharia</span>
        <span class="card-instrumento-desc">Art. 157 do RLC</span>
      </button>
      <button id="btn-nota-demanda" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_NOTA_DEMANDA}</span>
        <span class="card-instrumento-titulo">Nota de Demanda</span>
        <span class="card-instrumento-subtitulo">Estudo Preliminar — Área Demandante</span>
        <span class="card-instrumento-desc">Art. 108 e 115 do RLC</span>
      </button>
      <button id="btn-nota-tecnica" class="card-instrumento" type="button">
        <span class="card-instrumento-icone">${ICONE_NOTA_TECNICA}</span>
        <span class="card-instrumento-titulo">Nota Técnica</span>
        <span class="card-instrumento-subtitulo">Estudo Preliminar — Área Técnica</span>
        <span class="card-instrumento-desc">Art. 108, 115 e 123 do RLC</span>
      </button>
    </div>
  `;
  container.querySelector("#btn-tr-bens")
    .addEventListener("click", () => aoConfirmar("TR-Bens"));
  container.querySelector("#btn-tr-servicos")
    .addEventListener("click", () => aoConfirmar("TR-Serviços"));
  container.querySelector("#btn-tr-dispensa")
    .addEventListener("click", () => aoConfirmar("TR-Dispensa"));
  container.querySelector("#btn-tr-dispensa-avaliacao-imoveis")
    .addEventListener("click", () => aoConfirmar("TR-Dispensa-Avaliacao-Imoveis"));
  container.querySelector("#btn-pb")
    .addEventListener("click", () => aoConfirmar("PB"));
  container.querySelector("#btn-nota-demanda")
    .addEventListener("click", () => aoConfirmar("Nota-Demanda"));
  container.querySelector("#btn-nota-tecnica")
    .addEventListener("click", () => aoConfirmar("Nota-Tecnica"));
}
