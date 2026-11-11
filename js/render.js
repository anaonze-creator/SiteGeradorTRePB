import { campoVisivel, progresso } from "./modelo.js";

export function renderizarSumario(secoes, container) {
  container.innerHTML = "";
  const cartao = document.createElement("div");
  cartao.className = "sumario-card";
  const lista = document.createElement("ul");
  lista.className = "br-list";
  for (const secao of secoes) {
    const item = document.createElement("li");
    const link = document.createElement("a");
    link.href = `#secao-${secao.id}`;
    link.textContent = `${secao.id}. ${secao.titulo}`;
    item.appendChild(link);
    lista.appendChild(item);
  }
  cartao.appendChild(lista);
  container.appendChild(cartao);
}

function criarCampoInput(campo, valorAtual, aoMudar) {
  const wrapper = document.createElement("div");
  wrapper.className = campo.tipo === "especificacao" ? "campo-especificacao" : "campo-parametro";

  const label = document.createElement("label");
  label.id = `${campo.id}-label`;
  label.textContent = campo.pergunta;
  wrapper.appendChild(label);

  if (campo.ajuda) {
    // .br-tooltip (gov.br Design System) é o BALÃO em si, escondido por
    // padrão via CSS (opacity:0/visibility:hidden) até o componente JS
    // deles marcar [data-show] — JS que este site não carrega. Por isso o
    // "?" precisa de um gatilho próprio (sempre visível), que só revela o
    // balão via CSS :hover/:focus (ver .ajuda-gatilho em css/estilo.css).
    const gatilho = document.createElement("span");
    gatilho.className = "ajuda-gatilho";
    gatilho.textContent = "?";
    gatilho.tabIndex = 0;
    gatilho.setAttribute("role", "button");
    gatilho.setAttribute("aria-label", "Ajuda");
    const balao = document.createElement("span");
    balao.className = "br-tooltip";
    balao.textContent = campo.ajuda;
    gatilho.appendChild(balao);
    wrapper.appendChild(gatilho);
  }

  if (campo.avisoLimiar) {
    const aviso = document.createElement("div");
    aviso.className = "aviso-limiar";
    wrapper.appendChild(aviso);
    const atualizarAviso = () => {
      const campoValorEl = document.getElementById(campo.avisoLimiar.campoValor);
      const valorRef = campoValorEl ? Number(String(campoValorEl.value || "").replace(/[^\d,.-]/g, "").replace(",", ".")) : NaN;
      if (Number.isNaN(valorRef) || valorRef === 0) { aviso.textContent = ""; aviso.style.display = "none"; return; }
      aviso.style.display = "block";
      aviso.textContent = valorRef > campo.avisoLimiar.limite
        ? campo.avisoLimiar.mensagemAcima
        : campo.avisoLimiar.mensagemAbaixo;
    };
    setTimeout(atualizarAviso, 0);
    document.addEventListener("input", (e) => {
      if (e.target && e.target.id === campo.avisoLimiar.campoValor) atualizarAviso();
    });
  }

  if (campo.input === "tabela-linhas") {
    const tabelaFragmento = criarCampoTabelaLinhas(campo, valorAtual, aoMudar);
    tabelaFragmento.querySelector("table").setAttribute("aria-labelledby", label.id);
    wrapper.appendChild(tabelaFragmento);
    return wrapper;
  }

  if (campo.input === "select" && campo.opcoes.length <= 4) {
    const grupo = document.createElement("div");
    grupo.className = "opcoes-pill";
    grupo.id = campo.id;
    grupo.setAttribute("role", "group");
    grupo.setAttribute("aria-labelledby", label.id);
    // Antes do elaborador clicar em algo, `valorAtual` é undefined — mas se
    // o campo tem `padrao`, campos dependentes (mostrarSe) já o tratam como
    // respondido com esse valor (ver valorEfetivo em modelo.js), revelando
    // conteúdo condicionado a ele. Sem isto, nenhum botão aparecia marcado
    // mesmo com um campo dependente já visível — parecia que o campo
    // dependente tinha surgido do nada, sem resposta correspondente.
    const valorConsiderado = valorAtual ?? campo.padrao;
    for (const opcao of campo.opcoes) {
      const botao = document.createElement("button");
      botao.type = "button";
      botao.className = "pill" + (valorConsiderado === opcao ? " ativo" : "");
      botao.textContent = opcao;
      botao.addEventListener("click", () => {
        for (const irmao of grupo.querySelectorAll(".pill")) {
          irmao.classList.remove("ativo");
        }
        botao.classList.add("ativo");
        aoMudar(campo.id, opcao);
      });
      grupo.appendChild(botao);
    }
    wrapper.appendChild(grupo);
    return wrapper;
  }

  let input;
  if (campo.input === "textarea") {
    input = document.createElement("textarea");
  } else if (campo.input === "select") {
    input = document.createElement("select");
    for (const opcao of campo.opcoes) {
      const optEl = document.createElement("option");
      optEl.value = opcao;
      optEl.textContent = opcao;
      input.appendChild(optEl);
    }
  } else {
    input = document.createElement("input");
    // "numero" usa type="text" + inputmode="decimal" em vez de type="number":
    // muitos desses campos são percentuais com casas decimais (ex.: 0,3%,
    // 0,08%) escritos em notação brasileira (vírgula) — um <input
    // type="number"> nativo rejeita silenciosamente a vírgula (o valor final
    // fica "", sem nenhum aviso ao usuário), enquanto o teclado numérico
    // móvel continua aparecendo com inputmode="decimal".
    input.type = "text";
    if (campo.input === "numero") input.inputMode = "decimal";
  }
  input.id = campo.id;
  label.htmlFor = campo.id;
  input.value = valorAtual ?? "";
  input.addEventListener("input", () => aoMudar(campo.id, input.value));
  wrapper.appendChild(input);
  return wrapper;
}

function criarCampoTabelaLinhas(campo, valorAtual, aoMudar) {
  const fragmento = document.createElement("div");
  fragmento.className = "campo-tabela-linhas";

  const linhas = Array.isArray(valorAtual) ? valorAtual.map((l) => ({ ...l })) : [];

  const tabela = document.createElement("table");
  tabela.className = "br-table tabela-linhas-editavel";
  const corpo = document.createElement("tbody");
  tabela.appendChild(corpo);
  fragmento.appendChild(tabela);

  function renderizarLinhas() {
    corpo.innerHTML = "";
    linhas.forEach((linha, indice) => {
      const tr = document.createElement("tr");
      for (const coluna of campo.colunas) {
        const td = document.createElement("td");
        let input;
        if (coluna.input === "select") {
          input = document.createElement("select");
          for (const opcao of coluna.opcoes) {
            const optEl = document.createElement("option");
            optEl.value = opcao;
            optEl.textContent = opcao;
            input.appendChild(optEl);
          }
          input.value = linha[coluna.id] ?? coluna.opcoes[0];
        } else {
          input = document.createElement("input");
          input.type = "text";
          input.value = linha[coluna.id] ?? "";
        }
        input.addEventListener("input", () => {
          linhas[indice] = { ...linhas[indice], [coluna.id]: input.value };
          aoMudar(campo.id, linhas.map((l) => ({ ...l })));
        });
        td.appendChild(input);
        tr.appendChild(td);
      }
      const tdRemover = document.createElement("td");
      const botaoRemover = document.createElement("button");
      botaoRemover.type = "button";
      botaoRemover.className = "br-button small";
      botaoRemover.textContent = "Remover";
      botaoRemover.addEventListener("click", () => {
        linhas.splice(indice, 1);
        aoMudar(campo.id, linhas.map((l) => ({ ...l })));
        renderizarLinhas();
      });
      tdRemover.appendChild(botaoRemover);
      tr.appendChild(tdRemover);
      corpo.appendChild(tr);
    });
  }
  renderizarLinhas();

  const botaoAdicionar = document.createElement("button");
  botaoAdicionar.type = "button";
  botaoAdicionar.className = "br-button secondary";
  botaoAdicionar.textContent = "Adicionar linha";
  botaoAdicionar.addEventListener("click", () => {
    const linhaVazia = {};
    for (const coluna of campo.colunas) {
      linhaVazia[coluna.id] = coluna.input === "select" ? coluna.opcoes[0] : "";
    }
    linhas.push(linhaVazia);
    aoMudar(campo.id, linhas.map((l) => ({ ...l })));
    renderizarLinhas();
  });
  fragmento.appendChild(botaoAdicionar);

  return fragmento;
}

export function renderizarSecoes(secoes, respostas, container, aoMudar) {
  container.innerHTML = "";
  for (const secao of secoes) {
    const bloco = document.createElement("section");
    bloco.id = `secao-${secao.id}`;
    bloco.className = "br-accordion card-secao";

    const titulo = document.createElement("h2");
    titulo.textContent = `${secao.id}. ${secao.titulo}`;
    bloco.appendChild(titulo);

    for (const campo of secao.campos) {
      if (!campoVisivel(campo, respostas, secoes)) continue;
      bloco.appendChild(criarCampoInput(campo, respostas[campo.id], aoMudar));
    }
    container.appendChild(bloco);
  }
}

export function atualizarProgresso(secoes, respostas, elementoBarra) {
  const { totalEspecificacao, preenchidos } = progresso(secoes, respostas);
  const percentual = totalEspecificacao === 0 ? 100 : Math.round((preenchidos / totalEspecificacao) * 100);
  const preenchimento = elementoBarra.querySelector(".barra-progresso-preenchimento");
  const texto = elementoBarra.querySelector(".barra-progresso-texto");
  const mensagem = `${preenchidos} de ${totalEspecificacao} campos de especificação preenchidos`;
  preenchimento.style.width = `${percentual}%`;
  texto.textContent = mensagem;
  elementoBarra.setAttribute("aria-valuenow", String(percentual));
  elementoBarra.setAttribute("aria-valuetext", mensagem);
}
