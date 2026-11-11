// Conversão de valor monetário para texto por extenso em português,
// portada de Gerador_PB.html (commit d7ecea4, projeto Projeto-Basico-App)
// para manter a mesma regra de arredondamento e de junção de grupos.

const UNIDADES_EXTENSO = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove", "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"];
const DEZENAS_EXTENSO = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];
const CENTENAS_EXTENSO = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"];
const ESCALAS_EXTENSO = [
  [1000000000, "bilhão", "bilhões"],
  [1000000, "milhão", "milhões"],
  [1000, "mil", "mil"],
];

function ate999PorExtenso(n) {
  if (n === 0) return "";
  if (n === 100) return "cem";
  const partes = [];
  const centena = Math.floor(n / 100);
  const resto = n % 100;
  if (centena) partes.push(CENTENAS_EXTENSO[centena]);
  if (resto) {
    if (resto < 20) {
      partes.push(UNIDADES_EXTENSO[resto]);
    } else {
      const dezena = Math.floor(resto / 10);
      const unidade = resto % 10;
      partes.push(unidade ? DEZENAS_EXTENSO[dezena] + " e " + UNIDADES_EXTENSO[unidade] : DEZENAS_EXTENSO[dezena]);
    }
  }
  return partes.join(" e ");
}

// Escreve por extenso um inteiro não negativo, aplicando as regras de junção
// do português: grupos separados por vírgula, exceto o último, que recebe
// 'e' quando for menor que cem ou uma centena redonda. Quando o resto de
// unidades (0-999) é zero, o "último grupo" para essa regra é o último
// grupo de escala não nulo (ex.: os 250 de "duzentos e cinquenta mil" em
// 6.250.000) — não o resto vazio, que sempre daria falso positivo para 'e'.
function inteiroPorExtenso(n) {
  if (n === 0) return "zero";
  const partes = [];
  let resto = n;
  let ultimaQuantidadeEscala = null;
  ESCALAS_EXTENSO.forEach(([valorEscala, singular, plural]) => {
    const quantidade = Math.floor(resto / valorEscala);
    resto = resto % valorEscala;
    if (quantidade === 0) return;
    ultimaQuantidadeEscala = quantidade;
    if (valorEscala === 1000) {
      partes.push(quantidade === 1 ? "mil" : ate999PorExtenso(quantidade) + " mil");
    } else {
      partes.push(ate999PorExtenso(quantidade) + " " + (quantidade === 1 ? singular : plural));
    }
  });
  let ultimoValor;
  if (resto) {
    ultimoValor = resto;
    partes.push(ate999PorExtenso(resto));
  } else {
    ultimoValor = ultimaQuantidadeEscala === null ? 0 : ultimaQuantidadeEscala;
  }
  if (partes.length === 1) return partes[0];
  const usaE = ultimoValor < 100 || (ultimoValor % 100 === 0 && ultimoValor < 1000);
  return partes.slice(0, -1).join(", ") + (usaE ? " e " : ", ") + partes[partes.length - 1];
}

// Converte para centavos inteiros com arredondamento meio-para-cima a partir
// dos dígitos decimais. Math.round(valor * 100) NÃO serve: 1.005 * 100 dá
// 100.49999999999999 em ponto flutuante e arredondaria para baixo.
function centavosTotais(valor) {
  let texto = String(valor === null || valor === undefined ? "" : valor).trim();
  // Notação brasileira (ex.: "6.250.000,00"): pontos são separador de
  // milhar, vírgula é decimal. Só reinterpreta dessa forma quando há
  // vírgula — "6250000.00" (sem vírgula) continua lido com ponto decimal.
  if (texto.includes(",")) texto = texto.replace(/\./g, "").replace(",", ".");
  const numero = Number(texto === "" ? "0" : texto);
  if (!isFinite(numero)) return 0;
  const absoluto = Math.abs(numero);
  const [inteiros, decimais] = absoluto.toFixed(10).split(".");
  let centavos = parseInt(inteiros, 10) * 100 + parseInt(decimais.slice(0, 2), 10);
  if (parseInt(decimais.charAt(2), 10) >= 5) centavos += 1;
  return centavos;
}

export function valorPorExtenso(valor) {
  const total = centavosTotais(valor);
  const reais = Math.floor(total / 100);
  const centavos = total % 100;
  const partes = [];
  if (reais) {
    const nome = reais === 1 ? "real" : "reais";
    const conector = (reais >= 1000000 && reais % 1000000 === 0) ? " de " : " ";
    partes.push(inteiroPorExtenso(reais) + conector + nome);
  }
  if (centavos) {
    partes.push(inteiroPorExtenso(centavos) + " " + (centavos === 1 ? "centavo" : "centavos"));
  }
  if (partes.length === 0) return "zero reais";
  return partes.join(" e ");
}
