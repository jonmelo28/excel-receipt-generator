function onlyDigitsCpf(cpf) {
  return String(cpf ?? "").replace(/\D/g, "").padStart(11, "0").slice(0, 11);
}

function formatCpf(cpf) {
  const d = onlyDigitsCpf(cpf);
  return `${d.slice(0,3)}.${d.slice(3,6)}.${d.slice(6,9)}-${d.slice(9,11)}`;
}

const UNIDADES = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"];
const DEZ_A_DEZENOVE = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"];
const DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"];
const CENTENAS = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"];

function trioPorExtenso(n) {
  n = Number(n);
  if (n === 0) return "";
  if (n === 100) return "cem";

  const c = Math.floor(n / 100);
  const d = Math.floor((n % 100) / 10);
  const u = n % 10;

  const parts = [];

  if (c) parts.push(CENTENAS[c]);

  if (d === 1) {
    parts.push(DEZ_A_DEZENOVE[u]);
  } else {
    if (d) parts.push(DEZENAS[d]);
    if (u) parts.push(UNIDADES[u]);
  }

  return parts.filter(Boolean).join(" e ");
}

function numeroPorExtenso(n) {
  n = Number(n);
  if (!Number.isFinite(n) || n < 0 || n > 999999999999) {
    throw new Error("Número fora do limite (0 a 999.999.999.999)");
  }
  if (n === 0) return "zero";

  const trilhoes = Math.floor(n / 1_000_000_000_000);
  const bilhoes  = Math.floor((n % 1_000_000_000_000) / 1_000_000_000);
  const milhoes  = Math.floor((n % 1_000_000_000) / 1_000_000);
  const milhares = Math.floor((n % 1_000_000) / 1_000);
  const resto    = n % 1_000;

  const parts = [];

  function pushGrupo(valor, singular, plural) {
    if (!valor) return;
    const t = trioPorExtenso(valor);
    parts.push(`${t} ${valor === 1 ? singular : plural}`.trim());
  }

  pushGrupo(trilhoes, "trilhão", "trilhões");
  pushGrupo(bilhoes, "bilhão", "bilhões");
  pushGrupo(milhoes, "milhão", "milhões");

  if (milhares) {
    if (milhares === 1) parts.push("mil");
    else parts.push(`${trioPorExtenso(milhares)} mil`);
  }

  if (resto) parts.push(trioPorExtenso(resto));

  // Ajuste de “e” entre grupos (pt-br)
  return parts.join(parts.length > 1 ? " e " : "");
}

function valorEmReaisPorExtenso(valor) {
  const v = Number(valor);
  if (!Number.isFinite(v) || v < 0) throw new Error("Valor inválido");

  const inteiro = Math.floor(v);
  const cent = Math.round((v - inteiro) * 100);

  const partes = [];

  if (inteiro === 0) {
    partes.push("zero real");
  } else {
    const ext = numeroPorExtenso(inteiro);
    partes.push(`${ext} ${inteiro === 1 ? "real" : "reais"}`);
  }

  if (cent > 0) {
    const extc = numeroPorExtenso(cent);
    partes.push(`${extc} ${cent === 1 ? "centavo" : "centavos"}`);
  }

  return partes.join(" e ");
}

module.exports = {
  formatCpf,
  valorEmReaisPorExtenso,
};