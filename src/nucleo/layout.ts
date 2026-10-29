/**
 * Layout do Gantt: transforma uma página de linhas em formas primitivas (em pontos).
 * O mesmo resultado alimenta a prévia (SVG) e a inserção no PowerPoint.
 */
import { fdata } from "./datas";
import { Escala } from "./escala";
import { Indice, Linha, chave, desvio, niveis, paginar, selecionar } from "./selecao";
import type { Area, Config, Dia, Paleta, Primitiva, Tarefa } from "./tipos";

export const BRANCO = "#FFFFFF";

export const PALETAS: Paleta[] = [
  {
    nome: "Exxata", principal: "#09182B", aExecutar: "#C9D0D9", destaque: "#D51D07", destaqueClaro: "#F2B8AF",
    linhaBase: "#A7AEB8", fundoResumo: "#EEEEEE", fundoAlternado: "#F6F7F8", grade: "#D6D9DD",
    textoSecundario: "#5B6470", rotulo: "#A7AEB8", fonte: "Manrope",
  },
  {
    nome: "Azul Office", principal: "#156082", aExecutar: "#C0DCEB", destaque: "#C00000", destaqueClaro: "#F2B6B6",
    linhaBase: "#A6A6A6", fundoResumo: "#F2F2F2", fundoAlternado: "#F8F8F8", grade: "#D9D9D9",
    textoSecundario: "#595959", rotulo: "#8C8C8C", fonte: "Aptos",
  },
  {
    nome: "Grafite e laranja", principal: "#2F3437", aExecutar: "#D3D6D8", destaque: "#E46C0A", destaqueClaro: "#F7CFA8",
    linhaBase: "#9AA0A4", fundoResumo: "#EFF0F1", fundoAlternado: "#F7F7F8", grade: "#DADDDF",
    textoSecundario: "#5E6468", rotulo: "#9AA0A4", fonte: "Arial",
  },
];

export function configPadrao(): Config {
  return {
    modo: "lista", ramos: [], ocultos: [], busca: "",
    apenasMarcos: false, apenasCriticas: false, apenasEmAndamento: false, apenasPendentes: false,
    desvioMinimo: null, manterHierarquia: true,
    mostrarLinhaBase: true, mostrarDesvio: true, mostrarPercentual: true, mostrarTermino: true,
    destacarCriticas: true, mostrarLegenda: true, mostrarVinculos: false, umDiaComoMarco: false,
    escala: "auto", linhasPorPagina: 18, larguraNomes: 0.31, paleta: PALETAS[0],
    preencherTitulo: true, titulo: "",
  };
}

/** Área padrão do Gantt para um slide de largura x altura (pontos): abaixo da área de título. */
export function areaPadrao(largura = 960, altura = 540): Area {
  return { x: largura * 0.045, y: altura * 0.19, w: largura * 0.91, h: altura * 0.75 };
}

// ---------------------------------------------------------------- texto
export function larguraTexto(s: string, pt: number, negrito = false): number {
  let w = 0;
  for (const c of s) {
    if (c === " ") w += 0.27;
    else if (/[il.,;:|!'()]/.test(c)) w += 0.28;
    else if (/\d/.test(c)) w += 0.58;
    else if (c !== c.toLowerCase()) w += 0.68;
    else w += 0.54;
  }
  return w * pt * (negrito ? 1.06 : 1);
}

export function cortar(s: string, larg: number, pt: number, negrito = false): string {
  if (larguraTexto(s, pt, negrito) <= larg) return s;
  let r = s;
  while (r && larguraTexto(r + "…", pt, negrito) > larg) r = r.slice(0, -1);
  return r.replace(/[\s\-–,]+$/, "") + "…";
}

// ---------------------------------------------------------------- resultado
export interface Resultado {
  selecionadas: Tarefa[];
  paginas: Linha[][];
  ini: Dia;
  fim: Dia;
  titulo: string;
}

export function tituloAutomatico(sel: Tarefa[]): string {
  const fim = Math.max(...sel.map((t) => t.termino));
  const base = `Término previsto em ${fdata(fim, true)}`;
  const lbs = sel.filter((t) => t.lbTermino !== undefined).map((t) => t.lbTermino!);
  if (!lbs.length) return base;
  const d = fim - Math.max(...lbs);
  if (d > 0) return `${base}, ${d} dias após a linha de base`;
  if (d < 0) return `${base}, ${-d} dias antes da linha de base`;
  return `${base}, alinhado à linha de base`;
}

export function preparar(ind: Indice, cfg: Config): Resultado | null {
  const sel = selecionar(ind, cfg);
  if (!sel.length) return null;
  const datas: number[] = [];
  for (const t of sel) {
    datas.push(t.inicio, t.termino);
    if (cfg.mostrarLinhaBase && t.lbInicio !== undefined) datas.push(t.lbInicio, t.lbTermino!);
  }
  if (cfg.dataStatus !== undefined) datas.push(cfg.dataStatus);
  return {
    selecionadas: sel,
    paginas: paginar(sel, ind, cfg),
    ini: cfg.periodoInicio ?? Math.min(...datas),
    fim: cfg.periodoFim ?? Math.max(...datas),
    titulo: cfg.titulo.trim() || tituloAutomatico(sel),
  };
}

// ---------------------------------------------------------------- desenho
export function desenharPagina(
  linhas: Linha[],
  res: Resultado,
  cfg: Config,
  ind: Indice,
  A: Area,
): Primitiva[] {
  const P = cfg.paleta;
  const out: Primitiva[] = [];
  const ret = (x: number, y: number, w: number, h: number, cor: string) => out.push({ t: "ret", x, y, w, h, cor });
  const txt = (x: number, y: number, w: number, h: number, texto: string, tam: number, cor: string,
    o: { negrito?: boolean; alinh?: "l" | "c" | "r"; quebra?: boolean } = {}) =>
    out.push({ t: "texto", x, y, w, h, texto, tam, cor, ...o });
  const lin = (x1: number, y1: number, x2: number, y2: number, cor: string, esp: number, tracejado = false) =>
    out.push({ t: "linha", x1, y1, x2, y2, cor, esp, tracejado });

  const faixas = cfg.modo === "faixas";
  const status = cfg.dataStatus;
  const H_CAB = 30;
  const yCab = A.y;
  const yLin = yCab + H_CAB + 3;
  const yLeg = A.y + A.h - 14;
  const yFimMax = cfg.mostrarLegenda ? yLeg - 18 : A.y + A.h - 14;
  const direita = A.x + A.w;

  // colunas
  const wFaixa = faixas ? Math.min(97, A.w * 0.11) : 0;
  const xFaixa = A.x;
  const xNome = A.x + (wFaixa ? wFaixa + 6 : 0);
  const wTermino = cfg.mostrarTermino ? 40 : 0;
  const wDesvio = cfg.mostrarDesvio ? 30 : 0;
  const xInfo = direita - wTermino - wDesvio;
  const xGraf1 = xInfo - (wTermino || wDesvio ? 7 : 0);
  const xGraf0 = A.x + A.w * Math.min(Math.max(cfg.larguraNomes, 0.15), 0.6) + (faixas ? wFaixa * 0.6 : 0);
  const wPct = cfg.mostrarPercentual ? 30 : 0;
  const wNome = xGraf0 - 7 - wPct - xNome;
  const esc = new Escala(res.ini, res.fim, xGraf0, xGraf1, cfg.escala);

  const n = linhas.length;
  const rh = Math.min(24.5, (yFimMax - yLin) / Math.max(n, 8));
  const yFim = yLin + rh * n;
  const pt = rh >= 19.4 ? 8.5 : rh >= 15.8 ? 8 : 7;
  const nivelBase = Math.min(...linhas.map((l) => l.t.nivel));
  const naPagina = new Set(linhas.map((l) => l.t.uid));
  const temFilhaNaPagina = (t: Tarefa) => linhas.some((l) => l.t.pai === t.uid);

  // ---- cabeçalho de tempo
  const hh = H_CAB / 2;
  for (const f of esc.superior()) {
    const xa = esc.x(f.a), xb = esc.x(f.b);
    ret(xa, yCab, xb - xa, hh, P.principal);
    if (xb - xa > 18) txt(xa, yCab, xb - xa, hh, f.rotulo, 8, BRANCO, { negrito: true, alinh: "c" });
  }
  esc.inferior().forEach((f, i) => {
    const xa = esc.x(f.a), xb = esc.x(f.b);
    ret(xa, yCab + hh, xb - xa, hh, i % 2 === 0 ? P.fundoResumo : BRANCO);
    txt(xa, yCab + hh, xb - xa, hh, f.rotulo, xb - xa < 22 ? 7 : 7.5, P.textoSecundario, { alinh: "c" });
  });
  const cabecalhos: [number, string, number, "l" | "r"][] = [
    [xNome, "ATIVIDADE", wNome, "l"],
    [xGraf0 - 6 - wPct, wPct ? "%" : "", wPct, "r"],
    [xInfo, wTermino ? "TÉRMINO" : "", wTermino, "r"],
    [xInfo + wTermino, wDesvio ? "DESVIO" : "", wDesvio, "r"],
  ];
  for (const [x, s, w, al] of cabecalhos) if (s) txt(x, yCab + hh, w, hh, s, 7, P.rotulo, { negrito: true, alinh: al });
  if (faixas) txt(xFaixa, yCab + hh, wFaixa, hh, "GRUPO", 7, P.rotulo, { negrito: true });

  // ---- fundos e separadores
  const blocos: { g: Tarefa; i0: number; i1: number }[] = [];
  linhas.forEach((l, i) => {
    const y = yLin + i * rh;
    if (faixas) {
      const ult = blocos[blocos.length - 1];
      if (!ult || ult.g !== l.faixa) blocos.push({ g: l.faixa!, i0: i, i1: i + 1 });
      else ult.i1 = i + 1;
      if (blocos.length % 2 === 0) ret(xNome, y, xGraf1 - xNome, rh, P.fundoAlternado);
    } else if (temFilhaNaPagina(l.t)) {
      ret(xNome, y, direita - xNome, rh, P.fundoResumo);
    }
    lin(xNome, y + rh, direita, y + rh, P.grade, 0.5);
  });

  // ---- grade vertical
  for (const f of esc.inferior()) {
    const xa = esc.x(f.a);
    if (xa > esc.x0 + 0.5) {
      const ano = !!f.novoAno;
      lin(xa, yLin, xa, yFim, ano ? P.linhaBase : P.grade, ano ? 0.75 : 0.4);
    }
  }

  // ---- faixas (swimlanes)
  for (const { g, i0, i1 } of blocos) {
    const y0 = yLin + i0 * rh + 1, h = (i1 - i0) * rh - 2;
    const sozinho = i1 - i0 === 1 && linhas[i0].t === g;
    const w = sozinho ? xNome + wNome - xFaixa : wFaixa;
    ret(xFaixa, y0, w, h, P.principal);
    const ptF = h >= 29 || sozinho ? 8 : 7;
    const cabem = Math.max(1, Math.floor(h / (ptF * 1.3)));
    const cont = linhas.slice(i0, i1).some((l) => l.cont);
    const nome = g.nome + (cont ? " (cont.)" : "");
    const texto = cabem >= 3 ? nome : cortar(nome, (w - 8) * cabem * 0.92, ptF, true);
    txt(xFaixa + 4, y0, w - 8, h, texto, ptF, BRANCO, { negrito: true, quebra: !sozinho });
  }

  // ---- linhas
  const posicoes = new Map<number, { x0: number; x1: number; yc: number; yTopo: number; yBase: number }>();
  linhas.forEach((l, i) => {
    const t = l.t;
    const y = yLin + i * rh;
    const resumo = !faixas && temFilhaNaPagina(t);
    const negrito = resumo || (!faixas && t.nivel === nivelBase);
    const recuo = faixas ? 0 : (t.nivel - nivelBase) * 10;
    if (!(faixas && t === l.faixa)) {
      const nome = cortar(t.nome + (l.cont && !faixas ? " (cont.)" : ""), wNome - recuo - 3, pt, negrito);
      txt(xNome + recuo, y, wNome - recuo, rh, nome, pt, P.principal, { negrito });
    }
    if (wPct && t.pct !== undefined)
      txt(xGraf0 - 6 - wPct, y, wPct, rh, `${Math.round(t.pct * 100)}%`, pt - 0.5, P.textoSecundario, { alinh: "r" });
    if (wTermino) txt(xInfo, y, wTermino, rh, fdata(t.termino), pt - 0.5, P.textoSecundario, { alinh: "r" });
    const d = desvio(t);
    if (wDesvio && d)
      txt(xInfo + wTermino, y, wDesvio, rh, `${d > 0 ? "+" : "−"}${Math.abs(d)} d`, pt - 0.5,
        d > 0 ? P.destaque : P.textoSecundario, { negrito: d > 0, alinh: "r" });
    if (l.cont && !faixas) return;
    const marco = t.marco || (cfg.umDiaComoMarco && t.inicio === t.termino && !resumo);
    const pos = barra(out, esc, t, y, rh, resumo, marco, cfg);
    posicoes.set(t.uid, pos);
  });

  // ---- vínculos (só entre tarefas da mesma página)
  if (cfg.mostrarVinculos) {
    for (const l of linhas) {
      const suc = posicoes.get(l.t.uid);
      if (!suc || !l.t.predecessoras) continue;
      for (const v of l.t.predecessoras) {
        const p = ind.porId.get(v.id);
        if (!p || !naPagina.has(p.uid)) continue;
        const pre = posicoes.get(p.uid);
        if (!pre) continue;
        const cor = cfg.destacarCriticas && p.critica && l.t.critica ? P.destaque : P.textoSecundario;
        vinculo(out, pre, suc, v.tipo, cor, rh);
      }
    }
  }

  // ---- linha de status
  if (status !== undefined && status >= esc.ini && status <= esc.fim) {
    const xs = esc.x(status);
    lin(xs, yCab + H_CAB, xs, yFim, P.destaque, 1.25, true);
    txt(xs - 43, yFim + 1.5, 86, 13, `status ${fdata(status)}`, 7, P.destaque, { negrito: true, alinh: "c" });
  }

  if (cfg.mostrarLegenda) legenda(out, A.x, yLeg, cfg);
  return out;
}

function barra(out: Primitiva[], esc: Escala, t: Tarefa, y: number, rh: number, resumo: boolean, marco: boolean, cfg: Config) {
  const P = cfg.paleta;
  const crit = cfg.destacarCriticas && t.critica;
  const temLb = cfg.mostrarLinhaBase && t.lbInicio !== undefined && t.lbTermino !== undefined;
  const yc = y + rh / 2;
  if (marco) {
    const s = Math.min(rh * 0.62, 14.4);
    if (temLb && t.lbTermino !== t.termino) {
      const xb = esc.x(t.lbTermino!), sb = s * 0.75;
      out.push({ t: "losango", x: xb - sb / 2, y: yc - sb / 2, w: sb, h: sb, cor: P.linhaBase, cheio: false });
    }
    const xm = esc.x(t.termino);
    out.push({ t: "losango", x: xm - s / 2, y: yc - s / 2, w: s, h: s, cor: crit ? P.destaque : P.principal, cheio: (t.pct ?? 0) >= 1 });
    return { x0: xm - s / 2, x1: xm + s / 2, yc, yTopo: yc - s / 2, yBase: yc + s / 2 };
  }
  const x0 = esc.x(t.inicio), x1 = esc.x(t.termino + 1);
  const w = Math.max(x1 - x0, 2);
  let yb: number, h: number;
  if (resumo) {
    h = rh * 0.2;
    yb = yc - h / 2 - (temLb ? rh * 0.08 : 0);
    out.push({ t: "ret", x: x0, y: yb, w, h, cor: P.principal });
    for (const xt of [x0, x0 + w - 5]) out.push({ t: "tri", x: xt, y: yb + h - 0.4, w: 5, h: 4.3, cor: P.principal });
  } else {
    h = rh * (temLb ? 0.42 : 0.5);
    yb = yc - h / 2 - (temLb ? rh * 0.09 : 0);
    out.push({ t: "ret", x: x0, y: yb, w, h, cor: crit ? P.destaqueClaro : P.aExecutar });
    if (t.pct) out.push({ t: "ret", x: x0, y: yb, w: Math.max(w * Math.min(t.pct, 1), 0.7), h, cor: crit ? P.destaque : P.principal });
  }
  if (temLb) {
    const lx0 = esc.x(t.lbInicio!), lx1 = esc.x(t.lbTermino! + 1);
    out.push({ t: "ret", x: lx0, y: yb + h + 1.8, w: Math.max(lx1 - lx0, 2), h: Math.max(rh * 0.11, 1.8), cor: P.linhaBase });
  }
  return { x0, x1: x0 + w, yc, yTopo: yb, yBase: yb + h };
}

type Pos = { x0: number; x1: number; yc: number; yTopo: number; yBase: number };

/** Seta em cotovelo: sai da predecessora, desce/sobe e entra na sucessora pela borda. */
function vinculo(out: Primitiva[], pre: Pos, suc: Pos, tipo: string, cor: string, rh: number) {
  const ox = tipo === "II" || tipo === "IT" ? pre.x0 : pre.x1;
  const alvo = tipo === "TT" || tipo === "IT" ? suc.x1 - 3 : suc.x0 + 3;
  const desce = suc.yc > pre.yc;
  const yEntrada = desce ? suc.yTopo : suc.yBase;
  const seg = (x1: number, y1: number, x2: number, y2: number) => out.push({ t: "linha", x1, y1, x2, y2, cor, esp: 0.75 });
  if (alvo >= ox + 2) {
    seg(ox, pre.yc, alvo, pre.yc);
    seg(alvo, pre.yc, alvo, yEntrada + (desce ? -3 : 3));
  } else {
    const yMeio = desce ? pre.yc + rh / 2 : pre.yc - rh / 2;
    seg(ox, pre.yc, ox + 4, pre.yc);
    seg(ox + 4, pre.yc, ox + 4, yMeio);
    seg(ox + 4, yMeio, alvo, yMeio);
    seg(alvo, yMeio, alvo, yEntrada + (desce ? -3 : 3));
  }
  out.push({ t: "tri", x: alvo - 2.5, y: desce ? yEntrada - 3.5 : yEntrada, w: 5, h: 3.5, cor, cima: !desce });
}

function legenda(out: Primitiva[], x0: number, y: number, cfg: Config) {
  const P = cfg.paleta;
  const itens: [string, string, string][] = [
    ["barra", P.principal, "Executado"],
    ["barra", P.aExecutar, "A executar"],
  ];
  if (cfg.destacarCriticas) itens.push(["barra", P.destaque, "Caminho crítico"]);
  if (cfg.mostrarLinhaBase) itens.push(["lb", P.linhaBase, "Linha de base"]);
  itens.push(["marco", P.principal, "Marco concluído"], ["marcoVazio", P.principal, "Marco a cumprir"]);
  if (cfg.dataStatus !== undefined) itens.push(["status", P.destaque, "Data de status"]);
  let x = x0;
  for (const [tipo, cor, texto] of itens) {
    if (tipo === "barra") out.push({ t: "ret", x, y: y + 3.6, w: 20, h: 7.2, cor });
    else if (tipo === "lb") out.push({ t: "ret", x, y: y + 5.8, w: 20, h: 2.9, cor });
    else if (tipo.startsWith("marco")) out.push({ t: "losango", x, y: y + 2.5, w: 9.4, h: 9.4, cor, cheio: tipo === "marco" });
    else out.push({ t: "linha", x1: x + 10, y1: y, x2: x + 10, y2: y + 14.4, cor, esp: 1.25, tracejado: true });
    const larg = 5 * texto.length + 7;
    const dx = tipo.startsWith("marco") ? 12 : 24;
    out.push({ t: "texto", x: x + dx, y, w: larg, h: 14.4, texto, tam: 7.5, cor: P.textoSecundario });
    x += dx + larg + 13;
  }
}

export { chave, niveis };
