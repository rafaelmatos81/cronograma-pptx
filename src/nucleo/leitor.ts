/**
 * Leitura de cronogramas: planilha (linhas já extraídas do Excel/CSV/texto colado) ou XML do MS Project.
 * O cabeçalho é localizado sozinho e as colunas são reconhecidas pelo nome (português ou inglês).
 */
import { lerData } from "./datas";
import type { Tarefa, TipoVinculo, Vinculo } from "./tipos";

export type Campo =
  | "nome"
  | "inicio"
  | "termino"
  | "lbInicio"
  | "lbTermino"
  | "pct"
  | "edt"
  | "id"
  | "nivel"
  | "critica"
  | "duracao"
  | "marco"
  | "predecessoras";

export const CAMPOS: { campo: Campo; rotulo: string; obrigatorio?: boolean }[] = [
  { campo: "nome", rotulo: "Nome da tarefa", obrigatorio: true },
  { campo: "inicio", rotulo: "Início", obrigatorio: true },
  { campo: "termino", rotulo: "Término", obrigatorio: true },
  { campo: "lbInicio", rotulo: "Início da linha de base" },
  { campo: "lbTermino", rotulo: "Término da linha de base" },
  { campo: "pct", rotulo: "% concluída" },
  { campo: "edt", rotulo: "EDT" },
  { campo: "nivel", rotulo: "Nível" },
  { campo: "id", rotulo: "Id" },
  { campo: "critica", rotulo: "Crítica" },
  { campo: "duracao", rotulo: "Duração" },
  { campo: "marco", rotulo: "Marco" },
  { campo: "predecessoras", rotulo: "Predecessoras" },
];

const APELIDOS: Record<Campo, string[]> = {
  nome: ["nome da tarefa", "nome", "tarefa", "atividade", "descricao", "task name", "name"],
  inicio: ["inicio", "data de inicio", "data inicio", "start", "start date"],
  termino: ["termino", "fim", "data de termino", "data fim", "finish", "end", "finish date"],
  lbInicio: [
    "inicio da linha de base", "inicio da linha de base1", "inicio da linha de base 1", "inicio lb",
    "inicio original", "inicio previsto", "baseline start", "baseline1 start",
  ],
  lbTermino: [
    "termino da linha de base", "termino da linha de base1", "termino da linha de base 1", "termino lb",
    "termino original", "termino previsto", "baseline finish", "baseline1 finish",
  ],
  pct: ["% concluida", "% concluido", "percentual concluido", "% complete", "percent complete", "avanco", "% executado"],
  edt: ["edt", "eap", "wbs", "outline number"],
  id: ["id", "n", "item"],
  nivel: ["nivel da estrutura", "nivel", "outline level"],
  critica: ["critica", "critico", "critical"],
  duracao: ["duracao", "duration"],
  marco: ["marco", "milestone"],
  predecessoras: ["predecessoras", "predecessores", "predecessors"],
};

export type Mapa = Partial<Record<Campo, number>>;
export type Celula = string | number | boolean | Date | null | undefined;

export function normalizar(s: unknown): string {
  return String(s ?? "")
    .normalize("NFKD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

export function mapear(cabecalho: Celula[]): Mapa {
  const nomes = cabecalho.map(normalizar);
  const mapa: Mapa = {};
  for (const campo of Object.keys(APELIDOS) as Campo[]) {
    for (const a of APELIDOS[campo]) {
      const i = nomes.indexOf(a);
      if (i >= 0) {
        mapa[campo] = i;
        break;
      }
    }
  }
  return mapa;
}

/** Procura o cabeçalho nas 30 primeiras linhas: a primeira que tenha nome, início e término. */
export function acharCabecalho(linhas: Celula[][]): { linha: number; mapa: Mapa } | null {
  for (let i = 0; i < Math.min(30, linhas.length); i++) {
    const mapa = mapear(linhas[i] ?? []);
    if (mapa.nome !== undefined && mapa.inicio !== undefined && mapa.termino !== undefined) return { linha: i, mapa };
  }
  return null;
}

function lerPct(v: Celula): number | undefined {
  if (v === null || v === undefined || v === "") return undefined;
  if (typeof v === "string") {
    const n = parseFloat(v.replace("%", "").replace(",", ".").trim());
    return isNaN(n) ? undefined : n / 100;
  }
  const n = Number(v);
  if (isNaN(n)) return undefined;
  return n > 1.0001 ? n / 100 : n;
}

function lerBool(v: Celula): boolean {
  if (v === true) return true;
  return ["sim", "s", "yes", "y", "true", "verdadeiro", "1", "x"].includes(normalizar(v));
}

function marcoPorDuracao(v: Celula): boolean {
  if (v === null || v === undefined || v === "") return false;
  if (typeof v === "number") return v === 0;
  return /^0+([.,]0+)?\s*(d|dias?|ed|h|hr|days?|e?dias)?\??$/.test(normalizar(v));
}

/** "69TI+60 d;3TI+15 dd;21" -> vínculos (o atraso não é usado no desenho). */
export function lerPredecessoras(v: Celula): Vinculo[] | undefined {
  if (v === null || v === undefined || v === "") return undefined;
  const out: Vinculo[] = [];
  for (const parte of String(v).split(/[;,]/)) {
    const m = /^\s*(\d+)\s*(TI|II|TT|IT|FS|SS|FF|SF)?/i.exec(parte);
    if (!m) continue;
    const ing: Record<string, TipoVinculo> = { FS: "TI", SS: "II", FF: "TT", SF: "IT" };
    const t = (m[2] ?? "TI").toUpperCase();
    out.push({ id: m[1], tipo: (ing[t] ?? t) as TipoVinculo });
  }
  return out.length ? out : undefined;
}

export function ligarHierarquia(tarefas: Tarefa[]): Tarefa[] {
  const pilha: Tarefa[] = [];
  for (const t of tarefas) {
    while (pilha.length && pilha[pilha.length - 1].nivel >= t.nivel) pilha.pop();
    const pai = pilha[pilha.length - 1];
    t.pai = pai?.uid;
    if (pai) pai.resumo = true;
    pilha.push(t);
  }
  return tarefas;
}

/** Monta as tarefas a partir das linhas abaixo do cabeçalho. */
export function montar(linhas: Celula[][], mapa: Mapa): Tarefa[] {
  const tarefas: Tarefa[] = [];
  const g = (r: Celula[], c: Campo): Celula => (mapa[c] !== undefined ? r[mapa[c]!] : undefined);
  for (const r of linhas) {
    const bruto = g(r, "nome");
    if (bruto === null || bruto === undefined || String(bruto).trim() === "") continue;
    const nomeBruto = String(bruto);
    const edtV = g(r, "edt");
    const edt = edtV === null || edtV === undefined ? "" : String(edtV).trim();
    let nivel: number;
    const nivelV = g(r, "nivel");
    if (nivelV !== null && nivelV !== undefined && nivelV !== "" && !isNaN(Number(nivelV))) nivel = Math.trunc(Number(nivelV));
    else if (/^\d+(\.\d+)*$/.test(edt)) nivel = edt.split(".").length;
    else nivel = Math.floor((nomeBruto.length - nomeBruto.trimStart().length) / 3) + 1; // recuo do MS Project
    let inicio = lerData(g(r, "inicio"));
    let termino = lerData(g(r, "termino"));
    if (inicio === undefined && termino === undefined) continue;
    inicio ??= termino;
    termino ??= inicio;
    const t: Tarefa = {
      uid: tarefas.length,
      id: String(g(r, "id") ?? "").trim(),
      edt,
      nome: nomeBruto.trim(),
      nivel,
      inicio: inicio!,
      termino: Math.max(termino!, inicio!),
      lbInicio: lerData(g(r, "lbInicio")),
      lbTermino: lerData(g(r, "lbTermino")),
      pct: lerPct(g(r, "pct")),
      critica: lerBool(g(r, "critica")),
      marco: mapa.marco !== undefined ? lerBool(g(r, "marco")) : marcoPorDuracao(g(r, "duracao")),
      resumo: false,
      predecessoras: lerPredecessoras(g(r, "predecessoras")),
    };
    if (t.lbInicio !== undefined && t.lbTermino === undefined) t.lbTermino = t.lbInicio;
    if (t.lbTermino !== undefined && t.lbInicio === undefined) t.lbInicio = t.lbTermino;
    tarefas.push(t);
  }
  return ligarHierarquia(tarefas);
}

/** Texto colado do Excel (tabulações) ou CSV com ; ou , */
export function linhasDeTexto(texto: string): Celula[][] {
  const amostra = texto.slice(0, 4000);
  const sep = amostra.includes("\t") ? "\t" : (amostra.match(/;/g)?.length ?? 0) >= (amostra.match(/,/g)?.length ?? 0) ? ";" : ",";
  return texto
    .replace(/^﻿/, "")
    .split(/\r?\n/)
    .map((l) => dividirCsv(l, sep));
}

function dividirCsv(linha: string, sep: string): string[] {
  const out: string[] = [];
  let atual = "";
  let aspas = false;
  for (let i = 0; i < linha.length; i++) {
    const c = linha[i];
    if (c === '"') {
      if (aspas && linha[i + 1] === '"') {
        atual += '"';
        i++;
      } else aspas = !aspas;
    } else if (c === sep && !aspas) {
      out.push(atual);
      atual = "";
    } else atual += c;
  }
  out.push(atual);
  return out;
}

/** XML do MS Project (Arquivo › Salvar como › XML). */
export function lerProjectXml(xml: string): Tarefa[] {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("XML inválido.");
  // Compara nomes sem prefixo e sem diferenciar maiúsculas (exportações e analisadores variam).
  const nome = (el: Element) => (el.localName || el.tagName).split(":").pop()!.toLowerCase();
  const filho = (el: Element, tag: string) => {
    for (const c of Array.from(el.children)) if (nome(c) === tag.toLowerCase()) return (c.textContent ?? "").trim();
    return "";
  };
  const tarefas: Tarefa[] = [];
  const porUid = new Map<string, string>(); // UID -> ID (os vínculos do XML usam UID)
  const pendentes: { t: Tarefa; uids: { uid: string; tipo: string }[] }[] = [];
  const tasks = Array.from(doc.getElementsByTagName("*")).filter((e) => nome(e) === "task");
  for (const tk of tasks) {
    if (!filho(tk, "Name") || filho(tk, "IsNull") === "1" || filho(tk, "OutlineLevel") === "0") continue;
    const inicio = lerData(filho(tk, "Start"));
    if (inicio === undefined) continue;
    let lbI: number | undefined, lbT: number | undefined;
    const vinc: { uid: string; tipo: string }[] = [];
    for (const c of Array.from(tk.children)) {
      if (nome(c) === "baseline" && filho(c, "Number") === "0") {
        lbI = lerData(filho(c, "Start"));
        lbT = lerData(filho(c, "Finish"));
      }
      if (nome(c) === "predecessorlink") vinc.push({ uid: filho(c, "PredecessorUID"), tipo: filho(c, "Type") });
    }
    const pctTxt = filho(tk, "PercentComplete");
    const t: Tarefa = {
      uid: tarefas.length,
      id: filho(tk, "ID"),
      edt: filho(tk, "WBS") || filho(tk, "OutlineNumber"),
      nome: filho(tk, "Name"),
      nivel: parseInt(filho(tk, "OutlineLevel") || "1", 10),
      inicio,
      termino: Math.max(lerData(filho(tk, "Finish")) ?? inicio, inicio),
      lbInicio: lbI,
      lbTermino: lbT,
      pct: pctTxt ? parseFloat(pctTxt) / 100 : undefined,
      critica: filho(tk, "Critical") === "1",
      marco: filho(tk, "Milestone") === "1",
      resumo: false,
    };
    porUid.set(filho(tk, "UID"), t.id);
    if (vinc.length) pendentes.push({ t, uids: vinc });
    tarefas.push(t);
  }
  // Tipos do XML: 0 = TT, 1 = TI, 2 = IT, 3 = II
  const tipos: Record<string, TipoVinculo> = { "0": "TT", "1": "TI", "2": "IT", "3": "II" };
  for (const { t, uids } of pendentes) {
    t.predecessoras = uids
      .filter((v) => porUid.has(v.uid))
      .map((v) => ({ id: porUid.get(v.uid)!, tipo: tipos[v.tipo] ?? "TI" }));
  }
  return ligarHierarquia(tarefas);
}
