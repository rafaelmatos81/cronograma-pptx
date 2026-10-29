import type { Dia } from "./tipos";

const MS_DIA = 86_400_000;

export const MESES = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"];

export function dia(ano: number, mes: number, d: number): Dia {
  return Math.round(Date.UTC(ano, mes - 1, d) / MS_DIA);
}

export function partes(d: Dia): { ano: number; mes: number; dia: number; semana: number } {
  const dt = new Date(d * MS_DIA);
  return { ano: dt.getUTCFullYear(), mes: dt.getUTCMonth() + 1, dia: dt.getUTCDate(), semana: (dt.getUTCDay() + 6) % 7 };
}

export function primeiroDoMes(d: Dia): Dia {
  const p = partes(d);
  return dia(p.ano, p.mes, 1);
}

export function proximoMes(d: Dia, n = 1): Dia {
  const p = partes(d);
  const m = p.mes - 1 + n;
  return dia(p.ano + Math.floor(m / 12), (m % 12) + 1, 1);
}

const dois = (n: number) => String(n).padStart(2, "0");

export function fdata(d: Dia | undefined, ano4 = false): string {
  if (d === undefined) return "—";
  const p = partes(d);
  return `${dois(p.dia)}/${dois(p.mes)}/${ano4 ? p.ano : dois(p.ano % 100)}`;
}

export function isoParaDia(s: string | undefined | null): Dia | undefined {
  if (!s) return undefined;
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  return m ? dia(+m[1], +m[2], +m[3]) : undefined;
}

export function diaParaIso(d: Dia | undefined): string {
  if (d === undefined) return "";
  const p = partes(d);
  return `${p.ano}-${dois(p.mes)}-${dois(p.dia)}`;
}

export function hoje(): Dia {
  const h = new Date();
  return dia(h.getFullYear(), h.getMonth() + 1, h.getDate());
}

/**
 * Converte o que vier da planilha: Date, número serial do Excel, "Qua 05/11/25", "05/11/2025",
 * "2025-11-05T08:00:00". Devolve undefined para "NA", "ND", "#ERRO" etc.
 */
export function lerData(v: unknown): Dia | undefined {
  if (v === null || v === undefined || v === "") return undefined;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return undefined;
    return dia(v.getFullYear(), v.getMonth() + 1, v.getDate());
  }
  if (typeof v === "number") {
    if (v > 20000 && v < 80000) return dia(1899, 12, 30) + Math.floor(v);
    return undefined;
  }
  const s = String(v).trim();
  const br = /(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.exec(s);
  if (br) {
    const a = +br[3];
    return dia(a < 100 ? a + 2000 : a, +br[2], +br[1]);
  }
  return isoParaDia(s);
}
