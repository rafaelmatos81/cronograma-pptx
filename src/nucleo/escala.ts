import { MESES, dia, partes, primeiroDoMes, proximoMes } from "./datas";
import type { Dia, Unidade } from "./tipos";

export interface Faixa {
  a: Dia;
  b: Dia;
  rotulo: string;
  novoAno?: boolean;
}

/** Converte datas em posição horizontal e monta as duas linhas do cabeçalho de tempo. */
export class Escala {
  readonly unidade: Exclude<Unidade, "auto">;
  readonly ini: Dia;
  readonly fim: Dia;
  readonly k: number; // pontos por dia

  constructor(ini: Dia, fim: Dia, readonly x0: number, readonly x1: number, unidade: Unidade = "auto") {
    const span = fim - ini + 1;
    if (unidade === "auto") {
      const largMes = (x1 - x0) / Math.max(span / 30.4, 1);
      unidade = span <= 120 ? "semana" : largMes >= 17 ? "mes" : "trimestre";
    }
    this.unidade = unidade;
    if (unidade === "semana") {
      this.ini = ini - partes(ini).semana;
      this.fim = fim + (7 - partes(fim).semana);
    } else if (unidade === "trimestre") {
      const p = partes(ini);
      this.ini = dia(p.ano, Math.floor((p.mes - 1) / 3) * 3 + 1, 1);
      const q = partes(fim);
      this.fim = proximoMes(dia(q.ano, Math.floor((q.mes - 1) / 3) * 3 + 1, 1), 3);
    } else {
      this.ini = primeiroDoMes(ini);
      this.fim = proximoMes(fim);
    }
    this.k = (x1 - x0) / (this.fim - this.ini);
  }

  x(d: Dia): number {
    return this.x0 + (Math.min(Math.max(d, this.ini), this.fim) - this.ini) * this.k;
  }

  private *meses(): Generator<[Dia, Dia]> {
    for (let d = primeiroDoMes(this.ini); d < this.fim; d = proximoMes(d)) yield [Math.max(d, this.ini), Math.min(proximoMes(d), this.fim)];
  }

  superior(): Faixa[] {
    if (this.unidade === "semana") {
      return [...this.meses()].map(([a, b]) => {
        const p = partes(a);
        return { a, b, rotulo: `${MESES[p.mes - 1]}/${String(p.ano % 100).padStart(2, "0")}` };
      });
    }
    const out: Faixa[] = [];
    for (let a = this.ini; a < this.fim; ) {
      const ano = partes(a).ano;
      const b = Math.min(dia(ano + 1, 1, 1), this.fim);
      out.push({ a, b, rotulo: String(ano) });
      a = b;
    }
    return out;
  }

  inferior(): Faixa[] {
    if (this.unidade === "semana") {
      const out: Faixa[] = [];
      for (let d = this.ini; d < this.fim; d += 7) out.push({ a: d, b: d + 7, rotulo: String(partes(d).dia).padStart(2, "0") });
      return out;
    }
    if (this.unidade === "trimestre") {
      const out: Faixa[] = [];
      for (let a = this.ini; a < this.fim; a = proximoMes(a, 3)) {
        const p = partes(a);
        out.push({ a, b: Math.min(proximoMes(a, 3), this.fim), rotulo: `T${Math.floor((p.mes - 1) / 3) + 1}`, novoAno: p.mes === 1 });
      }
      return out;
    }
    return [...this.meses()].map(([a, b]) => {
      const p = partes(a);
      const larg = (b - a) * this.k;
      return { a, b, rotulo: larg >= 26 ? MESES[p.mes - 1] : MESES[p.mes - 1][0].toUpperCase(), novoAno: p.mes === 1 };
    });
  }
}
