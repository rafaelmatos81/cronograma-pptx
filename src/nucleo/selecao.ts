import { normalizar } from "./leitor";
import type { Config, Tarefa } from "./tipos";

/** Identifica a tarefa de forma estável entre versões do arquivo (para filtros e ocultos salvos). */
export function chave(t: Tarefa): string {
  return t.edt || t.id || `${t.nivel}:${t.nome}`;
}

export function desvio(t: Tarefa): number | undefined {
  return t.lbTermino !== undefined ? t.termino - t.lbTermino : undefined;
}

export class Indice {
  readonly porUid = new Map<number, Tarefa>();
  readonly porId = new Map<string, Tarefa>();
  readonly nivelRaiz: number;
  readonly raizUnica: boolean;

  constructor(readonly tarefas: Tarefa[]) {
    for (const t of tarefas) {
      this.porUid.set(t.uid, t);
      if (t.id) this.porId.set(t.id, t);
    }
    this.nivelRaiz = tarefas.length ? Math.min(...tarefas.map((t) => t.nivel)) : 1;
    this.raizUnica = tarefas.filter((t) => t.nivel === this.nivelRaiz).length === 1;
  }

  pai(t: Tarefa): Tarefa | undefined {
    return t.pai !== undefined ? this.porUid.get(t.pai) : undefined;
  }

  ancestrais(t: Tarefa): Tarefa[] {
    const out: Tarefa[] = [];
    for (let p = this.pai(t); p; p = this.pai(p)) out.unshift(p);
    return out;
  }

  /** Nível mínimo padrão: pula o resumo do projeto quando há um único nível 1. */
  nivelMinPadrao(): number {
    return this.raizUnica && this.tarefas.length > 1 ? this.nivelRaiz + 1 : this.nivelRaiz;
  }
}

export function niveis(cfg: Config, ind: Indice): { min: number; max: number } {
  const min = cfg.nivelMin ?? ind.nivelMinPadrao();
  return { min, max: Math.max(min, cfg.nivelMax ?? min + 1) };
}

export function selecionar(ind: Indice, cfg: Config): Tarefa[] {
  const { min, max } = niveis(cfg, ind);
  let sel = ind.tarefas.filter((t) => t.nivel >= min && t.nivel <= max);

  if (cfg.ramos.length) {
    const ramos = new Set(cfg.ramos);
    sel = sel.filter((t) => ramos.has(chave(t)) || ind.ancestrais(t).some((a) => ramos.has(chave(a))));
  }
  if (cfg.busca.trim()) {
    const termos = cfg.busca.split(";").map(normalizar).filter(Boolean);
    sel = sel.filter((t) => termos.some((b) => normalizar(t.nome).includes(b)));
  }

  const filtros: ((t: Tarefa) => boolean)[] = [];
  if (cfg.apenasMarcos) filtros.push((t) => t.marco || (cfg.umDiaComoMarco && t.inicio === t.termino && !t.resumo));
  if (cfg.apenasCriticas) filtros.push((t) => t.critica);
  if (cfg.desvioMinimo !== null && cfg.desvioMinimo !== undefined)
    filtros.push((t) => (desvio(t) ?? -Infinity) >= cfg.desvioMinimo!);
  if (cfg.apenasEmAndamento) filtros.push((t) => t.pct !== undefined && t.pct > 0 && t.pct < 1);
  if (cfg.apenasPendentes) filtros.push((t) => t.pct === undefined || t.pct < 1);
  if (cfg.busca.trim() || filtros.length) {
    sel = sel.filter((t) => filtros.every((f) => f(t)));
    // Mantém os pais como contexto para a lista não ficar solta.
    if (cfg.manterHierarquia) {
      const ids = new Set(sel.map((t) => t.uid));
      for (const t of sel) for (const a of ind.ancestrais(t)) if (a.nivel >= min) ids.add(a.uid);
      sel = ind.tarefas.filter((t) => ids.has(t.uid));
    }
  }

  if (cfg.periodoInicio !== undefined || cfg.periodoFim !== undefined) {
    sel = sel.filter(
      (t) =>
        (cfg.periodoFim === undefined || t.inicio <= cfg.periodoFim) &&
        (cfg.periodoInicio === undefined || t.termino >= cfg.periodoInicio),
    );
  }

  if (cfg.ocultos.length) {
    const ocultos = new Set(cfg.ocultos);
    sel = sel.filter((t) => !ocultos.has(chave(t)) && !ind.ancestrais(t).some((a) => ocultos.has(chave(a))));
  }
  return sel;
}

// ---------------------------------------------------------------- paginação
export interface Linha {
  t: Tarefa;
  faixa?: Tarefa; // modo faixas: o grupo da linha
  cont?: boolean; // repetida como contexto ("cont.")
}

export function paginarLista(sel: Tarefa[], ind: Indice, porPagina: number): Linha[][] {
  const visiveis = new Set(sel.map((t) => t.uid));
  const paginas: Linha[][] = [];
  let atual: Linha[] = [];
  for (const t of sel) {
    if (atual.length >= porPagina) {
      paginas.push(atual);
      // Repete os pais da primeira linha da nova página, marcados como continuação.
      atual = ind
        .ancestrais(t)
        .filter((a) => visiveis.has(a.uid))
        .map((a) => ({ t: a, cont: true }));
    }
    atual.push({ t });
  }
  if (atual.length) paginas.push(atual);
  return paginas;
}

export function paginarFaixas(sel: Tarefa[], porPagina: number, nivelGrupo: number): Linha[][] {
  const faixas: { g: Tarefa; filhas: Tarefa[] }[] = [];
  for (const t of sel) {
    if (t.nivel === nivelGrupo) faixas.push({ g: t, filhas: [] });
    else if (faixas.length && t.nivel > nivelGrupo) faixas[faixas.length - 1].filhas.push(t);
  }
  const paginas: Linha[][] = [];
  let atual: Linha[] = [];
  for (const { g, filhas } of faixas) {
    const linhas: Linha[] = filhas.length ? filhas.map((f) => ({ t: f, faixa: g })) : [{ t: g, faixa: g }];
    // Não quebra uma faixa entre slides se ela couber inteira no próximo.
    if (atual.length && atual.length + linhas.length > porPagina && linhas.length <= porPagina) {
      paginas.push(atual);
      atual = [];
    }
    for (const ln of linhas) {
      if (atual.length >= porPagina) {
        paginas.push(atual);
        atual = [];
        ln.cont = true;
      }
      atual.push(ln);
    }
  }
  if (atual.length) paginas.push(atual);
  return paginas;
}

export function paginar(sel: Tarefa[], ind: Indice, cfg: Config): Linha[][] {
  if (!sel.length) return [];
  const porPagina = Math.max(3, cfg.linhasPorPagina);
  if (cfg.modo === "faixas") {
    const nivelGrupo = cfg.agruparNivel ?? Math.min(...sel.map((t) => t.nivel));
    return paginarFaixas(sel, porPagina, nivelGrupo);
  }
  return paginarLista(sel, ind, porPagina);
}
