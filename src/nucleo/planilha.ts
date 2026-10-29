/** Abre arquivos de cronograma no navegador (dentro do PowerPoint) e devolve as linhas de cada aba. */
import * as XLSX from "xlsx";
import { type Celula, type Mapa, acharCabecalho, linhasDeTexto, lerProjectXml, montar } from "./leitor";
import type { Tarefa } from "./tipos";

export interface Aba {
  nome: string;
  linhas: Celula[][];
  cabecalho: number; // índice da linha de cabeçalho (-1 se não achou)
  mapa: Mapa;
}

export interface Fonte {
  arquivo: string;
  abas: Aba[]; // vazio para XML do Project
  xml?: Tarefa[];
}

function prepararAba(nome: string, linhas: Celula[][]): Aba {
  const c = acharCabecalho(linhas);
  return { nome, linhas, cabecalho: c?.linha ?? -1, mapa: c?.mapa ?? {} };
}

export async function abrirArquivo(arquivo: File): Promise<Fonte> {
  const ext = arquivo.name.split(".").pop()?.toLowerCase() ?? "";
  if (ext === "mpp")
    throw new Error("O arquivo .mpp não pode ser lido direto. No MS Project, use Arquivo › Salvar como › XML e carregue o XML.");
  if (ext === "xml") return { arquivo: arquivo.name, abas: [], xml: lerProjectXml(await arquivo.text()) };
  if (ext === "csv" || ext === "txt") return { arquivo: arquivo.name, abas: [prepararAba(arquivo.name, linhasDeTexto(await arquivo.text()))] };
  // Datas chegam como número serial (raw), evitando erro de fuso horário.
  const wb = XLSX.read(await arquivo.arrayBuffer(), { type: "array", cellDates: false, dense: true });
  return { arquivo: arquivo.name, abas: wb.SheetNames.map((n) => prepararAba(n, folhaParaLinhas(wb.Sheets[n]))) };
}

export function folhaParaLinhas(ws: XLSX.WorkSheet): Celula[][] {
  return XLSX.utils.sheet_to_json<Celula[]>(ws, { header: 1, raw: true, defval: null, blankrows: true });
}

export function deTextoColado(texto: string): Fonte {
  return { arquivo: "dados colados", abas: [prepararAba("dados colados", linhasDeTexto(texto))] };
}

/** Aba sugerida: a que tem mais colunas reconhecidas (com nome, início e término). */
export function abaSugerida(f: Fonte): number {
  let melhor = -1, pontos = -1;
  f.abas.forEach((a, i) => {
    if (a.cabecalho < 0) return;
    const p = Object.keys(a.mapa).length * 10_000 + Math.min(a.linhas.length, 9_999);
    if (p > pontos) {
      pontos = p;
      melhor = i;
    }
  });
  return melhor;
}

export function tarefasDaAba(a: Aba, mapa: Mapa = a.mapa): Tarefa[] {
  if (a.cabecalho < 0) return [];
  return montar(a.linhas.slice(a.cabecalho + 1), mapa);
}
