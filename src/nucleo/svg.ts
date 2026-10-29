import { larguraTexto } from "./layout";
import type { Primitiva } from "./tipos";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function quebrar(texto: string, larg: number, pt: number, negrito: boolean): string[] {
  const linhas: string[] = [];
  let atual = "";
  for (const p of texto.split(/\s+/)) {
    const tentativa = atual ? `${atual} ${p}` : p;
    if (atual && larguraTexto(tentativa, pt, negrito) > larg) {
      linhas.push(atual);
      atual = p;
    } else atual = tentativa;
  }
  if (atual) linhas.push(atual);
  return linhas;
}

/** Desenha as primitivas num SVG do tamanho do slide (prévia do painel). */
export function paraSvg(prims: Primitiva[], largura: number, altura: number, fonte: string): string {
  const partes: string[] = [];
  for (const p of prims) {
    switch (p.t) {
      case "ret":
        partes.push(`<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" fill="${p.cor}"/>`);
        break;
      case "losango": {
        const cx = p.x + p.w / 2, cy = p.y + p.h / 2;
        partes.push(
          `<polygon points="${cx},${p.y} ${p.x + p.w},${cy} ${cx},${p.y + p.h} ${p.x},${cy}" fill="${p.cheio ? p.cor : "#FFFFFF"}" stroke="${p.cor}" stroke-width="1.25"/>`,
        );
        break;
      }
      case "tri": {
        const pts = p.cima
          ? `${p.x},${p.y + p.h} ${p.x + p.w},${p.y + p.h} ${p.x + p.w / 2},${p.y}`
          : `${p.x},${p.y} ${p.x + p.w},${p.y} ${p.x + p.w / 2},${p.y + p.h}`;
        partes.push(`<polygon points="${pts}" fill="${p.cor}"/>`);
        break;
      }
      case "linha":
        partes.push(
          `<line x1="${p.x1}" y1="${p.y1}" x2="${p.x2}" y2="${p.y2}" stroke="${p.cor}" stroke-width="${p.esp}"${p.tracejado ? ' stroke-dasharray="4 3"' : ""}/>`,
        );
        break;
      case "texto": {
        const al = p.alinh ?? "l";
        const x = al === "c" ? p.x + p.w / 2 : al === "r" ? p.x + p.w : p.x;
        const anchor = al === "c" ? "middle" : al === "r" ? "end" : "start";
        const linhas = p.quebra ? quebrar(p.texto, p.w, p.tam, !!p.negrito) : [p.texto];
        const alt = p.tam * 1.2;
        const y0 = p.y + p.h / 2 - ((linhas.length - 1) * alt) / 2;
        const tspans = linhas.map((l, i) => `<tspan x="${x}" y="${y0 + i * alt}">${esc(l)}</tspan>`).join("");
        partes.push(
          `<text font-size="${p.tam}" fill="${p.cor}" text-anchor="${anchor}" dominant-baseline="central"${p.negrito ? ' font-weight="700"' : ""}>${tspans}</text>`,
        );
        break;
      }
    }
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${largura} ${altura}" font-family="${esc(fonte)}, Manrope, Aptos, Arial, sans-serif"><rect width="${largura}" height="${altura}" fill="#FFFFFF"/>${partes.join("")}</svg>`;
}
