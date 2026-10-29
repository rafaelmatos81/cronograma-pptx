// @vitest-environment happy-dom
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import * as XLSX from "xlsx";
import { describe, expect, it } from "vitest";
import { fdata, lerData } from "../src/nucleo/datas";
import { acharCabecalho, lerPredecessoras, lerProjectXml, linhasDeTexto, montar } from "../src/nucleo/leitor";
import { areaPadrao, configPadrao, desenharPagina, preparar } from "../src/nucleo/layout";
import { folhaParaLinhas } from "../src/nucleo/planilha";
import { Indice } from "../src/nucleo/selecao";
import { paraSvg } from "../src/nucleo/svg";
import type { Config } from "../src/nucleo/tipos";

describe("datas", () => {
  it("lê formatos do MS Project, serial do Excel e ISO", () => {
    expect(fdata(lerData("Qua 05/11/25"), true)).toBe("05/11/2025");
    expect(fdata(lerData(45966))).toBe("05/11/25");
    expect(fdata(lerData("2026-02-20T17:00:00"))).toBe("20/02/26");
    expect(lerData("NA")).toBeUndefined();
  });
});

describe("leitor", () => {
  it("monta hierarquia pela EDT e marcos pela duração", () => {
    const linhas = linhasDeTexto(
      "EDT;Nome da Tarefa;Duração;Início;Término;% concluída\n1;Obra;10 d;01/03/2026;20/03/2026;50%\n1.1;Escavação;5 d;01/03/2026;06/03/2026;100%\n1.2;Fim;0 d;20/03/2026;20/03/2026;0%",
    );
    const c = acharCabecalho(linhas)!;
    const ts = montar(linhas.slice(c.linha + 1), c.mapa);
    expect(ts.map((t) => t.nivel)).toEqual([1, 2, 2]);
    expect(ts[0].resumo).toBe(true);
    expect(ts[2].marco).toBe(true);
    expect(ts[1].pai).toBe(0);
    expect(ts[0].pct).toBeCloseTo(0.5);
  });

  it("lê predecessoras no formato do MS Project", () => {
    expect(lerPredecessoras("69TI+60 d;3II+15 dd;21")).toEqual([
      { id: "69", tipo: "TI" }, { id: "3", tipo: "II" }, { id: "21", tipo: "TI" },
    ]);
  });

  it("lê XML do MS Project com linha de base e vínculos", () => {
    const xml = `<?xml version="1.0"?><Project xmlns="http://schemas.microsoft.com/project"><Tasks>
      <Task><UID>0</UID><ID>0</ID><Name>P</Name><OutlineLevel>0</OutlineLevel><Start>2026-01-01T08:00:00</Start><Finish>2026-03-01T17:00:00</Finish></Task>
      <Task><UID>10</UID><ID>1</ID><Name>A</Name><WBS>1</WBS><OutlineLevel>1</OutlineLevel><Start>2026-01-05T08:00:00</Start><Finish>2026-02-20T17:00:00</Finish><PercentComplete>40</PercentComplete><Critical>1</Critical>
        <Baseline><Number>0</Number><Start>2026-01-05T08:00:00</Start><Finish>2026-02-10T17:00:00</Finish></Baseline></Task>
      <Task><UID>11</UID><ID>2</ID><Name>B</Name><WBS>2</WBS><OutlineLevel>1</OutlineLevel><Start>2026-02-21T08:00:00</Start><Finish>2026-02-21T08:00:00</Finish><Milestone>1</Milestone>
        <PredecessorLink><PredecessorUID>10</PredecessorUID><Type>1</Type></PredecessorLink></Task>
    </Tasks></Project>`;
    const ts = lerProjectXml(xml);
    expect(ts).toHaveLength(2);
    expect(ts[0].termino - ts[0].lbTermino!).toBe(10);
    expect(ts[0].critica).toBe(true);
    expect(ts[1].marco).toBe(true);
    expect(ts[1].predecessoras).toEqual([{ id: "1", tipo: "TI" }]);
  });
});

// ---- cronograma real (opcional, fora do git): CRONOGRAMA_TESTE=arquivo.xlsx [CRONOGRAMA_ABA=...] npm test
const ARQ = process.env.CRONOGRAMA_TESTE ?? "";
if (ARQ && existsSync(ARQ)) describe("cronograma real", () => {
  const wb = XLSX.read(readFileSync(ARQ), { type: "buffer", cellDates: false, dense: true });
  const aba = process.env.CRONOGRAMA_ABA ?? wb.SheetNames.find((n) => acharCabecalho(folhaParaLinhas(wb.Sheets[n]))) ?? "";
  const linhas = folhaParaLinhas(wb.Sheets[aba]);
  const c = acharCabecalho(linhas)!;
  const tarefas = montar(linhas.slice(c.linha + 1), c.mapa);
  const ind = new Indice(tarefas);
  const status = Math.max(...tarefas.map((t) => t.inicio));

  it("lê as tarefas", () => {
    expect(tarefas.length).toBeGreaterThan(0);
    for (const t of tarefas) expect(t.termino).toBeGreaterThanOrEqual(t.inicio);
  });

  const visoes: [string, Partial<Config>][] = [
    ["geral", {}],
    ["faixas", { modo: "faixas", umDiaComoMarco: true }],
    ["marcos", { nivelMax: 9, apenasMarcos: true }],
    ["desvios", { nivelMax: 9, desvioMinimo: 15 }],
    ["vinculos", { nivelMax: 9, mostrarVinculos: true }],
  ];
  mkdirSync("testes/saida", { recursive: true });
  for (const [nome, extra] of visoes) {
    it(`desenha a visão ${nome} dentro do slide`, () => {
      const cfg = { ...configPadrao(), dataStatus: status, ...extra } as Config;
      const res = preparar(ind, cfg);
      if (!res) return;
      res.paginas.forEach((pg, i) => {
        const prims = desenharPagina(pg, res, cfg, ind, areaPadrao());
        for (const p of prims) {
          if ("x" in p) {
            expect(p.x).toBeGreaterThanOrEqual(0);
            expect(p.x + p.w).toBeLessThanOrEqual(960.5);
            expect(p.y + p.h).toBeLessThanOrEqual(540);
          }
        }
        writeFileSync(`testes/saida/${nome}-${i + 1}.svg`, paraSvg(prims, 960, 540, cfg.paleta.fonte));
      });
    });
  }
});
