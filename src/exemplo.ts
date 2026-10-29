import { dia } from "./nucleo/datas";
import { ligarHierarquia } from "./nucleo/leitor";
import type { Cronograma, Tarefa } from "./nucleo/tipos";

// [edt, nome, início, término, início LB, término LB, %, crítica, predecessoras]
type L = [string, string, string, string, string, string, number, boolean, string?];

const LINHAS: L[] = [
  ["1", "Obra exemplo – edifício comercial", "2026-02-02", "2027-03-31", "2026-02-02", "2027-02-26", 0.48, true],
  ["1.1", "Assinatura do contrato", "2026-02-02", "2026-02-02", "2026-02-02", "2026-02-02", 1, true],
  ["1.2", "Mobilização e canteiro", "2026-02-03", "2026-03-13", "2026-02-03", "2026-03-06", 1, true, "2"],
  ["1.3", "Fundações", "2026-03-02", "2026-06-12", "2026-03-02", "2026-05-22", 1, true],
  ["1.3.1", "Estacas hélice contínua", "2026-03-02", "2026-04-24", "2026-03-02", "2026-04-10", 1, true, "3"],
  ["1.3.2", "Blocos e baldrames", "2026-04-13", "2026-06-12", "2026-04-06", "2026-05-22", 1, true, "5"],
  ["1.4", "Estrutura", "2026-05-18", "2026-11-27", "2026-05-04", "2026-10-30", 0.72, true],
  ["1.4.1", "Pilares e lajes – subsolo", "2026-05-18", "2026-07-10", "2026-05-04", "2026-06-19", 1, true, "6"],
  ["1.4.2", "Pilares e lajes – térreo ao 4º", "2026-07-06", "2026-10-02", "2026-06-15", "2026-09-04", 0.9, true, "8"],
  ["1.4.3", "Cobertura e casa de máquinas", "2026-09-28", "2026-11-27", "2026-08-31", "2026-10-30", 0.15, true, "9"],
  ["1.5", "Vedações e fachada", "2026-08-03", "2027-01-29", "2026-07-20", "2026-12-18", 0.35, false],
  ["1.5.1", "Alvenaria interna", "2026-08-03", "2026-11-13", "2026-07-20", "2026-10-30", 0.6, false, "9II"],
  ["1.5.2", "Fachada ventilada", "2026-10-05", "2027-01-29", "2026-09-14", "2026-12-18", 0.1, false, "12"],
  ["1.6", "Instalações", "2026-08-17", "2027-02-26", "2026-08-03", "2027-01-29", 0.3, true],
  ["1.6.1", "Elétrica e dados", "2026-08-17", "2027-01-29", "2026-08-03", "2027-01-15", 0.35, false, "12II"],
  ["1.6.2", "Hidráulica e incêndio", "2026-08-24", "2027-01-22", "2026-08-10", "2027-01-08", 0.3, false, "12II"],
  ["1.6.3", "Ar-condicionado central", "2026-10-19", "2027-02-26", "2026-10-05", "2027-01-29", 0.1, true, "10"],
  ["1.7", "Acabamentos", "2026-11-16", "2027-03-19", "2026-11-02", "2027-02-19", 0.05, true],
  ["1.7.1", "Pisos e revestimentos", "2026-11-16", "2027-02-19", "2026-11-02", "2027-01-29", 0.05, false, "12"],
  ["1.7.2", "Pintura e forros", "2027-01-11", "2027-03-19", "2026-12-14", "2027-02-19", 0, true, "17"],
  ["1.8", "Marcos", "2026-06-12", "2027-03-31", "2026-05-22", "2027-02-26", 0.25, false],
  ["1.8.1", "Fundações concluídas", "2026-06-12", "2026-06-12", "2026-05-22", "2026-05-22", 1, false, "6"],
  ["1.8.2", "Estrutura concluída", "2026-11-27", "2026-11-27", "2026-10-30", "2026-10-30", 0, true, "10"],
  ["1.8.3", "Vistoria do Corpo de Bombeiros", "2027-03-05", "2027-03-05", "2027-02-12", "2027-02-12", 0, false, "16"],
  ["1.8.4", "Entrega da obra", "2027-03-31", "2027-03-31", "2027-02-26", "2027-02-26", 0, true, "20"],
];

const d = (s: string) => {
  const [a, m, di] = s.split("-").map(Number);
  return dia(a, m, di);
};

export function cronogramaExemplo(): Cronograma {
  const tarefas: Tarefa[] = LINHAS.map(([edt, nome, i, t, li, lt, pct, critica, pred], uid) => ({
    uid,
    id: String(uid + 1),
    edt,
    nome,
    nivel: edt.split(".").length,
    inicio: d(i),
    termino: d(t),
    lbInicio: d(li),
    lbTermino: d(lt),
    pct,
    critica,
    marco: i === t,
    resumo: false,
    predecessoras: pred ? [{ id: pred.replace(/\D/g, ""), tipo: pred.endsWith("II") ? "II" : "TI" }] : undefined,
  }));
  return { origem: "Cronograma de exemplo", tarefas: ligarHierarquia(tarefas) };
}
