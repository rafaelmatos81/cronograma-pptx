/** Controlador do painel lateral: dados -> configuração -> prévia -> inserir/atualizar no PowerPoint. */
import { cronogramaExemplo } from "./exemplo";
import { diaParaIso, hoje, isoParaDia } from "./nucleo/datas";
import { CAMPOS, type Campo, type Mapa } from "./nucleo/leitor";
import { PALETAS, type Resultado, areaPadrao, configPadrao, cortar, desenharPagina, preparar } from "./nucleo/layout";
import { type Fonte, abaSugerida, abrirArquivo, deTextoColado, tarefasDaAba } from "./nucleo/planilha";
import { Indice, chave, niveis } from "./nucleo/selecao";
import { paraSvg } from "./nucleo/svg";
import type { Area, Config, Cronograma, Modo, Paleta, Unidade } from "./nucleo/tipos";
import * as ppt from "./powerpoint";

const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const inp = (id: string) => $<HTMLInputElement>(id);
const sel = (id: string) => $<HTMLSelectElement>(id);

const est = {
  noPowerPoint: false,
  fonte: null as Fonte | null,
  aba: -1,
  mapa: {} as Mapa,
  crono: null as Cronograma | null,
  ind: null as Indice | null,
  cfg: configPadrao(),
  pagina: 0,
  slide: { w: 960, h: 540 },
  area: null as Area | null,
  formaAreaId: undefined as string | undefined,
  edicao: null as ppt.Registro | null,
  res: null as Resultado | null,
};

// ---------------------------------------------------------------- mensagens
function mensagem(texto: string, tipo: "" | "erro" | "ok" = "") {
  const m = $("mensagem");
  m.textContent = texto;
  m.className = `mensagem ${tipo}`;
}

async function tentar(fn: () => Promise<void> | void, ocupado?: HTMLButtonElement) {
  if (ocupado) ocupado.disabled = true;
  try {
    await fn();
  } catch (e) {
    console.error(e);
    mensagem(e instanceof Error ? e.message : String(e), "erro");
  } finally {
    if (ocupado) ocupado.disabled = false;
    atualizarRodape();
  }
}

// ---------------------------------------------------------------- preferências (só aparência, por usuário)
const PREFS = "cronograma:preferencias";
const CAMPOS_PREFS: (keyof Config)[] = [
  "mostrarLinhaBase", "mostrarDesvio", "mostrarPercentual", "mostrarTermino", "destacarCriticas", "mostrarLegenda",
  "mostrarVinculos", "umDiaComoMarco", "escala", "linhasPorPagina", "larguraNomes", "paleta", "preencherTitulo",
];
function lerPrefs() {
  try {
    const p = JSON.parse(localStorage.getItem(PREFS) ?? "{}");
    for (const k of CAMPOS_PREFS) if (k in p) (est.cfg as unknown as Record<string, unknown>)[k] = p[k];
  } catch { /* sem armazenamento: usa o padrão */ }
}
function salvarPrefs() {
  try {
    localStorage.setItem(PREFS, JSON.stringify(Object.fromEntries(CAMPOS_PREFS.map((k) => [k, est.cfg[k]]))));
  } catch { /* ignora */ }
}

// ---------------------------------------------------------------- dados
function definirCronograma(c: Cronograma) {
  est.crono = c;
  est.ind = new Indice(c.tarefas);
  est.pagina = 0;
  const lvs = new Set(c.tarefas.map((t) => t.nivel));
  if (est.cfg.nivelMin !== undefined && !lvs.has(est.cfg.nivelMin)) est.cfg.nivelMin = undefined;
  montarVisao();
  montarTarefas();
  resumoDados();
  agendarPrevia();
}

function resumoDados() {
  const c = est.crono;
  if (!c) return void ($("resumo-dados").textContent = "");
  const ts = c.tarefas;
  const partes = [
    `${ts.length} tarefas`,
    `${ts.filter((t) => t.marco).length} marcos`,
    ts.some((t) => t.lbTermino !== undefined) ? "com linha de base" : "sem linha de base",
    ts.some((t) => t.critica) ? "com caminho crítico" : "",
    ts.some((t) => t.predecessoras?.length) ? "com vínculos" : "",
  ].filter(Boolean);
  $("resumo-dados").innerHTML = `<strong>${c.origem}</strong><br>${partes.join(" · ")}`;
}

function carregarAba(i: number) {
  const f = est.fonte!;
  est.aba = i;
  const aba = f.abas[i];
  if (aba.cabecalho < 0) {
    est.mapa = {};
    montarColunas();
    est.crono = est.ind = null;
    resumoDados();
    agendarPrevia();
    throw new Error(`Na aba "${aba.nome}" não achei as colunas de nome, início e término.`);
  }
  est.mapa = { ...aba.mapa };
  montarColunas();
  recarregarTarefas();
}

function recarregarTarefas() {
  const f = est.fonte!;
  const aba = f.abas[est.aba];
  const tarefas = tarefasDaAba(aba, est.mapa);
  if (!tarefas.length) throw new Error("Nenhuma tarefa com datas válidas nas colunas escolhidas.");
  definirCronograma({ origem: f.abas.length > 1 ? `${f.arquivo} › ${aba.nome}` : f.arquivo, tarefas });
}

async function usarFonte(f: Fonte) {
  est.fonte = f;
  if (f.xml) {
    $("bloco-aba").hidden = $("bloco-colunas").hidden = true;
    if (!f.xml.length) throw new Error("O XML não tem tarefas.");
    definirCronograma({ origem: f.arquivo, tarefas: f.xml });
    mensagem("");
    return;
  }
  const s = sel("aba");
  s.innerHTML = f.abas.map((a, i) => `<option value="${i}">${a.nome}${a.cabecalho < 0 ? " (sem cronograma)" : ""}</option>`).join("");
  $("bloco-aba").hidden = f.abas.length < 2;
  const i = abaSugerida(f);
  if (i < 0) throw new Error("Não encontrei nenhuma aba com as colunas de nome, início e término.");
  s.value = String(i);
  carregarAba(i);
  mensagem("");
}

function montarColunas() {
  const f = est.fonte;
  const aba = f?.abas[est.aba];
  $("bloco-colunas").hidden = !aba || aba.cabecalho < 0;
  if (!aba || aba.cabecalho < 0) return;
  const cab = aba.linhas[aba.cabecalho] ?? [];
  const opcoes = ['<option value="">—</option>']
    .concat(cab.map((c, i) => (c === null || c === "" ? "" : `<option value="${i}">${String(c)}</option>`)))
    .join("");
  $("colunas").innerHTML = CAMPOS.map(
    ({ campo, rotulo, obrigatorio }) =>
      `<label for="col-${campo}" class="${obrigatorio ? "obrigatorio" : ""}">${rotulo}</label><select id="col-${campo}" data-campo="${campo}">${opcoes}</select>`,
  ).join("");
  for (const { campo } of CAMPOS) sel(`col-${campo}`).value = est.mapa[campo] !== undefined ? String(est.mapa[campo]) : "";
}

// ---------------------------------------------------------------- visão
function montarVisao() {
  const ind = est.ind;
  if (!ind) return;
  const lvs = [...new Set(ind.tarefas.map((t) => t.nivel))].sort((a, b) => a - b);
  const { min, max } = niveis(est.cfg, ind);
  const opts = lvs.map((n) => `<option value="${n}">${n}</option>`).join("");
  sel("nivel-min").innerHTML = opts;
  sel("nivel-max").innerHTML = opts;
  sel("agrupar").innerHTML = opts;
  sel("nivel-min").value = String(min);
  sel("nivel-max").value = String(Math.min(max, lvs[lvs.length - 1]));
  sel("agrupar").value = String(est.cfg.agruparNivel ?? min);
  montarRamos();
}

function montarRamos() {
  const ind = est.ind;
  if (!ind) return;
  const { min } = niveis(est.cfg, ind);
  const ramos = ind.tarefas.filter((t) => t.nivel === min);
  const marcados = new Set(est.cfg.ramos);
  $("ramos").innerHTML = ramos.length
    ? ramos
        .map((t) => {
          const k = chave(t);
          const ok = !marcados.size || marcados.has(k);
          return `<label title="${t.nome}"><input type="checkbox" data-ramo="${k}" ${ok ? "checked" : ""}/><span>${t.nome}</span><small>${t.edt}</small></label>`;
        })
        .join("")
    : '<p class="vazio">Sem tarefas neste nível.</p>';
}

function montarTarefas() {
  const ind = est.ind;
  if (!ind) return void ($("tarefas").innerHTML = "");
  const { min, max } = niveis(est.cfg, ind);
  const filtro = inp("tarefas-busca").value.trim().toLowerCase();
  const ocultos = new Set(est.cfg.ocultos);
  const lista = ind.tarefas.filter((t) => t.nivel >= min && t.nivel <= max && (!filtro || t.nome.toLowerCase().includes(filtro)));
  const LIMITE = 400;
  $("tarefas").innerHTML =
    lista
      .slice(0, LIMITE)
      .map((t) => {
        const k = chave(t);
        const recuo = (t.nivel - min) * 12;
        return `<label style="padding-left:${8 + recuo}px" title="${t.nome}"><input type="checkbox" data-tarefa="${k}" ${ocultos.has(k) ? "" : "checked"}/><span>${t.nome}</span><small>${t.edt}</small></label>`;
      })
      .join("") +
    (lista.length > LIMITE ? `<p class="vazio">Mostrando ${LIMITE} de ${lista.length}. Use a busca para achar as demais.</p>` : "") ||
    '<p class="vazio">Nenhuma tarefa.</p>';
  const b = $("mostrar-ocultas");
  b.hidden = !ocultos.size;
  b.textContent = `Mostrar as ${ocultos.size} ocultas`;
}

// ---------------------------------------------------------------- controles <-> configuração
function clarear(hex: string, f: number): string {
  const n = parseInt(hex.slice(1), 16);
  const c = [n >> 16, (n >> 8) & 255, n & 255].map((v) => Math.round(v + (255 - v) * f));
  return "#" + c.map((v) => v.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function lerControles() {
  const c = est.cfg;
  const num = (id: string) => (inp(id).value === "" ? null : Number(inp(id).value));
  c.modo = (document.querySelector<HTMLInputElement>('input[name="modo"]:checked')?.value ?? "lista") as Modo;
  if (est.ind) {
    c.nivelMin = Number(sel("nivel-min").value);
    c.nivelMax = Math.max(c.nivelMin, Number(sel("nivel-max").value));
    c.agruparNivel = Number(sel("agrupar").value);
  }
  c.apenasMarcos = inp("f-marcos").checked;
  c.apenasCriticas = inp("f-criticas").checked;
  c.apenasEmAndamento = inp("f-andamento").checked;
  c.apenasPendentes = inp("f-pendentes").checked;
  c.desvioMinimo = num("f-desvio");
  c.busca = inp("f-busca").value;
  c.periodoInicio = isoParaDia(inp("f-de").value);
  c.periodoFim = isoParaDia(inp("f-ate").value);
  c.manterHierarquia = inp("f-hierarquia").checked;
  c.preencherTitulo = inp("preencher-titulo").checked;
  c.titulo = inp("titulo").value;
  c.dataStatus = isoParaDia(inp("status").value);
  c.mostrarLinhaBase = inp("a-lb").checked;
  c.mostrarDesvio = inp("a-desvio").checked;
  c.mostrarPercentual = inp("a-pct").checked;
  c.mostrarTermino = inp("a-termino").checked;
  c.destacarCriticas = inp("a-criticas").checked;
  c.mostrarVinculos = inp("a-vinculos").checked;
  c.mostrarLegenda = inp("a-legenda").checked;
  c.umDiaComoMarco = inp("a-umdia").checked;
  c.escala = sel("escala").value as Unidade;
  c.linhasPorPagina = Math.min(40, Math.max(3, Number(inp("linhas").value) || 18));
  c.larguraNomes = Number(inp("largura").value);
  c.paleta = { ...c.paleta, fonte: sel("fonte").value };
}

function escreverControles() {
  const c = est.cfg;
  const marcar = (id: string, v: boolean) => (inp(id).checked = v);
  document.querySelectorAll<HTMLInputElement>('input[name="modo"]').forEach((r) => (r.checked = r.value === c.modo));
  $("bloco-agrupar").hidden = c.modo !== "faixas";
  marcar("f-marcos", c.apenasMarcos);
  marcar("f-criticas", c.apenasCriticas);
  marcar("f-andamento", c.apenasEmAndamento);
  marcar("f-pendentes", c.apenasPendentes);
  inp("f-desvio").value = c.desvioMinimo === null ? "" : String(c.desvioMinimo);
  inp("f-busca").value = c.busca;
  inp("f-de").value = diaParaIso(c.periodoInicio);
  inp("f-ate").value = diaParaIso(c.periodoFim);
  marcar("f-hierarquia", c.manterHierarquia);
  marcar("preencher-titulo", c.preencherTitulo);
  inp("titulo").value = c.titulo;
  inp("status").value = diaParaIso(c.dataStatus);
  marcar("a-lb", c.mostrarLinhaBase);
  marcar("a-desvio", c.mostrarDesvio);
  marcar("a-pct", c.mostrarPercentual);
  marcar("a-termino", c.mostrarTermino);
  marcar("a-criticas", c.destacarCriticas);
  marcar("a-vinculos", c.mostrarVinculos);
  marcar("a-legenda", c.mostrarLegenda);
  marcar("a-umdia", c.umDiaComoMarco);
  sel("escala").value = c.escala;
  inp("linhas").value = String(c.linhasPorPagina);
  inp("largura").value = String(c.larguraNomes);
  escreverPaleta();
  montarVisao();
  montarTarefas();
}

function escreverPaleta() {
  const p = est.cfg.paleta;
  const i = PALETAS.findIndex((x) => x.nome === p.nome);
  sel("paleta").innerHTML =
    PALETAS.map((x, j) => `<option value="${j}">${x.nome}</option>`).join("") + (i < 0 ? `<option value="-1">Personalizada</option>` : "");
  sel("paleta").value = String(i);
  inp("c-principal").value = p.principal.toLowerCase();
  inp("c-aexecutar").value = p.aExecutar.toLowerCase();
  inp("c-destaque").value = p.destaque.toLowerCase();
  inp("c-lb").value = p.linhaBase.toLowerCase();
  if (![...sel("fonte").options].some((o) => o.value === p.fonte)) sel("fonte").add(new Option(p.fonte, p.fonte));
  sel("fonte").value = p.fonte;
}

function trocarCor() {
  const p: Paleta = {
    ...est.cfg.paleta,
    nome: "Personalizada",
    principal: inp("c-principal").value.toUpperCase(),
    aExecutar: inp("c-aexecutar").value.toUpperCase(),
    destaque: inp("c-destaque").value.toUpperCase(),
    linhaBase: inp("c-lb").value.toUpperCase(),
  };
  p.destaqueClaro = clarear(p.destaque, 0.68);
  est.cfg.paleta = p;
  escreverPaleta();
}

// ---------------------------------------------------------------- prévia
let tempo = 0;
function agendarPrevia() {
  clearTimeout(tempo);
  tempo = window.setTimeout(previa, 90);
}

function areaAtual(): Area {
  return est.area ?? est.edicao?.area ?? areaPadrao(est.slide.w, est.slide.h);
}

function previa() {
  const caixa = $("previa-svg");
  const ind = est.ind;
  est.res = ind ? preparar(ind, est.cfg) : null;
  const res = est.res;
  if (!ind || !res) {
    caixa.innerHTML = `<p class="vazio">${ind ? "Nenhuma tarefa com esses filtros." : "Abra um cronograma ou use o exemplo."}</p>`;
    $("pag-txt").textContent = "—";
    $("previa-resumo").textContent = "";
    atualizarRodape();
    return;
  }
  est.pagina = Math.min(est.pagina, res.paginas.length - 1);
  const area = areaAtual();
  const prims = desenharPagina(res.paginas[est.pagina], res, est.cfg, ind, area);
  if (est.cfg.preencherTitulo) {
    const total = res.paginas.length;
    const t = res.titulo + (total > 1 ? ` (${est.pagina + 1}/${total})` : "");
    prims.unshift({ t: "texto", x: area.x, y: est.slide.h * 0.06, w: area.w, h: 30, texto: cortar(t, area.w, 22, true), tam: 22,
      cor: est.cfg.paleta.principal, negrito: true });
  }
  caixa.innerHTML = paraSvg(prims, est.slide.w, est.slide.h, est.cfg.paleta.fonte);
  $("pag-txt").textContent = `${est.pagina + 1}/${res.paginas.length}`;
  const n = res.selecionadas.length;
  $("previa-resumo").textContent = `${n} ${n === 1 ? "tarefa" : "tarefas"} · ${res.paginas.length} ${res.paginas.length === 1 ? "slide" : "slides"}`;
  atualizarRodape();
}

function atualizarRodape() {
  const pronto = !!est.res;
  const editando = !!est.edicao;
  $("btn-inserir").hidden = editando;
  $("btn-atualizar").hidden = $("btn-inserir-novo").hidden = !editando;
  $<HTMLButtonElement>("btn-inserir").disabled = !pronto || !est.noPowerPoint;
  $<HTMLButtonElement>("btn-atualizar").disabled = !pronto;
  $<HTMLButtonElement>("btn-inserir-novo").disabled = !pronto;
}

// ---------------------------------------------------------------- PowerPoint
function gerarSaida(area: Area): ppt.Saida {
  const res = preparar(est.ind!, est.cfg)!;
  return { titulo: res.titulo, paginas: res.paginas.map((pg) => desenharPagina(pg, res, est.cfg, est.ind!, area)) };
}

async function inserirNovo() {
  if (!est.res || !est.crono) return;
  const area = { ...areaAtual() };
  const saida = gerarSaida(area);
  const reg: ppt.Registro = {
    id: ppt.novoId(), config: structuredClone(est.cfg), cronograma: est.crono, area, paginas: [], atualizado: "",
  };
  mensagem(`Inserindo ${saida.paginas.length} ${saida.paginas.length === 1 ? "slide" : "slides"}…`);
  await ppt.inserir(saida, reg, est.formaAreaId);
  est.formaAreaId = undefined;
  est.area = null;
  $("area-txt").textContent = "";
  $("area-padrao").hidden = true;
  mensagem(
    saida.paginas.length > 1 ? `Cronograma inserido em ${saida.paginas.length} slides.` : "Cronograma inserido.",
    "ok",
  );
}

async function atualizarExistente() {
  const reg = est.edicao;
  if (!reg || !est.crono) return;
  reg.config = structuredClone(est.cfg);
  reg.cronograma = est.crono;
  mensagem("Atualizando…");
  const r = await ppt.atualizar(reg, gerarSaida);
  mensagem(
    `Cronograma atualizado (${r.paginas} ${r.paginas === 1 ? "slide" : "slides"}).` +
      (r.removidas ? ` ${r.removidas} ${r.removidas === 1 ? "slide ficou" : "slides ficaram"} sem cronograma e podem ser apagados.` : ""),
    "ok",
  );
}

async function aoMudarSelecao() {
  if (!est.noPowerPoint) return;
  const s = await ppt.selecaoAtual().catch(() => null);
  const faixa = $("faixa-edicao");
  if (est.edicao) return; // em edição, a faixa já mostra o estado
  if (s?.cronogramaId && ppt.lerRegistro(s.cronogramaId)) {
    faixa.hidden = false;
    faixa.dataset.id = s.cronogramaId;
    $("faixa-texto").textContent = "Cronograma selecionado no slide.";
    $("btn-editar").hidden = false;
    $("btn-cancelar-edicao").hidden = true;
  } else faixa.hidden = true;
}

function entrarEdicao(id: string) {
  const reg = ppt.lerRegistro(id);
  if (!reg) throw new Error("Não encontrei os dados desse cronograma no arquivo.");
  est.edicao = reg;
  est.cfg = { ...configPadrao(), ...structuredClone(reg.config) };
  est.fonte = null;
  $("bloco-aba").hidden = $("bloco-colunas").hidden = true;
  escreverControles();
  definirCronograma(reg.cronograma);
  $("faixa-texto").textContent = `Editando o cronograma de ${reg.cronograma.origem}. Para trocar os dados, abra o arquivo novo na aba Dados.`;
  $("btn-editar").hidden = true;
  $("btn-cancelar-edicao").hidden = false;
  mensagem("");
}

function sairEdicao() {
  est.edicao = null;
  $("faixa-edicao").hidden = true;
  agendarPrevia();
  void aoMudarSelecao();
}

// ---------------------------------------------------------------- eventos
function ligarEventos() {
  document.querySelectorAll<HTMLButtonElement>(".abas button").forEach((b) =>
    b.addEventListener("click", () => {
      document.querySelectorAll(".abas button").forEach((x) => x.classList.toggle("ativa", x === b));
      document.querySelectorAll<HTMLElement>("main section").forEach((s) => (s.hidden = s.dataset.painel !== b.dataset.aba));
      if (b.dataset.aba === "tarefas") montarTarefas();
    }),
  );

  inp("arquivo").addEventListener("change", () =>
    tentar(async () => {
      const f = inp("arquivo").files?.[0];
      if (!f) return;
      mensagem(`Lendo ${f.name}…`);
      await usarFonte(await abrirArquivo(f));
      inp("arquivo").value = "";
    }),
  );
  $("btn-exemplo").addEventListener("click", () =>
    tentar(() => {
      est.fonte = null;
      $("bloco-aba").hidden = $("bloco-colunas").hidden = true;
      if (!inp("status").value) inp("status").value = "2026-10-30";
      lerControles();
      definirCronograma(cronogramaExemplo());
    }),
  );
  $("btn-colar").addEventListener("click", () =>
    tentar(() => usarFonte(deTextoColado($<HTMLTextAreaElement>("colar").value))),
  );
  sel("aba").addEventListener("change", () => tentar(() => carregarAba(Number(sel("aba").value))));
  $("colunas").addEventListener("change", (e) =>
    tentar(() => {
      const s = e.target as HTMLSelectElement;
      const campo = s.dataset.campo as Campo;
      if (s.value === "") delete est.mapa[campo];
      else est.mapa[campo] = Number(s.value);
      recarregarTarefas();
    }),
  );

  // qualquer controle de Visão/Aparência
  const reagir = (e: Event) => {
    const alvo = e.target as HTMLElement;
    if (alvo.closest("#colunas, #tarefas-busca, #colar, #arquivo, #aba")) return;
    if (alvo.dataset.ramo !== undefined) return trocarRamo();
    if (alvo.dataset.tarefa !== undefined) return trocarOculto(alvo as HTMLInputElement);
    if (alvo.id === "paleta") {
      const i = Number(sel("paleta").value);
      if (i >= 0) est.cfg.paleta = { ...PALETAS[i] };
      escreverPaleta();
    } else if (alvo.id.startsWith("c-")) trocarCor();
    const antes = `${est.cfg.nivelMin}`;
    lerControles();
    $("bloco-agrupar").hidden = est.cfg.modo !== "faixas";
    if (`${est.cfg.nivelMin}` !== antes) {
      est.cfg.ramos = [];
      montarRamos();
    }
    if (alvo.id === "nivel-min" || alvo.id === "nivel-max") montarTarefas();
    salvarPrefs();
    agendarPrevia();
  };
  document.querySelector("main")!.addEventListener("change", reagir);
  document.querySelector("main")!.addEventListener("input", (e) => {
    const id = (e.target as HTMLElement).id;
    if (["f-busca", "titulo", "f-desvio", "largura", "linhas"].includes(id)) reagir(e);
  });
  inp("tarefas-busca").addEventListener("input", montarTarefas);
  $("mostrar-ocultas").addEventListener("click", () => {
    est.cfg.ocultos = [];
    montarTarefas();
    agendarPrevia();
  });
  $("ramos-todos").addEventListener("click", () => {
    est.cfg.ramos = [];
    montarRamos();
    agendarPrevia();
  });
  $("status-hoje").addEventListener("click", () => {
    inp("status").value = diaParaIso(hoje());
    lerControles();
    agendarPrevia();
  });

  $("pag-ant").addEventListener("click", () => {
    est.pagina = Math.max(0, est.pagina - 1);
    previa();
  });
  $("pag-prox").addEventListener("click", () => {
    est.pagina = Math.min((est.res?.paginas.length ?? 1) - 1, est.pagina + 1);
    previa();
  });

  $("area-selecao").addEventListener("click", () =>
    tentar(async () => {
      if (!est.noPowerPoint) throw new Error("Disponível só dentro do PowerPoint.");
      const s = await ppt.selecaoAtual();
      if (!s.forma) throw new Error("Selecione no slide a forma que marca a área do cronograma.");
      if (s.cronogramaId) throw new Error("A forma selecionada já é um cronograma. Use Editar.");
      est.area = { x: s.forma.x, y: s.forma.y, w: s.forma.w, h: s.forma.h };
      est.formaAreaId = s.forma.id;
      $("area-txt").textContent = `Área: "${s.forma.nome}" (${Math.round(s.forma.w)} × ${Math.round(s.forma.h)} pt). Ela será substituída pelo cronograma.`;
      $("area-padrao").hidden = false;
      agendarPrevia();
    }),
  );
  $("area-padrao").addEventListener("click", () => {
    est.area = null;
    est.formaAreaId = undefined;
    $("area-txt").textContent = "";
    $("area-padrao").hidden = true;
    agendarPrevia();
  });

  const bIns = $<HTMLButtonElement>("btn-inserir");
  bIns.addEventListener("click", () => tentar(inserirNovo, bIns));
  const bNovo = $<HTMLButtonElement>("btn-inserir-novo");
  bNovo.addEventListener("click", () => tentar(inserirNovo, bNovo));
  const bAtu = $<HTMLButtonElement>("btn-atualizar");
  bAtu.addEventListener("click", () => tentar(atualizarExistente, bAtu));
  $("btn-editar").addEventListener("click", () => tentar(() => entrarEdicao($("faixa-edicao").dataset.id!)));
  $("btn-cancelar-edicao").addEventListener("click", sairEdicao);
}

function trocarRamo() {
  const caixas = [...document.querySelectorAll<HTMLInputElement>("#ramos input[data-ramo]")];
  const marcados = caixas.filter((c) => c.checked).map((c) => c.dataset.ramo!);
  est.cfg.ramos = marcados.length === caixas.length ? [] : marcados;
  agendarPrevia();
}

function trocarOculto(c: HTMLInputElement) {
  const k = c.dataset.tarefa!;
  const ocultos = new Set(est.cfg.ocultos);
  if (c.checked) ocultos.delete(k);
  else ocultos.add(k);
  est.cfg.ocultos = [...ocultos];
  $("mostrar-ocultas").hidden = !ocultos.size;
  $("mostrar-ocultas").textContent = `Mostrar as ${ocultos.size} ocultas`;
  agendarPrevia();
}

// ---------------------------------------------------------------- início
async function iniciar(host: Office.HostType | null) {
  est.noPowerPoint = typeof Office !== "undefined" && host === Office.HostType.PowerPoint;
  $("ambiente").textContent = est.noPowerPoint ? "" : "modo prévia (fora do PowerPoint)";
  lerPrefs();
  escreverControles();
  ligarEventos();
  if (est.noPowerPoint) {
    if (!ppt.suportado("1.4")) mensagem("Esta versão do PowerPoint não permite desenhar formas por suplemento. Atualize o Office.", "erro");
    est.slide = await ppt.tamanhoSlide();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => void aoMudarSelecao());
    void aoMudarSelecao();
  }
  if (new URLSearchParams(location.search).has("exemplo")) $("btn-exemplo").click();
  const aba = location.hash.slice(1);
  if (aba) document.querySelector<HTMLButtonElement>(`.abas [data-aba="${aba}"]`)?.click();
  atualizarRodape();
}

if (typeof Office !== "undefined") Office.onReady((info) => void iniciar(info.host));
else void iniciar(null);
