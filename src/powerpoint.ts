/**
 * Integração com o PowerPoint (Office.js): desenha as primitivas como formas nativas, agrupa cada página,
 * grava a configuração no próprio arquivo e localiza/atualiza cronogramas já inseridos.
 */
import type { Area, Config, Cronograma, Primitiva } from "./nucleo/tipos";

const PREFIXO = "Cronograma";
const CHAVE = (id: string) => `cronograma:${id}`;
const RE_NOME = /^Cronograma \[([a-z0-9]+)\] · (\d+)\//;

export interface PaginaRegistrada {
  slideId: string;
  dx: number; // posição do grupo em relação à área (para respeitar quando o usuário move o grupo)
  dy: number;
}

export interface Registro {
  id: string;
  config: Config;
  cronograma: Cronograma;
  area: Area;
  paginas: PaginaRegistrada[];
  atualizado: string;
}

export interface Selecao {
  slideId?: string;
  forma?: { id: string; nome: string; x: number; y: number; w: number; h: number };
  cronogramaId?: string;
}

export const suportado = (versao: string) => Office.context.requirements.isSetSupported("PowerPointApi", versao);

export function novoId(): string {
  return Math.random().toString(36).slice(2, 8);
}

// ---------------------------------------------------------------- registro no arquivo
export function lerRegistro(id: string): Registro | null {
  return (Office.context.document.settings.get(CHAVE(id)) as Registro | null) ?? null;
}

function salvarRegistro(reg: Registro): Promise<void> {
  Office.context.document.settings.set(CHAVE(reg.id), reg);
  return new Promise((ok, erro) =>
    Office.context.document.settings.saveAsync((r) => (r.status === Office.AsyncResultStatus.Succeeded ? ok() : erro(r.error))),
  );
}

// ---------------------------------------------------------------- consultas
export async function tamanhoSlide(): Promise<{ w: number; h: number }> {
  if (!suportado("1.10")) return { w: 960, h: 540 };
  return PowerPoint.run(async (ctx) => {
    const ps = ctx.presentation.pageSetup;
    ps.load("slideWidth,slideHeight");
    await ctx.sync();
    return { w: ps.slideWidth, h: ps.slideHeight };
  }).catch(() => ({ w: 960, h: 540 }));
}

export async function selecaoAtual(): Promise<Selecao> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items/id");
    const formas = ctx.presentation.getSelectedShapes();
    formas.load("items/id,items/name,items/left,items/top,items/width,items/height");
    await ctx.sync();
    const sel: Selecao = { slideId: slides.items[0]?.id };
    const f = formas.items[0];
    if (f) {
      sel.forma = { id: f.id, nome: f.name, x: f.left, y: f.top, w: f.width, h: f.height };
      sel.cronogramaId = RE_NOME.exec(f.name)?.[1];
    }
    return sel;
  });
}

// ---------------------------------------------------------------- desenho
function desenhar(shapes: PowerPoint.ShapeCollection, prims: Primitiva[], fonte: string): PowerPoint.Shape[] {
  const criadas: PowerPoint.Shape[] = [];
  const caixa = (x: number, y: number, w: number, h: number) => ({ left: x, top: y, width: Math.max(w, 0.1), height: Math.max(h, 0.1) });
  for (const p of prims) {
    let s: PowerPoint.Shape;
    switch (p.t) {
      case "ret":
        s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, caixa(p.x, p.y, p.w, p.h));
        s.fill.setSolidColor(p.cor);
        s.lineFormat.visible = false;
        break;
      case "losango":
        s = shapes.addGeometricShape(PowerPoint.GeometricShapeType.diamond, caixa(p.x, p.y, p.w, p.h));
        s.fill.setSolidColor(p.cheio ? p.cor : "#FFFFFF");
        s.lineFormat.color = p.cor;
        s.lineFormat.weight = 1.25;
        break;
      case "tri":
        s = shapes.addGeometricShape(
          p.cima ? PowerPoint.GeometricShapeType.triangle : PowerPoint.GeometricShapeType.flowChartMerge,
          caixa(p.x, p.y, p.w, p.h),
        );
        s.fill.setSolidColor(p.cor);
        s.lineFormat.visible = false;
        break;
      case "linha":
        s = shapes.addLine(PowerPoint.ConnectorType.straight, {
          left: Math.min(p.x1, p.x2),
          top: Math.min(p.y1, p.y2),
          width: Math.abs(p.x2 - p.x1),
          height: Math.abs(p.y2 - p.y1),
        });
        s.lineFormat.color = p.cor;
        s.lineFormat.weight = p.esp;
        if (p.tracejado) s.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.dash;
        break;
      case "texto": {
        s = shapes.addTextBox(p.texto, caixa(p.x, p.y, p.w, p.h));
        const tf = s.textFrame;
        tf.autoSizeSetting = PowerPoint.ShapeAutoSize.autoSizeNone;
        tf.wordWrap = !!p.quebra;
        tf.leftMargin = tf.rightMargin = tf.topMargin = tf.bottomMargin = 0;
        tf.verticalAlignment = PowerPoint.TextVerticalAlignment.middle;
        const r = tf.textRange;
        r.font.name = fonte;
        r.font.size = p.tam;
        r.font.color = p.cor;
        r.font.bold = !!p.negrito;
        r.paragraphFormat.horizontalAlignment =
          p.alinh === "c" ? PowerPoint.ParagraphHorizontalAlignment.center
          : p.alinh === "r" ? PowerPoint.ParagraphHorizontalAlignment.right
          : PowerPoint.ParagraphHorizontalAlignment.left;
        break;
      }
    }
    criadas.push(s);
  }
  return criadas;
}

/** Desenha uma página no slide, agrupa e nomeia. Devolve a posição do grupo. */
async function desenharPagina(
  ctx: PowerPoint.RequestContext,
  slide: PowerPoint.Slide,
  prims: Primitiva[],
  reg: Registro,
  k: number,
  total: number,
): Promise<{ x: number; y: number }> {
  const nome = `${PREFIXO} [${reg.id}] · ${k + 1}/${total}`;
  const formas = desenhar(slide.shapes, prims, reg.config.paleta.fonte);
  if (suportado("1.8")) {
    formas.forEach((f) => f.load("id"));
    await ctx.sync();
    const grupo = slide.shapes.addGroup(formas.map((f) => f.id));
    grupo.name = nome;
    grupo.tags.add("CRONOGRAMA_ID", reg.id);
    grupo.tags.add("CRONOGRAMA_PAGINA", String(k));
    grupo.load("left,top");
    await ctx.sync();
    return { x: grupo.left, y: grupo.top };
  }
  // Sem agrupamento disponível: cada forma leva o nome do cronograma.
  formas.forEach((f) => (f.name = nome));
  await ctx.sync();
  const xs = prims.map((p) => ("x" in p ? p.x : Math.min(p.x1, p.x2)));
  const ys = prims.map((p) => ("y" in p ? p.y : Math.min(p.y1, p.y2)));
  return { x: Math.min(...xs), y: Math.min(...ys) };
}

async function preencherTitulo(ctx: PowerPoint.RequestContext, slide: PowerPoint.Slide, titulo: string) {
  if (!suportado("1.8")) return;
  const shapes = slide.shapes;
  shapes.load("items/type,items/id");
  await ctx.sync();
  const phs = shapes.items.filter((s) => s.type === "Placeholder");
  phs.forEach((s) => s.placeholderFormat.load("type"));
  await ctx.sync();
  const t = phs.find((s) => ["Title", "CenterTitle"].includes(s.placeholderFormat.type));
  if (t) t.textFrame.textRange.text = titulo;
}

/** Cria um slide com o mesmo layout logo depois de `depoisDe` e devolve-o. */
async function novoSlide(ctx: PowerPoint.RequestContext, modelo: PowerPoint.Slide, posicao: number): Promise<PowerPoint.Slide> {
  modelo.layout.load("id");
  modelo.slideMaster.load("id");
  await ctx.sync();
  ctx.presentation.slides.add({ layoutId: modelo.layout.id, slideMasterId: modelo.slideMaster.id });
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  const novo = slides.items[slides.items.length - 1];
  if (suportado("1.8")) {
    novo.moveTo(posicao);
    await ctx.sync();
  }
  return novo;
}

export interface Saida {
  paginas: Primitiva[][];
  titulo: string;
}

/** Insere um cronograma novo a partir do slide selecionado (páginas extras viram slides novos logo depois). */
export async function inserir(saida: Saida, reg: Registro, formaAreaId?: string): Promise<void> {
  await PowerPoint.run(async (ctx) => {
    const sel = ctx.presentation.getSelectedSlides();
    sel.load("items/id");
    await ctx.sync();
    if (!sel.items.length) throw new Error("Selecione um slide antes de inserir.");
    const base = ctx.presentation.slides.getItem(sel.items[0].id);
    let indice = 0;
    if (suportado("1.8")) {
      base.load("index");
      await ctx.sync();
      indice = base.index;
    }
    if (formaAreaId) base.shapes.getItem(formaAreaId).delete();
    reg.paginas = [];
    const total = saida.paginas.length;
    for (let k = 0; k < total; k++) {
      const slide = k === 0 ? base : await novoSlide(ctx, base, indice + k);
      if (k > 0) slide.load("id");
      const pos = await desenharPagina(ctx, slide, saida.paginas[k], reg, k, total);
      if (reg.config.preencherTitulo) await preencherTitulo(ctx, slide, saida.titulo + (total > 1 ? ` (${k + 1}/${total})` : ""));
      await ctx.sync();
      reg.paginas.push({ slideId: k === 0 ? sel.items[0].id : slide.id, dx: pos.x - reg.area.x, dy: pos.y - reg.area.y });
    }
  });
  reg.atualizado = new Date().toISOString();
  await salvarRegistro(reg);
}

/** Localiza os grupos de um cronograma em todos os slides: página -> { slide, forma }. */
async function localizar(ctx: PowerPoint.RequestContext, id: string) {
  const slides = ctx.presentation.slides;
  slides.load("items/id");
  await ctx.sync();
  slides.items.forEach((s) => s.shapes.load("items/id,items/name,items/left,items/top"));
  await ctx.sync();
  const achados = new Map<number, { slide: PowerPoint.Slide; formas: PowerPoint.Shape[] }>();
  for (const s of slides.items) {
    for (const f of s.shapes.items) {
      const m = RE_NOME.exec(f.name);
      if (!m || m[1] !== id) continue;
      const k = +m[2] - 1;
      if (!achados.has(k)) achados.set(k, { slide: s, formas: [] });
      achados.get(k)!.formas.push(f);
    }
  }
  return achados;
}

/**
 * Redesenha um cronograma existente com a nova configuração/dados. Cada página volta para o slide onde
 * estava, na posição para onde o usuário tenha movido o grupo. Páginas a mais viram slides novos.
 */
export async function atualizar(reg: Registro, gerar: (area: Area) => Saida): Promise<{ removidas: number; paginas: number }> {
  let total = 0;
  let removidas = 0;
  await PowerPoint.run(async (ctx) => {
    const achados = await localizar(ctx, reg.id);
    if (!achados.size) throw new Error("Não encontrei esse cronograma nos slides. Ele pode ter sido apagado ou desagrupado.");
    const antigas = [...achados.keys()].sort((a, b) => a - b);
    // A primeira página define o deslocamento da área (o usuário pode ter movido o grupo).
    const p0 = achados.get(antigas[0])!;
    const reg0 = reg.paginas[antigas[0]];
    if (reg0) {
      reg.area.x = Math.min(...p0.formas.map((f) => f.left)) - reg0.dx;
      reg.area.y = Math.min(...p0.formas.map((f) => f.top)) - reg0.dy;
    }
    for (const { formas } of achados.values()) formas.forEach((f) => f.delete());
    await ctx.sync();

    const saida = gerar(reg.area);
    total = saida.paginas.length;
    const novas: PaginaRegistrada[] = [];
    let ultimo = achados.get(antigas[antigas.length - 1])!.slide;
    for (let k = 0; k < total; k++) {
      let slide = achados.get(k)?.slide;
      if (!slide) {
        ultimo.load("index");
        await ctx.sync();
        slide = await novoSlide(ctx, ultimo, (ultimo.index ?? 0) + 1);
        slide.load("id");
        await ctx.sync();
      }
      ultimo = slide;
      const pos = await desenharPagina(ctx, slide, saida.paginas[k], reg, k, total);
      if (reg.config.preencherTitulo) await preencherTitulo(ctx, slide, saida.titulo + (total > 1 ? ` (${k + 1}/${total})` : ""));
      await ctx.sync();
      novas.push({ slideId: slide.id, dx: pos.x - reg.area.x, dy: pos.y - reg.area.y });
    }
    removidas = antigas.filter((k) => k >= total).length;
    reg.paginas = novas;
  });
  reg.atualizado = new Date().toISOString();
  await salvarRegistro(reg);
  return { removidas, paginas: total };
}
