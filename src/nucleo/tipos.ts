/** Datas são guardadas como dias corridos desde 1970-01-01 (UTC): simples de somar e de salvar em JSON. */
export type Dia = number;

export type TipoVinculo = "TI" | "II" | "TT" | "IT";

export interface Vinculo {
  id: string; // Id da tarefa predecessora
  tipo: TipoVinculo;
}

export interface Tarefa {
  uid: number; // posição no arquivo original
  id: string;
  edt: string;
  nome: string;
  nivel: number;
  inicio: Dia;
  termino: Dia;
  lbInicio?: Dia;
  lbTermino?: Dia;
  pct?: number; // 0 a 1
  critica: boolean;
  marco: boolean;
  resumo: boolean; // tem filhas no arquivo original
  pai?: number; // uid do pai
  predecessoras?: Vinculo[];
}

export interface Cronograma {
  origem: string; // nome do arquivo (e aba)
  tarefas: Tarefa[];
}

export type Modo = "lista" | "faixas";
export type Unidade = "auto" | "semana" | "mes" | "trimestre";

export interface Paleta {
  nome: string;
  principal: string; // executado, cabeçalho, faixas, marcos
  aExecutar: string;
  destaque: string; // caminho crítico, desvio, data de status
  destaqueClaro: string;
  linhaBase: string;
  fundoResumo: string;
  fundoAlternado: string;
  grade: string;
  textoSecundario: string;
  rotulo: string;
  fonte: string;
}

export interface Config {
  // seleção
  modo: Modo;
  agruparNivel?: number;
  nivelMin?: number;
  nivelMax?: number;
  ramos: string[]; // chaves das tarefas do nível mínimo a incluir (vazio = todas)
  ocultos: string[]; // chaves de tarefas escondidas à mão
  busca: string;
  apenasMarcos: boolean;
  apenasCriticas: boolean;
  apenasEmAndamento: boolean;
  apenasPendentes: boolean;
  desvioMinimo: number | null;
  manterHierarquia: boolean;
  periodoInicio?: Dia;
  periodoFim?: Dia;
  // aparência
  dataStatus?: Dia;
  mostrarLinhaBase: boolean;
  mostrarDesvio: boolean;
  mostrarPercentual: boolean;
  mostrarTermino: boolean;
  destacarCriticas: boolean;
  mostrarLegenda: boolean;
  mostrarVinculos: boolean;
  umDiaComoMarco: boolean;
  escala: Unidade;
  linhasPorPagina: number;
  larguraNomes: number; // fração da largura da área (0,2 a 0,5)
  paleta: Paleta;
  // slide
  preencherTitulo: boolean;
  titulo: string; // vazio = automático
}

export interface Area {
  x: number;
  y: number;
  w: number;
  h: number;
}

/** Formas primitivas, em pontos, já posicionadas no slide. */
export type Primitiva =
  | { t: "ret"; x: number; y: number; w: number; h: number; cor: string }
  | { t: "losango"; x: number; y: number; w: number; h: number; cor: string; cheio: boolean }
  | { t: "tri"; x: number; y: number; w: number; h: number; cor: string; cima?: boolean } // ponta para baixo, salvo cima
  | { t: "linha"; x1: number; y1: number; x2: number; y2: number; cor: string; esp: number; tracejado?: boolean }
  | {
      t: "texto";
      x: number;
      y: number;
      w: number;
      h: number;
      texto: string;
      tam: number;
      cor: string;
      negrito?: boolean;
      alinh?: "l" | "c" | "r";
      quebra?: boolean;
    };
