export type MountType = "MONTADO_HS" | "MONTADO_TIMADEL" | "COMPLEMENTO_OBRA" | "REVENDA" | "MONTADO_ESTANCIA";
export type PackageStatus = "VALIDO" | "ATENCAO" | "INVALIDO";

export type AdditionalItem = {
  id: string;
  label: string;
  text: string;
  target: "ROMANEIO" | "ETIQUETA" | "AMBOS";
  position: "CABECALHO" | "ANTES_TABELA" | "OBSERVACAO" | "RODAPE";
  enabled: boolean;
};

export type LogisticsRule = {
  maxDoors: number;
  maxM3: number;
  maxTrimGames: number;
  maxFrameGames: number;
};

export type LogisticsConfig = {
  MONTADO_HS: LogisticsRule;
  MONTADO_TIMADEL: LogisticsRule;
  COMPLEMENTO_OBRA: LogisticsRule;
  REVENDA: LogisticsRule & { maxDoorsCardboard: number; maxDoorsThickness41: number };
  MONTADO_ESTANCIA: LogisticsRule;
  mixedMaxM3: number;
  additionalItems: AdditionalItem[];
  romaneioNote?: string;
  labelNote?: string;
};

export type OrderOptions = {
  mountType: MountType;
  motorista?: string;
  transportadora?: string;
  placa?: string;
  notaFiscal?: string;
  conferente?: string;
  separador?: string;
  romaneioExtraText?: string;
  etiquetaExtraText?: string;
};

export type PackageRow = {
  games?: number;
  quantity: number;
  lengthMm?: number;
  widthMm?: number;
  thicknessMm?: number;
  volume?: number;
  product: string;
  observation?: string;
  packaging?: "PLASTICO" | "PAPELAO" | "OUTRO";
  category?: "PORTA" | "MARCO" | "ALIZAR" | "FERRAGEM" | "KIT" | "OUTRO";
};

export type PackageData = {
  number: number;
  games?: number;
  totalVolume?: number;
  notes?: string;
  rows: PackageRow[];
  status?: PackageStatus;
  ruleApplied?: string;
  limitM3?: number;
  limitQuantity?: number;
  warnings?: string[];
};

export type ProcessingResult = {
  orderNumber: string;
  client: string;
  destination: string;
  delivery?: string;
  hasMachining: boolean;
  handSplit?: { right: number; left: number };
  mountType: MountType;
  mixedOrder?: boolean;
  packages: PackageData[];
  warnings: string[];
  config?: LogisticsConfig;
  orderOptions?: OrderOptions;
};
