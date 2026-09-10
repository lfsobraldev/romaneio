export type PackageRow = {
  quantity: number;
  lengthMm?: number;
  widthMm?: number;
  thicknessMm?: number;
  volume?: number;
  product: string;
  observation?: string;
};

export type PackageData = {
  number: number;
  games?: number;
  totalVolume?: number;
  notes?: string;
  rows: PackageRow[];
};

export type ProcessingResult = {
  orderNumber: string;
  client: string;
  destination: string;
  delivery?: string;
  hasMachining: boolean;
  handSplit?: { right: number; left: number };
  packages: PackageData[];
  warnings: string[];
};
