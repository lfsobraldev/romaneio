import * as XLSX from "xlsx";
import type { ProcessingResult, PackageData, PackageRow, MountType, LogisticsConfig } from "./types";
import { applyLogistics } from "./logistics";
import { DEFAULT_LOGISTICS_CONFIG } from "./settings";

type PdfItem = {
  item: string;
  code: string;
  description: string;
  unit?: string;
  quantity: number;
  volume?: number;
};

type MachiningEntry = {
  quantity: number;
  unit?: string;
  description: string;
  lengthMm?: number;
  widthMm?: number;
  thicknessMm?: number;
  totalVolume?: number;
  notes: string[];
};

type HandInfo = { right: number; left: number; application?: string };

const clean = (s: string) => s.replace(/\s+/g, " ").trim();
const num = (value: unknown): number | undefined => {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return undefined;
  const s = value.trim();
  if (!s) return undefined;
  const normalized = s.includes(",") ? s.replace(/\./g, "").replace(",", ".") : s;
  const n = Number(normalized);
  return Number.isFinite(n) ? n : undefined;
};

function parseHand(text: string): HandInfo {
  const normalized = clean(text);
  const right = Number(normalized.match(/(\d+)\s*Direita(?:s)?\b/i)?.[1] || 0);
  const left = Number(normalized.match(/(\d+)\s*Esquerda(?:s)?\b/i)?.[1] || 0);
  const application = normalized.match(/(?:Direita(?:s)?|Esquerda(?:s)?)[^\-–—]*[\-–—]\s*(.+)$/i)?.[1];
  return { right, left, application: application ? clean(application) : undefined };
}

function parseMachining(buffer?: Buffer): MachiningEntry[] {
  if (!buffer?.length) return [];
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: "" }) as unknown[][];
  const entries: MachiningEntry[] = [];
  let current: MachiningEntry | undefined;
  let hardwareSection = false;

  for (const rawRow of rows) {
    const cells = rawRow.map(v => clean(String(v ?? "")));
    const qty = num(cells[0]);
    const unit = cells[1] || undefined;
    const desc = cells[2] || "";

    if (/FERRAGENS\s+PARA\s+USINAGEM/i.test(desc)) {
      hardwareSection = true;
      current = undefined;
      continue;
    }

    if (qty !== undefined && desc) {
      current = {
        quantity: Math.round(qty),
        unit,
        description: desc,
        lengthMm: num(cells[7]),
        widthMm: num(cells[8]),
        thicknessMm: num(cells[9]),
        totalVolume: num(cells[10]),
        notes: [],
      };
      entries.push(current);
      continue;
    }

    if (hardwareSection && desc) {
      entries.push({ quantity: 0, unit, description: desc, notes: [] });
      current = undefined;
      continue;
    }

    if (current && desc) current.notes.push(desc);
  }

  return entries;
}

function parsePdfItems(pdfText: string): PdfItem[] {
  const text = pdfText.replace(/\r/g, "");
  const lines = text.split("\n").map(clean).filter(Boolean);
  const items: PdfItem[] = [];
  const flat = clean(text);
  const re = /(?:^|\s)(\d+(?:\.\d+)?)\s+(\d{5,})\s+(.+?)\s+(CJ|PC|UN|JG|PÇ|PCS|PÇS)\s+(\d{1,6}[\.,]\d{3})\s+\d{1,6}[\.,]\d{3}\s+\d{1,9}(?:\.\d{3})*[\.,]\d{3}\s+(\d{1,6}[\.,]\d{3})(?=\s+\d+(?:\.\d+)?\s+\d{5,}|\s+Total Peso|$)/gi;

  for (const m of flat.matchAll(re)) {
    items.push({
      item: m[1], code: m[2], description: clean(m[3]), unit: m[4],
      quantity: Math.round(num(m[5]) || 0), volume: num(m[6]),
    });
  }
  if (items.length >= 5) return items;

  for (let i = 0; i < lines.length; i++) {
    const start = lines[i].match(/^(\d+(?:\.\d+)?)\s+(\d{5,})\s*(.*)$/);
    if (!start) continue;
    const block: string[] = [start[3]].filter(Boolean);
    let terminal: RegExpMatchArray | null = null;
    let j = i + 1;
    for (; j < Math.min(lines.length, i + 14); j++) {
      if (/^\d+(?:\.\d+)?\s+\d{5,}\b/.test(lines[j])) break;
      const t = lines[j].match(/^(CJ|PC|UN|JG|PÇ|PCS|PÇS)\s+(\d+[\.,]\d{3}).*?(\d+[\.,]\d{3})$/i);
      if (t) { terminal = t; break; }
      if (!/^(Página:|Pedido:|Cliente:|CNPJ:|Bairro:|Data Emissão:|Observação:|Item Código|Líquido Peso)/i.test(lines[j])) block.push(lines[j]);
    }
    if (!terminal) continue;
    items.push({ item: start[1], code: start[2], description: clean(block.join(" ")), unit: terminal[1], quantity: Math.round(num(terminal[2]) || 0), volume: num(terminal[3]) });
    i = j;
  }
  return items;
}

function dimsFromDescription(desc: string): [number, number, number] | undefined {
  const m = desc.match(/(\d{3,4})\s*[Xx]\s*(\d{2,4})\s*[Xx]\s*(\d{1,3})/);
  return m ? [Number(m[1]), Number(m[2]), Number(m[3])] : undefined;
}

const isDoor = (d: string) => /FOLHA\s+DE\s+PORTA/i.test(d);
const isFrame = (d: string) => /\b(?:MARCO|BATENTE)\b/i.test(d);
const isTrim = (d: string) => /\bALIZAR\b/i.test(d);
const isHinge = (d: string) => /\bDOBR(?:ADI[ÇC]A)?\b/i.test(d);
const isKitCorrer = (d: string) => /KIT\s+DE\s+CORRER/i.test(d);
const isBaguete = (d: string) => /BAGUETE/i.test(d);
const isPivotDoor = (d: string) => isDoor(d) && /PIVOTANTE/i.test(d);
const isSlidingDoor = (d: string) => isDoor(d) && /DE\s+CORRER/i.test(d);
const isSwingDoor = (d: string) => isDoor(d) && !isPivotDoor(d) && !isSlidingDoor(d);

function itemChildren(items: PdfItem[]) {
  const parentWithChildren = new Set<string>();
  for (const item of items) {
    const p = item.item.match(/^(\d+)\./)?.[1];
    if (p) parentWithChildren.add(p);
  }
  return items.filter(item => item.item.includes(".") || !parentWithChildren.has(item.item));
}

function parentItem(item: string) { return item.split(".")[0]; }

function sourceItemsText(items: PdfItem[]) {
  const ids = items.map(i => i.item).filter(Boolean);
  if (!ids.length) return undefined;
  return ids.length === 1 ? `Item ${ids[0]}` : `Itens ${ids.join(" / ")}`;
}

function humanDoorBase(desc: string) {
  let s = clean(desc)
    .replace(/\bHDF\s*\d*\b/gi, "")
    .replace(/\b\d{3,4}\s*[Xx]\s*\d{2,4}\s*[Xx]\s*\d{1,3}\b/g, "")
    .replace(/\bC\/USINAGEM\s+COMPLETA\b/gi, "")
    .replace(/\bC\/USI\.?\s*P\/PIVOTANTE\b/gi, "C/Usi. Pivotante")
    .replace(/\bC\/PLASTICO\b/gi, "")
    .replace(/\s+/g, " ").trim();
  s = s.replace(/^FOLHA DE PORTA/i, "Folha de Porta")
       .replace(/SOLIDA/gi, "Sólida")
       .replace(/IMPRESSO/gi, "Impresso");
  return s;
}

function genericDoorBase(pdfDoors: PdfItem[], entry: MachiningEntry) {
  const candidate = pdfDoors.find(i => {
    const d = dimsFromDescription(i.description);
    return d && d[0] === entry.lengthMm && d[1] === entry.widthMm && d[2] === entry.thicknessMm;
  });
  return candidate ? humanDoorBase(candidate.description) : clean(entry.description).replace(/\s*-\s*plast\b/i, "");
}

function withPlastic(s: string) {
  const value = clean(s);
  if (/PAPEL[AÃ]O/i.test(value)) return value.replace(/C\/?PAPEL[AÃ]O/gi, "(C/Papelão)").replace(/\s+/g, " ");
  return /\(C\/Plast\.\)|C\/PLASTICO|COM PLASTICO/i.test(value) ? value.replace(/C\/PLASTICO/gi, "(C/Plast.)") : `${value} (C/Plast.)`;
}

function appName(app?: string) {
  if (!app) return "";
  if (/eletr[oô]nica/i.test(app)) return "Eletrônica";
  if (/externa/i.test(app)) return "Externa";
  if (/\bwc\b|banheiro/i.test(app)) return "WC";
  return clean(app);
}

function frameHeadLength(entry: MachiningEntry) {
  const notes = entry.notes.join(" ");
  const m = notes.match(/(?:cab|travessa)\s*(?:de)?\s*(\d{3,4})/i);
  if (m) return Number(m[1]);
  if (entry.lengthMm && entry.widthMm) {
    // A planilha informa o comprimento total do conjunto (2 pernas + cabeceira).
    const inferred = entry.lengthMm - 2 * 2110;
    if (inferred > 300 && inferred < 1600) return inferred;
  }
  return undefined;
}

function rowVolume(row: PackageRow) {
  if (typeof row.volume === "number") return row.volume;
  if (!row.lengthMm || !row.widthMm || !row.thicknessMm || !row.quantity) return 0;
  return row.quantity * row.lengthMm * row.widthMm * row.thicknessMm / 1_000_000_000;
}

function pkgVolume(rows: PackageRow[]) { return rows.reduce((s, r) => s + rowVolume(r), 0); }

function splitRowsIntoPallets(rows: PackageRow[], capacity: number, startNumber: number): PackageData[] {
  const result: PackageData[] = [];
  let current: PackageRow[] = [];
  let used = 0;

  const flush = () => {
    if (!current.length) return;
    result.push({ number: startNumber + result.length, totalVolume: pkgVolume(current), rows: current });
    current = [];
    used = 0;
  };

  for (const original of rows) {
    let remaining = original.quantity;
    while (remaining > 0) {
      const room = capacity - used;
      if (room <= 0) flush();
      const take = Math.min(remaining, capacity - used);
      current.push({ ...original, quantity: take, volume: original.lengthMm && original.widthMm && original.thicknessMm ? undefined : original.volume });
      used += take;
      remaining -= take;
      if (used === capacity) flush();
    }
  }
  flush();
  return result;
}

function buildMachinedDoorRows(machining: MachiningEntry[], pdfItems: PdfItem[]): PackageRow[] {
  const pdfDoors = pdfItems.filter(i => isSwingDoor(i.description));
  const entries = machining.filter(e => /Porta/i.test(e.description) && parseHand(e.notes.join(" ")).right + parseHand(e.notes.join(" ")).left > 0);
  const atomic: Array<{ row: PackageRow; width: number; app: string; hand: "R"|"L" }> = [];

  for (const e of entries) {
    const h = parseHand(e.notes.join(" "));
    const app = appName(h.application);
    const base = genericDoorBase(pdfDoors, e);
    const make = (quantity: number, side: "Direita"|"Esquerda", hand: "R"|"L") => {
      if (!quantity) return;
      atomic.push({
        width: e.widthMm || 0, app, hand,
        row: {
          quantity,
          lengthMm: e.lengthMm,
          widthMm: e.widthMm,
          thicknessMm: e.thicknessMm,
          product: withPlastic(`${base} ${side}${app ? ` ${app}` : ""}`),
        },
      });
    };
    // A ordem abaixo reproduz o padrão real observado no romaneio aprovado:
    // Externa = direita/esquerda; WC e Eletrônica = esquerda/direita.
    if (/externa/i.test(app)) { make(h.right, "Direita", "R"); make(h.left, "Esquerda", "L"); }
    else { make(h.left, "Esquerda", "L"); make(h.right, "Direita", "R"); }
  }

  atomic.sort((a, b) => (b.width - a.width) || (a.app.localeCompare(b.app)));
  return atomic.map(a => a.row);
}

function buildFrameLegPackages(machining: MachiningEntry[], startNumber: number): PackageData[] {
  const frames = machining.filter(e => /Batente|Marco/i.test(e.description) && e.widthMm && (parseHand(e.notes.join(" ")).right + parseHand(e.notes.join(" ")).left > 0));
  const rightRows: PackageRow[] = [];
  const leftRows: PackageRow[] = [];
  const travRows: PackageRow[] = [];

  for (const e of frames) {
    const h = parseHand(e.notes.join(" "));
    const app = appName(h.application);
    const extra = /eletr[oô]nica/i.test(app) ? " Eletrônica" : "";
    const W = e.widthMm;
    const T = e.thicknessMm || 30;
    const L = 2110;

    if (h.right) {
      rightRows.push({ quantity: h.right, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet 2L Borr. Aplic. \"I\" 42x10 Dobradiça Direita${extra}`) });
      rightRows.push({ quantity: h.right, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet 2L Borr. Aplic. \"I\" 42x10 Contratesta Direita${extra}`) });
    }
    if (h.left) {
      leftRows.push({ quantity: h.left, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet 2L Borr. Aplic. \"I\" 42x10 Dobradiça Esquerda${extra}`) });
      leftRows.push({ quantity: h.left, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet 2L Borr. Aplic. \"I\" 42x10 Contratesta Esquerda${extra}`) });
    }

    const head = frameHeadLength(e);
    if (head) travRows.push({ quantity: e.quantity, lengthMm: head, widthMm: W, thicknessMm: T, product: withPlastic(`Travessa de Marco Pet 2L Borr. Aplic. \"I\" 42x10`) });
  }

  const packages: PackageData[] = [];
  if (rightRows.length) packages.push({ number: startNumber + packages.length, totalVolume: pkgVolume(rightRows), rows: rightRows });
  if (leftRows.length) packages.push({ number: startNumber + packages.length, totalVolume: pkgVolume(leftRows), rows: leftRows });
  if (travRows.length) packages.push({ number: startNumber + packages.length, totalVolume: pkgVolume(travRows), rows: travRows });
  return packages;
}

function buildSimpleDoorPackage(items: PdfItem[], number: number): PackageData | undefined {
  if (!items.length) return undefined;
  const rows: PackageRow[] = items.map(i => {
    const d = dimsFromDescription(i.description);
    return { quantity: i.quantity, lengthMm: d?.[0], widthMm: d?.[1], thicknessMm: d?.[2], volume: i.volume, product: withPlastic(humanDoorBase(i.description)), observation: sourceItemsText([i]) };
  });
  return { number, totalVolume: pkgVolume(rows), rows };
}

function frameParts(item: PdfItem): { legL?: number; headL?: number; width?: number; thickness?: number } {
  const d = dimsFromDescription(item.description);
  const head = item.description.match(/(?:1\s*PC|TRAVESSA\s+DE)\s*(\d{3,4})/i)?.[1];
  return { legL: d?.[0], width: d?.[1], thickness: d?.[2], headL: head ? Number(head) : undefined };
}

function buildUnmachinedFrames(items: PdfItem[], number: number, pivot = false): PackageData | undefined {
  if (!items.length) return undefined;
  const rows: PackageRow[] = [];
  for (const i of items) {
    const p = frameParts(i);
    rows.push({ games: i.quantity, quantity: i.quantity * 2, lengthMm: p.legL, widthMm: p.width, thicknessMm: p.thickness, product: withPlastic(`Conj. Marco Pet 2L S/Rebaixo`), observation: sourceItemsText([i]) });
    if (p.headL) rows.push({ quantity: i.quantity, lengthMm: p.headL, widthMm: p.width, thicknessMm: p.thickness, product: "" });
  }
  if (pivot) {
    // dobradiça pivotante é agregada ao mesmo pallet no modelo real.
    rows.push({ quantity: items.reduce((s, i) => s + i.quantity, 0), product: "Dobradiça Pivotante C/Esfera 150KG Cromada (Pino)" });
  }
  return { number, totalVolume: pkgVolume(rows), rows };
}

function buildTrimRows(items: PdfItem[]): PackageRow[] {
  const grouped = new Map<string, { games: number; len1: number; len2: number; width: number; thick: number; reg: string; items: string[] }>();
  for (const i of items) {
    const desc = i.description;
    const m1 = desc.match(/4\s*PCS\s+(\d{3,4})X(\d{2,4})X(\d{1,3})/i);
    const m2 = desc.match(/2\s*PCS\s+(\d{3,4})/i);
    if (!m1 || !m2) continue;
    const reg = /MAIOR\s*55\s*mm/i.test(desc) ? "55" : "35";
    const key = `${reg}-${m1[1]}-${m2[1]}-${m1[2]}-${m1[3]}`;
    const g = grouped.get(key) || { games: 0, len1: Number(m1[1]), len2: Number(m2[1]), width: Number(m1[2]), thick: Number(m1[3]), reg, items: [] };
    g.games += i.quantity;
    g.items.push(i.item);
    grouped.set(key, g);
  }

  const rows: PackageRow[] = [];
  for (const g of [...grouped.values()].sort((a,b) => Number(b.reg)-Number(a.reg))) {
    rows.push({ games: g.games, quantity: g.games * 2, lengthMm: g.len1, widthMm: g.width, thicknessMm: g.thick, product: withPlastic(`Conj. L Maior ${g.reg}mm MDF Ultra Pet 2L R.1`), observation: g.items.length ? `Itens ${g.items.join(" / ")}` : undefined });
    rows.push({ quantity: g.games, lengthMm: g.len2, widthMm: g.width, thicknessMm: g.thick, product: "" });
    rows.push({ quantity: g.games * 2, lengthMm: g.len1, widthMm: g.width, thicknessMm: g.thick, product: withPlastic(`Conj. L Menor MDF Ultra Pet 2L R.1`) });
    rows.push({ quantity: g.games, lengthMm: g.len2, widthMm: g.width, thicknessMm: g.thick, product: "" });
  }
  return rows;
}

function buildKitPackage(items: PdfItem[], number: number): PackageData | undefined {
  if (!items.length) return undefined;
  const rows: PackageRow[] = [];
  for (const i of items) {
    const q = i.quantity;
    rows.push({ games: q, quantity: q * 2, lengthMm: 2200, widthMm: 55, thicknessMm: 45, product: withPlastic(`Conj. Kit de Correr \"Modelo B\" Pet + Ferragens`), observation: sourceItemsText([i]) });
    rows.push({ quantity: q, lengthMm: 2200, widthMm: 55, thicknessMm: 45, product: "" });
    rows.push({ quantity: q, lengthMm: 2200, widthMm: 100, thicknessMm: 9, product: "" });
    rows.push({ quantity: q * 3, lengthMm: 2200, widthMm: 50, thicknessMm: 9, product: "" });
  }
  return { number, totalVolume: pkgVolume(rows), rows };
}

export function buildProcessing(pdfText: string, machiningBuffer?: Buffer, options?: { mountType?: MountType; config?: LogisticsConfig; mixedOrder?: boolean }): ProcessingResult {
  const flat = clean(pdfText);
  const orderNumber = flat.match(/Pedido:\s*(\d+)/i)?.[1] || "NÃO IDENTIFICADO";
  const rawClient = clean(flat.match(/Cliente:\s*(.+?)(?=CNPJ:|Endere[çc]o:|Bairro:|Data Emiss[aã]o:)/i)?.[1] || "Cliente não identificado");
  const client = rawClient.replace(/^\d+\s*[-–—]\s*/, "");
  const city = clean(flat.match(/Cidade:\s*(.+?)(?=Data Emiss[aã]o:|Data Previs[aã]o:|Observa[çc][aã]o:)/i)?.[1] || "");
  const destination = city;
  const delivery = flat.match(/(\d+\s*[º°ª]?\s*ENTREGA)/i)?.[1]?.replace(/\s+/g, " ");

  const allItems = parsePdfItems(pdfText);
  const items = itemChildren(allItems);
  const machining = parseMachining(machiningBuffer);
  const warnings: string[] = [];
  const packages: PackageData[] = [];

  // 1) Folhas usinadas de giro: a planilha de usinagem é a fonte da quantidade por mão/aplicação.
  const doorRows = buildMachinedDoorRows(machining, items);
  if (doorRows.length) packages.push({ number: 1, totalVolume: pkgVolume(doorRows), rows: doorRows });

  // 2) Pernas direita, pernas esquerda e travessas, separadas exatamente como o romaneio real.
  const framePkgs = buildFrameLegPackages(machining, packages.length + 1);
  packages.push(...framePkgs);

  // 3) Portas especiais do pedido (pivotante e correr).
  const pivotDoors = items.filter(i => isPivotDoor(i.description));
  const pivotPkg = buildSimpleDoorPackage(pivotDoors, packages.length + 1);
  if (pivotPkg) packages.push(pivotPkg);

  const slidingDoors = items.filter(i => isSlidingDoor(i.description));
  const slidingPkg = buildSimpleDoorPackage(slidingDoors, packages.length + 1);
  if (slidingPkg) packages.push(slidingPkg);

  // 4) Marcos sem rebaixo/sem usinagem: separa altura 2130 e 2330 como no modelo aprovado.
  const unmachinedFrames = items.filter(i => isFrame(i.description) && /S\/REB|S\/REBAIXO/i.test(i.description));
  const frames2130 = unmachinedFrames.filter(i => /2130X/i.test(i.description));
  const frames2330 = unmachinedFrames.filter(i => /2330X/i.test(i.description));
  const p2130 = buildUnmachinedFrames(frames2130, packages.length + 1, false); if (p2130) packages.push(p2130);
  const p2330 = buildUnmachinedFrames(frames2330, packages.length + 1, true); if (p2330) packages.push(p2330);

  // 5) Alizares: mantém jogos e cada comprimento em linha própria.
  const trims = items.filter(i => isTrim(i.description));
  const trimRows = buildTrimRows(trims);
  if (trimRows.length) {
    // Ainda existe variação física de embalagem entre obras; divide em blocos manejáveis para conferência.
    const chunkSize = 8; // duas famílias completas (4 linhas cada) por pallet visual.
    for (let i = 0; i < trimRows.length; i += chunkSize) {
      const rows = trimRows.slice(i, i + chunkSize);
      packages.push({ number: packages.length + 1, totalVolume: pkgVolume(rows), rows });
    }
  }

  // 6) Baguetes ficam junto dos últimos alizares quando houver; se não, pallet próprio.
  const baguetes = items.filter(i => isBaguete(i.description));
  if (baguetes.length) {
    const rows = baguetes.map(i => {
      const d = dimsFromDescription(i.description);
      return { quantity: i.quantity, lengthMm: d?.[0], widthMm: d?.[1], thicknessMm: d?.[2], volume: i.volume, product: withPlastic(`Baguetes P/Porta Pivotante Pinus Rec. Pet C/Borr. Aplic. \"T\" R.1`), observation: sourceItemsText([i]) } as PackageRow;
    });
    const lastTrim = packages.length ? packages[packages.length - 1] : undefined;
    if (lastTrim && lastTrim.rows.some(r => /Conj\. L/i.test(r.product))) {
      lastTrim.rows.push(...rows); lastTrim.totalVolume = pkgVolume(lastTrim.rows);
    } else packages.push({ number: packages.length + 1, totalVolume: pkgVolume(rows), rows });
  }

  // 7) Kit de correr.
  const kit = buildKitPackage(items.filter(i => isKitCorrer(i.description)), packages.length + 1);
  if (kit) packages.push(kit);

  // 8) Dobradiças comuns do pedido em um único pallet/volume zero.
  const hinges = items.filter(i => isHinge(i.description) && !/PIVOTANTE/i.test(i.description));
  if (hinges.length) {
    const qty = hinges.reduce((s, i) => s + i.quantity, 0);
    packages.push({ number: packages.length + 1, totalVolume: 0, rows: [{ quantity: qty, product: "Dobradiça Aço Inox 304 Famossul 3x2,5 R16 ESC", observation: sourceItemsText(hinges) }] });
  }

  if (!items.length) warnings.push("Não foi possível identificar automaticamente os itens do pedido.");
  if (machiningBuffer && !machining.length) warnings.push("A planilha de usinagem foi recebida, mas não foi possível ler suas linhas.");
  if (trimRows.length) warnings.push("Alizares foram estruturados por regulagem e medidas. A divisão física final dos pallets de alizar deve ser conferida, pois varia conforme a embalagem da obra.");

  const handRows = machining.map(e => parseHand(e.notes.join(" "))).filter(h => h.right || h.left);
  const handSplit = handRows.length ? {
    right: handRows.reduce((s,h) => s+h.right, 0),
    left: handRows.reduce((s,h) => s+h.left, 0),
  } : undefined;

  const base: ProcessingResult = {
    orderNumber, client, destination, delivery,
    hasMachining: Boolean(machiningBuffer?.length),
    handSplit,
    mountType: options?.mountType || "MONTADO_HS",
    mixedOrder: Boolean(options?.mixedOrder),
    packages,
    warnings,
    config: options?.config || DEFAULT_LOGISTICS_CONFIG,
  };
  return applyLogistics(base, base.mountType, base.config);
}
