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
  const rawLines = pdfText.replace(/\r/g, "").split("\n");
  const lines = rawLines.map(clean);
  const found = new Map<string, PdfItem>();
  const put = (item: PdfItem) => {
    if (!item.item || !item.code || !item.quantity) return;
    const key = `${item.item}|${item.code}`;
    const prev = found.get(key);
    // Mantém a descrição mais completa quando duas estratégias encontram a mesma linha.
    if (!prev || item.description.length > prev.description.length) found.set(key, item);
  };

  // Estratégia 1: texto corrido retornado pelo pdf-parse.
  const flat = clean(pdfText);
  const record = /(?:^|\s)(\d+(?:\.\d+)?)\s+(\d{5,})\s+(.+?)\s+(CJ|PC|UN|JG|PÇ|PCS|PÇS)\s+(\d{1,6}[\.,]\d{3})\s+\d{1,6}[\.,]\d{3}\s+\d{1,9}(?:\.\d{3})*[\.,]\d{3}\s+(\d{1,6}[\.,]\d{3})(?=\s+\d+(?:\.\d+)?\s+\d{5,}|\s+Total Peso|$)/gi;
  for (const m of flat.matchAll(record)) put({ item:m[1], code:m[2], description:clean(m[3]), unit:m[4], quantity:Math.round(num(m[5])||0), volume:num(m[6]) });

  // Estratégia 2: parser tolerante ao layout do relatório. A descrição pode começar
  // antes da coluna Item/Código e continuar nas linhas seguintes.
  const header = /^(FAMOSSUL MADEIRAS|CNPJ:|R PAULO|ESTANCIA|Página:|Pedido:|Cliente:|Bairro:|Cidade:|Data Emiss|Data Previs|Observa|Item\s+Código|Peso Unit|Líquido|Total Peso)/i;
  const itemLine = /^(\d+(?:\.\d+)?)\s+(\d{5,})\s+(.*)$/;
  const terminal = /\b(CJ|PC|UN|JG|PÇ|PCS|PÇS)\s+(\d+[\.,]\d{3})\s+\d+[\.,]\d{3}\s+[\d\.]+[\.,]\d{3}\s+(\d+[\.,]\d{3})\s*$/i;
  for (let i=0;i<lines.length;i++) {
    const m=lines[i].match(itemLine); if(!m) continue;
    const item=m[1], code=m[2];
    const parts:string[]=[];
    const inline=m[3]; if(inline) parts.push(inline);
    // Captura até 4 linhas anteriores de descrição que não sejam cabeçalho nem outro item.
    for(let k=i-1, n=0;k>=0 && n<4;k--){
      const v=lines[k]; if(!v) continue; if(header.test(v)||itemLine.test(v)) break;
      if(terminal.test(v)) break; parts.unshift(v); n++;
    }
    let unit:string|undefined, quantity:number|undefined, volume:number|undefined;
    const own=lines[i].match(terminal);
    if(own){ unit=own[1]; quantity=num(own[2]); volume=num(own[3]); }
    for(let k=i+1, n=0;k<lines.length && n<8;k++){
      const v=lines[k]; if(!v) continue;
      if(itemLine.test(v)||header.test(v)) break;
      const t=v.match(terminal);
      if(t){ unit=t[1]; quantity=num(t[2]); volume=num(t[3]); break; }
      parts.push(v); n++;
    }
    // Em muitos PDFs os valores UN/Qtd estão na própria linha e a descrição vem no meio.
    const joined=clean(parts.join(" "));
    const t2=joined.match(/(.+?)\s+(CJ|PC|UN|JG|PÇ|PCS|PÇS)\s+(\d+[\.,]\d{3})\s+\d+[\.,]\d{3}\s+[\d\.]+[\.,]\d{3}\s+(\d+[\.,]\d{3})$/i);
    const desc=t2?clean(t2[1]):joined;
    if(t2){unit=t2[2];quantity=num(t2[3]);volume=num(t2[4]);}
    if(quantity!==undefined) put({item,code,description:desc,unit,quantity:Math.round(quantity),volume});
  }

  return [...found.values()].sort((a,b)=>{
    const pa=a.item.split('.').map(Number), pb=b.item.split('.').map(Number);
    return (pa[0]-pb[0]) || ((pa[1]||0)-(pb[1]||0));
  });
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

function compactItemIds(ids: string[]) {
  const cleanIds = [...new Set(ids.filter(Boolean))];
  if (!cleanIds.length) return "";
  // Compacta sequências do mesmo sufixo: 1.4 ... 10.4 => 1.4 ao 10.4.
  const groups = new Map<string, number[]>();
  const loose: string[] = [];
  for (const id of cleanIds) {
    const m = id.match(/^(\d+)\.(\d+)$/);
    if (!m) { loose.push(id); continue; }
    const arr = groups.get(m[2]) || [];
    arr.push(Number(m[1])); groups.set(m[2], arr);
  }
  const parts: string[] = [];
  for (const [suffix, nums0] of groups) {
    const nums = [...new Set(nums0)].sort((a,b)=>a-b);
    let start = nums[0], prev = nums[0];
    const flush = () => parts.push(start === prev ? `${start}.${suffix}` : `${start}.${suffix} ao ${prev}.${suffix}`);
    for (let i=1;i<nums.length;i++) {
      if (nums[i] === prev + 1) { prev = nums[i]; continue; }
      flush(); start = prev = nums[i];
    }
    flush();
  }
  parts.push(...loose);
  return parts.join(" / ");
}

function sourceItemsText(items: PdfItem[]) {
  const ids = items.map(i => i.item).filter(Boolean);
  if (!ids.length) return undefined;
  const compact = compactItemIds(ids);
  return ids.length === 1 ? `Item ${compact}` : `Itens ${compact}`;
}

function sourceMeta(items: PdfItem[]) {
  return {
    sourceItems: items.map(i=>i.item),
    sourceCodes: items.map(i=>i.code),
    itemText: sourceItemsText(items),
  };
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
  const itemByWidth = new Map<string, PdfItem[]>();
  for (const i of pdfDoors) {
    const d=dimsFromDescription(i.description); if(!d) continue;
    const key=`${d[0]}-${d[1]}-${d[2]}`; const arr=itemByWidth.get(key)||[]; arr.push(i); itemByWidth.set(key,arr);
  }
  const obsUsed = new Set<string>();

  for (const e of entries) {
    const h = parseHand(e.notes.join(" "));
    const app = appName(h.application);
    const base = genericDoorBase(pdfDoors, e);
    const key=`${e.lengthMm||0}-${e.widthMm||0}-${e.thicknessMm||0}`;
    const sources=itemByWidth.get(key)||[];
    const make = (quantity: number, side: "Direita"|"Esquerda", hand: "R"|"L") => {
      if (!quantity) return;
      const firstForSize=!obsUsed.has(key); if(firstForSize) obsUsed.add(key);
      atomic.push({
        width: e.widthMm || 0, app, hand,
        row: {
          quantity,
          lengthMm: e.lengthMm,
          widthMm: e.widthMm,
          thicknessMm: e.thicknessMm,
          product: withPlastic(`${base} ${side}${app ? ` ${app}` : ""}`),
          observation: firstForSize ? sourceItemsText(sources) : undefined,
          category:"PORTA", groupType:"PORTA", groupId:`PORTA-${key}`,
          application: app || undefined, hand: side === "Direita" ? "DIREITA":"ESQUERDA",
          originalDescription: sources[0]?.description || e.description,
          sourceItems: sources.map(i=>i.item), sourceCodes:sources.map(i=>i.code), itemText: firstForSize ? sourceItemsText(sources) : undefined,
          packaging:/papel[aã]o/i.test(e.description)?"PAPELAO":"PLASTICO",
          matchConfidence:sources.length?95:70,
        },
      });
    };
    if (/externa/i.test(app)) { make(h.right, "Direita", "R"); make(h.left, "Esquerda", "L"); }
    else { make(h.left, "Esquerda", "L"); make(h.right, "Direita", "R"); }
  }

  atomic.sort((a, b) => (b.width - a.width) || (a.app.localeCompare(b.app)) || a.hand.localeCompare(b.hand));
  return atomic.map(a => a.row);
}

function buildFrameLegPackages(machining: MachiningEntry[], startNumber: number, pdfItems: PdfItem[] = []): PackageData[] {
  const frames = machining.filter(e => /Batente|Marco/i.test(e.description) && e.widthMm && (parseHand(e.notes.join(" ")).right + parseHand(e.notes.join(" ")).left > 0));
  const pdfFrames=pdfItems.filter(i=>isFrame(i.description)&&!/S\/REB|S\/REBAIXO/i.test(i.description));
  const frameSourceText=sourceItemsText(pdfFrames);
  const meta=sourceMeta(pdfFrames);
  const rightRows: PackageRow[] = [];
  const leftRows: PackageRow[] = [];
  const travRows: PackageRow[] = [];
  let obsPlaced=false;

  for (const e of frames) {
    const h = parseHand(e.notes.join(" "));
    const app = appName(h.application);
    const extra = /eletr[oô]nica/i.test(app) ? " Eletrônica" : "";
    const W = e.widthMm;
    const T = e.thicknessMm || 30;
    const L = 2110;
    const group=`MARCO-${W}-${app||"PADRAO"}`;
    const baseMeta={category:"MARCO" as const,groupType:"MARCO" as const,application:app||undefined,sourceItems:meta.sourceItems,sourceCodes:meta.sourceCodes,originalDescription:e.description,matchConfidence:pdfFrames.length?90:70};

    if (h.right) {
      rightRows.push({ ...baseMeta, groupId:"MARCO-USINADO-DIREITA", quantity: h.right, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet Freijo Nogara 2L Borr. Aplic. "I" 42x10 Dobradiça Direita${extra}`), observation: !obsPlaced ? frameSourceText : undefined, itemText: !obsPlaced ? frameSourceText : undefined, hand:"DIREITA", mergeObservation:true });
      obsPlaced=true;
      rightRows.push({ ...baseMeta, groupId:"MARCO-USINADO-DIREITA", quantity: h.right, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet Freijo Nogara 2L Borr. Aplic. "I" 42x10 Contratesta Direita${extra}`), hand:"DIREITA" });
    }
    if (h.left) {
      leftRows.push({ ...baseMeta, groupId:"MARCO-USINADO-ESQUERDA", quantity: h.left, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet Freijo Nogara 2L Borr. Aplic. "I" 42x10 Dobradiça Esquerda${extra}`), hand:"ESQUERDA" });
      leftRows.push({ ...baseMeta, groupId:"MARCO-USINADO-ESQUERDA", quantity: h.left, lengthMm: L, widthMm: W, thicknessMm: T, product: withPlastic(`Perna de Marco Pet Freijo Nogara 2L Borr. Aplic. "I" 42x10 Contratesta Esquerda${extra}`), hand:"ESQUERDA" });
    }

    const head = frameHeadLength(e);
    if (head) travRows.push({ ...baseMeta, groupId:"MARCO-USINADO-TRAVESSA", quantity: e.quantity, lengthMm: head, widthMm: W, thicknessMm: T, product: withPlastic(`Travessa de Marco Pet Freijo Nogara 2L Borr. Aplic. "I" 42x10`), hand:"SEM_MAO" });
  }

  const packages: PackageData[] = [];
  if (rightRows.length) packages.push({ number: startNumber + packages.length, totalVolume: pkgVolume(rightRows), rows: rightRows,packageType:"MARCO" });
  if (leftRows.length) packages.push({ number: startNumber + packages.length, totalVolume: pkgVolume(leftRows), rows: leftRows,packageType:"MARCO" });
  if (travRows.length) packages.push({ number: startNumber + packages.length, totalVolume: pkgVolume(travRows), rows: travRows,packageType:"MARCO" });
  return packages;
}

function buildSimpleDoorPackage(items: PdfItem[], number: number): PackageData | undefined {
  if (!items.length) return undefined;
  const rows: PackageRow[] = items.map(i => {
    const d = dimsFromDescription(i.description);
    return { quantity: i.quantity, lengthMm: d?.[0], widthMm: d?.[1], thicknessMm: d?.[2], volume: i.volume, product: withPlastic(humanDoorBase(i.description)), observation: sourceItemsText([i]), category:"PORTA", groupType:"PORTA", groupId:`PORTA-${i.item}`, sourceItems:[i.item], sourceCode:i.code, sourceCodes:[i.code], originalDescription:i.description, itemText:sourceItemsText([i]), packaging:/papel[aã]o/i.test(i.description)?"PAPELAO":"PLASTICO", matchConfidence:100 };
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
    const gid=`MARCO-${i.item}`;
    const meta={category:"MARCO" as const,groupType:"MARCO" as const,groupId:gid,sourceItems:[i.item],sourceCode:i.code,sourceCodes:[i.code],originalDescription:i.description,itemText:sourceItemsText([i]),matchConfidence:100};
    rows.push({ ...meta, games: i.quantity, quantity: i.quantity * 2, lengthMm: p.legL, widthMm: p.width, thicknessMm: p.thickness, product: withPlastic(`Conj. Marco Pet Freijo Nogara 2L + Travessa 1L S/Rebaixo`), observation: sourceItemsText([i]), mergeProduct:true, mergeObservation:true });
    if (p.headL) rows.push({ ...meta, quantity: i.quantity, lengthMm: p.headL, widthMm: p.width, thicknessMm: p.thickness, product: "" });
  }
  if (pivot) {
    // dobradiça pivotante é agregada ao mesmo pallet no modelo real.
    rows.push({ quantity: items.reduce((s, i) => s + i.quantity, 0), product: "Dobradiça Pivotante C/Esfera 150KG Cromada (Pino)", category:"FERRAGEM", groupType:"FERRAGEM" });
  }
  return { number, totalVolume: pkgVolume(rows), rows };
}

function buildTrimRows(items: PdfItem[]): PackageRow[] {
  type G = { games:number; len1:number; len2:number; width:number; thick:number; reg:string; items:string[]; codes:string[]; original:string[] };
  const grouped = new Map<string,G>();
  const fallback: PackageRow[] = [];
  for (const i of items) {
    const desc = clean(i.description);
    // Aceita PCS, PC, PÇ/PÇS e espaços variados. A Famossul normalmente usa 4 pernas + 2 travessas.
    const m1 = desc.match(/4\s*(?:PCS?|PÇS?|PECAS?)\s*(?:DE\s*)?(\d{3,4})\s*[Xx]\s*(\d{2,4})\s*[Xx]\s*(\d{1,3})/i)
      || desc.match(/(\d{3,4})\s*[Xx]\s*(\d{2,4})\s*[Xx]\s*(\d{1,3}).*?4\s*(?:PCS?|PÇS?)/i);
    const m2 = desc.match(/2\s*(?:PCS?|PÇS?|PECAS?)\s*(?:DE\s*)?(\d{3,4})/i);
    const regMatch = desc.match(/MAIOR\s*(\d{2,3})\s*MM/i);
    const reg = regMatch?.[1] || (/MAIOR\s*55/i.test(desc)?"55":"35");
    if (!m1 || !m2) {
      const d=dimsFromDescription(desc);
      fallback.push({
        games:i.quantity, quantity:i.quantity, lengthMm:d?.[0], widthMm:d?.[1], thicknessMm:d?.[2], volume:i.volume,
        product:desc, observation:sourceItemsText([i]), itemText:sourceItemsText([i]),
        category:"ALIZAR",groupType:"ALIZAR",groupId:`ALIZAR-FALLBACK-${i.item}`,sourceItems:[i.item],sourceCode:i.code,sourceCodes:[i.code],originalDescription:desc,matchConfidence:70,
      });
      continue;
    }
    const key=`${reg}-${m1[1]}-${m2[1]}-${m1[2]}-${m1[3]}`;
    const g=grouped.get(key)||{games:0,len1:Number(m1[1]),len2:Number(m2[1]),width:Number(m1[2]),thick:Number(m1[3]),reg,items:[],codes:[],original:[]};
    g.games+=i.quantity;g.items.push(i.item);g.codes.push(i.code);g.original.push(desc);grouped.set(key,g);
  }

  const rows:PackageRow[]=[];
  for(const g of [...grouped.values()].sort((a,b)=>Number(b.reg)-Number(a.reg)||a.len1-b.len1)){
    const gid=`ALIZAR-${g.reg}-${g.len1}-${g.len2}-${g.width}-${g.thick}`;
    const txt=g.items.length?`Itens ${compactItemIds(g.items)}`:undefined;
    const meta={category:"ALIZAR" as const,groupType:"ALIZAR" as const,groupId:gid,sourceItems:g.items,sourceCodes:g.codes,itemText:txt,mergeObservation:true,matchConfidence:100,originalDescription:g.original.join(" | ")};
    rows.push({...meta,productGroupId:`${gid}-MAIOR`,games:g.games,quantity:g.games*2,lengthMm:g.len1,widthMm:g.width,thicknessMm:g.thick,product:withPlastic(`Conj. L Maior ${g.reg}mm MDF Ultra Pet Freijo Nogara 2L R.1`),observation:txt,mergeProduct:true});
    rows.push({...meta,productGroupId:`${gid}-MAIOR`,quantity:g.games,lengthMm:g.len2,widthMm:g.width,thicknessMm:g.thick,product:""});
    rows.push({...meta,productGroupId:`${gid}-MENOR`,quantity:g.games*2,lengthMm:g.len1,widthMm:g.width,thicknessMm:g.thick,product:withPlastic(`Conj. L Menor MDF Ultra Pet Freijo Nogara 2L R.1`),mergeProduct:true});
    rows.push({...meta,productGroupId:`${gid}-MENOR`,quantity:g.games,lengthMm:g.len2,widthMm:g.width,thicknessMm:g.thick,product:""});
  }
  return [...rows,...fallback];
}

function buildKitPackage(items: PdfItem[], number: number): PackageData | undefined {
  if (!items.length) return undefined;
  const rows: PackageRow[] = [];
  for (const i of items) {
    const q = i.quantity;
    const meta={category:"KIT" as const,groupType:"KIT" as const,groupId:`KIT-${i.item}`,sourceItems:[i.item],sourceCode:i.code,sourceCodes:[i.code],originalDescription:i.description,itemText:sourceItemsText([i]),mergeObservation:true,matchConfidence:100};
    rows.push({ ...meta, games: q, quantity: q * 2, lengthMm: 2200, widthMm: 55, thicknessMm: 45, product: withPlastic(`Conj. Kit de Correr \"Modelo B\" Pet Freijo Nogara + Ferragens`), observation: sourceItemsText([i]), mergeProduct:true });
    rows.push({ ...meta, quantity: q, lengthMm: 2200, widthMm: 55, thicknessMm: 45, product: "" });
    rows.push({ ...meta, quantity: q, lengthMm: 2200, widthMm: 100, thicknessMm: 9, product: "" });
    rows.push({ ...meta, quantity: q * 3, lengthMm: 2200, widthMm: 50, thicknessMm: 9, product: "" });
  }
  return { number, totalVolume: pkgVolume(rows), rows };
}

function catalogCategory(desc:string){
  if(isDoor(desc)) return "PORTA" as const;
  if(isFrame(desc)) return "MARCO" as const;
  if(isTrim(desc)) return "ALIZAR" as const;
  if(isHinge(desc)||/FECHADURA|FERRAGEM/i.test(desc)) return "FERRAGEM" as const;
  if(isKitCorrer(desc)) return "KIT" as const;
  return "OUTRO" as const;
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
  const framePkgs = buildFrameLegPackages(machining, packages.length + 1, items);
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
      return { quantity: i.quantity, lengthMm: d?.[0], widthMm: d?.[1], thicknessMm: d?.[2], volume: i.volume, product: withPlastic(`Baguetes P/Porta Pivotante Pinus Rec. Pet Freijo Nogara C/Borr. Aplic. \"T\" R.1`), observation: sourceItemsText([i]), category:"OUTRO", groupType:"OUTRO", groupId:`BAGUETE-${i.item}`, sourceItems:[i.item], sourceCode:i.code, sourceCodes:[i.code], originalDescription:i.description, itemText:sourceItemsText([i]), matchConfidence:100 } as PackageRow;
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
    packages.push({ number: packages.length + 1, totalVolume: 0, rows: [{ quantity: qty, product: "Dobradiça Aço Inox 304 Famossul 3x2,5 R16 ESC", observation: sourceItemsText(hinges), category:"FERRAGEM", groupType:"FERRAGEM", groupId:"DOBRADICAS", sourceItems:hinges.map(i=>i.item), sourceCodes:hinges.map(i=>i.code), itemText:sourceItemsText(hinges), matchConfidence:100 }] });
  }

  if (!items.length) warnings.push("Não foi possível identificar automaticamente os itens do pedido.");
  if (machiningBuffer && !machining.length) warnings.push("A planilha de usinagem foi recebida, mas não foi possível ler suas linhas.");
  if (trimRows.length) warnings.push("Alizares foram estruturados por regulagem e medidas. A divisão física final dos pallets de alizar deve ser conferida, pois varia conforme a embalagem da obra.");

  const handRows = machining.map(e => parseHand(e.notes.join(" "))).filter(h => h.right || h.left);
  const handSplit = handRows.length ? {
    right: handRows.reduce((s,h) => s+h.right, 0),
    left: handRows.reduce((s,h) => s+h.left, 0),
  } : undefined;

  const mapped = new Set(packages.flatMap(p=>p.rows.flatMap(r=>r.sourceItems||[])));
  const sourceCatalog = items.map(i=>({ item:i.item, code:i.code, description:i.description, unit:i.unit, quantity:i.quantity, volume:i.volume, category:catalogCategory(i.description), used:mapped.has(i.item) }));
  const unmappedItems = sourceCatalog.filter(i=>!i.used);
  if(unmappedItems.length){
    warnings.push(`${unmappedItems.length} item(ns) do desmembrado ainda precisam de conferência: ${unmappedItems.slice(0,8).map(i=>i.item).join(" / ")}${unmappedItems.length>8?" ...":""}.`);
  }

  const base: ProcessingResult = {
    orderNumber, client, destination, delivery,
    hasMachining: Boolean(machiningBuffer?.length),
    handSplit,
    mountType: options?.mountType || "MONTADO_HS",
    mixedOrder: Boolean(options?.mixedOrder),
    packages,
    warnings,
    sourceItemCount: items.length,
    sourceItems: items.map(i=>i.item),
    sourceCatalog,
    unmappedItems,
    sourceMode: "PEDIDO",
    config: options?.config || DEFAULT_LOGISTICS_CONFIG,
  };
  return applyLogistics(base, base.mountType, base.config);
}
