import * as XLSX from "xlsx";
import type { MountType, PackageData, PackageRow, ProcessingResult } from "./types";
import { DEFAULT_LOGISTICS_CONFIG, MOUNT_LABELS } from "./settings";
import { packageVolume } from "./logistics";

const clean = (v: unknown) => String(v ?? "").replace(/\s+/g, " ").trim();
const number = (v: unknown): number | undefined => {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  const s = clean(v); if (!s) return undefined;
  const n = Number(s.includes(",") ? s.replace(/\./g, "").replace(",", ".") : s);
  return Number.isFinite(n) ? n : undefined;
};
const norm = (s: string) => s.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase();

function inferMount(text: string): MountType {
  const t = norm(text);
  if (t.includes("TIMADEL")) return "MONTADO_TIMADEL";
  if (t.includes("COMPLEMENTO")) return "COMPLEMENTO_OBRA";
  if (t.includes("REVENDA")) return "REVENDA";
  if (t.includes("ESTANCIA")) return "MONTADO_ESTANCIA";
  return "MONTADO_HS";
}

function classify(product: string): PackageRow["category"] {
  const t = norm(product);
  if (t.includes("ALIZAR") || t.includes("CONJ. L MAIOR") || t.includes("CONJ L MAIOR") || t.includes("CONJ. L MENOR")) return "ALIZAR";
  if (t.includes("MARCO") || t.includes("BATENTE")) return "MARCO";
  if (t.includes("FOLHA DE PORTA") || t.startsWith("PORTA ")) return "PORTA";
  if (t.includes("DOBR") || t.includes("FECHADURA") || t.includes("FERRAGEM")) return "FERRAGEM";
  if (t.includes("KIT DE CORRER")) return "KIT";
  return "OUTRO";
}

export function readReadyRomaneio(buffer: Buffer, fileName = "Romaneio.xlsx"): ProcessingResult {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true, cellFormula: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error("A planilha não possui uma aba válida.");

  const ref = ws["!ref"] || "A1:K200";
  const range = XLSX.utils.decode_range(ref);
  const matrix: unknown[][] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: unknown[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) row.push(ws[XLSX.utils.encode_cell({ r, c })]?.v ?? "");
    matrix.push(row);
  }

  let headerIndex = -1;
  for (let r = 0; r < Math.min(matrix.length, 40); r++) {
    const joined = matrix[r].map(clean).join(" | ");
    const n = norm(joined);
    if ((n.includes("PACOTE") || n.includes("N º PACOTE") || n.includes("N° PACOTE")) && n.includes("PRODUTO")) { headerIndex = r; break; }
  }
  if (headerIndex < 0) throw new Error("Não encontrei o cabeçalho do romaneio. Use o modelo da Famossul ou um romaneio já preenchido.");

  const headers = matrix[headerIndex].map(v => norm(clean(v)));
  const idx = (terms: string[]) => headers.findIndex(h => terms.some(t => h.includes(t)));
  const cPackage = idx(["PACOTE"]);
  const cOrder = idx(["PEDIDO"]);
  const cGames = idx(["JOGOS", "JGS"]);
  const cQty = idx(["QTDE", "QTD"]);
  const cLen = idx(["COMPR"]);
  const cWid = idx(["LARG"]);
  const cThk = idx(["ESPESS", "ESP."]);
  const cM3 = idx(["M³", "M3"]);
  const cM3Pallet = headers.findIndex(h => h.includes("PALLET") && (h.includes("M3") || h.includes("M³")));
  const cProduct = idx(["PRODUTO"]);
  const cObs = idx(["OBS"]);
  if ([cPackage,cOrder,cQty,cProduct].some(v => v < 0)) throw new Error("O romaneio não possui as colunas mínimas esperadas (Pacote, Pedido, Qtde e Produto).");

  // Resolve values inherited by merged cells using !merges.
  const mergeLookup = new Map<string, unknown>();
  for (const m of (ws["!merges"] || [])) {
    const top = ws[XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })]?.v ?? "";
    for (let r=m.s.r;r<=m.e.r;r++) for (let c=m.s.c;c<=m.e.c;c++) mergeLookup.set(`${r}:${c}`, top);
  }
  const val = (r:number,c:number) => {
    if (c < 0) return "";
    const direct = ws[XLSX.utils.encode_cell({ r, c })]?.v;
    return direct === undefined || direct === null || direct === "" ? (mergeLookup.get(`${r}:${c}`) ?? "") : direct;
  };

  const packagesMap = new Map<number, PackageData>();
  let currentPackage = 0;
  const endRow = Math.min(range.e.r, headerIndex + 140);
  for (let r = headerIndex + 1; r <= endRow; r++) {
    const labelA = norm(clean(val(r, cPackage)));
    if (labelA.includes("FILTRO") || labelA === "OBS:" || labelA.startsWith("TOTAL")) break;
    const rawPackage = number(val(r,cPackage));
    if (rawPackage !== undefined && rawPackage > 0) currentPackage = Math.round(rawPackage);
    if (!currentPackage) continue;

    const product = clean(val(r,cProduct));
    const qty = number(val(r,cQty));
    const games = number(val(r,cGames));
    const len = number(val(r,cLen)), wid = number(val(r,cWid)), thk = number(val(r,cThk));
    const m3 = number(val(r,cM3));
    const obs = clean(val(r,cObs));
    if (!product && qty === undefined && !games && !obs) continue;

    let pkg = packagesMap.get(currentPackage);
    if (!pkg) { pkg = { number: currentPackage, rows: [], totalVolume: 0, status: "VALIDO", warnings: [] }; packagesMap.set(currentPackage,pkg); }
    const row: PackageRow = {
      games: games ? Math.round(games) : undefined,
      quantity: Math.round(qty || 0),
      lengthMm: len, widthMm: wid, thicknessMm: thk,
      volume: m3,
      product,
      observation: obs || undefined,
      itemText: /\bItens?\b/i.test(obs) ? obs : undefined,
      category: classify(product), groupType: classify(product),
      groupId: `READY-${currentPackage}-${pkg.rows.length+1}`,
      matchConfidence: 100,
    };
    pkg.rows.push(row);
    const pv = number(val(r,cM3Pallet));
    if (pv !== undefined && pv > 0) pkg.totalVolume = pv;
  }

  const packages = [...packagesMap.values()].sort((a,b)=>a.number-b.number);
  for (const p of packages) {
    if (!p.totalVolume) p.totalVolume = packageVolume(p.rows);
    p.ruleApplied = "Romaneio pronto importado • revisar e emitir etiquetas";
  }
  if (!packages.length) throw new Error("Não encontrei linhas de produtos no romaneio enviado.");

  const cell = (address:string) => clean(ws[address]?.v);
  const orderNumber = cell("B10") || clean(val(headerIndex+1,cOrder)) || "NÃO IDENTIFICADO";
  const client = cell("B8") || "Cliente não identificado";
  const destination = cell("B9") || "";
  const delivery = cell("B11") || undefined;
  const footerText = matrix.slice(Math.max(0,matrix.length-12)).flat().map(clean).join(" ");
  const mountType = inferMount(footerText);

  return {
    orderNumber,
    client,
    destination,
    delivery,
    hasMachining: false,
    mountType,
    mixedOrder: false,
    packages,
    warnings: ["Romaneio pronto importado. Confira os pallets e gere as etiquetas sem reprocessar o pedido."],
    config: DEFAULT_LOGISTICS_CONFIG,
    orderOptions: { mountType },
    sourceMode: "ROMANEIO_PRONTO",
    sourceFileName: fileName,
    sourceItemCount: packages.reduce((s,p)=>s+p.rows.length,0),
  };
}
