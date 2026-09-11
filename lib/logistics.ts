import type { LogisticsConfig, MountType, PackageData, PackageRow, ProcessingResult } from "./types";
import { DEFAULT_LOGISTICS_CONFIG, MOUNT_LABELS, sanitizeConfig } from "./settings";

export function rowVolume(row: PackageRow): number {
  if (typeof row.volume === "number" && Number.isFinite(row.volume)) return row.volume;
  if (!row.lengthMm || !row.widthMm || !row.thicknessMm || !row.quantity) return 0;
  return row.quantity * row.lengthMm * row.widthMm * row.thicknessMm / 1_000_000_000;
}

export function packageVolume(rows: PackageRow[]): number {
  return rows.reduce((sum, row) => sum + rowVolume(row), 0);
}

function isDoor(row: PackageRow) { return row.category === "PORTA" || /FOLHA DE PORTA|PORTA /i.test(row.product); }
function isTrim(row: PackageRow) { return row.category === "ALIZAR" || /ALIZAR|CONJ\. L (?:MAIOR|MENOR)/i.test(row.product); }
function isFrame(row: PackageRow) { return row.category === "MARCO" || /MARCO|BATENTE/i.test(row.product); }
function isCardboard(row: PackageRow) { return row.packaging === "PAPELAO" || /PAPEL[AÃ]O|CANTONEIRA/i.test(`${row.product} ${row.observation || ""}`); }

function splitRow(row: PackageRow, maxQty: number, maxM3: number): PackageRow[] {
  const parts: PackageRow[] = [];
  let remaining = Math.max(0, Math.round(row.quantity || 0));
  if (!remaining) return [{ ...row }];
  const unitVol = rowVolume({ ...row, quantity: 1, volume: undefined });
  const byM3 = unitVol > 0 ? Math.max(1, Math.floor((maxM3 + 1e-9) / unitVol)) : maxQty;
  const cap = Math.max(1, Math.min(maxQty, byM3));
  while (remaining > 0) {
    const take = Math.min(remaining, cap);
    parts.push({ ...row, quantity: take, volume: row.volume !== undefined && row.lengthMm && row.widthMm && row.thicknessMm ? undefined : row.volume });
    remaining -= take;
  }
  return parts;
}

function doorMaxQty(row: PackageRow, mountType: MountType, config: LogisticsConfig) {
  const standard = config[mountType].maxDoors;
  if (mountType !== "REVENDA") return standard;
  if (row.thicknessMm === 41) return Math.min(standard, config.REVENDA.maxDoorsThickness41);
  if (isCardboard(row)) return Math.min(standard, config.REVENDA.maxDoorsCardboard);
  return standard;
}

function packDoorRows(rows: PackageRow[], mountType: MountType, config: LogisticsConfig, mixedOrder: boolean): PackageData[] {
  const baseMaxM3 = config[mountType].maxM3;
  const maxM3 = mixedOrder ? Math.min(baseMaxM3, config.mixedMaxM3) : baseMaxM3;
  const result: PackageData[] = [];
  let currentRows: PackageRow[] = [];
  let currentQty = 0;
  let currentM3 = 0;
  let currentQtyLimit = Number.POSITIVE_INFINITY;
  let special600 = false;
  let hasCardboard = false;
  let has41 = false;

  const flush = () => {
    if (!currentRows.length) return;
    const qtyLimit = Number.isFinite(currentQtyLimit) ? currentQtyLimit : config[mountType].maxDoors;
    const rules = [
      `${MOUNT_LABELS[mountType]}: máx. ${qtyLimit} portas`,
      `máx. ${maxM3.toFixed(3).replace(".", ",")} m³`,
    ];
    if (hasCardboard) rules.push("embalagem papelão/cantoneira");
    if (has41) rules.push("espessura 41 mm: limite reduzido");
    if (special600) rules.push("regra especial 800/820 mm → pallet base 600 mm");
    const invalid = currentQty > qtyLimit || currentM3 > maxM3 + 0.0005;
    result.push({
      number: 0,
      rows: currentRows,
      totalVolume: currentM3,
      limitM3: maxM3,
      limitQuantity: qtyLimit,
      status: invalid ? "INVALIDO" : "VALIDO",
      ruleApplied: rules.join(" • "),
      notes: special600 ? "Pallet base 600 mm" : undefined,
      warnings: special600 ? ["Regra especial aplicada: porta 800/820 mm → pallet de 600 mm."] : [],
    });
    currentRows = []; currentQty = 0; currentM3 = 0; currentQtyLimit = Number.POSITIVE_INFINITY;
    special600 = false; hasCardboard = false; has41 = false;
  };

  // Mantém a ordem industrial fornecida pelo parser (larguras maiores primeiro e mão/aplicação preservadas),
  // mas usa toda a capacidade do pallet antes de abrir o próximo, como no romaneio real 25948.
  for (const original of rows) {
    let remaining = Math.max(0, Math.round(original.quantity || 0));
    const rowMaxQty = doorMaxQty(original, mountType, config);
    const unitM3 = rowVolume({ ...original, quantity: 1, volume: undefined });
    while (remaining > 0) {
      const candidateQtyLimit = Math.min(currentQtyLimit, rowMaxQty);
      // Uma regra mais restritiva (papelão/41mm) não pode tornar inválido o que já estava no pallet.
      if (currentRows.length && currentQty > candidateQtyLimit) { flush(); continue; }
      currentQtyLimit = candidateQtyLimit;
      const qtyRoom = Math.max(0, currentQtyLimit - currentQty);
      const m3Room = Math.max(0, maxM3 - currentM3);
      const m3Units = unitM3 > 0 ? Math.floor((m3Room + 1e-9) / unitM3) : qtyRoom;
      const take = Math.min(remaining, qtyRoom, Math.max(0, m3Units));
      if (take <= 0) { flush(); continue; }
      const part: PackageRow = { ...original, quantity: take, volume: original.volume !== undefined && original.lengthMm && original.widthMm && original.thicknessMm ? undefined : original.volume };
      currentRows.push(part);
      currentQty += take;
      currentM3 += rowVolume(part);
      remaining -= take;
      special600 ||= part.widthMm === 800 || part.widthMm === 820;
      hasCardboard ||= isCardboard(part);
      has41 ||= part.thicknessMm === 41;
      if (currentQty >= currentQtyLimit || currentM3 >= maxM3 - 0.0005) flush();
    }
  }
  flush();
  return result;
}

function splitGamesRows(rows: PackageRow[], maxGames: number, type: "ALIZAR" | "MARCO", maxM3: number): PackageData[] {
  const result: PackageData[] = [];
  // Grupos iniciam em linha com games; linhas seguintes vazias pertencem ao mesmo conjunto.
  const groups: PackageRow[][] = [];
  let current: PackageRow[] = [];
  for (const row of rows) {
    if (row.games && current.length) { groups.push(current); current = []; }
    current.push(row);
  }
  if (current.length) groups.push(current);

  for (const group of groups) {
    const games = group.find(r => r.games)?.games || 0;
    const totalVol = packageVolume(group);
    const perGameVol = games > 0 ? totalVol / games : 0;
    const byM3 = perGameVol > 0 ? Math.max(1, Math.floor((maxM3 + 1e-9) / perGameVol)) : maxGames;
    const effectiveMax = Math.max(1, Math.min(maxGames, byM3));
    if (!games) {
      result.push({ number: 0, rows: group, totalVolume: totalVol, limitM3: maxM3, limitQuantity: maxGames, status: totalVol <= maxM3 + 0.0005 ? "VALIDO" : "ATENCAO", ruleApplied: `${type}: composição sem jogos explícitos; conferir` });
      continue;
    }
    let remaining = games;
    while (remaining > 0) {
      const take = Math.min(remaining, effectiveMax);
      const ratio = take / games;
      const split = group.map((row, idx) => ({
        ...row,
        games: idx === 0 ? take : undefined,
        quantity: Math.round(row.quantity * ratio),
        volume: row.volume !== undefined ? row.volume * ratio : undefined,
      }));
      const splitVol = packageVolume(split);
      result.push({ number: 0, rows: split, games: take, totalVolume: splitVol, limitM3: maxM3, limitQuantity: maxGames, status: splitVol <= maxM3 + 0.0005 ? "VALIDO" : "INVALIDO", ruleApplied: `${type}: até ${maxGames} jogos e ${maxM3.toFixed(3).replace(".", ",")} m³ por pallet` });
      remaining -= take;
    }
  }
  return result;
}

export function applyLogistics(input: ProcessingResult, mountType: MountType, rawConfig?: LogisticsConfig): ProcessingResult {
  const config = sanitizeConfig(rawConfig || DEFAULT_LOGISTICS_CONFIG);
  // Linhas vazias de produto no romaneio são continuação física do conjunto anterior
  // (ex.: travessa de alizar/marco). Propaga a categoria para não perder a composição.
  for (const pkg of input.packages) {
    let inherited: PackageRow["category"] | undefined;
    for (const row of pkg.rows) {
      if (row.product) {
        if (isDoor(row)) inherited = row.category = "PORTA";
        else if (isTrim(row)) inherited = row.category = "ALIZAR";
        else if (isFrame(row)) inherited = row.category = "MARCO";
        else if (/DOBR|FECHADURA|FERRAGEM/i.test(row.product)) inherited = row.category = "FERRAGEM";
        else if (/KIT DE CORRER/i.test(row.product)) inherited = row.category = "KIT";
        else inherited = row.category || "OUTRO";
      } else if (!row.category && inherited) row.category = inherited;
    }
  }
  const allRows = input.packages.flatMap(p => p.rows);
  const doorRows = allRows.filter(isDoor);
  const mixedOrder = Boolean(input.mixedOrder || allRows.some(r => /\bMISTO\b/i.test(`${r.product} ${r.observation || ""}`)));
  const maxM3 = mixedOrder ? Math.min(config[mountType].maxM3, config.mixedMaxM3) : config[mountType].maxM3;

  const consumed = new Set<PackageRow>();
  const out: PackageData[] = [];
  const doors = allRows.filter(r => isDoor(r)); doors.forEach(r => consumed.add(r));
  out.push(...packDoorRows(doors, mountType, config, mixedOrder));

  for (const pkg of input.packages) {
    const remaining = pkg.rows.filter(r => !consumed.has(r));
    if (!remaining.length) continue;
    const trim = remaining.filter(isTrim);
    const frame = remaining.filter(isFrame);
    const other = remaining.filter(r => !isTrim(r) && !isFrame(r));
    trim.forEach(r => consumed.add(r)); frame.forEach(r => consumed.add(r)); other.forEach(r => consumed.add(r));

    if (trim.length) out.push(...splitGamesRows(trim, config[mountType].maxTrimGames, "ALIZAR", maxM3));
    if (frame.length) out.push(...splitGamesRows(frame, config[mountType].maxFrameGames, "MARCO", maxM3));
    if (other.length) {
      const vol = packageVolume(other);
      out.push({ ...pkg, number: 0, rows: other, totalVolume: vol, limitM3: maxM3, status: vol > maxM3 + 0.0005 ? "ATENCAO" : "VALIDO", ruleApplied: `${MOUNT_LABELS[mountType]} • conferência por cubagem` });
    }
  }

  out.forEach((pkg, i) => { pkg.number = i + 1; pkg.totalVolume = packageVolume(pkg.rows); });
  const warnings = [...input.warnings];
  if (mixedOrder) warnings.unshift(`Pedido misto detectado: limite de cubagem aplicado em ${config.mixedMaxM3.toFixed(3).replace(".", ",")} m³ por pallet.`);
  const invalid = out.filter(p => p.status === "INVALIDO").length;
  if (invalid) warnings.unshift(`${invalid} pallet(s) excedem regra logística e precisam de correção antes da emissão.`);

  return { ...input, mountType, mixedOrder, packages: out, warnings, config, orderOptions: { ...(input.orderOptions || { mountType }), mountType } };
}
