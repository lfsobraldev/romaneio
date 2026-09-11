import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { readFile } from "node:fs/promises";
import path from "node:path";
import type { ProcessingResult, PackageRow } from "./types";
import { MOUNT_LABELS } from "./settings";
import { rowVolume } from "./logistics";

const BLACK = rgb(0, 0, 0);
const WHITE = rgb(1, 1, 1);
function fmtM3(v?: number) { return (v || 0).toFixed(3).replace(".", ","); }
function fitText(text: string, max = 92) { const clean = text.replace(/\s+/g, " ").trim(); return clean.length <= max ? clean : `${clean.slice(0, max - 3)}...`; }
function rowM3(row:PackageRow){return rowVolume(row)}

export async function createLabelsPdf(data: ProcessingResult) {
  const templatePath = path.join(process.cwd(), "public", "templates", "etiqueta-base.png");
  const templateBytes = await readFile(templatePath);
  const doc = await PDFDocument.create();
  const bg = await doc.embedPng(templateBytes);
  const regular = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const boldItalic = await doc.embedFont(StandardFonts.HelveticaBoldOblique);
  const W = 684, H = 367, totalPallets = data.packages.length;
  const configExtra=(data.config?.additionalItems||[]).filter(i=>i.enabled&&(i.target==="ETIQUETA"||i.target==="AMBOS")).map(i=>i.text||i.label).filter(Boolean);
  const globalExtra=[data.config?.labelNote,data.orderOptions?.etiquetaExtraText,...configExtra].filter(Boolean).join(" | ");

  for (const pkg of data.packages) {
    const page = doc.addPage([W, H]);
    page.drawImage(bg, { x: 0, y: 0, width: W, height: H });
    page.drawRectangle({ x: 79, y: H - 213, width: 136, height: 40, color: WHITE });
    page.drawRectangle({ x: 356, y: H - 213, width: 326, height: 40, color: WHITE });
    page.drawRectangle({ x: 80, y: H - 258, width: 90, height: 41, color: WHITE });
    page.drawRectangle({ x: 230, y: H - 258, width: 124, height: 41, color: WHITE });
    page.drawRectangle({ x: 2, y: H - 342, width: 75, height: 59, color: WHITE });
    page.drawRectangle({ x: 79, y: H - 342, width: 512, height: 58, color: WHITE });
    page.drawRectangle({ x: 593, y: H - 342, width: 89, height: 58, color: WHITE });
    page.drawRectangle({ x: 80, y: 2, width: 602, height: 20, color: WHITE });

    page.drawText(String(data.orderNumber), { x: 92, y: H - 202, size: 13, font: boldItalic, color: BLACK });
    page.drawText(fitText(data.destination || "-", 45), { x: 365, y: H - 202, size: 12, font: boldItalic, color: BLACK });
    page.drawText(String(pkg.number), { x: 111, y: H - 252, size: 38, font: bold, color: BLACK });
    // O "DE" vermelho já existe no modelo oficial; não redesenha para evitar duplicação.
    page.drawText(String(totalPallets), { x: 267, y: H - 252, size: 38, font: bold, color: BLACK });
    page.drawText(String(pkg.number), { x: 22, y: H - 337, size: 47, font: bold, color: BLACK });

    const rows = pkg.rows.slice(0, 3); const lineHeight = rows.length > 1 ? 15 : 18; let y = H - 302;
    rows.forEach((row, index) => {
      const jgs = row.games ? String(row.games) : (index===0&&pkg.games?String(pkg.games):"");
      const qty = row.quantity ? String(row.quantity) : "";
      const comp = row.lengthMm ? String(row.lengthMm) : "";
      const larg = row.widthMm ? String(row.widthMm) : "";
      const esp = row.thicknessMm ? String(row.thicknessMm) : "";
      const m3 = rowM3(row) ? fmtM3(rowM3(row)) : "";
      if (jgs) page.drawText(jgs, { x: 91, y, size: 8, font: regular, color: BLACK });
      if (qty) page.drawText(qty, { x: 128, y, size: 8, font: regular, color: BLACK });
      if (comp) page.drawText(comp, { x: 174, y, size: 8, font: regular, color: BLACK });
      if (larg) page.drawText(larg, { x: 220, y, size: 8, font: regular, color: BLACK });
      if (esp) page.drawText(esp, { x: 269, y, size: 8, font: regular, color: BLACK });
      if (m3) page.drawText(m3, { x: 316, y, size: 7, font: regular, color: BLACK });
      page.drawText(fitText(row.product, 65), { x: 365, y, size: 7.3, font: regular, color: BLACK });
      const obs=fitText([row.observation,pkg.notes,index===0?globalExtra:""].filter(Boolean).join(" | "),18);
      if(obs)page.drawText(obs, { x: 598, y, size: 6.2, font: regular, color: BLACK });
      y -= lineHeight;
    });

    const totalGames=pkg.rows.reduce((s,r)=>s+(r.games||0),0)||pkg.games||0;
    page.drawText(String(totalGames), { x: 140, y: 7, size: 10, font: bold, color: BLACK });
    page.drawText(`Total m³ ${fmtM3(pkg.totalVolume)}`, { x: 263, y: 7, size: 10, font: bold, color: BLACK });
    page.drawText(fitText(MOUNT_LABELS[data.mountType]||data.mountType,18), { x: 476, y: 7, size: 9, font: bold, color: BLACK });
  }
  return Buffer.from(await doc.save());
}
