import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { readFile } from "node:fs/promises";
import path from "node:path";
import type { ProcessingResult } from "./types";

const BLACK = rgb(0, 0, 0);
const RED = rgb(1, 0, 0);
const WHITE = rgb(1, 1, 1);

function fmtM3(v?: number) {
  return (v || 0).toFixed(3).replace(".", ",");
}

function fitText(text: string, max = 92) {
  const clean = text.replace(/\s+/g, " ").trim();
  return clean.length <= max ? clean : `${clean.slice(0, max - 3)}...`;
}

export async function createLabelsPdf(data: ProcessingResult) {
  // Usa como FUNDO a própria etiqueta oficial enviada pelo usuário.
  // O PDF apenas preenche os campos variáveis por cima do modelo.
  const templatePath = path.join(process.cwd(), "public", "templates", "etiqueta-base.png");
  const templateBytes = await readFile(templatePath);

  const doc = await PDFDocument.create();
  const bg = await doc.embedPng(templateBytes);
  const regular = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const boldItalic = await doc.embedFont(StandardFonts.HelveticaBoldOblique);

  // Mesma proporção do modelo original enviado: 684 x 367.
  const W = 684;
  const H = 367;
  const totalPallets = data.packages.length;

  for (const pkg of data.packages) {
    const page = doc.addPage([W, H]);
    page.drawImage(bg, { x: 0, y: 0, width: W, height: H });

    // Limpa apenas o interior dos campos que continham os valores de exemplo no modelo.
    // As bordas permanecem visíveis.
    page.drawRectangle({ x: 79, y: H - 213, width: 136, height: 40, color: WHITE }); // pedido
    page.drawRectangle({ x: 356, y: H - 213, width: 326, height: 40, color: WHITE }); // destino
    page.drawRectangle({ x: 80, y: H - 258, width: 90, height: 41, color: WHITE });  // pallet atual
    page.drawRectangle({ x: 230, y: H - 258, width: 124, height: 41, color: WHITE }); // total pallets
    page.drawRectangle({ x: 2, y: H - 342, width: 75, height: 59, color: WHITE });   // pacote nº
    page.drawRectangle({ x: 79, y: H - 342, width: 512, height: 58, color: WHITE }); // dados produto
    page.drawRectangle({ x: 593, y: H - 342, width: 89, height: 58, color: WHITE }); // obs
    page.drawRectangle({ x: 80, y: 2, width: 602, height: 20, color: WHITE });       // totais

    // Cabeçalho variável.
    page.drawText(String(data.orderNumber), {
      x: 92, y: H - 202, size: 13, font: boldItalic, color: BLACK,
    });
    page.drawText(fitText(data.destination || "-", 45), {
      x: 365, y: H - 202, size: 12, font: boldItalic, color: BLACK,
    });

    // Pallet X DE Y.
    page.drawText(String(pkg.number), {
      x: 111, y: H - 252, size: 38, font: bold, color: BLACK,
    });
    page.drawText("DE", {
      x: 199, y: H - 246, size: 20, font: boldItalic, color: RED,
    });
    page.drawText(String(totalPallets), {
      x: 267, y: H - 252, size: 38, font: bold, color: BLACK,
    });

    // Número do pacote.
    page.drawText(String(pkg.number), {
      x: 22, y: H - 337, size: 47, font: bold, color: BLACK,
    });

    // Conteúdo central da etiqueta.
    const rows = pkg.rows.slice(0, 3);
    const lineHeight = rows.length > 1 ? 15 : 18;
    let y = H - 302;

    rows.forEach((row, index) => {
      const jgs = index === 0 && pkg.games ? String(pkg.games) : "";
      const qty = row.quantity ? String(row.quantity) : "";
      const comp = row.lengthMm ? String(row.lengthMm) : "";
      const larg = row.widthMm ? String(row.widthMm) : "";
      const esp = row.thicknessMm ? String(row.thicknessMm) : "";
      const m3 = row.volume !== undefined ? fmtM3(row.volume) : "";

      if (jgs) page.drawText(jgs, { x: 91, y, size: 8, font: regular, color: BLACK });
      if (qty) page.drawText(qty, { x: 128, y, size: 8, font: regular, color: BLACK });
      if (comp) page.drawText(comp, { x: 174, y, size: 8, font: regular, color: BLACK });
      if (larg) page.drawText(larg, { x: 220, y, size: 8, font: regular, color: BLACK });
      if (esp) page.drawText(esp, { x: 269, y, size: 8, font: regular, color: BLACK });
      if (m3) page.drawText(m3, { x: 316, y, size: 7, font: regular, color: BLACK });
      page.drawText(fitText(row.product, 65), { x: 365, y, size: 7.3, font: regular, color: BLACK });
      page.drawText(fitText(row.observation || pkg.notes || "", 14), { x: 603, y, size: 7, font: regular, color: BLACK });
      y -= lineHeight;
    });

    // Total da etiqueta.
    page.drawText(pkg.games ? String(pkg.games) : "0", {
      x: 140, y: 7, size: 10, font: bold, color: BLACK,
    });
    page.drawText(`Total m³ ${fmtM3(pkg.totalVolume)}`, {
      x: 263, y: 7, size: 10, font: bold, color: BLACK,
    });
    page.drawText("0 - toneladas", {
      x: 476, y: 7, size: 11, font: bold, color: BLACK,
    });
  }

  return Buffer.from(await doc.save());
}
