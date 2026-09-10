import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import type { ProcessingResult } from "./types";

const mm = (n:number) => n * 2.83465;
export async function createLabelsPdf(data: ProcessingResult) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const bold = await doc.embedFont(StandardFonts.HelveticaBold);
  const total = data.packages.length;
  for (const pkg of data.packages) {
    const page = doc.addPage([mm(180), mm(96)]);
    const { height } = page.getSize();
    const x=mm(7), w=mm(166);
    page.drawRectangle({x, y:mm(6), width:w, height:mm(84), borderWidth:1, borderColor:rgb(0.1,0.1,0.1)});
    page.drawText("famossul", {x:x+mm(5), y:height-mm(20), size:24, font:bold});
    page.drawText("Portas para o hoje. E para o amanhã.", {x:x+mm(5), y:height-mm(27), size:7, font});
    page.drawText("Rua A - Quadra 04 - Lote 29 a 34 - Distrito Industrial, Estância/SE", {x:x+mm(58), y:height-mm(15), size:8, font});
    page.drawText("(79) 3522-1228 - www.famossul.com.br", {x:x+mm(58), y:height-mm(21), size:8, font});
    page.drawLine({start:{x,y:height-mm(34)},end:{x:x+w,y:height-mm(34)},thickness:1});
    page.drawText(`Pedido: ${data.orderNumber}`, {x:x+mm(4), y:height-mm(45), size:11, font:bold});
    page.drawText(`Destino: ${data.destination || "-"}`, {x:x+mm(70), y:height-mm(45), size:11, font:bold});
    page.drawText("Pallet", {x:x+mm(4), y:height-mm(58), size:11, font:bold});
    page.drawText(String(pkg.number), {x:x+mm(23), y:height-mm(65), size:32, font:bold});
    page.drawText("DE", {x:x+mm(45), y:height-mm(60), size:17, font:bold});
    page.drawText(String(total), {x:x+mm(62), y:height-mm(65), size:32, font:bold});
    page.drawText(`Pacote nº ${pkg.number}`, {x:x+mm(4), y:height-mm(74), size:8, font:bold});
    const summary = pkg.rows.map(r=>`${r.quantity}x ${r.product}${r.lengthMm?` (${r.lengthMm}x${r.widthMm||0}x${r.thicknessMm||0})`:""}`).join(" | ");
    const safe = summary.length>120 ? summary.slice(0,117)+"..." : summary;
    page.drawText(safe, {x:x+mm(32), y:height-mm(74), size:7, font, maxWidth:mm(126)});
    page.drawText(`m³ pallet: ${(pkg.totalVolume||0).toFixed(3).replace('.',',')}`, {x:x+mm(115), y:mm(9), size:9, font:bold});
  }
  return Buffer.from(await doc.save());
}
