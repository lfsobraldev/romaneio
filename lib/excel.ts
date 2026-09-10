import ExcelJS from "exceljs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import type { ProcessingResult, PackageRow } from "./types";

function rowVolume(row: PackageRow): number | undefined {
  if (typeof row.volume === "number") return row.volume;
  if (!row.lengthMm || !row.widthMm || !row.thicknessMm || !row.quantity) return undefined;
  return (row.quantity * row.lengthMm * row.widthMm * row.thicknessMm) / 1_000_000_000;
}

function setValue(ws: ExcelJS.Worksheet, address: string, value: string | number | null | undefined) {
  ws.getCell(address).value = value ?? "";
}

export async function createRomaneioWorkbook(data: ProcessingResult) {
  // IMPORTANTE: o arquivo final parte do modelo oficial da Famossul.
  // Assim logo, bordas, larguras, alturas, cabeçalho e impressão permanecem iguais ao modelo.
  const templatePath = path.join(process.cwd(), "public", "templates", "Romaneio.xlsx");
  const template = await readFile(templatePath);

  const wb = new ExcelJS.Workbook();

const templateBuffer = Buffer.from(template);
await wb.xlsx.load(templateBuffer as any);

wb.creator = "Famossul | Gerador de Romaneios";

  const ws = wb.getWorksheet("Romaneio") || wb.worksheets[0];
  if (!ws) throw new Error("Modelo de romaneio inválido.");

  // Cabeçalho do próprio modelo.
  setValue(ws, "B7", new Date());
  ws.getCell("B7").numFmt = "dd/mm/yyyy";
  setValue(ws, "B8", data.client);
  setValue(ws, "B9", data.destination);
  setValue(ws, "B10", data.orderNumber);
  setValue(ws, "B11", data.delivery || "");

  // Limpa somente os VALORES da área preenchível; preserva toda a formatação do modelo.
  for (let r = 13; r <= 109; r++) {
    for (let c = 1; c <= 11; c++) ws.getCell(r, c).value = null;
    ws.getRow(r).hidden = false;
  }

  // Remove mesclagens antigas da área de itens, caso existam em um arquivo reaproveitado.
  // O template em branco normalmente não possui mesclagens nesta área.
  const merges = (ws as any)._merges || {};
  for (const key of Object.keys(merges)) {
    const range = String(key);
    const m = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (m && Number(m[2]) >= 13 && Number(m[4]) <= 109) {
      try { ws.unMergeCells(range); } catch {}
    }
  }

  let r = 13;
  const MAX_DATA_ROW = 108;

  for (const pkg of data.packages) {
    if (!pkg.rows.length) continue;
    const start = r;

    for (const item of pkg.rows) {
      if (r > MAX_DATA_ROW) throw new Error("Quantidade de linhas excede o modelo de romaneio.");

      setValue(ws, `D${r}`, item.quantity || "");
      setValue(ws, `E${r}`, item.lengthMm || "");
      setValue(ws, `F${r}`, item.widthMm || "");
      setValue(ws, `G${r}`, item.thicknessMm || "");

      const vol = rowVolume(item);
      setValue(ws, `H${r}`, typeof vol === "number" ? vol : "");
      ws.getCell(`H${r}`).numFmt = "0.000";

      setValue(ws, `J${r}`, item.product);
      setValue(ws, `K${r}`, item.observation || pkg.notes || "");

      // Alinhamentos seguindo o modelo preenchido enviado.
      for (const col of [1,2,3,4,5,6,7,8,9]) {
        ws.getCell(r, col).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      }
      ws.getCell(`J${r}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      ws.getCell(`K${r}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      r++;
    }

    const end = r - 1;
    setValue(ws, `A${start}`, pkg.number);
    setValue(ws, `B${start}`, data.orderNumber);
    setValue(ws, `C${start}`, pkg.games || "");
    setValue(ws, `I${start}`, typeof pkg.totalVolume === "number" ? pkg.totalVolume : "");
    ws.getCell(`I${start}`).numFmt = "0.000";

    // Um número de pacote/pedido/jogos/m³ pallet para o grupo inteiro, como no modelo real.
    if (end > start) {
      ws.mergeCells(`A${start}:A${end}`);
      ws.mergeCells(`B${start}:B${end}`);
      ws.mergeCells(`C${start}:C${end}`);
      ws.mergeCells(`I${start}:I${end}`);
    }
    for (const col of ["A","B","C","I"]) {
      ws.getCell(`${col}${start}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    }
  }

  // Total imediatamente depois do último item, igual ao romaneio preenchido.
  const totalRow = Math.max(r, 13);
  const totalM3 = data.packages.reduce((sum, pkg) => sum + (pkg.totalVolume || 0), 0);

  setValue(ws, `G${totalRow}`, "Total:");
  setValue(ws, `H${totalRow}`, totalM3);
  ws.getCell(`H${totalRow}`).numFmt = "0.000";
  ws.getCell(`G${totalRow}`).font = { ...(ws.getCell(`G${totalRow}`).font || {}), bold: true };
  ws.getCell(`H${totalRow}`).font = { ...(ws.getCell(`H${totalRow}`).font || {}), bold: true };

  // Mantém o rodapé do modelo e esconde as linhas não utilizadas, reproduzindo a aparência do exemplo.
  for (let rr = totalRow + 1; rr <= 109; rr++) ws.getRow(rr).hidden = true;

  // Rodapé existente no modelo.
  setValue(ws, "B111", data.hasMachining ? "Mont. HS" : "");

  ws.pageSetup = {
    ...ws.pageSetup,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    margins: { left: 0.15, right: 0.15, top: 0.2, bottom: 0.2, header: 0, footer: 0 },
    printArea: "A1:K111",
  };

  return Buffer.from(await wb.xlsx.writeBuffer());
}
