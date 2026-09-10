import ExcelJS from "exceljs";
import type { ProcessingResult } from "./types";

const border: Partial<ExcelJS.Borders> = {
  top: { style: "thin" }, left: { style: "thin" }, bottom: { style: "thin" }, right: { style: "thin" }
};

export async function createRomaneioWorkbook(data: ProcessingResult) {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Famossul | Gerador de Romaneios";
  const ws = wb.addWorksheet("Romaneio", { pageSetup: { paperSize: 9, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1 } });
  ws.columns = [
    { key: "pkg", width: 10 }, { key: "order", width: 11 }, { key: "games", width: 8 }, { key: "qty", width: 8 },
    { key: "length", width: 10 }, { key: "width", width: 10 }, { key: "thick", width: 10 }, { key: "volume", width: 10 },
    { key: "pkgvol", width: 11 }, { key: "product", width: 56 }, { key: "obs", width: 18 }
  ];
  ws.mergeCells("A1:K1"); ws.getCell("A1").value = "FAMOSSUL MADEIRAS NORDESTE LTDA";
  ws.getCell("A1").font = { bold: true, size: 16 }; ws.getCell("A1").alignment = { horizontal: "center" };
  ws.mergeCells("A2:K2"); ws.getCell("A2").value = "ROMANEIO DE SAÍDA DE PRODUTOS - MERCADO INTERNO";
  ws.getCell("A2").font = { bold: true }; ws.getCell("A2").alignment = { horizontal: "center" };
  ws.getCell("A4").value = "Cliente:"; ws.getCell("B4").value = data.client;
  ws.getCell("A5").value = "Destino:"; ws.getCell("B5").value = data.destination;
  ws.getCell("A6").value = "Pedido:"; ws.getCell("B6").value = data.orderNumber;

  const headerRow = 8;
  const headers = ["Nº pacote","Pedido","Jogos","Qtde","Compr.","Largura","Espessura","m³","m³ pallet","Produto","Obs."];
  headers.forEach((h, i) => { const c=ws.getCell(headerRow, i+1); c.value=h; c.font={bold:true}; c.alignment={horizontal:"center",vertical:"middle"}; c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFE7F1DF"}}; c.border=border; });

  let r = headerRow + 1;
  for (const pkg of data.packages) {
    const start = r;
    for (const row of pkg.rows) {
      ws.getCell(r,1).value = pkg.number;
      ws.getCell(r,2).value = data.orderNumber;
      ws.getCell(r,3).value = pkg.games || null;
      ws.getCell(r,4).value = row.quantity;
      ws.getCell(r,5).value = row.lengthMm || null;
      ws.getCell(r,6).value = row.widthMm || null;
      ws.getCell(r,7).value = row.thicknessMm || null;
      ws.getCell(r,8).value = row.volume || null;
      ws.getCell(r,9).value = pkg.totalVolume || null;
      ws.getCell(r,10).value = row.product;
      ws.getCell(r,11).value = row.observation || pkg.notes || "";
      for (let c=1;c<=11;c++){ ws.getCell(r,c).border=border; ws.getCell(r,c).alignment={vertical:"middle",wrapText:true}; }
      r++;
    }
    if (r > start + 1) {
      for (const col of [1,2,3,9]) ws.mergeCells(start,col,r-1,col);
      for (const col of [1,2,3,9]) ws.getCell(start,col).alignment={vertical:"middle",horizontal:"center",wrapText:true};
    }
  }
  const total = data.packages.reduce((s,p)=>s+(p.totalVolume||0),0);
  ws.getCell(r+1,7).value="Total:"; ws.getCell(r+1,7).font={bold:true};
  ws.getCell(r+1,8).value=total; ws.getCell(r+1,8).numFmt="0.000"; ws.getCell(r+1,8).font={bold:true};
  ws.getRow(1).height=28; ws.getRow(2).height=22;
  ws.views = [{ state: "frozen", ySplit: 8 }];
  return Buffer.from(await wb.xlsx.writeBuffer());
}
