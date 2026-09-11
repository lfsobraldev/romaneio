import ExcelJS from "exceljs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import type { ProcessingResult, PackageRow } from "./types";
import { MOUNT_LABELS } from "./settings";

function rowVolume(row: PackageRow): number {
  if (typeof row.volume === "number") return row.volume;
  if (!row.lengthMm || !row.widthMm || !row.thicknessMm || !row.quantity) return 0;
  return (row.quantity * row.lengthMm * row.widthMm * row.thicknessMm) / 1_000_000_000;
}

function setValue(ws: ExcelJS.Worksheet, address: string, value: string | number | Date | null | undefined) {
  ws.getCell(address).value = value ?? "";
}

function mergeSafe(ws: ExcelJS.Worksheet, range: string) {
  try { ws.mergeCells(range); } catch {}
}

export async function createRomaneioWorkbook(data: ProcessingResult) {
  // Este template é um ROMANEIO REAL aprovado da Famossul.
  // O sistema apenas limpa os dados variáveis e preenche o novo pedido,
  // mantendo cores, bordas, larguras, alturas, fontes e impressão originais.
  const templatePath = path.join(process.cwd(), "public", "templates", "Romaneio.xlsx");
  const template = await readFile(templatePath);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(template) as any);
  wb.creator = "Famossul | Gerador de Romaneios";

  const ws = wb.getWorksheet("Romaneio") || wb.worksheets[0];
  if (!ws) throw new Error("Modelo de romaneio inválido.");

  // Remove as mesclagens do romaneio usado como modelo somente na área variável.
  const merges = (ws as any)._merges || {};
  for (const key of Object.keys(merges)) {
    const range = String(key);
    const m = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!m) continue;
    const r1 = Number(m[2]);
    const r2 = Number(m[4]);
    if (r1 >= 13 && r2 <= 126) {
      try { ws.unMergeCells(range); } catch {}
    }
  }

  // Cabeçalho do modelo oficial.
  setValue(ws, "B7", new Date());
  ws.getCell("B7").numFmt = "dd/mm/yyyy";
  setValue(ws, "B8", data.client);
  setValue(ws, "B9", data.destination);
  setValue(ws, "B10", data.orderNumber);
  setValue(ws, "B11", data.delivery || "");
  setValue(ws, "H7", data.orderOptions?.motorista || "");
  setValue(ws, "H8", data.orderOptions?.transportadora || "");
  setValue(ws, "H9", data.orderOptions?.placa || "");
  setValue(ws, "H11", data.orderOptions?.notaFiscal || "");

  // Limpa apenas valores. A aparência do arquivo real continua intacta.
  for (let r = 13; r <= 126; r++) {
    for (let c = 1; c <= 10; c++) ws.getCell(r, c).value = null;
    ws.getRow(r).hidden = false;
  }

  // Limpa valores específicos do rodapé do pedido usado como template.
  setValue(ws, "B127", "");
  setValue(ws, "D127", "");
  setValue(ws, "B128", MOUNT_LABELS[data.mountType] || data.mountType);
  const extraRomaneio = [
    data.config?.romaneioNote,
    data.orderOptions?.romaneioExtraText,
    ...(data.config?.additionalItems || []).filter(i => i.enabled && (i.target === "ROMANEIO" || i.target === "AMBOS")).map(i => i.text || i.label),
    data.orderOptions?.conferente ? `Conferente: ${data.orderOptions.conferente}` : "",
    data.orderOptions?.separador ? `Separado por: ${data.orderOptions.separador}` : "",
  ].filter(Boolean).join(" | ");
  setValue(ws, "B129", extraRomaneio);

  const START = 13;
  const MAX_DATA_ROW = 125;
  let r = START;

  for (const pkg of data.packages) {
    if (!pkg.rows?.length) continue;
    const packageStart = r;

    for (const item of pkg.rows) {
      if (r > MAX_DATA_ROW) throw new Error("Quantidade de linhas excede o modelo real do romaneio (até linha 125).");

      setValue(ws, `C${r}`, item.games ?? "");
      setValue(ws, `D${r}`, item.quantity || "");
      setValue(ws, `E${r}`, item.lengthMm || "");
      setValue(ws, `F${r}`, item.widthMm || "");
      setValue(ws, `G${r}`, item.thicknessMm || "");
      const vol = rowVolume(item);
      setValue(ws, `H${r}`, vol || 0);
      ws.getCell(`H${r}`).numFmt = "0.000";
      setValue(ws, `I${r}`, item.product || "");
      setValue(ws, `J${r}`, item.observation || "");

      for (let c = 1; c <= 10; c++) {
        ws.getCell(r, c).alignment = {
          ...(ws.getCell(r, c).alignment || {}),
          horizontal: c >= 9 ? "center" : "center",
          vertical: "middle",
          wrapText: true,
        };
      }
      r++;
    }

    const packageEnd = r - 1;
    setValue(ws, `A${packageStart}`, pkg.number);
    setValue(ws, `B${packageStart}`, data.orderNumber);

    if (packageEnd > packageStart) {
      mergeSafe(ws, `A${packageStart}:A${packageEnd}`);
      mergeSafe(ws, `B${packageStart}:B${packageEnd}`);
    }
    ws.getCell(`A${packageStart}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    ws.getCell(`B${packageStart}`).alignment = { horizontal: "center", vertical: "middle", wrapText: true };

    // Jogos podem mudar dentro do mesmo pallet (acontece no romaneio 25948).
    // Mescla C somente dentro de cada subgrupo iniciado por um valor de jogos.
    let gamesStart: number | undefined;
    for (let rr = packageStart; rr <= packageEnd + 1; rr++) {
      const hasGames = rr <= packageEnd && ws.getCell(`C${rr}`).value !== "" && ws.getCell(`C${rr}`).value != null;
      if (hasGames) {
        if (gamesStart !== undefined && rr - 1 > gamesStart) mergeSafe(ws, `C${gamesStart}:C${rr - 1}`);
        gamesStart = rr;
      }
      if (rr === packageEnd + 1 && gamesStart !== undefined && packageEnd > gamesStart) mergeSafe(ws, `C${gamesStart}:C${packageEnd}`);
    }

    // Produto e observação: quando as linhas seguintes estão vazias, elas pertencem ao item acima.
    for (const col of ["I", "J"]) {
      let groupStart: number | undefined;
      for (let rr = packageStart; rr <= packageEnd + 1; rr++) {
        const value = rr <= packageEnd ? ws.getCell(`${col}${rr}`).value : "__END__";
        const nonEmpty = value !== "" && value != null;
        if (nonEmpty) {
          if (groupStart !== undefined && rr - 1 > groupStart) mergeSafe(ws, `${col}${groupStart}:${col}${rr - 1}`);
          groupStart = rr <= packageEnd ? rr : undefined;
        }
      }
    }
  }

  // Totais no mesmo formato do arquivo aprovado.
  const totalRow = r;
  const allRows = data.packages.flatMap(p => p.rows || []);
  const totalGames = allRows.reduce((s, row) => s + (row.games || 0), 0);
  const totalQty = allRows.reduce((s, row) => s + (row.quantity || 0), 0);
  const totalM3 = allRows.reduce((s, row) => s + rowVolume(row), 0);

  if (totalRow > 126) throw new Error("O romaneio ultrapassou a área disponível do modelo.");
  setValue(ws, `C${totalRow}`, totalGames || "");
  setValue(ws, `D${totalRow}`, totalQty || "");
  setValue(ws, `G${totalRow}`, "Total m³");
  setValue(ws, `H${totalRow}`, totalM3);
  ws.getCell(`H${totalRow}`).numFmt = "0.000";

  // Esconde linhas que sobraram e mantém o rodapé encostado no conteúdo, como na planilha real.
  for (let rr = totalRow + 1; rr <= 126; rr++) ws.getRow(rr).hidden = true;

  ws.pageSetup = {
    ...ws.pageSetup,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    printArea: "A1:J129",
  };

  return Buffer.from(await wb.xlsx.writeBuffer());
}
