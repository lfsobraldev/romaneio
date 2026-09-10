import * as XLSX from "xlsx";
import type { ProcessingResult, PackageData } from "./types";

const clean = (s: string) => s.replace(/\s+/g, " ").trim();
const toNumber = (v?: string) => v ? Number(v.replace(/\./g, "").replace(",", ".")) : undefined;

function qtyNear(text: string, keyword: string): number | undefined {
  const i = text.toUpperCase().indexOf(keyword.toUpperCase());
  if (i < 0) return;
  const chunk = text.slice(i, i + 550);
  const m = chunk.match(/(?:UN|PC|CJ)\s+(\d+[\.,]\d{3})/i);
  return m ? Math.round(toNumber(m[1]) || 0) : undefined;
}

function dimsNear(text: string, keyword: string): [number, number, number] | undefined {
  const i = text.toUpperCase().indexOf(keyword.toUpperCase());
  if (i < 0) return;
  const chunk = text.slice(i, i + 350);
  const m = chunk.match(/(\d{3,4})\s*[Xx]\s*(\d{2,4})\s*[Xx]\s*(\d{1,3})/);
  return m ? [Number(m[1]), Number(m[2]), Number(m[3])] : undefined;
}

function volumeNear(text: string, keyword: string): number | undefined {
  const i = text.toUpperCase().indexOf(keyword.toUpperCase());
  if (i < 0) return;
  const chunk = text.slice(i, i + 600);
  const matches = [...chunk.matchAll(/\b(\d+[\.,]\d{3})\b/g)].map(m => toNumber(m[1])!);
  return matches.length ? matches[matches.length - 1] : undefined;
}

export function parseMachining(buffer?: Buffer) {
  if (!buffer?.length) return { right: undefined as number|undefined, left: undefined as number|undefined, raw: "" };
  const wb = XLSX.read(buffer, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false }) as unknown[][];
  const raw = data.flat().filter(Boolean).join(" ");
  const m = raw.match(/(\d+)\s*Direitas?\s*\/\s*(\d+)\s*Esquerdas?/i);
  return { right: m ? Number(m[1]) : undefined, left: m ? Number(m[2]) : undefined, raw };
}

export function buildProcessing(pdfText: string, machiningBuffer?: Buffer): ProcessingResult {
  const text = clean(pdfText);
  const orderNumber = text.match(/Pedido:\s*(\d+)/i)?.[1] || "NÃO IDENTIFICADO";
  const client = clean(text.match(/Cliente:\s*(.+?)(?=CNPJ:|Endereço:|Bairro:|Data Emissão:)/i)?.[1] || "Cliente não identificado");
  const city = text.match(/Cidade:\s*([^\s]+(?:\s+[^\s]+)*?-[A-Z]{2}|Cidade:\s*(.+?)(?=Data Emissão:|Data Previsão:))/i);
  const destination = clean(city?.[1] || city?.[2] || "");

  const machining = parseMachining(machiningBuffer);
  const doorQty = qtyNear(text, "FOLHA DE PORTA") || 0;
  const doorDims = dimsNear(text, "FOLHA DE PORTA") || [0,0,0];
  const doorVol = volumeNear(text, "FOLHA DE PORTA");
  const frameQty = qtyNear(text, "MARCO") || doorQty;
  const frameDims = dimsNear(text, "MARCO") || [0,0,0];
  const frameVol = volumeNear(text, "MARCO");
  const trimQty = qtyNear(text, "CONJUNTO ALIZAR") || doorQty;
  const trimDims = dimsNear(text, "CONJUNTO ALIZAR") || [0,0,0];
  const trimVol = volumeNear(text, "CONJUNTO ALIZAR");
  const lockQty = qtyNear(text, "FECHADURA") || 0;
  const hingeQty = qtyNear(text, "DOBR") || 0;

  const packages: PackageData[] = [];
  const warnings: string[] = [];

  const right = machining.right;
  const left = machining.left;
  const doorRows = [];
  if (right !== undefined && left !== undefined && right + left === doorQty) {
    doorRows.push({ quantity: right, lengthMm: doorDims[0], widthMm: doorDims[1], thicknessMm: doorDims[2], product: "Folha de porta — mão direita", observation: "Usinagem" });
    doorRows.push({ quantity: left, lengthMm: doorDims[0], widthMm: doorDims[1], thicknessMm: doorDims[2], product: "Folha de porta — mão esquerda", observation: "Usinagem" });
  } else {
    doorRows.push({ quantity: doorQty, lengthMm: doorDims[0], widthMm: doorDims[1], thicknessMm: doorDims[2], product: "Folha de porta" });
    if (machiningBuffer) warnings.push("Não foi possível validar automaticamente a divisão Direita/Esquerda; confira o pacote 1.");
  }
  packages.push({ number: 1, totalVolume: doorVol, rows: doorRows });

  // Marco: o pedido informa 2 pernas por jogo + 1 travessa. A mão vem da usinagem.
  const frameRows = [];
  const legL = frameDims[0] || 2110;
  const legW = frameDims[1];
  const legT = frameDims[2];
  if (right !== undefined && left !== undefined && right + left === frameQty) {
    frameRows.push({ quantity: right, lengthMm: legL, widthMm: legW, thicknessMm: legT, product: "Perna de marco — jogos mão direita" });
    frameRows.push({ quantity: left, lengthMm: legL, widthMm: legW, thicknessMm: legT, product: "Perna de marco — jogos mão esquerda" });
    frameRows.push({ quantity: right, lengthMm: legL, widthMm: legW, thicknessMm: legT, product: "Segunda perna de marco — jogos mão direita" });
    frameRows.push({ quantity: left, lengthMm: legL, widthMm: legW, thicknessMm: legT, product: "Segunda perna de marco — jogos mão esquerda" });
  } else {
    frameRows.push({ quantity: frameQty * 2, lengthMm: legL, widthMm: legW, thicknessMm: legT, product: "Pernas de marco" });
  }
  const trav = text.match(/TRAV(?:ESSA)?\s+DE\s+(\d{3,4})/i)?.[1] || text.match(/TRAV\.\s*(\d{3,4})/i)?.[1];
  if (trav) frameRows.push({ quantity: frameQty, lengthMm: Number(trav), widthMm: legW, thicknessMm: legT, product: "Travessa de marco" });
  packages.push({ number: 2, totalVolume: frameVol, rows: frameRows });

  packages.push({ number: 3, games: trimQty, totalVolume: trimVol, rows: [{ quantity: trimQty, lengthMm: trimDims[0], widthMm: trimDims[1], thicknessMm: trimDims[2], product: "Conjunto de alizar MDF" }] });
  packages.push({ number: 4, rows: [
    ...(lockQty ? [{ quantity: lockQty, product: "Fechadura" }] : []),
    ...(hingeQty ? [{ quantity: hingeQty, product: "Dobradiça" }] : [])
  ] });

  if (!doorQty) warnings.push("Quantidade de folhas de porta não identificada automaticamente.");
  if (!destination) warnings.push("Destino não identificado automaticamente.");

  return {
    orderNumber, client, destination, hasMachining: Boolean(machiningBuffer),
    handSplit: right !== undefined && left !== undefined ? { right, left } : undefined,
    packages, warnings
  };
}
