import ExcelJS from "exceljs";
import { readFile } from "node:fs/promises";
import path from "node:path";
import type { PackageRow, ProcessingResult } from "./types";
import { MOUNT_LABELS } from "./settings";
import { rowVolume } from "./logistics";

// Modelo oficial em branco: 13..108 dados, 109 total, 110 filtro/página, 111 observação.
const START_ROW = 13, LAST_DATA_ROW = 108, TOTAL_ROW = 109, FILTER_ROW = 110, OBS_ROW = 111;
const COL = { pkg:"A", order:"B", games:"C", qty:"D", len:"E", wid:"F", thk:"G", m3:"H", palletM3:"I", product:"J", obs:"K" } as const;

function setValue(ws:ExcelJS.Worksheet,address:string,value:string|number|Date|null|undefined){ ws.getCell(address).value = value ?? ""; }
function mergeSafe(ws:ExcelJS.Worksheet,ref:string){ try{ ws.mergeCells(ref); }catch{} }
function itemText(row:PackageRow){
  return row.itemText || row.observation || (row.sourceItems?.length ? (row.sourceItems.length===1 ? `Item ${row.sourceItems[0]}` : `Itens ${row.sourceItems.join(" / ")}`) : "");
}
function unmergeDataArea(ws:ExcelJS.Worksheet){
  const merges=(ws as any)._merges||{};
  for(const ref of Object.keys(merges)){
    const m=ref.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/); if(!m) continue;
    const a=Number(m[2]),b=Number(m[4]);
    if(a<=LAST_DATA_ROW && b>=START_ROW){ try{ws.unMergeCells(ref)}catch{} }
  }
}
function clearDataValues(ws:ExcelJS.Worksheet){
  for(let r=START_ROW;r<=LAST_DATA_ROW;r++){
    ws.getRow(r).hidden=false;
    for(let c=1;c<=11;c++) ws.getCell(r,c).value=null;
  }
}
function groups(rows:PackageRow[]){
  const out:Array<{start:number;end:number;id:string}> = [];
  let start=0, id=rows[0]?.groupId || `ROW-0`;
  for(let i=1;i<=rows.length;i++){
    const next=i<rows.length ? (rows[i].groupId || `ROW-${i}`) : "__END__";
    if(next!==id){ out.push({start,end:i-1,id}); start=i; id=next; }
  }
  return out;
}
function mergeProducts(ws:ExcelJS.Worksheet,rows:PackageRow[],excelStart:number){
  // Cada descrição não vazia inicia um subgrupo; linhas vazias seguintes são continuação física da mesma descrição.
  let start=-1;
  for(let i=0;i<=rows.length;i++){
    const has = i<rows.length && Boolean(rows[i].product?.trim());
    if(has){
      if(start>=0 && i-start>1) mergeSafe(ws,`${COL.product}${excelStart+start}:${COL.product}${excelStart+i-1}`);
      start=i;
    }
  }
  if(start>=0 && rows.length-start>1) mergeSafe(ws,`${COL.product}${excelStart+start}:${COL.product}${excelStart+rows.length-1}`);
}

export async function createRomaneioWorkbook(data:ProcessingResult){
  const template = await readFile(path.join(process.cwd(),"public","templates","Romaneio.xlsx"));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(Buffer.from(template) as any);
  wb.creator="Famossul | Gerador de Romaneios V6";
  const ws = wb.getWorksheet("Romaneio") || wb.worksheets[0];
  if(!ws) throw new Error("Modelo oficial Romaneio.xlsx inválido.");

  // NÃO altera labels, cores, bordas, larguras, logo ou estrutura do modelo.
  setValue(ws,"B7",new Date()); ws.getCell("B7").numFmt="dd/mm/yyyy";
  setValue(ws,"B8",data.client);
  setValue(ws,"B9",data.destination);
  setValue(ws,"B10",data.orderNumber);
  setValue(ws,"B11",data.delivery||"");
  // No modelo oficial os rótulos ficam em H7/H8/H9 e G11:J11. O valor entra em K.
  setValue(ws,"K7",data.orderOptions?.motorista||"");
  setValue(ws,"K8",data.orderOptions?.transportadora||"");
  setValue(ws,"K9",data.orderOptions?.placa||"");
  setValue(ws,"K11",data.orderOptions?.notaFiscal||"");

  unmergeDataArea(ws);
  clearDataValues(ws);

  let r=START_ROW;
  for(const pkg of data.packages){
    if(!pkg.rows?.length) continue;
    if(r+pkg.rows.length-1>LAST_DATA_ROW) throw new Error("O romaneio possui mais linhas do que o modelo oficial comporta. Divida o carregamento ou gere um segundo romaneio.");
    const pStart=r;
    for(const row of pkg.rows){
      setValue(ws,`${COL.games}${r}`,row.games??"");
      setValue(ws,`${COL.qty}${r}`,row.quantity||"");
      setValue(ws,`${COL.len}${r}`,row.lengthMm||"");
      setValue(ws,`${COL.wid}${r}`,row.widthMm||"");
      setValue(ws,`${COL.thk}${r}`,row.thicknessMm||"");
      // Mantém a fórmula original do modelo para m³, quando há dimensões; caso contrário usa valor vindo do documento.
      if(row.lengthMm && row.widthMm && row.thicknessMm && row.quantity){
        ws.getCell(`${COL.m3}${r}`).value={formula:`G${r}*F${r}*E${r}*D${r}/1000000000`,result:rowVolume(row)} as any;
      } else setValue(ws,`${COL.m3}${r}`,rowVolume(row)||0);
      ws.getCell(`${COL.m3}${r}`).numFmt="0.000";
      setValue(ws,`${COL.product}${r}`,row.product||"");
      setValue(ws,`${COL.obs}${r}`,itemText(row));
      ws.getCell(`${COL.product}${r}`).alignment={...(ws.getCell(`${COL.product}${r}`).alignment||{}),vertical:"middle",horizontal:"center",wrapText:true};
      ws.getCell(`${COL.obs}${r}`).alignment={...(ws.getCell(`${COL.obs}${r}`).alignment||{}),vertical:"middle",horizontal:"center",wrapText:true};
      r++;
    }
    const pEnd=r-1;
    setValue(ws,`${COL.pkg}${pStart}`,pkg.number);
    setValue(ws,`${COL.order}${pStart}`,data.orderNumber);
    setValue(ws,`${COL.palletM3}${pStart}`,pkg.totalVolume ?? pkg.rows.reduce((s,x)=>s+rowVolume(x),0));
    ws.getCell(`${COL.palletM3}${pStart}`).numFmt="0.000";
    if(pEnd>pStart){
      mergeSafe(ws,`${COL.pkg}${pStart}:${COL.pkg}${pEnd}`);
      mergeSafe(ws,`${COL.order}${pStart}:${COL.order}${pEnd}`);
      mergeSafe(ws,`${COL.palletM3}${pStart}:${COL.palletM3}${pEnd}`);
    }

    const local=pkg.rows;
    for(const g of groups(local)){
      const s=pStart+g.start,e=pStart+g.end;
      const games=local[g.start]?.games;
      if(e>s && games) mergeSafe(ws,`${COL.games}${s}:${COL.games}${e}`);
      const obsIndex=local.slice(g.start,g.end+1).findIndex(x=>Boolean(itemText(x)));
      if(e>s && obsIndex>=0){
        const master=s+obsIndex;
        if(master!==s) setValue(ws,`${COL.obs}${s}`,itemText(local[g.start+obsIndex]));
        mergeSafe(ws,`${COL.obs}${s}:${COL.obs}${e}`);
      }
      mergeProducts(ws,local.slice(g.start,g.end+1),s);
    }
  }

  for(let rr=r;rr<=LAST_DATA_ROW;rr++) ws.getRow(rr).hidden=true;

  // Total e rodapé permanecem nas linhas e com o estilo do arquivo original.
  ws.getCell(`C${TOTAL_ROW}`).value={formula:`SUM(C${START_ROW}:C${LAST_DATA_ROW})`} as any;
  ws.getCell(`D${TOTAL_ROW}`).value={formula:`SUM(D${START_ROW}:D${LAST_DATA_ROW})`} as any;
  setValue(ws,`G${TOTAL_ROW}`,"Total m³");
  ws.getCell(`H${TOTAL_ROW}`).value={formula:`SUM(H${START_ROW}:H${LAST_DATA_ROW})`} as any;
  ws.getCell(`H${TOTAL_ROW}`).numFmt="0.000";
  ws.getCell(`K${TOTAL_ROW}`).value={formula:`ROUNDUP(((H${TOTAL_ROW}*420)/1000),1)&\" - toneladas\"`} as any;

  setValue(ws,`A${FILTER_ROW}`,"Filtro:"); setValue(ws,`B${FILTER_ROW}`,data.orderOptions?.filtro||"");
  setValue(ws,`C${FILTER_ROW}`,"Pag.:"); setValue(ws,`D${FILTER_ROW}`,data.orderOptions?.pagina||"");
  setValue(ws,`A${OBS_ROW}`,"Obs:");
  const mount=MOUNT_LABELS[data.mountType]||data.mountType;
  const extra=[mount,data.config?.romaneioNote,data.orderOptions?.romaneioExtraText,...(data.config?.additionalItems||[]).filter(i=>i.enabled&&(i.target==="ROMANEIO"||i.target==="AMBOS")).map(i=>i.text||i.label),data.orderOptions?.conferente?`Conferente: ${data.orderOptions.conferente}`:"",data.orderOptions?.separador?`Separado por: ${data.orderOptions.separador}`:""].filter(Boolean).join(" | ");
  setValue(ws,`B${OBS_ROW}`,extra);
  ws.getCell(`B${OBS_ROW}`).alignment={...(ws.getCell(`B${OBS_ROW}`).alignment||{}),wrapText:true,vertical:"middle"};

  // Preserva a configuração de impressão existente no modelo oficial; apenas garante a área correta.
  ws.pageSetup.printArea="A1:K111";
  return Buffer.from(await wb.xlsx.writeBuffer());
}
