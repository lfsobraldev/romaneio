import { PDFDocument, StandardFonts, rgb, type PDFFont, type PDFPage } from "pdf-lib";
import { readFile } from "node:fs/promises";
import path from "node:path";
import type { PackageRow, ProcessingResult } from "./types";
import { rowVolume } from "./logistics";

const BLACK=rgb(0,0,0); const W=684,H=367;
const fmtM3=(v?:number)=>(v||0).toFixed(3).replace(".",",");
const clean=(s?:string)=>(s||"").replace(/\s+/g," ").trim();

function textWidth(font:PDFFont,text:string,size:number){return font.widthOfTextAtSize(text,size)}
function fit(font:PDFFont,text:string,maxWidth:number,maxSize:number,minSize=4.2){let size=maxSize;while(size>minSize&&textWidth(font,text,size)>maxWidth)size-=.2;return size}
function drawCentered(page:PDFPage,font:PDFFont,text:string,x:number,width:number,y:number,size:number){const s=fit(font,text,width-4,size);const w=textWidth(font,text,s);page.drawText(text,{x:x+Math.max(2,(width-w)/2),y,size:s,font,color:BLACK})}
function wrap(font:PDFFont,text:string,size:number,maxWidth:number,maxLines=20){
 const words=clean(text).split(" ").filter(Boolean);const lines:string[]=[];let line="";
 for(const word of words){const n=line?`${line} ${word}`:word;if(textWidth(font,n,size)<=maxWidth)line=n;else{if(line)lines.push(line);line=word;if(lines.length>=maxLines-1)break}}
 if(line&&lines.length<maxLines)lines.push(line);return lines;
}
function itemObs(row:PackageRow){return clean([row.itemText,row.observation].filter(Boolean).join(" | "))}

function drawCellLines(page:PDFPage,font:PDFFont,lines:string[],x:number,width:number,topY:number,bottomY:number,size:number,center=true){
 if(!lines.length)return; const available=topY-bottomY; const step=Math.min(size+1.1,available/Math.max(lines.length,1)); const start=topY-step+1;
 lines.forEach((t,i)=>{const y=start-i*step;if(y<bottomY)return;const s=fit(font,t,width-4,size,3.8);if(center)drawCentered(page,font,t,x,width,y,s);else page.drawText(t,{x:x+2,y,size:s,font,color:BLACK})});
}

export async function createLabelsPdf(data:ProcessingResult){
 const base=await readFile(path.join(process.cwd(),"public","templates","etiqueta-base.png"));
 const doc=await PDFDocument.create(); const bg=await doc.embedPng(base);
 const regular=await doc.embedFont(StandardFonts.Helvetica),bold=await doc.embedFont(StandardFonts.HelveticaBold),boldItalic=await doc.embedFont(StandardFonts.HelveticaBoldOblique);
 const total=data.packages.length;
 const cfgExtra=(data.config?.additionalItems||[]).filter(i=>i.enabled&&(i.target==="ETIQUETA"||i.target==="AMBOS")).map(i=>i.text||i.label);
 const globalExtra=clean([data.config?.labelNote,data.orderOptions?.etiquetaExtraText,...cfgExtra].filter(Boolean).join(" | "));

 for(const pkg of data.packages){
  const page=doc.addPage([W,H]); page.drawImage(bg,{x:0,y:0,width:W,height:H});
  // Cabeçalho variável: coordenadas medidas diretamente no modelo 684x367.
  const order=String(data.orderNumber||"");
  page.drawText(order,{x:82,y:166,size:fit(boldItalic,order,130,13,8),font:boldItalic,color:BLACK});
  const destLines=wrap(boldItalic,data.destination||"-",11,318,2);destLines.forEach((l,i)=>page.drawText(l,{x:360,y:166-i*11,size:fit(boldItalic,l,318,11,7),font:boldItalic,color:BLACK}));
  drawCentered(page,bold,String(pkg.number),79,92,114,36);
  drawCentered(page,bold,String(total),230,125,114,36);
  drawCentered(page,bold,String(pkg.number),1,77,30,43);

  // A etiqueta oficial tem uma única área física de conteúdo. Todas as linhas do pallet são encaixadas nela.
  const rows=pkg.rows||[];
  const maxRows=Math.max(1,rows.length);
  const baseSize=maxRows<=4?7.1:maxRows<=7?6.0:maxRows<=10?5.1:4.2;
  const dataTop=84, dataBottom=25;
  const values={
    games: rows.map(r=>r.games?String(r.games):""),
    qty: rows.map(r=>r.quantity?String(r.quantity):""),
    len: rows.map(r=>r.lengthMm?String(r.lengthMm):""),
    wid: rows.map(r=>r.widthMm?String(r.widthMm):""),
    thk: rows.map(r=>r.thicknessMm?String(r.thicknessMm):""),
    m3: rows.map(r=>rowVolume(r)?fmtM3(rowVolume(r)):""),
    product: rows.map(r=>clean(r.product)||"↳ continuação"),
    obs: rows.map((r,i)=>clean([itemObs(r),i===0?pkg.notes:"",i===0?globalExtra:""].filter(Boolean).join(" | "))),
  };
  drawCellLines(page,regular,values.games,79,44,dataTop,dataBottom,baseSize,true);
  drawCellLines(page,regular,values.qty,123,48,dataTop,dataBottom,baseSize,true);
  drawCellLines(page,regular,values.len,171,45,dataTop,dataBottom,baseSize,true);
  drawCellLines(page,regular,values.wid,216,46,dataTop,dataBottom,baseSize,true);
  drawCellLines(page,regular,values.thk,262,45,dataTop,dataBottom,baseSize,true);
  drawCellLines(page,regular,values.m3,307,48,dataTop,dataBottom,Math.max(3.9,baseSize-.3),true);
  drawCellLines(page,regular,values.product,355,238,dataTop,dataBottom,Math.max(3.8,baseSize-.2),false);
  drawCellLines(page,regular,values.obs,593,91,dataTop,dataBottom,Math.max(3.8,baseSize-.5),false);

  const games=rows.reduce((s,r)=>s+(r.games||0),0)||pkg.games||0;
  drawCentered(page,bold,String(games||""),79,92,5,8.5);
  drawCentered(page,bold,fmtM3(pkg.totalVolume ?? rows.reduce((s,r)=>s+rowVolume(r),0)),307,48,5,9.5);
  drawCentered(page,bold,"0 - toneladas",356,237,5,9.5);
 }
 return Buffer.from(await doc.save());
}
