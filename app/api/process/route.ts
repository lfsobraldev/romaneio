import { NextResponse } from "next/server";
import pdf from "pdf-parse";
import { buildProcessing } from "@/lib/parser";
import { isAuthenticated } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
import { sanitizeConfig } from "@/lib/settings";
import type { LogisticsConfig, MountType } from "@/lib/types";
export const runtime="nodejs";

const allowedMounts = new Set<MountType>(["MONTADO_HS","MONTADO_TIMADEL","COMPLEMENTO_OBRA","REVENDA","MONTADO_ESTANCIA"]);

export async function POST(req:Request){
 if(!(await isAuthenticated()))return NextResponse.json({error:"Não autorizado"},{status:401});
 try{
   const fd=await req.formData();
   const order=fd.get("pedido"); const mach=fd.get("usinagem");
   if(!(order instanceof File))return NextResponse.json({error:"Envie o PDF do pedido."},{status:400});
   if(!order.name.toLowerCase().endsWith(".pdf"))return NextResponse.json({error:"O pedido precisa estar em PDF."},{status:400});

   const mountRaw=String(fd.get("mountType")||"MONTADO_HS") as MountType;
   const mountType:MountType=allowedMounts.has(mountRaw)?mountRaw:"MONTADO_HS";
   let config:LogisticsConfig|undefined;
   const configRaw=fd.get("config");
   if(typeof configRaw==="string"&&configRaw){try{config=sanitizeConfig(JSON.parse(configRaw));}catch{config=sanitizeConfig();}}

   const orderBuf=Buffer.from(await order.arrayBuffer());
   const parsed=await pdf(orderBuf);
   const machBuf=mach instanceof File?Buffer.from(await mach.arrayBuffer()):undefined;
   const mixedOrder=String(fd.get("mixedOrder")||"")==="true";
   const result=buildProcessing(parsed.text,machBuf,{mountType,config,mixedOrder});

   try {
     await prisma.processing.create({data:{
       orderNumber:result.orderNumber, client:result.client, destination:result.destination, delivery:result.delivery,
       hasMachining:result.hasMachining, status:mountType, totalPackages:result.packages.length, sourceFileName:order.name,
       packages:{create:result.packages.map(pkg=>({number:pkg.number,games:pkg.games,totalVolume:pkg.totalVolume,notes:[pkg.notes,pkg.ruleApplied].filter(Boolean).join(" | ")||undefined,
         rows:{create:pkg.rows.map(row=>({quantity:row.quantity,lengthMm:row.lengthMm,widthMm:row.widthMm,thicknessMm:row.thicknessMm,volume:row.volume,product:row.product,observation:row.observation}))}}))}
     }});
   } catch (dbError) { console.warn("Histórico não salvo:", dbError); }
   return NextResponse.json(result);
 }catch(e:any){console.error(e);return NextResponse.json({error:e?.message||"Não foi possível interpretar os documentos. Confira os arquivos e tente novamente."},{status:500})}
}
