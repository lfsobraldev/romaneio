import { NextResponse } from "next/server";
import pdf from "pdf-parse";
import { buildProcessing } from "@/lib/parser";
import { isAuthenticated } from "@/lib/auth";
import { prisma } from "@/lib/prisma";
export const runtime="nodejs";
export async function POST(req:Request){
 if(!(await isAuthenticated()))return NextResponse.json({error:"Não autorizado"},{status:401});
 try{
   const fd=await req.formData();
   const order=fd.get("pedido"); const mach=fd.get("usinagem");
   if(!(order instanceof File))return NextResponse.json({error:"Envie o PDF do pedido."},{status:400});
   if(!order.name.toLowerCase().endsWith(".pdf"))return NextResponse.json({error:"O pedido precisa estar em PDF."},{status:400});
   const orderBuf=Buffer.from(await order.arrayBuffer());
   const parsed=await pdf(orderBuf);
   const machBuf=mach instanceof File?Buffer.from(await mach.arrayBuffer()):undefined;
   const result=buildProcessing(parsed.text,machBuf);
   try {
     await prisma.processing.create({data:{
       orderNumber:result.orderNumber, client:result.client, destination:result.destination,
       hasMachining:result.hasMachining, totalPackages:result.packages.length, sourceFileName:order.name,
       packages:{create:result.packages.map(pkg=>({number:pkg.number,games:pkg.games,totalVolume:pkg.totalVolume,notes:pkg.notes,
         rows:{create:pkg.rows.map(row=>({quantity:row.quantity,lengthMm:row.lengthMm,widthMm:row.widthMm,thicknessMm:row.thicknessMm,volume:row.volume,product:row.product,observation:row.observation}))}}))}
     }});
   } catch (dbError) { console.warn("Histórico não salvo:", dbError); }
   return NextResponse.json(result);
 }catch(e:any){console.error(e);return NextResponse.json({error:"Não foi possível interpretar os documentos. Confira os arquivos e tente novamente."},{status:500})}
}
