import { NextResponse } from "next/server";
import { isAuthenticated } from "@/lib/auth";
import { createLabelsPdf } from "@/lib/labels";
import type { ProcessingResult } from "@/lib/types";
export const runtime="nodejs";
export async function POST(req:Request){if(!(await isAuthenticated()))return new NextResponse("Não autorizado",{status:401});const data=await req.json() as ProcessingResult;const buf=await createLabelsPdf(data);return new NextResponse(new Uint8Array(buf),{headers:{"content-type":"application/pdf","content-disposition":`attachment; filename=\"Etiquetas-${data.orderNumber}.pdf\"`}})}
