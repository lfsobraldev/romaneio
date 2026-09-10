import { NextResponse } from "next/server";
import { isAuthenticated } from "@/lib/auth";
import { createRomaneioWorkbook } from "@/lib/excel";
import type { ProcessingResult } from "@/lib/types";
export const runtime="nodejs";
export async function POST(req:Request){if(!(await isAuthenticated()))return new NextResponse("Não autorizado",{status:401});const data=await req.json() as ProcessingResult;const buf=await createRomaneioWorkbook(data);return new NextResponse(new Uint8Array(buf),{headers:{"content-type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","content-disposition":`attachment; filename=\"Romaneio-${data.orderNumber}.xlsx\"`}})}
