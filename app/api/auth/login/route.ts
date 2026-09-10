import { NextResponse } from "next/server";
import { z } from "zod";
import { createSession, validateCredentials } from "@/lib/auth";
export async function POST(req:Request){const parsed=z.object({user:z.string(),password:z.string()}).safeParse(await req.json());if(!parsed.success)return NextResponse.json({error:"Dados inválidos"},{status:400});if(!(await validateCredentials(parsed.data.user,parsed.data.password)))return NextResponse.json({error:"Credenciais inválidas"},{status:401});await createSession();return NextResponse.json({ok:true})}
