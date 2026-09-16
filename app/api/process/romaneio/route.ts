import { NextResponse } from "next/server";
import { isAuthenticated } from "@/lib/auth";
import { readReadyRomaneio } from "@/lib/romaneio-reader";
export const runtime = "nodejs";

export async function POST(req: Request) {
  if (!(await isAuthenticated())) return NextResponse.json({ error: "Não autorizado" }, { status: 401 });
  try {
    const fd = await req.formData();
    const file = fd.get("romaneio");
    if (!(file instanceof File)) return NextResponse.json({ error: "Envie o romaneio pronto." }, { status: 400 });
    if (!/\.(xlsx|xls)$/i.test(file.name)) return NextResponse.json({ error: "O romaneio precisa ser XLS ou XLSX." }, { status: 400 });
    const result = readReadyRomaneio(Buffer.from(await file.arrayBuffer()), file.name);
    return NextResponse.json(result);
  } catch (e:any) {
    console.error(e);
    return NextResponse.json({ error: e?.message || "Não foi possível ler o romaneio pronto." }, { status: 500 });
  }
}
