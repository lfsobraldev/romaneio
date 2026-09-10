import Link from "next/link";
import { redirect } from "next/navigation";
import { ArrowLeft, PackageCheck } from "lucide-react";
import { isAuthenticated } from "@/lib/auth";
import { prisma } from "@/lib/prisma";

export const dynamic = "force-dynamic";
export default async function HistoryPage(){
  if(!(await isAuthenticated())) redirect("/login");
  let items: Awaited<ReturnType<typeof prisma.processing.findMany>> = [];
  let dbError = false;
  try { items = await prisma.processing.findMany({orderBy:{createdAt:"desc"},take:100}); }
  catch { dbError = true; }
  return <div className="shell"><header className="topbar"><div className="brand"><span className="leaf"/><div>famossul<small>Gerador de romaneios e etiquetas</small></div></div><Link className="btn ghost" href="/dashboard"><ArrowLeft size={16}/> Voltar</Link></header><main className="main"><div className="eyebrow">Rastreabilidade</div><h1 className="title">Histórico de processamentos</h1><p className="subtitle">Últimos pedidos interpretados pelo gerador.</p>{dbError&&<div className="warning">O histórico depende do Neon. Configure DATABASE_URL e execute <strong>npm run db:push</strong>.</div>}<div className="card recent">{!items.length?<div className="empty"><PackageCheck size={30} style={{marginBottom:10}}/><div>Nenhum processamento salvo ainda.</div></div>:<div className="tablewrap"><table className="table"><thead><tr><th>Pedido</th><th>Cliente</th><th>Destino</th><th>Pallets</th><th>Usinagem</th><th>Data</th></tr></thead><tbody>{items.map(i=><tr key={i.id}><td><strong>{i.orderNumber}</strong></td><td>{i.client}</td><td>{i.destination||"—"}</td><td>{i.totalPackages}</td><td>{i.hasMachining?"Sim":"Não"}</td><td>{new Intl.DateTimeFormat("pt-BR",{dateStyle:"short",timeStyle:"short",timeZone:"America/Maceio"}).format(i.createdAt)}</td></tr>)}</tbody></table></div>}</div></main></div>
}
