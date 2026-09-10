"use client";
import { useState } from "react";
import { useRouter } from "next/navigation";
export default function Login(){
 const [user,setUser]=useState(""); const [password,setPassword]=useState(""); const [error,setError]=useState(""); const [loading,setLoading]=useState(false); const router=useRouter();
 async function submit(e:React.FormEvent){e.preventDefault();setLoading(true);setError("");const r=await fetch("/api/auth/login",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({user,password})});setLoading(false);if(r.ok){router.push("/dashboard");router.refresh()}else setError("Usuário ou senha inválidos.")}
 return <main className="loginwrap"><div className="card login"><div className="brand"><span className="leaf"/><div>famossul<small>Portas para o hoje. E para o amanhã.</small></div></div><div style={{marginTop:32}}><div className="eyebrow">Acesso interno</div><h1 className="title" style={{fontSize:27}}>Gerador de romaneios</h1><p className="subtitle">Acesse para processar pedidos, conferir pallets e emitir etiquetas.</p></div><form onSubmit={submit}><div className="field"><label>Usuário</label><input value={user} onChange={e=>setUser(e.target.value)} autoComplete="username"/></div><div className="field"><label>Senha</label><input type="password" value={password} onChange={e=>setPassword(e.target.value)} autoComplete="current-password"/></div>{error&&<p className="error">{error}</p>}<button className="btn primary" disabled={loading}>{loading?"Entrando...":"Entrar"}</button></form></div></main>
}
