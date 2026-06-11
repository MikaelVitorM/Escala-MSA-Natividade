import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ── Supabase config ───────────────────────────────────────────────────────────
const SUPABASE_URL = "https://sdhxphydfowlftoiwjwt.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNkaHhwaHlkZm93bGZ0b2l3and0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEyMTA1MjAsImV4cCI6MjA5Njc4NjUyMH0._CFXjTwjx9I42MeRxIX7AJP5v3Qp5Cf1mdCSu1urC-0";

async function sb(path, opts={}) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: { "apikey": SUPABASE_KEY, "Authorization": `Bearer ${SUPABASE_KEY}`, "Content-Type": "application/json", "Prefer": "return=representation", ...opts.headers },
    ...opts
  });
  if (!res.ok) { const e = await res.text(); throw new Error(e); }
  const txt = await res.text();
  return txt ? JSON.parse(txt) : [];
}

const dbGet    = (table, qs="")          => sb(`${table}?${qs}`);
const dbInsert = (table, body)           => sb(table, { method:"POST", body: JSON.stringify(body) });
const dbUpsert = (table, body)           => sb(table, { method:"POST", headers:{"Prefer":"resolution=merge-duplicates,return=representation"}, body: JSON.stringify(body) });
const dbDelete = (table, qs)             => sb(`${table}?${qs}`, { method:"DELETE" });
const dbUpdate = (table, qs, body)       => sb(`${table}?${qs}`, { method:"PATCH", body: JSON.stringify(body) });

// ── constantes ────────────────────────────────────────────────────────────────
const IGREJAS = [
  "Paroquia Natividade - Presbitério","Paroquia Natividade - Capela do Santíssimo",
  "Menino Jesus de Praga","Nossa Senhora das Graças","São Sebastião",
  "São Pedro e São Paulo","Nossa Senhora Auxiliadora","Núcleo Santa Mônica","Núcleo Caminhando com Maria"
];
const IGREJAS_RESTRITAS = { "Paroquia Natividade - Capela do Santíssimo": ["Librifero","Credencial"] };
const HORARIOS_FIXOS = ["7h","8h","9h","18h","19h","19h30","20h"];
const FUNCOES_BASE = ["Cerimoniário Principal","Cerimoniário Regente","Librifero","Microfone","Credencial","Ceroferário","Turiferário","Naveteiro","Baculífero","Mitrífero"];
const FUNCOES_SEM_REPETICAO = ["Turiferário","Credencial"];

const isDomingo = d => d ? new Date(d+"T12:00:00").getDay()===0 : false;
const fmtData   = d => d ? new Date(d+"T12:00:00").toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long"}) : "";

const C = {
  vermelho:"#8B0000", ouro:"#C9A84C", ouroClaro:"#E8C96D",
  ouroPale:"#F5E6C0", branco:"#FFFFFF", brancoOff:"#FAF5EC",
};
const bgApp = {minHeight:"100vh",background:"linear-gradient(160deg,#6B0000 0%,#8B0000 40%,#5A0000 100%)",fontFamily:"Georgia,serif",color:C.brancoOff};

// ── lógica de sorteio ─────────────────────────────────────────────────────────
function getFuncoesNecessarias(dia) {
  if (IGREJAS_RESTRITAS[dia.igreja]) return [{nome:"Librifero",qtd:1},{nome:"Credencial",qtd:1}];
  const dom = isDomingo(dia.data);
  const funcs = [
    {nome:"Cerimoniário Principal",qtd:1},{nome:"Cerimoniário Regente",qtd:1},
    {nome:"Librifero",qtd:1},{nome:"Microfone",qtd:1},
    {nome:"Credencial",qtd:dom?3:2},{nome:"Ceroferário",qtd:2},
  ];
  if (dia.tem_incenso||dia.temIncenso) { funcs.find(f=>f.nome==="Ceroferário").qtd=6; funcs.push({nome:"Turiferário",qtd:1},{nome:"Naveteiro",qtd:1}); }
  else if (dom) { funcs.find(f=>f.nome==="Ceroferário").qtd=6; }
  if (dia.tem_bispo||dia.temBispo) { funcs.push({nome:"Baculífero",qtd:1},{nome:"Mitrífero",qtd:1}); }
  return funcs;
}

function sortearEscala(dias, membros, disponibilidades, historico) {
  const escHist={}, fnHist={};
  membros.forEach(m=>{ escHist[m.nome]=0; fnHist[m.nome]={}; });
  historico.forEach(h=>{ (h.resultado||[]).forEach(d=>{ Object.entries(d.alocados||{}).forEach(([nome,fn])=>{ escHist[nome]=(escHist[nome]||0)+1; fnHist[nome]=fnHist[nome]||{}; fnHist[nome][fn]=(fnHist[nome][fn]||0)+1; }); }); });
  const escAtual={};
  membros.forEach(m=>{ escAtual[m.nome]=0; });
  const escaladosDomingo=new Set();
  const resultado=[]; const observacoes=[];
  for (const dia of dias) {
    const funcs=getFuncoesNecessarias(dia);
    const dispDia=disponibilidades[dia.id]||[];
    const disponiveis=membros.filter(m=>dispDia.includes(m.nome));
    const dispFiltrados=isDomingo(dia.data)?disponiveis.filter(m=>!escaladosDomingo.has(m.nome)):disponiveis;
    const alocados={}; const usados=new Set();
    for (const {nome:fn,qtd} of funcs) {
      const semRep=FUNCOES_SEM_REPETICAO.includes(fn);
      let cands=dispFiltrados.filter(m=>m.funcoes.includes(fn)&&!usados.has(m.nome));
      if (semRep&&historico.length>0) {
        const ult=historico[historico.length-1];
        const fizeraUlt=new Set();
        (ult.resultado||[]).forEach(d=>Object.entries(d.alocados||{}).forEach(([nome,f])=>{ if(f===fn) fizeraUlt.add(nome); }));
        const sr=cands.filter(m=>!fizeraUlt.has(m.nome));
        if (sr.length>0) cands=sr;
        else observacoes.push({motivo:`Não foi possível evitar repetição de ${fn}.`});
      }
      cands.sort((a,b)=>{
        const pA=(a.score||0)*2-(escHist[a.nome]||0)*3-(escAtual[a.nome]||0)*5-((fnHist[a.nome]||{})[fn]||0)*2;
        const pB=(b.score||0)*2-(escHist[b.nome]||0)*3-(escAtual[b.nome]||0)*5-((fnHist[b.nome]||{})[fn]||0)*2;
        return pB-pA;
      });
      for (let i=0;i<qtd;i++) {
        if (!cands.length) { observacoes.push({motivo:`Sem candidatos para ${fn} em ${fmtData(dia.data)}.`}); break; }
        const pick=cands[0];
        alocados[pick.nome]=fn; usados.add(pick.nome); escAtual[pick.nome]=(escAtual[pick.nome]||0)+1;
        if (isDomingo(dia.data)) escaladosDomingo.add(pick.nome);
        cands=cands.slice(1);
      }
    }
    resultado.push({...dia,funcs,alocados});
  }
  const pendentes=membros.filter(m=>{ const d=Object.values(disponibilidades).filter(a=>a.includes(m.nome)).length; return d>0&&(escAtual[m.nome]||0)===0; }).map(m=>m.nome);
  return {resultado,observacoes,pendentes,escAtual};
}

// ── componentes base ──────────────────────────────────────────────────────────
function Tag({label,onRemove,color}) {
  const c=color||C.ouro;
  return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:`${c}25`,color:c,border:`1px solid ${c}60`,borderRadius:6,padding:"2px 8px",fontSize:12}}>
    {label}{onRemove&&<button onClick={onRemove} style={{background:"none",border:"none",cursor:"pointer",color:c,fontSize:14,lineHeight:1,padding:0}}>×</button>}
  </span>;
}
function Card({children,style={}}) { return <div style={{background:"rgba(0,0,0,0.25)",border:`1px solid ${C.ouro}35`,borderRadius:12,padding:18,...style}}>{children}</div>; }
function SecTitle({children}) { return <div style={{fontSize:11,color:C.ouro,letterSpacing:2,textTransform:"uppercase",marginBottom:12}}>{children}</div>; }
function Spinner() { return <div style={{textAlign:"center",padding:40,color:C.ouro,fontSize:13}}>⏳ Carregando…</div>; }

// ── TELA DO MEMBRO ────────────────────────────────────────────────────────────
function TelaMembro({onVoltar}) {
  const [membros,setMembros]=useState([]);
  const [dias,setDias]=useState([]);
  const [disps,setDisps]=useState({});
  const [nome,setNome]=useState("");
  const [busca,setBusca]=useState("");
  const [confirmado,setConfirmado]=useState(false);
  const [loading,setLoading]=useState(true);
  const [salvando,setSalvando]=useState(false);

  useEffect(()=>{
    (async()=>{
      try {
        const [ms,ds,dps]=await Promise.all([
          dbGet("membros","select=id,nome&order=nome"),
          dbGet("dias_escala","select=*&order=data,horario"),
          dbGet("disponibilidades","select=membro_id,dia_id")
        ]);
        setMembros(ms);
        setDias(ds);
        const map={};
        for (const d of ds) map[d.id]=[];
        for (const dp of dps) { if (map[dp.dia_id]) map[dp.dia_id].push(dp.membro_id); }
        setDisps(map);
      } catch(e){ console.error(e); }
      setLoading(false);
    })();
  },[]);

  const membroObj = membros.find(m=>m.nome===nome);
  const dispMembro = diaId => membroObj && (disps[diaId]||[]).includes(membroObj.id);
  const total = dias.filter(d=>dispMembro(d.id)).length;

  const toggleDisp = async(diaId) => {
    if (!membroObj) return;
    const marcado = dispMembro(diaId);
    setDisps(prev=>{ const a=prev[diaId]||[]; return {...prev,[diaId]:marcado?a.filter(x=>x!==membroObj.id):[...a,membroObj.id]}; });
    try {
      if (marcado) await dbDelete("disponibilidades",`membro_id=eq.${membroObj.id}&dia_id=eq.${diaId}`);
      else await dbUpsert("disponibilidades",[{membro_id:membroObj.id,dia_id:diaId,disponivel:true}]);
    } catch(e){ console.error(e); }
  };

  const confirmar = async() => { setSalvando(true); await new Promise(r=>setTimeout(r,500)); setSalvando(false); setConfirmado(true); };

  const filtrados=[...membros].sort((a,b)=>a.nome.localeCompare(b.nome)).filter(m=>m.nome.toLowerCase().includes(busca.toLowerCase()));

  if (loading) return <div style={{...bgApp,display:"flex",alignItems:"center",justifyContent:"center"}}><Spinner/></div>;

  if (!nome) return (
    <div style={{...bgApp,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{textAlign:"center",marginBottom:24}}><div style={{fontSize:32,color:C.ouro}}>✦</div>
        <h1 style={{color:C.branco,fontWeight:400,fontSize:22,margin:"4px 0",letterSpacing:3}}>Escala Litúrgica</h1>
        <div style={{width:80,height:1,background:`linear-gradient(90deg,transparent,${C.ouro},transparent)`,margin:"10px auto"}}/>
        <p style={{color:C.ouroPale,fontSize:13,margin:0}}>Selecione seu nome</p>
      </div>
      <div style={{width:"100%",maxWidth:420,background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}40`,borderRadius:16,padding:20}}>
        <input placeholder="🔍 Buscar…" value={busca} onChange={e=>setBusca(e.target.value)} style={{...inp,marginBottom:8,width:"100%",boxSizing:"border-box"}}/>
        <div style={{maxHeight:340,overflowY:"auto",display:"flex",flexDirection:"column",gap:3}}>
          {filtrados.map(m=><button key={m.id} onClick={()=>setNome(m.nome)} style={{textAlign:"left",padding:"10px 14px",border:`1px solid ${C.ouro}30`,borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:13,background:"rgba(255,255,255,0.07)",color:C.brancoOff}}>{m.nome}</button>)}
        </div>
      </div>
      {onVoltar&&<button onClick={onVoltar} style={{...btnSec,marginTop:14,fontSize:13}}>← Voltar</button>}
    </div>
  );

  if (confirmado) return (
    <div style={{...bgApp,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,textAlign:"center"}}>
      <div style={{fontSize:52}}>✅</div>
      <h2 style={{color:C.branco,fontWeight:400,fontSize:20,margin:"12px 0 0"}}>Confirmado!</h2>
      <p style={{color:C.ouroPale}}><strong style={{color:C.ouroClaro}}>{nome.split(" ")[0]}</strong> marcou <strong style={{color:C.ouroClaro}}>{total}</strong> dia(s).</p>
      <button onClick={()=>{setNome("");setBusca("");setConfirmado(false);}} style={{...btnPri,marginTop:16}}>Voltar ao início</button>
    </div>
  );

  return (
    <div style={{...bgApp,paddingBottom:48}}>
      <div style={{background:"rgba(0,0,0,0.3)",borderBottom:`1px solid ${C.ouro}40`,padding:"14px 20px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
        <div><div style={{fontSize:10,color:C.ouro,letterSpacing:3,textTransform:"uppercase"}}>Olá,</div><div style={{color:C.branco,fontSize:16,fontWeight:600}}>{nome}</div></div>
        <button onClick={()=>setNome("")} style={btnSec}>Trocar</button>
      </div>
      <div style={{maxWidth:640,margin:"0 auto",padding:"18px 14px"}}>
        <p style={{color:C.ouroPale,fontSize:13,marginBottom:14}}>Marque os dias disponíveis:</p>
        {!dias.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:40}}>Nenhum dia cadastrado.</div>}
        {dias.map(dia=>{const m=dispMembro(dia.id);return(
          <div key={dia.id} onClick={()=>toggleDisp(dia.id)} style={{background:m?`${C.ouro}20`:"rgba(0,0,0,0.2)",border:`2px solid ${m?C.ouro:C.ouro+"30"}`,borderRadius:12,padding:"13px 16px",marginBottom:9,cursor:"pointer",display:"flex",alignItems:"center",gap:14}}>
            <div style={{width:24,height:24,borderRadius:"50%",border:`2px solid ${m?C.ouro:C.ouro+"50"}`,background:m?C.ouro:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:13,color:C.vermelho,fontWeight:700}}>{m?"✓":""}</div>
            <div><div style={{color:C.branco,fontWeight:600,fontSize:13,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
              <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>{dia.horario} · {dia.igreja}{isDomingo(dia.data)&&<span style={{marginLeft:6,fontSize:10,color:C.ouro,background:`${C.ouro}25`,borderRadius:4,padding:"1px 5px"}}>Domingo</span>}{dia.tem_incenso&&" 🔥"}{dia.tem_bispo&&" 👑"}</div>
            </div>
          </div>
        );})}
        {dias.length>0&&<button onClick={confirmar} disabled={salvando} style={{...btnPri,width:"100%",marginTop:8,padding:13,fontSize:14,opacity:salvando?.7:1}}>
          {salvando?"⏳ Salvando…":`✓ Confirmar (${total} dia${total!==1?"s":""})`}
        </button>}
      </div>
    </div>
  );
}

// ── APP PRINCIPAL ─────────────────────────────────────────────────────────────
const VIEWS=["📅 Dias","👥 Membros","🔗 Link","📊 Estatísticas","🎲 Sorteio","📋 Histórico"];
const isMembro = window.location.pathname.includes("/responder") || window.location.search.includes("membro=1");

export default function App() {
  if (isMembro) return <TelaMembro/>;

  const [view,setView]=useState(0);
  const [simMembro,setSimMembro]=useState(false);
  const [dias,setDias]=useState([]);
  const [membros,setMembros]=useState([]);
  const [disponibilidades,setDisponibilidades]=useState({});
  const [historico,setHistorico]=useState([]);
  const [sorteioAtual,setSorteioAtual]=useState(null);
  const [formDia,setFormDia]=useState({data:"",horario:"",horarioCustom:"",usarCustom:false,igreja:"",temIncenso:false,temBispo:false});
  const [xlsxStatus,setXlsxStatus]=useState(null);
  const [xlsxMsg,setXlsxMsg]=useState("");
  const [nomeEscala,setNomeEscala]=useState("");
  const [loading,setLoading]=useState(true);
  const [salvando,setSalvando]=useState(false);

  // carregar dados do Supabase
  useEffect(()=>{
    (async()=>{
      try {
        const [ds,ms,dps,hs]=await Promise.all([
          dbGet("dias_escala","select=*&order=data,horario"),
          dbGet("membros","select=*&order=nome"),
          dbGet("disponibilidades","select=membro_id,dia_id"),
          dbGet("historico_escalas","select=*&order=criado_em.desc")
        ]);
        setDias(ds);
        setMembros(ms);
        setHistorico(hs.map(h=>({...h,resultado:h.resultado||[]})));
        const map={};
        for (const d of ds) map[d.id]=[];
        for (const dp of dps) {
          const m=ms.find(x=>x.id===dp.membro_id);
          if (m&&map[dp.dia_id]) map[dp.dia_id].push(m.nome);
        }
        setDisponibilidades(map);
      } catch(e){ console.error(e); }
      setLoading(false);
    })();
  },[]);

  const addDia=async()=>{
    const horario=formDia.usarCustom?formDia.horarioCustom:formDia.horario;
    if (!formDia.data||!horario||!formDia.igreja) return;
    setSalvando(true);
    try {
      const [novo]=await dbInsert("dias_escala",[{data:formDia.data,horario,igreja:formDia.igreja,tem_incenso:formDia.temIncenso,tem_bispo:formDia.temBispo}]);
      setDias(p=>[...p,novo]);
      setDisponibilidades(p=>({...p,[novo.id]:[]}));
      setFormDia({data:"",horario:"",horarioCustom:"",usarCustom:false,igreja:"",temIncenso:false,temBispo:false});
    } catch(e){ alert("Erro ao salvar dia: "+e.message); }
    setSalvando(false);
  };

  const removerDia=async(id)=>{
    await dbDelete("dias_escala",`id=eq.${id}`);
    setDias(p=>p.filter(d=>d.id!==id));
    setDisponibilidades(p=>{const n={...p};delete n[id];return n;});
  };

  const importarXlsx=async(e)=>{
    const file=e.target.files?.[0]; if (!file) return;
    setXlsxStatus("carregando");
    try {
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
      if (!rows.length) { setXlsxStatus("erro");setXlsxMsg("Planilha vazia.");return; }
      const keys=Object.keys(rows[0]);
      const nomeKey=keys.find(k=>k.toLowerCase().includes("nome"))||keys[0];
      const scoreKey=keys.find(k=>k.toLowerCase().includes("score")||k.toLowerCase().includes("escore"));
      const funcCols=keys.filter(k=>k!==nomeKey&&k!==scoreKey);
      const novos=[];
      rows.forEach(row=>{
        const nome=String(row[nomeKey]||"").trim(); if (!nome) return;
        const score=scoreKey?parseFloat(String(row[scoreKey]||"0").replace(",","."))||0:0;
        const funcoes=funcCols.filter(col=>{const v=String(row[col]||"").trim().toLowerCase();return v==="sim"||v==="s"||v==="yes"||v==="x"||v==="1";})
          .map(col=>FUNCOES_BASE.find(f=>f.toLowerCase()===col.toLowerCase())||col).filter(f=>FUNCOES_BASE.includes(f));
        novos.push({nome,funcoes,score});
      });
      // upsert no Supabase por nome
      await dbDelete("membros","id=neq.00000000-0000-0000-0000-000000000000"); // limpa todos
      const inseridos=await dbInsert("membros",novos.map(m=>({nome:m.nome,funcoes:m.funcoes,score:m.score})));
      setMembros(inseridos);
      setXlsxStatus("ok"); setXlsxMsg(`${inseridos.length} membro(s) importado(s).`);
    } catch(e){ setXlsxStatus("erro");setXlsxMsg("Erro: "+e.message); }
    e.target.value="";
  };

  const baixarModelo=()=>{
    const header=["Nome","Score",...FUNCOES_BASE];
    const ex=[["Maria da Silva",85,"Sim","Não","Sim","Sim","Não","Não","Não","Não","Não","Não"]];
    const ws=XLSX.utils.aoa_to_sheet([header,...ex]);
    const wb={SheetNames:["Membros"],Sheets:{Membros:ws}};
    XLSX.writeFile(wb,"modelo-escala.xlsx");
  };

  const gerarSorteio=()=>{
    const r=sortearEscala(dias,membros,disponibilidades,historico);
    setSorteioAtual(r); setView(4);
  };

  const salvarHistorico=async()=>{
    if (!sorteioAtual) return;
    setSalvando(true);
    try {
      const [novo]=await dbInsert("historico_escalas",[{nome:nomeEscala||`Escala ${new Date().toLocaleDateString("pt-BR")}`,resultado:sorteioAtual.resultado,pendentes:sorteioAtual.pendentes}]);
      setHistorico(p=>[novo,...p]);
      alert("Escala salva!");
    } catch(e){ alert("Erro: "+e.message); }
    setSalvando(false);
  };

  const excluirHistorico=async(id)=>{
    await dbDelete("historico_escalas",`id=eq.${id}`);
    setHistorico(p=>p.filter(h=>h.id!==id));
  };

  // link do membro
  const linkMembro = `${window.location.origin}${window.location.pathname}?membro=1`;

  // estatísticas
  const statsPre=useMemo_pre(dias,membros,disponibilidades);
  const statsPos=sorteioAtual?useMemo_pos(sorteioAtual,membros,disponibilidades):null;

  if (simMembro) return <TelaMembro onVoltar={()=>setSimMembro(false)}/>;
  if (loading) return <div style={bgApp}><Spinner/></div>;

  return (
    <div style={bgApp}>
      <div style={{textAlign:"center",padding:"24px 24px 14px",borderBottom:`1px solid ${C.ouro}40`,background:"rgba(0,0,0,0.2)"}}>
        <div style={{fontSize:10,letterSpacing:6,color:C.ouro,textTransform:"uppercase",marginBottom:5}}>Painel do Organizador</div>
        <h1 style={{margin:0,fontSize:"clamp(18px,4vw,30px)",fontWeight:400,color:C.branco,letterSpacing:3}}>
          <span style={{color:C.ouro}}>✦</span> Escala Litúrgica <span style={{color:C.ouro}}>✦</span>
        </h1>
        <div style={{width:100,height:1,background:`linear-gradient(90deg,transparent,${C.ouro},transparent)`,margin:"10px auto 0"}}/>
      </div>
      <div style={{display:"flex",justifyContent:"center",flexWrap:"wrap",gap:3,padding:"12px 12px 0"}}>
        {VIEWS.map((v,i)=>(
          <button key={v} onClick={()=>setView(i)} style={{padding:"7px 13px",border:`1px solid ${view===i?C.ouro:C.ouro+"40"}`,borderRadius:7,cursor:"pointer",fontFamily:"inherit",fontSize:11,background:view===i?`linear-gradient(135deg,${C.ouro},#A07828)`:"rgba(0,0,0,0.2)",color:view===i?C.vermelho:C.ouroPale,fontWeight:view===i?"700":"400"}}>{v}</button>
        ))}
      </div>

      <div style={{maxWidth:960,margin:"0 auto",padding:"18px 14px 48px"}}>

        {/* ── DIAS ── */}
        {view===0&&(
          <div>
            <p style={{color:C.ouroPale,fontSize:13,marginBottom:14}}>Credenciais: <strong style={{color:C.ouroClaro}}>2</strong> dias de semana · <strong style={{color:C.ouroClaro}}>3</strong> domingos.</p>
            <Card style={{marginBottom:16}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(145px,1fr))",gap:10,marginBottom:10}}>
                <div>
                  <label style={lbl}>Data</label>
                  <input type="date" value={formDia.data} onChange={e=>setFormDia(p=>({...p,data:e.target.value}))} style={inp}/>
                  {formDia.data&&<div style={{fontSize:10,color:isDomingo(formDia.data)?C.ouro:C.ouroPale,marginTop:3}}>{isDomingo(formDia.data)?"☀️ Domingo — 3 cred.":"📅 Semana — 2 cred."}</div>}
                </div>
                <div>
                  <label style={lbl}>Horário</label>
                  <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:4}}>
                    {HORARIOS_FIXOS.map(h=>(
                      <button key={h} onClick={()=>setFormDia(p=>({...p,horario:h,usarCustom:false}))} style={{padding:"4px 8px",border:`1px solid ${formDia.horario===h&&!formDia.usarCustom?C.ouro:C.ouro+"40"}`,borderRadius:5,cursor:"pointer",fontFamily:"inherit",fontSize:11,background:formDia.horario===h&&!formDia.usarCustom?`linear-gradient(135deg,${C.ouro},#A07828)`:"rgba(0,0,0,0.2)",color:formDia.horario===h&&!formDia.usarCustom?C.vermelho:C.ouroPale}}>{h}</button>
                    ))}
                    <button onClick={()=>setFormDia(p=>({...p,usarCustom:true,horario:""}))} style={{padding:"4px 8px",border:`1px solid ${formDia.usarCustom?C.ouro:C.ouro+"40"}`,borderRadius:5,cursor:"pointer",fontFamily:"inherit",fontSize:11,background:formDia.usarCustom?`linear-gradient(135deg,${C.ouro},#A07828)`:"rgba(0,0,0,0.2)",color:formDia.usarCustom?C.vermelho:C.ouroPale}}>Outro</button>
                  </div>
                  {formDia.usarCustom&&<input type="time" value={formDia.horarioCustom} onChange={e=>setFormDia(p=>({...p,horarioCustom:e.target.value}))} style={inp}/>}
                </div>
                <div>
                  <label style={lbl}>Igreja</label>
                  <select value={formDia.igreja} onChange={e=>setFormDia(p=>({...p,igreja:e.target.value}))} style={inp}>
                    <option value="">Selecione…</option>
                    {IGREJAS.map(ig=><option key={ig}>{ig}</option>)}
                  </select>
                </div>
              </div>
              <div style={{display:"flex",gap:20,marginBottom:12}}>
                {[["temIncenso","🔥 Com incenso"],["temBispo","👑 Com bispo"]].map(([k,t])=>(
                  <label key={k} style={{display:"flex",alignItems:"center",gap:7,cursor:"pointer",fontSize:12,color:C.ouroPale}}>
                    <input type="checkbox" checked={formDia[k]} onChange={e=>setFormDia(p=>({...p,[k]:e.target.checked}))} style={{accentColor:C.ouro,width:15,height:15}}/>{t}
                  </label>
                ))}
              </div>
              <button onClick={addDia} disabled={salvando} style={{...btnPri,opacity:salvando?.7:1}}>+ Adicionar dia</button>
            </Card>
            {!dias.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:32,fontSize:13}}>Nenhum dia adicionado.</div>}
            <div style={{display:"flex",flexDirection:"column",gap:7}}>
              {dias.map(dia=>(
                <div key={dia.id} style={{background:"rgba(0,0,0,0.2)",border:`1px solid ${C.ouro}30`,borderRadius:9,padding:"10px 14px",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:600,color:C.branco,fontSize:13,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                    <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>{dia.horario} · {dia.igreja}{isDomingo(dia.data)&&<span style={{marginLeft:5,fontSize:10,color:C.ouro,background:`${C.ouro}25`,borderRadius:3,padding:"1px 4px"}}>Dom</span>}{dia.tem_incenso&&" 🔥"}{dia.tem_bispo&&" 👑"}</div>
                  </div>
                  <div style={{fontSize:11,color:`${C.ouro}70`}}>{(disponibilidades[dia.id]||[]).length} disp.</div>
                  <button onClick={()=>removerDia(dia.id)} style={{background:"rgba(150,0,0,0.6)",border:`1px solid ${C.ouro}40`,color:C.ouroPale,borderRadius:5,padding:"3px 9px",cursor:"pointer",fontSize:11}}>✕</button>
                </div>
              ))}
            </div>
            {dias.length>0&&<div style={{textAlign:"right",marginTop:14}}><button onClick={()=>setView(1)} style={btnPri}>Próximo →</button></div>}
          </div>
        )}

        {/* ── MEMBROS ── */}
        {view===1&&(
          <div>
            <Card style={{marginBottom:14}}>
              <SecTitle>📂 Importar via planilha (.xlsx)</SecTitle>
              <p style={{color:C.ouroPale,fontSize:12,margin:"0 0 10px"}}>Colunas: <strong style={{color:C.ouroClaro}}>Nome</strong>, <strong style={{color:C.ouroClaro}}>Score</strong> e uma por função com <strong style={{color:C.ouroClaro}}>Sim/Não</strong>.</p>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,alignItems:"center"}}>
                <label style={{...btnPri,display:"inline-flex",alignItems:"center",gap:6,cursor:"pointer"}}>
                  📎 Selecionar .xlsx
                  <input type="file" accept=".xlsx,.xls" onChange={importarXlsx} style={{position:"absolute",opacity:0,width:0,height:0}}/>
                </label>
                <button onClick={baixarModelo} style={{...btnSec,fontSize:12,padding:"8px 14px"}}>⬇ Modelo</button>
                {xlsxStatus==="carregando"&&<span style={{color:C.ouro,fontSize:12}}>⏳ Importando…</span>}
                {xlsxStatus==="ok"&&<span style={{color:"#6fcf6f",fontSize:12}}>✅ {xlsxMsg}</span>}
                {xlsxStatus==="erro"&&<span style={{color:"#ff7044",fontSize:12}}>❌ {xlsxMsg}</span>}
              </div>
            </Card>
            {!membros.length
              ?<div style={{textAlign:"center",color:`${C.ouro}50`,padding:40,fontSize:13}}>Nenhum membro importado.</div>
              :<Card>
                <SecTitle>{membros.length} membro(s)</SecTitle>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:7,maxHeight:400,overflowY:"auto"}}>
                  {membros.map(m=>(
                    <div key={m.id} style={{background:"rgba(0,0,0,0.2)",border:`1px solid ${(m.funcoes||[]).length>0?C.ouro+"50":C.ouro+"20"}`,borderRadius:7,padding:"9px 11px"}}>
                      <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}>
                        <div style={{color:C.branco,fontSize:12,fontWeight:600}}>{m.nome.split(" ").slice(0,2).join(" ")}</div>
                        <span style={{fontSize:11,color:C.ouro,background:`${C.ouro}20`,borderRadius:4,padding:"1px 6px"}}>★ {m.score||0}</span>
                      </div>
                      <div style={{display:"flex",flexWrap:"wrap",gap:2}}>
                        {(m.funcoes||[]).length===0?<span style={{color:`${C.ouro}50`,fontSize:10}}>Sem funções</span>
                          :(m.funcoes||[]).map(f=><span key={f} style={{fontSize:9,background:`${C.ouro}20`,color:C.ouro,border:`1px solid ${C.ouro}40`,borderRadius:3,padding:"1px 4px"}}>{f}</span>)}
                      </div>
                    </div>
                  ))}
                </div>
              </Card>
            }
            <div style={{textAlign:"right",marginTop:14}}><button onClick={()=>setView(2)} style={btnPri}>Próximo →</button></div>
          </div>
        )}

        {/* ── LINK ── */}
        {view===2&&(
          <div>
            <Card style={{marginBottom:14,textAlign:"center"}}>
              <SecTitle>🔗 Link para os membros</SecTitle>
              <div style={{background:"rgba(0,0,0,0.3)",borderRadius:7,padding:"10px 14px",fontSize:12,color:C.ouroPale,fontFamily:"monospace",marginBottom:12,border:`1px solid ${C.ouro}30`,wordBreak:"break-all"}}>{linkMembro}</div>
              <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
                <button onClick={()=>navigator.clipboard.writeText(linkMembro)} style={btnPri}>📋 Copiar link</button>
                <button onClick={()=>setSimMembro(true)} style={btnSec}>👁 Simular</button>
              </div>
              <p style={{color:`${C.ouro}70`,fontSize:11,marginTop:10}}>Envie este link no grupo do WhatsApp. Cada membro abre e marca disponibilidade.</p>
            </Card>
            <Card>
              <SecTitle>📊 Respostas em tempo real</SecTitle>
              {!dias.length&&<div style={{color:`${C.ouro}50`,fontSize:12}}>Nenhum dia cadastrado.</div>}
              {dias.map(dia=>{const resp=disponibilidades[dia.id]||[];return(
                <div key={dia.id} style={{marginBottom:12,paddingBottom:12,borderBottom:`1px solid ${C.ouro}20`}}>
                  <div style={{color:C.branco,fontSize:12,fontWeight:600,textTransform:"capitalize"}}>{fmtData(dia.data)} · {dia.horario}</div>
                  <div style={{color:C.ouroPale,fontSize:11,marginBottom:5}}>{dia.igreja} — <strong style={{color:C.ouroClaro}}>{resp.length}</strong> resp.</div>
                  {!resp.length?<span style={{color:`${C.ouro}50`,fontSize:11}}>Sem respostas</span>
                    :<div style={{display:"flex",flexWrap:"wrap",gap:4}}>{resp.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")}/>)}</div>}
                </div>
              );})}
            </Card>
            <div style={{textAlign:"right",marginTop:14}}><button onClick={()=>setView(3)} style={btnPri}>Ver Estatísticas →</button></div>
          </div>
        )}

        {/* ── ESTATÍSTICAS ── */}
        {view===3&&statsPre&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <Card>
              <SecTitle>📊 Pré-sorteio</SecTitle>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:10,marginBottom:14}}>
                <div style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:12,textAlign:"center"}}>
                  <div style={{fontSize:22,color:C.ouroClaro,fontWeight:700}}>{statsPre.mediaDisp}</div>
                  <div style={{fontSize:11,color:C.ouroPale}}>média dias marcados</div>
                </div>
                <div style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:12,textAlign:"center"}}>
                  <div style={{fontSize:14,color:C.ouroClaro,fontWeight:700,textTransform:"capitalize"}}>{statsPre.maisDia?`${fmtData(statsPre.maisDia.data).split(",")[0]} ${statsPre.maisDia.horario}`:"—"}</div>
                  <div style={{fontSize:11,color:C.ouroPale}}>+ disponível ({statsPre.maisDia?.count||0})</div>
                </div>
                <div style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:12,textAlign:"center"}}>
                  <div style={{fontSize:14,color:"#ff9966",fontWeight:700,textTransform:"capitalize"}}>{statsPre.menosDia?`${fmtData(statsPre.menosDia.data).split(",")[0]} ${statsPre.menosDia.horario}`:"—"}</div>
                  <div style={{fontSize:11,color:C.ouroPale}}>- disponível ({statsPre.menosDia?.count||0})</div>
                </div>
              </div>
              <div style={{fontSize:11,color:C.ouro,marginBottom:7}}>Disponibilidade por membro</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:4,maxHeight:180,overflowY:"auto",marginBottom:12}}>
                {membros.sort((a,b)=>(statsPre.dispPorMembro[b.nome]||0)-(statsPre.dispPorMembro[a.nome]||0)).map(m=>(
                  <div key={m.id} style={{display:"flex",justifyContent:"space-between",background:"rgba(0,0,0,0.15)",borderRadius:5,padding:"4px 8px",fontSize:11}}>
                    <span style={{color:C.brancoOff}}>{m.nome.split(" ").slice(0,2).join(" ")}</span>
                    <span style={{color:C.ouro}}>{statsPre.dispPorMembro[m.nome]||0}d</span>
                  </div>
                ))}
              </div>
              {statsPre.naoPreencheram.length>0&&(
                <div>
                  <div style={{fontSize:11,color:"#ff9966",marginBottom:5}}>⚠️ Não preencheram ({statsPre.naoPreencheram.length})</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:4}}>{statsPre.naoPreencheram.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")} color="#ff9966"/>)}</div>
                </div>
              )}
            </Card>
            {statsPos&&(
              <Card>
                <SecTitle>🎯 Pós-sorteio</SecTitle>
                <div style={{marginBottom:8,fontSize:12,color:C.ouroPale}}>Média de escalações: <strong style={{color:C.ouroClaro}}>{statsPos.media}</strong></div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:4,maxHeight:220,overflowY:"auto",marginBottom:12}}>
                  {membros.filter(m=>statsPos.dispPor[m.nome]>0).sort((a,b)=>(statsPos.escaladoPor[b.nome]||0)-(statsPos.escaladoPor[a.nome]||0)).map(m=>{
                    const esc=statsPos.escaladoPor[m.nome]||0,disp=statsPos.dispPor[m.nome]||0,pct=disp>0?Math.round(esc/disp*100):0;
                    return(
                      <div key={m.id} style={{background:"rgba(0,0,0,0.15)",borderRadius:5,padding:"6px 9px",fontSize:11}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                          <span style={{color:C.brancoOff}}>{m.nome.split(" ").slice(0,2).join(" ")}</span>
                          <span style={{color:esc>0?C.ouroClaro:`${C.ouro}60`}}>{esc}/{disp}</span>
                        </div>
                        <div style={{height:3,background:`${C.ouro}20`,borderRadius:2}}><div style={{width:`${pct}%`,height:"100%",background:`linear-gradient(90deg,${C.ouro},${C.ouroClaro})`,borderRadius:2}}/></div>
                      </div>
                    );
                  })}
                </div>
                {statsPos.pendentes?.length>0&&(
                  <div>
                    <div style={{fontSize:11,color:"#88ddff",marginBottom:5}}>🔖 Prioridade próxima escala</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:4}}>{statsPos.pendentes.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")} color="#88ddff"/>)}</div>
                  </div>
                )}
              </Card>
            )}
            <div style={{textAlign:"right"}}><button onClick={()=>{gerarSorteio();}} style={btnPri}>🎲 Gerar Sorteio →</button></div>
          </div>
        )}

        {/* ── SORTEIO ── */}
        {view===4&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
              <SecTitle>🎲 Sorteio</SecTitle>
              <div style={{display:"flex",gap:7,flexWrap:"wrap"}}>
                <button onClick={gerarSorteio} style={btnPri}>🔀 {sorteioAtual?"Ressortear":"Gerar"}</button>
                {sorteioAtual&&<button onClick={salvarHistorico} disabled={salvando} style={{...btnSec,fontSize:12,opacity:salvando?.7:1}}>💾 Salvar</button>}
              </div>
            </div>
            {sorteioAtual&&<input placeholder="Nome desta escala (ex: Julho 2025)…" value={nomeEscala} onChange={e=>setNomeEscala(e.target.value)} style={{...inp,maxWidth:320,marginBottom:12}}/>}
            {sorteioAtual?.observacoes?.length>0&&(
              <Card style={{marginBottom:10,border:`1px solid #ff884450`}}>
                <SecTitle>⚠️ Observações</SecTitle>
                {sorteioAtual.observacoes.map((o,i)=><div key={i} style={{color:"#ffaa66",fontSize:12,marginBottom:3}}>• {o.motivo}</div>)}
              </Card>
            )}
            {sorteioAtual?.pendentes?.length>0&&(
              <Card style={{marginBottom:10,border:`1px solid #88ddff50`}}>
                <SecTitle>🔖 Pendentes próxima escala</SecTitle>
                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>{sorteioAtual.pendentes.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")} color="#88ddff"/>)}</div>
              </Card>
            )}
            {!sorteioAtual&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:50,fontSize:13}}>Clique em "Gerar" para sortear.</div>}
            {sorteioAtual?.resultado?.map((dia,idx)=>{
              const falta=dia.funcs.some(({nome:fn,qtd})=>Object.entries(dia.alocados||{}).filter(([,f])=>f===fn).length<qtd);
              return(
                <Card key={idx} style={{marginBottom:10,border:`1px solid ${falta?"#ff884460":C.ouro+"30"}`}}>
                  <div style={{marginBottom:10}}>
                    <div style={{fontSize:14,color:C.branco,fontWeight:600,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                    <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>{dia.horario} · {dia.igreja}{isDomingo(dia.data)&&<span style={{marginLeft:5,fontSize:10,color:C.ouro,background:`${C.ouro}25`,borderRadius:3,padding:"1px 4px"}}>Dom</span>}{(dia.tem_incenso||dia.temIncenso)&&" 🔥"}{(dia.tem_bispo||dia.temBispo)&&" 👑"}{falta&&<span style={{marginLeft:8,color:"#ffaa66",fontSize:11}}>⚠️ Vagas em aberto</span>}</div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(175px,1fr))",gap:6}}>
                    {dia.funcs.map(({nome:fn,qtd})=>{
                      const pessoas=Object.entries(dia.alocados||{}).filter(([,f])=>f===fn).map(([n])=>n);
                      const f=qtd-pessoas.length;
                      return(
                        <div key={fn} style={{background:f>0?"rgba(180,40,0,0.2)":`${C.ouro}08`,border:`1px solid ${f>0?"#ff440035":C.ouro+"35"}`,borderRadius:7,padding:"8px 10px"}}>
                          <div style={{fontSize:10,letterSpacing:1,color:C.ouro,marginBottom:4,textTransform:"uppercase"}}>{fn} ({qtd})</div>
                          {pessoas.map(p=>{const mb=membros.find(m=>m.nome===p);return(
                            <div key={p} style={{fontSize:11,color:C.ouroPale,marginBottom:2,display:"flex",justifyContent:"space-between"}}>
                              <span>✓ {p.split(" ").slice(0,2).join(" ")}</span>
                              {mb?.score>0&&<span style={{color:C.ouro,fontSize:10}}>★{mb.score}</span>}
                            </div>
                          );})}
                          {Array.from({length:f}).map((_,i)=><div key={i} style={{fontSize:11,color:"#ff8844",marginBottom:2}}>⚠ Sem candidato</div>)}
                        </div>
                      );
                    })}
                  </div>
                </Card>
              );
            })}
          </div>
        )}

        {/* ── HISTÓRICO ── */}
        {view===5&&(
          <div>
            <SecTitle>📋 Histórico de escalas</SecTitle>
            {!historico.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:40,fontSize:13}}>Nenhuma escala salva.</div>}
            {historico.map(h=>(
              <Card key={h.id} style={{marginBottom:10}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                  <div>
                    <div style={{color:C.branco,fontWeight:600,fontSize:13}}>{h.nome}</div>
                    <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>{new Date(h.criado_em).toLocaleDateString("pt-BR",{day:"2-digit",month:"long",year:"numeric"})}</div>
                  </div>
                  <button onClick={()=>excluirHistorico(h.id)} style={{background:"rgba(150,0,0,0.6)",border:`1px solid ${C.ouro}40`,color:C.ouroPale,borderRadius:5,padding:"3px 9px",cursor:"pointer",fontSize:11}}>Excluir</button>
                </div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4,marginBottom:6}}>
                  {(h.resultado||[]).map((dia,i)=>(
                    <span key={i} style={{fontSize:10,background:`${C.ouro}15`,color:C.ouro,border:`1px solid ${C.ouro}30`,borderRadius:4,padding:"2px 6px",textTransform:"capitalize"}}>
                      {fmtData(dia.data).split(",")[0]} {dia.horario}
                    </span>
                  ))}
                </div>
                {(h.pendentes||[]).length>0&&<div style={{fontSize:11,color:"#88ddff"}}>🔖 Pendentes: {h.pendentes.slice(0,4).map(n=>n.split(" ")[0]).join(", ")}{h.pendentes.length>4?` +${h.pendentes.length-4}`:""}</div>}
              </Card>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

// ── helpers de estatística ────────────────────────────────────────────────────
function useMemo_pre(dias,membros,disponibilidades) {
  const dispPorMembro={},dispPorDia={};
  membros.forEach(m=>{dispPorMembro[m.nome]=0;});
  dias.forEach(dia=>{const arr=disponibilidades[dia.id]||[];dispPorDia[dia.id]={...dia,count:arr.length};arr.forEach(n=>{dispPorMembro[n]=(dispPorMembro[n]||0)+1;});});
  const naoPreencheram=membros.filter(m=>!Object.values(disponibilidades).some(a=>a.includes(m.nome))).map(m=>m.nome);
  const vals=Object.values(dispPorDia).sort((a,b)=>b.count-a.count);
  const total=Object.values(dispPorMembro).reduce((s,v)=>s+v,0);
  return {dispPorMembro,naoPreencheram,maisDia:vals[0],menosDia:vals[vals.length-1],mediaDisp:membros.length>0?(total/membros.length).toFixed(1):0};
}
function useMemo_pos({resultado,pendentes},membros,disponibilidades) {
  const escaladoPor={},dispPor={};
  membros.forEach(m=>{escaladoPor[m.nome]=0;dispPor[m.nome]=Object.values(disponibilidades).filter(a=>a.includes(m.nome)).length;});
  resultado.forEach(dia=>Object.keys(dia.alocados||{}).forEach(n=>{escaladoPor[n]=(escaladoPor[n]||0)+1;}));
  const total=Object.values(escaladoPor).reduce((s,v)=>s+v,0);
  const count=Object.values(escaladoPor).filter(v=>v>0).length;
  return {escaladoPor,dispPor,media:count>0?(total/count).toFixed(1):0,pendentes};
}

const lbl={display:"block",fontSize:10,color:C.ouro,letterSpacing:1,marginBottom:5,textTransform:"uppercase"};
const inp={background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}50`,borderRadius:7,padding:"8px 11px",color:C.brancoOff,fontSize:13,fontFamily:"Georgia,serif",outline:"none",width:"100%",boxSizing:"border-box"};
const btnPri={background:`linear-gradient(135deg,${C.ouro},#A07828)`,border:"none",color:C.vermelho,borderRadius:7,padding:"9px 18px",cursor:"pointer",fontSize:13,fontFamily:"Georgia,serif",letterSpacing:.5,fontWeight:700};
const btnSec={background:"rgba(0,0,0,0.25)",border:`1px solid ${C.ouro}50`,color:C.ouroPale,borderRadius:7,padding:"9px 18px",cursor:"pointer",fontSize:13,fontFamily:"Georgia,serif"};