import { useState, useEffect, useMemo, useRef, createContext, useContext } from "react";
import * as XLSX from "xlsx";

const SUPABASE_URL = "https://sdhxphydfowlftoiwjwt.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNkaHhwaHlkZm93bGZ0b2l3and0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODEyMTA1MjAsImV4cCI6MjA5Njc4NjUyMH0._CFXjTwjx9I42MeRxIX7AJP5v3Qp5Cf1mdCSu1urC-0";
const BRASAO = "https://i.imgur.com/ejDFAD0.png";

// ── acesso ao Supabase ────────────────────────────────────────────────────────
async function sbFetch(path, opts = {}) {
  const { headers: extra = {}, ...rest } = opts;
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...rest,
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${SUPABASE_KEY}`,
      "Content-Type": "application/json",
      "Prefer": "return=representation",
      ...extra,
    },
  });
  if (!res.ok) { const e = await res.text(); throw new Error(e); }
  const txt = await res.text();
  return txt ? JSON.parse(txt) : [];
}

const dbGet    = (t, qs = "") => sbFetch(`${t}?${qs}`);
const dbInsert = (t, body)    => sbFetch(t, { method: "POST", body: JSON.stringify(body) });
const dbDelete = (t, qs)      => sbFetch(`${t}?${qs}`, { method: "DELETE" });
const dbPatch  = (t, qs, body) => sbFetch(`${t}?${qs}`, { method: "PATCH", body: JSON.stringify(body) });

// o Supabase devolve no máximo 1000 linhas por consulta — esta função busca tudo em lotes
async function dbGetAll(t, qs = "") {
  const todos = [];
  const passo = 1000;
  for (let off = 0; ; off += passo) {
    const lote = await dbGet(t, `${qs}&limit=${passo}&offset=${off}`);
    todos.push(...lote);
    if (lote.length < passo) break;
  }
  return todos;
}

async function dbUpsertDisp(membro_id, dia_id) {
  await fetch(`${SUPABASE_URL}/rest/v1/disponibilidades`, {
    method: "POST",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": `Bearer ${SUPABASE_KEY}`,
      "Content-Type": "application/json",
      "Prefer": "resolution=merge-duplicates",
    },
    body: JSON.stringify([{ membro_id, dia_id, disponivel: true }]),
  });
}

// ── constantes ────────────────────────────────────────────────────────────────
const IGREJAS = [
  "Paroquia Natividade - Presbitério","Paroquia Natividade - Capela do Santíssimo",
  "Menino Jesus de Praga","Nossa Senhora das Graças","São Sebastião",
  "São Pedro e São Paulo","Nossa Senhora Auxiliadora","Núcleo Santa Mônica",
  "Núcleo Caminhando com Maria","São José","Nossa Senhora de Guadalupe",
];
const IGREJAS_RESTRITAS = { "Paroquia Natividade - Capela do Santíssimo": ["Librifero","Credencial"] };
// comunidades com no máximo 4 escalados: 1 librifero, 1 microfone, 2 credenciais
const COMUNIDADES_4 = ["São José","Nossa Senhora de Guadalupe"];
const HORARIOS_FIXOS = ["7h","8h","9h","18h","19h","19h30","20h"];
const FUNCOES_BASE = [
  "Cerimoniário Principal","Cerimoniário Regente",
  "Librifero","Microfone","Credencial",
  "Ceroferário","Turiferário","Naveteiro",
  "Baculífero","Mitrífero",
];
const FUNCOES_SEM_REPETICAO = ["Turiferário","Credencial"];
const HIERARQUIA_SCORE = [
  "Credencial","Ceroferário","Microfone","Librifero",
  "Turiferário","Naveteiro","Baculífero","Mitrífero",
  "Cerimoniário Regente","Cerimoniário Principal",
];
const DIA_VAZIO = { data:"", horario:"", horarioCustom:"", usarCustom:false, igreja:"", temIncenso:false, temBispo:false, solenidade:false, nomeSolenidade:"" };

// ── pesos do ESCORE SUGERIDO (altere aqui se quiser mudar o cálculo) ──────────
// o maior peso é a presença em escalas; o menor é a presença em reuniões
const PESOS_ESCORE = {
  escala:  { presenca: 3, falta: -3, justificada: 0 },
  reuniao: { presenca: 1, falta: -1, justificada: 0 },
};

const STATUS = {
  presenca:    { label:"PRESENÇA",    bg:"#c6efce", fg:"#006100" },
  falta:       { label:"FALTA",       bg:"#ffc7ce", fg:"#9c0006" },
  justificada: { label:"JUSTIFICADA", bg:"#ffeb9c", fg:"#9c5700" },
};
const proximoStatus = s => s==="presenca" ? "falta" : s==="falta" ? "justificada" : s==="justificada" ? undefined : "presenca";

const norm = s => s ? s.normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim() : "";
const isDomingo = d => d ? new Date(d+"T12:00:00").getDay()===0 : false;
const fmtData   = d => d ? new Date(d+"T12:00:00").toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long"}) : "";
const fmtDataCurta = d => d ? new Date(d+"T12:00:00").toLocaleDateString("pt-BR") : "";
const fmtISO = iso => iso ? new Date(iso).toLocaleDateString("pt-BR") : "";
const diasDesde = iso => {
  const t = new Date(iso).getTime();
  if (!t) return null;
  return Math.max(0, Math.floor((Date.now() - t) / 86400000));
};

const ordenarEventos = arr => [...arr].sort((a,b)=>{
  if (a.data && b.data && a.data!==b.data) return a.data<b.data ? -1 : 1;
  if (a.data && !b.data) return -1;
  if (!a.data && b.data) return 1;
  return String(a.criado_em||"")<String(b.criado_em||"") ? -1 : 1;
});

const C = {
  vermelho:"#8B0000", ouro:"#C9A84C", ouroClaro:"#E8C96D",
  ouroPale:"#F5E6C0", branco:"#FFFFFF", brancoOff:"#FAF5EC",
};
const bgApp = { minHeight:"100vh", background:"linear-gradient(160deg,#6B0000 0%,#8B0000 40%,#5A0000 100%)", fontFamily:"Georgia,serif", color:C.brancoOff };

// ── temas por aba: Cadastro = preto, Escalas = vermelho, Reuniões e Eventos = dourado ──
const TEMAS = {
  cadastro: {
    fundo: "linear-gradient(160deg,#141414 0%,#000000 55%,#0a0a0a 100%)",
    titulo: "#FFFFFF", subtitulo: "#C9A84C", linha: "#C9A84C",
    abaAtiva: "#E8C96D", abaInativa: "#8d7f55", abaBorda: "#E8C96D",
    card: "rgba(255,255,255,0.05)",
  },
  escalas: {
    fundo: "linear-gradient(160deg,#6B0000 0%,#8B0000 40%,#5A0000 100%)",
    titulo: "#FFFFFF", subtitulo: "#C9A84C", linha: "#C9A84C",
    abaAtiva: "#FFFFFF", abaInativa: "#e0b9a0", abaBorda: "#C9A84C",
    card: "rgba(0,0,0,0.25)",
  },
  eventos: {
    fundo: "linear-gradient(160deg,#E6C463 0%,#C9A84C 45%,#B08F34 100%)",
    titulo: "#2b0000", subtitulo: "#5a0000", linha: "#5a0000",
    abaAtiva: "#2b0000", abaInativa: "#7a5a10", abaBorda: "#2b0000",
    card: "rgba(0,0,0,0.25)",
  },
};
const TemaCtx = createContext("escalas");

// ── funções necessárias por missa ─────────────────────────────────────────────
function getFuncoesNecessarias(dia) {
  if (COMUNIDADES_4.includes(dia.igreja)) {
    return [
      {nome:"Librifero",  qtd:1},
      {nome:"Microfone",  qtd:1},
      {nome:"Credencial", qtd:2},
    ];
  }
  if (IGREJAS_RESTRITAS[dia.igreja]) return [{nome:"Librifero",qtd:1},{nome:"Credencial",qtd:1}];
  const dom     = isDomingo(dia.data);
  const bispo   = dia.tem_bispo || dia.temBispo;
  const incenso = dia.tem_incenso || dia.temIncenso;
  const funcs = [
    {nome:"Cerimoniário Principal",qtd:1},
    {nome:"Cerimoniário Regente",  qtd:1},
    {nome:"Librifero",   qtd:1},
    {nome:"Microfone",   qtd:1},
    {nome:"Ceroferário", qtd:2},
    {nome:"Credencial",  qtd:dom?3:2},
  ];
  if (incenso || dom) {
    funcs.find(f=>f.nome==="Ceroferário").qtd = 6;
    if (incenso) funcs.push({nome:"Turiferário",qtd:1},{nome:"Naveteiro",qtd:1});
    if (dom && !incenso) funcs.find(f=>f.nome==="Credencial").qtd = 3;
  }
  if (bispo) funcs.push({nome:"Baculífero",qtd:1},{nome:"Mitrífero",qtd:1});
  return funcs;
}

// ── algoritmo de sorteio ──────────────────────────────────────────────────────
function sortearEscala(dias, membros, disponibilidades, historico, manualSlots={}) {
  const escHist={}, fnHist={}, bispoHist=new Set();
  membros.forEach(m=>{ escHist[m.nome]=0; fnHist[m.nome]={}; });
  historico.forEach(h=>(h.resultado||[]).forEach(d=>{
    const eraBispo=d.tem_bispo||d.temBispo;
    Object.entries(d.alocados||{}).forEach(([nome,fn])=>{
      escHist[nome]=(escHist[nome]||0)+1;
      fnHist[nome]=fnHist[nome]||{};
      fnHist[nome][fn]=(fnHist[nome][fn]||0)+1;
      if(eraBispo) bispoHist.add(nome);
    });
  }));

  const escAtual={};
  membros.forEach(m=>{ escAtual[m.nome]=0; });
  const escaladosDomingo={};
  const resultado=[], observacoes=[];

  for (const dia of dias) {
    const funcs=getFuncoesNecessarias(dia);
    const dispDia=disponibilidades[dia.id]||[];
    const disponiveis=membros.filter(m=>dispDia.includes(m.nome));
    const ehDomingo=isDomingo(dia.data);
    const ehBispo=dia.tem_bispo||dia.temBispo;

    if(ehDomingo) escaladosDomingo[dia.data]=escaladosDomingo[dia.data]||new Set();

    const manual=manualSlots[dia.id]||{};
    const alocados={...manual};
    const usados=new Set(Object.keys(manual));

    for (const {nome:fn, qtd} of funcs) {
      const jaAlocados=Object.entries(alocados).filter(([,f])=>f===fn).length;
      const faltam=qtd-jaAlocados;
      if(faltam<=0) continue;

      const temFuncao=m=>(m.funcoes||[]).some(f=>norm(f)===norm(fn));

      let cands=disponiveis.filter(m=>temFuncao(m)&&!usados.has(m.nome)&&(!ehDomingo||!escaladosDomingo[dia.data].has(m.nome)));

      if(ehBispo) cands=cands.filter(m=>!bispoHist.has(m.nome));

      if(FUNCOES_SEM_REPETICAO.some(f=>norm(f)===norm(fn))&&historico.length>0){
        const ult=historico[historico.length-1];
        const fizeraUlt=new Set();
        (ult.resultado||[]).forEach(d=>Object.entries(d.alocados||{}).forEach(([nome,f])=>{ if(norm(f)===norm(fn)) fizeraUlt.add(nome); }));
        const sr=cands.filter(m=>!fizeraUlt.has(m.nome));
        if(sr.length>0) cands=sr;
        else observacoes.push({motivo:`Repetição inevitável de ${fn}.`});
      }

      const ordenar=arr=>[...arr].sort((a,b)=>{
        const escA=(escHist[a.nome]||0)+(escAtual[a.nome]||0)*3;
        const escB=(escHist[b.nome]||0)+(escAtual[b.nome]||0)*3;
        if(escA!==escB) return escA-escB;
        const fnA=(fnHist[a.nome]||{})[fn]||0, fnB=(fnHist[b.nome]||{})[fn]||0;
        if(fnA!==fnB) return fnA-fnB;
        const hi=HIERARQUIA_SCORE.findIndex(x=>norm(x)===norm(fn));
        return hi>=5?(b.score||0)-(a.score||0):(a.score||0)-(b.score||0);
      });

      cands=ordenar(cands);

      for(let i=0;i<faltam;i++){
        if(!cands.length){
          const extras=ordenar(disponiveis.filter(m=>temFuncao(m)&&!usados.has(m.nome)));
          if(extras.length>0){
            const pick=extras[0];
            alocados[pick.nome]=fn; usados.add(pick.nome);
            escAtual[pick.nome]=(escAtual[pick.nome]||0)+1;
            if(ehDomingo) escaladosDomingo[dia.data].add(pick.nome);
            if(ehBispo) bispoHist.add(pick.nome);
            observacoes.push({motivo:`${pick.nome.split(" ")[0]} repetiu no domingo ${fmtData(dia.data)} - ${dia.horario} (necessário para ${fn}).`});
          } else {
            observacoes.push({motivo:`Sem candidatos para ${fn} em ${fmtData(dia.data)} ${dia.horario}.`});
          }
          break;
        }
        const pick=cands[0];
        alocados[pick.nome]=fn; usados.add(pick.nome);
        escAtual[pick.nome]=(escAtual[pick.nome]||0)+1;
        if(ehDomingo) escaladosDomingo[dia.data].add(pick.nome);
        if(ehBispo) bispoHist.add(pick.nome);
        cands=cands.slice(1);
      }
    }
    resultado.push({...dia,funcs,alocados});
  }

  const pendentes=membros.filter(m=>{
    const d=Object.values(disponibilidades).filter(a=>a.includes(m.nome)).length;
    return d>0&&(escAtual[m.nome]||0)===0;
  }).map(m=>({nome:m.nome,score:m.score||0}));

  const maxEsc=Math.max(...Object.values(escAtual),0);
  const minEsc=Math.min(...Object.values(escAtual).filter(v=>v>0),maxEsc);
  if(maxEsc-minEsc>1){
    const muitoEsc=Object.entries(escAtual).filter(([,v])=>v===maxEsc).map(([n])=>n.split(" ")[0]);
    const poucoEsc=Object.entries(escAtual).filter(([,v])=>v===minEsc&&v>0).map(([n])=>n.split(" ")[0]);
    if(muitoEsc.length&&poucoEsc.length)
      observacoes.push({motivo:`Atenção: ${muitoEsc.join(", ")} escalado(s) ${maxEsc}x; ${poucoEsc.join(", ")} apenas ${minEsc}x. Compensar na próxima escala.`});
  }

  return {resultado,observacoes,pendentes,escAtual};
}

// ── estatísticas ──────────────────────────────────────────────────────────────
function calcPre(dias,membros,disponibilidades){
  const dispPorMembro={},dispPorDia={};
  membros.forEach(m=>{dispPorMembro[m.nome]=0;});
  dias.forEach(dia=>{
    const arr=disponibilidades[dia.id]||[];
    dispPorDia[dia.id]={...dia,count:arr.length};
    arr.forEach(n=>{dispPorMembro[n]=(dispPorMembro[n]||0)+1;});
  });
  const naoPreencheram=membros.filter(m=>!Object.values(disponibilidades).some(a=>a.includes(m.nome))).map(m=>m.nome);
  const vals=Object.values(dispPorDia).sort((a,b)=>b.count-a.count);
  const total=Object.values(dispPorMembro).reduce((s,v)=>s+v,0);
  return{dispPorMembro,naoPreencheram,maisDia:vals[0],menosDia:vals[vals.length-1],mediaDisp:membros.length>0?(total/membros.length).toFixed(1):0};
}

function calcPos({resultado,pendentes},membros,disponibilidades){
  const escaladoPor={},dispPor={};
  membros.forEach(m=>{escaladoPor[m.nome]=0;dispPor[m.nome]=Object.values(disponibilidades).filter(a=>a.includes(m.nome)).length;});
  resultado.forEach(dia=>Object.keys(dia.alocados||{}).forEach(n=>{escaladoPor[n]=(escaladoPor[n]||0)+1;}));
  const total=Object.values(escaladoPor).reduce((s,v)=>s+v,0);
  const count=Object.values(escaladoPor).filter(v=>v>0).length;
  return{escaladoPor,dispPor,media:count>0?(total/count).toFixed(1):0,pendentes};
}

// estatísticas de cada membro: presenças em eventos, presenças nas escalas salvas e trocas
function calcStatsMembros(membros, historico, presMap) {
  const out = {};
  const idPorNome = {};
  membros.forEach(m => {
    out[m.id] = { pres:0, falta:0, just:0, ePres:0, eFalta:0, eJust:0, trocas:0, escalado:0, fnCount:{}, fnTop:null };
    idPorNome[norm(m.nome)] = m.id;
  });
  Object.entries(presMap).forEach(([k, st]) => {
    const s = out[k.split("|")[1]];
    if (!s) return;
    if (st === "presenca") s.pres++;
    else if (st === "falta") s.falta++;
    else if (st === "justificada") s.just++;
  });
  historico.forEach(h => (h.resultado || []).forEach(d => {
    Object.entries(d.alocados || {}).forEach(([nome, fn]) => {
      const s = out[idPorNome[norm(nome)]];
      if (!s) return;
      s.escalado++;
      s.fnCount[fn] = (s.fnCount[fn] || 0) + 1;
      const stEsc = (d.status || {})[nome];
      if (stEsc === "presenca") s.ePres++;
      else if (stEsc === "falta") s.eFalta++;
      else if (stEsc === "justificada") s.eJust++;
    });
    // trocas solicitadas = quem foi trocado (saiu da escala)
    const quemSaiu = Array.isArray(d.trocas) ? d.trocas.map(t => t.de) : Object.values(d.substituicoes || {});
    quemSaiu.forEach(nome => {
      const s = out[idPorNome[norm(nome)]];
      if (s) s.trocas++;
    });
  }));
  Object.values(out).forEach(s => {
    const top = Object.entries(s.fnCount).sort((a, b) => b[1] - a[1])[0];
    if (top) s.fnTop = { nome: top[0], qtd: top[1] };
  });
  return out;
}

// escore calculado a partir das presenças e faltas
function escoreSugerido(s) {
  const p = PESOS_ESCORE;
  const v =
    s.ePres * p.escala.presenca + s.eFalta * p.escala.falta + s.eJust * p.escala.justificada +
    s.pres  * p.reuniao.presenca + s.falta  * p.reuniao.falta + s.just  * p.reuniao.justificada;
  return Math.round(v * 10) / 10;
}

// o escore foi digitado manualmente? (sem a coluna nova, vale o comportamento antigo: escore diferente de 0 = manual)
function escoreManual(m) {
  if (m.score_editado === undefined) return (Number(m.score) || 0) !== 0;
  return m.score_editado === true;
}

// em quantas escalas cada pessoa se disponibilizou e qual foi a última
function calcDispStats(membros, historico, dias, disponibilidades, dispUltima) {
  const out = {};
  const idPorNome = {};
  membros.forEach(m => {
    out[m.id] = { escalasDisp:0, total:0, ultima:null };
    idPorNome[norm(m.nome)] = m.id;
  });
  const diaIdsAtuais = new Set(dias.map(d => d.id));
  const hist = [...historico].sort((a, b) => String(b.criado_em).localeCompare(String(a.criado_em)));
  let atualSalva = false;
  hist.forEach(h => {
    const ids = new Set();
    (h.resultado || []).forEach(d => {
      if (diaIdsAtuais.has(d.id)) atualSalva = true;
      // escalas salvas antes do registro de disponíveis: considera quem foi escalado
      const lista = Array.isArray(d.disponiveis)
        ? d.disponiveis
        : [...Object.keys(d.alocados || {}), ...Object.values(d.substituicoes || {})];
      lista.forEach(n => { const id = idPorNome[norm(n)]; if (id) ids.add(id); });
    });
    ids.forEach(id => {
      out[id].escalasDisp++;
      if (!out[id].ultima) out[id].ultima = { tipo:"salva", nome:h.nome, data:h.criado_em };
    });
  });
  const atualAberta = dias.length > 0 && !atualSalva;
  const total = hist.length + (atualAberta ? 1 : 0);
  membros.forEach(m => {
    const s = out[m.id];
    s.total = total;
    if (atualAberta && Object.values(disponibilidades).some(a => a.includes(m.nome))) {
      s.escalasDisp++;
      s.ultima = { tipo:"atual", nome:null, data:dispUltima[m.nome] || null };
    }
  });
  return out;
}

// resumo de presenças de uma escala salva
function resumoEscala(h) {
  let pres = 0, falta = 0, just = 0, total = 0;
  (h.resultado || []).forEach(d => Object.keys(d.alocados || {}).forEach(n => {
    total++;
    const s = (d.status || {})[n];
    if (s === "presenca") pres++;
    else if (s === "falta") falta++;
    else if (s === "justificada") just++;
  }));
  return { pres, falta, just, total, pend: total - pres - falta - just };
}

// escalados de uma missa, na ordem das funções
function ordenarEscalados(dia) {
  const ordem = (dia.funcs || []).map(f => f.nome);
  return Object.entries(dia.alocados || {}).sort((a, b) => {
    const ia = ordem.indexOf(a[1]), ib = ordem.indexOf(b[1]);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
  });
}

// ── componentes base ──────────────────────────────────────────────────────────
function Tag({label,onRemove,color}){
  const c=color||C.ouro;
  return <span style={{display:"inline-flex",alignItems:"center",gap:4,background:`${c}25`,color:c,border:`1px solid ${c}60`,borderRadius:6,padding:"2px 8px",fontSize:12}}>
    {label}{onRemove&&<button onClick={onRemove} style={{background:"none",border:"none",cursor:"pointer",color:c,fontSize:14,lineHeight:1,padding:0}}>×</button>}
  </span>;
}
function Card({children,style={}}){
  const tema=useContext(TemaCtx);
  return <div style={{background:TEMAS[tema].card,border:`1px solid ${C.ouro}35`,borderRadius:12,padding:18,...style}}>{children}</div>;
}
function SecTitle({children}){ return <div style={{fontSize:11,color:C.ouro,letterSpacing:2,textTransform:"uppercase",marginBottom:12}}>{children}</div>; }
function Spinner(){ return <div style={{textAlign:"center",padding:60,color:C.ouro,fontSize:14}}>⏳ Carregando…</div>; }
function Mini({label,valor,cor}){
  return <div style={{background:"rgba(0,0,0,0.25)",borderRadius:6,padding:"5px 4px",textAlign:"center"}}>
    <div style={{fontSize:16,fontWeight:700,color:cor}}>{valor}</div>
    <div style={{fontSize:9,color:C.ouroPale,letterSpacing:.5,textTransform:"uppercase"}}>{label}</div>
  </div>;
}

// botões PRESENÇA / FALTA / JUSTIFICADA (clicar de novo no ativo limpa a marcação)
function BotoesStatus({valor,onChange}){
  return(
    <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
      {Object.entries(STATUS).map(([k,cfg])=>{
        const ativo=valor===k;
        return(
          <button key={k} onClick={()=>onChange(ativo?undefined:k)} style={{border:`1px solid ${cfg.bg}`,background:ativo?cfg.bg:"transparent",color:ativo?cfg.fg:cfg.bg,borderRadius:5,padding:"3px 8px",fontSize:10,fontWeight:700,cursor:"pointer",fontFamily:"inherit",opacity:ativo?1:.85}}>
            {ativo?"✓ ":""}{cfg.label}
          </button>
        );
      })}
    </div>
  );
}

// formulário de cadastro / edição de membro
function FormMembro({inicial,onSalvar,onCancelar,salvando,obsAtiva}){
  const [nome,setNome]=useState(inicial?.nome||"");
  const [score,setScore]=useState(inicial&&inicial.scoreManual?String(inicial.score):"");
  const [funcoes,setFuncoes]=useState(inicial?.funcoes||[]);
  const [obs,setObs]=useState(inicial?.observacoes||"");
  const sug=inicial?.scoreSugerido??0;
  const tem=f=>funcoes.some(x=>norm(x)===norm(f));
  const toggle=f=>setFuncoes(p=>tem(f)?p.filter(x=>norm(x)!==norm(f)):[...p,f]);
  const salvar=()=>{
    if(!nome.trim()) return;
    const txt=String(score).trim().replace(",",".");
    const n=txt===""?null:parseFloat(txt);
    onSalvar({nome:nome.trim(),funcoes,score:(n===null||Number.isNaN(n))?null:n,observacoes:obs.trim()});
  };
  return(
    <div style={{background:"rgba(0,0,0,0.35)",border:`1px solid ${C.ouro}60`,borderRadius:12,padding:16}}>
      <SecTitle>{inicial?"✏️ Editar membro":"➕ Novo membro"}</SecTitle>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:10,marginBottom:10}}>
        <div>
          <label style={lbl}>Nome completo</label>
          <input value={nome} onChange={e=>setNome(e.target.value)} placeholder="Nome do membro" style={inp}/>
          {inicial&&<div style={{fontSize:10,color:`${C.ouro}80`,marginTop:3}}>Alterar o nome não atualiza escalas já salvas no histórico.</div>}
        </div>
        <div>
          <label style={lbl}>Escore (manual)</label>
          <input type="number" step="any" value={score} onChange={e=>setScore(e.target.value)} placeholder={`Automático (${sug})`} style={inp}/>
          <div style={{fontSize:10,color:`${C.ouro}80`,marginTop:3}}>Deixe em branco para usar o escore sugerido ({sug}). Se digitar um valor, ele prevalece.</div>
        </div>
      </div>
      <label style={lbl}>Funções que pode desempenhar</label>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:6,marginBottom:12}}>
        {FUNCOES_BASE.map(f=>(
          <label key={f} style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",fontSize:12,color:tem(f)?C.ouroClaro:C.ouroPale,background:tem(f)?`${C.ouro}20`:"rgba(0,0,0,0.2)",border:`1px solid ${tem(f)?C.ouro:C.ouro+"30"}`,borderRadius:6,padding:"6px 8px"}}>
            <input type="checkbox" checked={tem(f)} onChange={()=>toggle(f)} style={{accentColor:C.ouro}}/>{f}
          </label>
        ))}
      </div>
      <div style={{marginBottom:12}}>
        <label style={lbl}>Observações (opcional)</label>
        <textarea value={obs} onChange={e=>setObs(e.target.value)} disabled={!obsAtiva} rows={3}
          placeholder={obsAtiva?"Anotações sobre o membro…":"Rode o SQL das observações no Supabase para ativar este campo"}
          style={{...inp,resize:"vertical",fontFamily:"Georgia,serif",opacity:obsAtiva?1:.6}}/>
      </div>
      <div style={{display:"flex",gap:8}}>
        <button onClick={salvar} disabled={salvando||!nome.trim()} style={{...btnPri,opacity:(salvando||!nome.trim())?.6:1}}>{salvando?"⏳ Salvando…":"💾 Salvar"}</button>
        <button onClick={onCancelar} style={btnSec}>Cancelar</button>
      </div>
    </div>
  );
}

// ── TELA DO MEMBRO ────────────────────────────────────────────────────────────
function TelaMembro({onVoltar}){
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
      try{
        const [ms,ds,dps]=await Promise.all([
          dbGet("membros","select=id,nome&order=nome"),
          dbGet("dias_escala","select=*&order=data,horario"),
          dbGetAll("disponibilidades","select=membro_id,dia_id&order=id"),
        ]);
        setMembros(ms); setDias(ds);
        const map={};
        for(const d of ds) map[d.id]=[];
        for(const dp of dps){ if(map[dp.dia_id]) map[dp.dia_id].push(dp.membro_id); }
        setDisps(map);
      }catch(e){console.error(e);}
      setLoading(false);
    })();
  },[]);

  const membroObj=membros.find(m=>m.nome===nome);
  const dispMembro=diaId=>membroObj&&(disps[diaId]||[]).includes(membroObj.id);
  const total=dias.filter(d=>dispMembro(d.id)).length;

  const toggleDisp=async(diaId)=>{
    if(!membroObj)return;
    const marcado=dispMembro(diaId);
    setDisps(prev=>{const a=prev[diaId]||[];return{...prev,[diaId]:marcado?a.filter(x=>x!==membroObj.id):[...a,membroObj.id]};});
    try{
      if(marcado){
        await fetch(`${SUPABASE_URL}/rest/v1/disponibilidades?membro_id=eq.${membroObj.id}&dia_id=eq.${diaId}`,{
          method:"DELETE",headers:{"apikey":SUPABASE_KEY,"Authorization":`Bearer ${SUPABASE_KEY}`}
        });
      } else {
        await dbUpsertDisp(membroObj.id,diaId);
      }
    }catch(e){console.error(e);alert("Erro ao salvar: "+e.message);}
  };

  const confirmar=async()=>{setSalvando(true);await new Promise(r=>setTimeout(r,400));setSalvando(false);setConfirmado(true);};
  const filtrados=[...membros].sort((a,b)=>a.nome.localeCompare(b.nome)).filter(m=>m.nome.toLowerCase().includes(busca.toLowerCase()));

  if(loading) return <div style={{...bgApp,display:"flex",alignItems:"center",justifyContent:"center"}}><Spinner/></div>;

  if(!nome) return(
    <div style={{...bgApp,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{textAlign:"center",marginBottom:20}}>
        <img src={BRASAO} alt="Brasão" style={{width:80,height:80,objectFit:"contain",marginBottom:8,filter:"drop-shadow(0 2px 8px rgba(201,168,76,0.4))"}}/>
        <h1 style={{color:C.branco,fontWeight:400,fontStyle:"italic",fontSize:32,margin:"4px 0 0",letterSpacing:8,paddingLeft:8}}>ALTARE</h1>
        <div style={{fontSize:10,color:C.ouro,letterSpacing:4,textTransform:"uppercase",marginTop:2}}>Software Administrativo</div>
        <div style={{width:70,height:1,background:`linear-gradient(90deg,transparent,${C.ouro},transparent)`,margin:"10px auto"}}/>
        <p style={{color:C.ouroPale,fontSize:13,margin:0}}>Selecione seu nome</p>
      </div>
      <div style={{width:"100%",maxWidth:420,background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}40`,borderRadius:16,padding:18}}>
        <input placeholder="🔍 Buscar…" value={busca} onChange={e=>setBusca(e.target.value)} style={{...inp,marginBottom:8,width:"100%",boxSizing:"border-box"}}/>
        <div style={{maxHeight:340,overflowY:"auto",display:"flex",flexDirection:"column",gap:3}}>
          {filtrados.map(m=><button key={m.id} onClick={()=>setNome(m.nome)} style={{textAlign:"left",padding:"10px 14px",border:`1px solid ${C.ouro}30`,borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:13,background:"rgba(255,255,255,0.07)",color:C.brancoOff}}>{m.nome}</button>)}
        </div>
      </div>
      {onVoltar&&<button onClick={onVoltar} style={{...btnSec,marginTop:12,fontSize:12}}>← Voltar</button>}
    </div>
  );

  if(confirmado) return(
    <div style={{...bgApp,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,textAlign:"center"}}>
      <img src={BRASAO} alt="Brasão" style={{width:60,height:60,objectFit:"contain",marginBottom:10,opacity:.8}}/>
      <div style={{fontSize:50,marginBottom:6}}>✅</div>
      <h2 style={{color:C.branco,fontWeight:400,fontSize:19,margin:"0 0 6px"}}>Confirmado!</h2>
      <p style={{color:C.ouroPale,margin:0}}><strong style={{color:C.ouroClaro}}>{nome.split(" ")[0]}</strong> — <strong style={{color:C.ouroClaro}}>{total}</strong> dia(s) marcado(s).</p>
      <button onClick={()=>{setNome("");setBusca("");setConfirmado(false);}} style={{...btnPri,marginTop:16}}>Voltar ao início</button>
    </div>
  );

  return(
    <div style={{...bgApp,paddingBottom:48}}>
      <div style={{background:"rgba(0,0,0,0.3)",borderBottom:`1px solid ${C.ouro}40`,padding:"12px 18px",display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <img src={BRASAO} alt="Brasão" style={{width:36,height:36,objectFit:"contain"}}/>
          <div><div style={{fontSize:10,color:C.ouro,letterSpacing:2,textTransform:"uppercase"}}>Olá,</div><div style={{color:C.branco,fontSize:15,fontWeight:600}}>{nome}</div></div>
        </div>
        <button onClick={()=>setNome("")} style={btnSec}>Trocar</button>
      </div>
      <div style={{maxWidth:640,margin:"0 auto",padding:"16px 14px"}}>
        <p style={{color:C.ouroPale,fontSize:13,marginBottom:12}}>Marque os dias em que você está disponível:</p>
        {!dias.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:40}}>Nenhum dia cadastrado.</div>}
        {dias.map(dia=>{
          const m=dispMembro(dia.id);
          return(
            <div key={dia.id} onClick={()=>toggleDisp(dia.id)} style={{background:m?`${C.ouro}20`:"rgba(0,0,0,0.2)",border:`2px solid ${m?C.ouro:C.ouro+"30"}`,borderRadius:12,padding:"12px 16px",marginBottom:8,cursor:"pointer",display:"flex",alignItems:"center",gap:12}}>
              <div style={{width:24,height:24,borderRadius:"50%",border:`2px solid ${m?C.ouro:C.ouro+"50"}`,background:m?C.ouro:"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:13,color:C.vermelho,fontWeight:700}}>{m?"✓":""}</div>
              <div>
                <div style={{color:C.branco,fontWeight:600,fontSize:13,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>
                  {dia.horario} · {dia.igreja}
                  {isDomingo(dia.data)&&<span style={{marginLeft:5,fontSize:10,color:C.ouro,background:`${C.ouro}25`,borderRadius:3,padding:"1px 4px"}}>Domingo</span>}
                  {dia.solenidade&&<span style={{marginLeft:5,fontSize:10,color:"#f0c060",background:"rgba(240,192,96,0.2)",borderRadius:3,padding:"1px 4px"}}>✨ {dia.nome_solenidade||"Solenidade"}</span>}
                  {dia.tem_incenso&&" 🔥"}{dia.tem_bispo&&" 👑"}
                </div>
              </div>
            </div>
          );
        })}
        {dias.length>0&&<button onClick={confirmar} disabled={salvando} style={{...btnPri,width:"100%",marginTop:8,padding:12,fontSize:14,opacity:salvando?.7:1}}>{salvando?"⏳ Salvando…":`✓ Confirmar (${total} dia${total!==1?"s":""})`}</button>}
      </div>
    </div>
  );
}

// ── PAINEL DO ORGANIZADOR ─────────────────────────────────────────────────────
const ABAS=[["cadastro","CADASTRO"],["escalas","ESCALAS"],["eventos","REUNIÕES E EVENTOS"]];
const SUBS=[["dias","📅 Dias"],["link","🔗 Link"],["estatisticas","📊 Estatísticas"],["sorteio","🎲 Sorteio"],["historico","📋 Histórico"],["externas","🌐 EXTERNAS"]];

function Painel(){
  const [aba,setAba]=useState("escalas");
  const [sub,setSub]=useState("dias");
  const [simMembro,setSimMembro]=useState(false);

  const [dias,setDias]=useState([]);
  const [membros,setMembros]=useState([]);
  const [disponibilidades,setDisponibilidades]=useState({});
  const [dispUltima,setDispUltima]=useState({});
  const [dispDataOk,setDispDataOk]=useState(true);
  const [historico,setHistorico]=useState([]);
  const [eventos,setEventos]=useState([]);
  const [presMap,setPresMap]=useState({});
  const [eventosOk,setEventosOk]=useState(true);

  const [externos,setExternos]=useState([]);
  const [extPart,setExtPart]=useState({});
  const [externosOk,setExternosOk]=useState(true);
  const [formExterno,setFormExterno]=useState({nome:"",descricao:"",data:""});
  const [pickerId,setPickerId]=useState(null);
  const [pickerSel,setPickerSel]=useState([]);
  const [pickerBusca,setPickerBusca]=useState("");
  const [mostrarSemExt,setMostrarSemExt]=useState(false);

  const [sorteioAtual,setSorteioAtual]=useState(null);
  const [manualSlots,setManualSlots]=useState({});
  const [editandoManual,setEditandoManual]=useState(null);
  const [buscaSorteio,setBuscaSorteio]=useState("");
  const [nomeEscala,setNomeEscala]=useState("");
  const [formDia,setFormDia]=useState(DIA_VAZIO);

  const [formEvento,setFormEvento]=useState({nome:"",descricao:"",data:""});
  const [buscaMembro,setBuscaMembro]=useState("");
  const [ordemCadastro,setOrdemCadastro]=useState("nome");
  const [buscaEvento,setBuscaEvento]=useState("");
  const [editandoId,setEditandoId]=useState(null);
  const [xlsxStatus,setXlsxStatus]=useState(null);
  const [xlsxMsg,setXlsxMsg]=useState("");

  const [abertoId,setAbertoId]=useState(null);
  const [trocandoKey,setTrocandoKey]=useState(null);
  const filaRef=useRef(Promise.resolve());
  const colDispRef=useRef(true);

  const [loading,setLoading]=useState(true);
  const [salvando,setSalvando]=useState(false);

  // ── carregamento ──
  // busca as disponibilidades com a data de preenchimento (se a coluna já existir no banco)
  const buscarDisp=async()=>{
    if(colDispRef.current){
      try{
        const r=await dbGetAll("disponibilidades","select=membro_id,dia_id,criado_em&order=id");
        setDispDataOk(true);
        return r;
      }catch(e){
        if(!String(e.message).includes("criado_em")) throw e;
        colDispRef.current=false;
        setDispDataOk(false);
      }
    }
    return dbGetAll("disponibilidades","select=membro_id,dia_id&order=id");
  };

  const carregarDados=async(completo=false)=>{
    try{
      const [ds,ms,dps]=await Promise.all([
        dbGet("dias_escala","select=*&order=data,horario"),
        dbGet("membros","select=*&order=nome"),
        buscarDisp(),
      ]);
      setDias(ds); setMembros(ms);
      const map={};
      const ult={};
      for(const d of ds) map[d.id]=[];
      for(const dp of dps){
        const m=ms.find(x=>x.id===dp.membro_id);
        if(m&&map[dp.dia_id]){
          map[dp.dia_id].push(m.nome);
          if(dp.criado_em&&(!ult[m.nome]||dp.criado_em>ult[m.nome])) ult[m.nome]=dp.criado_em;
        }
      }
      setDisponibilidades(map);
      setDispUltima(ult);
    }catch(e){console.error(e);}

    if(completo){
      try{
        const hs=await dbGet("historico_escalas","select=*&order=criado_em.desc");
        setHistorico(hs.map(h=>({...h,resultado:h.resultado||[]})));
      }catch(e){console.error(e);}
      try{
        const [evs,prs]=await Promise.all([
          dbGet("eventos","select=*&order=criado_em"),
          dbGetAll("presencas","select=evento_id,membro_id,status&order=id"),
        ]);
        setEventos(ordenarEventos(evs));
        const mapa={};
        prs.forEach(p=>{mapa[`${p.evento_id}|${p.membro_id}`]=p.status;});
        setPresMap(mapa);
        setEventosOk(true);
      }catch(e){console.error(e);setEventosOk(false);}
      try{
        const [exs,eps]=await Promise.all([
          dbGet("externos","select=*&order=criado_em"),
          dbGetAll("externos_membros","select=externo_id,membro_id&order=id"),
        ]);
        setExternos(ordenarEventos(exs));
        const mp={};
        exs.forEach(x=>{mp[x.id]=[];});
        eps.forEach(p=>{if(mp[p.externo_id])mp[p.externo_id].push(p.membro_id);});
        setExtPart(mp);
        setExternosOk(true);
      }catch(e){console.error(e);setExternosOk(false);}
    }
    setLoading(false);
  };

  useEffect(()=>{
    carregarDados(true);
    const iv=setInterval(()=>carregarDados(false),15000);
    return()=>clearInterval(iv);
  },[]);

  // ── derivados (antes de qualquer retorno antecipado) ──
  const statsMembros=useMemo(()=>calcStatsMembros(membros,historico,presMap),[membros,historico,presMap]);

  // membros com o escore EM USO: o manual prevalece; sem manual vale o sugerido
  const membrosEf=useMemo(()=>membros.map(m=>{
    const s=statsMembros[m.id];
    const sugerido=s?escoreSugerido(s):0;
    const manual=escoreManual(m);
    return {...m,scoreManual:manual,scoreSugerido:sugerido,score:manual?(Number(m.score)||0):sugerido};
  }),[membros,statsMembros]);

  const dispStats=useMemo(()=>calcDispStats(membros,historico,dias,disponibilidades,dispUltima),[membros,historico,dias,disponibilidades,dispUltima]);

  // quantidade de eventos externos por membro
  const extCont=useMemo(()=>{
    const c={};
    Object.values(extPart).forEach(arr=>arr.forEach(id=>{c[id]=(c[id]||0)+1;}));
    return c;
  },[extPart]);

  const colScoreOk=membros.length===0||membros[0].score_editado!==undefined;
  const colObsOk=membros.length===0||membros[0].observacoes!==undefined;

  // ── dias da escala ──
  const addDia=async()=>{
    const horario=formDia.usarCustom?formDia.horarioCustom:formDia.horario;
    if(!formDia.data||!horario||!formDia.igreja)return;
    setSalvando(true);
    try{
      const [novo]=await dbInsert("dias_escala",[{data:formDia.data,horario,igreja:formDia.igreja,tem_incenso:formDia.temIncenso,tem_bispo:formDia.temBispo,solenidade:formDia.solenidade,nome_solenidade:formDia.nomeSolenidade||null}]);
      setDias(p=>[...p,novo]);
      setDisponibilidades(p=>({...p,[novo.id]:[]}));
      setFormDia(DIA_VAZIO);
    }catch(e){alert("Erro: "+e.message);}
    setSalvando(false);
  };

  const removerDia=async(id)=>{
    await dbDelete("dias_escala",`id=eq.${id}`);
    setDias(p=>p.filter(d=>d.id!==id));
    setDisponibilidades(p=>{const n={...p};delete n[id];return n;});
  };

  // ── cadastro de membros ──
  const salvarMembro=async(dados,id)=>{
    if(!id && membros.some(m=>norm(m.nome)===norm(dados.nome))){
      alert("Já existe um membro com esse nome.");
      return;
    }
    // score nulo = automático (usa o sugerido); com valor = manual (prevalece)
    const payload={nome:dados.nome,funcoes:dados.funcoes};
    if(dados.score===null){
      if(colScoreOk) payload.score_editado=false;
      else payload.score=0;
    } else {
      payload.score=dados.score;
      if(colScoreOk) payload.score_editado=true;
    }
    if(colObsOk){
      payload.observacoes=dados.observacoes||null;
    } else if(dados.observacoes){
      alert("Para salvar observações, rode o SQL das observações no Supabase e recarregue a página.");
      return;
    }
    setSalvando(true);
    try{
      if(id){
        const [atual]=await dbPatch("membros",`id=eq.${id}`,payload);
        setMembros(p=>p.map(m=>m.id===id?{...m,...(atual||payload)}:m));
      } else {
        const [novo]=await dbInsert("membros",[payload]);
        setMembros(p=>[...p,novo]);
      }
      setEditandoId(null);
    }catch(e){alert("Erro ao salvar membro: "+e.message);}
    setSalvando(false);
  };

  const usarEscoreSugerido=async(m)=>{
    try{
      await dbPatch("membros",`id=eq.${m.id}`,{score_editado:false});
      setMembros(p=>p.map(x=>x.id===m.id?{...x,score_editado:false}:x));
    }catch(e){alert("Erro: "+e.message);}
  };

  const usarSugeridoTodos=async()=>{
    if(!window.confirm("Usar o escore sugerido para TODOS os membros? Os escores digitados manualmente deixam de valer (o valor digitado continua guardado, mas não é usado)."))return;
    try{
      await dbPatch("membros","id=neq.00000000-0000-0000-0000-000000000000",{score_editado:false});
      setMembros(p=>p.map(x=>({...x,score_editado:false})));
    }catch(e){alert("Erro: "+e.message);}
  };

  const removerMembro=async(m)=>{
    if(!window.confirm(`Remover ${m.nome}? As disponibilidades, presenças e participações em eventos externos dessa pessoa também serão apagadas.`))return;
    try{
      await dbDelete("membros",`id=eq.${m.id}`);
      setMembros(p=>p.filter(x=>x.id!==m.id));
      setDisponibilidades(p=>{const n={};Object.entries(p).forEach(([k,v])=>{n[k]=v.filter(x=>x!==m.nome);});return n;});
      setPresMap(p=>{const n={};Object.entries(p).forEach(([k,v])=>{if(k.split("|")[1]!==m.id)n[k]=v;});return n;});
      setExtPart(p=>{const n={};Object.entries(p).forEach(([k,v])=>{n[k]=v.filter(id=>id!==m.id);});return n;});
    }catch(e){alert("Erro ao remover: "+e.message);}
  };

  const importarXlsx=async(e)=>{
    const file=e.target.files?.[0];if(!file)return;
    setXlsxStatus("carregando");
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
      if(!rows.length){setXlsxStatus("erro");setXlsxMsg("Planilha vazia.");return;}
      const keys=Object.keys(rows[0]);
      const nomeKey=keys.find(k=>k.toLowerCase().includes("nome"))||keys[0];
      const scoreKey=keys.find(k=>k.toLowerCase().includes("score")||k.toLowerCase().includes("escore"));
      const funcCols=keys.filter(k=>k!==nomeKey&&k!==scoreKey);
      const novos=[];
      rows.forEach(row=>{
        const nome=String(row[nomeKey]||"").trim();if(!nome)return;
        const txtScore=scoreKey?String(row[scoreKey]).trim():"";
        const temScore=txtScore!=="";
        const score=temScore?(parseFloat(txtScore.replace(",","."))||0):0;
        const funcoes=funcCols
          .filter(col=>{const v=String(row[col]||"").trim().toLowerCase();return v==="sim"||v==="s"||v==="yes"||v==="x"||v==="1";})
          .map(col=>FUNCOES_BASE.find(f=>norm(f)===norm(col))||col)
          .filter(f=>FUNCOES_BASE.some(fb=>norm(fb)===norm(f)));
        novos.push({nome,funcoes,score,temScore});
      });

      // atualiza quem já existe (pelo nome) e cria quem é novo — sem apagar ninguém
      const existentes=await dbGetAll("membros","select=id,nome&order=id");
      const mapa={};
      existentes.forEach(m=>{mapa[norm(m.nome)]=m.id;});
      const patches=[];
      const inserir=[];
      novos.forEach(m=>{
        const id=mapa[norm(m.nome)];
        if(id){
          const corpo={funcoes:m.funcoes};
          if(m.temScore){
            corpo.score=m.score;
            if(colScoreOk) corpo.score_editado=true;
          }
          patches.push(dbPatch("membros",`id=eq.${id}`,corpo));
        } else {
          const novo={nome:m.nome,funcoes:m.funcoes,score:m.score};
          if(colScoreOk) novo.score_editado=m.temScore;
          inserir.push(novo);
        }
      });
      await Promise.all(patches);
      if(inserir.length) await dbInsert("membros",inserir);
      await carregarDados(false);
      setXlsxStatus("ok");
      setXlsxMsg(`${patches.length} atualizado(s), ${inserir.length} novo(s).`);
    }catch(err){setXlsxStatus("erro");setXlsxMsg("Erro: "+err.message);}
    e.target.value="";
  };

  const baixarModelo=()=>{
    const header=["Nome","Score",...FUNCOES_BASE];
    const ex=[["Maria da Silva",85,"Sim","Não","Sim","Sim","Não","Não","Não","Não","Não","Não"]];
    const ws=XLSX.utils.aoa_to_sheet([header,...ex]);
    const wb={SheetNames:["Membros"],Sheets:{Membros:ws}};
    XLSX.writeFile(wb,"modelo-escala.xlsx");
  };

  // ── eventos e presenças ──
  const addEvento=async()=>{
    if(!formEvento.nome.trim())return;
    setSalvando(true);
    try{
      const [novo]=await dbInsert("eventos",[{nome:formEvento.nome.trim(),descricao:formEvento.descricao.trim()||null,data:formEvento.data||null}]);
      setEventos(p=>ordenarEventos([...p,novo]));
      setFormEvento({nome:"",descricao:"",data:""});
    }catch(e){alert("Erro ao salvar evento: "+e.message);}
    setSalvando(false);
  };

  const removerEvento=async(ev)=>{
    if(!window.confirm(`Remover o evento "${ev.nome}" e todas as presenças registradas nele?`))return;
    try{
      await dbDelete("eventos",`id=eq.${ev.id}`);
      setEventos(p=>p.filter(x=>x.id!==ev.id));
      setPresMap(p=>{const n={};Object.entries(p).forEach(([k,v])=>{if(!k.startsWith(ev.id+"|"))n[k]=v;});return n;});
    }catch(e){alert("Erro ao remover: "+e.message);}
  };

  const definirStatus=async(eventoId,membroId,novo)=>{
    const k=`${eventoId}|${membroId}`;
    const anterior=presMap[k];
    setPresMap(p=>{const n={...p};if(novo)n[k]=novo;else delete n[k];return n;});
    try{
      if(novo){
        await sbFetch("presencas?on_conflict=evento_id,membro_id",{
          method:"POST",
          headers:{Prefer:"resolution=merge-duplicates,return=representation"},
          body:JSON.stringify([{evento_id:eventoId,membro_id:membroId,status:novo}]),
        });
      } else {
        await dbDelete("presencas",`evento_id=eq.${eventoId}&membro_id=eq.${membroId}`);
      }
    }catch(e){
      setPresMap(p=>{const n={...p};if(anterior)n[k]=anterior;else delete n[k];return n;});
      alert("Erro ao salvar presença: "+e.message);
    }
  };

  const marcarTodosPresentes=async(ev)=>{
    const pendentes=membros.filter(m=>!presMap[`${ev.id}|${m.id}`]);
    if(!pendentes.length)return;
    if(!window.confirm(`Marcar PRESENÇA para os ${pendentes.length} membro(s) que estão em branco?`))return;
    const antes=presMap;
    setPresMap(p=>{const n={...p};pendentes.forEach(m=>{n[`${ev.id}|${m.id}`]="presenca";});return n;});
    try{
      await sbFetch("presencas?on_conflict=evento_id,membro_id",{
        method:"POST",
        headers:{Prefer:"resolution=merge-duplicates,return=representation"},
        body:JSON.stringify(pendentes.map(m=>({evento_id:ev.id,membro_id:m.id,status:"presenca"}))),
      });
    }catch(e){
      setPresMap(antes);
      alert("Erro ao salvar: "+e.message);
    }
  };

  // ── eventos externos ──
  const addExterno=async()=>{
    if(!formExterno.nome.trim())return;
    setSalvando(true);
    try{
      const [novo]=await dbInsert("externos",[{nome:formExterno.nome.trim(),descricao:formExterno.descricao.trim()||null,data:formExterno.data||null}]);
      setExternos(p=>ordenarEventos([...p,novo]));
      setExtPart(p=>({...p,[novo.id]:[]}));
      setFormExterno({nome:"",descricao:"",data:""});
    }catch(e){alert("Erro ao salvar evento externo: "+e.message);}
    setSalvando(false);
  };

  const removerExterno=async(ex)=>{
    if(!window.confirm(`Remover o evento externo "${ex.nome}" e a lista de participantes?`))return;
    try{
      await dbDelete("externos",`id=eq.${ex.id}`);
      setExternos(p=>p.filter(x=>x.id!==ex.id));
      setExtPart(p=>{const n={...p};delete n[ex.id];return n;});
      if(pickerId===ex.id) setPickerId(null);
    }catch(e){alert("Erro ao remover: "+e.message);}
  };

  const adicionarParticipantes=async(ex)=>{
    const jaTem=extPart[ex.id]||[];
    const novos=pickerSel.filter(id=>!jaTem.includes(id));
    setPickerId(null);setPickerSel([]);setPickerBusca("");
    if(!novos.length) return;
    const antes=extPart;
    setExtPart(p=>({...p,[ex.id]:[...(p[ex.id]||[]),...novos]}));
    try{
      await sbFetch("externos_membros?on_conflict=externo_id,membro_id",{
        method:"POST",
        headers:{Prefer:"resolution=merge-duplicates,return=representation"},
        body:JSON.stringify(novos.map(id=>({externo_id:ex.id,membro_id:id}))),
      });
    }catch(e){
      setExtPart(antes);
      alert("Erro ao adicionar participantes: "+e.message);
    }
  };

  const removerParticipante=async(ex,membroId)=>{
    const antes=extPart;
    setExtPart(p=>({...p,[ex.id]:(p[ex.id]||[]).filter(id=>id!==membroId)}));
    try{
      await dbDelete("externos_membros",`externo_id=eq.${ex.id}&membro_id=eq.${membroId}`);
    }catch(e){
      setExtPart(antes);
      alert("Erro ao remover participante: "+e.message);
    }
  };

  // ── presença e trocas nas escalas salvas (gravadas dentro do próprio histórico) ──
  const salvarResultadoHistorico=(id,resultado,anterior)=>{
    setHistorico(p=>p.map(x=>x.id===id?{...x,resultado}:x));
    filaRef.current=filaRef.current
      .then(()=>dbPatch("historico_escalas",`id=eq.${id}`,{resultado}))
      .catch(e=>{
        setHistorico(p=>p.map(x=>x.id===id?{...x,resultado:anterior}:x));
        alert("Erro ao salvar alteração: "+e.message);
      });
  };

  const definirStatusEscala=(h,idx,nome,novo)=>{
    const resultado=h.resultado.map((d,i)=>{
      if(i!==idx)return d;
      const status={...(d.status||{})};
      if(novo)status[nome]=novo;else delete status[nome];
      return {...d,status};
    });
    salvarResultadoHistorico(h.id,resultado,h.resultado);
  };

  const marcarTodosPresentesEscala=(h,idx)=>{
    const resultado=h.resultado.map((d,i)=>{
      if(i!==idx)return d;
      const status={...(d.status||{})};
      Object.keys(d.alocados||{}).forEach(n=>{if(!status[n])status[n]="presenca";});
      return {...d,status};
    });
    salvarResultadoHistorico(h.id,resultado,h.resultado);
  };

  const trocarEscalado=(h,idx,antigo,novo)=>{
    const resultado=h.resultado.map((d,i)=>{
      if(i!==idx)return d;
      const alocados={};
      Object.entries(d.alocados||{}).forEach(([n,fn])=>{alocados[n===antigo?novo:n]=fn;});
      const status={...(d.status||{})};
      delete status[antigo];
      const subs={...(d.substituicoes||{})};
      const original=subs[antigo]||antigo;
      delete subs[antigo];
      if(norm(original)!==norm(novo))subs[novo]=original;
      // registro das trocas (quem saiu e quem entrou); aproveita as trocas antigas se ainda não havia o registro
      const trocas=Array.isArray(d.trocas)?[...d.trocas]:Object.entries(d.substituicoes||{}).map(([para,de])=>({de,para}));
      const iDesfazer=trocas.map(t=>norm(t.de)===norm(novo)&&norm(t.para)===norm(antigo)).lastIndexOf(true);
      if(iDesfazer>=0)trocas.splice(iDesfazer,1);
      else trocas.push({de:antigo,para:novo});
      return {...d,alocados,status,substituicoes:subs,trocas};
    });
    salvarResultadoHistorico(h.id,resultado,h.resultado);
  };

  // ── sorteio ──
  const gerarSorteio=()=>{
    const r=sortearEscala(dias,membrosEf,disponibilidades,historico,manualSlots);
    setSorteioAtual(r);
    setAba("escalas");
    setSub("sorteio");
  };

  const salvarHistorico=async()=>{
    if(!sorteioAtual)return;
    setSalvando(true);
    try{
      // guarda também quem estava disponível em cada missa (para as estatísticas do cadastro)
      const resultado=sorteioAtual.resultado.map(d=>({...d,disponiveis:disponibilidades[d.id]||[]}));
      const [novo]=await dbInsert("historico_escalas",[{nome:nomeEscala||`Escala ${new Date().toLocaleDateString("pt-BR")}`,resultado,pendentes:(sorteioAtual.pendentes||[]).map(p=>p.nome||p)}]);
      setHistorico(p=>[{...novo,resultado:novo.resultado||[]},...p]);
      alert("Escala salva!");
    }catch(e){alert("Erro: "+e.message);}
    setSalvando(false);
  };

  const excluirHistorico=async(id)=>{
    if(!window.confirm("Excluir esta escala salva? As presenças e trocas registradas nela também serão perdidas."))return;
    try{
      await dbDelete("historico_escalas",`id=eq.${id}`);
      setHistorico(p=>p.filter(h=>h.id!==id));
    }catch(e){alert("Erro ao excluir: "+e.message);}
  };

  // ── mais derivados ──
  const linkMembro=`${window.location.origin}${window.location.pathname}?membro=1`;
  const statsPre=calcPre(dias,membros,disponibilidades);
  const statsPos=sorteioAtual?calcPos(sorteioAtual,membros,disponibilidades):null;

  const chaveParado=m=>{
    const du=(dispStats[m.id]||{}).ultima;
    if(!du) return 0;
    if(du.tipo==="atual"&&!du.data) return Date.now();
    return new Date(du.data).getTime()||0;
  };
  const membrosCadastro=membrosEf.filter(m=>norm(m.nome).includes(norm(buscaMembro))).sort((a,b)=>{
    if(ordemCadastro==="escore") return (b.score-a.score)||a.nome.localeCompare(b.nome);
    if(ordemCadastro==="parado") return (chaveParado(a)-chaveParado(b))||a.nome.localeCompare(b.nome);
    if(ordemCadastro==="trocas") return ((statsMembros[b.id]?.trocas||0)-(statsMembros[a.id]?.trocas||0))||a.nome.localeCompare(b.nome);
    if(ordemCadastro==="externos") return ((extCont[a.id]||0)-(extCont[b.id]||0))||a.nome.localeCompare(b.nome);
    return a.nome.localeCompare(b.nome);
  });
  const membrosEvento=[...membros].filter(m=>norm(m.nome).includes(norm(buscaEvento))).sort((a,b)=>a.nome.localeCompare(b.nome));

  if(simMembro) return <TelaMembro onVoltar={()=>setSimMembro(false)}/>;
  if(loading) return <div style={bgApp}><Spinner/></div>;

  const T=TEMAS[aba];
  const thBase={background:"#3d0000",color:C.ouro,padding:"8px 6px",fontSize:11,letterSpacing:.5,borderBottom:`1px solid ${C.ouro}50`,borderRight:`1px solid ${C.ouro}20`,textAlign:"center",verticalAlign:"top"};
  const P=PESOS_ESCORE;

  return(
    <TemaCtx.Provider value={aba}>
    <div style={{...bgApp,background:T.fundo}}>
      {/* HEADER */}
      <div style={{textAlign:"center",padding:"22px 24px 8px"}}>
        <img src={BRASAO} alt="Brasão" style={{width:96,height:96,objectFit:"contain",marginBottom:6,filter:"drop-shadow(0 2px 8px rgba(0,0,0,0.35))"}}/>
        <h1 style={{margin:"4px 0 0",fontSize:"clamp(30px,7vw,50px)",fontWeight:400,fontStyle:"italic",letterSpacing:10,paddingLeft:10,color:T.titulo,fontFamily:"Georgia,serif"}}>ALTARE</h1>
        <div style={{fontSize:11,letterSpacing:5,textTransform:"uppercase",color:T.subtitulo,marginTop:2}}>Software Administrativo</div>
        <div style={{width:110,height:1,background:`linear-gradient(90deg,transparent,${T.linha},transparent)`,margin:"12px auto 0"}}/>
      </div>

      {/* ABAS PRINCIPAIS */}
      <div style={{display:"flex",justifyContent:"center",alignItems:"flex-end",flexWrap:"wrap",gap:"2px 38px",padding:"8px 10px 0",borderBottom:`1px solid ${T.linha}55`}}>
        {ABAS.map(([k,t])=>{
          const ativo=aba===k;
          return(
            <button key={k} onClick={()=>setAba(k)} style={{
              background:"none",border:"none",borderBottom:`4px solid ${ativo?T.abaBorda:"transparent"}`,
              padding:"8px 6px 10px",cursor:"pointer",fontFamily:"Georgia,serif",textTransform:"uppercase",
              letterSpacing:ativo?3:2,fontWeight:ativo?700:400,
              fontSize:ativo?"clamp(18px,4.6vw,28px)":"clamp(12px,3vw,16px)",
              color:ativo?T.abaAtiva:T.abaInativa,
              textShadow:ativo&&k!=="eventos"?"0 0 14px rgba(232,201,109,0.45)":"none",
              transition:"all .2s",
            }}>{t}</button>
          );
        })}
      </div>

      {/* SUB-ABAS DE ESCALAS */}
      {aba==="escalas"&&(
        <div style={{display:"flex",justifyContent:"center",flexWrap:"wrap",gap:3,padding:"10px 10px 0"}}>
          {SUBS.map(([k,t])=>(
            <button key={k} onClick={()=>setSub(k)} style={{padding:"6px 12px",border:`1px solid ${sub===k?C.ouro:C.ouro+"40"}`,borderRadius:7,cursor:"pointer",fontFamily:"inherit",fontSize:11,background:sub===k?`${C.ouro}35`:"rgba(0,0,0,0.2)",color:sub===k?C.ouroClaro:C.ouroPale,fontWeight:sub===k?"700":"400"}}>{t}</button>
          ))}
        </div>
      )}

      <div style={{maxWidth:aba==="eventos"?1400:960,margin:"0 auto",padding:"16px 12px 48px"}}>
        <div style={aba==="eventos"?{background:"rgba(35,0,0,0.94)",border:`1px solid ${C.ouro}60`,borderRadius:16,padding:16,boxShadow:"0 6px 24px rgba(0,0,0,0.35)"}:undefined}>

        {/* ═════════ CADASTRO ═════════ */}
        {aba==="cadastro"&&(
          <div>
            {!colScoreOk&&(
              <Card style={{marginBottom:12,border:"1px solid #ff884460"}}>
                <SecTitle>⚠️ Escore sugerido ainda não ativado</SecTitle>
                <div style={{fontSize:12,color:"#ffaa66"}}>Rode o SQL do escore (coluna <strong>score_editado</strong>) no Supabase e recarregue a página para ativar o escore sugerido.</div>
              </Card>
            )}
            {!colObsOk&&(
              <Card style={{marginBottom:12,border:"1px solid #ff884460"}}>
                <SecTitle>⚠️ Observações ainda não ativadas</SecTitle>
                <div style={{fontSize:12,color:"#ffaa66"}}>Rode o SQL da coluna <strong>observacoes</strong> em membros no Supabase e recarregue a página para usar o campo de observações.</div>
              </Card>
            )}
            {!dispDataOk&&(
              <Card style={{marginBottom:12,border:"1px solid #ff884460"}}>
                <SecTitle>⚠️ Data do formulário não registrada</SecTitle>
                <div style={{fontSize:12,color:"#ffaa66"}}>Rode o SQL da coluna <strong>criado_em</strong> em disponibilidades para registrar a data em que cada pessoa preenche o formulário.</div>
              </Card>
            )}

            <Card style={{marginBottom:12}}>
              <SecTitle>➕ Adicionar membros</SecTitle>
              <div style={{display:"flex",flexWrap:"wrap",gap:8,alignItems:"center"}}>
                <button onClick={()=>setEditandoId("novo")} style={btnPri}>+ Novo membro</button>
                <label style={{...btnSec,display:"inline-flex",alignItems:"center",gap:6,cursor:"pointer"}}>
                  📎 Importar planilha (.xlsx)
                  <input type="file" accept=".xlsx,.xls" onChange={importarXlsx} style={{position:"absolute",opacity:0,width:0,height:0}}/>
                </label>
                <button onClick={baixarModelo} style={{...btnSec,fontSize:12,padding:"8px 14px"}}>⬇ Modelo</button>
                {xlsxStatus==="carregando"&&<span style={{color:C.ouro,fontSize:12}}>⏳ Importando…</span>}
                {xlsxStatus==="ok"&&<span style={{color:"#6fcf6f",fontSize:12}}>✅ {xlsxMsg}</span>}
                {xlsxStatus==="erro"&&<span style={{color:"#ff7044",fontSize:12}}>❌ {xlsxMsg}</span>}
              </div>
              <p style={{color:`${C.ouroPale}`,fontSize:11,margin:"10px 0 0",opacity:.8}}>
                A planilha precisa das colunas <strong>Nome</strong>, <strong>Score</strong> e uma coluna por função com Sim/Não. Quem já está cadastrado (mesmo nome) é atualizado; quem é novo é adicionado. Ninguém é apagado pela importação. Escores preenchidos na planilha passam a valer como escores manuais.
              </p>
            </Card>

            <Card style={{marginBottom:12,padding:12}}>
              <div style={{fontSize:11,color:C.ouroPale,lineHeight:1.6}}>
                <strong style={{color:C.ouro}}>Como o escore sugerido é calculado:</strong> presença em escala {P.escala.presenca>0?"+":""}{P.escala.presenca}, falta em escala {P.escala.falta}, presença em reunião +{P.reuniao.presenca}, falta em reunião {P.reuniao.falta}. Justificadas não pontuam. O escore digitado manualmente sempre prevalece.
              </div>
            </Card>

            {editandoId==="novo"&&(
              <div style={{marginBottom:12}}>
                <FormMembro onSalvar={d=>salvarMembro(d,null)} onCancelar={()=>setEditandoId(null)} salvando={salvando} obsAtiva={colObsOk}/>
              </div>
            )}

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:10}}>
              <div style={{fontSize:12,color:C.ouro,letterSpacing:2,textTransform:"uppercase"}}>{membros.length} membro(s) cadastrado(s)</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
                <select value={ordemCadastro} onChange={e=>setOrdemCadastro(e.target.value)} style={{...inp,width:"auto",fontSize:12}}>
                  <option value="nome">Ordenar: Nome (A–Z)</option>
                  <option value="escore">Ordenar: Maior escore</option>
                  <option value="parado">Ordenar: Mais tempo sem se disponibilizar</option>
                  <option value="trocas">Ordenar: Mais trocas solicitadas</option>
                  <option value="externos">Ordenar: Menos eventos externos</option>
                </select>
                <input placeholder="🔍 Buscar membro…" value={buscaMembro} onChange={e=>setBuscaMembro(e.target.value)} style={{...inp,maxWidth:230}}/>
                {colScoreOk&&membros.length>0&&<button onClick={usarSugeridoTodos} style={{...btnSec,fontSize:11,padding:"8px 12px"}} title="Remove os escores manuais de todos">↺ Usar escore sugerido em todos</button>}
              </div>
            </div>

            {!membros.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:36,fontSize:13}}>Nenhum membro cadastrado. Adicione um novo ou importe a planilha.</div>}

            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:10}}>
              {membrosCadastro.map(m=>{
                if(editandoId===m.id) return(
                  <div key={m.id} style={{gridColumn:"1 / -1"}}>
                    <FormMembro inicial={m} onSalvar={d=>salvarMembro(d,m.id)} onCancelar={()=>setEditandoId(null)} salvando={salvando} obsAtiva={colObsOk}/>
                  </div>
                );
                const s=statsMembros[m.id]||{pres:0,falta:0,just:0,ePres:0,eFalta:0,eJust:0,trocas:0,escalado:0,fnTop:null};
                const dd=dispStats[m.id]||{escalasDisp:0,total:0,ultima:null};
                const du=dd.ultima;
                let ultTxt="Nunca se disponibilizou";
                let ultCor="#ff8a8a";
                if(du){
                  const rotulo=du.tipo==="atual"?"Escala atual (em aberto)":du.nome;
                  const nd=du.data?diasDesde(du.data):null;
                  ultTxt=`${rotulo}${du.data?` · ${fmtISO(du.data)}`:""}${nd!==null?` (${nd===0?"hoje":`há ${nd} dia(s)`})`:""}`;
                  ultCor=nd===null||nd<=21?"#7bd88f":nd<=60?"#f0d060":"#ff8a8a";
                }
                const mostraSug=m.scoreManual&&Math.abs(m.scoreSugerido-m.score)>=0.05;
                return(
                  <Card key={m.id} style={{padding:12,border:`1px solid ${(m.funcoes||[]).length>0?C.ouro+"50":C.ouro+"25"}`}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,marginBottom:6}}>
                      <div style={{color:C.branco,fontSize:13,fontWeight:700,lineHeight:1.3}}>{m.nome}</div>
                      <div style={{display:"flex",flexWrap:"wrap",justifyContent:"flex-end",alignItems:"center",gap:4,flexShrink:0,maxWidth:200}}>
                        <span style={{fontSize:11,color:C.ouro,background:`${C.ouro}20`,borderRadius:4,padding:"2px 7px"}} title={m.scoreManual?"Escore definido manualmente":"Escore calculado automaticamente"}>★ {m.score}{m.scoreManual?" ✏️":" · auto"}</span>
                        {mostraSug&&(
                          <span style={{fontSize:10,color:"#88ddff",background:"rgba(136,221,255,0.12)",border:"1px solid rgba(136,221,255,0.35)",borderRadius:4,padding:"2px 6px",display:"inline-flex",alignItems:"center",gap:4}} title="Calculado a partir das presenças e faltas">
                            ESCORE SUGERIDO: {m.scoreSugerido}
                            {colScoreOk&&<button onClick={()=>usarEscoreSugerido(m)} style={{background:"none",border:`1px solid rgba(136,221,255,0.5)`,color:"#88ddff",borderRadius:3,fontSize:9,padding:"0 4px",cursor:"pointer",fontFamily:"inherit"}}>usar</button>}
                          </span>
                        )}
                      </div>
                    </div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:3,marginBottom:8,minHeight:18}}>
                      {(m.funcoes||[]).length===0
                        ?<span style={{color:`${C.ouro}60`,fontSize:10}}>Sem funções cadastradas</span>
                        :(m.funcoes||[]).map(f=><span key={f} style={{fontSize:10,background:`${C.ouro}20`,color:C.ouro,border:`1px solid ${C.ouro}40`,borderRadius:4,padding:"1px 5px"}}>{f}</span>)}
                    </div>
                    <div style={{fontSize:9,color:`${C.ouro}90`,letterSpacing:1,textTransform:"uppercase",marginBottom:3}}>Reuniões e eventos</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:4}}>
                      <Mini label="Presenças" valor={s.pres} cor="#7bd88f"/>
                      <Mini label="Faltas" valor={s.falta} cor="#ff8a8a"/>
                      <Mini label="Justificadas" valor={s.just} cor="#f0d060"/>
                    </div>
                    <div style={{fontSize:9,color:`${C.ouro}90`,letterSpacing:1,textTransform:"uppercase",margin:"7px 0 3px"}}>Escalas</div>
                    <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:4}}>
                      <Mini label="Presenças" valor={s.ePres} cor="#7bd88f"/>
                      <Mini label="Faltas" valor={s.eFalta} cor="#ff8a8a"/>
                      <Mini label="Justificadas" valor={s.eJust} cor="#f0d060"/>
                    </div>
                    <div style={{marginTop:7,fontSize:11,color:C.ouroPale,lineHeight:1.6}}>
                      🎲 Escalado <strong style={{color:C.ouroClaro}}>{s.escalado}</strong> vez(es)<br/>
                      ⭐ Função que mais desempenha: <strong style={{color:C.ouroClaro}}>{s.fnTop?`${s.fnTop.nome} (${s.fnTop.qtd}x)`:"—"}</strong><br/>
                      🔄 Trocas solicitadas: <strong style={{color:C.ouroClaro}}>{s.trocas}</strong><br/>
                      🌐 Eventos externos: <strong style={{color:C.ouroClaro}}>{extCont[m.id]||0}</strong><br/>
                      📝 Disponível em <strong style={{color:C.ouroClaro}}>{dd.escalasDisp}</strong> de <strong style={{color:C.ouroClaro}}>{dd.total}</strong> escala(s)<br/>
                      🕒 Último formulário: <strong style={{color:ultCor}}>{ultTxt}</strong>
                    </div>
                    {m.observacoes&&(
                      <div style={{marginTop:7,fontSize:11,color:C.ouroPale,background:"rgba(0,0,0,0.3)",borderLeft:`3px solid ${C.ouro}`,borderRadius:4,padding:"5px 8px",whiteSpace:"pre-wrap"}}>
                        <strong style={{color:C.ouro}}>📌 Observações:</strong> {m.observacoes}
                      </div>
                    )}
                    <div style={{display:"flex",gap:6,marginTop:9}}>
                      <button onClick={()=>setEditandoId(m.id)} style={{...btnSec,fontSize:11,padding:"5px 12px"}}>✏️ Editar</button>
                      <button onClick={()=>removerMembro(m)} style={{background:"rgba(150,0,0,0.6)",border:`1px solid ${C.ouro}40`,color:C.ouroPale,borderRadius:7,padding:"5px 12px",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>🗑 Remover</button>
                    </div>
                  </Card>
                );
              })}
            </div>
          </div>
        )}

        {/* ═════════ ESCALAS ═════════ */}
        {aba==="escalas"&&sub==="dias"&&(
          <div>
            <p style={{color:C.ouroPale,fontSize:13,marginBottom:12}}>Credenciais: <strong style={{color:C.ouroClaro}}>2</strong> semana · <strong style={{color:C.ouroClaro}}>3</strong> domingos.</p>
            <Card style={{marginBottom:14}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10,marginBottom:10}}>
                <div>
                  <label style={lbl}>Data</label>
                  <input type="date" value={formDia.data} onChange={e=>setFormDia(p=>({...p,data:e.target.value}))} style={inp}/>
                  {formDia.data&&<div style={{fontSize:10,color:isDomingo(formDia.data)?C.ouro:C.ouroPale,marginTop:3}}>{isDomingo(formDia.data)?"☀️ Domingo — 3 cred.":"📅 Semana — 2 cred."}</div>}
                </div>
                <div>
                  <label style={lbl}>Horário</label>
                  <div style={{display:"flex",flexWrap:"wrap",gap:3,marginBottom:4}}>
                    {HORARIOS_FIXOS.map(h=>(
                      <button key={h} onClick={()=>setFormDia(p=>({...p,horario:h,usarCustom:false}))} style={{padding:"3px 7px",border:`1px solid ${formDia.horario===h&&!formDia.usarCustom?C.ouro:C.ouro+"40"}`,borderRadius:5,cursor:"pointer",fontFamily:"inherit",fontSize:11,background:formDia.horario===h&&!formDia.usarCustom?`linear-gradient(135deg,${C.ouro},#A07828)`:"rgba(0,0,0,0.2)",color:formDia.horario===h&&!formDia.usarCustom?C.vermelho:C.ouroPale}}>{h}</button>
                    ))}
                    <button onClick={()=>setFormDia(p=>({...p,usarCustom:true,horario:""}))} style={{padding:"3px 7px",border:`1px solid ${formDia.usarCustom?C.ouro:C.ouro+"40"}`,borderRadius:5,cursor:"pointer",fontFamily:"inherit",fontSize:11,background:formDia.usarCustom?`linear-gradient(135deg,${C.ouro},#A07828)`:"rgba(0,0,0,0.2)",color:formDia.usarCustom?C.vermelho:C.ouroPale}}>Outro</button>
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
              <div style={{display:"flex",gap:16,marginBottom:10,flexWrap:"wrap"}}>
                {[["temIncenso","🔥 Incenso"],["temBispo","👑 Bispo"]].map(([k,t])=>(
                  <label key={k} style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",fontSize:12,color:C.ouroPale}}>
                    <input type="checkbox" checked={formDia[k]} onChange={e=>setFormDia(p=>({...p,[k]:e.target.checked}))} style={{accentColor:C.ouro,width:14,height:14}}/>{t}
                  </label>
                ))}
                <label style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer",fontSize:12,color:"#f0c060"}}>
                  <input type="checkbox" checked={formDia.solenidade} onChange={e=>setFormDia(p=>({...p,solenidade:e.target.checked}))} style={{accentColor:C.ouro,width:14,height:14}}/>✨ Solenidade
                </label>
              </div>
              {formDia.solenidade&&(
                <div style={{marginBottom:10}}>
                  <label style={lbl}>Nome da Solenidade</label>
                  <input placeholder="Ex: Corpus Christi…" value={formDia.nomeSolenidade} onChange={e=>setFormDia(p=>({...p,nomeSolenidade:e.target.value}))} style={inp}/>
                </div>
              )}
              <button onClick={addDia} disabled={salvando} style={{...btnPri,opacity:salvando?.7:1}}>+ Adicionar dia</button>
            </Card>
            {!dias.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:28,fontSize:13}}>Nenhum dia adicionado.</div>}
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {dias.map(dia=>(
                <div key={dia.id} style={{background:"rgba(0,0,0,0.2)",border:`1px solid ${C.ouro}30`,borderRadius:9,padding:"9px 12px",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:600,color:C.branco,fontSize:13,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                    <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>
                      {dia.horario} · {dia.igreja}
                      {isDomingo(dia.data)&&<span style={{marginLeft:4,fontSize:10,color:C.ouro,background:`${C.ouro}25`,borderRadius:3,padding:"1px 4px"}}>Dom</span>}
                      {dia.solenidade&&<span style={{marginLeft:4,fontSize:10,color:"#f0c060",background:"rgba(240,192,96,0.15)",borderRadius:3,padding:"1px 4px"}}>✨ {dia.nome_solenidade||"Solenidade"}</span>}
                      {dia.tem_incenso&&" 🔥"}{dia.tem_bispo&&" 👑"}
                    </div>
                  </div>
                  <div style={{fontSize:11,color:`${C.ouro}70`}}>{(disponibilidades[dia.id]||[]).length} disp.</div>
                  <button onClick={()=>removerDia(dia.id)} style={{background:"rgba(150,0,0,0.6)",border:`1px solid ${C.ouro}40`,color:C.ouroPale,borderRadius:5,padding:"3px 8px",cursor:"pointer",fontSize:11}}>✕</button>
                </div>
              ))}
            </div>
            {dias.length>0&&<div style={{textAlign:"right",marginTop:12}}><button onClick={()=>setSub("link")} style={btnPri}>Próximo →</button></div>}
          </div>
        )}

        {aba==="escalas"&&sub==="link"&&(
          <div>
            <Card style={{marginBottom:12,textAlign:"center"}}>
              <SecTitle>🔗 Link para os membros</SecTitle>
              <div style={{background:"rgba(0,0,0,0.3)",borderRadius:7,padding:"9px 12px",fontSize:12,color:C.ouroPale,fontFamily:"monospace",marginBottom:10,border:`1px solid ${C.ouro}30`,wordBreak:"break-all"}}>{linkMembro}</div>
              <div style={{display:"flex",gap:8,justifyContent:"center",flexWrap:"wrap"}}>
                <button onClick={()=>navigator.clipboard.writeText(linkMembro)} style={btnPri}>📋 Copiar link</button>
                <button onClick={()=>setSimMembro(true)} style={btnSec}>👁 Simular</button>
                <button onClick={()=>carregarDados(false)} style={{...btnSec,fontSize:12}}>🔄 Atualizar</button>
              </div>
              <p style={{color:`${C.ouro}60`,fontSize:11,marginTop:8}}>Atualiza automaticamente a cada 15s.</p>
            </Card>
            <Card>
              <SecTitle>📊 Respostas em tempo real</SecTitle>
              {!dias.length&&<div style={{color:`${C.ouro}50`,fontSize:12}}>Nenhum dia cadastrado.</div>}
              {dias.map(dia=>{const resp=disponibilidades[dia.id]||[];return(
                <div key={dia.id} style={{marginBottom:10,paddingBottom:10,borderBottom:`1px solid ${C.ouro}20`}}>
                  <div style={{color:C.branco,fontSize:12,fontWeight:600,textTransform:"capitalize"}}>{fmtData(dia.data)} · {dia.horario}</div>
                  <div style={{color:C.ouroPale,fontSize:11,marginBottom:4}}>{dia.igreja} — <strong style={{color:C.ouroClaro}}>{resp.length}</strong> resp.</div>
                  {!resp.length?<span style={{color:`${C.ouro}50`,fontSize:11}}>Sem respostas</span>
                    :<div style={{display:"flex",flexWrap:"wrap",gap:3}}>{resp.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")}/>)}</div>}
                </div>
              );})}
            </Card>
            <div style={{textAlign:"right",marginTop:12}}><button onClick={()=>setSub("estatisticas")} style={btnPri}>Ver Estatísticas →</button></div>
          </div>
        )}

        {aba==="escalas"&&sub==="estatisticas"&&(
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            <Card>
              <SecTitle>📊 Pré-sorteio</SecTitle>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:8,marginBottom:12}}>
                <div style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:10,textAlign:"center"}}>
                  <div style={{fontSize:20,color:C.ouroClaro,fontWeight:700}}>{statsPre.mediaDisp}</div>
                  <div style={{fontSize:10,color:C.ouroPale}}>média dias marcados</div>
                </div>
                <div style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:10,textAlign:"center"}}>
                  <div style={{fontSize:12,color:C.ouroClaro,fontWeight:700,textTransform:"capitalize"}}>{statsPre.maisDia?`${fmtData(statsPre.maisDia.data).split(",")[0]} ${statsPre.maisDia.horario}`:"—"}</div>
                  <div style={{fontSize:10,color:C.ouroPale}}>+ disponível ({statsPre.maisDia?.count||0})</div>
                </div>
                <div style={{background:"rgba(0,0,0,0.2)",borderRadius:8,padding:10,textAlign:"center"}}>
                  <div style={{fontSize:12,color:"#ff9966",fontWeight:700,textTransform:"capitalize"}}>{statsPre.menosDia?`${fmtData(statsPre.menosDia.data).split(",")[0]} ${statsPre.menosDia.horario}`:"—"}</div>
                  <div style={{fontSize:10,color:C.ouroPale}}>- disponível ({statsPre.menosDia?.count||0})</div>
                </div>
              </div>
              <div style={{fontSize:11,color:C.ouro,marginBottom:6}}>Disponibilidade por membro</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(175px,1fr))",gap:3,maxHeight:160,overflowY:"auto",marginBottom:10}}>
                {[...membros].sort((a,b)=>(statsPre.dispPorMembro[b.nome]||0)-(statsPre.dispPorMembro[a.nome]||0)).map(m=>(
                  <div key={m.id} style={{display:"flex",justifyContent:"space-between",background:"rgba(0,0,0,0.15)",borderRadius:4,padding:"3px 7px",fontSize:11}}>
                    <span style={{color:C.brancoOff}}>{m.nome.split(" ").slice(0,2).join(" ")}</span>
                    <span style={{color:C.ouro}}>{statsPre.dispPorMembro[m.nome]||0}d</span>
                  </div>
                ))}
              </div>
              {statsPre.naoPreencheram.length>0&&(
                <div>
                  <div style={{fontSize:11,color:"#ff9966",marginBottom:4}}>⚠️ Não preencheram ({statsPre.naoPreencheram.length})</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{statsPre.naoPreencheram.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")} color="#ff9966"/>)}</div>
                </div>
              )}
            </Card>
            {statsPos&&(
              <Card>
                <SecTitle>🎯 Pós-sorteio</SecTitle>
                <div style={{marginBottom:8,fontSize:12,color:C.ouroPale}}>Média: <strong style={{color:C.ouroClaro}}>{statsPos.media}</strong> escalação(ões)</div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(195px,1fr))",gap:3,maxHeight:200,overflowY:"auto",marginBottom:10}}>
                  {membros.filter(m=>statsPos.dispPor[m.nome]>0).sort((a,b)=>(statsPos.escaladoPor[b.nome]||0)-(statsPos.escaladoPor[a.nome]||0)).map(m=>{
                    const esc=statsPos.escaladoPor[m.nome]||0,disp=statsPos.dispPor[m.nome]||0,pct=disp>0?Math.round(esc/disp*100):0;
                    return(
                      <div key={m.id} style={{background:"rgba(0,0,0,0.15)",borderRadius:4,padding:"5px 8px",fontSize:11}}>
                        <div style={{display:"flex",justifyContent:"space-between",marginBottom:2}}>
                          <span style={{color:C.brancoOff}}>{m.nome.split(" ").slice(0,2).join(" ")}</span>
                          <span style={{color:esc>0?C.ouroClaro:`${C.ouro}50`}}>{esc}/{disp}</span>
                        </div>
                        <div style={{height:2,background:`${C.ouro}20`,borderRadius:2}}><div style={{width:`${pct}%`,height:"100%",background:`linear-gradient(90deg,${C.ouro},${C.ouroClaro})`,borderRadius:2}}/></div>
                      </div>
                    );
                  })}
                </div>
                {statsPos.pendentes?.length>0&&(
                  <div>
                    <div style={{fontSize:11,color:"#88ddff",marginBottom:4}}>🔖 Prioridade próxima escala ({statsPos.pendentes.length})</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:3}}>{(statsPos.pendentes||[]).map(p=>{const n=p.nome||p;return<Tag key={n} label={n.split(" ").slice(0,2).join(" ")} color="#88ddff"/>;})}</div>
                  </div>
                )}
              </Card>
            )}
            <div style={{textAlign:"right"}}><button onClick={gerarSorteio} style={btnPri}>🎲 Gerar Sorteio →</button></div>
          </div>
        )}

        {aba==="escalas"&&sub==="sorteio"&&(
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
              <SecTitle>🎲 Sorteio</SecTitle>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                <button onClick={gerarSorteio} style={btnPri}>🔀 {sorteioAtual?"Ressortear":"Gerar"}</button>
                {sorteioAtual&&<button onClick={salvarHistorico} disabled={salvando} style={{...btnSec,fontSize:12,opacity:salvando?.7:1}}>💾 Salvar</button>}
              </div>
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:12}}>
              <div>
                <label style={lbl}>Nome da escala</label>
                <input placeholder="Ex: Julho 2025…" value={nomeEscala} onChange={e=>setNomeEscala(e.target.value)} style={inp}/>
              </div>
              <div>
                <label style={lbl}>🔍 Buscar nome</label>
                <input placeholder="Digite um nome…" value={buscaSorteio} onChange={e=>setBuscaSorteio(e.target.value)} style={inp}/>
              </div>
            </div>
            {sorteioAtual?.observacoes?.length>0&&(
              <Card style={{marginBottom:10,border:`1px solid #ff884450`}}>
                <SecTitle>⚠️ Observações</SecTitle>
                {sorteioAtual.observacoes.map((o,i)=><div key={i} style={{color:"#ffaa66",fontSize:12,marginBottom:3}}>• {o.motivo}</div>)}
              </Card>
            )}
            {sorteioAtual?.pendentes?.length>0&&(
              <Card style={{marginBottom:10,border:`1px solid #88ddff40`}}>
                <SecTitle>🔖 Prioridade na próxima escala</SecTitle>
                <div style={{display:"flex",flexWrap:"wrap",gap:3}}>
                  {(sorteioAtual.pendentes||[]).map(p=>{const n=p.nome||p;return<Tag key={n} label={`${n.split(" ").slice(0,2).join(" ")}${p.score>0?" ★"+p.score:""}`} color="#88ddff"/>;})}</div>
              </Card>
            )}
            {!sorteioAtual&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:40,fontSize:13}}>Clique em "Gerar" para sortear.</div>}
            {sorteioAtual?.resultado?.map((dia,idx)=>{
              const falta=dia.funcs.some(({nome:fn,qtd})=>Object.entries(dia.alocados||{}).filter(([,f])=>f===fn).length<qtd);
              return(
                <Card key={idx} style={{marginBottom:10,border:`1px solid ${falta?"#ff884460":dia.solenidade?"#f0c06050":C.ouro+"30"}`}}>
                  <div style={{marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:6}}>
                    <div>
                      <div style={{fontSize:14,color:C.branco,fontWeight:600,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                      <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>
                        {dia.horario} · {dia.igreja}
                        {isDomingo(dia.data)&&<span style={{marginLeft:4,fontSize:10,color:C.ouro,background:`${C.ouro}25`,borderRadius:3,padding:"1px 4px"}}>Dom</span>}
                        {dia.solenidade&&<span style={{marginLeft:4,fontSize:10,color:"#f0c060",background:"rgba(240,192,96,0.15)",borderRadius:3,padding:"1px 4px"}}>✨ {dia.nome_solenidade||"Solenidade"}</span>}
                        {(dia.tem_incenso||dia.temIncenso)&&" 🔥"}{(dia.tem_bispo||dia.temBispo)&&" 👑"}
                        {falta&&<span style={{marginLeft:6,color:"#ffaa66",fontSize:10}}>⚠️ Vagas em aberto</span>}
                      </div>
                    </div>
                    <button onClick={()=>setEditandoManual(editandoManual===dia.id?null:dia.id)} style={{...btnSec,fontSize:11,padding:"4px 10px"}}>
                      {editandoManual===dia.id?"✕ Fechar":"✏️ Manual"}
                    </button>
                  </div>

                  {editandoManual===dia.id&&(
                    <div style={{background:"rgba(0,0,0,0.3)",borderRadius:8,padding:12,marginBottom:10,border:`1px solid ${C.ouro}30`}}>
                      <div style={{fontSize:11,color:C.ouro,marginBottom:8}}>Defina manualmente — o restante será preenchido automaticamente:</div>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))",gap:6}}>
                        {dia.funcs.map(({nome:fn})=>{
                          const atualNome=Object.entries(manualSlots[dia.id]||{}).find(([,f])=>f===fn)?.[0]||"";
                          return(
                            <div key={fn} style={{background:"rgba(0,0,0,0.2)",borderRadius:6,padding:"7px 9px"}}>
                              <div style={{fontSize:10,color:C.ouro,marginBottom:4,textTransform:"uppercase"}}>{fn}</div>
                              <select value={atualNome} onChange={e=>{
                                const v=e.target.value;
                                setManualSlots(prev=>{
                                  const d={...(prev[dia.id]||{})};
                                  Object.keys(d).forEach(k=>{if(d[k]===fn)delete d[k];});
                                  if(v) d[v]=fn;
                                  return{...prev,[dia.id]:d};
                                });
                              }} style={{...inp,fontSize:12,padding:"5px 8px"}}>
                                <option value="">— Automático —</option>
                                {(disponibilidades[dia.id]||[]).filter(n=>{const m=membros.find(x=>x.nome===n);return(m?.funcoes||[]).some(f=>norm(f)===norm(fn));}).map(n=><option key={n}>{n}</option>)}
                              </select>
                            </div>
                          );
                        })}
                      </div>
                      <button onClick={()=>{gerarSorteio();setEditandoManual(null);}} style={{...btnPri,marginTop:10,fontSize:12}}>🎲 Gerar com manual aplicado</button>
                    </div>
                  )}

                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(170px,1fr))",gap:6}}>
                    {dia.funcs.map(({nome:fn,qtd})=>{
                      const pessoas=Object.entries(dia.alocados||{}).filter(([,f])=>f===fn).map(([n])=>n);
                      const f=qtd-pessoas.length;
                      return(
                        <div key={fn} style={{background:f>0?"rgba(180,40,0,0.2)":`${C.ouro}08`,border:`1px solid ${f>0?"#ff440035":C.ouro+"35"}`,borderRadius:7,padding:"7px 9px"}}>
                          <div style={{fontSize:10,letterSpacing:1,color:C.ouro,marginBottom:3,textTransform:"uppercase"}}>{fn} ({qtd})</div>
                          {pessoas.map(p=>{
                            const mb=membrosEf.find(m=>m.nome===p);
                            const highlight=buscaSorteio&&p.toLowerCase().includes(buscaSorteio.toLowerCase());
                            return(
                              <div key={p} style={{fontSize:11,color:highlight?"#FFD700":C.ouroPale,marginBottom:1,display:"flex",justifyContent:"space-between",background:highlight?"rgba(255,215,0,0.1)":"transparent",borderRadius:3,padding:"1px 3px"}}>
                                <span>✓ {p.split(" ").slice(0,2).join(" ")}</span>
                                {mb&&mb.score!==0&&<span style={{color:C.ouro,fontSize:10}}>★{mb.score}</span>}
                              </div>
                            );
                          })}
                          {Array.from({length:f}).map((_,i)=><div key={i} style={{fontSize:11,color:"#ff8844",marginBottom:1}}>⚠ Sem candidato</div>)}
                        </div>
                      );
                    })}
                  </div>

                  {/* SUGESTÕES DE SUBSTITUTOS */}
                  {(()=>{
                    const escaladosNoDia=new Set(Object.keys(dia.alocados||{}));
                    const dispDia=disponibilidades[dia.id]||[];
                    const substitutos=membrosEf
                      .filter(m=>dispDia.includes(m.nome)&&!escaladosNoDia.has(m.nome))
                      .sort((a,b)=>{
                        const escA=sorteioAtual?.escAtual?.[a.nome]||0;
                        const escB=sorteioAtual?.escAtual?.[b.nome]||0;
                        if(escA!==escB) return escA-escB;
                        return (b.score||0)-(a.score||0);
                      })
                      .slice(0,5);
                    if(!substitutos.length) return null;
                    return(
                      <div style={{marginTop:10,background:"rgba(0,0,0,0.2)",borderRadius:8,padding:"8px 10px",border:`1px solid ${C.ouro}20`}}>
                        <div style={{fontSize:10,color:`${C.ouro}90`,letterSpacing:1,marginBottom:5,textTransform:"uppercase"}}>💡 Sugestões caso haja faltas</div>
                        <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                          {substitutos.map((m,i)=>(
                            <div key={m.nome} style={{display:"flex",alignItems:"center",gap:4,background:`${C.ouro}12`,border:`1px solid ${C.ouro}30`,borderRadius:5,padding:"2px 7px",fontSize:11,color:C.ouroPale}}>
                              <span style={{color:C.ouro,fontSize:10,fontWeight:700}}>{i+1}º</span>
                              {m.nome.split(" ").slice(0,2).join(" ")}
                              {m.score!==0&&<span style={{color:C.ouro,fontSize:10}}>★{m.score}</span>}
                            </div>
                          ))}
                        </div>
                      </div>
                    );
                  })()}
                </Card>
              );
            })}
          </div>
        )}

        {/* ═════════ HISTÓRICO ═════════ */}
        {aba==="escalas"&&sub==="historico"&&(
          <div>
            <SecTitle>📋 Histórico de escalas</SecTitle>
            {!historico.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:36,fontSize:13}}>Nenhuma escala salva.</div>}
            {historico.map(h=>{
              const r=resumoEscala(h);
              const aberto=abertoId===h.id;
              return(
                <Card key={h.id} style={{marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,flexWrap:"wrap",marginBottom:7}}>
                    <div>
                      <div style={{color:C.branco,fontWeight:600,fontSize:13}}>{h.nome}</div>
                      <div style={{color:C.ouroPale,fontSize:11,marginTop:1}}>{new Date(h.criado_em).toLocaleDateString("pt-BR",{day:"2-digit",month:"long",year:"numeric"})}</div>
                    </div>
                    <div style={{display:"flex",gap:6}}>
                      <button onClick={()=>{setAbertoId(aberto?null:h.id);setTrocandoKey(null);}} style={{...btnSec,fontSize:11,padding:"4px 12px"}}>{aberto?"▲ Fechar":"📂 Abrir"}</button>
                      <button onClick={()=>excluirHistorico(h.id)} style={{background:"rgba(150,0,0,0.6)",border:`1px solid ${C.ouro}40`,color:C.ouroPale,borderRadius:7,padding:"4px 10px",cursor:"pointer",fontSize:11,fontFamily:"inherit"}}>Excluir</button>
                    </div>
                  </div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:3,marginBottom:6}}>
                    {(h.resultado||[]).map((dia,i)=>(
                      <span key={i} style={{fontSize:10,background:`${C.ouro}15`,color:C.ouro,border:`1px solid ${C.ouro}30`,borderRadius:4,padding:"1px 5px",textTransform:"capitalize"}}>
                        {fmtData(dia.data).split(",")[0]} {dia.horario}
                      </span>
                    ))}
                  </div>
                  <div style={{display:"flex",gap:12,flexWrap:"wrap",fontSize:11,marginBottom:6}}>
                    <span style={{color:"#7bd88f"}}>✅ {r.pres} presença(s)</span>
                    <span style={{color:"#ff8a8a"}}>❌ {r.falta} falta(s)</span>
                    <span style={{color:"#f0d060"}}>📝 {r.just} justificada(s)</span>
                    <span style={{color:C.ouroPale,opacity:.8}}>⏳ {r.pend} sem registro</span>
                  </div>

                  {aberto&&(
                    <div style={{display:"flex",flexDirection:"column",gap:10,margin:"10px 0"}}>
                      {(h.resultado||[]).map((dia,idx)=>{
                        const escalados=ordenarEscalados(dia);
                        const nomesNoDia=new Set(escalados.map(([n])=>n));
                        const ehBispo=dia.tem_bispo||dia.temBispo;
                        const ehIncenso=dia.tem_incenso||dia.temIncenso;
                        return(
                          <div key={idx} style={{background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}35`,borderRadius:10,padding:12}}>
                            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,flexWrap:"wrap",marginBottom:6}}>
                              <div>
                                <div style={{color:C.branco,fontWeight:700,fontSize:13,textTransform:"capitalize"}}>{fmtData(dia.data)} · {dia.horario}</div>
                                <div style={{color:C.ouroPale,fontSize:11,marginTop:3,display:"flex",flexWrap:"wrap",gap:4,alignItems:"center"}}>
                                  <span>{dia.igreja}</span>
                                  {isDomingo(dia.data)&&<span style={badgeSt(C.ouro)}>☀️ Domingo</span>}
                                  {dia.solenidade&&<span style={badgeSt("#f0c060")}>✨ {dia.nome_solenidade||"Solenidade"}</span>}
                                  {ehBispo&&<span style={badgeSt("#e8a0ff")}>👑 Bispo</span>}
                                  {ehIncenso&&<span style={badgeSt("#ff9966")}>🔥 Incenso</span>}
                                </div>
                              </div>
                              <button onClick={()=>marcarTodosPresentesEscala(h,idx)} title="Marcar PRESENÇA para quem está sem registro" style={{background:STATUS.presenca.bg,color:STATUS.presenca.fg,border:"none",borderRadius:5,padding:"4px 9px",cursor:"pointer",fontSize:10,fontWeight:700,fontFamily:"inherit"}}>✓ todos presentes</button>
                            </div>

                            {escalados.map(([nome,fn])=>{
                              const st=(dia.status||{})[nome];
                              const orig=(dia.substituicoes||{})[nome];
                              const chave=`${h.id}|${idx}|${nome}`;
                              const temF=m=>(m.funcoes||[]).some(f=>norm(f)===norm(fn));
                              return(
                                <div key={nome} style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap",padding:"7px 0",borderTop:`1px solid ${C.ouro}18`}}>
                                  <div style={{flex:"1 1 180px",minWidth:160}}>
                                    <div style={{fontSize:9,color:C.ouro,letterSpacing:1,textTransform:"uppercase"}}>{fn}</div>
                                    <div style={{fontSize:12,color:C.branco,fontWeight:600}}>{nome}</div>
                                    {orig&&<div style={{fontSize:10,color:"#88ddff"}}>↔ no lugar de {orig.split(" ").slice(0,2).join(" ")}</div>}
                                  </div>
                                  <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap"}}>
                                    <BotoesStatus valor={st} onChange={v=>definirStatusEscala(h,idx,nome,v)}/>
                                    {trocandoKey===chave?(
                                      <div style={{display:"flex",gap:4,alignItems:"center"}}>
                                        <select autoFocus value="" onChange={e=>{if(e.target.value){trocarEscalado(h,idx,nome,e.target.value);setTrocandoKey(null);}}} style={{...inp,width:"auto",maxWidth:230,fontSize:11,padding:"4px 6px"}}>
                                          <option value="">Trocar por…</option>
                                          {membros.filter(m=>!nomesNoDia.has(m.nome)).sort((a,b)=>{
                                            const ta=temF(a)?0:1,tb=temF(b)?0:1;
                                            return ta!==tb?ta-tb:a.nome.localeCompare(b.nome);
                                          }).map(m=><option key={m.id} value={m.nome}>{m.nome}{temF(m)?"":" (sem esta função)"}</option>)}
                                        </select>
                                        <button onClick={()=>setTrocandoKey(null)} style={{background:"none",border:"none",color:C.ouroPale,cursor:"pointer",fontSize:14}}>✕</button>
                                      </div>
                                    ):(
                                      <button onClick={()=>setTrocandoKey(chave)} style={{background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}50`,color:C.ouroPale,borderRadius:5,padding:"3px 8px",fontSize:10,cursor:"pointer",fontFamily:"inherit"}}>↔ Trocar</button>
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })}
                    </div>
                  )}

                  {(h.pendentes||[]).length>0&&<div style={{fontSize:11,color:"#88ddff"}}>🔖 Pendentes: {h.pendentes.slice(0,5).map(p=>(p.nome||p).split(" ")[0]).join(", ")}{h.pendentes.length>5?` +${h.pendentes.length-5}`:""}</div>}
                </Card>
              );
            })}
          </div>
        )}

        {/* ═════════ EXTERNAS ═════════ */}
        {aba==="escalas"&&sub==="externas"&&(
          <div>
            {!externosOk&&(
              <Card style={{marginBottom:12,border:"1px solid #ff884460"}}>
                <SecTitle>⚠️ Tabelas de eventos externos não encontradas</SecTitle>
                <div style={{fontSize:12,color:"#ffaa66"}}>Rode o SQL das tabelas <strong>externos</strong> e <strong>externos_membros</strong> no Supabase e recarregue a página.</div>
              </Card>
            )}

            <Card style={{marginBottom:12}}>
              <SecTitle>➕ Novo evento externo</SecTitle>
              <p style={{color:C.ouroPale,fontSize:11,margin:"0 0 10px",opacity:.85}}>Registre eventos fora da paróquia e quem participou, sem limite de participantes. Serve para saber quem já teve a oportunidade e quem ainda não.</p>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:10,marginBottom:10}}>
                <div>
                  <label style={lbl}>Nome do evento externo</label>
                  <input value={formExterno.nome} onChange={e=>setFormExterno(p=>({...p,nome:e.target.value}))} placeholder="Ex: Encontro Diocesano" style={inp}/>
                </div>
                <div>
                  <label style={lbl}>Data (opcional)</label>
                  <input type="date" value={formExterno.data} onChange={e=>setFormExterno(p=>({...p,data:e.target.value}))} style={inp}/>
                </div>
                <div>
                  <label style={lbl}>Descrição (opcional)</label>
                  <input value={formExterno.descricao} onChange={e=>setFormExterno(p=>({...p,descricao:e.target.value}))} placeholder="Detalhes do evento" style={inp}/>
                </div>
              </div>
              <button onClick={addExterno} disabled={salvando||!formExterno.nome.trim()} style={{...btnPri,opacity:(salvando||!formExterno.nome.trim())?.6:1}}>+ Adicionar evento externo</button>
            </Card>

            {membros.length>0&&(()=>{
              const semExt=membros.filter(m=>!extCont[m.id]).sort((a,b)=>a.nome.localeCompare(b.nome));
              return(
                <Card style={{marginBottom:12,padding:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                    <div style={{fontSize:12,color:C.ouroPale}}>
                      🌐 <strong style={{color:C.ouroClaro}}>{membros.length-semExt.length}</strong> de {membros.length} membros já participaram de eventos externos · <strong style={{color:"#7bd88f"}}>{semExt.length}</strong> ainda sem oportunidade
                    </div>
                    {semExt.length>0&&<button onClick={()=>setMostrarSemExt(!mostrarSemExt)} style={{...btnSec,fontSize:11,padding:"4px 10px"}}>{mostrarSemExt?"ocultar":"ver quem"}</button>}
                  </div>
                  {mostrarSemExt&&(
                    <div style={{display:"flex",flexWrap:"wrap",gap:4,marginTop:10}}>
                      {semExt.map(m=><Tag key={m.id} label={m.nome.split(" ").slice(0,2).join(" ")} color="#7bd88f"/>)}
                    </div>
                  )}
                </Card>
              );
            })()}

            {!externos.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:36,fontSize:13}}>Nenhum evento externo cadastrado.</div>}

            {externos.map(ex=>{
              const ids=extPart[ex.id]||[];
              const participantes=ids.map(id=>membros.find(m=>m.id===id)).filter(Boolean).sort((a,b)=>a.nome.localeCompare(b.nome));
              const aberto=pickerId===ex.id;
              const candidatos=aberto
                ?membros.filter(m=>!ids.includes(m.id)&&norm(m.nome).includes(norm(pickerBusca))).sort((a,b)=>((extCont[a.id]||0)-(extCont[b.id]||0))||a.nome.localeCompare(b.nome))
                :[];
              return(
                <Card key={ex.id} style={{marginBottom:12}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:8,flexWrap:"wrap",marginBottom:8}}>
                    <div style={{flex:"1 1 220px"}}>
                      <div style={{color:C.branco,fontWeight:700,fontSize:15}}>{ex.nome}</div>
                      <div style={{color:C.ouroPale,fontSize:11,marginTop:2}}>
                        {ex.data?`📅 ${fmtDataCurta(ex.data)} · `:""}<strong style={{color:C.ouroClaro}}>{participantes.length}</strong> participante(s)
                      </div>
                      {ex.descricao&&<div style={{color:C.ouroPale,opacity:.85,fontSize:12,marginTop:4,whiteSpace:"pre-wrap"}}>{ex.descricao}</div>}
                    </div>
                    <div style={{display:"flex",gap:6}}>
                      <button onClick={()=>{setPickerId(aberto?null:ex.id);setPickerSel([]);setPickerBusca("");}} style={{...btnPri,fontSize:12,padding:"6px 12px"}}>{aberto?"✕ Fechar":"➕ Adicionar membros"}</button>
                      <button onClick={()=>removerExterno(ex)} title="Remover evento externo" style={{background:"rgba(150,0,0,0.6)",border:`1px solid ${C.ouro}40`,color:C.ouroPale,borderRadius:7,padding:"6px 10px",cursor:"pointer",fontSize:12,fontFamily:"inherit"}}>🗑</button>
                    </div>
                  </div>

                  {participantes.length===0
                    ?<div style={{fontSize:11,color:`${C.ouro}70`}}>Nenhum participante adicionado ainda.</div>
                    :<div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                      {participantes.map(m=><Tag key={m.id} label={m.nome.split(" ").slice(0,2).join(" ")} onRemove={()=>removerParticipante(ex,m.id)}/>)}
                    </div>}

                  {aberto&&(
                    <div style={{marginTop:10,background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}40`,borderRadius:10,padding:12}}>
                      <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center",marginBottom:8}}>
                        <input placeholder="🔍 Buscar membro…" value={pickerBusca} onChange={e=>setPickerBusca(e.target.value)} style={{...inp,maxWidth:240}}/>
                        <button onClick={()=>setPickerSel(candidatos.map(m=>m.id))} style={{...btnSec,fontSize:11,padding:"6px 10px"}}>Selecionar todos da lista</button>
                        <button onClick={()=>setPickerSel([])} style={{...btnSec,fontSize:11,padding:"6px 10px"}}>Limpar</button>
                      </div>
                      <div style={{fontSize:10,color:`${C.ouro}90`,marginBottom:6}}>Quem tem menos eventos externos aparece primeiro.</div>
                      <div style={{maxHeight:260,overflowY:"auto",display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(250px,1fr))",gap:4}}>
                        {candidatos.map(m=>{
                          const sel=pickerSel.includes(m.id);
                          const n=extCont[m.id]||0;
                          return(
                            <label key={m.id} style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:12,color:sel?C.ouroClaro:C.ouroPale,background:sel?`${C.ouro}20`:"rgba(0,0,0,0.2)",border:`1px solid ${sel?C.ouro:C.ouro+"25"}`,borderRadius:6,padding:"6px 8px"}}>
                              <input type="checkbox" checked={sel} onChange={()=>setPickerSel(p=>sel?p.filter(x=>x!==m.id):[...p,m.id])} style={{accentColor:C.ouro}}/>
                              <span style={{flex:1}}>{m.nome}</span>
                              <span style={{fontSize:10,color:n===0?"#7bd88f":C.ouro,whiteSpace:"nowrap"}}>{n===0?"nunca participou":`${n} externo(s)`}</span>
                            </label>
                          );
                        })}
                        {!candidatos.length&&<div style={{fontSize:12,color:`${C.ouro}70`,padding:8}}>Nenhum membro encontrado.</div>}
                      </div>
                      <div style={{display:"flex",gap:8,marginTop:10}}>
                        <button onClick={()=>adicionarParticipantes(ex)} disabled={!pickerSel.length} style={{...btnPri,opacity:pickerSel.length?1:.5}}>Adicionar ({pickerSel.length})</button>
                        <button onClick={()=>{setPickerId(null);setPickerSel([]);setPickerBusca("");}} style={btnSec}>Cancelar</button>
                      </div>
                    </div>
                  )}
                </Card>
              );
            })}
          </div>
        )}

        {/* ═════════ REUNIÕES E EVENTOS ═════════ */}
        {aba==="eventos"&&(
          <div>
            {!eventosOk&&(
              <Card style={{marginBottom:12,border:"1px solid #ff884460"}}>
                <SecTitle>⚠️ Tabelas de eventos não encontradas</SecTitle>
                <div style={{fontSize:12,color:"#ffaa66"}}>Rode o SQL das tabelas <strong>eventos</strong> e <strong>presencas</strong> no Supabase e recarregue a página.</div>
              </Card>
            )}

            <Card style={{marginBottom:12}}>
              <SecTitle>➕ Novo evento / reunião</SecTitle>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:10,marginBottom:10}}>
                <div>
                  <label style={lbl}>Nome do evento</label>
                  <input value={formEvento.nome} onChange={e=>setFormEvento(p=>({...p,nome:e.target.value}))} placeholder="Ex: Formação Prática" style={inp}/>
                </div>
                <div>
                  <label style={lbl}>Data (opcional)</label>
                  <input type="date" value={formEvento.data} onChange={e=>setFormEvento(p=>({...p,data:e.target.value}))} style={inp}/>
                </div>
                <div>
                  <label style={lbl}>Descrição (opcional)</label>
                  <input value={formEvento.descricao} onChange={e=>setFormEvento(p=>({...p,descricao:e.target.value}))} placeholder="Detalhes do evento" style={inp}/>
                </div>
              </div>
              <button onClick={addEvento} disabled={salvando||!formEvento.nome.trim()} style={{...btnPri,opacity:(salvando||!formEvento.nome.trim())?.6:1}}>+ Adicionar evento</button>
            </Card>

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:10}}>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",fontSize:11,color:C.ouroPale}}>
                Clique na célula para alternar:
                {Object.values(STATUS).map(s=><span key={s.label} style={{background:s.bg,color:s.fg,borderRadius:4,padding:"2px 8px",fontSize:10,fontWeight:700}}>{s.label}</span>)}
              </div>
              <input placeholder="🔍 Buscar membro…" value={buscaEvento} onChange={e=>setBuscaEvento(e.target.value)} style={{...inp,maxWidth:260}}/>
            </div>

            {!membros.length&&<div style={{textAlign:"center",color:`${C.ouro}50`,padding:36,fontSize:13}}>Cadastre os membros na aba CADASTRO para registrar presenças.</div>}

            {membros.length>0&&(
              <div style={{overflow:"auto",border:`1px solid ${C.ouro}40`,borderRadius:10,maxHeight:"70vh"}}>
                <table style={{borderCollapse:"separate",borderSpacing:0,fontSize:12,minWidth:"100%"}}>
                  <thead>
                    <tr>
                      <th style={{...thBase,position:"sticky",left:0,top:0,zIndex:4,minWidth:40,width:40}}>Nº</th>
                      <th style={{...thBase,position:"sticky",left:40,top:0,zIndex:4,minWidth:220,textAlign:"left"}}>NOMES</th>
                      <th style={{...thBase,position:"sticky",top:0,zIndex:2,minWidth:62,color:"#7bd88f"}}>PRES.</th>
                      <th style={{...thBase,position:"sticky",top:0,zIndex:2,minWidth:62,color:"#ff8a8a"}}>FALTA</th>
                      <th style={{...thBase,position:"sticky",top:0,zIndex:2,minWidth:62,color:"#f0d060"}}>JUST.</th>
                      {eventos.map(ev=>(
                        <th key={ev.id} style={{...thBase,position:"sticky",top:0,zIndex:2,minWidth:140,maxWidth:170}}>
                          <div style={{fontWeight:700,color:C.ouroClaro,fontSize:12,whiteSpace:"normal"}}>{ev.nome}</div>
                          {ev.data&&<div style={{color:C.ouroPale,fontSize:11,marginTop:2}}>{fmtDataCurta(ev.data)}</div>}
                          {ev.descricao&&<div title={ev.descricao} style={{color:`${C.ouroPale}`,opacity:.75,fontSize:10,marginTop:2,whiteSpace:"normal",fontWeight:400,letterSpacing:0}}>{ev.descricao}</div>}
                          <div style={{display:"flex",gap:4,justifyContent:"center",marginTop:6}}>
                            <button onClick={()=>marcarTodosPresentes(ev)} title="Marcar PRESENÇA para quem está em branco" style={{background:STATUS.presenca.bg,color:STATUS.presenca.fg,border:"none",borderRadius:4,padding:"2px 6px",cursor:"pointer",fontSize:10,fontWeight:700,fontFamily:"inherit"}}>✓ todos</button>
                            <button onClick={()=>removerEvento(ev)} title="Remover evento" style={{background:"rgba(150,0,0,0.7)",color:C.ouroPale,border:"none",borderRadius:4,padding:"2px 6px",cursor:"pointer",fontSize:10,fontFamily:"inherit"}}>✕</button>
                          </div>
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {membrosEvento.map((m,i)=>{
                      const s=statsMembros[m.id]||{pres:0,falta:0,just:0};
                      const fundo=i%2===0?"#5c0000":"#500000";
                      const cel={background:fundo,borderBottom:`1px solid ${C.ouro}18`,borderRight:`1px solid ${C.ouro}18`,padding:"5px 6px",color:C.brancoOff};
                      return(
                        <tr key={m.id}>
                          <td style={{...cel,position:"sticky",left:0,zIndex:1,textAlign:"center",color:`${C.ouroPale}`,opacity:.8}}>{i+1}</td>
                          <td style={{...cel,position:"sticky",left:40,zIndex:1,whiteSpace:"nowrap",fontWeight:600}}>{m.nome}</td>
                          <td style={{...cel,textAlign:"center",color:"#7bd88f",fontWeight:700}}>{s.pres}</td>
                          <td style={{...cel,textAlign:"center",color:"#ff8a8a",fontWeight:700}}>{s.falta}</td>
                          <td style={{...cel,textAlign:"center",color:"#f0d060",fontWeight:700}}>{s.just}</td>
                          {eventos.map(ev=>{
                            const st=presMap[`${ev.id}|${m.id}`];
                            const cfg=st?STATUS[st]:null;
                            return(
                              <td key={ev.id} style={{padding:0,borderBottom:`1px solid ${C.ouro}18`,borderRight:`1px solid ${C.ouro}18`,background:fundo}}>
                                <button onClick={()=>definirStatus(ev.id,m.id,proximoStatus(st))} style={{display:"block",width:"100%",height:30,border:"none",cursor:"pointer",background:cfg?cfg.bg:"transparent",color:cfg?cfg.fg:`${C.ouro}50`,fontSize:11,fontWeight:700,fontFamily:"inherit"}}>
                                  {cfg?cfg.label:"—"}
                                </button>
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
            {membros.length>0&&!eventos.length&&<div style={{textAlign:"center",color:`${C.ouro}60`,padding:20,fontSize:12}}>Nenhum evento ainda. Adicione o primeiro acima — ele aparece como uma nova coluna.</div>}
          </div>
        )}

        </div>
      </div>
    </div>
    </TemaCtx.Provider>
  );
}

export default function App(){
  const isMembro=window.location.search.includes("membro=1");
  return isMembro ? <TelaMembro/> : <Painel/>;
}

const lbl={display:"block",fontSize:10,color:C.ouro,letterSpacing:1,marginBottom:4,textTransform:"uppercase"};
const inp={background:"rgba(0,0,0,0.3)",border:`1px solid ${C.ouro}50`,borderRadius:7,padding:"8px 10px",color:C.brancoOff,fontSize:13,fontFamily:"Georgia,serif",outline:"none",width:"100%",boxSizing:"border-box"};
const btnPri={background:`linear-gradient(135deg,${C.ouro},#A07828)`,border:"none",color:C.vermelho,borderRadius:7,padding:"9px 18px",cursor:"pointer",fontSize:13,fontFamily:"Georgia,serif",letterSpacing:.5,fontWeight:700};
const btnSec={background:"rgba(0,0,0,0.25)",border:`1px solid ${C.ouro}50`,color:C.ouroPale,borderRadius:7,padding:"9px 18px",cursor:"pointer",fontSize:13,fontFamily:"Georgia,serif"};
const badgeSt=cor=>({fontSize:10,color:cor,background:`${cor}25`,borderRadius:3,padding:"1px 5px"});