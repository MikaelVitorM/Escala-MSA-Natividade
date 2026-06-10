import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ── dados ─────────────────────────────────────────────────────────────────────
const NOMES_INICIAIS = [
  "Ágata Carolina Senna da Silva","Alef Renan de Sousa Sena","Alexsandro Guimarães Silveira",
  "Alice Negrão de Albuquerque","Amanda Vitória Gomes Batista","Ana Beatriz Santana da Silva",
  "Ana Rosa Liboa Figueiredo","Anna Leticia Rodrigues de Souza","Antonni Gabriel dos Santos Albuquerque",
  "Arianny Souza Braz","Beatriz do Espírito Santo Almeida","Bento Miguel da Silva Leitão",
  "Bruna Jaqueline Farias Silva","Bruna Maria Dias de Carvalho","Bruno Borges",
  "Carlos Daniel Trindade Silva","Celso José Souza dos Santos Filho","Davi de Carvalho Carneiro Sousa",
  "David Sousa Xavier","Eduarda de Jesus Araújo","Elizabeth Cristinny Araújo da Costa",
  "Elton Kayo Nunes Lima","Emanuelle Oliveira","Emily lohanny dos Santos Sacramenta",
  "Enzo Estefânio Teixeira Lopes","Estela Luz da Silva","Fabrício Daniel Rodrigues Fonseca",
  "Felipe Daniel Rodrigues Fonseca","Fernando Sarmento da Cruz","Gabriel Henrique dos Santos Costa",
  "Gabrielle Anjos dos Santos","Giovana de Nazaré Santos Martins","Guilherme Vasconcelos do Remédio",
  "Heidi Christinny Menezes de Oliveira","Hillary Leticia do Nascimento da Cruz","Hyuri Gabriel Saraiva Lopes",
  "Isabela Pinto Rodrigues","Ivana Pereira dos Santos","James Marcelo Ribeiro",
  "José Antonio de Sousa Favacho","José Gabriel Barbosa de Paula","Júlia Luane da Cruz de Souza",
  "Juliana Corrêa Nunes","Juliana Ferreira de Senna","Klara Coelho Silva",
  "Larissa Espíndola","Lucas Kauã Silva","Luis Francisco Pacheco Ferreira",
  "Luiz Eduardo Silva da Silva","Luma de Almeida Cardoso","Manuela Cunha Lima",
  "Marcos Alexandre Gomes Barata","Maria Antônia da Silva Leitão","Maria Cecília Sousa Xavier",
  "Maria Clara Barbosa da Silva","Maria Eduarda Lopes dos Santos","Maria Eduarda Silva Machado dos Santos",
  "Maria Elisa Alves de Brito","Maria Luiza Santos Cardoso","Maria Sophia Itapovica da Cruz",
  "Maria Sophia Menezes de Oliveira","Paulo Augusto Trindade Silva","Rafaela Araújo Pina",
  "Renan Arruda","Renan Sousa da Silva","Richard Gabriel Guimarães das Mercês",
  "Sabrina Nascimento Silva","Samuel Saymon Ramalho Damarceno","Samya Fabile Ramalho Damasceno",
  "Sayuri Eduarda Santos lima","Sophia Mendonça Pereira","Sophia Rodrigues dos Santos",
  "Sophia Tocantins Viana","Thaila Cardoso Ribeiro","Thalles Rafael da Silva Coelho",
  "Vitor Devison Lustoso Ribeiro","Vitória Cristo Ribeiro","Vitória Jinkings Carneiro",
  "Weslley Raiol","Yasmim Fabiane da Silva Marques","Yuri Rafael Ribeiro de Almeida Silva","Luis Otávio"
];

const IGREJAS = [
  "Paróquia Matriz","Menino Jesus de Praga","Nossa Senhora das Graças","São Sebastião",
  "São Pedro e São Paulo","Auxiliadora","Núcleo Santa Mônica","Núcleo Caminhando com Maria"
];

const HORARIOS_FIXOS = ["7h","8h","9h","18h","19h","19h30"];

const FUNCOES_BASE = [
  "Cerimoniário Principal","Cerimoniário Regente",
  "Librifero","Microfone","Credencial","Ceroferário",
  "Turiferário","Naveteiro","Baculífero","Mitrífero"
];

const isDomingo = (data) => {
  if (!data) return false;
  return new Date(data + "T12:00:00").getDay() === 0;
};

function getFuncoesNecessarias(dia) {
  const domingo = isDomingo(dia.data);
  const funcs = [
    { nome: "Cerimoniário Principal", qtd: 1 },
    { nome: "Cerimoniário Regente", qtd: 1 },
    { nome: "Librifero", qtd: 1 },
    { nome: "Microfone", qtd: 1 },
    { nome: "Credencial", qtd: domingo ? 3 : 2 },
    { nome: "Ceroferário", qtd: domingo ? 6 : 4 },
  ];
  if (dia.temIncenso) { funcs.push({ nome: "Turiferário", qtd: 1 }); funcs.push({ nome: "Naveteiro", qtd: 1 }); }
  if (dia.temBispo)   { funcs.push({ nome: "Baculífero", qtd: 1 }); funcs.push({ nome: "Mitrífero", qtd: 1 }); }
  return funcs;
}

function sortearEscala(dias, membros, disponibilidades) {
  const contagem = {};
  membros.forEach(m => { contagem[m.nome] = 0; });
  return dias.map(dia => {
    const funcs = getFuncoesNecessarias(dia);
    const dispDia = disponibilidades[dia.id] || [];
    const disponiveis = membros.filter(m => dispDia.includes(m.nome));
    const alocados = {};
    const usados = new Set();
    for (const { nome: fn, qtd } of funcs) {
      let cands = disponiveis.filter(m => m.funcoes.includes(fn) && !usados.has(m.nome));
      cands.sort((a,b) => contagem[a.nome] - contagem[b.nome]);
      for (let i = 0; i < qtd; i++) {
        if (!cands.length) break;
        const minC = contagem[cands[0].nome];
        const pool = cands.filter(c => contagem[c.nome] <= minC + 1);
        const pick = pool[Math.floor(Math.random() * pool.length)];
        alocados[pick.nome] = fn;
        usados.add(pick.nome);
        contagem[pick.nome]++;
        cands = cands.filter(c => c.nome !== pick.nome);
      }
    }
    return { ...dia, funcs, alocados };
  });
}

// ── helpers ───────────────────────────────────────────────────────────────────
function fmtData(d) {
  if (!d) return "";
  return new Date(d+"T12:00:00").toLocaleDateString("pt-BR",{weekday:"long",day:"2-digit",month:"long"});
}

function Tag({ label, onRemove }) {
  return (
    <span style={{display:"inline-flex",alignItems:"center",gap:4,background:"#1a5fa820",color:"#7ab3d9",border:"1px solid #1a5fa840",borderRadius:6,padding:"2px 8px",fontSize:12}}>
      {label}
      {onRemove && <button onClick={onRemove} style={{background:"none",border:"none",cursor:"pointer",color:"#7ab3d9",fontSize:14,lineHeight:1,padding:0}}>×</button>}
    </span>
  );
}

function MultiSelect({ options, selected, onChange, placeholder }) {
  const [open, setOpen] = useState(false);
  const toggle = v => onChange(selected.includes(v) ? selected.filter(x=>x!==v) : [...selected,v]);
  return (
    <div style={{position:"relative"}}>
      <div onClick={()=>setOpen(!open)} style={{border:"1px solid #3a5a78",borderRadius:8,padding:"8px 10px",cursor:"pointer",background:"#0d1b2a",minHeight:38,display:"flex",flexWrap:"wrap",gap:4,alignItems:"center"}}>
        {!selected.length && <span style={{color:"#556",fontSize:13}}>{placeholder}</span>}
        {selected.map(s=><Tag key={s} label={s} onRemove={e=>{e.stopPropagation();toggle(s);}}/>)}
        <span style={{marginLeft:"auto",fontSize:11,color:"#556"}}>{open?"▲":"▼"}</span>
      </div>
      {open && (
        <div style={{position:"absolute",zIndex:200,top:"100%",left:0,right:0,background:"#0d1b2a",border:"1px solid #3a5a78",borderRadius:8,boxShadow:"0 4px 20px #0006",maxHeight:200,overflowY:"auto"}}>
          {options.map(opt=>(
            <label key={opt} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",cursor:"pointer",background:selected.includes(opt)?"#1a5fa820":"transparent",fontSize:13,borderBottom:"1px solid #ffffff08",color:"#c8d6e5"}}>
              <input type="checkbox" checked={selected.includes(opt)} onChange={()=>toggle(opt)} style={{accentColor:"#1a5fa8"}}/>
              {opt}
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

// ── TELA DO MEMBRO ────────────────────────────────────────────────────────────
function TelaMembro({ dias, membros, onVoltar, disponibilidades, setDisponibilidades }) {
  const [nome, setNome] = useState("");
  const [busca, setBusca] = useState("");
  const [confirmado, setConfirmado] = useState(false);
  const nomesOrdenados = [...membros.map(m=>m.nome)].sort((a,b)=>a.localeCompare(b));
  const filtrados = nomesOrdenados.filter(n=>n.toLowerCase().includes(busca.toLowerCase()));
  const toggleDisp = diaId => {
    setDisponibilidades(prev => {
      const atual = prev[diaId]||[];
      return {...prev,[diaId]: atual.includes(nome)?atual.filter(n=>n!==nome):[...atual,nome]};
    });
  };
  const dispMembro = diaId => (disponibilidades[diaId]||[]).includes(nome);
  const total = dias.filter(d=>dispMembro(d.id)).length;

  if (!nome) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0d1b2a,#1a3a5c)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{textAlign:"center",marginBottom:28}}>
        <div style={{fontSize:32}}>✦</div>
        <h1 style={{color:"#fff",fontWeight:400,fontSize:22,margin:"8px 0",letterSpacing:2}}>Escala Litúrgica</h1>
        <p style={{color:"#7ab3d9",fontSize:13,margin:0}}>Selecione seu nome para marcar disponibilidade</p>
      </div>
      <div style={{width:"100%",maxWidth:420,background:"#ffffff0d",border:"1px solid #ffffff18",borderRadius:16,padding:20}}>
        <input placeholder="🔍 Buscar seu nome…" value={busca} onChange={e=>setBusca(e.target.value)} style={{...inp,marginBottom:10,width:"100%",boxSizing:"border-box"}}/>
        <div style={{maxHeight:340,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
          {filtrados.map(n=>(
            <button key={n} onClick={()=>setNome(n)} style={{textAlign:"left",padding:"10px 14px",border:"1px solid #ffffff18",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:13,background:"#ffffff08",color:"#c8d6e5"}}>{n}</button>
          ))}
        </div>
      </div>
      <button onClick={onVoltar} style={{...btnSec,marginTop:16,fontSize:13}}>← Voltar ao organizador</button>
    </div>
  );

  if (confirmado) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0d1b2a,#1a3a5c)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:24,textAlign:"center"}}>
      <div style={{fontSize:64,marginBottom:12}}>✅</div>
      <h2 style={{color:"#fff",fontWeight:400,fontSize:20,margin:0}}>Disponibilidade confirmada!</h2>
      <p style={{color:"#7ab3d9",marginTop:8}}>Obrigado, <strong style={{color:"#fff"}}>{nome.split(" ")[0]}</strong>! Você marcou <strong style={{color:"#7ab3d9"}}>{total}</strong> dia(s).</p>
      <button onClick={()=>{setNome("");setBusca("");setConfirmado(false);}} style={{...btnPri,marginTop:20}}>Voltar ao início</button>
    </div>
  );

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0d1b2a,#1a3a5c)",paddingBottom:48}}>
      <div style={{background:"#ffffff0d",borderBottom:"1px solid #ffffff18",padding:"16px 24px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:11,color:"#7ab3d9",letterSpacing:2,textTransform:"uppercase"}}>Olá,</div>
          <div style={{color:"#fff",fontSize:17,fontWeight:600}}>{nome}</div>
        </div>
        <button onClick={()=>{setNome("");setBusca("");}} style={btnSec}>Trocar nome</button>
      </div>
      <div style={{maxWidth:640,margin:"0 auto",padding:"20px 16px"}}>
        <p style={{color:"#7ab3d9",fontSize:13,marginBottom:16}}>Marque os dias em que você está disponível:</p>
        {dias.length===0 && <div style={{textAlign:"center",color:"#556",padding:40}}>Nenhum dia cadastrado.</div>}
        {dias.map(dia=>{
          const marcado = dispMembro(dia.id);
          const dom = isDomingo(dia.data);
          return (
            <div key={dia.id} onClick={()=>toggleDisp(dia.id)} style={{background:marcado?"#1a5fa830":"#ffffff0a",border:`2px solid ${marcado?"#1a5fa8":"#ffffff18"}`,borderRadius:12,padding:"14px 18px",marginBottom:10,cursor:"pointer",transition:"all .2s",display:"flex",alignItems:"center",gap:14}}>
              <div style={{width:26,height:26,borderRadius:"50%",border:`2px solid ${marcado?"#1a5fa8":"#3a5a78"}`,background:marcado?"#1a5fa8":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:14,color:"#fff",transition:"all .2s"}}>{marcado?"✓":""}</div>
              <div>
                <div style={{color:"#fff",fontWeight:600,fontSize:14,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                <div style={{color:"#7ab3d9",fontSize:12,marginTop:2}}>
                  {dia.horario} · {dia.igreja}
                  {dom&&<span style={{marginLeft:6,fontSize:11,color:"#a0c0e0",background:"#1a5fa840",borderRadius:4,padding:"1px 5px"}}>Domingo</span>}
                  {dia.temIncenso&&<span style={{marginLeft:6}}>🔥</span>}
                  {dia.temBispo&&<span style={{marginLeft:4}}>👑</span>}
                </div>
              </div>
            </div>
          );
        })}
        {dias.length>0 && (
          <button onClick={()=>setConfirmado(true)} style={{...btnPri,width:"100%",marginTop:8,padding:14,fontSize:15}}>
            ✓ Confirmar ({total} dia{total!==1?"s":""})
          </button>
        )}
      </div>
    </div>
  );
}

// ── APP PRINCIPAL ─────────────────────────────────────────────────────────────
const VIEWS = ["📅 Dias","👥 Membros","🔗 Link / Preview","🎲 Sorteio"];

export default function App() {
  const [view, setView] = useState(0);
  const [simulandoMembro, setSimulandoMembro] = useState(false);

  const [dias, setDias] = useState([]);
  const [formDia, setFormDia] = useState({ data:"", horario:"", horarioCustom:"", usarCustom:false, igreja:"", temIncenso:false, temBispo:false });

  const [membros, setMembros] = useState(NOMES_INICIAIS.map(nome=>({ nome, funcoes:[] })));
  const [membroAtivo, setMembroAtivo] = useState(null);
  const [busca, setBusca] = useState("");
  const [xlsxStatus, setXlsxStatus] = useState(null); // null | "ok" | "erro" | "carregando"
  const [xlsxMsg, setXlsxMsg] = useState("");
  const fileRef = useRef();

  const [disponibilidades, setDisponibilidades] = useState({});
  const [resultado, setResultado] = useState(null);

  // ── adicionar dia ──
  const addDia = () => {
    const horario = formDia.usarCustom ? formDia.horarioCustom : formDia.horario;
    if (!formDia.data || !horario || !formDia.igreja) return;
    setDias(prev=>[...prev,{ ...formDia, horario, id: Date.now().toString() }]);
    setFormDia({ data:"", horario:"", horarioCustom:"", usarCustom:false, igreja:"", temIncenso:false, temBispo:false });
  };

  // ── importar xlsx ──
  const importarXlsx = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setXlsxStatus("carregando");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type:"array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });

      if (!rows.length) { setXlsxStatus("erro"); setXlsxMsg("Planilha vazia."); return; }

      // detectar coluna de nome
      const keys = Object.keys(rows[0]);
      const nomeKey = keys.find(k => k.toLowerCase().includes("nome")) || keys[0];

      // colunas de funções = todas exceto a de nome
      const funcCols = keys.filter(k=>k!==nomeKey);

      let atualizados = 0;
      const novosMembros = [...membros];

      rows.forEach(row => {
        const nomeXlsx = String(row[nomeKey]||"").trim();
        if (!nomeXlsx) return;
        const funcoes = funcCols.filter(col => {
          const val = String(row[col]||"").trim().toLowerCase();
          return val === "sim" || val === "s" || val === "yes" || val === "x" || val === "1";
        });
        // mapear coluna para nome de função (case-insensitive)
        const funcoesMatch = funcoes.map(col => {
          const match = FUNCOES_BASE.find(f=>f.toLowerCase()===col.toLowerCase());
          return match || col;
        }).filter(f=>FUNCOES_BASE.includes(f));

        const idx = novosMembros.findIndex(m=>m.nome.toLowerCase()===nomeXlsx.toLowerCase());
        if (idx>=0) {
          novosMembros[idx] = { ...novosMembros[idx], funcoes: funcoesMatch };
          atualizados++;
        } else {
          // membro novo não cadastrado ainda — adiciona
          novosMembros.push({ nome: nomeXlsx, funcoes: funcoesMatch });
          atualizados++;
        }
      });

      setMembros(novosMembros);
      setXlsxStatus("ok");
      setXlsxMsg(`${atualizados} membro(s) atualizado(s) com sucesso.`);
    } catch(err) {
      setXlsxStatus("erro");
      setXlsxMsg("Erro ao ler o arquivo: " + err.message);
    }
    e.target.value = "";
  };

  const gerarEscala = () => setResultado(sortearEscala(dias, membros, disponibilidades));

  const membroCurrent = membroAtivo ? membros.find(m=>m.nome===membroAtivo) : null;
  const membrosFiltrados = membros.filter(m=>m.nome.toLowerCase().includes(busca.toLowerCase()));

  if (simulandoMembro) return (
    <TelaMembro dias={dias} membros={membros} disponibilidades={disponibilidades} setDisponibilidades={setDisponibilidades} onVoltar={()=>setSimulandoMembro(false)}/>
  );

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0d1b2a 0%,#1a3a5c 50%,#0d1b2a 100%)",fontFamily:"Georgia,serif",color:"#e8f0f8"}}>
      <div style={{textAlign:"center",padding:"24px 24px 14px",borderBottom:"1px solid #ffffff18"}}>
        <div style={{fontSize:11,letterSpacing:6,color:"#7ab3d9",textTransform:"uppercase",marginBottom:5}}>Painel do Organizador</div>
        <h1 style={{margin:0,fontSize:"clamp(18px,4vw,30px)",fontWeight:400,color:"#fff",letterSpacing:2}}>✦ Escala Litúrgica ✦</h1>
      </div>

      <div style={{display:"flex",justifyContent:"center",flexWrap:"wrap",gap:4,padding:"14px 16px 0"}}>
        {VIEWS.map((v,i)=>(
          <button key={v} onClick={()=>setView(i)} style={{padding:"8px 16px",border:"1px solid",borderRadius:8,cursor:"pointer",fontFamily:"inherit",fontSize:12,background:view===i?"#1a5fa8":"#ffffff08",borderColor:view===i?"#1a5fa8":"#ffffff20",color:view===i?"#fff":"#7ab3d9",letterSpacing:.5}}>{v}</button>
        ))}
      </div>

      <div style={{maxWidth:900,margin:"0 auto",padding:"20px 16px 48px"}}>

        {/* ── DIAS ── */}
        {view===0 && (
          <div>
            <p style={{color:"#7ab3d9",fontSize:13,marginBottom:18}}>
              Configure os dias da escala. Credenciais: <strong style={{color:"#fff"}}>2</strong> em dias de semana, <strong style={{color:"#fff"}}>3</strong> aos domingos.
            </p>
            <div style={{background:"#ffffff0d",borderRadius:12,padding:20,marginBottom:20,border:"1px solid #ffffff18"}}>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12,marginBottom:12}}>
                <div>
                  <label style={lbl}>Data</label>
                  <input type="date" value={formDia.data} onChange={e=>setFormDia(p=>({...p,data:e.target.value}))} style={inp}/>
                  {formDia.data && (
                    <div style={{fontSize:11,color:isDomingo(formDia.data)?"#7ab3d9":"#5a7a90",marginTop:4}}>
                      {isDomingo(formDia.data)?"☀️ Domingo — 3 credenciais":"📅 Dia de semana — 2 credenciais"}
                    </div>
                  )}
                </div>
                <div>
                  <label style={lbl}>Horário</label>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:6}}>
                    {HORARIOS_FIXOS.map(h=>(
                      <button key={h} onClick={()=>setFormDia(p=>({...p,horario:h,usarCustom:false}))} style={{
                        padding:"5px 10px",border:"1px solid",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,
                        background:formDia.horario===h&&!formDia.usarCustom?"#1a5fa8":"#ffffff10",
                        borderColor:formDia.horario===h&&!formDia.usarCustom?"#1a5fa8":"#3a5a78",
                        color:formDia.horario===h&&!formDia.usarCustom?"#fff":"#c8d6e5"
                      }}>{h}</button>
                    ))}
                    <button onClick={()=>setFormDia(p=>({...p,usarCustom:true,horario:""}))} style={{
                      padding:"5px 10px",border:"1px solid",borderRadius:6,cursor:"pointer",fontFamily:"inherit",fontSize:12,
                      background:formDia.usarCustom?"#1a5fa8":"#ffffff10",
                      borderColor:formDia.usarCustom?"#1a5fa8":"#3a5a78",
                      color:formDia.usarCustom?"#fff":"#c8d6e5"
                    }}>Outro</button>
                  </div>
                  {formDia.usarCustom && (
                    <input type="time" value={formDia.horarioCustom} onChange={e=>setFormDia(p=>({...p,horarioCustom:e.target.value}))} style={inp}/>
                  )}
                </div>
                <div>
                  <label style={lbl}>Igreja / Comunidade</label>
                  <select value={formDia.igreja} onChange={e=>setFormDia(p=>({...p,igreja:e.target.value}))} style={inp}>
                    <option value="">Selecione…</option>
                    {IGREJAS.map(ig=><option key={ig}>{ig}</option>)}
                  </select>
                </div>
              </div>
              <div style={{display:"flex",gap:24,marginBottom:14}}>
                <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:"#c8d6e5"}}>
                  <input type="checkbox" checked={formDia.temIncenso} onChange={e=>setFormDia(p=>({...p,temIncenso:e.target.checked}))} style={{accentColor:"#1a5fa8",width:16,height:16}}/>
                  🔥 Com incenso
                </label>
                <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:"#c8d6e5"}}>
                  <input type="checkbox" checked={formDia.temBispo} onChange={e=>setFormDia(p=>({...p,temBispo:e.target.checked}))} style={{accentColor:"#1a5fa8",width:16,height:16}}/>
                  👑 Com bispo
                </label>
              </div>
              <button onClick={addDia} style={btnPri}>+ Adicionar dia</button>
            </div>

            {!dias.length && <div style={{textAlign:"center",color:"#446",padding:36,fontSize:14}}>Nenhum dia adicionado ainda.</div>}
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {dias.map(dia=>(
                <div key={dia.id} style={{background:"#ffffff0d",border:"1px solid #ffffff18",borderRadius:10,padding:"12px 16px",display:"flex",alignItems:"center",gap:12}}>
                  <div style={{flex:1}}>
                    <div style={{fontWeight:600,color:"#fff",fontSize:14,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                    <div style={{color:"#7ab3d9",fontSize:12,marginTop:2}}>
                      {dia.horario} · {dia.igreja}
                      {isDomingo(dia.data)&&<span style={{marginLeft:6,fontSize:11,color:"#a0c0e0",background:"#1a5fa840",borderRadius:4,padding:"1px 5px"}}>Dom · 3 cred.</span>}
                      {dia.temIncenso&&<span style={{marginLeft:6}}>🔥</span>}
                      {dia.temBispo&&<span style={{marginLeft:4}}>👑</span>}
                    </div>
                  </div>
                  <div style={{fontSize:12,color:"#5a7a90"}}>{(disponibilidades[dia.id]||[]).length} disp.</div>
                  <button onClick={()=>setDias(p=>p.filter(d=>d.id!==dia.id))} style={{background:"#ff4444",border:"none",color:"#fff",borderRadius:6,padding:"4px 10px",cursor:"pointer",fontSize:12}}>✕</button>
                </div>
              ))}
            </div>
            {dias.length>0 && <div style={{textAlign:"right",marginTop:16}}><button onClick={()=>setView(1)} style={btnPri}>Próximo: Membros →</button></div>}
          </div>
        )}

        {/* ── MEMBROS ── */}
        {view===1 && (
          <div>
            {/* importar xlsx */}
            <div style={{background:"#1a5fa815",border:"1px solid #1a5fa840",borderRadius:12,padding:18,marginBottom:20}}>
              <div style={{fontSize:13,color:"#7ab3d9",letterSpacing:1,marginBottom:10,textTransform:"uppercase"}}>📂 Importar funções via planilha (.xlsx)</div>
              <p style={{color:"#a0b8cc",fontSize:12,margin:"0 0 12px"}}>
                A planilha deve ter uma coluna <strong style={{color:"#fff"}}>Nome</strong> e uma coluna para cada função (ex: "Librifero", "Ceroferário"…), preenchidas com <strong style={{color:"#fff"}}>Sim</strong> ou <strong style={{color:"#fff"}}>Não</strong>.
              </p>
              <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                <button onClick={()=>fileRef.current?.click()} style={btnPri}>📎 Selecionar arquivo .xlsx</button>
                <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={importarXlsx} style={{display:"none"}}/>
                {xlsxStatus==="carregando" && <span style={{color:"#7ab3d9",fontSize:13}}>⏳ Processando…</span>}
                {xlsxStatus==="ok" && <span style={{color:"#4caf50",fontSize:13}}>✅ {xlsxMsg}</span>}
                {xlsxStatus==="erro" && <span style={{color:"#ff7044",fontSize:13}}>❌ {xlsxMsg}</span>}
              </div>

              {/* modelo de download */}
              <button onClick={()=>{
                const header = ["Nome", ...FUNCOES_BASE];
                const ex = [["Maria da Silva","Sim","Não","Sim","Sim","Não","Não","Não","Não","Não","Não"]];
                const ws = XLSX.utils.aoa_to_sheet([header,...ex]);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb,"Membros",ws);
                XLSX.writeFile(wb,"modelo-escala.xlsx");
              }} style={{...btnSec,marginTop:10,fontSize:12,padding:"6px 14px"}}>
                ⬇ Baixar modelo de planilha
              </button>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"240px 1fr",gap:14}}>
              <div>
                <div style={{fontSize:11,color:"#7ab3d9",letterSpacing:2,textTransform:"uppercase",marginBottom:8}}>Membros ({membros.length})</div>
                <input placeholder="Buscar…" value={busca} onChange={e=>setBusca(e.target.value)} style={{...inp,marginBottom:8,width:"100%",boxSizing:"border-box"}}/>
                <div style={{maxHeight:460,overflowY:"auto",display:"flex",flexDirection:"column",gap:3}}>
                  {membrosFiltrados.map(m=>(
                    <button key={m.nome} onClick={()=>setMembroAtivo(m.nome)} style={{textAlign:"left",padding:"7px 10px",border:"1px solid",borderRadius:7,cursor:"pointer",fontFamily:"inherit",fontSize:12,background:membroAtivo===m.nome?"#1a5fa8":"#ffffff08",borderColor:membroAtivo===m.nome?"#1a5fa8":m.funcoes.length>0?"#2d6a2d44":"#ffffff15",color:membroAtivo===m.nome?"#fff":"#c8d6e5"}}>
                      <span style={{marginRight:4,opacity:.6}}>{m.funcoes.length>0?"✅":"○"}</span>
                      {m.nome.split(" ").slice(0,2).join(" ")}
                    </button>
                  ))}
                </div>
              </div>
              <div>
                {!membroCurrent
                  ? <div style={{textAlign:"center",color:"#446",padding:60,fontSize:13}}>← Selecione um membro para editar funções manualmente</div>
                  : (
                    <div style={{background:"#ffffff0d",border:"1px solid #ffffff18",borderRadius:12,padding:18}}>
                      <h3 style={{margin:"0 0 4px",color:"#fff",fontSize:15,fontWeight:400}}>{membroCurrent.nome}</h3>
                      <p style={{margin:"0 0 14px",color:"#7ab3d9",fontSize:12}}>Funções que pode desempenhar:</p>
                      <MultiSelect options={FUNCOES_BASE} selected={membroCurrent.funcoes}
                        onChange={v=>setMembros(prev=>prev.map(m=>m.nome===membroCurrent.nome?{...m,funcoes:v}:m))}
                        placeholder="Selecione as funções…"/>
                      <div style={{marginTop:10,fontSize:12,color:"#5a7a90"}}>{membroCurrent.funcoes.length} função(ões) selecionada(s)</div>
                    </div>
                  )
                }
                <div style={{marginTop:16,textAlign:"right"}}>
                  <button onClick={()=>setView(2)} style={btnPri}>Próximo: Link / Preview →</button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ── LINK / PREVIEW ── */}
        {view===2 && (
          <div>
            <p style={{color:"#7ab3d9",fontSize:13,marginBottom:20}}>Neste protótipo o link é simulado localmente. Na versão final (Supabase), um link real seria gerado para enviar no WhatsApp.</p>
            <div style={{background:"#1a5fa820",border:"1px solid #1a5fa860",borderRadius:12,padding:22,marginBottom:20,textAlign:"center"}}>
              <div style={{fontSize:12,color:"#7ab3d9",marginBottom:8,letterSpacing:1,textTransform:"uppercase"}}>Link que seria enviado aos membros</div>
              <div style={{background:"#0d1b2a",borderRadius:8,padding:"10px 16px",fontSize:13,color:"#a0c8f0",fontFamily:"monospace",marginBottom:14,wordBreak:"break-all"}}>
                https://escala-liturgica.app/responder?cod=ABC123
              </div>
              <button onClick={()=>setSimulandoMembro(true)} style={{...btnPri,fontSize:14,padding:"11px 26px"}}>
                👁 Simular tela do membro
              </button>
              <p style={{color:"#5a7a90",fontSize:12,marginTop:8}}>Veja exatamente como o membro verá ao abrir o link</p>
            </div>

            <div style={{background:"#ffffff0d",border:"1px solid #ffffff18",borderRadius:12,padding:18}}>
              <div style={{fontSize:12,color:"#7ab3d9",letterSpacing:1,marginBottom:14,textTransform:"uppercase"}}>📊 Respostas recebidas</div>
              {!dias.length && <div style={{color:"#446",fontSize:13}}>Nenhum dia cadastrado.</div>}
              {dias.map(dia=>{
                const resp = disponibilidades[dia.id]||[];
                return (
                  <div key={dia.id} style={{marginBottom:14,paddingBottom:14,borderBottom:"1px solid #ffffff10"}}>
                    <div style={{color:"#fff",fontSize:13,fontWeight:600,textTransform:"capitalize"}}>{fmtData(dia.data)} · {dia.horario}</div>
                    <div style={{color:"#7ab3d9",fontSize:12,marginBottom:6}}>{dia.igreja} — <strong style={{color:"#fff"}}>{resp.length}</strong> resposta(s)</div>
                    {!resp.length
                      ? <span style={{color:"#446",fontSize:12}}>Nenhuma resposta ainda</span>
                      : <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{resp.map(n=><Tag key={n} label={n.split(" ").slice(0,2).join(" ")}/>)}</div>
                    }
                  </div>
                );
              })}
            </div>
            <div style={{textAlign:"right",marginTop:16}}><button onClick={()=>setView(3)} style={btnPri}>Próximo: Sorteio →</button></div>
          </div>
        )}

        {/* ── SORTEIO ── */}
        {view===3 && (
          <div>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18,flexWrap:"wrap",gap:10}}>
              <h2 style={{color:"#7ab3d9",fontWeight:400,letterSpacing:2,fontSize:15,margin:0}}>🎲 SORTEIO</h2>
              <button onClick={gerarEscala} style={btnPri}>🔀 {resultado?"Ressortear":"Gerar Escala"}</button>
            </div>
            {!resultado && <div style={{textAlign:"center",color:"#446",padding:50,fontSize:14}}>Clique em "Gerar Escala" para sortear com base nas disponibilidades.</div>}
            {resultado && resultado.map((dia,idx)=>{
              const falta = dia.funcs.some(({nome:fn,qtd})=>Object.entries(dia.alocados).filter(([,f])=>f===fn).length<qtd);
              return (
                <div key={idx} style={{background:"#ffffff0a",border:`1px solid ${falta?"#ff884460":"#ffffff18"}`,borderRadius:12,padding:18,marginBottom:14}}>
                  <div style={{marginBottom:12}}>
                    <div style={{fontSize:15,color:"#fff",fontWeight:600,textTransform:"capitalize"}}>{fmtData(dia.data)}</div>
                    <div style={{color:"#7ab3d9",fontSize:12,marginTop:2}}>
                      {dia.horario} · {dia.igreja}
                      {isDomingo(dia.data)&&<span style={{marginLeft:6,fontSize:11,color:"#a0c0e0",background:"#1a5fa840",borderRadius:4,padding:"1px 5px"}}>Dom · 3 cred.</span>}
                      {dia.temIncenso&&<span style={{marginLeft:6}}>🔥</span>}
                      {dia.temBispo&&<span style={{marginLeft:4}}>👑</span>}
                      {falta&&<span style={{marginLeft:10,fontSize:11,color:"#ffaa66"}}>⚠️ Vagas em aberto</span>}
                    </div>
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))",gap:8}}>
                    {dia.funcs.map(({nome:fn,qtd})=>{
                      const pessoas = Object.entries(dia.alocados).filter(([,f])=>f===fn).map(([n])=>n);
                      const f = qtd-pessoas.length;
                      return (
                        <div key={fn} style={{background:f>0?"#ff440012":"#00aa4412",border:`1px solid ${f>0?"#ff440030":"#00aa4430"}`,borderRadius:8,padding:"9px 12px"}}>
                          <div style={{fontSize:10,letterSpacing:1,color:"#7ab3d9",marginBottom:5,textTransform:"uppercase"}}>{fn} ({qtd})</div>
                          {pessoas.map(p=><div key={p} style={{fontSize:12,color:"#d0e8d0",marginBottom:2}}>✓ {p.split(" ").slice(0,2).join(" ")}</div>)}
                          {Array.from({length:f}).map((_,i)=><div key={i} style={{fontSize:12,color:"#ff8844",marginBottom:2}}>⚠ Sem candidato</div>)}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

const lbl = { display:"block",fontSize:11,color:"#7ab3d9",letterSpacing:1,marginBottom:6,textTransform:"uppercase" };
const inp = { background:"#0d1b2a",border:"1px solid #3a5a78",borderRadius:8,padding:"9px 12px",color:"#e8f0f8",fontSize:14,fontFamily:"Georgia,serif",outline:"none",width:"100%",boxSizing:"border-box" };
const btnPri = { background:"linear-gradient(135deg,#1a5fa8,#0d3d6e)",border:"1px solid #2a7ad8",color:"#fff",borderRadius:8,padding:"10px 20px",cursor:"pointer",fontSize:14,fontFamily:"Georgia,serif",letterSpacing:1 };
const btnSec = { background:"#ffffff10",border:"1px solid #ffffff30",color:"#c8d6e5",borderRadius:8,padding:"10px 20px",cursor:"pointer",fontSize:14,fontFamily:"Georgia,serif" };
