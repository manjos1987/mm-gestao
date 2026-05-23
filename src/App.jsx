import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } from "recharts";
import { initializeApp } from "firebase/app";
import { getFirestore, doc, setDoc, onSnapshot } from "firebase/firestore";
import html2canvas from "html2canvas";

// ─────────────────────────────────────────────────────────────────────────────
// 🔴 PASSO 1: Cole aqui os dados do seu projeto Firebase
//    (você vai obter esses valores no console.firebase.google.com)
// ─────────────────────────────────────────────────────────────────────────────
const firebaseConfig = {
  apiKey:            "AIzaSyB6EVH7jY2qTMaD5wBUeawyI7eAPAf_PeI",
  authDomain:        "mm-gestao1987.firebaseapp.com",
  projectId:         "mm-gestao1987",
  storageBucket:     "mm-gestao1987.firebasestorage.app",
  messagingSenderId: "660580949425",
  appId:             "1:660580949425:web:556bc217a483733092b190",
};

// ─────────────────────────────────────────────────────────────────────────────
// 🔴 PASSO 2: Cole aqui sua chave da API Anthropic (console.anthropic.com)
//    Deixe vazio ("") se não quiser usar a função de gerar mensagem com IA
// ─────────────────────────────────────────────────────────────────────────────
const ANTHROPIC_KEY = "";

// ── Constantes ────────────────────────────────────────────────────────────────
const TODAY   = new Date().toISOString().split("T")[0];
const DAYS_PT = ["domingo","segunda-feira","terça-feira","quarta-feira","quinta-feira","sexta-feira","sábado"];
const RED     = "#C8102E";
const COLORS  = ["#C8102E","#2563EB","#059669","#D97706","#7C3AED","#0891B2","#B45309","#0D9488","#4F46E5","#065F46","#1D4ED8","#6D28D9","#BE185D","#92400E","#0E7490","#9D174D"];
const ROLES   = ["Sócio","Coordenador","Arquiteto","Estagiário","Parceiro"];
const SERVICES= ["Projeto arquitetônico","Projeto legal","Projeto executivo","Imagens / Renderização","Maquete","Interiores","Urbanismo","Estudo de Viabilidade"];
const DW      = {Sócio:3,Coordenador:2.5,Arquiteto:2,Estagiário:1,Parceiro:1.5};
const ST_CYC  = {ativo:"aguardando",aguardando:"bloqueado",bloqueado:"ativo"};
const STATUS  = {
  ativo:     {label:"Ativo",      bg:"#dcfce7",color:"#15803d",hex:"#16a34a"},
  aguardando:{label:"Aguardando", bg:"#fef9c3",color:"#a16207",hex:"#d97706"},
  bloqueado: {label:"Bloqueado",  bg:"#fee2e2",color:"#dc2626",hex:"#dc2626"},
  concluido: {label:"Concluído",  bg:"#f3f4f6",color:"#6b7280",hex:"#9ca3af"},
};
const ROLE_COLOR={Sócio:"#111",Coordenador:"#7C3AED",Arquiteto:"#2563EB",Estagiário:"#059669",Parceiro:"#D97706"};

const UFS=["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"];
// ── Dados iniciais ────────────────────────────────────────────────────────────
const INIT_C=[
  {id:"c1",name:"Canaã",contact:""},{id:"c2",name:"Santorini",contact:""},
  {id:"c3",name:"Ilha de Creta",contact:""},{id:"c4",name:"Cristalina",contact:""},
  {id:"c5",name:"Júlio Cezar",contact:""},{id:"c6",name:"Executivo",contact:""},
  {id:"c7",name:"Astro",contact:""},{id:"c8",name:"Luciano",contact:""},
  {id:"c9",name:"Terra do Meio",contact:""},{id:"c10",name:"Capitólio",contact:""},
  {id:"c11",name:"Hado",contact:""},
];
const INIT_P=[
  {id:"p1", cId:"c1", name:"Condomínio de Canaã",  svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p2", cId:"c2", name:"Santorini",             svcs:[], city:"",st:"",country:"Brasil",status:"aguardando",note:"Aguardando cliente"},
  {id:"p3", cId:"c3", name:"Ilha de Creta",         svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p4", cId:"c4", name:"Casas Cristalina",      svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p5", cId:"c1", name:"Casas Canaã",           svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p6", cId:"c5", name:"Kitnet Júlio Cezar",    svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p7", cId:"c6", name:"Executivo Towers",      svcs:[], city:"",st:"",country:"Brasil",status:"bloqueado", note:"Aguardando Haroldo / Guilherme"},
  {id:"p8", cId:"c6", name:"Executivo Mall",        svcs:[], city:"",st:"",country:"Brasil",status:"bloqueado", note:"Calçada errada — sem atualização"},
  {id:"p9", cId:"c7", name:"Apto Astro",            svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Verificar docs do Heli"},
  {id:"p10",cId:"c8", name:"Casas Luciano",         svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Seguir com projeto legal"},
  {id:"p11",cId:"c9", name:"Terra do Meio",         svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Pórtico pendente"},
  {id:"p12",cId:"c10",name:"Capitólio",             svcs:[], city:"",st:"",country:"Brasil",status:"bloqueado", note:"Aguardando Manoel"},
  {id:"p13",cId:"c11",name:"Hado San",              svcs:[], city:"",st:"",country:"Brasil",status:"aguardando",note:"Protocolado na prefeitura"},
  {id:"p14",cId:"c1", name:"Posto Canaã",           svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Atualizar planilha"},
];
const INIT_T=[
  {id:"m1",name:"Alice",   role:"Arquiteto"},{id:"m2",name:"Eloisa",  role:"Arquiteto"},
  {id:"m3",name:"Gustavo", role:"Arquiteto"},{id:"m4",name:"Letícia", role:"Estagiário"},
  {id:"m5",name:"Marcello",role:"Parceiro"}, {id:"m6",name:"Murilo",  role:"Arquiteto"},
];

// ── Helpers ───────────────────────────────────────────────────────────────────
const fmtDate=d=>{const[y,m,day]=d.split("-");return`${day}/${m}/${y}`;};
const fmtDay =d=>DAYS_PT[new Date(d+"T12:00:00").getDay()];
const cap    =s=>s.charAt(0).toUpperCase()+s.slice(1);
const uid    =()=>Date.now().toString(36)+Math.random().toString(36).slice(2,6);
const wOf    =(role,w)=>+(w[role]??DW[role]??1);

function normAlloc(a){
  const r={};
  Object.entries(a||{}).forEach(([mid,v])=>{r[mid]=Array.isArray(v)?v.filter(Boolean):typeof v==="string"&&v?[v]:[];});
  return r;
}

function calcReport(alloc,team,projects,weights={}){
const n=normAlloc(alloc),bp={};
team.forEach(m=>{
const pids=n[m.id]||[];if(!pids.length)return;
const w=wOf(m.role,weights);
// Peso FIXO: cada projeto recebe o peso inteiro da pessoa
pids.forEach(pid=>{
if(!bp[pid])bp[pid]={wpts:0,people:[],detail:[]};
bp[pid].wpts+=w;bp[pid].people.push(m.name);
bp[pid].detail.push({name:m.name,role:m.role,w:+w.toFixed(2)});
});
});
const tot=Object.values(bp).reduce((s,v)=>s+v.wpts,0);
return Object.entries(bp).map(([pid,{wpts,people,detail}])=>{
const p=projects.find(x=>x.id===pid);
return{pid,name:p?.name||pid,people,detail,wpts:+wpts.toFixed(2),pct:tot?wpts/tot*100:0};
}).sort((a,b)=>b.wpts-a.wpts);
}

function calcTeamLoad(alloc,team,weights={}){
  const n=normAlloc(alloc);
  const totW=team.reduce((s,m)=>{const pids=n[m.id]||[];return pids.length?s+wOf(m.role,weights):s;},0);
  return team.map(m=>{
    const pids=n[m.id]||[],w=wOf(m.role,weights);
    return{...m,weight:w,pids,pct:totW&&pids.length?w/totW*100:0,active:pids.length>0};
  });
}

// ── Export Excel / PDF ────────────────────────────────────────────────────────
function exportExcel(report,date,team,normA,projects,cont,weights){
  const wb=XLSX.utils.book_new();
  const s1=[["M+M Arquitetura — Relatório de Dedicação"],["Data:",fmtDate(date),cap(fmtDay(date))],[],
    ["Projeto","Contratante","Colaboradores","Pts Pond.","%"],
    ...report.map(r=>{const p=projects.find(x=>x.id===r.pid);const c=cont.find(x=>x.id===p?.cId);return[r.name,c?.name||"",r.people.join(", "),r.wpts,+r.pct.toFixed(1)];}),
    [],[`Pesos: ${Object.entries(weights).map(([r,w])=>`${r}=${w}`).join(", ")}`]
  ];
  const ws1=XLSX.utils.aoa_to_sheet(s1);ws1["!cols"]=[{wch:30},{wch:20},{wch:40},{wch:12},{wch:8}];
  XLSX.utils.book_append_sheet(wb,ws1,"Dedicação");
  const s2=[["Alocação por Colaborador"],["Data:",fmtDate(date)],[],["Colaborador","Cargo","Peso","Projetos"],
    ...team.map(m=>{const pids=normA[m.id]||[];return[m.name,m.role,wOf(m.role,weights),pids.map(pid=>projects.find(p=>p.id===pid)?.name||pid).join("; ")];})
  ];
  const ws2=XLSX.utils.aoa_to_sheet(s2);ws2["!cols"]=[{wch:20},{wch:15},{wch:8},{wch:70}];
  XLSX.utils.book_append_sheet(wb,ws2,"Colaboradores");
  XLSX.writeFile(wb,`MM_Relatorio_${date}.xlsx`);
}

function exportPDF(report,date,projects,cont,normA,team,weights){
  const cidx={};projects.forEach((p,i)=>{cidx[p.id]=i;});
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>M+M ${fmtDate(date)}</title>
<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:'Helvetica Neue',Arial,sans-serif;color:#111;-webkit-print-color-adjust:exact;print-color-adjust:exact}.page{padding:44px 48px;max-width:720px;margin:0 auto}.stripe{height:4px;background:#C8102E}.hdr{display:flex;justify-content:space-between;align-items:flex-end;padding:24px 0 16px;border-bottom:2px solid #111;margin-bottom:22px}.logo{font-size:28px;font-weight:800;letter-spacing:-1px}.plus{color:#C8102E}.sub{font-size:11px;color:#aaa;letter-spacing:.07em;text-transform:uppercase;margin-top:3px}.meta{text-align:right;font-size:12px;color:#888;line-height:1.7}.mdate{font-size:15px;font-weight:700;color:#111;display:block}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:22px}.stat{padding:12px;border:1px solid #f0f0f0;border-radius:6px;text-align:center}.sv{font-size:22px;font-weight:700}.sl{font-size:9px;text-transform:uppercase;letter-spacing:.06em;color:#bbb;margin-top:3px}.sec{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:#bbb;font-weight:700;margin-bottom:12px}.proj{margin-bottom:14px}.prow{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px}.pname{font-size:13px;font-weight:600}.pcli{font-size:10px;color:#aaa;margin-left:6px}.ppct{font-size:12px;color:#888;font-weight:600}.bar{height:7px;background:#f0f0f0;border-radius:4px;overflow:hidden;margin-bottom:3px}.fill{height:100%;border-radius:4px}.ppl{font-size:10px;color:#bbb}.div{border:none;border-top:1px solid #f0f0f0;margin:18px 0}.arow{display:flex;gap:10px;padding:5px 0;border-bottom:1px solid #f8f8f8;font-size:11px}.an{font-weight:600;min-width:80px;color:#555}.ar{color:#aaa;min-width:80px}.ap{color:#777}.foot{margin-top:22px;text-align:center;font-size:9px;color:#ddd;padding-top:10px;border-top:1px solid #f0f0f0}@media print{@page{margin:1cm}}</style>
</head><body><div class="stripe"></div><div class="page">
<div class="hdr"><div><div class="logo">M<span class="plus">+</span>M</div><div class="sub">Arquitetura · Relatório de Dedicação</div></div><div class="meta"><span class="mdate">${fmtDate(date)}</span>${cap(fmtDay(date))}</div></div>
<div class="stats"><div class="stat"><div class="sv">${report.length}</div><div class="sl">Projetos</div></div><div class="stat"><div class="sv">${team.filter(m=>(normA[m.id]||[]).length>0).length}</div><div class="sl">Pessoas</div></div><div class="stat"><div class="sv">${report.reduce((s,r)=>s+r.wpts,0).toFixed(1)}</div><div class="sl">Pts pond.</div></div><div class="stat"><div class="sv">${Object.values(weights).join("/")}</div><div class="sl">Pesos</div></div></div>
<div class="sec">Dedicação por projeto</div>
${report.map(r=>{const p=projects.find(x=>x.id===r.pid);const c=cont.find(x=>x.id===p?.cId);const col=COLORS[(cidx[r.pid]||0)%COLORS.length];return`<div class="proj"><div class="prow"><span><span class="pname">${r.name}</span><span class="pcli">${c?.name||""}</span></span><span class="ppct">${r.pct.toFixed(1)}%</span></div><div class="bar"><div class="fill" style="width:${r.pct}%;background:${col}"></div></div><div class="ppl">${r.people.join(", ")} · ${r.wpts} pts</div></div>`;}).join("")}
<div class="div"></div><div class="sec">Por colaborador</div>
${team.map(m=>{const pids=normA[m.id]||[];const w=wOf(m.role,weights);const names=pids.map(pid=>projects.find(p=>p.id===pid)?.name||pid);return`<div class="arow"><span class="an">${m.name}</span><span class="ar">${m.role} (${w})</span><span class="ap">${names.length?names.join(", "):"—"}</span></div>`;}).join("")}
<div class="foot">Gerado em ${new Date().toLocaleString("pt-BR")} · M+M Arquitetura</div>
</div><script>setTimeout(()=>window.print(),400)</script></body></html>`;
  const w=window.open("","_blank","width=820,height=720");if(w){w.document.write(html);w.document.close();}
}

// ── Estilos base ──────────────────────────────────────────────────────────────
const B={
  inp:{fontSize:"13px",padding:"8px 10px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:"#111",width:"100%",boxSizing:"border-box",outline:"none"},
  card:{background:"#fff",border:"1px solid #f0f0f0",borderRadius:"8px",padding:"12px 16px",marginBottom:"8px",boxShadow:"0 1px 2px rgba(0,0,0,.03)"},
  pri:{background:RED,color:"#fff",border:"none",borderRadius:"6px",padding:"8px 18px",fontSize:"13px",fontWeight:600,cursor:"pointer"},
  sec:{background:"transparent",color:"#6b7280",border:"1px solid #e5e7eb",borderRadius:"6px",padding:"7px 14px",fontSize:"13px",cursor:"pointer"},
  ghost:{background:"transparent",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:"12px",padding:"2px 6px"},
  lbl:{fontSize:"10px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".08em",marginBottom:"10px"},
};

// ── App principal ─────────────────────────────────────────────────────────────
export default function App(){
  const dbRef    = useRef(null);
  const [tab,setTab]          =useState("visao");
  const [cont,setCont]        =useState(INIT_C);
  const [projects,setProjects]=useState(INIT_P);
  const [team,setTeam]        =useState(INIT_T);
  const [weights,setWeights]  =useState(DW);
  const [diary,setDiaryState] =useState({});
  const [date,setDate]        =useState(TODAY);
  const [alloc,setAlloc]      =useState({});
  const [history,setHistory]  =useState({});
  const [loading,setLoading]  =useState(true);
  const [syncing,setSyncing]  =useState(false);
  const [flash,setFlash]      =useState("");
  const [aiMsg,setAiMsg]      =useState("");
  const [aiLoad,setAiLoad]    =useState(false);
  const [online,setOnline]    =useState(true);

  // ── Inicializa Firebase e ouve dados em tempo real ────────────────────────
  useEffect(()=>{
    // Verifica se o config foi preenchido
    if(firebaseConfig.apiKey==="COLE_AQUI"){
      setLoading(false);
      return;
    }
    try{
      const app=initializeApp(firebaseConfig);
      const db=getFirestore(app);
      dbRef.current=db;

      const cfgRef=doc(db,"mm_app","config");
      const histRef=doc(db,"mm_app","history");

      // Listener 1: config (projetos, equipe, contratantes, pesos, diário)
      const unsub1=onSnapshot(cfgRef,(snap)=>{
        if(!snap.exists()){
          // Primeiro acesso: inicializa o banco com os dados padrão
          setDoc(cfgRef,{cont:INIT_C,projects:INIT_P,team:INIT_T,weights:DW,diary:{}});
        } else {
          const d=snap.data();
          setCont(d.cont||INIT_C);
          setProjects(d.projects||INIT_P);
          setTeam(d.team||INIT_T);
          setWeights(d.weights||DW);
          setDiaryState(d.diary||{});
        }
      },(err)=>{ console.error(err); setOnline(false); });

      // Listener 2: histórico de alocações
      const unsub2=onSnapshot(histRef,(snap)=>{
        const h=snap.exists()?snap.data():{};
        setHistory(h);
        if(h[TODAY]) setAlloc(normAlloc(h[TODAY].alloc||{}));
        setLoading(false);
      },(err)=>{ console.error(err); setLoading(false); });

      return ()=>{ unsub1(); unsub2(); };
    }catch(err){
      console.error("Firebase error:",err);
      setLoading(false);
      setOnline(false);
    }
  },[]);

  useEffect(()=>{ setAlloc(normAlloc(history[date]?.alloc||{})); },[date]);

  const colorMap={};projects.forEach((p,i)=>{colorMap[p.id]=COLORS[i%COLORS.length];});
  const normA      =normAlloc(alloc);
  const report     =calcReport(alloc,team,projects,weights);
  const teamLoad   =calcTeamLoad(alloc,team,weights);
  const savedDate  =!!history[date]?.savedAt;
  const assignedCnt=team.filter(m=>(normA[m.id]||[]).length>0).length;

  function toggleProj(mid,pid){
    const curr=normA[mid]||[];
    setAlloc({...normA,[mid]:curr.includes(pid)?curr.filter(p=>p!==pid):[...curr,pid]});
  }

  async function persist(key,value){
    if(!dbRef.current)return;
    setSyncing(true);
    try{
      if(key==="history"){
        await setDoc(doc(dbRef.current,"mm_app","history"),value,{merge:true});
      } else {
        await setDoc(doc(dbRef.current,"mm_app","config"),{[key]:value},{merge:true});
      }
    }catch(e){ console.error(e); }
    setSyncing(false);
  }

  async function saveAlloc(){
    const entry={alloc:normA,savedAt:new Date().toISOString()};
    await persist("history",{[date]:entry});
    setFlash("Salvo!");setTimeout(()=>setFlash(""),1800);
  }
  const saveCont    =(c)=>{ setCont(c);     persist("cont",c);     };
  const saveProjects=(p)=>{ setProjects(p); persist("projects",p); };
  const saveTeam    =(t)=>{ setTeam(t);     persist("team",t);     };
  const saveWeights =(w)=>{ setWeights(w);  persist("weights",w);  };
  const saveDiary   =(d)=>{ setDiaryState(d); persist("diary",d);  };

  async function genMsg(){
    if(!report.length)return;
    const key=ANTHROPIC_KEY||prompt("Cole sua chave da API Anthropic (sk-ant-...):");
    if(!key)return;
    setAiLoad(true);setAiMsg("");
    try{
      const lines=report.map(r=>{const p=projects.find(x=>x.id===r.pid);const c=cont.find(x=>x.id===p?.cId);return`- ${r.name}${c?` (${c.name})`:""}: ${r.people.join(", ")} — ${r.pct.toFixed(0)}%`;}).join("\n");
      const blocked=projects.filter(p=>["bloqueado","aguardando"].includes(p.status)&&p.note).map(p=>`- ${p.name}: ${p.note}`).join("\n");
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":key,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:300,messages:[{role:"user",content:`Mensagem curta para o Manoel (sócio M+M) sobre alocação de ${fmtDate(date)}.\n\nDedicação:\n${lines}\n\nBloqueados:\n${blocked||"Nenhum"}\n\nTom WhatsApp, informal mas profissional. Máximo 4 linhas. Sem bullets.`}]})
      });
      const data=await res.json();
      setAiMsg(data.content?.find(c=>c.type==="text")?.text||"Erro.");
    }catch(e){ setAiMsg("Erro: "+e.message); }
    setAiLoad(false);
  }

  // ── Tela de configuração inicial ──────────────────────────────────────────
  if(!loading && firebaseConfig.apiKey==="COLE_AQUI"){
    return (
      <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",alignItems:"center",justifyContent:"center",padding:"2rem"}}>
        <div style={{background:"#fff",borderRadius:"16px",padding:"2.5rem",maxWidth:"480px",width:"100%",boxShadow:"0 4px 24px rgba(0,0,0,.08)",textAlign:"center"}}>
          <div style={{fontSize:"28px",fontWeight:800,letterSpacing:"-1px",marginBottom:"8px"}}>M<span style={{color:RED}}>+</span>M</div>
          <div style={{fontSize:"14px",color:"#6b7280",marginBottom:"2rem"}}>Configure o Firebase para ativar a sincronização em tempo real</div>
          <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:"8px",padding:"16px",textAlign:"left",fontSize:"13px",color:"#dc2626",lineHeight:1.7}}>
            <b>⚠️ Configuração necessária:</b><br/>
            Abra o arquivo <code style={{background:"#fee2e2",padding:"1px 4px",borderRadius:"3px"}}>App.jsx</code> e substitua os valores de <code style={{background:"#fee2e2",padding:"1px 4px",borderRadius:"3px"}}>firebaseConfig</code> pelos dados do seu projeto Firebase.<br/><br/>
            Siga o guia <b>INSTRUCOES.md</b> para criar o projeto Firebase gratuitamente.
          </div>
        </div>
      </div>
    );
  }

  if(loading) return (
    <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:"24px",fontWeight:800,marginBottom:"16px"}}>M<span style={{color:RED}}>+</span>M</div>
        <div style={{fontSize:"13px",color:"#9ca3af"}}>Conectando ao banco de dados…</div>
      </div>
    </div>
  );

  const TABS=[["visao","Visão Geral"],["hoje","Hoje"],["relatorio","Relatório"],["projetos","Projetos"],["equipe","Equipe"],["historico","Histórico"]];

  return (
    <div style={{background:"#f9fafb",minHeight:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>
      <div style={{height:"3px",background:RED}}/>
      {/* Header */}
      <div style={{background:"#fff",borderBottom:"1px solid #f0f0f0",padding:"0 1.25rem",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0 0"}}>
          <div style={{display:"flex",alignItems:"baseline",gap:"10px"}}>
            <span style={{fontSize:"22px",fontWeight:800,letterSpacing:"-1px",color:"#111"}}>M<span style={{color:RED}}>+</span>M</span>
            <span style={{fontSize:"11px",color:"#9ca3af",letterSpacing:".05em",textTransform:"uppercase",fontWeight:500}}>Arquitetura</span>
          </div>
          <div style={{display:"flex",gap:"12px",alignItems:"center"}}>
            {/* Indicador de sincronização */}
            <div style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"11px"}}>
              <div style={{width:"7px",height:"7px",borderRadius:"50%",background:online?(syncing?"#f59e0b":"#16a34a"):"#dc2626",animation:syncing?"pulse 1s infinite":"none"}}/>
              <span style={{color:"#9ca3af"}}>{syncing?"Salvando…":online?"Sincronizado":"Offline"}</span>
            </div>
            <div style={{display:"flex",gap:"12px",fontSize:"12px",color:"#9ca3af"}}>
              <span><b style={{color:"#111"}}>{projects.filter(p=>p.status!=="concluido").length}</b> projetos</span>
              <span><b style={{color:"#111"}}>{team.length}</b> pessoas</span>
            </div>
          </div>
        </div>
        <div style={{display:"flex",overflowX:"auto",marginTop:"4px"}}>
          {TABS.map(([id,l])=>(
            <button key={id} onClick={()=>setTab(id)} style={{background:"transparent",border:"none",cursor:"pointer",padding:"7px 12px",fontSize:"12px",fontWeight:tab===id?700:400,color:tab===id?"#111":"#9ca3af",borderBottom:tab===id?`2.5px solid ${RED}`:"2.5px solid transparent",whiteSpace:"nowrap",flexShrink:0}}>
              {l}
            </button>
          ))}
        </div>
      </div>

      <div style={{padding:"1.25rem",maxWidth:"800px",margin:"0 auto"}}>
        {tab==="visao"    && <DashTab report={report} teamLoad={teamLoad} projects={projects} cont={cont} team={team} history={history} colorMap={colorMap} weights={weights}/>}
        {tab==="hoje"     && <HojeTab team={team} projects={projects} cont={cont} normA={normA} colorMap={colorMap} toggleProj={toggleProj} date={date} setDate={setDate} savedDate={savedDate} assignedCnt={assignedCnt} saveAlloc={saveAlloc} flash={flash} report={report}/>}
        {tab==="relatorio"&& <RelTab report={report} date={date} setDate={setDate} projects={projects} cont={cont} normA={normA} team={team} weights={weights} assignedCnt={assignedCnt} colorMap={colorMap} genMsg={genMsg} aiLoad={aiLoad} aiMsg={aiMsg} exportExcel={()=>exportExcel(report,date,team,normA,projects,cont,weights)} exportPDF={()=>exportPDF(report,date,projects,cont,normA,team,weights)}/>}
        {tab==="projetos" && <ProjTab projects={projects} cont={cont} saveProjects={saveProjects} saveCont={saveCont} colorMap={colorMap} diary={diary} saveDiary={saveDiary} team={team} history={history}/>}
        {tab==="equipe"   && <EquipeTab team={team} saveTeam={saveTeam} weights={weights} saveWeights={saveWeights}/>}
        {tab==="historico"&& <HistTab history={history} projects={projects} cont={cont} team={team} colorMap={colorMap}/>}
      </div>

      <style>{`@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}`}</style>
    </div>
  );
}

// ── Visão Geral ────────────────────────────────────────────────────────────────
function DashTab({report,teamLoad,projects,cont,team,history,colorMap,weights}){
  const active =projects.filter(p=>p.status==="ativo").length;
  const blocked=projects.filter(p=>p.status==="bloqueado").length;
  const waiting=projects.filter(p=>p.status==="aguardando").length;
  const done   =projects.filter(p=>p.status==="concluido").length;
  const chartProj=report.map(r=>({name:r.name.length>22?r.name.slice(0,20)+"…":r.name,pct:+r.pct.toFixed(1),pid:r.pid,wpts:r.wpts}));
  const chartTeam=teamLoad.filter(m=>m.active).sort((a,b)=>b.pct-a.pct);
  async function exportPNG(){
const el=document.getElementById("dashboard-content");
if(!el)return;
const canvas=await html2canvas(el,{backgroundColor:"#f9fafb",scale:2});
const link=document.createElement("a");
link.download=MM_Dashboard_${TODAY}.png;
link.href=canvas.toDataURL("image/png");
link.click();
}
const statusPie=[
    {name:"Ativos",value:active,fill:STATUS.ativo.hex},
    {name:"Bloqueados",value:blocked,fill:STATUS.bloqueado.hex},
    {name:"Aguardando",value:waiting,fill:STATUS.aguardando.hex},
    {name:"Concluídos",value:done,fill:STATUS.concluido.hex},
  ].filter(d=>d.value>0);
  return(
    <div>
  <div style={{display:"flex",justifyContent:"flex-end",marginBottom:"12px"}}>
    <button onClick={exportPNG} style={{...B.sec,fontSize:"12px",padding:"6px 14px",display:"flex",alignItems:"center",gap:"5px"}}>↓ Exportar PNG</button>
  </div>
  <div id="dashboard-content">
  <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"8px",marginBottom:"1.25rem"}}>
        {[["Ativos",active,STATUS.ativo],["Bloqueados",blocked,STATUS.bloqueado],["Aguardando",waiting,STATUS.aguardando],["Concluídos",done,STATUS.concluido]].map(([l,v,s])=>(
          <div key={l} style={{...B.card,textAlign:"center",borderTop:`3px solid ${s.hex}`,padding:"12px 8px"}}>
            <div style={{fontSize:"9px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#9ca3af",marginBottom:"6px"}}>{l}</div>
            <div style={{fontSize:"26px",fontWeight:800,color:s.color}}>{v}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px",marginBottom:"1.25rem"}}>
        <div style={{...B.card,textAlign:"center"}}><div style={B.lbl}>Equipe</div><div style={{fontSize:"22px",fontWeight:700}}>{team.length}</div><div style={{fontSize:"11px",color:"#9ca3af"}}>{teamLoad.filter(m=>m.active).length} ativos hoje</div></div>
        <div style={{...B.card,textAlign:"center"}}><div style={B.lbl}>Contratantes</div><div style={{fontSize:"22px",fontWeight:700}}>{cont.length}</div></div>
        <div style={{...B.card,textAlign:"center"}}><div style={B.lbl}>Dias registrados</div><div style={{fontSize:"22px",fontWeight:700}}>{Object.keys(history).length}</div></div>
      </div>
      {chartProj.length>0&&(
        <div style={{...B.card,marginBottom:"1rem"}}>
          <div style={B.lbl}>Dedicação por projeto (ponderada)</div>
          <ResponsiveContainer width="100%" height={Math.max(160,chartProj.length*34)}>
            <BarChart data={chartProj} layout="vertical" margin={{top:0,right:48,left:8,bottom:0}}>
              <XAxis type="number" domain={[0,100]} tickFormatter={v=>`${v}%`} tick={{fontSize:10}} axisLine={false} tickLine={false}/>
              <YAxis type="category" dataKey="name" width={135} tick={{fontSize:11}} axisLine={false} tickLine={false}/>
              <Tooltip formatter={(v,n,p)=>[`${v}% (${p.payload.wpts} pts)`,"Dedicação"]} contentStyle={{fontSize:12,borderRadius:6}}/>
              <Bar dataKey="pct" radius={[0,4,4,0]} maxBarSize={18}>
                {chartProj.map((e,i)=><Cell key={i} fill={colorMap[e.pid]||COLORS[i%COLORS.length]}/>)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
      {chartTeam.length>0&&(
        <div style={{...B.card,marginBottom:"1rem"}}>
          <div style={B.lbl}>Carga da equipe</div>
          <div style={{fontSize:"11px",color:"#9ca3af",marginBottom:"12px"}}>
            Pesos: {Object.entries(weights).map(([r,w])=><span key={r} style={{marginRight:"10px"}}><b>{r}</b>={w}</span>)}
          </div>
          {chartTeam.map(m=>{const rc=ROLE_COLOR[m.role]||"#888";return(
            <div key={m.name} style={{marginBottom:"12px"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"4px"}}>
                <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                  <span style={{fontSize:"13px",fontWeight:600}}>{m.name}</span>
                  <span style={{fontSize:"10px",padding:"1px 7px",borderRadius:"10px",background:rc+"18",color:rc,fontWeight:600}}>{m.role} · {m.weight}</span>
                </div>
                <span style={{fontSize:"13px",fontWeight:700,color:"#374151"}}>{m.pct.toFixed(1)}%</span>
              </div>
              <div style={{height:"7px",background:"#f3f4f6",borderRadius:"4px",overflow:"hidden"}}>
                <div style={{height:"100%",width:`${Math.min(m.pct,100)}%`,background:rc,borderRadius:"4px"}}/>
              </div>
            </div>
          );})}
        </div>
      )}
      {statusPie.length>0&&(
        <div style={B.card}>
          <div style={B.lbl}>Status dos projetos</div>
          <div style={{display:"flex",alignItems:"center",gap:"16px"}}>
            <PieChart width={110} height={110}>
              <Pie data={statusPie} cx={50} cy={50} innerRadius={26} outerRadius={48} dataKey="value" paddingAngle={2}>
                {statusPie.map((e,i)=><Cell key={i} fill={e.fill}/>)}
              </Pie>
            </PieChart>
            <div style={{flex:1}}>
              {statusPie.map(s=>(
                <div key={s.name} style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"6px"}}>
                  <div style={{width:"10px",height:"10px",borderRadius:"2px",background:s.fill,flexShrink:0}}/>
                  <span style={{fontSize:"12px",color:"#6b7280",flex:1}}>{s.name}</span>
                  <span style={{fontSize:"14px",fontWeight:700}}>{s.value}</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ── Hoje ──────────────────────────────────────────────────────────────────────
function HojeTab({team,projects,cont,normA,colorMap,toggleProj,date,setDate,savedDate,assignedCnt,saveAlloc,flash,report}){
  return <>
    <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"1.25rem",flexWrap:"wrap"}}>
      <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{fontSize:"13px",padding:"6px 10px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:"#111"}}/>
      <span style={{fontSize:"13px",color:"#6b7280"}}>{cap(fmtDay(date))}</span>
      {savedDate&&<span style={{fontSize:"11px",background:"#dcfce7",color:"#15803d",padding:"2px 8px",borderRadius:"20px",fontWeight:600}}>✓ Salvo</span>}
      <span style={{marginLeft:"auto",fontSize:"12px",color:"#9ca3af"}}>{assignedCnt}/{team.length} alocados</span>
    </div>
    {ROLES.map(role=>{
      const members=team.filter(m=>m.role===role);
      if(!members.length)return null;
      const rc=ROLE_COLOR[role]||"#888";
      return(
        <div key={role} style={{marginBottom:"1.5rem"}}>
          <div style={{display:"flex",alignItems:"center",gap:"7px",marginBottom:"8px"}}>
            <div style={{width:"3px",height:"14px",borderRadius:"2px",background:rc}}/>
            <span style={{fontSize:"11px",fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:".07em"}}>{role}</span>
          </div>
          {members.map(m=>{
            const pids=normA[m.id]||[];const sCol=pids.length===1?colorMap[pids[0]]:null;
            return(
              <div key={m.id} style={{...B.card,display:"flex",alignItems:"flex-start",gap:"12px"}}>
                <div style={{width:"36px",height:"36px",borderRadius:"50%",flexShrink:0,marginTop:"1px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"11px",fontWeight:700,background:sCol?sCol+"20":pids.length>1?RED+"12":"#f3f4f6",color:sCol||(pids.length>1?RED:"#9ca3af"),border:`1.5px solid ${sCol?sCol+"30":pids.length>1?RED+"25":"#e5e7eb"}`}}>
                  {m.name.slice(0,2).toUpperCase()}
                </div>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:"14px",fontWeight:600,color:"#111",marginBottom:"6px"}}>{m.name}</div>
                  {pids.length>0&&(
                    <div style={{display:"flex",flexWrap:"wrap",gap:"4px",marginBottom:"6px"}}>
                      {pids.map(pid=>{
                        const proj=projects.find(p=>p.id===pid);const c=colorMap[pid];
                        const ctr=cont.find(x=>x.id===proj?.cId);
                        return(
                          <span key={pid} style={{fontSize:"11px",padding:"3px 7px 3px 9px",borderRadius:"20px",background:c+"18",color:c,border:`1px solid ${c}35`,display:"inline-flex",alignItems:"center",gap:"3px",lineHeight:1.5,fontWeight:500}}>
                            {ctr?`${ctr.name} · `:""}{proj?.name||pid}
                            <button onClick={()=>toggleProj(m.id,pid)} style={{background:"none",border:"none",cursor:"pointer",color:"inherit",padding:"0 1px",fontSize:"13px",opacity:.5}}>×</button>
                          </span>
                        );
                      })}
                    </div>
                  )}
                  <select onChange={e=>{if(e.target.value){toggleProj(m.id,e.target.value);e.target.value="";}}} style={{fontSize:"12px",padding:"3px 8px",borderRadius:"6px",border:"1px dashed #e5e7eb",background:"transparent",color:"#9ca3af",cursor:"pointer"}}>
                    <option value="">+ adicionar projeto</option>
                    {cont.map(ct=>{
                      const ps=projects.filter(p=>p.cId===ct.id&&p.status!=="concluido"&&!pids.includes(p.id));
                      if(!ps.length)return null;
                      return <optgroup key={ct.id} label={ct.name}>{ps.map(p=><option key={p.id} value={p.id}>{p.name}{p.status!=="ativo"?` (${STATUS[p.status]?.label})`:""}</option>)}</optgroup>;
                    })}
                  </select>
                </div>
              </div>
            );
          })}
        </div>
      );
    })}
    <div style={{display:"flex",justifyContent:"flex-end",marginBottom:"1.5rem"}}>
      <button onClick={saveAlloc} style={{...B.pri,background:flash?"#16a34a":RED,minWidth:"140px"}}>{flash||"Salvar alocação"}</button>
    </div>
    {report.length>0&&<>
      <div style={B.lbl}>Prévia ponderada</div>
      {report.map(r=>{
        const p=projects.find(x=>x.id===r.pid);const c=cont.find(x=>x.id===p?.cId);
        return(
          <div key={r.pid} style={{marginBottom:"12px"}}>
            <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
              <div><span style={{fontSize:"13px",fontWeight:600,color:"#111"}}>{r.name}</span>{c&&<span style={{fontSize:"11px",color:"#9ca3af",marginLeft:"5px"}}>{c.name}</span>}</div>
              <span style={{fontSize:"12px",fontWeight:700,color:"#6b7280"}}>{r.pct.toFixed(1)}%</span>
            </div>
            <div style={{height:"5px",background:"#f3f4f6",borderRadius:"3px",overflow:"hidden"}}>
              <div style={{height:"100%",width:`${r.pct}%`,background:colorMap[r.pid],borderRadius:"3px"}}/>
            </div>
            <div style={{fontSize:"11px",color:"#9ca3af",marginTop:"2px"}}>{r.people.join(", ")} · {r.wpts} pts</div>
          </div>
        );
      })}
    </>}
  </>;
}

// ── Relatório ─────────────────────────────────────────────────────────────────
function RelTab({report,date,setDate,projects,cont,normA,team,weights,assignedCnt,colorMap,genMsg,aiLoad,aiMsg,exportExcel,exportPDF}){
  const totalPts=report.reduce((s,r)=>s+r.wpts,0);
  return <>
    <div style={{display:"flex",gap:"10px",alignItems:"center",marginBottom:"1.5rem"}}>
      <input type="date" value={date} onChange={e=>setDate(e.target.value)} style={{fontSize:"13px",padding:"6px 10px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:"#111"}}/>
      <span style={{fontSize:"13px",color:"#6b7280"}}>{cap(fmtDay(date))}, {fmtDate(date)}</span>
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px",marginBottom:"1.5rem"}}>
      {[["Projetos",report.length],["Pessoas",assignedCnt],["Pts pond.",totalPts.toFixed(1)]].map(([l,v])=>(
        <div key={l} style={{background:"#f9fafb",borderRadius:"8px",padding:"14px",textAlign:"center",border:"1px solid #f0f0f0"}}>
          <div style={{fontSize:"9px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#9ca3af",marginBottom:"6px"}}>{l}</div>
          <div style={{fontSize:"24px",fontWeight:700,color:"#111"}}>{v}</div>
        </div>
      ))}
    </div>
    {report.length===0?(
      <div style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}>Nenhuma alocação para esta data.</div>
    ):<>
      {report.map(r=>{
        const p=projects.find(x=>x.id===r.pid);const c=cont.find(x=>x.id===p?.cId);
        return(
          <div key={r.pid} style={{marginBottom:"18px"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"5px"}}>
              <div><div style={{fontSize:"14px",fontWeight:700,color:"#111"}}>{r.name}</div>{c&&<div style={{fontSize:"11px",color:"#9ca3af",marginTop:"1px"}}>{c.name}</div>}</div>
              <span style={{fontSize:"13px",fontWeight:700,color:"#6b7280",marginLeft:"8px",flexShrink:0}}>{r.pct.toFixed(1)}%</span>
            </div>
            <div style={{height:"8px",background:"#f3f4f6",borderRadius:"4px",overflow:"hidden",marginBottom:"4px"}}>
              <div style={{height:"100%",width:`${r.pct}%`,background:colorMap[r.pid],borderRadius:"4px"}}/>
            </div>
            <div style={{fontSize:"12px",color:"#9ca3af"}}>{r.detail.map(d=>`${d.name} (${d.role}, ${d.w} pts)`).join(" · ")}</div>
          </div>
        );
      })}
      <div style={{borderTop:"1px solid #f0f0f0",paddingTop:"1.25rem",marginTop:"0.5rem"}}>
        <div style={B.lbl}>Exportar</div>
        <div style={{display:"flex",gap:"8px",flexWrap:"wrap",marginBottom:"10px"}}>
          <button onClick={exportExcel} style={{...B.pri,display:"flex",alignItems:"center",gap:"5px"}}>↓ Excel</button>
          <button onClick={exportPDF} style={{...B.sec,color:"#374151",fontWeight:500,display:"flex",alignItems:"center",gap:"5px"}}>↓ PDF / Imprimir</button>
        </div>
        <button onClick={genMsg} disabled={aiLoad} style={{...B.sec,color:aiLoad?"#9ca3af":RED,width:"100%",textAlign:"center"}}>
          {aiLoad?"Gerando…":"Gerar mensagem para o Manoel ↗"}
        </button>
        {aiMsg&&(
          <div style={{background:"#f9fafb",borderRadius:"8px",padding:"12px 14px",marginTop:"10px",fontSize:"13px",lineHeight:1.7,whiteSpace:"pre-wrap",border:"1px solid #f0f0f0"}}>
            {aiMsg}
            <button onClick={()=>navigator.clipboard?.writeText(aiMsg)} style={{...B.ghost,display:"block",marginTop:"6px",paddingTop:"6px",borderTop:"1px solid #f0f0f0",width:"100%",textAlign:"left"}}>Copiar mensagem</button>
          </div>
        )}
      </div>
    </>}
  </>;
}

// ── Projetos ──────────────────────────────────────────────────────────────────
function ProjTab({projects,cont,saveProjects,saveCont,colorMap,diary,saveDiary,team,history}){
  const [view,setView]       =useState("cont");
  const [showForm,setShowForm]=useState(false);
  const [editProj,setEditProj]=useState(null);
  const [detail,setDetail]   =useState(null);
  const [showDone,setShowDone]=useState(false);
  const [search,setSearch]   =useState("");
  const [editCont,setEditCont]=useState(null);
  const active  =projects.filter(p=>p.status!=="concluido");
  const done    =projects.filter(p=>p.status==="concluido");
  const filtered=search?active.filter(p=>p.name.toLowerCase().includes(search.toLowerCase())||(cont.find(c=>c.id===p.cId)?.name||"").toLowerCase().includes(search.toLowerCase())):active;
  function cycleStatus(proj){saveProjects(projects.map(p=>p.id===proj.id?{...p,status:ST_CYC[p.status]||"ativo"}:p));}
  function baixa(proj){saveProjects(projects.map(p=>p.id===proj.id?{...p,status:"concluido"}:p));}
  if(detail){
    const proj=projects.find(p=>p.id===detail);
    const ctr=cont.find(c=>c.id===proj?.cId);
    if(!proj)return null;
    return <ProjDetail proj={proj} ctr={ctr} colorMap={colorMap} diary={diary} saveDiary={saveDiary} team={team} onBack={()=>setDetail(null)} projects={projects} saveProjects={saveProjects} history={history}/>;
  }
  return <>
    <div style={{display:"flex",gap:"8px",marginBottom:"1.25rem",flexWrap:"wrap"}}>
      <input placeholder="Buscar…" value={search} onChange={e=>setSearch(e.target.value)} style={{...B.inp,width:"auto",flex:1,minWidth:"160px"}}/>
      <button onClick={()=>setShowForm(v=>!v)} style={B.pri}>+ Novo projeto</button>
    </div>
    <div style={{display:"flex",gap:"0",marginBottom:"1.25rem",background:"#f3f4f6",borderRadius:"8px",padding:"3px",width:"fit-content"}}>
      {[["cont","Por contratante"],["status","Por status"]].map(([v,l])=>(
        <button key={v} onClick={()=>setView(v)} style={{background:view===v?"#fff":"transparent",border:"none",cursor:"pointer",padding:"5px 12px",borderRadius:"6px",fontSize:"12px",fontWeight:view===v?600:400,color:view===v?"#111":"#9ca3af",boxShadow:view===v?"0 1px 3px rgba(0,0,0,.07)":"none"}}>{l}</button>
      ))}
    </div>
    {showForm&&<ProjForm cont={cont} saveCont={saveCont} onSave={p=>{saveProjects([...projects,{...p,id:"p"+uid()}]);setShowForm(false);}} onCancel={()=>setShowForm(false)}/>}
    {view==="cont"&&cont.map(ct=>{
      const ps=filtered.filter(p=>p.cId===ct.id);if(!ps.length)return null;
      return(
        <div key={ct.id} style={{marginBottom:"1.75rem"}}>
          <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px",paddingBottom:"8px",borderBottom:`2px solid ${RED}20`}}>
            <div style={{width:"6px",height:"6px",borderRadius:"50%",background:RED}}/>
            {editCont===ct.id?(
          <input defaultValue={ct.name} autoFocus onBlur={e=>{saveCont(cont.map(c=>c.id===ct.id?{...c,name:e.target.value}:c));setEditCont(null);}} style={{...B.inp,fontSize:"14px",fontWeight:700,padding:"2px 6px",width:"auto",flex:1}} onKeyDown={e=>e.key==="Enter"&&e.target.blur()}/>
        ):(
          <span style={{fontSize:"14px",fontWeight:700,color:"#111",cursor:"text"}} onDoubleClick={()=>setEditCont(ct.id)}>{ct.name}</span>
        )}
        <span style={{fontSize:"11px",color:"#9ca3af"}}>{ps.length} proj.</span>
        <div style={{marginLeft:"auto",display:"flex",gap:"8px",alignItems:"center"}}>
          <span style={{fontSize:"9px",color:"#d1d5db"}}>duplo clique para editar</span>
          {ps.length===0&&(
            <button onClick={()=>{if(window.confirm(`Apagar contratante "${ct.name}"?`))saveCont(cont.filter(c=>c.id!==ct.id));}} style={{...B.ghost,color:"#dc2626",fontSize:"11px",padding:"1px 6px",border:"1px solid #fecaca",borderRadius:"4px"}}>apagar</button>
          )}
        </div>
          </div>
          {ps.map(proj=><ProjCard key={proj.id} proj={proj} cont={cont} colorMap={colorMap} editProj={editProj} setEditProj={setEditProj} cycleStatus={cycleStatus} baixa={baixa} projects={projects} saveProjects={saveProjects} saveCont={saveCont} diaryCount={(diary[proj.id]||[]).length} onDetail={()=>setDetail(proj.id)}/>)}
        </div>
      );
    })}
    {view==="status"&&["ativo","aguardando","bloqueado"].map(sk=>{
      const ps=filtered.filter(p=>p.status===sk);if(!ps.length)return null;
      return(
        <div key={sk} style={{marginBottom:"1.75rem"}}>
          <div style={{...B.lbl,display:"flex",alignItems:"center",gap:"6px"}}><span style={{display:"inline-block",width:"6px",height:"6px",borderRadius:"50%",background:STATUS[sk].hex}}/>{STATUS[sk].label} ({ps.length})</div>
          {ps.map(proj=><ProjCard key={proj.id} proj={proj} cont={cont} colorMap={colorMap} editProj={editProj} setEditProj={setEditProj} cycleStatus={cycleStatus} baixa={baixa} projects={projects} saveProjects={saveProjects} saveCont={saveCont} diaryCount={(diary[proj.id]||[]).length} onDetail={()=>setDetail(proj.id)}/>)}
        </div>
      );
    })}
    {done.length>0&&(
      <div>
        <button onClick={()=>setShowDone(v=>!v)} style={{...B.ghost,display:"flex",alignItems:"center",gap:"5px",color:"#9ca3af"}}>
          <span style={{fontSize:"9px"}}>{showDone?"▾":"▸"}</span> Concluídos ({done.length})
        </button>
        {showDone&&done.map(proj=>(
          <div key={proj.id} style={{...B.card,opacity:.55,background:"#f9fafb",marginTop:"6px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:"8px"}}>
              <div>
                <div style={{fontSize:"13px",color:"#6b7280",textDecoration:"line-through"}}>{proj.name}</div>
                <div style={{fontSize:"11px",color:"#9ca3af"}}>{cont.find(c=>c.id===proj.cId)?.name}</div>
              </div>
              <button onClick={()=>saveProjects(projects.map(p=>p.id===proj.id?{...p,status:"ativo"}:p))} style={{...B.sec,fontSize:"11px",padding:"3px 10px"}}>reativar</button>
            </div>
          </div>
        ))}
      </div>
    )}
  </>;
}

function ProjCard({proj,cont,colorMap,editProj,setEditProj,cycleStatus,baixa,projects,saveProjects,saveCont,diaryCount,onDetail}){
  const sc=STATUS[proj.status]||STATUS.ativo;
  const color=colorMap[proj.id];
  if(editProj?.id===proj.id) return <ProjForm cont={cont} saveCont={saveCont} initial={editProj} onSave={p=>{saveProjects(projects.map(x=>x.id===editProj.id?{...editProj,...p}:x));setEditProj(null);}} onCancel={()=>setEditProj(null)} onDelete={()=>{saveProjects(projects.filter(x=>x.id!==proj.id));setEditProj(null);}}/>;
  return(
    <div style={{...B.card,borderLeft:`3px solid ${color}`}}>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:"10px"}}>
        <div style={{flex:1,minWidth:0,cursor:"pointer"}} onClick={onDetail}>
          <div style={{fontSize:"14px",fontWeight:600,color:"#111"}}>{proj.name}</div>
          {(proj.city||proj.st)&&<div style={{fontSize:"11px",color:"#9ca3af",marginTop:"2px"}}>{[proj.city,proj.st,proj.country].filter(Boolean).join(", ")}</div>}
          {proj.svcs?.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:"3px",marginTop:"6px"}}>{proj.svcs.map(s=><span key={s} style={{fontSize:"10px",padding:"2px 6px",borderRadius:"4px",background:"#f3f4f6",color:"#6b7280",border:"1px solid #e5e7eb"}}>{s}</span>)}</div>}
          {proj.note&&<div style={{fontSize:"12px",color:"#9ca3af",marginTop:"5px",fontStyle:"italic"}}>{proj.note}</div>}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:"5px",alignItems:"flex-end",flexShrink:0}}>
          {proj.status!=="concluido"&&(
            <button onClick={()=>cycleStatus(proj)} style={{fontSize:"11px",fontWeight:600,padding:"3px 9px",borderRadius:"20px",background:sc.bg,color:sc.color,border:"none",cursor:"pointer",display:"inline-flex",alignItems:"center",gap:"4px"}}>
              {sc.label}<span style={{fontSize:"9px",opacity:.55}}>↻</span>
            </button>
          )}
          <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
            {diaryCount>0&&<span style={{fontSize:"10px",padding:"1px 6px",borderRadius:"10px",background:"#eff6ff",color:"#2563EB",fontWeight:600,cursor:"pointer"}} onClick={onDetail}>{diaryCount} entrada{diaryCount!==1?"s":""}</span>}
            <button onClick={()=>setEditProj({...proj})} style={B.ghost}>editar</button>
            <button onClick={()=>baixa(proj)} style={B.ghost}>baixa</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ProjForm({cont,saveCont,initial,onSave,onCancel,onDelete}){
  const blank={cId:"",name:"",svcs:[],city:"",st:"",country:"Brasil",status:"ativo",note:""};
  const [f,setF]=useState(initial||blank);
  const [nc,setNc]=useState("");const [addC,setAddC]=useState(false);
const ncRef=useRef(null);
function tSvc(s){setF(p=>({...p,svcs:p.svcs.includes(s)?p.svcs.filter(x=>x!==s):[...p.svcs,s]}));}
async function createCont(){if(!nc.trim())return;const c={id:"c"+uid(),name:nc.trim(),contact:""};await saveCont([...cont,c]);setF(p=>({...p,cId:c.id}));setNc("");setAddC(false);}
const Lbl=({t,children})=><div style={{marginBottom:"10px"}}><div style={{fontSize:"10px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".06em",marginBottom:"5px"}}>{t}</div>{children}</div>;
  return(
    <div style={{background:"#fff",borderRadius:"10px",padding:"18px",marginBottom:"1.25rem",border:`1px solid ${RED}30`,boxShadow:"0 2px 8px rgba(200,16,46,.08)"}}>
      <div style={{fontSize:"13px",fontWeight:700,color:"#111",marginBottom:"14px",paddingBottom:"10px",borderBottom:"1px solid #f0f0f0"}}>{initial?"Editar projeto":"Novo projeto"}</div>
      <Lbl t="Contratante *">
        {addC?(
      <div style={{display:"flex",gap:"6px"}}><input ref={ncRef} placeholder="Nome" value={nc} onChange={e=>setNc(e.target.value)} onKeyDown={e=>e.key==="Enter"&&createCont()} style={B.inp}/><button onClick={createCont} style={{...B.pri,padding:"7px 14px"}}>Criar</button><button onClick={()=>setAddC(false)} style={{...B.sec,padding:"7px 10px"}}>×</button></div>
        ):(
          <div style={{display:"flex",gap:"6px"}}><select value={f.cId} onChange={e=>setF(p=>({...p,cId:e.target.value}))} style={{...B.inp,flex:1}}><option value="">Selecionar…</option>{cont.map(c=><option key={c.id} value={c.id}>{c.name}</option>)}</select><button onClick={()=>setAddC(true)} style={{...B.sec,fontSize:"12px",padding:"7px 10px"}}>+ Novo</button></div>
        )}
      </Lbl>
      <Lbl t="Nome *"><input placeholder="Ex: Residência Lago Sul" value={f.name} onChange={e=>setF(p=>({...p,name:e.target.value}))} style={B.inp}/></Lbl>
      <Lbl t="Serviços">
        <div style={{display:"flex",flexWrap:"wrap",gap:"5px"}}>
          {SERVICES.map(s=>{const on=f.svcs.includes(s);return(<button key={s} onClick={()=>tSvc(s)} style={{fontSize:"11px",padding:"4px 10px",borderRadius:"20px",border:`1px solid ${on?RED:"#e5e7eb"}`,background:on?RED+"18":"transparent",color:on?RED:"#6b7280",cursor:"pointer",fontWeight:on?600:400}}>{s}</button>);})}
        </div>
      </Lbl>
      <Lbl t="Localização">
    <div style={{display:"grid",gridTemplateColumns:"1fr 90px 100px",gap:"6px"}}>
      <input placeholder="Cidade" value={f.city} onChange={e=>setF(p=>({...p,city:e.target.value}))} style={B.inp}/>
      <select value={f.st} onChange={e=>setF(p=>({...p,st:e.target.value}))} style={B.inp}>
        <option value="">UF</option>
        {UFS.map(u=><option key={u} value={u}>{u}</option>)}
      </select>
      <input placeholder="País" value={f.country} onChange={e=>setF(p=>({...p,country:e.target.value}))} style={B.inp}/>
    </div>
  </Lbl>
      <Lbl t="Status">
        <div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>
          {["ativo","aguardando","bloqueado"].map(s=>{const sc=STATUS[s];const on=f.status===s;return(<button key={s} onClick={()=>setF(p=>({...p,status:s}))} style={{fontSize:"12px",padding:"5px 12px",borderRadius:"20px",border:`1px solid ${on?sc.hex:"#e5e7eb"}`,background:on?sc.bg:"transparent",color:on?sc.color:"#9ca3af",cursor:"pointer",fontWeight:on?600:400}}>{sc.label}</button>);})}</div>
      </Lbl>
      <Lbl t="Observação"><textarea placeholder="Bloqueios, pendências, notas…" value={f.note} onChange={e=>setF(p=>({...p,note:e.target.value}))} style={{...B.inp,height:"60px",resize:"vertical"}}/></Lbl>
      <div style={{display:"flex",gap:"8px"}}>
        <button onClick={()=>f.name.trim()&&f.cId&&onSave(f)} style={B.pri} disabled={!f.name.trim()||!f.cId}>Salvar</button>
        <button onClick={onCancel} style={B.sec}>Cancelar</button>
        {onDelete&&<button onClick={onDelete} style={{...B.ghost,color:"#dc2626",border:"1px solid #fecaca",borderRadius:"6px",padding:"7px 12px",marginLeft:"auto"}}>Excluir</button>}
      </div>
    </div>
  );
}

function ProjDetail({proj,ctr,colorMap,diary,saveDiary,team,onBack,projects,saveProjects,history}){
  const entries=(diary[proj.id]||[]).sort((a,b)=>b.date.localeCompare(a.date));
  const [showForm,setShowForm]=useState(false);
  const [eDate,setEDate]=useState(TODAY);
  const [eAuthor,setEAuthor]=useState(team[0]?.name||"");
  const [eText,setEText]=useState("");
  const color=colorMap[proj.id];const sc=STATUS[proj.status]||STATUS.ativo;
  const daysWorked=Object.entries(history).filter(([,h])=>{const n=normAlloc(h.alloc||{});return Object.values(n).some(pids=>pids.includes(proj.id));}).length;
  function addEntry(){
    if(!eText.trim())return;
    const entry={id:"e"+uid(),date:eDate,author:eAuthor,text:eText.trim(),createdAt:new Date().toISOString()};
    saveDiary({...diary,[proj.id]:[...(diary[proj.id]||[]),entry]});
    setEText("");setShowForm(false);
  }
  function delEntry(id){saveDiary({...diary,[proj.id]:(diary[proj.id]||[]).filter(e=>e.id!==id)});}
  return(
    <div>
      <button onClick={onBack} style={{...B.ghost,color:RED,padding:0,fontWeight:600,fontSize:"13px",marginBottom:"1rem"}}>← Projetos</button>
      <div style={{...B.card,borderLeft:`4px solid ${color}`,marginBottom:"1.25rem"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:"12px"}}>
          <div>
            <div style={{fontSize:"18px",fontWeight:700,color:"#111"}}>{proj.name}</div>
            {ctr&&<div style={{fontSize:"13px",color:"#9ca3af",marginTop:"2px"}}>{ctr.name}</div>}
            {(proj.city||proj.st)&&<div style={{fontSize:"12px",color:"#9ca3af",marginTop:"2px"}}>{[proj.city,proj.st,proj.country].filter(Boolean).join(", ")}</div>}
          </div>
          <button onClick={()=>saveProjects(projects.map(p=>p.id===proj.id?{...p,status:ST_CYC[p.status]||"ativo"}:p))} style={{fontSize:"12px",fontWeight:600,padding:"4px 12px",borderRadius:"20px",background:sc.bg,color:sc.color,border:"none",cursor:"pointer",flexShrink:0}}>
            {sc.label} ↻
          </button>
        </div>
        {proj.svcs?.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:"3px",marginTop:"10px"}}>{proj.svcs.map(s=><span key={s} style={{fontSize:"11px",padding:"3px 8px",borderRadius:"4px",background:"#f3f4f6",color:"#6b7280",border:"1px solid #e5e7eb"}}>{s}</span>)}</div>}
        {proj.note&&<div style={{fontSize:"12px",color:"#9ca3af",marginTop:"8px",padding:"8px 10px",borderRadius:"6px",background:"#f9fafb",fontStyle:"italic"}}>{proj.note}</div>}
        <div style={{marginTop:"12px",paddingTop:"10px",borderTop:"1px solid #f0f0f0",display:"flex",gap:"20px"}}>
          <div style={{textAlign:"center"}}><div style={{fontSize:"18px",fontWeight:700}}>{daysWorked}</div><div style={{fontSize:"10px",color:"#9ca3af",textTransform:"uppercase",letterSpacing:".05em"}}>dias trabalhados</div></div>
          <div style={{textAlign:"center"}}><div style={{fontSize:"18px",fontWeight:700}}>{entries.length}</div><div style={{fontSize:"10px",color:"#9ca3af",textTransform:"uppercase",letterSpacing:".05em"}}>entradas no diário</div></div>
        </div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
        <div style={B.lbl}>Diário de obra</div>
        <button onClick={()=>setShowForm(v=>!v)} style={{...B.pri,fontSize:"12px",padding:"6px 14px"}}>+ Nova entrada</button>
      </div>
      {showForm&&(
        <div style={{background:"#fff",borderRadius:"10px",padding:"16px",marginBottom:"1rem",border:`1px solid ${RED}25`,boxShadow:"0 2px 8px rgba(200,16,46,.06)"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"8px"}}>
            <input type="date" value={eDate} onChange={e=>setEDate(e.target.value)} style={B.inp}/>
            <select value={eAuthor} onChange={e=>setEAuthor(e.target.value)} style={B.inp}>{team.map(m=><option key={m.id} value={m.name}>{m.name}</option>)}</select>
          </div>
          <textarea placeholder="O que foi feito, decisões, pendências, observações…" value={eText} onChange={e=>setEText(e.target.value)} style={{...B.inp,height:"80px",resize:"vertical",marginBottom:"8px"}}/>
          <div style={{display:"flex",gap:"8px"}}>
            <button onClick={addEntry} style={B.pri} disabled={!eText.trim()}>Salvar entrada</button>
            <button onClick={()=>setShowForm(false)} style={B.sec}>Cancelar</button>
          </div>
        </div>
      )}
      {entries.length===0?(
        <div style={{textAlign:"center",padding:"2.5rem",color:"#9ca3af",fontSize:"13px"}}>Nenhuma entrada ainda.<br/>Registre reuniões, decisões e atualizações aqui.</div>
      ):entries.map(e=>(
        <div key={e.id} style={{...B.card,borderLeft:"3px solid #e5e7eb"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"8px"}}>
            <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
              <span style={{fontSize:"12px",fontWeight:700,color:RED}}>{fmtDate(e.date)}</span>
              <span style={{fontSize:"11px",padding:"1px 7px",borderRadius:"10px",background:"#f3f4f6",color:"#6b7280",fontWeight:600}}>{e.author}</span>
            </div>
            <button onClick={()=>delEntry(e.id)} style={{...B.ghost,fontSize:"16px",opacity:.3,lineHeight:1}}>×</button>
          </div>
          <div style={{fontSize:"13px",color:"#374151",lineHeight:1.6,whiteSpace:"pre-wrap"}}>{e.text}</div>
        </div>
      ))}
    </div>
  );
}

// ── Equipe ────────────────────────────────────────────────────────────────────
function EquipeTab({team,saveTeam,weights,saveWeights}){
  const [showForm,setShowForm]=useState(false);
  const [name,setName]=useState("");const [role,setRole]=useState("Arquiteto");
  const [editW,setEditW]=useState(false);
  const [wLocal,setWLocal]=useState({...weights});
  function addMember(){if(!name.trim())return;saveTeam([...team,{id:"m"+uid(),name:name.trim(),role}]);setName("");setRole("Arquiteto");setShowForm(false);}
  function delMember(id){saveTeam(team.filter(m=>m.id!==id));}
  return <>
    <div style={{...B.card,marginBottom:"1.5rem",borderTop:`3px solid ${RED}`}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"10px"}}>
        <div>
          <div style={{fontSize:"14px",fontWeight:700,color:"#111"}}>Pesos por cargo</div>
          <div style={{fontSize:"12px",color:"#9ca3af",marginTop:"2px"}}>Contribuição relativa de cada cargo no cálculo de dedicação</div>
        </div>
        <button onClick={()=>{setEditW(v=>!v);setWLocal({...weights});}} style={{...B.sec,fontSize:"12px",padding:"5px 12px"}}>{editW?"Cancelar":"Editar"}</button>
      </div>
      {editW?(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px",marginBottom:"12px"}}>
            {ROLES.map(r=>(
              <div key={r} style={{display:"flex",alignItems:"center",gap:"10px",padding:"8px 10px",borderRadius:"8px",background:"#f9fafb",border:"1px solid #f0f0f0"}}>
                <div style={{width:"8px",height:"8px",borderRadius:"50%",background:ROLE_COLOR[r]||"#888",flexShrink:0}}/>
                <span style={{fontSize:"13px",fontWeight:600,flex:1}}>{r}</span>
                <input type="number" step="0.5" min="0.1" value={wLocal[r]??DW[r]??1} onChange={e=>setWLocal(p=>({...p,[r]:+e.target.value}))} style={{...B.inp,width:"65px",textAlign:"center",fontWeight:700,fontSize:"14px",padding:"4px"}}/>
              </div>
            ))}
          </div>
          <div style={{fontSize:"11px",color:"#9ca3af",marginBottom:"10px",padding:"8px 10px",borderRadius:"6px",background:"#f9fafb"}}>
            Exemplo: Sócio ({wLocal["Sócio"]||3}) contribui {Math.round((wLocal["Sócio"]||3)/(wLocal["Estagiário"]||0.5))}× mais que Estagiário ({wLocal["Estagiário"]||0.5}) no cálculo.
          </div>
          <button onClick={()=>{saveWeights(wLocal);setEditW(false);}} style={B.pri}>Salvar pesos</button>
        </div>
      ):(
        <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
          {ROLES.map(r=>(
            <div key={r} style={{display:"flex",alignItems:"center",gap:"6px",padding:"6px 12px",borderRadius:"20px",background:ROLE_COLOR[r]+"15",border:`1px solid ${ROLE_COLOR[r]}25`}}>
              <span style={{fontSize:"12px",fontWeight:600,color:ROLE_COLOR[r]}}>{r}</span>
              <span style={{fontSize:"14px",fontWeight:800,color:"#111"}}>{weights[r]??DW[r]}</span>
            </div>
          ))}
        </div>
      )}
    </div>
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1.25rem"}}>
      <span style={{fontSize:"13px",color:"#9ca3af"}}>{team.length} pessoa{team.length!==1?"s":""}</span>
      <button onClick={()=>setShowForm(v=>!v)} style={B.pri}>+ Adicionar pessoa</button>
    </div>
    {showForm&&(
      <div style={{background:"#fff",borderRadius:"10px",padding:"16px",marginBottom:"1.25rem",border:`1px solid ${RED}25`}}>
        <div style={{fontSize:"13px",fontWeight:700,color:"#111",marginBottom:"12px"}}>Nova pessoa</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 160px",gap:"8px",marginBottom:"10px"}}>
          <input autoFocus placeholder="Nome completo" value={name} onChange={e=>setName(e.target.value)} onKeyDown={e=>e.key==="Enter"&&addMember()} style={B.inp}/>
          <select value={role} onChange={e=>setRole(e.target.value)} style={B.inp}>{ROLES.map(r=><option key={r} value={r}>{r}</option>)}</select>
        </div>
        <div style={{display:"flex",gap:"8px"}}>
          <button onClick={addMember} style={B.pri}>Adicionar</button>
          <button onClick={()=>setShowForm(false)} style={B.sec}>Cancelar</button>
        </div>
      </div>
    )}
    {ROLES.map(role=>{
      const members=team.filter(m=>m.role===role);if(!members.length)return null;
      const rc=ROLE_COLOR[role]||"#888";
      return(
        <div key={role} style={{marginBottom:"1.5rem"}}>
          <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"10px"}}>
            <div style={{width:"3px",height:"16px",borderRadius:"2px",background:rc}}/>
            <span style={{fontSize:"12px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#374151"}}>{role}</span>
            <span style={{fontSize:"11px",padding:"1px 7px",borderRadius:"10px",background:rc+"15",color:rc,fontWeight:700}}>peso {weights[role]??DW[role]}</span>
          </div>
          {members.map(m=>(
            <div key={m.id} style={{...B.card,display:"flex",alignItems:"center",gap:"12px",padding:"10px 14px"}}>
              <div style={{width:"36px",height:"36px",borderRadius:"50%",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:700,background:rc+"15",color:rc,border:`1.5px solid ${rc}30`}}>{m.name.slice(0,2).toUpperCase()}</div>
              <div style={{flex:1}}><div style={{fontSize:"14px",fontWeight:600,color:"#111"}}>{m.name}</div><div style={{fontSize:"11px",color:"#9ca3af"}}>{m.role}</div></div>
              <button onClick={()=>delMember(m.id)} style={{...B.ghost,fontSize:"18px",lineHeight:1,opacity:.35}}>×</button>
            </div>
          ))}
        </div>
      );
    })}
  </>;
}

// ── Histórico ─────────────────────────────────────────────────────────────────
function HistTab({history,projects,cont,team,colorMap}){
  const [sel,setSel]=useState(null);
  const dates=Object.keys(history).sort((a,b)=>b.localeCompare(a));
  if(!dates.length)return <div style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}>Nenhum histórico salvo ainda.</div>;
  if(sel){
    const rep=calcReport(history[sel]?.alloc||{},team,projects,{});
    return(
      <div>
        <button onClick={()=>setSel(null)} style={{...B.ghost,color:RED,padding:0,fontWeight:600,fontSize:"13px",marginBottom:"1rem"}}>← Voltar</button>
        <div style={{fontSize:"18px",fontWeight:700,color:"#111",marginBottom:"2px"}}>{fmtDate(sel)}</div>
        <div style={{fontSize:"13px",color:"#9ca3af",marginBottom:"1.25rem"}}>{cap(fmtDay(sel))}</div>
        {rep.map(r=>{
          const p=projects.find(x=>x.id===r.pid);const c=cont.find(x=>x.id===p?.cId);
          return(
            <div key={r.pid} style={{marginBottom:"14px"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:"4px"}}>
                <div><span style={{fontSize:"13px",fontWeight:600,color:"#111"}}>{r.name}</span>{c&&<span style={{fontSize:"11px",color:"#9ca3af",marginLeft:"6px"}}>{c.name}</span>}</div>
                <span style={{fontSize:"12px",color:"#9ca3af"}}>{r.pct.toFixed(1)}%</span>
              </div>
              <div style={{height:"6px",background:"#f3f4f6",borderRadius:"3px",overflow:"hidden",marginBottom:"3px"}}>
                <div style={{height:"100%",width:`${r.pct}%`,background:colorMap[r.pid]||"#888",borderRadius:"3px"}}/>
              </div>
              <div style={{fontSize:"11px",color:"#9ca3af"}}>{r.people.join(", ")}</div>
            </div>
          );
        })}
      </div>
    );
  }
  return(
    <div>
      <div style={{fontSize:"13px",color:"#9ca3af",marginBottom:"1rem"}}>{dates.length} dia{dates.length!==1?"s":""} registrado{dates.length!==1?"s":""}</div>
      {dates.map(d=>{
        const a=normAlloc(history[d]?.alloc||{});
        const pids=[...new Set(Object.values(a).flat().filter(Boolean))];
        const cnt=team.filter(m=>(a[m.id]||[]).length>0).length;
        return(
          <div key={d} onClick={()=>setSel(d)} style={{...B.card,display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer"}}>
            <div><div style={{fontSize:"14px",fontWeight:600,color:"#111"}}>{fmtDate(d)}</div><div style={{fontSize:"12px",color:"#9ca3af",marginTop:"1px"}}>{cap(fmtDay(d))}</div></div>
            <div style={{display:"flex",alignItems:"center",gap:"10px"}}>
              <div style={{display:"flex",gap:"3px"}}>{pids.slice(0,7).map(pid=><div key={pid} style={{width:"8px",height:"8px",borderRadius:"50%",background:colorMap[pid]||"#888"}}/>)}</div>
              <div style={{textAlign:"right",fontSize:"12px",color:"#9ca3af",lineHeight:1.6}}>
                <div>{pids.length} proj.</div><div>{cnt} pessoas</div>
              </div>
            </div>
          </div>
        );
      })}
    </div>
  );
}
