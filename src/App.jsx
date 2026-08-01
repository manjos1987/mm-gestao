import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell, PieChart, Pie } from "recharts";
import { initializeApp, getApps, getApp, deleteApp } from "firebase/app";
import { getFirestore, doc, setDoc, onSnapshot } from "firebase/firestore";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged, createUserWithEmailAndPassword, updatePassword, sendPasswordResetEmail, EmailAuthProvider, reauthenticateWithCredential } from "firebase/auth";
import html2canvas from "html2canvas";

const firebaseConfig = {
  apiKey:            "AIzaSyB6EVH7jY2qTMaD5wBUeawyI7eAPAf_PeI",
  authDomain:        "mm-gestao1987.firebaseapp.com",
  projectId:         "mm-gestao1987",
  storageBucket:     "mm-gestao1987.firebasestorage.app",
  messagingSenderId: "660580949425",
  appId:             "1:660580949425:web:556bc217a483733092b190",
};
const ANTHROPIC_KEY = "";

const TODAY   = new Date().toISOString().split("T")[0];
const DAYS_PT = ["domingo","segunda-feira","terca-feira","quarta-feira","quinta-feira","sexta-feira","sabado"];
const RED     = "#C8102E";
const ADMIN_EMAILS=["marcelloanjos.arq@gmail.com","marcello@mmarquitetura.com"];
const ADMIN_NAMES={"marcelloanjos.arq@gmail.com":"Marcello","marcello@mmarquitetura.com":"Marcello"};
const COLORS  = ["#C8102E","#2563EB","#059669","#D97706","#7C3AED","#0891B2","#B45309","#0D9488","#4F46E5","#065F46","#1D4ED8","#6D28D9","#BE185D","#92400E","#0E7490","#9D174D"];
const ROLES   = ["Socio","Coordenador","Arquiteto","Estagiario","Parceiro"];
const ROLES_DISPLAY = {"Socio":"Sócio","Coordenador":"Coordenador","Arquiteto":"Arquiteto","Estagiario":"Estagiário","Parceiro":"Parceiro"};
const SERVICES= ["Projeto arquitetonico","Projeto legal","Projeto executivo","Imagens / Renderizacao","Maquete","Interiores","Urbanismo","Estudo de Viabilidade"];
const SERVICES_DISPLAY = {
  "Projeto arquitetonico":"Projeto arquitetônico",
  "Projeto legal":"Projeto legal",
  "Projeto executivo":"Projeto executivo",
  "Imagens / Renderizacao":"Imagens / Renderização",
  "Maquete":"Maquete",
  "Interiores":"Interiores",
  "Urbanismo":"Urbanismo",
  "Estudo de Viabilidade":"Estudo de Viabilidade"
};
const DW = {Socio:3,Coordenador:2.5,Arquiteto:2,Estagiario:1,Parceiro:1.5};
const ST_CYC = {ativo:"aguardando",aguardando:"bloqueado",bloqueado:"ativo"};
const STATUS = {
  ativo:     {label:"Ativo",      bg:"#dcfce7",color:"#15803d",hex:"#16a34a"},
  aguardando:{label:"Aguardando", bg:"#fef9c3",color:"#a16207",hex:"#d97706"},
  bloqueado: {label:"Bloqueado",  bg:"#fee2e2",color:"#dc2626",hex:"#dc2626"},
  concluido: {label:"Concluido",  bg:"#f3f4f6",color:"#6b7280",hex:"#9ca3af"},
};
const ROLE_COLOR={Socio:"#111",Coordenador:"#7C3AED",Arquiteto:"#2563EB",Estagiario:"#059669",Parceiro:"#D97706"};
const UFS=["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"];
const SHIFTS={falta:"Falta",manha:"Manhã",tarde:"Tarde",integral:"Integral"};
const SHIFT_COLORS={falta:"#dc2626",manha:"#2563EB",tarde:"#D97706",integral:"#7C3AED"};
const SHIFT_BG={falta:"#fee2e2",manha:"#eff6ff",tarde:"#fffbeb",integral:"#f5f3ff"};
const MONTH_PT=["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

const INIT_C=[
  {id:"c1",name:"Canaa",contact:""},{id:"c2",name:"Santorini",contact:""},
  {id:"c3",name:"Ilha de Creta",contact:""},{id:"c4",name:"Cristalina",contact:""},
  {id:"c5",name:"Julio Cezar",contact:""},{id:"c6",name:"Executivo",contact:""},
  {id:"c7",name:"Astro",contact:""},{id:"c8",name:"Luciano",contact:""},
  {id:"c9",name:"Terra do Meio",contact:""},{id:"c10",name:"Capitolio",contact:""},
  {id:"c11",name:"Hado",contact:""},
];
const INIT_P=[
  {id:"p1", cId:"c1", name:"Condominio de Canaa",  svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p2", cId:"c2", name:"Santorini",             svcs:[], city:"",st:"",country:"Brasil",status:"aguardando",note:"Aguardando cliente"},
  {id:"p3", cId:"c3", name:"Ilha de Creta",         svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p4", cId:"c4", name:"Casas Cristalina",      svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p5", cId:"c1", name:"Casas Canaa",           svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p6", cId:"c5", name:"Kitnet Julio Cezar",    svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:""},
  {id:"p7", cId:"c6", name:"Executivo Towers",      svcs:[], city:"",st:"",country:"Brasil",status:"bloqueado", note:"Aguardando Haroldo / Guilherme"},
  {id:"p8", cId:"c6", name:"Executivo Mall",        svcs:[], city:"",st:"",country:"Brasil",status:"bloqueado", note:"Calcada errada - sem atualizacao"},
  {id:"p9", cId:"c7", name:"Apto Astro",            svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Verificar docs do Heli"},
  {id:"p10",cId:"c8", name:"Casas Luciano",         svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Seguir com projeto legal"},
  {id:"p11",cId:"c9", name:"Terra do Meio",         svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Portico pendente"},
  {id:"p12",cId:"c10",name:"Capitolio",             svcs:[], city:"",st:"",country:"Brasil",status:"bloqueado", note:"Aguardando Manoel"},
  {id:"p13",cId:"c11",name:"Hado San",              svcs:[], city:"",st:"",country:"Brasil",status:"aguardando",note:"Protocolado na prefeitura"},
  {id:"p14",cId:"c1", name:"Posto Canaa",           svcs:[], city:"",st:"",country:"Brasil",status:"ativo",     note:"Atualizar planilha"},
];
const INIT_T=[
  {id:"m1",name:"Alice",   role:"Arquiteto"},{id:"m2",name:"Eloisa",  role:"Arquiteto"},
  {id:"m3",name:"Gustavo", role:"Arquiteto"},{id:"m4",name:"Leticia", role:"Estagiario"},
  {id:"m5",name:"Marcello",role:"Parceiro"}, {id:"m6",name:"Murilo",  role:"Arquiteto"},
];

function fmtDate(d){ var p=d.split("-"); return p[2]+"/"+p[1]+"/"+p[0]; }
function fmtDay(d){ return DAYS_PT[new Date(d+"T12:00:00").getDay()]; }
function cap(s){ return s.charAt(0).toUpperCase()+s.slice(1); }
function uid(){ return Date.now().toString(36)+Math.random().toString(36).slice(2,6); }
function wOf(role,w){ return +(w[role]!=null?w[role]:(DW[role]!=null?DW[role]:1)); }

function normAlloc(a){
  var r={};
  Object.entries(a||{}).forEach(function(e){
    var mid=e[0],v=e[1];
    r[mid]=Array.isArray(v)?v.filter(Boolean):typeof v==="string"&&v?[v]:[];
  });
  return r;
}

function calcReport(alloc,team,projects,weights){
  if(!weights)weights={};
  var n=normAlloc(alloc),bp={};
  team.forEach(function(m){
    var pids=n[m.id]||[];
    if(!pids.length)return;
    var w=wOf(m.role,weights);
    pids.forEach(function(pid){
      if(!bp[pid])bp[pid]={wpts:0,people:[],detail:[]};
      bp[pid].wpts+=w;
      bp[pid].people.push(m.name);
      bp[pid].detail.push({name:m.name,role:m.role,w:+(w.toFixed(2))});
    });
  });
  var tot=Object.values(bp).reduce(function(s,v){return s+v.wpts;},0);
  return Object.entries(bp).map(function(e){
    var pid=e[0],val=e[1];
    var p=projects.find(function(x){return x.id===pid;});
    return {pid:pid,name:p?p.name:pid,people:val.people,detail:val.detail,wpts:+(val.wpts.toFixed(2)),pct:tot?val.wpts/tot*100:0};
  }).sort(function(a,b){return b.wpts-a.wpts;});
}

function calcTeamLoad(alloc,team,weights){
  if(!weights)weights={};
  var n=normAlloc(alloc);
  var totW=team.reduce(function(s,m){var pids=n[m.id]||[];return pids.length?s+wOf(m.role,weights):s;},0);
  return team.map(function(m){
    var pids=n[m.id]||[],w=wOf(m.role,weights);
    return Object.assign({},m,{weight:w,pids:pids,pct:totW&&pids.length?w/totW*100:0,active:pids.length>0});
  });
}

function exportExcel(report,date,team,normA,projects,cont,weights,periodLabel){
  var wb=XLSX.utils.book_new();
  var pesoStr=Object.entries(weights).map(function(e){return e[0]+"="+e[1];}).join(", ");
  var dateLabel=periodLabel||(fmtDate(date)+", "+cap(fmtDay(date)));
  var s1=[
    ["M+M Arquitetura - Relatorio de Dedicacao"],
    ["Período:",dateLabel],
    [],
    ["Projeto","Contratante","Colaboradores","Pts Pond.","%"]
  ];
  report.forEach(function(r){
    var p=projects.find(function(x){return x.id===r.pid;});
    var c=cont.find(function(x){return x.id===(p?p.cId:"");});
    s1.push([r.name,c?c.name:"",r.people.join(", "),r.wpts,+(r.pct.toFixed(1))]);
  });
  s1.push([]);
  s1.push(["Pesos: "+pesoStr]);
  var ws1=XLSX.utils.aoa_to_sheet(s1);
  ws1["!cols"]=[{wch:30},{wch:20},{wch:40},{wch:12},{wch:8}];
  XLSX.utils.book_append_sheet(wb,ws1,"Dedicacao");
  var s2=[["Alocacao por Colaborador"],["Data:",fmtDate(date)],[],["Colaborador","Cargo","Peso","Projetos"]];
  team.forEach(function(m){
    var pids=normA[m.id]||[];
    var names=pids.map(function(pid){var p=projects.find(function(x){return x.id===pid;});return p?p.name:pid;});
    s2.push([m.name,m.role,wOf(m.role,weights),names.join("; ")]);
  });
  var ws2=XLSX.utils.aoa_to_sheet(s2);
  ws2["!cols"]=[{wch:20},{wch:15},{wch:8},{wch:70}];
  XLSX.utils.book_append_sheet(wb,ws2,"Colaboradores");
  XLSX.writeFile(wb,"MM_Relatorio_"+date+".xlsx");
}

function exportPDF(report,date,projects,cont,normA,team,weights,periodLabel){
  var cidx={};
  projects.forEach(function(p,i){cidx[p.id]=i;});
  var projRows=report.map(function(r){
    var p=projects.find(function(x){return x.id===r.pid;});
    var c=cont.find(function(x){return x.id===(p?p.cId:"");});
    var col=COLORS[(cidx[r.pid]||0)%COLORS.length];
    var cname=c?c.name:"";
    var pct=r.pct.toFixed(1);
    var ppl=r.people.join(", ")+" - "+r.wpts+" pts";
    return '<div class="proj"><div class="prow"><span><span class="pname">'+r.name+'</span><span class="pcli">'+cname+'</span></span><span class="ppct">'+pct+'%</span></div><div class="bar"><div class="fill" style="width:'+pct+'%;background:'+col+'"></div></div><div class="ppl">'+ppl+'</div></div>';
  }).join("");
  var teamRows=team.map(function(m){
    var pids=normA[m.id]||[];
    var w=wOf(m.role,weights);
    var names=pids.map(function(pid){var p=projects.find(function(x){return x.id===pid;});return p?p.name:pid;});
    return '<div class="arow"><span class="an">'+m.name+'</span><span class="ar">'+m.role+' ('+w+')</span><span class="ap">'+(names.length?names.join(", "):"&mdash;")+'</span></div>';
  }).join("");
  var totalPts=report.reduce(function(s,r){return s+r.wpts;},0).toFixed(1);
  var peopleCnt=team.filter(function(m){return (normA[m.id]||[]).length>0;}).length;
  var pesoStr=Object.values(weights).join("/");
  var dateLabel=periodLabel||fmtDate(date);
  var dayLabel=periodLabel?'':cap(fmtDay(date));
  var html='<!DOCTYPE html><html><head><meta charset="utf-8"><title>M+M '+dateLabel+'</title>'
    +'<style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:Helvetica Neue,Arial,sans-serif;color:#111;-webkit-print-color-adjust:exact;print-color-adjust:exact}.page{padding:44px 48px;max-width:720px;margin:0 auto}.stripe{height:4px;background:#C8102E}.hdr{display:flex;justify-content:space-between;align-items:flex-end;padding:24px 0 16px;border-bottom:2px solid #111;margin-bottom:22px}.logo{font-size:28px;font-weight:800;letter-spacing:-1px}.plus{color:#C8102E}.sub{font-size:11px;color:#aaa;letter-spacing:.07em;text-transform:uppercase;margin-top:3px}.meta{text-align:right;font-size:12px;color:#888;line-height:1.7}.mdate{font-size:15px;font-weight:700;color:#111;display:block}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:22px}.stat{padding:12px;border:1px solid #f0f0f0;border-radius:6px;text-align:center}.sv{font-size:22px;font-weight:700}.sl{font-size:9px;text-transform:uppercase;letter-spacing:.06em;color:#bbb;margin-top:3px}.sec{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:#bbb;font-weight:700;margin-bottom:12px}.proj{margin-bottom:14px}.prow{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px}.pname{font-size:13px;font-weight:600}.pcli{font-size:10px;color:#aaa;margin-left:6px}.ppct{font-size:12px;color:#888;font-weight:600}.bar{height:7px;background:#f0f0f0;border-radius:4px;overflow:hidden;margin-bottom:3px}.fill{height:100%;border-radius:4px}.ppl{font-size:10px;color:#bbb}.div{border:none;border-top:1px solid #f0f0f0;margin:18px 0}.arow{display:flex;gap:10px;padding:5px 0;border-bottom:1px solid #f8f8f8;font-size:11px}.an{font-weight:600;min-width:80px;color:#555}.ar{color:#aaa;min-width:80px}.ap{color:#777}.foot{margin-top:22px;text-align:center;font-size:9px;color:#ddd;padding-top:10px;border-top:1px solid #f0f0f0}@media print{@page{margin:1cm}}</style>'
    +'</head><body><div class="stripe"></div><div class="page">'
    +'<div class="hdr"><div><div class="logo">M<span class="plus">+</span>M</div><div class="sub">Arquitetura - Relatorio de Dedicacao</div></div><div class="meta"><span class="mdate">'+dateLabel+'</span>'+dayLabel+'</div></div>'
    +'<div class="stats"><div class="stat"><div class="sv">'+report.length+'</div><div class="sl">Projetos</div></div><div class="stat"><div class="sv">'+peopleCnt+'</div><div class="sl">Pessoas</div></div><div class="stat"><div class="sv">'+totalPts+'</div><div class="sl">Pts pond.</div></div><div class="stat"><div class="sv">'+pesoStr+'</div><div class="sl">Pesos</div></div></div>'
    +'<div class="sec">Dedicacao por projeto</div>'+projRows
    +'<div class="div"></div><div class="sec">Por colaborador</div>'+teamRows
    +'<div class="foot">Gerado em '+new Date().toLocaleString("pt-BR")+' - M+M Arquitetura</div>'
    +'</div><script>setTimeout(function(){window.print();},400);<\/script></body></html>';
  var w=window.open("","_blank","width=820,height=720");
  if(w){w.document.write(html);w.document.close();}
}

var B={
  inp:{fontSize:"13px",padding:"8px 10px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:"#111",width:"100%",boxSizing:"border-box",outline:"none"},
  card:{background:"#fff",border:"1px solid #f0f0f0",borderRadius:"8px",padding:"12px 16px",marginBottom:"8px",boxShadow:"0 1px 2px rgba(0,0,0,.03)"},
  pri:{background:RED,color:"#fff",border:"none",borderRadius:"6px",padding:"8px 18px",fontSize:"13px",fontWeight:600,cursor:"pointer"},
  sec:{background:"transparent",color:"#6b7280",border:"1px solid #e5e7eb",borderRadius:"6px",padding:"7px 14px",fontSize:"13px",cursor:"pointer"},
  ghost:{background:"transparent",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:"12px",padding:"2px 6px"},
  lbl:{fontSize:"10px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".08em",marginBottom:"10px"},
};

function LoginScreen(){
  var es=useState(""); var email=es[0],setEmail=es[1];
  var ps=useState(""); var pass=ps[0],setPass=ps[1];
  var errs=useState(""); var err=errs[0],setErr=errs[1];
  var lds=useState(false); var load=lds[0],setLoad=lds[1];

  async function handleLogin(e){
    e.preventDefault();
    if(!email.trim()||!pass)return;
    setLoad(true);setErr("");
    try{
      var app=getApps().length?getApp():initializeApp(firebaseConfig);
      var auth=getAuth(app);
      await signInWithEmailAndPassword(auth,email.trim(),pass);
    }catch(ex){
      var msgs={
        "auth/invalid-credential":"E-mail ou senha incorretos.",
        "auth/user-not-found":"E-mail não encontrado.",
        "auth/wrong-password":"Senha incorreta.",
        "auth/too-many-requests":"Muitas tentativas. Aguarde alguns minutos.",
        "auth/invalid-email":"E-mail inválido.",
      };
      setErr(msgs[ex.code]||"Erro ao entrar. Tente novamente.");
    }
    setLoad(false);
  }

  return (
    <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",alignItems:"center",justifyContent:"center",padding:"1.5rem"}}>
      <div style={{background:"#fff",borderRadius:"16px",padding:"2.5rem",maxWidth:"380px",width:"100%",boxShadow:"0 4px 24px rgba(0,0,0,.07)",border:"1px solid #f0f0f0"}}>
        <div style={{textAlign:"center",marginBottom:"2rem"}}>
          <div style={{fontSize:"32px",fontWeight:800,letterSpacing:"-1px",color:"#111",marginBottom:"4px"}}>M<span style={{color:RED}}>+</span>M</div>
          <div style={{fontSize:"11px",color:"#9ca3af",letterSpacing:".08em",textTransform:"uppercase",fontWeight:500}}>Arquitetura — Gestão</div>
        </div>
        <form onSubmit={handleLogin}>
          <div style={{marginBottom:"12px"}}>
            <div style={{fontSize:"11px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".07em",marginBottom:"6px"}}>E-mail</div>
            <input type="email" value={email} onChange={function(e){setEmail(e.target.value);}} placeholder="seu@email.com" autoFocus style={Object.assign({},B.inp,{fontSize:"14px",padding:"10px 12px"})} required/>
          </div>
          <div style={{marginBottom:"20px"}}>
            <div style={{fontSize:"11px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".07em",marginBottom:"6px"}}>Senha</div>
            <input type="password" value={pass} onChange={function(e){setPass(e.target.value);}} placeholder="••••••••" style={Object.assign({},B.inp,{fontSize:"14px",padding:"10px 12px"})} required/>
          </div>
          {err&&(
            <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:"8px",padding:"10px 12px",fontSize:"13px",color:"#dc2626",marginBottom:"14px"}}>
              {err}
            </div>
          )}
          <button type="submit" disabled={load} style={Object.assign({},B.pri,{width:"100%",padding:"11px",fontSize:"14px",opacity:load?.6:1})}>
            {load?"Entrando...":"Entrar"}
          </button>
        </form>
        <div style={{textAlign:"center",marginTop:"1.5rem",fontSize:"11px",color:"#d1d5db"}}>
          Acesso restrito — M+M Arquitetura
        </div>
      </div>
    </div>
  );
}

export default function App(){
  var dbRef=useRef(null);
  var ts=useState("visao"); var tab=ts[0],setTab=ts[1];
  var cs=useState(INIT_C); var cont=cs[0],setCont=cs[1];
  var ps=useState(INIT_P); var projects=ps[0],setProjects=ps[1];
  var tms=useState(INIT_T); var team=tms[0],setTeam=tms[1];
  var ws=useState(DW); var weights=ws[0],setWeights=ws[1];
  var ds=useState({}); var diary=ds[0],setDiaryState=ds[1];
  var dts=useState(TODAY); var date=dts[0],setDate=dts[1];
  var als=useState({}); var alloc=als[0],setAlloc=als[1];
  var hs=useState({}); var history=hs[0],setHistory=hs[1];
  var aus=useState(null); var authUser=aus[0],setAuthUser=aus[1];
  var als3=useState(true); var authLoading=als3[0],setAuthLoading=als3[1];
  var ls=useState(true); var loading=ls[0],setLoading=ls[1];
  var ss=useState(false); var syncing=ss[0],setSyncing=ss[1];
  var fs=useState(""); var flash=fs[0],setFlash=fs[1];
  var pls=useState({}); var planning=pls[0],setPlanning=pls[1];
  var sis=useState({}); var allocSubitems=sis[0],setAllocSubitems=sis[1];
  var ams=useState(""); var aiMsg=ams[0],setAiMsg=ams[1];
  var als2=useState(false); var aiLoad=als2[0],setAiLoad=als2[1];
  var os=useState(true); var online=os[0],setOnline=os[1];
  var tfs=useState([]); var tarefas=tfs[0],setTarefas=tfs[1];

  useEffect(function(){
    if(firebaseConfig.apiKey==="COLE_AQUI"){setLoading(false);setAuthLoading(false);return;}
    try{
      var app=getApps().length?getApp():initializeApp(firebaseConfig);
      var db=getFirestore(app);
      var auth=getAuth(app);
      dbRef.current=db;
      var unsub1=null,unsub2=null;
      var unsubAuth=onAuthStateChanged(auth,function(user){
        setAuthUser(user);
        setAuthLoading(false);
        // Cancela listeners anteriores se existirem
        if(unsub1){unsub1();unsub1=null;}
        if(unsub2){unsub2();unsub2=null;}
        // Só abre conexão com o Firestore depois do login confirmado
        if(!user){setLoading(false);return;}
        var cfgRef=doc(db,"mm_app","config");
        var histRef=doc(db,"mm_app","history");
        unsub1=onSnapshot(cfgRef,function(snap){
          if(!snap.exists()){
            setDoc(cfgRef,{cont:INIT_C,projects:INIT_P,team:INIT_T,weights:DW,diary:{}});
          } else {
            var d=snap.data();
            setCont(d.cont||INIT_C);
            setProjects(d.projects||INIT_P);
            setTeam(d.team||INIT_T);
            setWeights(d.weights||DW);
            setDiaryState(d.diary||{});
            setPlanning(d.planning||{});
            setTarefas(d.tarefas||[]);
          }
        },function(err){console.error(err);setOnline(false);});
        unsub2=onSnapshot(histRef,function(snap){
          var h=snap.exists()?snap.data():{};
          setHistory(h);
          if(h[TODAY]){setAlloc(normAlloc(h[TODAY].alloc||{}));setAllocSubitems(h[TODAY].allocSubitems||{});}
          setLoading(false);
        },function(err){console.error(err);setLoading(false);});
      });
      return function(){unsubAuth();if(unsub1)unsub1();if(unsub2)unsub2();};
    }catch(err){console.error(err);setLoading(false);setAuthLoading(false);setOnline(false);}
  },[]);

  useEffect(function(){
    setAlloc(normAlloc(history[date]?history[date].alloc||{}:{}));
    setAllocSubitems(history[date]?history[date].allocSubitems||{}:{});
  },[date]);

  useEffect(function(){
    if(date!==TODAY)return;
    setAlloc(function(prev){
      var a=Object.assign({},prev);var ch=false;
      tarefas.filter(function(t){return t.status!=="concluido";}).forEach(function(t){
        if(!a[t.membroId])a[t.membroId]=[];
        if(a[t.membroId].indexOf(t.projetoId)===-1){a[t.membroId]=[...a[t.membroId],t.projetoId];ch=true;}
      });
      return ch?a:prev;
    });
    setAllocSubitems(function(prev){
      var sis=Object.assign({},prev);var ch=false;
      tarefas.filter(function(t){return t.status!=="concluido"&&t.atividadeId;}).forEach(function(t){
        if(!sis[t.membroId])sis[t.membroId]={};
        var ps=(sis[t.membroId][t.projetoId]||[]);
        if(ps.indexOf(t.atividadeId)===-1){sis[t.membroId][t.projetoId]=[...ps,t.atividadeId];ch=true;}
      });
      return ch?sis:prev;
    });
  },[tarefas]);

  var colorMap={};
  projects.forEach(function(p,i){colorMap[p.id]=COLORS[i%COLORS.length];});
  var normA=normAlloc(alloc);
  var report=calcReport(alloc,team,projects,weights);
  var teamLoad=calcTeamLoad(alloc,team,weights);
  var savedDate=!!(history[date]&&history[date].savedAt);
  var assignedCnt=team.filter(function(m){return (normA[m.id]||[]).length>0;}).length;

  async function persist(key,value){
    if(!dbRef.current)return;
    setSyncing(true);
    try{
      if(key==="history"){
        await setDoc(doc(dbRef.current,"mm_app","history"),value,{merge:true});
      } else {
        await setDoc(doc(dbRef.current,"mm_app","config"),{[key]:value},{merge:true});
      }
    }catch(e){console.error(e);}
    setSyncing(false);
  }

  function persistAlloc(a,subs){
    var snap=tarefas.filter(function(t){return t.status!=="pendente"&&(t.status!=="concluido"||(t.concluidoEm&&t.concluidoEm.startsWith(date)));}).map(function(t){return {id:t.id,membroNome:t.membroNome,projetoId:t.projetoId||"",projetoNome:t.projetoNome,atividadeNome:t.atividadeNome||"",descricao:t.descricao,status:t.status,iniciadoEm:t.iniciadoEm||null,concluidoEm:t.concluidoEm||null,delegadoHistorico:t.delegadoHistorico||[],atribuidoPor:t.atribuidoPor||""};});
    var entry=Object.assign({},history[date]||{},{alloc:a,allocSubitems:subs,savedAt:new Date().toISOString(),tarefasSnap:snap});
    return persist("history",{[date]:entry});
  }

  function toggleProj(mid,pid){
    var curr=normA[mid]||[];
    var next=curr.includes(pid)?curr.filter(function(p){return p!==pid;}):[...curr,pid];
    var nextAlloc=Object.assign({},normA,{[mid]:next});
    setAlloc(nextAlloc);
    persistAlloc(nextAlloc,allocSubitems);
  }

  function updateAllocSubitems(next){
    setAllocSubitems(next);
    persistAlloc(normA,next);
  }

  async function saveAlloc(){
    await persistAlloc(normA,allocSubitems);
    setFlash("Salvo!");setTimeout(function(){setFlash("");},1800);
  }
  function saveCont(c){setCont(c);persist("cont",c);}
  function saveProjects(p){setProjects(p);persist("projects",p);}
  function saveTeam(t){setTeam(t);persist("team",t);}
  function saveWeights(w){setWeights(w);persist("weights",w);}
  function saveDiary(d){setDiaryState(d);persist("diary",d);}
  function savePlanning(p){setPlanning(p);persist("planning",p);}
  function saveTarefas(t){
    setTarefas(t);
    persist("tarefas",t);
    var snap=t.filter(function(x){return x.status!=="pendente"&&(x.status!=="concluido"||(x.concluidoEm&&x.concluidoEm.startsWith(TODAY)));}).map(function(x){return {id:x.id,membroNome:x.membroNome,projetoId:x.projetoId||"",projetoNome:x.projetoNome,atividadeNome:x.atividadeNome||"",descricao:x.descricao,status:x.status,iniciadoEm:x.iniciadoEm||null,concluidoEm:x.concluidoEm||null,delegadoHistorico:x.delegadoHistorico||[],atribuidoPor:x.atribuidoPor||""};});
    persist("history",{[TODAY]:Object.assign({},history[TODAY]||{},{tarefasSnap:snap})});
  }

  async function genMsg(){
    if(!report.length)return;
    var key=ANTHROPIC_KEY||prompt("Cole sua chave da API Anthropic (sk-ant-...):");
    if(!key)return;
    setAiLoad(true);setAiMsg("");
    try{
      var lines=report.map(function(r){
        var p=projects.find(function(x){return x.id===r.pid;});
        var c=cont.find(function(x){return x.id===(p?p.cId:"");});
        return "- "+r.name+(c?" ("+c.name+")":"")+": "+r.people.join(", ")+" - "+r.pct.toFixed(0)+"%";
      }).join("\n");
      var blocked=projects.filter(function(p){return ["bloqueado","aguardando"].includes(p.status)&&p.note;}).map(function(p){return "- "+p.name+": "+p.note;}).join("\n");
      var prompt_text="Mensagem curta para o Manoel (socio M+M) sobre alocacao de "+fmtDate(date)+".\n\nDedicacao:\n"+lines+"\n\nBloqueados:\n"+(blocked||"Nenhum")+"\n\nTom WhatsApp, informal mas profissional. Maximo 4 linhas. Sem bullets.";
      var res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",
        headers:{"Content-Type":"application/json","x-api-key":key,"anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:300,messages:[{role:"user",content:prompt_text}]})
      });
      var data=await res.json();
      var txt=data.content&&data.content.find(function(c){return c.type==="text";});
      setAiMsg(txt?txt.text:"Erro.");
    }catch(e){setAiMsg("Erro: "+e.message);}
    setAiLoad(false);
  }

  function handleLogout(){
    var auth=getAuth();
    signOut(auth);
  }

  if(authLoading){
    return (
      <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:"24px",fontWeight:800,marginBottom:"16px"}}>M<span style={{color:RED}}>+</span>M</div>
          <div style={{fontSize:"13px",color:"#9ca3af"}}>Verificando acesso...</div>
        </div>
      </div>
    );
  }

  if(!authUser){
    return <LoginScreen/>;
  }

  if(!loading&&firebaseConfig.apiKey==="COLE_AQUI"){
    return (
      <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",alignItems:"center",justifyContent:"center",padding:"2rem"}}>
        <div style={{background:"#fff",borderRadius:"16px",padding:"2.5rem",maxWidth:"480px",width:"100%",textAlign:"center"}}>
          <div style={{fontSize:"28px",fontWeight:800,marginBottom:"8px"}}>M<span style={{color:RED}}>+</span>M</div>
          <div style={{fontSize:"14px",color:"#6b7280",marginBottom:"2rem"}}>Configure o Firebase para ativar a sincronizacao em tempo real</div>
          <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:"8px",padding:"16px",textAlign:"left",fontSize:"13px",color:"#dc2626",lineHeight:1.7}}>
            Abra o arquivo App.jsx e substitua os valores de firebaseConfig pelos dados do seu projeto Firebase.
          </div>
        </div>
      </div>
    );
  }

  if(loading){
    return (
      <div style={{minHeight:"100vh",background:"#f9fafb",display:"flex",alignItems:"center",justifyContent:"center"}}>
        <div style={{textAlign:"center"}}>
          <div style={{fontSize:"24px",fontWeight:800,marginBottom:"16px"}}>M<span style={{color:RED}}>+</span>M</div>
          <div style={{fontSize:"13px",color:"#9ca3af"}}>Conectando ao banco de dados...</div>
        </div>
      </div>
    );
  }

  var memberByEmail=team.find(function(m){return m.email&&m.email.toLowerCase()===(authUser?authUser.email.toLowerCase():"");});
  var memberAcesso=memberByEmail?memberByEmail.acesso||"usuario":"usuario";
  var isAdmin=!!(authUser&&(ADMIN_EMAILS.includes(authUser.email.toLowerCase())||memberAcesso==="admin"));
  var isGestor=!!(authUser&&!isAdmin&&memberAcesso==="gestor");
  var currentMember=!isAdmin?memberByEmail:null;
  if(!isAdmin&&isGestor){
    return <GestorApp authUser={authUser} handleLogout={handleLogout} tarefas={tarefas} saveTarefas={saveTarefas} currentMember={memberByEmail} team={team} projects={projects}/>;
  }
  if(!isAdmin){
    return <UsuarioApp authUser={authUser} handleLogout={handleLogout} tarefas={tarefas} currentMember={currentMember} saveTarefas={saveTarefas} team={team}/>;
  }

  var TABS=[["visao","Visao Geral"],["hoje","Hoje"],["relatorio","Relatorio"],["projetos","Projetos"],["equipe","Equipe"],["historico","Historico"],["planejamento","Planejamento"],["tarefas","Tarefas"]];
  var TABS_DISPLAY={"visao":"Visão Geral","hoje":"Hoje","relatorio":"Relatório","projetos":"Projetos","equipe":"Equipe","historico":"Histórico","planejamento":"Planejamento","tarefas":"Tarefas"};

  return (
    <div style={{background:"#f9fafb",minHeight:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>
      <div style={{height:"3px",background:RED}}/>
      <div style={{background:"#fff",borderBottom:"1px solid #f0f0f0",padding:"0 1.25rem",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0 0"}}>
          <div style={{display:"flex",alignItems:"baseline",gap:"10px"}}>
            <span style={{fontSize:"22px",fontWeight:800,letterSpacing:"-1px",color:"#111"}}>M<span style={{color:RED}}>+</span>M</span>
            <span style={{fontSize:"11px",color:"#9ca3af",letterSpacing:".05em",textTransform:"uppercase",fontWeight:500}}>Arquitetura</span>
          </div>
          <div style={{display:"flex",gap:"12px",alignItems:"center"}}>
            <div style={{display:"flex",alignItems:"center",gap:"5px",fontSize:"11px"}}>
              <div style={{width:"7px",height:"7px",borderRadius:"50%",background:online?(syncing?"#f59e0b":"#16a34a"):"#dc2626"}}/>
              <span style={{color:"#9ca3af"}}>{syncing?"Salvando...":online?"Sincronizado":"Offline"}</span>
            </div>
            <div style={{display:"flex",gap:"12px",fontSize:"12px",color:"#9ca3af"}}>
              <span><b style={{color:"#111"}}>{projects.filter(function(p){return p.status!=="concluido";}).length}</b> projetos</span>
              <span><b style={{color:"#111"}}>{team.length}</b> pessoas</span>
            </div>
            {authUser&&(
              <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                <span style={{fontSize:"11px",color:"#9ca3af",maxWidth:"120px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.email}</span>
                <button onClick={handleLogout} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",cursor:"pointer"}}>Sair</button>
              </div>
            )}
          </div>
        </div>
        <div style={{display:"flex",overflowX:"auto",marginTop:"4px"}}>
          {TABS.map(function(t){
            var id=t[0];
            return (
              <button key={id} onClick={function(){setTab(id);}} style={{background:"transparent",border:"none",cursor:"pointer",padding:"7px 12px",fontSize:"12px",fontWeight:tab===id?700:400,color:tab===id?"#111":"#9ca3af",borderBottom:tab===id?"2.5px solid "+RED:"2.5px solid transparent",whiteSpace:"nowrap",flexShrink:0}}>
                {TABS_DISPLAY[id]}
              </button>
            );
          })}
        </div>
      </div>
      <div style={{padding:"1.25rem",maxWidth:"800px",margin:"0 auto"}}>
        {tab==="visao"    && <DashTab report={report} teamLoad={teamLoad} projects={projects} cont={cont} team={team} history={history} colorMap={colorMap} weights={weights} tarefas={tarefas}/>}
        {tab==="hoje"     && <HojeTab team={team} projects={projects} cont={cont} normA={normA} colorMap={colorMap} toggleProj={toggleProj} date={date} setDate={setDate} savedDate={savedDate} assignedCnt={assignedCnt} saveAlloc={saveAlloc} flash={flash} report={report} allocSubitems={allocSubitems} setAllocSubitems={updateAllocSubitems}/>}
        {tab==="relatorio"&& <RelTab report={report} date={date} setDate={setDate} projects={projects} cont={cont} normA={normA} team={team} weights={weights} assignedCnt={assignedCnt} colorMap={colorMap} genMsg={genMsg} aiLoad={aiLoad} aiMsg={aiMsg} history={history} exportExcel={exportExcel} exportPDF={exportPDF}/>}
        {tab==="projetos" && <ProjTab projects={projects} cont={cont} saveProjects={saveProjects} saveCont={saveCont} colorMap={colorMap} diary={diary} saveDiary={saveDiary} team={team} history={history}/>}
        {tab==="equipe"   && <EquipeTab team={team} saveTeam={saveTeam} weights={weights} saveWeights={saveWeights}/>}
        {tab==="historico"&& <HistTab history={history} projects={projects} cont={cont} team={team} colorMap={colorMap} saveHistory={function(h){setHistory(h);persist("history",h);}} tarefas={tarefas}/>}
        {tab==="planejamento"&& <PlanTab team={team} planning={planning} savePlanning={savePlanning}/>}
        {tab==="tarefas"    && <TarefasTab tarefas={tarefas} saveTarefas={saveTarefas} team={team} projects={projects} authUser={authUser} currentMember={memberByEmail}/>}
      </div>
      <div style={{textAlign:"center",padding:"2rem 1rem 1.5rem",fontSize:"11px",color:"#9ca3af"}}>
        Criado por MAR Consultoria 2026 - v2.1
      </div>
    </div>
  );
}

function DashTab(props){
  var report=props.report,teamLoad=props.teamLoad,projects=props.projects,cont=props.cont,team=props.team,history=props.history,colorMap=props.colorMap,weights=props.weights,tarefas=props.tarefas||[];
  var active=projects.filter(function(p){return p.status==="ativo";}).length;
  var blocked=projects.filter(function(p){return p.status==="bloqueado";}).length;
  var waiting=projects.filter(function(p){return p.status==="aguardando";}).length;
  var done=projects.filter(function(p){return p.status==="concluido";}).length;
  var chartProj=report.map(function(r){return {name:r.name.length>22?r.name.slice(0,20)+"...":r.name,pct:+(r.pct.toFixed(1)),pid:r.pid,wpts:r.wpts};});
  var chartTeam=teamLoad.filter(function(m){return m.active;}).sort(function(a,b){return b.pct-a.pct;});
  var statusPie=[
    {name:"Ativos",value:active,fill:STATUS.ativo.hex},
    {name:"Bloqueados",value:blocked,fill:STATUS.bloqueado.hex},
    {name:"Aguardando",value:waiting,fill:STATUS.aguardando.hex},
    {name:"Concluidos",value:done,fill:STATUS.concluido.hex},
  ].filter(function(d){return d.value>0;});
  var tPendente=tarefas.filter(function(t){return t.status==="pendente";}).length;
  var tEmAnd=tarefas.filter(function(t){return t.status==="em_andamento";}).length;
  var tPausado=tarefas.filter(function(t){return t.status==="pausado";}).length;
  var tConc=tarefas.filter(function(t){return t.status==="concluido";}).length;
  var tarefasPie=[
    {name:"Pendentes",value:tPendente,fill:"#d97706"},
    {name:"Em andamento",value:tEmAnd,fill:"#15803d"},
    {name:"Pausadas",value:tPausado,fill:"#dc2626"},
    {name:"Concluídas",value:tConc,fill:"#6b7280"},
  ].filter(function(d){return d.value>0;});
  var membroTarefas=team.map(function(m){
    var ativas=tarefas.filter(function(t){return t.membroId===m.id&&t.status!=="concluido";}).length;
    var conc=tarefas.filter(function(t){return t.membroId===m.id&&t.status==="concluido";}).length;
    return {name:m.name,ativas:ativas,conc:conc};
  }).filter(function(m){return m.ativas+m.conc>0;}).sort(function(a,b){return b.ativas-a.ativas;});

  async function exportPNG(){
    var el=document.getElementById("dashboard-content");
    if(!el)return;
    var canvas=await html2canvas(el,{backgroundColor:"#f9fafb",scale:2});
    var link=document.createElement("a");
    link.download="MM_Dashboard_"+TODAY+".png";
    link.href=canvas.toDataURL("image/png");
    link.click();
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"flex-end",marginBottom:"12px"}}>
        <button onClick={exportPNG} style={Object.assign({},B.sec,{fontSize:"12px",padding:"6px 14px"})}>Exportar PNG</button>
      </div>
      <div id="dashboard-content">
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"8px",marginBottom:"1.25rem"}}>
          {[["Ativos",active,STATUS.ativo],["Bloqueados",blocked,STATUS.bloqueado],["Aguardando",waiting,STATUS.aguardando],["Concluidos",done,STATUS.concluido]].map(function(item){
            var l=item[0],v=item[1],s=item[2];
            return (
              <div key={l} style={Object.assign({},B.card,{textAlign:"center",borderTop:"3px solid "+s.hex,padding:"12px 8px"})}>
                <div style={{fontSize:"9px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#9ca3af",marginBottom:"6px"}}>{l}</div>
                <div style={{fontSize:"26px",fontWeight:800,color:s.color}}>{v}</div>
              </div>
            );
          })}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px",marginBottom:"1.25rem"}}>
          <div style={Object.assign({},B.card,{textAlign:"center"})}><div style={B.lbl}>Equipe</div><div style={{fontSize:"22px",fontWeight:700}}>{team.length}</div><div style={{fontSize:"11px",color:"#9ca3af"}}>{teamLoad.filter(function(m){return m.active;}).length} ativos hoje</div></div>
          <div style={Object.assign({},B.card,{textAlign:"center"})}><div style={B.lbl}>Contratantes</div><div style={{fontSize:"22px",fontWeight:700}}>{cont.length}</div></div>
          <div style={Object.assign({},B.card,{textAlign:"center"})}><div style={B.lbl}>Dias registrados</div><div style={{fontSize:"22px",fontWeight:700}}>{Object.keys(history).length}</div></div>
        </div>
        {chartProj.length>0&&(
          <div style={Object.assign({},B.card,{marginBottom:"1rem"})}>
            <div style={B.lbl}>Dedicacao por projeto (ponderada)</div>
            <ResponsiveContainer width="100%" height={Math.max(160,chartProj.length*34)}>
              <BarChart data={chartProj} layout="vertical" margin={{top:0,right:48,left:8,bottom:0}}>
                <XAxis type="number" domain={[0,100]} tickFormatter={function(v){return v+"%";}} tick={{fontSize:10}} axisLine={false} tickLine={false}/>
                <YAxis type="category" dataKey="name" width={135} tick={{fontSize:11}} axisLine={false} tickLine={false}/>
                <Tooltip formatter={function(v,n,p){return [v+"% ("+p.payload.wpts+" pts)","Dedicacao"];}} contentStyle={{fontSize:12,borderRadius:6}}/>
                <Bar dataKey="pct" radius={[0,4,4,0]} maxBarSize={18}>
                  {chartProj.map(function(e,i){return <Cell key={i} fill={colorMap[e.pid]||COLORS[i%COLORS.length]}/>;}) }
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
        {chartTeam.length>0&&(
          <div style={Object.assign({},B.card,{marginBottom:"1rem"})}>
            <div style={B.lbl}>Carga da equipe</div>
            <div style={{fontSize:"11px",color:"#9ca3af",marginBottom:"12px"}}>
              Pesos: {Object.entries(weights).map(function(e){return <span key={e[0]} style={{marginRight:"10px"}}><b>{e[0]}</b>={e[1]}</span>;})}
            </div>
            {chartTeam.map(function(m){
              var rc=ROLE_COLOR[m.role]||"#888";
              return (
                <div key={m.name} style={{marginBottom:"12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"4px"}}>
                    <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                      <span style={{fontSize:"13px",fontWeight:600}}>{m.name}</span>
                      <span style={{fontSize:"10px",padding:"1px 7px",borderRadius:"10px",background:rc+"18",color:rc,fontWeight:600}}>{m.role} - {m.weight}</span>
                    </div>
                    <span style={{fontSize:"13px",fontWeight:700,color:"#374151"}}>{m.pct.toFixed(1)}%</span>
                  </div>
                  <div style={{height:"7px",background:"#f3f4f6",borderRadius:"4px",overflow:"hidden"}}>
                    <div style={{height:"100%",width:Math.min(m.pct,100)+"%",background:rc,borderRadius:"4px"}}/>
                  </div>
                </div>
              );
            })}
          </div>
        )}
        {tarefas.length>0&&(
          <div style={Object.assign({},B.card,{marginBottom:"1rem"})}>
            <div style={B.lbl}>Status das tarefas</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:"8px",marginBottom:"1rem"}}>
              {[
                {label:"Pendentes",value:tPendente,color:"#92400e",bg:"#fef9c3"},
                {label:"Em andamento",value:tEmAnd,color:"#15803d",bg:"#dcfce7"},
                {label:"Pausadas",value:tPausado,color:"#dc2626",bg:"#fee2e2"},
                {label:"Concluídas",value:tConc,color:"#6b7280",bg:"#f3f4f6"},
              ].map(function(item){
                return (
                  <div key={item.label} style={{background:item.bg,borderRadius:"8px",padding:"10px 8px",textAlign:"center"}}>
                    <div style={{fontSize:"9px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:item.color,marginBottom:"4px"}}>{item.label}</div>
                    <div style={{fontSize:"26px",fontWeight:800,color:item.color}}>{item.value}</div>
                  </div>
                );
              })}
            </div>
            {tarefasPie.length>0&&(
              <div style={{display:"flex",alignItems:"center",gap:"20px"}}>
                <PieChart width={120} height={120}>
                  <Pie data={tarefasPie} cx={55} cy={55} innerRadius={28} outerRadius={52} dataKey="value" paddingAngle={3}>
                    {tarefasPie.map(function(e,i){return <Cell key={i} fill={e.fill}/>;}) }
                  </Pie>
                  <Tooltip formatter={function(v,n){return [v+" tarefa"+(v!==1?"s":""),n];}} contentStyle={{fontSize:11,borderRadius:6}}/>
                </PieChart>
                <div style={{flex:1}}>
                  {tarefasPie.map(function(s){
                    return (
                      <div key={s.name} style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"7px"}}>
                        <div style={{width:"10px",height:"10px",borderRadius:"2px",background:s.fill,flexShrink:0}}/>
                        <span style={{fontSize:"12px",color:"#6b7280",flex:1}}>{s.name}</span>
                        <span style={{fontSize:"14px",fontWeight:700,color:"#111"}}>{s.value}</span>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
            {membroTarefas.length>0&&(
              <div style={{marginTop:"1rem",borderTop:"1px solid #f0f0f0",paddingTop:"1rem"}}>
                <div style={{fontSize:"10px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#9ca3af",marginBottom:"8px"}}>Por colaborador</div>
                {membroTarefas.map(function(m){
                  return (
                    <div key={m.name} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #f9fafb"}}>
                      <span style={{fontSize:"13px",color:"#374151",fontWeight:500}}>{m.name}</span>
                      <div style={{display:"flex",gap:"6px"}}>
                        {m.ativas>0&&<span style={{fontSize:"11px",padding:"2px 9px",borderRadius:"20px",background:"#dbeafe",color:"#2563EB",fontWeight:600}}>{m.ativas} ativa{m.ativas!==1?"s":""}</span>}
                        {m.conc>0&&<span style={{fontSize:"11px",padding:"2px 9px",borderRadius:"20px",background:"#f3f4f6",color:"#6b7280",fontWeight:600}}>{m.conc} concluída{m.conc!==1?"s":""}</span>}
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
        {statusPie.length>0&&(
          <div style={B.card}>
            <div style={B.lbl}>Status dos projetos</div>
            <div style={{display:"flex",alignItems:"center",gap:"16px"}}>
              <PieChart width={110} height={110}>
                <Pie data={statusPie} cx={50} cy={50} innerRadius={26} outerRadius={48} dataKey="value" paddingAngle={2}>
                  {statusPie.map(function(e,i){return <Cell key={i} fill={e.fill}/>;}) }
                </Pie>
              </PieChart>
              <div style={{flex:1}}>
                {statusPie.map(function(s){
                  return (
                    <div key={s.name} style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"6px"}}>
                      <div style={{width:"10px",height:"10px",borderRadius:"2px",background:s.fill,flexShrink:0}}/>
                      <span style={{fontSize:"12px",color:"#6b7280",flex:1}}>{s.name}</span>
                      <span style={{fontSize:"14px",fontWeight:700}}>{s.value}</span>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function getSIs(allocSubitems,mid,pid){
  var val=(allocSubitems[mid]||{})[pid];
  if(!val)return [];
  return Array.isArray(val)?val:(val?[val]:[]);
}

function HojeTab(props){
  var team=props.team,projects=props.projects,cont=props.cont,normA=props.normA,colorMap=props.colorMap,toggleProj=props.toggleProj,date=props.date,setDate=props.setDate,savedDate=props.savedDate,assignedCnt=props.assignedCnt,saveAlloc=props.saveAlloc,flash=props.flash,report=props.report,allocSubitems=props.allocSubitems||{},setAllocSubitems=props.setAllocSubitems||function(){};
  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"1.25rem",flexWrap:"wrap"}}>
        <input type="date" value={date} onChange={function(e){setDate(e.target.value);}} style={{fontSize:"13px",padding:"6px 10px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:"#111"}}/>
        <span style={{fontSize:"13px",color:"#6b7280"}}>{cap(fmtDay(date))}</span>
        {savedDate&&<span style={{fontSize:"11px",background:"#dcfce7",color:"#15803d",padding:"2px 8px",borderRadius:"20px",fontWeight:600}}>Salvo</span>}
        <span style={{marginLeft:"auto",fontSize:"12px",color:"#9ca3af"}}>{assignedCnt}/{team.length} alocados</span>
      </div>
      {ROLES.map(function(role){
        var members=team.filter(function(m){return m.role===role;});
        if(!members.length)return null;
        var rc=ROLE_COLOR[role]||"#888";
        return (
          <div key={role} style={{marginBottom:"1.5rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:"7px",marginBottom:"8px"}}>
              <div style={{width:"3px",height:"14px",borderRadius:"2px",background:rc}}/>
              <span style={{fontSize:"11px",fontWeight:700,color:"#374151",textTransform:"uppercase",letterSpacing:".07em"}}>{ROLES_DISPLAY[role]}</span>
            </div>
            {members.map(function(m){
              var pids=normA[m.id]||[];
              var sCol=pids.length===1?colorMap[pids[0]]:null;
              return (
                <div key={m.id} style={Object.assign({},B.card,{display:"flex",alignItems:"flex-start",gap:"12px"})}>
                  <div style={{width:"36px",height:"36px",borderRadius:"50%",flexShrink:0,marginTop:"1px",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"11px",fontWeight:700,background:sCol?sCol+"20":pids.length>1?RED+"12":"#f3f4f6",color:sCol||(pids.length>1?RED:"#9ca3af"),border:"1.5px solid "+(sCol?sCol+"30":pids.length>1?RED+"25":"#e5e7eb")}}>
                    {m.name.slice(0,2).toUpperCase()}
                  </div>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:"14px",fontWeight:600,color:"#111",marginBottom:"6px"}}>{m.name}</div>
                    {pids.length>0&&(
                      <div style={{display:"flex",flexWrap:"wrap",gap:"4px",marginBottom:"6px"}}>
                        {pids.map(function(pid){
                          var proj=projects.find(function(p){return p.id===pid;});
                          var c=colorMap[pid];
                          var ctr=cont.find(function(x){return x.id===(proj?proj.cId:"");});
                          var subitems=(proj&&proj.subitems)||[];
                          var selSIs=getSIs(allocSubitems,m.id,pid);
                          return (
                            <div key={pid} style={{display:"inline-flex",flexDirection:"column",gap:"3px"}}>
                              <span style={{fontSize:"11px",padding:"3px 7px 3px 9px",borderRadius:"20px",background:c+"18",color:c,border:"1px solid "+c+"35",display:"inline-flex",alignItems:"center",gap:"3px",lineHeight:1.5,fontWeight:500}}>
                                {ctr?ctr.name+" - ":""}{proj?proj.name:pid}
                                <button onClick={function(){
                                  toggleProj(m.id,pid);
                                  var ns=Object.assign({},allocSubitems[m.id]||{});
                                  delete ns[pid];
                                  setAllocSubitems(Object.assign({},allocSubitems,{[m.id]:ns}));
                                }} style={{background:"none",border:"none",cursor:"pointer",color:"inherit",padding:"0 1px",fontSize:"13px",opacity:.5}}>x</button>
                              </span>
                              {subitems.length>0&&(
                                <div style={{display:"flex",flexWrap:"wrap",gap:"3px",paddingLeft:"4px"}}>
                                  {selSIs.map(function(siId){
                                    var si=subitems.find(function(x){return x.id===siId;});
                                    if(!si)return null;
                                    return (
                                      <span key={siId} style={{fontSize:"10px",padding:"2px 5px 2px 7px",borderRadius:"10px",background:c+"15",color:c,border:"1px solid "+c+"30",display:"inline-flex",alignItems:"center",gap:"2px",lineHeight:1.5}}>
                                        {si.name}
                                        <button onClick={function(){
                                          var next=selSIs.filter(function(x){return x!==siId;});
                                          var ns=Object.assign({},allocSubitems[m.id]||{},{[pid]:next});
                                          setAllocSubitems(Object.assign({},allocSubitems,{[m.id]:ns}));
                                        }} style={{background:"none",border:"none",cursor:"pointer",color:"inherit",padding:"0 1px",fontSize:"11px",opacity:.6,lineHeight:1}}>×</button>
                                      </span>
                                    );
                                  })}
                                  {selSIs.length<subitems.length&&(
                                    <select value="" onChange={function(e){
                                      if(!e.target.value)return;
                                      var next=[...selSIs,e.target.value];
                                      var ns=Object.assign({},allocSubitems[m.id]||{},{[pid]:next});
                                      setAllocSubitems(Object.assign({},allocSubitems,{[m.id]:ns}));
                                      e.target.value="";
                                    }} style={{fontSize:"10px",padding:"2px 6px",borderRadius:"10px",border:"1px dashed "+c+"50",background:"transparent",color:c,cursor:"pointer",outline:"none"}}>
                                      <option value="">+ atividade</option>
                                      {subitems.filter(function(si){return !selSIs.includes(si.id);}).map(function(si){
                                        return <option key={si.id} value={si.id}>{si.name}</option>;
                                      })}
                                    </select>
                                  )}
                                </div>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    )}
                    <select onChange={function(e){if(e.target.value){toggleProj(m.id,e.target.value);e.target.value="";}}} style={{fontSize:"12px",padding:"3px 8px",borderRadius:"6px",border:"1px dashed #e5e7eb",background:"transparent",color:"#9ca3af",cursor:"pointer"}}>
                      <option value="">+ adicionar projeto</option>
                      {cont.map(function(ct){
                        var ps=projects.filter(function(p){return p.cId===ct.id&&p.status!=="concluido"&&!pids.includes(p.id);});
                        if(!ps.length)return null;
                        return <optgroup key={ct.id} label={ct.name}>{ps.map(function(p){return <option key={p.id} value={p.id}>{p.name}{p.status!=="ativo"?" ("+STATUS[p.status].label+")":""}</option>;})}</optgroup>;
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
        <button onClick={saveAlloc} style={Object.assign({},B.pri,{background:flash?"#16a34a":RED,minWidth:"140px"})}>{flash||"Salvar alocacao"}</button>
      </div>
      {report.length>0&&(
        <div>
          <div style={B.lbl}>Previa ponderada</div>
          {report.map(function(r){
            var p=projects.find(function(x){return x.id===r.pid;});
            var c=cont.find(function(x){return x.id===(p?p.cId:"");});
            return (
              <div key={r.pid} style={{marginBottom:"12px"}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:"3px"}}>
                  <div>
                    <span style={{fontSize:"13px",fontWeight:600,color:"#111"}}>{r.name}</span>
                    {c&&<span style={{fontSize:"11px",color:"#9ca3af",marginLeft:"5px"}}>{c.name}</span>}
                  </div>
                  <span style={{fontSize:"12px",fontWeight:700,color:"#6b7280"}}>{r.pct.toFixed(1)}%</span>
                </div>
                <div style={{height:"5px",background:"#f3f4f6",borderRadius:"3px",overflow:"hidden"}}>
                  <div style={{height:"100%",width:r.pct+"%",background:colorMap[r.pid],borderRadius:"3px"}}/>
                </div>
                <div style={{fontSize:"11px",color:"#9ca3af",marginTop:"2px"}}>{r.people.join(", ")} - {r.wpts} pts</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function getWeekDates(dateStr){
  var d=new Date(dateStr+"T12:00:00");
  var dow=d.getDay()||7;
  var mon=new Date(d);mon.setDate(d.getDate()-(dow-1));
  var dates=[];
  for(var i=0;i<5;i++){var x=new Date(mon);x.setDate(mon.getDate()+i);dates.push(x.toISOString().split("T")[0]);}
  return dates;
}

function calcAggReport(dates,history,team,projects,weights){
  if(!weights)weights={};
  var bp={};
  dates.forEach(function(d){
    var h=history[d];
    if(!h||!h.alloc)return;
    var normA=normAlloc(h.alloc);
    team.forEach(function(m){
      var pids=normA[m.id]||[];
      if(!pids.length)return;
      var w=wOf(m.role,weights);
      pids.forEach(function(pid){
        if(!bp[pid])bp[pid]={wpts:0,byMember:{}};
        bp[pid].wpts+=w;
        if(!bp[pid].byMember[m.id])bp[pid].byMember[m.id]={name:m.name,role:m.role,w:w,days:0};
        bp[pid].byMember[m.id].days++;
      });
    });
  });
  var tot=Object.values(bp).reduce(function(s,v){return s+v.wpts;},0);
  return Object.entries(bp).map(function(e){
    var pid=e[0],val=e[1];
    var p=projects.find(function(x){return x.id===pid;});
    var members=Object.values(val.byMember);
    return {
      pid:pid,
      name:p?p.name:pid,
      people:members.map(function(x){return x.name;}),
      detail:members.map(function(x){return {name:x.name,role:x.role,w:+(x.w.toFixed(2)),days:x.days};}),
      wpts:+(val.wpts.toFixed(2)),
      pct:tot?val.wpts/tot*100:0
    };
  }).sort(function(a,b){return b.wpts-a.wpts;});
}

function RelTab(props){
  var report=props.report,date=props.date,setDate=props.setDate,projects=props.projects,cont=props.cont,normA=props.normA,team=props.team,weights=props.weights,assignedCnt=props.assignedCnt,colorMap=props.colorMap,genMsg=props.genMsg,aiLoad=props.aiLoad,aiMsg=props.aiMsg,exportExcel=props.exportExcel,exportPDF=props.exportPDF,history=props.history||{};
  var pts=useState("dia"); var period=pts[0],setPeriod=pts[1];

  // Calcula datas do período e relatório agregado
  var periodDates=[];
  var periodLabel="";
  if(period==="dia"){
    periodDates=[date];
    periodLabel=cap(fmtDay(date))+", "+fmtDate(date);
  } else if(period==="semana"){
    periodDates=getWeekDates(date);
    periodLabel="Semana de "+fmtDate(periodDates[0])+" a "+fmtDate(periodDates[4]);
  } else {
    var ym=date.slice(0,7);
    var parts2=ym.split("-");
    periodDates=Object.keys(history).filter(function(d){return d.startsWith(ym);}).sort();
    periodLabel=MONTH_PT[parseInt(parts2[1])-1]+" "+parts2[0];
  }

  var activeReport=period==="dia"?report:calcAggReport(periodDates,history,team,projects,weights);
  var totalPts=activeReport.reduce(function(s,r){return s+r.wpts;},0);
  var activePeople=period==="dia"?assignedCnt:[...new Set(activeReport.flatMap(function(r){return r.people;}))].length;
  var diasComDados=period==="dia"?1:periodDates.filter(function(d){return history[d]&&history[d].alloc;}).length;

  function shiftDate(delta){
    var d=new Date(date+"T12:00:00");
    if(period==="semana") d.setDate(d.getDate()+delta*7);
    else if(period==="mes") d.setMonth(d.getMonth()+delta);
    else d.setDate(d.getDate()+delta);
    setDate(d.toISOString().split("T")[0]);
  }

  return (
    <div>
      {/* Seletor de período */}
      <div style={{display:"flex",gap:"6px",marginBottom:"1rem",background:"#f3f4f6",borderRadius:"8px",padding:"3px",width:"fit-content"}}>
        {[["dia","Diário"],["semana","Semanal"],["mes","Mensal"]].map(function(item){
          return (
            <button key={item[0]} onClick={function(){setPeriod(item[0]);}} style={{background:period===item[0]?"#fff":"transparent",border:"none",cursor:"pointer",padding:"5px 14px",borderRadius:"6px",fontSize:"12px",fontWeight:period===item[0]?700:400,color:period===item[0]?"#111":"#9ca3af"}}>
              {item[1]}
            </button>
          );
        })}
      </div>
      {/* Navegação de data */}
      <div style={{display:"flex",gap:"8px",alignItems:"center",marginBottom:"1.5rem",flexWrap:"wrap"}}>
        <button onClick={function(){shiftDate(-1);}} style={Object.assign({},B.sec,{padding:"5px 12px"})}>‹</button>
        {period==="dia"&&(
          <input type="date" value={date} onChange={function(e){setDate(e.target.value);}} style={{fontSize:"13px",padding:"6px 10px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:"#111"}}/>
        )}
        <span style={{fontSize:"13px",fontWeight:600,color:"#374151"}}>{periodLabel}</span>
        <button onClick={function(){shiftDate(1);}} style={Object.assign({},B.sec,{padding:"5px 12px"})}>›</button>
        {period!=="dia"&&<span style={{fontSize:"11px",color:"#9ca3af",marginLeft:"4px"}}>{diasComDados} dia{diasComDados!==1?"s":""} com registro</span>}
      </div>
      {/* Cards de resumo */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"8px",marginBottom:"1.5rem"}}>
        {[["Projetos",activeReport.length],["Pessoas",activePeople],["Pts pond.",totalPts.toFixed(1)]].map(function(item){
          return (
            <div key={item[0]} style={{background:"#f9fafb",borderRadius:"8px",padding:"14px",textAlign:"center",border:"1px solid #f0f0f0"}}>
              <div style={{fontSize:"9px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#9ca3af",marginBottom:"6px"}}>{item[0]}</div>
              <div style={{fontSize:"24px",fontWeight:700,color:"#111"}}>{item[1]}</div>
            </div>
          );
        })}
      </div>
      {activeReport.length===0?(
        <div style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}>Nenhuma alocação registrada neste período.</div>
      ):(
        <div>
          {activeReport.map(function(r){
            var p=projects.find(function(x){return x.id===r.pid;});
            var c=cont.find(function(x){return x.id===(p?p.cId:"");});
            return (
              <div key={r.pid} style={{marginBottom:"18px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"5px"}}>
                  <div>
                    <div style={{fontSize:"14px",fontWeight:700,color:"#111"}}>{r.name}</div>
                    {c&&<div style={{fontSize:"11px",color:"#9ca3af",marginTop:"1px"}}>{c.name}</div>}
                  </div>
                  <span style={{fontSize:"13px",fontWeight:700,color:"#6b7280",marginLeft:"8px",flexShrink:0}}>{r.pct.toFixed(1)}%</span>
                </div>
                <div style={{height:"8px",background:"#f3f4f6",borderRadius:"4px",overflow:"hidden",marginBottom:"4px"}}>
                  <div style={{height:"100%",width:r.pct+"%",background:colorMap[r.pid],borderRadius:"4px"}}/>
                </div>
                <div style={{fontSize:"12px",color:"#9ca3af"}}>
                  {r.detail.map(function(d){
                    return d.name+" ("+d.role+(d.days?" — "+d.days+"d":"")+", "+d.w+" pts)";
                  }).join(" · ")}
                </div>
              </div>
            );
          })}
          <div style={{borderTop:"1px solid #f0f0f0",paddingTop:"1.25rem",marginTop:"0.5rem"}}>
            <div style={B.lbl}>Exportar</div>
            <div style={{display:"flex",gap:"8px",flexWrap:"wrap",marginBottom:"10px"}}>
              <button onClick={function(){
                var expNormA=period==="dia"?normA:(function(){var r={};activeReport.forEach(function(ar){ar.detail.forEach(function(d){var m=team.find(function(x){return x.name===d.name;});if(m){if(!r[m.id])r[m.id]=[];if(r[m.id].indexOf(ar.pid)===-1)r[m.id].push(ar.pid);}});});return r;})();
                var expLabel=period==="dia"?null:periodLabel;
                var expDate=periodDates[0]||date;
                exportExcel(activeReport,expDate,team,expNormA,projects,cont,weights,expLabel);
              }} style={Object.assign({},B.pri,{display:"flex",alignItems:"center",gap:"5px"})}>Excel (.xlsx)</button>
              <button onClick={function(){
                var expNormA=period==="dia"?normA:(function(){var r={};activeReport.forEach(function(ar){ar.detail.forEach(function(d){var m=team.find(function(x){return x.name===d.name;});if(m){if(!r[m.id])r[m.id]=[];if(r[m.id].indexOf(ar.pid)===-1)r[m.id].push(ar.pid);}});});return r;})();
                var expLabel=period==="dia"?null:periodLabel;
                var expDate=periodDates[0]||date;
                exportPDF(activeReport,expDate,projects,cont,expNormA,team,weights,expLabel);
              }} style={Object.assign({},B.sec,{color:"#374151",fontWeight:500,display:"flex",alignItems:"center",gap:"5px"})}>PDF / Imprimir</button>
            </div>
            <button onClick={genMsg} disabled={aiLoad} style={Object.assign({},B.sec,{color:aiLoad?"#9ca3af":RED,width:"100%",textAlign:"center"})}>
              {aiLoad?"Gerando...":"Gerar mensagem para o Manoel"}
            </button>
            {aiMsg&&(
              <div style={{background:"#f9fafb",borderRadius:"8px",padding:"12px 14px",marginTop:"10px",fontSize:"13px",lineHeight:1.7,whiteSpace:"pre-wrap",border:"1px solid #f0f0f0"}}>
                {aiMsg}
                <button onClick={function(){if(navigator.clipboard)navigator.clipboard.writeText(aiMsg);}} style={Object.assign({},B.ghost,{display:"block",marginTop:"6px",paddingTop:"6px",borderTop:"1px solid #f0f0f0",width:"100%",textAlign:"left"})}>Copiar mensagem</button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function ProjTab(props){
  var projects=props.projects,cont=props.cont,saveProjects=props.saveProjects,saveCont=props.saveCont,colorMap=props.colorMap,diary=props.diary,saveDiary=props.saveDiary,team=props.team,history=props.history;
  var vs=useState("cont"); var view=vs[0],setView=vs[1];
  var sfs=useState(false); var showForm=sfs[0],setShowForm=sfs[1];
  var eps=useState(null); var editProj=eps[0],setEditProj=eps[1];
  var dts=useState(null); var detail=dts[0],setDetail=dts[1];
  var sds=useState(false); var showDone=sds[0],setShowDone=sds[1];
  var srs=useState(""); var search=srs[0],setSearch=srs[1];
  var ecs=useState(null); var editCont=ecs[0],setEditCont=ecs[1];
  var active=projects.filter(function(p){return p.status!=="concluido";});
  var done=projects.filter(function(p){return p.status==="concluido";});
  var filtered=search?active.filter(function(p){
    var cname=(cont.find(function(c){return c.id===p.cId;})||{}).name||"";
    return p.name.toLowerCase().includes(search.toLowerCase())||cname.toLowerCase().includes(search.toLowerCase());
  }):active;
  function cycleStatus(proj){saveProjects(projects.map(function(p){return p.id===proj.id?Object.assign({},p,{status:ST_CYC[p.status]||"ativo"}):p;}));}
  function baixa(proj){saveProjects(projects.map(function(p){return p.id===proj.id?Object.assign({},p,{status:"concluido"}):p;}));}
  if(detail){
    var proj=projects.find(function(p){return p.id===detail;});
    var ctr=cont.find(function(c){return c.id===(proj?proj.cId:"");});
    if(!proj)return null;
    return <ProjDetail proj={proj} ctr={ctr} colorMap={colorMap} diary={diary} saveDiary={saveDiary} team={team} onBack={function(){setDetail(null);}} projects={projects} saveProjects={saveProjects} history={history}/>;
  }
  return (
    <div>
      <div style={{display:"flex",gap:"8px",marginBottom:"1.25rem",flexWrap:"wrap"}}>
        <input placeholder="Buscar..." value={search} onChange={function(e){setSearch(e.target.value);}} style={Object.assign({},B.inp,{width:"auto",flex:1,minWidth:"160px"})}/>
        <button onClick={function(){setShowForm(function(v){return !v;});}} style={B.pri}>+ Novo projeto</button>
      </div>
      <div style={{display:"flex",marginBottom:"1.25rem",background:"#f3f4f6",borderRadius:"8px",padding:"3px",width:"fit-content"}}>
        {[["cont","Por contratante"],["status","Por status"]].map(function(item){
          return <button key={item[0]} onClick={function(){setView(item[0]);}} style={{background:view===item[0]?"#fff":"transparent",border:"none",cursor:"pointer",padding:"5px 12px",borderRadius:"6px",fontSize:"12px",fontWeight:view===item[0]?600:400,color:view===item[0]?"#111":"#9ca3af"}}>{item[1]}</button>;
        })}
      </div>
      {showForm&&<ProjForm cont={cont} saveCont={saveCont} onSave={function(p){saveProjects([...projects,Object.assign({},p,{id:"p"+uid()})]);setShowForm(false);}} onCancel={function(){setShowForm(false);}}/>}
      {view==="cont"&&cont.map(function(ct){
        var ps=filtered.filter(function(p){return p.cId===ct.id;});
        if(!ps.length)return null;
        return (
          <div key={ct.id} style={{marginBottom:"1.75rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"8px",paddingBottom:"8px",borderBottom:"2px solid "+RED+"20"}}>
              <div style={{width:"6px",height:"6px",borderRadius:"50%",background:RED}}/>
              {editCont===ct.id?(
                <input defaultValue={ct.name} autoFocus onBlur={function(e){saveCont(cont.map(function(c){return c.id===ct.id?Object.assign({},c,{name:e.target.value}):c;}));setEditCont(null);}} style={Object.assign({},B.inp,{fontSize:"14px",fontWeight:700,padding:"2px 6px",width:"auto",flex:1})} onKeyDown={function(e){if(e.key==="Enter")e.target.blur();}}/>
              ):(
                <span style={{fontSize:"14px",fontWeight:700,color:"#111",cursor:"text"}} onDoubleClick={function(){setEditCont(ct.id);}}>{ct.name}</span>
              )}
              <span style={{fontSize:"11px",color:"#9ca3af"}}>{ps.length} proj.</span>
              <div style={{marginLeft:"auto",display:"flex",gap:"8px",alignItems:"center"}}>
                <span style={{fontSize:"9px",color:"#d1d5db"}}>duplo clique para editar</span>
                {(
                  <button onClick={function(){if(window.confirm("Apagar contratante "+ct.name+"?"))saveCont(cont.filter(function(c){return c.id!==ct.id;}));}} style={Object.assign({},B.ghost,{color:"#dc2626",fontSize:"11px",padding:"1px 6px",border:"1px solid #fecaca",borderRadius:"4px"})}>apagar</button>
                )}
              </div>
            </div>
            {ps.map(function(proj){return <ProjCard key={proj.id} proj={proj} cont={cont} colorMap={colorMap} editProj={editProj} setEditProj={setEditProj} cycleStatus={cycleStatus} baixa={baixa} projects={projects} saveProjects={saveProjects} saveCont={saveCont} diaryCount={(diary[proj.id]||[]).length} onDetail={function(){setDetail(proj.id);}}/>;}) }
          </div>
        );
      })}
      {view==="status"&&["ativo","aguardando","bloqueado"].map(function(sk){
        var ps=filtered.filter(function(p){return p.status===sk;});
        if(!ps.length)return null;
        return (
          <div key={sk} style={{marginBottom:"1.75rem"}}>
            <div style={Object.assign({},B.lbl,{display:"flex",alignItems:"center",gap:"6px"})}><span style={{display:"inline-block",width:"6px",height:"6px",borderRadius:"50%",background:STATUS[sk].hex}}/>{STATUS[sk].label} ({ps.length})</div>
            {ps.map(function(proj){return <ProjCard key={proj.id} proj={proj} cont={cont} colorMap={colorMap} editProj={editProj} setEditProj={setEditProj} cycleStatus={cycleStatus} baixa={baixa} projects={projects} saveProjects={saveProjects} saveCont={saveCont} diaryCount={(diary[proj.id]||[]).length} onDetail={function(){setDetail(proj.id);}}/>;}) }
          </div>
        );
      })}
      {done.length>0&&(
        <div>
          <button onClick={function(){setShowDone(function(v){return !v;});}} style={Object.assign({},B.ghost,{display:"flex",alignItems:"center",gap:"5px",color:"#9ca3af"})}>
            <span style={{fontSize:"9px"}}>{showDone?"v":">"}</span> Concluidos ({done.length})
          </button>
          {showDone&&done.map(function(proj){
            return (
              <div key={proj.id} style={Object.assign({},B.card,{opacity:.55,background:"#f9fafb",marginTop:"6px"})}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:"8px"}}>
                  <div>
                    <div style={{fontSize:"13px",color:"#6b7280",textDecoration:"line-through"}}>{proj.name}</div>
                    <div style={{fontSize:"11px",color:"#9ca3af"}}>{(cont.find(function(c){return c.id===proj.cId;})||{}).name}</div>
                  </div>
                  <button onClick={function(){saveProjects(projects.map(function(p){return p.id===proj.id?Object.assign({},p,{status:"ativo"}):p;}));}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"3px 10px"})}>reativar</button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function ProjCard(props){
  var proj=props.proj,cont=props.cont,colorMap=props.colorMap,editProj=props.editProj,setEditProj=props.setEditProj,cycleStatus=props.cycleStatus,baixa=props.baixa,projects=props.projects,saveProjects=props.saveProjects,saveCont=props.saveCont,diaryCount=props.diaryCount,onDetail=props.onDetail;
  var sc=STATUS[proj.status]||STATUS.ativo;
  var color=colorMap[proj.id];
  if(editProj&&editProj.id===proj.id){
    return <ProjForm cont={cont} saveCont={saveCont} initial={editProj} onSave={function(p){saveProjects(projects.map(function(x){return x.id===editProj.id?Object.assign({},editProj,p):x;}));setEditProj(null);}} onCancel={function(){setEditProj(null);}} onDelete={function(){saveProjects(projects.filter(function(x){return x.id!==proj.id;}));setEditProj(null);}}/>;
  }
  return (
    <div style={Object.assign({},B.card,{borderLeft:"3px solid "+color})}>
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:"10px"}}>
        <div style={{flex:1,minWidth:0,cursor:"pointer"}} onClick={onDetail}>
          <div style={{fontSize:"14px",fontWeight:600,color:"#111"}}>{proj.name}</div>
          {(proj.city||proj.st)&&<div style={{fontSize:"11px",color:"#9ca3af",marginTop:"2px"}}>{[proj.city,proj.st,proj.country].filter(Boolean).join(", ")}</div>}
          {proj.svcs&&proj.svcs.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:"3px",marginTop:"6px"}}>{proj.svcs.map(function(s){return <span key={s} style={{fontSize:"10px",padding:"2px 6px",borderRadius:"4px",background:"#f3f4f6",color:"#6b7280",border:"1px solid #e5e7eb"}}>{SERVICES_DISPLAY[s]||s}</span>;})}</div>}
          {proj.note&&<div style={{fontSize:"12px",color:"#9ca3af",marginTop:"5px",fontStyle:"italic"}}>{proj.note}</div>}
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:"5px",alignItems:"flex-end",flexShrink:0}}>
          {proj.status!=="concluido"&&(
            <button onClick={function(){cycleStatus(proj);}} style={{fontSize:"11px",fontWeight:600,padding:"3px 9px",borderRadius:"20px",background:sc.bg,color:sc.color,border:"none",cursor:"pointer",display:"inline-flex",alignItems:"center",gap:"4px"}}>
              {sc.label} ↻
            </button>
          )}
          <div style={{display:"flex",gap:"4px",alignItems:"center"}}>
            {diaryCount>0&&<span style={{fontSize:"10px",padding:"1px 6px",borderRadius:"10px",background:"#eff6ff",color:"#2563EB",fontWeight:600,cursor:"pointer"}} onClick={onDetail}>{diaryCount} entrada{diaryCount!==1?"s":""}</span>}
            <button onClick={function(){setEditProj(Object.assign({},proj));}} style={B.ghost}>editar</button>
            <button onClick={function(){baixa(proj);}} style={B.ghost}>baixa</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function Lbl(lprops){return <div style={{marginBottom:"10px"}}><div style={{fontSize:"10px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".06em",marginBottom:"5px"}}>{lprops.t}</div>{lprops.children}</div>;}

function ContratanteField(props){
  var cont=props.cont,saveCont=props.saveCont,value=props.value,onChange=props.onChange;
  var init=cont.find(function(c){return c.id===value;});
  var ivs=useState(init?init.name:""); var inputVal=ivs[0],setInputVal=ivs[1];
  var ops=useState(false); var open=ops[0],setOpen=ops[1];
  var filtered=cont.filter(function(c){return c.name.toLowerCase().includes(inputVal.toLowerCase());});
  var exactMatch=cont.find(function(c){return c.name.toLowerCase()===inputVal.trim().toLowerCase();});
  function select(c){onChange(c.id);setInputVal(c.name);setOpen(false);}
  function createNew(){
    if(!inputVal.trim())return;
    if(exactMatch){select(exactMatch);return;}
    var c={id:"c"+uid(),name:inputVal.trim(),contact:""};
    saveCont([...cont,c]);
    onChange(c.id);
    setOpen(false);
  }
  function deleteCont(id,e){
    e.preventDefault();e.stopPropagation();
    if(!window.confirm("Apagar contratante da lista?"))return;
    saveCont(cont.filter(function(c){return c.id!==id;}));
    if(value===id){onChange("");setInputVal("");}
  }
  return (
    <div style={{position:"relative"}}>
      <input
        value={inputVal}
        onChange={function(e){setInputVal(e.target.value);setOpen(true);onChange("");}}
        onFocus={function(){setOpen(true);}}
        onBlur={function(){setTimeout(function(){setOpen(false);},200);}}
        onKeyDown={function(e){if(e.key==="Enter"){e.preventDefault();createNew();}if(e.key==="Escape")setOpen(false);}}
        placeholder="Digite ou selecione o contratante..."
        style={Object.assign({},B.inp,value?{borderColor:RED+"60"}:{})}
      />
      {open&&(
        <div style={{position:"absolute",top:"100%",left:0,right:0,background:"#fff",border:"1px solid #e5e7eb",borderRadius:"6px",boxShadow:"0 4px 12px rgba(0,0,0,.1)",zIndex:50,maxHeight:"180px",overflowY:"auto",marginTop:"2px"}}>
          {filtered.map(function(c){
            return (
              <div key={c.id} onMouseDown={function(){select(c);}} style={{display:"flex",alignItems:"center",padding:"8px 10px",cursor:"pointer",background:value===c.id?RED+"0d":"transparent"}}>
                <span style={{flex:1,fontSize:"13px",color:"#111"}}>{c.name}</span>
                <button onMouseDown={function(e){deleteCont(c.id,e);}} style={{background:"none",border:"none",cursor:"pointer",color:"#dc2626",fontSize:"16px",padding:"0 4px",lineHeight:1,opacity:.4}}>×</button>
              </div>
            );
          })}
          {inputVal.trim()&&!exactMatch&&(
            <div onMouseDown={createNew} style={{padding:"8px 10px",fontSize:"13px",color:RED,cursor:"pointer",borderTop:filtered.length?"1px solid #f0f0f0":"none",fontWeight:600}}>
              + Criar "{inputVal.trim()}"
            </div>
          )}
          {!filtered.length&&!inputVal.trim()&&(
            <div style={{padding:"10px",fontSize:"12px",color:"#9ca3af",textAlign:"center"}}>Nenhum contratante cadastrado</div>
          )}
        </div>
      )}
    </div>
  );
}

function ProjForm(props){
  var cont=props.cont,saveCont=props.saveCont,initial=props.initial,onSave=props.onSave,onCancel=props.onCancel,onDelete=props.onDelete;
  var blank={cId:"",name:"",svcs:[],subitems:[],city:"",st:"",country:"Brasil",status:"ativo",note:""};
  var fs=useState(initial?Object.assign({subitems:[]},initial):blank); var f=fs[0],setF=fs[1];
  var nsis=useState(""); var newSi=nsis[0],setNewSi=nsis[1];
  function tSvc(s){setF(function(p){return Object.assign({},p,{svcs:p.svcs.includes(s)?p.svcs.filter(function(x){return x!==s;}):[...p.svcs,s]});});}
  function addSubitem(){
    if(!newSi.trim())return;
    setF(function(p){return Object.assign({},p,{subitems:[...(p.subitems||[]),{id:"si"+uid(),name:newSi.trim()}]});});
    setNewSi("");
  }
  function delSubitem(id){setF(function(p){return Object.assign({},p,{subitems:(p.subitems||[]).filter(function(x){return x.id!==id;})});});}

  return (
    <div style={{background:"#fff",borderRadius:"10px",padding:"18px",marginBottom:"1.25rem",border:"1px solid "+RED+"30"}}>
      <div style={{fontSize:"13px",fontWeight:700,color:"#111",marginBottom:"14px",paddingBottom:"10px",borderBottom:"1px solid #f0f0f0"}}>{initial?"Editar projeto":"Novo projeto"}</div>
      <Lbl t="Contratante *">
        <ContratanteField cont={cont} saveCont={saveCont} value={f.cId} onChange={function(cId){setF(function(p){return Object.assign({},p,{cId:cId});});}}/>
      </Lbl>
      <Lbl t="Nome *"><input placeholder="Ex: Residencia Lago Sul" value={f.name} onChange={function(e){setF(function(p){return Object.assign({},p,{name:e.target.value});});}} style={B.inp}/></Lbl>
      <Lbl t="Servicos">
        <div style={{display:"flex",flexWrap:"wrap",gap:"5px"}}>
          {SERVICES.map(function(s){
            var on=f.svcs.includes(s);
            return <button key={s} onClick={function(){tSvc(s);}} style={{fontSize:"11px",padding:"4px 10px",borderRadius:"20px",border:"1px solid "+(on?RED:"#e5e7eb"),background:on?RED+"18":"transparent",color:on?RED:"#6b7280",cursor:"pointer",fontWeight:on?600:400}}>{SERVICES_DISPLAY[s]||s}</button>;
          })}
        </div>
      </Lbl>
      <Lbl t="Atividades / Subitens">
        <div>
          {(f.subitems||[]).map(function(si){
            return (
              <div key={si.id} style={{display:"flex",gap:"6px",alignItems:"center",marginBottom:"4px"}}>
                <span style={{flex:1,fontSize:"12px",padding:"5px 9px",borderRadius:"6px",background:"#f3f4f6",color:"#374151"}}>{si.name}</span>
                <button onClick={function(){delSubitem(si.id);}} style={Object.assign({},B.ghost,{color:"#dc2626",fontSize:"15px",padding:"1px 6px"})}>×</button>
              </div>
            );
          })}
          <div style={{display:"flex",gap:"6px",marginTop:"4px"}}>
            <input placeholder="Ex: Fundações, Estrutura, Detalhamento..." value={newSi} onChange={function(e){setNewSi(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter"){e.preventDefault();addSubitem();}}} style={B.inp}/>
            <button onClick={addSubitem} style={Object.assign({},B.sec,{fontSize:"12px",padding:"6px 12px",flexShrink:0})}>+ Add</button>
          </div>
        </div>
      </Lbl>
      <Lbl t="Localizacao">
        <div style={{display:"grid",gridTemplateColumns:"1fr 90px 100px",gap:"6px"}}>
          <input placeholder="Cidade" value={f.city} onChange={function(e){setF(function(p){return Object.assign({},p,{city:e.target.value});});}} style={B.inp}/>
          <select value={f.st} onChange={function(e){setF(function(p){return Object.assign({},p,{st:e.target.value});});}} style={B.inp}>
            <option value="">UF</option>
            {UFS.map(function(u){return <option key={u} value={u}>{u}</option>;})}
          </select>
          <input placeholder="Pais" value={f.country} onChange={function(e){setF(function(p){return Object.assign({},p,{country:e.target.value});});}} style={B.inp}/>
        </div>
      </Lbl>
      <Lbl t="Status">
        <div style={{display:"flex",gap:"6px",flexWrap:"wrap"}}>
          {["ativo","aguardando","bloqueado"].map(function(s){
            var sc=STATUS[s];var on=f.status===s;
            return <button key={s} onClick={function(){setF(function(p){return Object.assign({},p,{status:s});});}} style={{fontSize:"12px",padding:"5px 12px",borderRadius:"20px",border:"1px solid "+(on?sc.hex:"#e5e7eb"),background:on?sc.bg:"transparent",color:on?sc.color:"#9ca3af",cursor:"pointer",fontWeight:on?600:400}}>{sc.label}</button>;
          })}
        </div>
      </Lbl>
      <Lbl t="Observacao"><textarea placeholder="Bloqueios, pendencias, notas..." value={f.note} onChange={function(e){setF(function(p){return Object.assign({},p,{note:e.target.value});});}} style={Object.assign({},B.inp,{height:"60px",resize:"vertical"})}/></Lbl>
      <div style={{display:"flex",gap:"8px"}}>
        <button onClick={function(){if(f.name.trim()&&f.cId)onSave(f);}} style={B.pri} disabled={!f.name.trim()||!f.cId}>Salvar</button>
        <button onClick={onCancel} style={B.sec}>Cancelar</button>
        {onDelete&&<button onClick={onDelete} style={Object.assign({},B.ghost,{color:"#dc2626",border:"1px solid #fecaca",borderRadius:"6px",padding:"7px 12px",marginLeft:"auto"})}>Excluir</button>}
      </div>
    </div>
  );
}

function ProjDetail(props){
  var proj=props.proj,ctr=props.ctr,colorMap=props.colorMap,diary=props.diary,saveDiary=props.saveDiary,team=props.team,onBack=props.onBack,projects=props.projects,saveProjects=props.saveProjects,history=props.history;
  var entries=(diary[proj.id]||[]).sort(function(a,b){return b.date.localeCompare(a.date);});
  var sfs=useState(false); var showForm=sfs[0],setShowForm=sfs[1];
  var eds=useState(TODAY); var eDate=eds[0],setEDate=eds[1];
  var eas=useState(team[0]?team[0].name:""); var eAuthor=eas[0],setEAuthor=eas[1];
  var ets=useState(""); var eText=ets[0],setEText=ets[1];
  var color=colorMap[proj.id];
  var sc=STATUS[proj.status]||STATUS.ativo;
  var daysWorked=Object.entries(history).filter(function(e){
    var n=normAlloc((e[1].alloc)||{});
    return Object.values(n).some(function(pids){return pids.includes(proj.id);});
  }).length;
  function addEntry(){
    if(!eText.trim())return;
    var entry={id:"e"+uid(),date:eDate,author:eAuthor,text:eText.trim(),createdAt:new Date().toISOString()};
    saveDiary(Object.assign({},diary,{[proj.id]:[...(diary[proj.id]||[]),entry]}));
    setEText("");setShowForm(false);
  }
  function delEntry(id){saveDiary(Object.assign({},diary,{[proj.id]:(diary[proj.id]||[]).filter(function(e){return e.id!==id;})}));}
  return (
    <div>
      <button onClick={onBack} style={Object.assign({},B.ghost,{color:RED,padding:0,fontWeight:600,fontSize:"13px",marginBottom:"1rem"})}>Voltar</button>
      <div style={Object.assign({},B.card,{borderLeft:"4px solid "+color,marginBottom:"1.25rem"})}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",gap:"12px"}}>
          <div>
            <div style={{fontSize:"18px",fontWeight:700,color:"#111"}}>{proj.name}</div>
            {ctr&&<div style={{fontSize:"13px",color:"#9ca3af",marginTop:"2px"}}>{ctr.name}</div>}
            {(proj.city||proj.st)&&<div style={{fontSize:"12px",color:"#9ca3af",marginTop:"2px"}}>{[proj.city,proj.st,proj.country].filter(Boolean).join(", ")}</div>}
          </div>
          <button onClick={function(){saveProjects(projects.map(function(p){return p.id===proj.id?Object.assign({},p,{status:ST_CYC[p.status]||"ativo"}):p;}));}} style={{fontSize:"12px",fontWeight:600,padding:"4px 12px",borderRadius:"20px",background:sc.bg,color:sc.color,border:"none",cursor:"pointer",flexShrink:0}}>
            {sc.label} ↻
          </button>
        </div>
        {proj.svcs&&proj.svcs.length>0&&<div style={{display:"flex",flexWrap:"wrap",gap:"3px",marginTop:"10px"}}>{proj.svcs.map(function(s){return <span key={s} style={{fontSize:"11px",padding:"3px 8px",borderRadius:"4px",background:"#f3f4f6",color:"#6b7280",border:"1px solid #e5e7eb"}}>{SERVICES_DISPLAY[s]||s}</span>;})}</div>}
        {proj.note&&<div style={{fontSize:"12px",color:"#9ca3af",marginTop:"8px",padding:"8px 10px",borderRadius:"6px",background:"#f9fafb",fontStyle:"italic"}}>{proj.note}</div>}
        <div style={{marginTop:"12px",paddingTop:"10px",borderTop:"1px solid #f0f0f0",display:"flex",gap:"20px"}}>
          <div style={{textAlign:"center"}}><div style={{fontSize:"18px",fontWeight:700}}>{daysWorked}</div><div style={{fontSize:"10px",color:"#9ca3af",textTransform:"uppercase",letterSpacing:".05em"}}>dias trabalhados</div></div>
          <div style={{textAlign:"center"}}><div style={{fontSize:"18px",fontWeight:700}}>{entries.length}</div><div style={{fontSize:"10px",color:"#9ca3af",textTransform:"uppercase",letterSpacing:".05em"}}>entradas no diario</div></div>
        </div>
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
        <div style={B.lbl}>Diario de obra</div>
        <button onClick={function(){setShowForm(function(v){return !v;});}} style={Object.assign({},B.pri,{fontSize:"12px",padding:"6px 14px"})}>+ Nova entrada</button>
      </div>
      {showForm&&(
        <div style={{background:"#fff",borderRadius:"10px",padding:"16px",marginBottom:"1rem",border:"1px solid "+RED+"25"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"8px"}}>
            <input type="date" value={eDate} onChange={function(e){setEDate(e.target.value);}} style={B.inp}/>
            <select value={eAuthor} onChange={function(e){setEAuthor(e.target.value);}} style={B.inp}>{team.map(function(m){return <option key={m.id} value={m.name}>{m.name}</option>;})}</select>
          </div>
          <textarea placeholder="O que foi feito, decisoes, pendencias, observacoes..." value={eText} onChange={function(e){setEText(e.target.value);}} style={Object.assign({},B.inp,{height:"80px",resize:"vertical",marginBottom:"8px"})}/>
          <div style={{display:"flex",gap:"8px"}}>
            <button onClick={addEntry} style={B.pri} disabled={!eText.trim()}>Salvar entrada</button>
            <button onClick={function(){setShowForm(false);}} style={B.sec}>Cancelar</button>
          </div>
        </div>
      )}
      {entries.length===0?(
        <div style={{textAlign:"center",padding:"2.5rem",color:"#9ca3af",fontSize:"13px"}}>Nenhuma entrada ainda.</div>
      ):entries.map(function(e){
        return (
          <div key={e.id} style={Object.assign({},B.card,{borderLeft:"3px solid #e5e7eb"})}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"8px"}}>
              <div style={{display:"flex",gap:"8px",alignItems:"center"}}>
                <span style={{fontSize:"12px",fontWeight:700,color:RED}}>{fmtDate(e.date)}</span>
                <span style={{fontSize:"11px",padding:"1px 7px",borderRadius:"10px",background:"#f3f4f6",color:"#6b7280",fontWeight:600}}>{e.author}</span>
              </div>
              <button onClick={function(){delEntry(e.id);}} style={Object.assign({},B.ghost,{fontSize:"16px",opacity:.3,lineHeight:1})}>x</button>
            </div>
            <div style={{fontSize:"13px",color:"#374151",lineHeight:1.6,whiteSpace:"pre-wrap"}}>{e.text}</div>
          </div>
        );
      })}
    </div>
  );
}

function MemberCard(props){
  var m=props.m,rc=props.rc,delMember=props.delMember,saveEmail=props.saveEmail,saveAcesso=props.saveAcesso;
  var ees=useState(false); var editE=ees[0],setEditE=ees[1];
  var evs=useState(m.email||""); var emailVal=evs[0],setEmailVal=evs[1];
  var cms=useState(false); var criarMode=cms[0],setCriarMode=cms[1];
  var cps=useState(""); var criarPass=cps[0],setCriarPass=cps[1];
  var cls=useState(false); var criarLoad=cls[0],setCriarLoad=cls[1];
  var css2=useState(""); var criarStatus=css2[0],setCriarStatus=css2[1];
  var rls=useState(false); var resetLoad=rls[0],setResetLoad=rls[1];
  var rss=useState(""); var resetStatus=rss[0],setResetStatus=rss[1];

  async function handleCriarAcesso(){
    if(!criarPass||criarPass.length<6){setCriarStatus("Mínimo 6 caracteres.");return;}
    setCriarLoad(true);setCriarStatus("");
    try{
      var appName="cria-"+Date.now();
      var sec=initializeApp(firebaseConfig,appName);
      var secAuth=getAuth(sec);
      await createUserWithEmailAndPassword(secAuth,m.email,criarPass);
      await secAuth.signOut();
      await deleteApp(sec);
      setCriarStatus("ok");setCriarMode(false);setCriarPass("");
    }catch(ex){
      var msgs={"auth/email-already-in-use":"Acesso já existe para este e-mail.","auth/invalid-email":"E-mail inválido.","auth/weak-password":"Senha muito fraca (mín. 6 caracteres)."};
      setCriarStatus(msgs[ex.code]||"Erro: "+ex.message);
    }
    setCriarLoad(false);
  }

  async function handleRedefinirSenha(){
    setResetLoad(true);setResetStatus("");
    try{
      var app=getApps().length?getApp():initializeApp(firebaseConfig);
      await sendPasswordResetEmail(getAuth(app),m.email);
      setResetStatus("ok");
    }catch(ex){
      setResetStatus("Erro ao enviar e-mail.");
    }
    setResetLoad(false);
  }

  return (
    <div style={Object.assign({},B.card,{padding:"12px 14px"})}>
      <div style={{display:"flex",alignItems:"center",gap:"12px"}}>
        <div style={{width:"36px",height:"36px",borderRadius:"50%",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",fontSize:"12px",fontWeight:700,background:rc+"15",color:rc,border:"1.5px solid "+rc+"30"}}>{m.name.slice(0,2).toUpperCase()}</div>
        <div style={{flex:1,minWidth:0}}>
          <div style={{fontSize:"14px",fontWeight:600,color:"#111"}}>{m.name}</div>
          <div style={{fontSize:"11px",color:"#9ca3af"}}>{ROLES_DISPLAY[m.role]||m.role}</div>
          {editE?(
            <div style={{display:"flex",gap:"4px",marginTop:"4px"}}>
              <input type="email" value={emailVal} onChange={function(e){setEmailVal(e.target.value);}} placeholder="email@..." style={Object.assign({},B.inp,{fontSize:"11px",padding:"3px 7px",height:"26px"})}/>
              <button onClick={function(){if(saveEmail)saveEmail(m.id,emailVal);setEditE(false);}} style={Object.assign({},B.pri,{fontSize:"11px",padding:"3px 9px"})}>✓</button>
              <button onClick={function(){setEditE(false);setEmailVal(m.email||"");}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"3px 9px"})}>✕</button>
            </div>
          ):(
            <div onClick={function(){setEditE(true);setCriarMode(false);}} style={{fontSize:"11px",color:m.email?"#6b7280":"#d1d5db",marginTop:"2px",cursor:"pointer"}}>
              {m.email||"+ definir e-mail de acesso"}
            </div>
          )}
        </div>
        {saveAcesso&&(
          <div style={{display:"flex",flexDirection:"column",alignItems:"center",gap:"2px",flexShrink:0}}>
            <span style={{fontSize:"9px",color:"#9ca3af",textTransform:"uppercase",letterSpacing:".06em"}}>Acesso</span>
            {ADMIN_EMAILS.includes((m.email||"").toLowerCase())
              ? <span style={{fontSize:"11px",padding:"3px 10px",borderRadius:"20px",background:RED+"15",color:RED,fontWeight:700}}>Admin ★</span>
              : <select value={m.acesso||"usuario"} onChange={function(e){saveAcesso(m.id,e.target.value);}} style={{fontSize:"11px",padding:"3px 7px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"#fff",color:m.acesso==="admin"?RED:m.acesso==="gestor"?"#2563EB":"#9ca3af",fontWeight:m.acesso&&m.acesso!=="usuario"?700:400,cursor:"pointer"}}>
                  <option value="usuario">Usuário</option>
                  <option value="gestor">Gestor</option>
                  <option value="admin">Admin</option>
                </select>
            }
          </div>
        )}
        <button onClick={function(){if(window.confirm("Remover "+m.name+" da equipe?"))delMember(m.id);}} style={{background:"none",border:"1px solid #fecaca",borderRadius:"6px",cursor:"pointer",color:"#dc2626",fontSize:"12px",padding:"3px 9px",fontWeight:600,flexShrink:0}}>Remover</button>
      </div>
      {m.email&&!editE&&(
        <div style={{marginTop:"10px",paddingTop:"10px",borderTop:"1px solid #f0f0f0"}}>
          {criarMode?(
            <div style={{display:"flex",gap:"6px",alignItems:"center",flexWrap:"wrap",marginBottom:"6px"}}>
              <span style={{fontSize:"11px",color:"#9ca3af",flexShrink:0}}>{m.email}</span>
              <input type="password" value={criarPass} onChange={function(e){setCriarPass(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")handleCriarAcesso();}} placeholder="Senha..." style={Object.assign({},B.inp,{fontSize:"11px",padding:"4px 8px",width:"130px"})}/>
              <button onClick={handleCriarAcesso} disabled={criarLoad} style={Object.assign({},B.pri,{fontSize:"11px",padding:"4px 12px"})}>{criarLoad?"Criando...":"Criar"}</button>
              <button onClick={function(){setCriarMode(false);setCriarPass("");setCriarStatus("");}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"4px 10px"})}>Cancelar</button>
            </div>
          ):(
            <div style={{display:"flex",gap:"6px",flexWrap:"wrap",alignItems:"center"}}>
              <button onClick={function(){setCriarMode(true);setCriarStatus("");setResetStatus("");}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"4px 12px"})}>+ Criar acesso</button>
              <button onClick={handleRedefinirSenha} disabled={resetLoad} style={Object.assign({},B.sec,{fontSize:"11px",padding:"4px 12px"})}>{resetLoad?"Enviando...":"Redefinir senha"}</button>
              {criarStatus==="ok"&&<span style={{fontSize:"11px",color:"#15803d"}}>✓ Acesso criado!</span>}
              {criarStatus&&criarStatus!=="ok"&&<span style={{fontSize:"11px",color:"#dc2626"}}>{criarStatus}</span>}
              {resetStatus==="ok"&&<span style={{fontSize:"11px",color:"#15803d"}}>✓ E-mail de redefinição enviado!</span>}
              {resetStatus&&resetStatus!=="ok"&&<span style={{fontSize:"11px",color:"#dc2626"}}>{resetStatus}</span>}
            </div>
          )}
          {criarMode&&criarStatus&&criarStatus!=="ok"&&<div style={{fontSize:"11px",color:"#dc2626",marginTop:"4px"}}>{criarStatus}</div>}
        </div>
      )}
    </div>
  );
}

function EquipeTab(props){
  var team=props.team,saveTeam=props.saveTeam,weights=props.weights,saveWeights=props.saveWeights;
  var sfs=useState(false); var showForm=sfs[0],setShowForm=sfs[1];
  var ns=useState(""); var name=ns[0],setName=ns[1];
  var rs=useState("Arquiteto"); var role=rs[0],setRole=rs[1];
  var ews=useState(false); var editW=ews[0],setEditW=ews[1];
  var wls=useState(Object.assign({},weights)); var wLocal=wls[0],setWLocal=wls[1];
  function addMember(){if(!name.trim())return;saveTeam([...team,{id:"m"+uid(),name:name.trim(),role:role,email:""}]);setName("");setRole("Arquiteto");setShowForm(false);}
  function delMember(id){saveTeam(team.filter(function(m){return m.id!==id;}));}
  function saveEmail(id,email){saveTeam(team.map(function(m){return m.id===id?Object.assign({},m,{email:email.trim().toLowerCase()}):m;}));}
  function saveAcesso(id,acesso){saveTeam(team.map(function(m){return m.id===id?Object.assign({},m,{acesso:acesso}):m;}));}
  return (
    <div>
      <div style={Object.assign({},B.card,{marginBottom:"1.5rem",borderTop:"3px solid "+RED})}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"10px"}}>
          <div>
            <div style={{fontSize:"14px",fontWeight:700,color:"#111"}}>Pesos por cargo</div>
            <div style={{fontSize:"12px",color:"#9ca3af",marginTop:"2px"}}>Contribuicao relativa de cada cargo no calculo de dedicacao</div>
          </div>
          <button onClick={function(){setEditW(function(v){return !v;});setWLocal(Object.assign({},weights));}} style={Object.assign({},B.sec,{fontSize:"12px",padding:"5px 12px"})}>{editW?"Cancelar":"Editar"}</button>
        </div>
        {editW?(
          <div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"10px",marginBottom:"12px"}}>
              {ROLES.map(function(r){
                var rc=ROLE_COLOR[r]||"#888";
                return (
                  <div key={r} style={{display:"flex",alignItems:"center",gap:"10px",padding:"8px 10px",borderRadius:"8px",background:"#f9fafb",border:"1px solid #f0f0f0"}}>
                    <div style={{width:"8px",height:"8px",borderRadius:"50%",background:rc,flexShrink:0}}/>
                    <span style={{fontSize:"13px",fontWeight:600,flex:1}}>{ROLES_DISPLAY[r]}</span>
                    <input type="number" step="0.5" min="0.1" value={wLocal[r]!=null?wLocal[r]:(DW[r]||1)} onChange={function(e){var val=+e.target.value;setWLocal(function(p){return Object.assign({},p,{[r]:val});});}} style={Object.assign({},B.inp,{width:"65px",textAlign:"center",fontWeight:700,fontSize:"14px",padding:"4px"})}/>
                  </div>
                );
              })}
            </div>
            <button onClick={function(){saveWeights(wLocal);setEditW(false);}} style={B.pri}>Salvar pesos</button>
          </div>
        ):(
          <div style={{display:"flex",gap:"8px",flexWrap:"wrap"}}>
            {ROLES.map(function(r){
              var rc=ROLE_COLOR[r]||"#888";
              return (
                <div key={r} style={{display:"flex",alignItems:"center",gap:"6px",padding:"6px 12px",borderRadius:"20px",background:rc+"15",border:"1px solid "+rc+"25"}}>
                  <span style={{fontSize:"12px",fontWeight:600,color:rc}}>{ROLES_DISPLAY[r]}</span>
                  <span style={{fontSize:"14px",fontWeight:800,color:"#111"}}>{weights[r]!=null?weights[r]:(DW[r]||1)}</span>
                </div>
              );
            })}
          </div>
        )}
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1.25rem"}}>
        <span style={{fontSize:"13px",color:"#9ca3af"}}>{team.length} pessoa{team.length!==1?"s":""}</span>
        <button onClick={function(){setShowForm(function(v){return !v;});}} style={B.pri}>+ Adicionar pessoa</button>
      </div>
      {showForm&&(
        <div style={{background:"#fff",borderRadius:"10px",padding:"16px",marginBottom:"1.25rem",border:"1px solid "+RED+"25"}}>
          <div style={{fontSize:"13px",fontWeight:700,color:"#111",marginBottom:"12px"}}>Nova pessoa</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 160px",gap:"8px",marginBottom:"10px"}}>
            <input autoFocus placeholder="Nome completo" value={name} onChange={function(e){setName(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")addMember();}} style={B.inp}/>
            <select value={role} onChange={function(e){setRole(e.target.value);}} style={B.inp}>{ROLES.map(function(r){return <option key={r} value={r}>{ROLES_DISPLAY[r]}</option>;})}</select>
          </div>
          <div style={{display:"flex",gap:"8px"}}>
            <button onClick={addMember} style={B.pri}>Adicionar</button>
            <button onClick={function(){setShowForm(false);}} style={B.sec}>Cancelar</button>
          </div>
        </div>
      )}
      {ROLES.map(function(role){
        var members=team.filter(function(m){return m.role===role;});
        if(!members.length)return null;
        var rc=ROLE_COLOR[role]||"#888";
        return (
          <div key={role} style={{marginBottom:"1.5rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"10px"}}>
              <div style={{width:"3px",height:"16px",borderRadius:"2px",background:rc}}/>
              <span style={{fontSize:"12px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#374151"}}>{ROLES_DISPLAY[role]}</span>
              <span style={{fontSize:"11px",padding:"1px 7px",borderRadius:"10px",background:rc+"15",color:rc,fontWeight:700}}>peso {weights[role]!=null?weights[role]:(DW[role]||1)}</span>
            </div>
            {members.map(function(m){
              return <MemberCard key={m.id} m={m} rc={rc} weights={weights} delMember={delMember} saveEmail={saveEmail} saveAcesso={saveAcesso}/>;
            })}
          </div>
        );
      })}
      {(function(){
        var outros=team.filter(function(m){return !ROLES.includes(m.role);});
        if(!outros.length)return null;
        return (
          <div style={{marginBottom:"1.5rem"}}>
            <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"10px"}}>
              <div style={{width:"3px",height:"16px",borderRadius:"2px",background:"#9ca3af"}}/>
              <span style={{fontSize:"12px",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em",color:"#374151"}}>Outros</span>
              <span style={{fontSize:"11px",padding:"1px 7px",borderRadius:"10px",background:"#f3f4f6",color:"#9ca3af",fontWeight:700}}>{outros.length} cadastro{outros.length!==1?"s":""}</span>
            </div>
            {outros.map(function(m){return <MemberCard key={m.id} m={m} rc="#9ca3af" weights={weights} delMember={delMember} saveEmail={saveEmail}/>;}) }
          </div>
        );
      })()}
    </div>
  );
}

function buildCalendar(monthStr){
  var pts=monthStr.split('-');
  var yr=parseInt(pts[0]),mo=parseInt(pts[1]);
  var dim=new Date(yr,mo,0).getDate();
  var weeks=[],week=null;
  for(var d=1;d<=dim;d++){
    var ds=pts[0]+'-'+pts[1]+'-'+String(d).padStart(2,'0');
    var dow=new Date(ds+'T12:00:00').getDay();
    if(dow===0||dow===6)continue;
    var idx=dow-1;
    if(week===null||idx===0){week=[null,null,null,null,null];weeks.push(week);}
    week[idx]=ds;
  }
  return weeks;
}

function PlanTab(props){
  var team=props.team,planning=props.planning,savePlanning=props.savePlanning;
  var todayMonth=TODAY.slice(0,7);
  var ms=useState(todayMonth); var month=ms[0],setMonth=ms[1];
  var sds=useState(null); var selDate=sds[0],setSelDate=sds[1];
  var smbs=useState(""); var selMemberId=smbs[0],setSelMemberId=smbs[1];
  var nes=useState({}); var noteEdits=nes[0],setNoteEdits=nes[1];

  var weeks=buildCalendar(month);
  var mParts=month.split('-');
  var monthLabel=MONTH_PT[parseInt(mParts[1])-1]+' '+mParts[0];

  function prevMonth(){
    var d=new Date(parseInt(mParts[0]),parseInt(mParts[1])-2,1);
    setMonth(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));
    setSelDate(null);
  }
  function nextMonth(){
    var d=new Date(parseInt(mParts[0]),parseInt(mParts[1]),1);
    setMonth(d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'));
    setSelDate(null);
  }

  var dayEntries=selDate?(planning[selDate]||{}):{};

  function setShift(memberId,shift){
    var newDay=Object.assign({},planning[selDate]||{});
    if(!shift){
      delete newDay[memberId];
    } else {
      newDay[memberId]=Object.assign({},newDay[memberId]||{},{shift:shift});
    }
    var np=Object.assign({},planning,{[selDate]:newDay});
    if(!Object.keys(newDay).length)delete np[selDate];
    savePlanning(np);
  }

  function saveNote(memberId){
    var key=selDate+'_'+memberId;
    var val=noteEdits[key]!=null?noteEdits[key]:((planning[selDate]||{})[memberId]||{}).note||"";
    var newDay=Object.assign({},planning[selDate]||{});
    newDay[memberId]=Object.assign({},newDay[memberId]||{},{note:val});
    savePlanning(Object.assign({},planning,{[selDate]:newDay}));
  }

  var memberPlanDays=[];
  if(selMemberId){
    Object.keys(planning).filter(function(d){return d.startsWith(month);}).sort().forEach(function(d){
      var entry=(planning[d]||{})[selMemberId];
      if(entry&&entry.shift)memberPlanDays.push({date:d,shift:entry.shift,note:entry.note||""});
    });
  }

  var DAY_LABELS=["Seg","Ter","Qua","Qui","Sex"];

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:"10px",marginBottom:"1.25rem"}}>
        <button onClick={prevMonth} style={Object.assign({},B.sec,{padding:"6px 14px",fontSize:"16px"})}>‹</button>
        <span style={{fontSize:"15px",fontWeight:700,color:"#111",flex:1,textAlign:"center"}}>{monthLabel}</span>
        <button onClick={nextMonth} style={Object.assign({},B.sec,{padding:"6px 14px",fontSize:"16px"})}>›</button>
      </div>

      <div style={Object.assign({},B.card,{padding:"10px",marginBottom:"1rem",overflowX:"auto"})}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:"4px",marginBottom:"6px",minWidth:"340px"}}>
          {DAY_LABELS.map(function(l){
            return <div key={l} style={{textAlign:"center",fontSize:"10px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".06em",padding:"4px"}}>{l}</div>;
          })}
        </div>
        {weeks.map(function(week,wi){
          return (
            <div key={wi} style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:"4px",marginBottom:"4px",minWidth:"340px"}}>
              {week.map(function(date,di){
                if(!date)return <div key={di} style={{minHeight:"62px",background:"#f9fafb",borderRadius:"6px",border:"1px solid #f0f0f0"}}/>;
                var isToday=date===TODAY;
                var isSel=date===selDate;
                var entries=planning[date]||{};
                var assigned=Object.keys(entries).filter(function(mid){return entries[mid]&&entries[mid].shift&&entries[mid].shift!=="falta";});
                return (
                  <div key={date} onClick={function(){setSelDate(isSel?null:date);}} style={{minHeight:"62px",background:isSel?"#fef2f2":isToday?"#fff7ed":"#fff",borderRadius:"6px",border:"1px solid "+(isSel?RED:isToday?"#fed7aa":"#e5e7eb"),cursor:"pointer",padding:"6px",boxSizing:"border-box"}}>
                    <div style={{fontSize:"12px",fontWeight:isToday?700:500,color:isToday?RED:"#374151",marginBottom:"3px"}}>{parseInt(date.slice(-2))}</div>
                    <div style={{display:"flex",flexWrap:"wrap",gap:"2px"}}>
                      {assigned.slice(0,3).map(function(mid){
                        var m=team.find(function(x){return x.id===mid;});
                        var sh=entries[mid].shift;
                        var sc=SHIFT_COLORS[sh]||"#888";
                        return <span key={mid} style={{fontSize:"8px",fontWeight:700,padding:"1px 3px",borderRadius:"3px",background:SHIFT_BG[sh]||"#f3f4f6",color:sc,border:"1px solid "+sc+"30",lineHeight:1.6}}>{m?m.name.split(' ')[0].slice(0,5):mid}</span>;
                      })}
                      {assigned.length>3&&<span style={{fontSize:"8px",color:"#9ca3af"}}>+{assigned.length-3}</span>}
                    </div>
                  </div>
                );
              })}
            </div>
          );
        })}
      </div>

      {selDate&&(
        <div style={Object.assign({},B.card,{marginBottom:"1rem",borderTop:"3px solid "+RED})}>
          <div style={{fontSize:"13px",fontWeight:700,color:"#111",marginBottom:"12px",paddingBottom:"8px",borderBottom:"1px solid #f0f0f0"}}>
            {fmtDate(selDate)} — {cap(fmtDay(selDate))}
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:"6px"}}>
            {team.map(function(m){
              var entry=dayEntries[m.id]||{};
              var hasShift=!!entry.shift;
              var rc=ROLE_COLOR[m.role]||"#888";
              var noteKey=selDate+'_'+m.id;
              var noteVal=noteEdits[noteKey]!=null?noteEdits[noteKey]:(entry.note||"");
              return (
                <div key={m.id} style={{display:"flex",alignItems:"flex-start",gap:"8px",padding:"8px 10px",borderRadius:"8px",background:hasShift?"#fef2f2":"#f9fafb",border:"1px solid "+(hasShift?RED+"25":"#f0f0f0"),flexWrap:"wrap"}}>
                  <div style={{display:"flex",alignItems:"center",gap:"8px",flex:1,minWidth:"160px"}}>
                    <div style={{width:"26px",height:"26px",borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:"10px",fontWeight:700,background:rc+"15",color:rc,flexShrink:0}}>{m.name.slice(0,2).toUpperCase()}</div>
                    <span style={{fontSize:"13px",fontWeight:500,color:"#111"}}>{m.name}</span>
                  </div>
                  <div style={{display:"flex",gap:"4px",flexWrap:"wrap"}}>
                    {Object.keys(SHIFTS).map(function(sh){
                      var on=entry.shift===sh;
                      var sc=SHIFT_COLORS[sh];
                      return (
                        <button key={sh} onClick={function(){setShift(m.id,on?null:sh);}} style={{fontSize:"11px",fontWeight:on?700:400,padding:"3px 9px",borderRadius:"12px",border:"1px solid "+(on?sc:"#e5e7eb"),background:on?SHIFT_BG[sh]:"transparent",color:on?sc:"#9ca3af",cursor:"pointer"}}>
                          {SHIFTS[sh]}
                        </button>
                      );
                    })}
                  </div>
                  {hasShift&&(
                    <div style={{display:"flex",gap:"4px",width:"100%",marginTop:"4px"}}>
                      <input placeholder="Nota (motivo de alteração, falta...)" value={noteVal} onChange={function(e){var k=selDate+'_'+m.id;setNoteEdits(function(p){return Object.assign({},p,{[k]:e.target.value});});}} style={Object.assign({},B.inp,{fontSize:"11px",padding:"4px 8px",flex:1})}/>
                      <button onClick={function(){saveNote(m.id);}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"4px 10px",flexShrink:0})}>Salvar</button>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      <div style={B.card}>
        <div style={B.lbl}>Resumo por colaborador — {monthLabel}</div>
        <select value={selMemberId} onChange={function(e){setSelMemberId(e.target.value);}} style={Object.assign({},B.inp,{marginBottom:"12px"})}>
          <option value="">Selecionar colaborador...</option>
          {team.map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
        </select>
        {selMemberId&&(
          memberPlanDays.length===0?(
            <div style={{textAlign:"center",padding:"1.5rem",color:"#9ca3af",fontSize:"13px"}}>Sem planejamento para este mês.</div>
          ):(
            <div>
              <div style={{fontSize:"12px",color:"#6b7280",marginBottom:"10px"}}>{memberPlanDays.length} dia{memberPlanDays.length!==1?"s":""} planejado{memberPlanDays.length!==1?"s":""}</div>
              {memberPlanDays.map(function(item){
                var sc=SHIFT_COLORS[item.shift]||"#888";
                var sb=SHIFT_BG[item.shift]||"#f3f4f6";
                return (
                  <div key={item.date} style={{display:"flex",alignItems:"center",gap:"10px",padding:"7px 0",borderBottom:"1px solid #f0f0f0",flexWrap:"wrap"}}>
                    <span style={{fontSize:"12px",fontWeight:600,color:"#374151",minWidth:"95px"}}>{fmtDate(item.date)}</span>
                    <span style={{fontSize:"11px",padding:"2px 9px",borderRadius:"10px",background:sb,color:sc,fontWeight:600}}>{SHIFTS[item.shift]}</span>
                    {item.note&&<span style={{fontSize:"11px",color:"#9ca3af",fontStyle:"italic"}}>{item.note}</span>}
                  </div>
                );
              })}
            </div>
          )
        )}
      </div>
    </div>
  );
}

var TASK_ST={
  pendente:    {label:"Pendente",    bg:"#fef9c3",color:"#a16207"},
  em_andamento:{label:"Em andamento",bg:"#dcfce7",color:"#15803d"},
  pausado:     {label:"Pausado",     bg:"#fee2e2",color:"#dc2626"},
  concluido:   {label:"Concluído",   bg:"#f3f4f6",color:"#6b7280"},
};

function TarefasTab(props){
  var tarefas=props.tarefas,saveTarefas=props.saveTarefas,team=props.team,projects=props.projects,authUser=props.authUser,currentMember=props.currentMember;
  var ps=useState(""); var pessoa=ps[0],setPessoa=ps[1];
  var prs=useState(""); var projeto=prs[0],setProjeto=prs[1];
  var ats=useState(""); var atividade=ats[0],setAtividade=ats[1];
  var ds=useState(""); var desc=ds[0],setDesc=ds[1];
  var errs=useState(""); var err=errs[0],setErr=errs[1];
  var dels=useState(null); var delegando=dels[0],setDelegando=dels[1];
  var faps=useState(""); var fAtivaPessoa=faps[0],setFAtivaPessoa=faps[1];
  var faprs=useState(""); var fAtivaProj=faprs[0],setFAtivaProj=faprs[1];
  var fcps=useState(""); var fConcPessoa=fcps[0],setFConcPessoa=fcps[1];
  var fcprs=useState(""); var fConcProj=fcprs[0],setFConcProj=fcprs[1];
  var projSel=projects.find(function(p){return p.id===projeto;});
  var subitens=(projSel&&projSel.subitems)||[];

  function addTarefa(){
    if(!pessoa||!projeto||!desc.trim()){setErr("Preencha todos os campos.");return;}
    setErr("");
    var membro=team.find(function(m){return m.id===pessoa;});
    var proj=projects.find(function(p){return p.id===projeto;});
    var adminNome=ADMIN_NAMES[(authUser?authUser.email.toLowerCase():"")] || (authUser?authUser.email.split("@")[0]:"Admin");
    var atNome=atividade?(subitens.find(function(s){return s.id===atividade;})||{}).name||"":"";
    var nova={id:uid(),membroId:pessoa,membroNome:membro?membro.name:"",projetoId:projeto,projetoNome:proj?proj.name:"",atividadeId:atividade||null,atividadeNome:atNome,descricao:desc.trim(),status:"pendente",lida:false,criadoEm:new Date().toISOString(),atribuidoPor:adminNome,criadoPorEmail:authUser?authUser.email:"",iniciadoEm:null,pausadoEm:null,concluidoEm:null};
    saveTarefas([...tarefas,nova]);
    setPessoa("");setProjeto("");setAtividade("");setDesc("");
  }

  function changeStatus(taskId,newStatus){
    var now=new Date().toISOString();
    saveTarefas(tarefas.map(function(t){
      if(t.id!==taskId)return t;
      return Object.assign({},t,{
        status:newStatus,
        iniciadoEm:newStatus==="em_andamento"&&!t.iniciadoEm?now:t.iniciadoEm,
        pausadoEm:newStatus==="pausado"?now:t.pausadoEm,
        concluidoEm:newStatus==="concluido"?now:t.concluidoEm,
      });
    }));
  }

  function getActions(status){
    if(status==="pendente")    return [{label:"Iniciar",icon:"▷",next:"em_andamento"}];
    if(status==="em_andamento")return [{label:"Pausar",icon:"⏸",next:"pausado"},{label:"Concluir",icon:"✓",next:"concluido"}];
    if(status==="pausado")     return [{label:"Retomar",icon:"▷",next:"em_andamento"},{label:"Concluir",icon:"✓",next:"concluido"}];
    return [];
  }

  var ativas=tarefas.filter(function(t){return t.status!=="concluido";});
  var concluidas=tarefas.filter(function(t){return t.status==="concluido";});
  var ativasFiltradas=ativas.filter(function(t){return(!fAtivaProj||t.projetoNome===fAtivaProj)&&(!fAtivaPessoa||t.membroNome===fAtivaPessoa);});

  return (
    <div>
      <div style={B.lbl}>Atribuir nova tarefa</div>
      <div style={Object.assign({},B.card,{marginBottom:"1.5rem"})}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"8px"}}>
          <select value={pessoa} onChange={function(e){setPessoa(e.target.value);}} style={B.inp}>
            <option value="">Colaborador...</option>
            {team.map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
          </select>
          <select value={projeto} onChange={function(e){setProjeto(e.target.value);setAtividade("");}} style={B.inp}>
            <option value="">Projeto...</option>
            {projects.filter(function(p){return p.status!=="concluido";}).map(function(p){return <option key={p.id} value={p.id}>{p.name}</option>;})}
          </select>
        </div>
        {subitens.length>0&&(
          <div style={{marginBottom:"8px"}}>
            <select value={atividade} onChange={function(e){setAtividade(e.target.value);}} style={B.inp}>
              <option value="">Atividade (opcional)...</option>
              {subitens.map(function(si){return <option key={si.id} value={si.id}>{si.name}</option>;})}
            </select>
          </div>
        )}
        <div style={{display:"flex",gap:"8px"}}>
          <input type="text" value={desc} onChange={function(e){setDesc(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")addTarefa();}} placeholder="Descrição da tarefa..." style={Object.assign({},B.inp,{flex:1})}/>
          <button onClick={addTarefa} style={B.pri}>Atribuir</button>
        </div>
        {err&&<div style={{fontSize:"12px",color:"#dc2626",marginTop:"6px"}}>{err}</div>}
      </div>

      <div style={B.lbl}>Todas as tarefas ativas</div>
      <div style={{display:"flex",gap:"8px",marginBottom:"10px",flexWrap:"wrap"}}>
        <select value={fAtivaProj} onChange={function(e){setFAtivaProj(e.target.value);}} style={Object.assign({},B.inp,{flex:"1",minWidth:"140px",fontSize:"12px"})}>
          <option value="">Todos os projetos</option>
          {[...new Set(ativas.map(function(t){return t.projetoNome;}).filter(Boolean))].sort().map(function(p){return <option key={p} value={p}>{p}</option>;})}
        </select>
        <select value={fAtivaPessoa} onChange={function(e){setFAtivaPessoa(e.target.value);}} style={Object.assign({},B.inp,{flex:"1",minWidth:"140px",fontSize:"12px"})}>
          <option value="">Todos os membros</option>
          {[...new Set(ativas.map(function(t){return t.membroNome;}).filter(Boolean))].sort().map(function(p){return <option key={p} value={p}>{p}</option>;})}
        </select>
      </div>
      {ativasFiltradas.length===0?(
        <div style={{textAlign:"center",padding:"2.5rem",color:"#9ca3af",fontSize:"13px"}}>{ativas.length===0?"Nenhuma tarefa ativa no momento.":"Nenhuma tarefa corresponde ao filtro."}</div>
      ):(
        ativasFiltradas.map(function(t){
          var st=TASK_ST[t.status]||TASK_ST.pendente;
          var isDel=delegando&&delegando.taskId===t.id;
          var isOwn=!!(currentMember&&t.membroId===currentMember.id);
          var acts=isOwn?getActions(t.status):[];
          return (
            <div key={t.id} style={Object.assign({},B.card,{padding:"10px 14px"})}>
              <div style={{display:"flex",alignItems:"center",gap:"12px",flexWrap:"wrap"}}>
                <div style={{fontWeight:700,fontSize:"14px",color:"#111",minWidth:"80px",flexShrink:0}}>{t.membroNome}</div>
                <div style={{fontSize:"13px",color:"#9ca3af",flexShrink:0,minWidth:"130px"}}>{t.projetoNome}{t.atividadeNome?" · "+t.atividadeNome:""}</div>
                <div style={{fontSize:"13px",color:"#374151",flex:1,minWidth:"100px"}}>{t.descricao}</div>
                <span style={{fontSize:"11px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:st.bg,color:st.color,flexShrink:0}}>{st.label}</span>
                {isDel?(
                  <div style={{display:"flex",gap:"5px",alignItems:"center",flexShrink:0}}>
                    <select value={delegando.novoMembro} onChange={function(e){setDelegando({taskId:t.id,novoMembro:e.target.value});}} style={Object.assign({},B.inp,{fontSize:"11px",padding:"3px 8px",width:"130px"})}>
                      <option value="">Para quem?</option>
                      {team.filter(function(m){return m.id!==t.membroId;}).map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
                    </select>
                    <button onClick={function(){
                      if(!delegando.novoMembro)return;
                      var nm=team.find(function(m){return m.id===delegando.novoMembro;});
                      saveTarefas(tarefas.map(function(x){
                        if(x.id!==t.id)return x;
                        return Object.assign({},x,{membroId:delegando.novoMembro,membroNome:nm?nm.name:"",lida:false,delegadoHistorico:[...(x.delegadoHistorico||[]),{de:x.membroNome,para:nm?nm.name:"",quando:new Date().toISOString()}]});
                      }));
                      setDelegando(null);
                    }} disabled={!delegando.novoMembro} style={Object.assign({},B.pri,{fontSize:"11px",padding:"3px 10px"})}>✓</button>
                    <button onClick={function(){setDelegando(null);}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"3px 10px"})}>✕</button>
                  </div>
                ):(
                  <button onClick={function(){setDelegando({taskId:t.id,novoMembro:"",});}} style={Object.assign({},B.ghost,{color:"#2563EB",fontSize:"12px",flexShrink:0,border:"1px solid #dbeafe",borderRadius:"6px",padding:"2px 8px"})}>Delegar</button>
                )}
                <button onClick={function(){if(window.confirm("Remover tarefa?"))saveTarefas(tarefas.filter(function(x){return x.id!==t.id;}));}} style={Object.assign({},B.ghost,{color:"#dc2626",fontSize:"18px",flexShrink:0,lineHeight:1})}>×</button>
              </div>
              {acts.length>0&&(
                <div style={{display:"grid",gridTemplateColumns:"repeat("+acts.length+",1fr)",gap:"6px",marginTop:"8px"}}>
                  {acts.map(function(a){return <button key={a.next} onClick={function(){changeStatus(t.id,a.next);}} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"7px",padding:"7px",fontSize:"12px",color:"#374151",cursor:"pointer",textAlign:"center",fontWeight:500}}>{a.icon} {a.label}</button>;})}
                </div>
              )}
              {t.delegadoHistorico&&t.delegadoHistorico.length>0&&(
                <div style={{marginTop:"6px",paddingTop:"6px",borderTop:"1px solid #f0f0f0",fontSize:"11px",color:"#9ca3af"}}>
                  {t.delegadoHistorico.map(function(h,i){return <span key={i} style={{marginRight:"10px"}}>↳ {h.de} → {h.para} ({new Date(h.quando).toLocaleDateString("pt-BR")})</span>;})}
                </div>
              )}
            </div>
          );
        })
      )}
      {concluidas.length>0&&(
        <div style={{marginTop:"1.25rem",borderTop:"1px solid #f0f0f0",paddingTop:"1rem"}}>
          <div style={Object.assign({},B.lbl,{cursor:"pointer"})} onClick={function(){var el=document.getElementById("conc-list");if(el)el.style.display=el.style.display==="none"?"block":"none";}}>
            Concluídas ({concluidas.length}) ▾
          </div>
          <div id="conc-list" style={{display:"none"}}>
            <div style={{display:"flex",gap:"8px",marginBottom:"10px",marginTop:"8px"}}>
              <select value={fConcProj} onChange={function(e){setFConcProj(e.target.value);}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px"})}>
                <option value="">Todos os projetos</option>
                {[...new Set(concluidas.map(function(t){return t.projetoNome;}))].map(function(p){return <option key={p} value={p}>{p}</option>;}) }
              </select>
              <select value={fConcPessoa} onChange={function(e){setFConcPessoa(e.target.value);}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px"})}>
                <option value="">Todas as pessoas</option>
                {[...new Set(concluidas.map(function(t){return t.membroNome;}))].map(function(p){return <option key={p} value={p}>{p}</option>;}) }
              </select>
            </div>
            {concluidas.filter(function(t){return(!fConcProj||t.projetoNome===fConcProj)&&(!fConcPessoa||t.membroNome===fConcPessoa);}).map(function(t){
              var pessoas=[...(t.delegadoHistorico||[]).map(function(d){return d.de;}),t.membroNome].filter(function(v,idx,a){return a.indexOf(v)===idx;});
              return (
                <div key={t.id} style={Object.assign({},B.card,{padding:"10px 14px"})}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:"6px"}}>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"4px",flexWrap:"wrap"}}>
                        <span style={{fontWeight:700,fontSize:"13px",color:"#111"}}>{pessoas.join(" → ")}</span>
                        <span style={{fontSize:"10px",fontWeight:600,padding:"2px 8px",borderRadius:"20px",background:"#f3f4f6",color:"#6b7280"}}>Concluído</span>
                      </div>
                      <div style={{fontSize:"12px",color:"#9ca3af",marginBottom:"6px"}}>{t.projetoNome}{t.atividadeNome?" · "+t.atividadeNome:""}{t.descricao?" — "+t.descricao:""}</div>
                      <div style={{display:"flex",gap:"14px",flexWrap:"wrap"}}>
                        {t.iniciadoEm&&<span style={{fontSize:"11px",color:"#6b7280"}}>▶ Início: {new Date(t.iniciadoEm).toLocaleDateString("pt-BR")}</span>}
                        {t.concluidoEm&&<span style={{fontSize:"11px",color:"#15803d",fontWeight:600}}>✓ Conclusão: {new Date(t.concluidoEm).toLocaleDateString("pt-BR")}</span>}
                        {(t.delegadoHistorico||[]).map(function(h,hi){return <span key={hi} style={{fontSize:"11px",color:"#9ca3af"}}>↳ {h.de} → {h.para} ({new Date(h.quando).toLocaleDateString("pt-BR")})</span>;})}
                      </div>
                    </div>
                    <button onClick={function(){saveTarefas(tarefas.filter(function(x){return x.id!==t.id;}));}} style={Object.assign({},B.ghost,{color:"#dc2626",fontSize:"18px",lineHeight:1,flexShrink:0})}>×</button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

function UsuarioApp(props){
  var authUser=props.authUser,handleLogout=props.handleLogout,tarefas=props.tarefas,currentMember=props.currentMember,saveTarefas=props.saveTarefas,team=props.team||[];
  var ts=useState("minhas"); var tab=ts[0],setTab=ts[1];
  var sms=useState(false); var senhaMode=sms[0],setSenhaMode=sms[1];
  var udels=useState(null); var uDelegando=udels[0],setUDelegando=udels[1];
  var ucfp=useState(""); var uConcFiltProj=ucfp[0],setUConcFiltProj=ucfp[1];
  var sas=useState(""); var senhaAtual=sas[0],setSenhaAtual=sas[1];
  var sns=useState(""); var senhaNova=sns[0],setSenhaNova=sns[1];
  var scs=useState(""); var senhaConf=scs[0],setSenhaConf=scs[1];
  var sls=useState(false); var senhaLoad=sls[0],setSenhaLoad=sls[1];
  var sss=useState(""); var senhaStatus=sss[0],setSenhaStatus=sss[1];

  async function handleAlterarSenha(){
    if(!senhaAtual||!senhaNova||!senhaConf){setSenhaStatus("Preencha todos os campos.");return;}
    if(senhaNova!==senhaConf){setSenhaStatus("A nova senha e a confirmação não coincidem.");return;}
    if(senhaNova.length<6){setSenhaStatus("Senha mínima de 6 caracteres.");return;}
    setSenhaLoad(true);setSenhaStatus("");
    try{
      var app=getApps().length?getApp():initializeApp(firebaseConfig);
      var auth=getAuth(app);
      var credential=EmailAuthProvider.credential(authUser.email,senhaAtual);
      await reauthenticateWithCredential(auth.currentUser,credential);
      await updatePassword(auth.currentUser,senhaNova);
      setSenhaStatus("ok");setSenhaAtual("");setSenhaNova("");setSenhaConf("");
      setTimeout(function(){setSenhaMode(false);setSenhaStatus("");},2000);
    }catch(ex){
      var msgs={"auth/wrong-password":"Senha atual incorreta.","auth/invalid-credential":"Senha atual incorreta.","auth/weak-password":"Nova senha muito fraca."};
      setSenhaStatus(msgs[ex.code]||"Erro ao alterar senha.");
    }
    setSenhaLoad(false);
  }

  var minhasTarefas=currentMember?tarefas.filter(function(t){return t.membroId===currentMember.id;}):[];
  var naoLidas=minhasTarefas.filter(function(t){return !t.lida&&t.status==="pendente";});
  var ativas=minhasTarefas.filter(function(t){return t.status!=="concluido";});
  var concluidas=minhasTarefas.filter(function(t){return t.status==="concluido";});

  function marcarLidas(){
    var updated=tarefas.map(function(t){return (currentMember&&t.membroId===currentMember.id&&!t.lida)?Object.assign({},t,{lida:true}):t;});
    saveTarefas(updated);
  }

  function changeStatus(taskId,newStatus){
    var now=new Date().toISOString();
    var updated=tarefas.map(function(t){
      if(t.id!==taskId)return t;
      return Object.assign({},t,{
        status:newStatus,
        iniciadoEm:newStatus==="em_andamento"&&!t.iniciadoEm?now:t.iniciadoEm,
        pausadoEm:newStatus==="pausado"?now:t.pausadoEm,
        concluidoEm:newStatus==="concluido"?now:t.concluidoEm,
      });
    });
    saveTarefas(updated);
  }

  function getActions(status){
    if(status==="pendente")    return [{label:"Iniciar",icon:"▷",next:"em_andamento"}];
    if(status==="em_andamento")return [{label:"Pausar",icon:"⏸",next:"pausado"},{label:"Concluir",icon:"✓",next:"concluido"}];
    if(status==="pausado")     return [{label:"Retomar",icon:"▷",next:"em_andamento"},{label:"Concluir",icon:"✓",next:"concluido"}];
    return [];
  }

  return (
    <div style={{background:"#f9fafb",minHeight:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>
      <div style={{height:"3px",background:RED}}/>
      <div style={{background:"#fff",borderBottom:"1px solid #f0f0f0",padding:"0 1.25rem",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0 0"}}>
          <div style={{display:"flex",alignItems:"baseline",gap:"10px"}}>
            <span style={{fontSize:"22px",fontWeight:800,letterSpacing:"-1px",color:"#111"}}>M<span style={{color:RED}}>+</span>M</span>
            <span style={{fontSize:"11px",color:"#9ca3af",letterSpacing:".05em",textTransform:"uppercase",fontWeight:500}}>Arquitetura</span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
            <span style={{fontSize:"11px",color:"#9ca3af",maxWidth:"120px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.email}</span>
            <button onClick={function(){setSenhaMode(function(v){return !v;});setSenhaStatus("");}} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",cursor:"pointer"}}>Senha</button>
            <button onClick={handleLogout} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",cursor:"pointer"}}>Sair</button>
          </div>
        </div>
        <div style={{display:"flex",marginTop:"4px"}}>
          {[["minhas","Minhas tarefas"],["historico","Meu histórico"]].map(function(t){
            return (
              <button key={t[0]} onClick={function(){setTab(t[0]);}} style={{background:"transparent",border:"none",cursor:"pointer",padding:"7px 12px",fontSize:"12px",fontWeight:tab===t[0]?700:400,color:tab===t[0]?"#111":"#9ca3af",borderBottom:tab===t[0]?"2.5px solid "+RED:"2.5px solid transparent",whiteSpace:"nowrap"}}>
                {t[1]}
              </button>
            );
          })}
        </div>
      </div>

      <div style={{padding:"1.25rem",maxWidth:"640px",margin:"0 auto"}}>
        {senhaMode&&(
          <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"1rem 1.25rem",marginBottom:"1.25rem"}}>
            <div style={{fontSize:"14px",fontWeight:700,color:"#111",marginBottom:"12px"}}>Alterar senha</div>
            <div style={{display:"flex",flexDirection:"column",gap:"8px",marginBottom:"12px"}}>
              <input type="password" value={senhaAtual} onChange={function(e){setSenhaAtual(e.target.value);}} placeholder="Senha atual" style={B.inp}/>
              <input type="password" value={senhaNova} onChange={function(e){setSenhaNova(e.target.value);}} placeholder="Nova senha (mín. 6 caracteres)" style={B.inp}/>
              <input type="password" value={senhaConf} onChange={function(e){setSenhaConf(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")handleAlterarSenha();}} placeholder="Confirmar nova senha" style={B.inp}/>
            </div>
            {senhaStatus&&senhaStatus!=="ok"&&<div style={{fontSize:"12px",color:"#dc2626",marginBottom:"8px"}}>{senhaStatus}</div>}
            {senhaStatus==="ok"&&<div style={{fontSize:"12px",color:"#15803d",marginBottom:"8px"}}>✓ Senha alterada com sucesso!</div>}
            <div style={{display:"flex",gap:"8px"}}>
              <button onClick={handleAlterarSenha} disabled={senhaLoad} style={B.pri}>{senhaLoad?"Salvando...":"Salvar senha"}</button>
              <button onClick={function(){setSenhaMode(false);setSenhaAtual("");setSenhaNova("");setSenhaConf("");setSenhaStatus("");}} style={B.sec}>Cancelar</button>
            </div>
          </div>
        )}
        {!currentMember&&(
          <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:"12px",padding:"1.25rem",textAlign:"center",color:"#dc2626",fontSize:"13px",lineHeight:1.7}}>
            Seu e-mail não está vinculado a nenhum membro da equipe.<br/>Peça ao administrador para configurar seu e-mail na aba Equipe.
          </div>
        )}

        {naoLidas.length>0&&(
          <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.45)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <div style={{background:"#fff",borderRadius:"16px",padding:"1.5rem",maxWidth:"360px",width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.3)"}}>
              <div style={{textAlign:"center",marginBottom:"1rem"}}>
                <div style={{fontSize:"36px",marginBottom:"8px"}}>🔔</div>
                <div style={{fontSize:"16px",fontWeight:700,color:"#111"}}>Nova{naoLidas.length!==1?"s":""} tarefa{naoLidas.length!==1?"s":""}!</div>
                <div style={{fontSize:"13px",color:"#6b7280",marginTop:"4px"}}>Atribuída{naoLidas.length!==1?"s":""} por {naoLidas[0].atribuidoPor}</div>
              </div>
              {naoLidas.map(function(t){return(
                <div key={t.id} style={{background:"#f9fafb",borderRadius:"8px",padding:"10px 12px",marginBottom:"8px"}}>
                  <div style={{fontSize:"13px",fontWeight:600,color:"#111"}}>{t.projetoNome}</div>
                  <div style={{fontSize:"12px",color:"#6b7280",marginTop:"2px"}}>{t.descricao}</div>
                </div>
              );})}
              <button onClick={marcarLidas} style={Object.assign({},B.pri,{width:"100%",padding:"11px",fontSize:"14px",marginTop:"4px"})}>Entendido</button>
            </div>
          </div>
        )}

        {currentMember&&tab==="minhas"&&(
          <div>

            <div style={B.lbl}>Suas tarefas em andamento</div>

            {ativas.length===0?(
              <div style={{textAlign:"center",padding:"2.5rem",color:"#9ca3af",fontSize:"13px"}}>Nenhuma tarefa pendente no momento.</div>
            ):(
              ativas.map(function(t){
                var st=TASK_ST[t.status]||TASK_ST.pendente;
                var acts=getActions(t.status);
                return (
                  <div key={t.id} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"1rem 1.25rem",marginBottom:"12px"}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:acts.length?"10px":"0"}}>
                      <div>
                        <div style={{fontSize:"15px",fontWeight:700,color:"#111",marginBottom:"3px"}}>{t.projetoNome}{t.atividadeNome?" — "+t.atividadeNome:""}</div>
                        <div style={{fontSize:"13px",color:"#9ca3af"}}>{t.descricao}</div>
                      </div>
                      <span style={{fontSize:"11px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:st.bg,color:st.color,flexShrink:0,marginLeft:"8px",whiteSpace:"nowrap"}}>{st.label}</span>
                    </div>
                    {acts.length>0&&(
                      <div style={{display:"grid",gridTemplateColumns:"repeat("+acts.length+",1fr)",gap:"8px",marginBottom:"6px"}}>
                        {acts.map(function(a){
                          return (
                            <button key={a.next} onClick={function(){changeStatus(t.id,a.next);}} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"8px",padding:"9px",fontSize:"13px",color:"#374151",cursor:"pointer",textAlign:"center",fontWeight:500}}>
                              {a.icon} {a.label}
                            </button>
                          );
                        })}
                      </div>
                    )}
                    {uDelegando&&uDelegando.taskId===t.id?(
                      <div style={{display:"flex",gap:"6px",alignItems:"center",marginTop:"4px"}}>
                        <select value={uDelegando.novoMembro} onChange={function(e){setUDelegando({taskId:t.id,novoMembro:e.target.value});}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px"})}>
                          <option value="">Escolher colaborador...</option>
                          {team.filter(function(m){return currentMember&&m.id!==currentMember.id;}).map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
                        </select>
                        <button onClick={function(){
                          if(!uDelegando.novoMembro)return;
                          var nm=team.find(function(m){return m.id===uDelegando.novoMembro;});
                          saveTarefas(tarefas.map(function(x){
                            if(x.id!==t.id)return x;
                            return Object.assign({},x,{membroId:uDelegando.novoMembro,membroNome:nm?nm.name:"",lida:false,delegadoHistorico:[...(x.delegadoHistorico||[]),{de:x.membroNome,para:nm?nm.name:"",quando:new Date().toISOString()}]});
                          }));
                          setUDelegando(null);
                        }} disabled={!uDelegando.novoMembro} style={Object.assign({},B.pri,{fontSize:"12px",padding:"6px 14px",flexShrink:0})}>Confirmar</button>
                        <button onClick={function(){setUDelegando(null);}} style={Object.assign({},B.sec,{fontSize:"12px",padding:"6px 12px",flexShrink:0})}>Cancelar</button>
                      </div>
                    ):(
                      <button onClick={function(){setUDelegando({taskId:t.id,novoMembro:"",});}} style={{background:"transparent",border:"1px solid #dbeafe",borderRadius:"8px",padding:"7px",fontSize:"12px",color:"#2563EB",cursor:"pointer",width:"100%",textAlign:"center",marginTop:"4px"}}>Delegar para outro colaborador</button>
                    )}
                  </div>
                );
              })
            )}

            {concluidas.length>0&&(
              <div style={{marginTop:"1.25rem",borderTop:"1px solid #f0f0f0",paddingTop:"1rem"}}>
                <div style={Object.assign({},B.lbl,{marginBottom:"8px",cursor:"pointer"})} onClick={function(){var el=document.getElementById("u-conc");if(el)el.style.display=el.style.display==="none"?"block":"none";}}>
                  Concluídas ({concluidas.length}) ▾
                </div>
                <div id="u-conc" style={{display:"none"}}>
                  <select value={uConcFiltProj} onChange={function(e){setUConcFiltProj(e.target.value);}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px",marginBottom:"8px"})}>
                    <option value="">Todos os projetos</option>
                    {[...new Set(concluidas.map(function(t){return t.projetoNome;}))].map(function(p){return <option key={p} value={p}>{p}</option>;}) }
                  </select>
                  {concluidas.filter(function(t){return !uConcFiltProj||t.projetoNome===uConcFiltProj;}).map(function(t){
                    var pessoas=[...(t.delegadoHistorico||[]).map(function(d){return d.de;}),t.membroNome].filter(function(v,idx,a){return a.indexOf(v)===idx;});
                    return (
                      <div key={t.id} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"1rem 1.25rem",marginBottom:"8px"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"6px"}}>
                          <div style={{flex:1}}>
                            <div style={{fontSize:"14px",fontWeight:700,color:"#111",marginBottom:"2px"}}>{t.projetoNome}{t.atividadeNome?" — "+t.atividadeNome:""}</div>
                            <div style={{fontSize:"12px",color:"#9ca3af"}}>{t.descricao}</div>
                            {pessoas.length>1&&<div style={{fontSize:"11px",color:"#6b7280",marginTop:"4px"}}>Envolvidos: {pessoas.join(" → ")}</div>}
                          </div>
                          <span style={{fontSize:"11px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:"#f3f4f6",color:"#6b7280",marginLeft:"8px",flexShrink:0}}>Concluído</span>
                        </div>
                        <div style={{display:"flex",gap:"14px",flexWrap:"wrap"}}>
                          {t.iniciadoEm&&<span style={{fontSize:"11px",color:"#6b7280"}}>▶ Início: {new Date(t.iniciadoEm).toLocaleDateString("pt-BR")}</span>}
                          {t.concluidoEm&&<span style={{fontSize:"11px",color:"#15803d",fontWeight:600}}>✓ Conclusão: {new Date(t.concluidoEm).toLocaleDateString("pt-BR")}</span>}
                          {(t.delegadoHistorico||[]).map(function(h,hi){return <span key={hi} style={{fontSize:"11px",color:"#9ca3af"}}>↳ {h.de} → {h.para} ({new Date(h.quando).toLocaleDateString("pt-BR")})</span>;})}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {tab==="historico"&&(
          <div style={{textAlign:"center",padding:"3rem",color:"#9ca3af",fontSize:"13px"}}>Histórico de tarefas em breve.</div>
        )}
      </div>
      <div style={{textAlign:"center",padding:"2rem 1rem 1.5rem",fontSize:"11px",color:"#9ca3af"}}>
        Criado por MAR Consultoria 2026 - v2.1
      </div>
    </div>
  );
}

function GestorApp(props){
  var authUser=props.authUser,handleLogout=props.handleLogout,tarefas=props.tarefas,saveTarefas=props.saveTarefas,currentMember=props.currentMember,team=props.team||[],projects=props.projects||[];
  var ts=useState("minhas"); var tab=ts[0],setTab=ts[1];
  var sms=useState(false); var senhaMode=sms[0],setSenhaMode=sms[1];
  var sas=useState(""); var senhaAtual=sas[0],setSenhaAtual=sas[1];
  var sns=useState(""); var senhaNova=sns[0],setSenhaNova=sns[1];
  var scs=useState(""); var senhaConf=scs[0],setSenhaConf=scs[1];
  var sls=useState(false); var senhaLoad=sls[0],setSenhaLoad=sls[1];
  var sss=useState(""); var senhaStatus=sss[0],setSenhaStatus=sss[1];
  var ps=useState(""); var pessoa=ps[0],setPessoa=ps[1];
  var prs=useState(""); var projeto=prs[0],setProjeto=prs[1];
  var ats=useState(""); var atividade=ats[0],setAtividade=ats[1];
  var ds=useState(""); var desc=ds[0],setDesc=ds[1];
  var errs=useState(""); var err=errs[0],setErr=errs[1];
  var dels=useState(null); var delegando=dels[0],setDelegando=dels[1];
  var udels=useState(null); var uDelegando=udels[0],setUDelegando=udels[1];

  var projSel=projects.find(function(p){return p.id===projeto;});
  var subitens=(projSel&&projSel.subitems)||[];
  var gestorNome=currentMember?currentMember.name:(authUser?authUser.email.split("@")[0]:"Gestor");
  var gestorEmail=authUser?authUser.email:"";

  async function handleAlterarSenha(){
    if(!senhaAtual||!senhaNova||!senhaConf){setSenhaStatus("Preencha todos os campos.");return;}
    if(senhaNova!==senhaConf){setSenhaStatus("As senhas não coincidem.");return;}
    if(senhaNova.length<6){setSenhaStatus("Mínimo 6 caracteres.");return;}
    setSenhaLoad(true);setSenhaStatus("");
    try{
      var app=getApps().length?getApp():initializeApp(firebaseConfig);
      var auth=getAuth(app);
      var cred=EmailAuthProvider.credential(authUser.email,senhaAtual);
      await reauthenticateWithCredential(auth.currentUser,cred);
      await updatePassword(auth.currentUser,senhaNova);
      setSenhaStatus("ok");setSenhaAtual("");setSenhaNova("");setSenhaConf("");
      setTimeout(function(){setSenhaMode(false);setSenhaStatus("");},2000);
    }catch(ex){
      var msgs={"auth/wrong-password":"Senha atual incorreta.","auth/invalid-credential":"Senha atual incorreta.","auth/weak-password":"Senha muito fraca."};
      setSenhaStatus(msgs[ex.code]||"Erro ao alterar senha.");
    }
    setSenhaLoad(false);
  }

  function addTarefa(){
    if(!pessoa||!projeto||!desc.trim()){setErr("Preencha todos os campos.");return;}
    setErr("");
    var membro=team.find(function(m){return m.id===pessoa;});
    var proj=projects.find(function(p){return p.id===projeto;});
    var atNome=atividade?(subitens.find(function(s){return s.id===atividade;})||{}).name||"":"";
    var nova={id:uid(),membroId:pessoa,membroNome:membro?membro.name:"",projetoId:projeto,projetoNome:proj?proj.name:"",atividadeId:atividade||null,atividadeNome:atNome,descricao:desc.trim(),status:"pendente",lida:false,criadoEm:new Date().toISOString(),atribuidoPor:gestorNome,criadoPorEmail:gestorEmail,iniciadoEm:null,pausadoEm:null,concluidoEm:null};
    saveTarefas([...tarefas,nova]);
    setPessoa("");setProjeto("");setAtividade("");setDesc("");
  }

  function changeStatus(taskId,newStatus){
    var now=new Date().toISOString();
    saveTarefas(tarefas.map(function(t){
      if(t.id!==taskId)return t;
      return Object.assign({},t,{status:newStatus,iniciadoEm:newStatus==="em_andamento"&&!t.iniciadoEm?now:t.iniciadoEm,pausadoEm:newStatus==="pausado"?now:t.pausadoEm,concluidoEm:newStatus==="concluido"?now:t.concluidoEm});
    }));
  }
  function getActions(status){
    if(status==="pendente")    return [{label:"Iniciar",icon:"▷",next:"em_andamento"}];
    if(status==="em_andamento")return [{label:"Pausar",icon:"⏸",next:"pausado"},{label:"Concluir",icon:"✓",next:"concluido"}];
    if(status==="pausado")     return [{label:"Retomar",icon:"▷",next:"em_andamento"},{label:"Concluir",icon:"✓",next:"concluido"}];
    return [];
  }

  var criadas=tarefas.filter(function(t){return t.criadoPorEmail===gestorEmail||(t.criadoPorEmail==null&&t.atribuidoPor===gestorNome);});
  var criAtivas=criadas.filter(function(t){return t.status!=="concluido";});
  var minhasTarefas=currentMember?tarefas.filter(function(t){return t.membroId===currentMember.id;}):[];
  var mAtivas=minhasTarefas.filter(function(t){return t.status!=="concluido";});
  var mConcluidas=minhasTarefas.filter(function(t){return t.status==="concluido";});
  var naoLidas=mAtivas.filter(function(t){return !t.lida&&t.status==="pendente";});

  function marcarLidas(){saveTarefas(tarefas.map(function(t){return(currentMember&&t.membroId===currentMember.id&&!t.lida)?Object.assign({},t,{lida:true}):t;}));}

  return (
    <div style={{background:"#f9fafb",minHeight:"100vh",fontFamily:"-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif"}}>
      <div style={{height:"3px",background:RED}}/>
      <div style={{background:"#fff",borderBottom:"1px solid #f0f0f0",padding:"0 1.25rem",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 3px rgba(0,0,0,.04)"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"10px 0 0"}}>
          <div style={{display:"flex",alignItems:"baseline",gap:"10px"}}>
            <span style={{fontSize:"22px",fontWeight:800,letterSpacing:"-1px",color:"#111"}}>M<span style={{color:RED}}>+</span>M</span>
            <span style={{fontSize:"11px",color:"#9ca3af",letterSpacing:".05em",textTransform:"uppercase",fontWeight:500}}>Arquitetura</span>
            <span style={{fontSize:"10px",padding:"2px 8px",borderRadius:"20px",background:"#dbeafe",color:"#2563EB",fontWeight:700}}>Gestor</span>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
            <span style={{fontSize:"11px",color:"#9ca3af",maxWidth:"140px",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{authUser.email}</span>
            <button onClick={function(){setSenhaMode(function(v){return !v;});setSenhaStatus("");}} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",cursor:"pointer"}}>Senha</button>
            <button onClick={handleLogout} style={{fontSize:"11px",padding:"3px 9px",borderRadius:"6px",border:"1px solid #e5e7eb",background:"transparent",color:"#6b7280",cursor:"pointer"}}>Sair</button>
          </div>
        </div>
        <div style={{display:"flex",marginTop:"4px"}}>
          {[["atribuir","Tarefas"],["minhas","Minhas tarefas"]].map(function(t){
            return <button key={t[0]} onClick={function(){setTab(t[0]);}} style={{background:"transparent",border:"none",cursor:"pointer",padding:"7px 12px",fontSize:"12px",fontWeight:tab===t[0]?700:400,color:tab===t[0]?"#111":"#9ca3af",borderBottom:tab===t[0]?"2.5px solid "+RED:"2.5px solid transparent",whiteSpace:"nowrap"}}>{t[1]}{t[0]==="minhas"&&naoLidas.length>0?" ("+naoLidas.length+")":""}</button>;
          })}
        </div>
      </div>

      <div style={{padding:"1.25rem",maxWidth:"700px",margin:"0 auto"}}>
        {senhaMode&&(
          <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"1rem 1.25rem",marginBottom:"1.25rem"}}>
            <div style={{fontSize:"14px",fontWeight:700,color:"#111",marginBottom:"12px"}}>Alterar senha</div>
            <div style={{display:"flex",flexDirection:"column",gap:"8px",marginBottom:"12px"}}>
              <input type="password" value={senhaAtual} onChange={function(e){setSenhaAtual(e.target.value);}} placeholder="Senha atual" style={B.inp}/>
              <input type="password" value={senhaNova} onChange={function(e){setSenhaNova(e.target.value);}} placeholder="Nova senha (mín. 6 caracteres)" style={B.inp}/>
              <input type="password" value={senhaConf} onChange={function(e){setSenhaConf(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")handleAlterarSenha();}} placeholder="Confirmar nova senha" style={B.inp}/>
            </div>
            {senhaStatus&&senhaStatus!=="ok"&&<div style={{fontSize:"12px",color:"#dc2626",marginBottom:"8px"}}>{senhaStatus}</div>}
            {senhaStatus==="ok"&&<div style={{fontSize:"12px",color:"#15803d",marginBottom:"8px"}}>✓ Senha alterada!</div>}
            <div style={{display:"flex",gap:"8px"}}>
              <button onClick={handleAlterarSenha} disabled={senhaLoad} style={B.pri}>{senhaLoad?"Salvando...":"Salvar senha"}</button>
              <button onClick={function(){setSenhaMode(false);setSenhaAtual("");setSenhaNova("");setSenhaConf("");setSenhaStatus("");}} style={B.sec}>Cancelar</button>
            </div>
          </div>
        )}

        {tab==="atribuir"&&(
          <div>
            <div style={B.lbl}>Atribuir nova tarefa</div>
            <div style={Object.assign({},B.card,{marginBottom:"1.5rem"})}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:"8px",marginBottom:"8px"}}>
                <select value={pessoa} onChange={function(e){setPessoa(e.target.value);}} style={B.inp}>
                  <option value="">Colaborador...</option>
                  {currentMember&&<option value={currentMember.id}>Eu mesmo ({gestorNome})</option>}
                  {team.filter(function(m){return !currentMember||m.id!==currentMember.id;}).map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
                </select>
                <select value={projeto} onChange={function(e){setProjeto(e.target.value);setAtividade("");}} style={B.inp}>
                  <option value="">Projeto...</option>
                  {projects.filter(function(p){return p.status!=="concluido";}).map(function(p){return <option key={p.id} value={p.id}>{p.name}</option>;})}
                </select>
              </div>
              {subitens.length>0&&(
                <div style={{marginBottom:"8px"}}>
                  <select value={atividade} onChange={function(e){setAtividade(e.target.value);}} style={B.inp}>
                    <option value="">Atividade (opcional)...</option>
                    {subitens.map(function(si){return <option key={si.id} value={si.id}>{si.name}</option>;})}
                  </select>
                </div>
              )}
              <div style={{display:"flex",gap:"8px"}}>
                <input type="text" value={desc} onChange={function(e){setDesc(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")addTarefa();}} placeholder="Descrição da tarefa..." style={Object.assign({},B.inp,{flex:1})}/>
                <button onClick={addTarefa} style={B.pri}>Atribuir</button>
              </div>
              {err&&<div style={{fontSize:"12px",color:"#dc2626",marginTop:"6px"}}>{err}</div>}
            </div>

            <div style={B.lbl}>Tarefas que atribuí</div>
            {criAtivas.length===0?(
              <div style={{textAlign:"center",padding:"2rem",color:"#9ca3af",fontSize:"13px"}}>Nenhuma tarefa ativa atribuída por você.</div>
            ):criAtivas.map(function(t){
              var st=TASK_ST[t.status]||TASK_ST.pendente;
              var isDel=delegando&&delegando.taskId===t.id;
              return (
                <div key={t.id} style={Object.assign({},B.card,{padding:"10px 14px"})}>
                  <div style={{display:"flex",alignItems:"center",gap:"10px",flexWrap:"wrap"}}>
                    <div style={{fontWeight:700,fontSize:"13px",color:"#111",minWidth:"80px",flexShrink:0}}>{t.membroNome}</div>
                    <div style={{fontSize:"12px",color:"#9ca3af",flex:1}}>{t.projetoNome}{t.atividadeNome?" · "+t.atividadeNome:""} — {t.descricao}</div>
                    <span style={{fontSize:"11px",fontWeight:600,padding:"2px 9px",borderRadius:"20px",background:st.bg,color:st.color,flexShrink:0}}>{st.label}</span>
                    {isDel?(
                      <div style={{display:"flex",gap:"5px",alignItems:"center"}}>
                        <select value={delegando.novoMembro} onChange={function(e){setDelegando({taskId:t.id,novoMembro:e.target.value});}} style={Object.assign({},B.inp,{fontSize:"11px",padding:"3px 7px",width:"130px"})}>
                          <option value="">Para quem?</option>
                          {team.filter(function(m){return m.id!==t.membroId;}).map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
                        </select>
                        <button onClick={function(){
                          if(!delegando.novoMembro)return;
                          var nm=team.find(function(m){return m.id===delegando.novoMembro;});
                          saveTarefas(tarefas.map(function(x){if(x.id!==t.id)return x;return Object.assign({},x,{membroId:delegando.novoMembro,membroNome:nm?nm.name:"",lida:false,delegadoHistorico:[...(x.delegadoHistorico||[]),{de:x.membroNome,para:nm?nm.name:"",quando:new Date().toISOString()}]});}));
                          setDelegando(null);
                        }} disabled={!delegando.novoMembro} style={Object.assign({},B.pri,{fontSize:"11px",padding:"3px 9px"})}>✓</button>
                        <button onClick={function(){setDelegando(null);}} style={Object.assign({},B.sec,{fontSize:"11px",padding:"3px 9px"})}>✕</button>
                      </div>
                    ):(
                      <button onClick={function(){setDelegando({taskId:t.id,novoMembro:"",});}} style={Object.assign({},B.ghost,{color:"#2563EB",fontSize:"12px",border:"1px solid #dbeafe",borderRadius:"6px",padding:"2px 8px"})}>Delegar</button>
                    )}
                  </div>
                  {t.delegadoHistorico&&t.delegadoHistorico.length>0&&(
                    <div style={{marginTop:"5px",fontSize:"11px",color:"#9ca3af",borderTop:"1px solid #f0f0f0",paddingTop:"5px"}}>
                      {t.delegadoHistorico.map(function(h,i){return <span key={i} style={{marginRight:"10px"}}>↳ {h.de} → {h.para} ({new Date(h.quando).toLocaleDateString("pt-BR")})</span>;})}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        )}

        {naoLidas.length>0&&(
          <div style={{position:"fixed",top:0,left:0,right:0,bottom:0,background:"rgba(0,0,0,0.45)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}>
            <div style={{background:"#fff",borderRadius:"16px",padding:"1.5rem",maxWidth:"360px",width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.3)"}}>
              <div style={{textAlign:"center",marginBottom:"1rem"}}>
                <div style={{fontSize:"36px",marginBottom:"8px"}}>🔔</div>
                <div style={{fontSize:"16px",fontWeight:700,color:"#111"}}>Nova{naoLidas.length!==1?"s":""} tarefa{naoLidas.length!==1?"s":""}!</div>
                <div style={{fontSize:"13px",color:"#6b7280",marginTop:"4px"}}>Atribuída{naoLidas.length!==1?"s":""} por {naoLidas[0].atribuidoPor}</div>
              </div>
              {naoLidas.map(function(t){return(
                <div key={t.id} style={{background:"#f9fafb",borderRadius:"8px",padding:"10px 12px",marginBottom:"8px"}}>
                  <div style={{fontSize:"13px",fontWeight:600,color:"#111"}}>{t.projetoNome}</div>
                  <div style={{fontSize:"12px",color:"#6b7280",marginTop:"2px"}}>{t.descricao}</div>
                </div>
              );})}
              <button onClick={marcarLidas} style={Object.assign({},B.pri,{width:"100%",padding:"11px",fontSize:"14px",marginTop:"4px"})}>Entendido</button>
            </div>
          </div>
        )}

        {tab==="minhas"&&(
          <div>
            {!currentMember&&(
              <div style={{background:"#fef2f2",border:"1px solid #fecaca",borderRadius:"12px",padding:"1.25rem",textAlign:"center",color:"#dc2626",fontSize:"13px",lineHeight:1.7}}>
                Seu e-mail não está vinculado a nenhum membro da equipe.<br/>Peça ao administrador para configurar seu e-mail na aba Equipe.
              </div>
            )}
            <div style={B.lbl}>Suas tarefas</div>
            {mAtivas.length===0?(
              <div style={{textAlign:"center",padding:"2.5rem",color:"#9ca3af",fontSize:"13px"}}>Nenhuma tarefa pendente.</div>
            ):mAtivas.map(function(t){
              var st=TASK_ST[t.status]||TASK_ST.pendente;
              var acts=getActions(t.status);
              return (
                <div key={t.id} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"1rem 1.25rem",marginBottom:"12px"}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:acts.length?"10px":"0"}}>
                    <div>
                      <div style={{fontSize:"15px",fontWeight:700,color:"#111",marginBottom:"3px"}}>{t.projetoNome}{t.atividadeNome?" — "+t.atividadeNome:""}</div>
                      <div style={{fontSize:"13px",color:"#9ca3af"}}>{t.descricao}</div>
                    </div>
                    <span style={{fontSize:"11px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:st.bg,color:st.color,flexShrink:0,marginLeft:"8px"}}>{st.label}</span>
                  </div>
                  {acts.length>0&&(
                    <div style={{display:"grid",gridTemplateColumns:"repeat("+acts.length+",1fr)",gap:"8px",marginBottom:"6px"}}>
                      {acts.map(function(a){return <button key={a.next} onClick={function(){changeStatus(t.id,a.next);}} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"8px",padding:"9px",fontSize:"13px",color:"#374151",cursor:"pointer",textAlign:"center",fontWeight:500}}>{a.icon} {a.label}</button>;})}
                    </div>
                  )}
                  {uDelegando&&uDelegando.taskId===t.id?(
                    <div style={{display:"flex",gap:"6px",alignItems:"center",marginTop:"4px"}}>
                      <select value={uDelegando.novoMembro} onChange={function(e){setUDelegando({taskId:t.id,novoMembro:e.target.value});}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px"})}>
                        <option value="">Escolher colaborador...</option>
                        {team.filter(function(m){return !currentMember||m.id!==currentMember.id;}).map(function(m){return <option key={m.id} value={m.id}>{m.name}</option>;})}
                      </select>
                      <button onClick={function(){
                        if(!uDelegando.novoMembro)return;
                        var nm=team.find(function(m){return m.id===uDelegando.novoMembro;});
                        saveTarefas(tarefas.map(function(x){if(x.id!==t.id)return x;return Object.assign({},x,{membroId:uDelegando.novoMembro,membroNome:nm?nm.name:"",lida:false,delegadoHistorico:[...(x.delegadoHistorico||[]),{de:x.membroNome,para:nm?nm.name:"",quando:new Date().toISOString()}]});}));
                        setUDelegando(null);
                      }} disabled={!uDelegando.novoMembro} style={Object.assign({},B.pri,{fontSize:"12px",padding:"6px 14px",flexShrink:0})}>Confirmar</button>
                      <button onClick={function(){setUDelegando(null);}} style={Object.assign({},B.sec,{fontSize:"12px",padding:"6px 12px",flexShrink:0})}>Cancelar</button>
                    </div>
                  ):(
                    <button onClick={function(){setUDelegando({taskId:t.id,novoMembro:"",});}} style={{background:"transparent",border:"1px solid #dbeafe",borderRadius:"8px",padding:"7px",fontSize:"12px",color:"#2563EB",cursor:"pointer",width:"100%",textAlign:"center",marginTop:"4px"}}>Delegar para outro colaborador</button>
                  )}
                </div>
              );
            })}
            {mConcluidas.length>0&&(
              <div style={{marginTop:"1.25rem",borderTop:"1px solid #f0f0f0",paddingTop:"1rem"}}>
                <div style={Object.assign({},B.lbl,{cursor:"pointer"})} onClick={function(){var el=document.getElementById("g-conc");if(el)el.style.display=el.style.display==="none"?"block":"none";}}>Concluídas ({mConcluidas.length}) ▾</div>
                <div id="g-conc" style={{display:"none"}}>
                  {mConcluidas.map(function(t){
                    var pessoas=[...(t.delegadoHistorico||[]).map(function(d){return d.de;}),t.membroNome].filter(function(v,idx,a){return a.indexOf(v)===idx;});
                    return(
                      <div key={t.id} style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:"12px",padding:"1rem 1.25rem",marginBottom:"8px"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:"6px"}}>
                          <div style={{flex:1}}>
                            <div style={{fontSize:"14px",fontWeight:700,color:"#111",marginBottom:"2px"}}>{t.projetoNome}{t.atividadeNome?" — "+t.atividadeNome:""}</div>
                            <div style={{fontSize:"12px",color:"#9ca3af"}}>{t.descricao}</div>
                            {pessoas.length>1&&<div style={{fontSize:"11px",color:"#6b7280",marginTop:"3px"}}>Envolvidos: {pessoas.join(" → ")}</div>}
                          </div>
                          <span style={{fontSize:"11px",fontWeight:600,padding:"3px 10px",borderRadius:"20px",background:"#f3f4f6",color:"#6b7280",marginLeft:"8px",flexShrink:0}}>Concluído</span>
                        </div>
                        <div style={{display:"flex",gap:"14px",flexWrap:"wrap"}}>
                          {t.iniciadoEm&&<span style={{fontSize:"11px",color:"#6b7280"}}>▶ Início: {new Date(t.iniciadoEm).toLocaleDateString("pt-BR")}</span>}
                          {t.concluidoEm&&<span style={{fontSize:"11px",color:"#15803d",fontWeight:600}}>✓ Conclusão: {new Date(t.concluidoEm).toLocaleDateString("pt-BR")}</span>}
                          {(t.delegadoHistorico||[]).map(function(h,hi){return <span key={hi} style={{fontSize:"11px",color:"#9ca3af"}}>↳ {h.de} → {h.para} ({new Date(h.quando).toLocaleDateString("pt-BR")})</span>;})}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
      <div style={{textAlign:"center",padding:"2rem 1rem 1.5rem",fontSize:"11px",color:"#9ca3af"}}>Criado por MAR Consultoria 2026 - v2.1</div>
    </div>
  );
}

function HistTab(props){
  var history=props.history,projects=props.projects,cont=props.cont,team=props.team,colorMap=props.colorMap,saveHistory=props.saveHistory,tarefas=props.tarefas||[];
  var ss=useState(null); var sel=ss[0],setSel=ss[1];
  var es=useState({}); var expanded=es[0],setExpanded=es[1];
  var ns=useState({}); var notes=ns[0],setNotes=ns[1];
  var fps=useState(""); var fPessoa=fps[0],setFPessoa=fps[1];
  var fprjs=useState(""); var fProjHist=fprjs[0],setFProjHist=fprjs[1];
  var dates=Object.keys(history).sort(function(a,b){return b.localeCompare(a);});
  function toggleExpand(d){setExpanded(function(p){return Object.assign({},p,{[d]:!p[d]});});}
  function saveNote(d){
    var note=notes[d]!=null?notes[d]:(history[d]&&history[d].note?history[d].note:"");
    saveHistory(Object.assign({},history,{[d]:Object.assign({},history[d],{note:note})}));
  }
  if(!dates.length)return <div style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}>Nenhum historico salvo ainda.</div>;
  if(sel){
    var rep=calcReport(history[sel]?history[sel].alloc||{}:{},team,projects,{});
    var tarefasDia=(sel===TODAY?tarefas.filter(function(t){return t.status!=="pendente"&&(t.status!=="concluido"||(t.concluidoEm&&t.concluidoEm.startsWith(TODAY)));}):((history[sel]&&history[sel].tarefasSnap)||[]).filter(function(t){return t.status!=="pendente";}));
    return (
      <div>
        <button onClick={function(){setSel(null);}} style={Object.assign({},B.ghost,{color:RED,padding:0,fontWeight:600,fontSize:"13px",marginBottom:"1rem"})}>Voltar</button>
        <div style={{fontSize:"18px",fontWeight:700,color:"#111",marginBottom:"2px"}}>{fmtDate(sel)}</div>
        <div style={{fontSize:"13px",color:"#9ca3af",marginBottom:"1.25rem"}}>{cap(fmtDay(sel))}</div>
        <div style={Object.assign({},B.card,{marginBottom:"1.25rem",borderLeft:"3px solid "+RED})}>
          <div style={{fontSize:"11px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".06em",marginBottom:"8px"}}>Nota do dia</div>
          <textarea placeholder="Registre observacoes sobre este dia (max. 500 caracteres)..." maxLength={1000} value={notes[sel]!=null?notes[sel]:(history[sel]&&history[sel].note?history[sel].note:"")} onChange={function(e){setNotes(function(p){return Object.assign({},p,{[sel]:e.target.value});});}} style={Object.assign({},B.inp,{height:"72px",resize:"vertical",marginBottom:"6px"})}/>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <span style={{fontSize:"11px",color:"#d1d5db"}}>{(notes[sel]!=null?notes[sel]:(history[sel]&&history[sel].note?history[sel].note:"")).length}/1000</span>
            <button onClick={function(){saveNote(sel);}} style={Object.assign({},B.pri,{fontSize:"12px",padding:"5px 14px"})}>Salvar nota</button>
          </div>
        </div>
        {rep.map(function(r){
          var p=projects.find(function(x){return x.id===r.pid;});
          var c=cont.find(function(x){return x.id===(p?p.cId:"");});
          return (
            <div key={r.pid} style={{marginBottom:"14px"}}>
              <div style={{display:"flex",justifyContent:"space-between",marginBottom:"4px"}}>
                <div><span style={{fontSize:"13px",fontWeight:600,color:"#111"}}>{r.name}</span>{c&&<span style={{fontSize:"11px",color:"#9ca3af",marginLeft:"6px"}}>{c.name}</span>}</div>
                <span style={{fontSize:"12px",color:"#9ca3af"}}>{r.pct.toFixed(1)}%</span>
              </div>
              <div style={{height:"6px",background:"#f3f4f6",borderRadius:"3px",overflow:"hidden",marginBottom:"3px"}}>
                <div style={{height:"100%",width:r.pct+"%",background:colorMap[r.pid]||"#888",borderRadius:"3px"}}/>
              </div>
              <div style={{fontSize:"11px",color:"#9ca3af"}}>{r.people.join(", ")}</div>
            </div>
          );
        })}
        {tarefasDia.length>0&&(
          <div style={{marginTop:"1.5rem",borderTop:"1px solid #f0f0f0",paddingTop:"1rem"}}>
            <div style={B.lbl}>Tarefas do dia</div>
            <div style={{display:"flex",gap:"8px",marginBottom:"10px"}}>
              <select value={fProjHist} onChange={function(e){setFProjHist(e.target.value);}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px"})}>
                <option value="">Todos os projetos</option>
                {[...new Set(tarefasDia.map(function(t){return t.projetoNome;}))].map(function(p){return <option key={p} value={p}>{p}</option>;}) }
              </select>
              <select value={fPessoa} onChange={function(e){setFPessoa(e.target.value);}} style={Object.assign({},B.inp,{fontSize:"12px",padding:"5px 8px"})}>
                <option value="">Todas as pessoas</option>
                {[...new Set(tarefasDia.map(function(t){return t.membroNome;}))].map(function(p){return <option key={p} value={p}>{p}</option>;}) }
              </select>
            </div>
            {tarefasDia.filter(function(t){return(!fPessoa||t.membroNome===fPessoa)&&(!fProjHist||t.projetoNome===fProjHist);}).map(function(t,i){
              var st=TASK_ST[t.status]||TASK_ST.pendente;
              var pessoas=[...(t.delegadoHistorico||[]).map(function(d){return d.de;}),t.membroNome].filter(function(v,idx,a){return a.indexOf(v)===idx;});
              return (
                <div key={t.id||i} style={Object.assign({},B.card,{padding:"10px 14px"})}>
                  <div style={{display:"flex",alignItems:"flex-start",gap:"10px",flexWrap:"wrap"}}>
                    <div style={{flex:1}}>
                      <div style={{display:"flex",alignItems:"center",gap:"8px",marginBottom:"4px",flexWrap:"wrap"}}>
                        <span style={{fontWeight:700,fontSize:"13px",color:"#111"}}>{pessoas.join(" → ")}</span>
                        <span style={{fontSize:"10px",fontWeight:600,padding:"2px 8px",borderRadius:"20px",background:st.bg,color:st.color}}>{st.label}</span>
                      </div>
                      <div style={{fontSize:"12px",color:"#9ca3af"}}>{t.projetoNome}{t.atividadeNome?" · "+t.atividadeNome:""}{t.descricao?" — "+t.descricao:""}</div>
                      <div style={{display:"flex",gap:"14px",marginTop:"5px",flexWrap:"wrap"}}>
                        {t.iniciadoEm&&<span style={{fontSize:"11px",color:"#6b7280"}}>▶ Início: {new Date(t.iniciadoEm).toLocaleDateString("pt-BR")}</span>}
                        {t.concluidoEm&&<span style={{fontSize:"11px",color:"#15803d",fontWeight:600}}>✓ Conclusão: {new Date(t.concluidoEm).toLocaleDateString("pt-BR")}</span>}
                        {(t.delegadoHistorico||[]).map(function(h,hi){return <span key={hi} style={{fontSize:"11px",color:"#9ca3af"}}>↳ {h.de} → {h.para} ({new Date(h.quando).toLocaleDateString("pt-BR")})</span>;})}
                      </div>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    );
  }
  return (
    <div>
      <div style={{fontSize:"13px",color:"#9ca3af",marginBottom:"1rem"}}>{dates.length} dia{dates.length!==1?"s":""} registrado{dates.length!==1?"s":""}</div>
      {dates.map(function(d){
        var a=normAlloc(history[d]?history[d].alloc||{}:{});
        var pids=[...new Set(Object.values(a).reduce(function(acc,v){return acc.concat(v);},[])).values()].filter(Boolean);
        var cnt=team.filter(function(m){return (a[m.id]||[]).length>0;}).length;
        var isOpen=!!expanded[d];
        var hasNote=!!(history[d]&&history[d].note);
        return (
          <div key={d} style={Object.assign({},B.card,{marginBottom:"6px"})}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer"}} onClick={function(){toggleExpand(d);}}>
              <div>
                <div style={{display:"flex",alignItems:"center",gap:"8px"}}>
                  <div style={{fontSize:"14px",fontWeight:600,color:"#111"}}>{fmtDate(d)}</div>
                  {hasNote&&<span style={{fontSize:"10px",padding:"1px 6px",borderRadius:"10px",background:"#fef3c7",color:"#d97706",fontWeight:600}}>nota</span>}
                </div>
                <div style={{fontSize:"12px",color:"#9ca3af",marginTop:"1px"}}>{cap(fmtDay(d))}</div>
              </div>
              <div style={{display:"flex",alignItems:"center",gap:"10px"}}>
                <div style={{display:"flex",gap:"3px"}}>{pids.slice(0,7).map(function(pid){return <div key={pid} style={{width:"8px",height:"8px",borderRadius:"50%",background:colorMap[pid]||"#888"}}/>;})}</div>
                <div style={{textAlign:"right",fontSize:"12px",color:"#9ca3af",lineHeight:1.6}}>
                  <div>{pids.length} proj.</div><div>{cnt} pessoas</div>
                </div>
                <span style={{fontSize:"12px",color:"#9ca3af"}}>{isOpen?"v":">"}</span>
              </div>
            </div>
            {isOpen&&(
              <div style={{marginTop:"12px",paddingTop:"12px",borderTop:"1px solid #f0f0f0"}}>
                <div style={{fontSize:"11px",fontWeight:700,color:"#9ca3af",textTransform:"uppercase",letterSpacing:".06em",marginBottom:"8px"}}>Nota do dia</div>
                <textarea placeholder="Registre observacoes sobre este dia (max. 500 caracteres)..." maxLength={1000} value={notes[d]!=null?notes[d]:(history[d]&&history[d].note?history[d].note:"")} onChange={function(e){var val=e.target.value;setNotes(function(p){return Object.assign({},p,{[d]:val});});}} style={Object.assign({},B.inp,{height:"64px",resize:"vertical",marginBottom:"6px"})}/>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"12px"}}>
                  <span style={{fontSize:"11px",color:"#d1d5db"}}>{(notes[d]!=null?notes[d]:(history[d]&&history[d].note?history[d].note:"")).length}/1000</span>
                  <button onClick={function(e){e.stopPropagation();saveNote(d);}} style={Object.assign({},B.pri,{fontSize:"12px",padding:"5px 14px"})}>Salvar nota</button>
                </div>
                <button onClick={function(){setSel(d);}} style={Object.assign({},B.sec,{fontSize:"12px",width:"100%",textAlign:"center"})}>Ver relatorio completo deste dia</button>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
