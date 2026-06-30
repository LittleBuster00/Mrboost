import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── DATE HELPERS ─────────────────────────────────────────────────────────────
const getMonday = (d) => {
  const date = new Date(d instanceof Date ? d : d + "T00:00:00");
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  date.setHours(0,0,0,0);
  return date;
};
const addDays = (d, n) => { const r = new Date(d); r.setDate(r.getDate() + n); return r; };
const toISO = (d) => {
  const x = new Date(d);
  return `${x.getFullYear()}-${String(x.getMonth()+1).padStart(2,"0")}-${String(x.getDate()).padStart(2,"0")}`;
};
const fmtS = (d) => new Date(d instanceof Date ? d : d + "T00:00:00").toLocaleDateString("id-ID",{day:"2-digit",month:"short"});
const fmtL = (d) => new Date(d instanceof Date ? d : d + "T00:00:00").toLocaleDateString("id-ID",{day:"2-digit",month:"long",year:"numeric"});
const weekLabel = (mondayISO) => {
  const mon = new Date(mondayISO + "T00:00:00");
  const sun = addDays(mon, 6);
  return `${fmtS(mon)} – ${fmtL(sun)}`;
};
const thisMonday = toISO(getMonday(new Date()));

// ─── FORMAT HELPERS ───────────────────────────────────────────────────────────
const fmt = (n) => new Intl.NumberFormat("id-ID",{style:"currency",currency:"IDR",minimumFractionDigits:0}).format(n||0);
const num = (s) => Math.max(0, parseInt(String(s||0).replace(/[^\d]/g,""))||0);
const uid = () => Math.random().toString(36).slice(2,9);

const LOC_COLORS = [
  {bg:"#4ade8018",text:"#4ade80",border:"#4ade8040"},
  {bg:"#60a5fa18",text:"#60a5fa",border:"#60a5fa40"},
  {bg:"#fb923c18",text:"#fb923c",border:"#fb923c40"},
  {bg:"#f472b618",text:"#f472b6",border:"#f472b640"},
  {bg:"#a78bfa18",text:"#a78bfa",border:"#a78bfa40"},
  {bg:"#34d39918",text:"#34d399",border:"#34d39940"},
  {bg:"#fbbf2418",text:"#fbbf24",border:"#fbbf2440"},
  {bg:"#f8717118",text:"#f87171",border:"#f8717140"},
];

// ─── STORAGE ──────────────────────────────────────────────────────────────────
async function sGet(k){try{const r=await window.storage.get(k);return r?JSON.parse(r.value):null;}catch{return null;}}
async function sSave(k,v){try{await window.storage.set(k,JSON.stringify(v));}catch(e){console.error(e);}}

// ─── CSS ──────────────────────────────────────────────────────────────────────
const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=Azeret+Mono:wght@400;500;600&display=swap');
  *{box-sizing:border-box;margin:0;padding:0}
  html,body{background:#111318;color:#e2e8f0;font-family:'Syne',sans-serif;font-size:14px}
  ::-webkit-scrollbar{width:4px;height:4px}
  ::-webkit-scrollbar-track{background:#111318}
  ::-webkit-scrollbar-thumb{background:#2d3748;border-radius:2px}
  input,select{font-family:'Syne',sans-serif;background:#1a1f2e;border:1px solid #2d3748;color:#e2e8f0;border-radius:7px;padding:8px 11px;font-size:13px;outline:none;transition:border .15s,box-shadow .15s;width:100%}
  input:focus,select:focus{border-color:#4ade80;box-shadow:0 0 0 2px #4ade8018}
  input[type=number]{-moz-appearance:textfield}
  input[type=number]::-webkit-inner-spin-button{-webkit-appearance:none}
  input[type=checkbox]{width:16px;height:16px;flex-shrink:0;accent-color:#4ade80;cursor:pointer;padding:0}
  .mono{font-family:'Azeret Mono',monospace!important}
  .btn{border:none;cursor:pointer;font-family:'Syne',sans-serif;font-weight:700;border-radius:7px;padding:8px 16px;font-size:12px;transition:all .15s;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
  .btn:disabled{opacity:.5;cursor:not-allowed}
  .btn-g{background:#4ade80;color:#0a0d12}.btn-g:hover:not(:disabled){background:#22c55e}
  .btn-r{background:#ef4444;color:#fff}.btn-r:hover:not(:disabled){background:#dc2626}
  .btn-b{background:#3b82f6;color:#fff}.btn-b:hover:not(:disabled){background:#2563eb}
  .btn-s{background:#1a1f2e;color:#94a3b8;border:1px solid #2d3748}.btn-s:hover:not(:disabled){background:#252d40;color:#e2e8f0}
  .card{background:#161b27;border:1px solid #1e2535;border-radius:12px;padding:18px}
  .t-wrap{border:1px solid #1e2535;border-radius:10px;overflow:hidden}
  table{width:100%;border-collapse:collapse}
  th{background:#0e1420;color:#475569;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;padding:9px 12px;text-align:left;white-space:nowrap}
  td{padding:9px 12px;border-bottom:1px solid #1a2030;font-size:13px;vertical-align:top}
  tr:last-child td{border-bottom:none}
  .hov:hover td{background:#1a2030}
  .label{font-size:10px;font-weight:700;color:#475569;margin-bottom:4px;text-transform:uppercase;letter-spacing:.06em}
  .bdg{display:inline-flex;align-items:center;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:700}
  .bdg-g{background:#4ade8018;color:#4ade80;border:1px solid #4ade8030}
  .bdg-r{background:#ef444418;color:#ef4444;border:1px solid #ef444430}
  .bdg-y{background:#fbbf2418;color:#fbbf24;border:1px solid #fbbf2430}
  .overlay{position:fixed;inset:0;background:#00000088;display:flex;align-items:center;justify-content:center;z-index:300;backdrop-filter:blur(3px)}
  .modal{background:#161b27;border:1px solid #2d3748;border-radius:14px;padding:26px;width:500px;max-width:96vw;max-height:90vh;overflow-y:auto}
  .loc-chip{display:inline-flex;align-items:center;padding:2px 8px;border-radius:5px;font-size:11px;font-weight:700;border:1px solid}
  @keyframes fi{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}
  .fi{animation:fi .2s ease}
  /* ── Hamburger button (mobile only) ── */
  #hamburger{display:none;position:fixed;top:12px;left:14px;z-index:500;background:#0c0f17;border:1px solid #1e2535;border-radius:8px;width:40px;height:40px;align-items:center;justify-content:center;cursor:pointer;flex-direction:column;gap:5px;padding:0}
  #hamburger span{display:block;width:18px;height:2px;background:#4ade80;border-radius:2px;transition:all .25s}
  #hamburger.open span:nth-child(1){transform:translateY(7px) rotate(45deg)}
  #hamburger.open span:nth-child(2){opacity:0;transform:scaleX(0)}
  #hamburger.open span:nth-child(3){transform:translateY(-7px) rotate(-45deg)}
  /* ── Nav drawer transition ── */
  #sidenav{transition:transform .28s cubic-bezier(.4,0,.2,1)}
  /* ── Nav overlay (mobile) ── */
  #nav-overlay{display:none;position:fixed;inset:0;background:#00000070;z-index:149;backdrop-filter:blur(1px)}
  #nav-overlay.show{display:block}
  @media(max-width:720px){
    #hamburger{display:flex}
    #sidenav{transform:translateX(-100%);z-index:150}
    #sidenav.open{transform:translateX(0)}
    #main-content{margin-left:0!important;padding:64px 14px 50px!important}
  }
  @media print{
    #noprint,#sidenav,#hamburger,#nav-overlay{display:none!important}
    #slip-area{position:fixed;inset:0;padding:30px;background:#fff;z-index:9999}
    body *{visibility:hidden}
    #slip-area,#slip-area *{visibility:visible}
  }
`;

// ─── PASSWORD CONFIG ──────────────────────────────────────────────────────────
// Ganti nilai di bawah ini dengan password yang kamu inginkan
const APP_PASSWORD = "payroll2025";

// ─── LOGIN SCREEN ─────────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState(false);
  const [shake, setShake] = useState(false);

  const tryLogin = () => {
    if (pw === APP_PASSWORD) {
      sessionStorage.setItem("pr_auth", "1");
      onLogin();
    } else {
      setErr(true);
      setShake(true);
      setTimeout(() => setShake(false), 500);
      setTimeout(() => setErr(false), 2500);
      setPw("");
    }
  };

  return (
    <div style={{display:"flex",height:"100vh",alignItems:"center",justifyContent:"center",background:"#111318",fontFamily:"'Syne',sans-serif"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&display=swap');
        @keyframes shake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-8px)}40%,80%{transform:translateX(8px)}}
        .shake{animation:shake .4s ease}
        .login-input{background:#1a1f2e;border:1px solid #2d3748;color:#e2e8f0;border-radius:9px;padding:12px 14px;font-size:15px;outline:none;width:100%;font-family:'Syne',sans-serif;transition:border .15s,box-shadow .15s}
        .login-input:focus{border-color:#4ade80;box-shadow:0 0 0 3px #4ade8020}
        .login-input.err{border-color:#ef4444;box-shadow:0 0 0 3px #ef444420}
        .login-btn{background:#4ade80;color:#0a0d12;border:none;border-radius:9px;padding:13px;font-size:14px;font-weight:800;cursor:pointer;width:100%;font-family:'Syne',sans-serif;transition:background .15s;letter-spacing:.02em}
        .login-btn:hover{background:#22c55e}
      `}</style>
      <div className={shake ? "shake" : ""} style={{background:"#161b27",border:"1px solid #1e2535",borderRadius:16,padding:"36px 32px",width:360,maxWidth:"94vw",textAlign:"center"}}>
        <div style={{fontSize:32,marginBottom:10}}>🔐</div>
        <div style={{fontSize:20,fontWeight:800,color:"#f1f5f9",marginBottom:4}}>PayTrack</div>
        <div style={{fontSize:12,color:"#475569",marginBottom:28,fontWeight:600}}>Sistem Penggajian · Masukkan password untuk lanjut</div>
        <input
          className={`login-input${err?" err":""}`}
          type="password"
          placeholder="Password..."
          value={pw}
          onChange={e=>{setPw(e.target.value);setErr(false);}}
          onKeyDown={e=>e.key==="Enter"&&tryLogin()}
          autoFocus
        />
        {err && <div style={{color:"#ef4444",fontSize:12,marginTop:8,fontWeight:600}}>❌ Password salah, coba lagi</div>}
        <button className="login-btn" style={{marginTop:16}} onClick={tryLogin}>Masuk →</button>
        <div style={{marginTop:20,fontSize:10,color:"#2d3748"}}>Akses terbatas · Hanya untuk pengguna yang berwenang</div>
      </div>
    </div>
  );
}

// ─── APP ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage]     = useState("dashboard");
  const [workers, setW]     = useState([]);
  const [locs, setL]        = useState([]);
  const [payroll, setP]     = useState({});
  const [ready, setReady]   = useState(false);
  const [navOpen, setNavOpen] = useState(false);
  const [authed, setAuthed] = useState(!!sessionStorage.getItem("pr_auth"));
  const navigate = (id) => { setPage(id); setNavOpen(false); };

  // Tampilkan login jika belum autentikasi
  if (!authed) return <LoginScreen onLogin={() => setAuthed(true)} />;

  useEffect(()=>{
    (async()=>{
      const w=await sGet("pr2_w"); if(w) setW(w);
      const l=await sGet("pr2_l"); if(l) setL(l);
      const p=await sGet("pr2_p"); if(p) setP(p);
      setReady(true);
    })();
  },[]);

  const saveW=async v=>{setW(v);await sSave("pr2_w",v);};
  const saveL=async v=>{setL(v);await sSave("pr2_l",v);};
  const saveP=async v=>{setP(v);await sSave("pr2_p",v);};

  if(!ready) return(
    <div style={{display:"flex",height:"100vh",alignItems:"center",justifyContent:"center",background:"#111318",color:"#4ade80",fontFamily:"Syne,sans-serif",fontSize:16,gap:12}}>
      <span style={{animation:"spin 1s linear infinite",display:"inline-block"}}>⟳</span> Memuat sistem…
    </div>
  );

  const NAV=[
    {id:"dashboard",icon:"⬡",label:"Dashboard"},
    {id:"lokasi",   icon:"◉",label:"Lokasi / Proyek"},
    {id:"pekerja",  icon:"⊞",label:"Pekerja"},
    {id:"gaji",     icon:"◈",label:"Penggajian"},
    {id:"riwayat",  icon:"≋",label:"Riwayat"},
    {id:"slip",     icon:"◻",label:"Slip Gaji"},
    {id:"export",   icon:"↧",label:"Export"},
  ];

  return(
    <>
      <style>{CSS}</style>
      <div style={{display:"flex",minHeight:"100vh"}}>

        {/* ── HAMBURGER BUTTON ── */}
        <button id="hamburger" className={navOpen?"open":""} onClick={()=>setNavOpen(o=>!o)} aria-label="Menu">
          <span/><span/><span/>
        </button>

        {/* ── NAV OVERLAY (mobile tap-outside to close) ── */}
        <div id="nav-overlay" className={navOpen?"show":""} onClick={()=>setNavOpen(false)}/>

        {/* SIDEBAR */}
        <nav id="sidenav" className={navOpen?"open":""} style={{width:204,background:"#0c0f17",borderRight:"1px solid #1e2535",display:"flex",flexDirection:"column",padding:"18px 10px",gap:2,position:"fixed",top:0,bottom:0,left:0,zIndex:150}}>
          <div style={{padding:"2px 10px 16px",borderBottom:"1px solid #1e2535",marginBottom:4,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div>
              <div style={{fontSize:18,fontWeight:800,color:"#4ade80",letterSpacing:"-0.4px"}}>PayTrack</div>
              <div style={{fontSize:9,color:"#2d3748",fontWeight:700,letterSpacing:".1em",marginTop:1}}>HARIAN · MINGGUAN</div>
            </div>
          </div>
          {NAV.map(n=>(
            <button key={n.id} onClick={()=>navigate(n.id)} style={{
              display:"flex",alignItems:"center",gap:9,padding:"8px 12px",borderRadius:8,border:"none",
              cursor:"pointer",fontFamily:"'Syne',sans-serif",fontSize:12,fontWeight:page===n.id?800:500,
              transition:"all .15s",textAlign:"left",
              background:page===n.id?"#4ade80":"transparent",
              color:page===n.id?"#0a0d12":"#4a5568",
            }}>
              <span style={{fontSize:13,fontFamily:"monospace"}}>{n.icon}</span>{n.label}
            </button>
          ))}
          <div style={{marginTop:"auto",padding:"10px 8px 0",borderTop:"1px solid #1e2535",fontSize:10,color:"#2d3748",fontWeight:600,lineHeight:1.8}}>
            {workers.filter(w=>w.aktif).length} pekerja aktif<br/>
            {locs.filter(l=>l.aktif).length} lokasi aktif
          </div>
        </nav>

        {/* MAIN */}
        <main id="main-content" style={{flex:1,marginLeft:204,padding:"26px 26px 50px",minHeight:"100vh"}}>
          {page==="dashboard" && <Dashboard workers={workers} locs={locs} payroll={payroll} setPage={navigate}/>}
          {page==="lokasi"    && <Lokasi    locs={locs} saveL={saveL} workers={workers}/>}
          {page==="pekerja"   && <Pekerja   workers={workers} locs={locs} saveW={saveW}/>}
          {page==="gaji"      && <Penggajian workers={workers} locs={locs} payroll={payroll} saveP={saveP}/>}
          {page==="riwayat"   && <Riwayat   workers={workers} locs={locs} payroll={payroll}/>}
          {page==="slip"      && <SlipGaji  workers={workers} locs={locs} payroll={payroll}/>}
          {page==="export"    && <ExportPage workers={workers} locs={locs} payroll={payroll}/>}
        </main>
      </div>
    </>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({workers,locs,payroll,setPage}){
  const active=workers.filter(w=>w.aktif);
  const wkEntries=Object.values(payroll).filter(e=>e.weekStart===thisMonday);
  const totGaji=wkEntries.reduce((a,e)=>a+(e.totalGaji||0),0);
  const totHari=wkEntries.reduce((a,e)=>a+(e.totalHari||0),0);
  const totPaid=wkEntries.filter(e=>e.sudahDibayar).length;
  const totHutang=wkEntries.reduce((a,e)=>a+(e.hutang||0),0);

  const trend=[];
  for(let i=5;i>=0;i--){
    const mon=toISO(getMonday(addDays(new Date(),-(i*7))));
    const es=Object.values(payroll).filter(e=>e.weekStart===mon);
    trend.push({label:i===0?"Ini":`-${i}w`,total:es.reduce((a,e)=>a+(e.totalGaji||0),0)});
  }
  const maxT=Math.max(...trend.map(t=>t.total),1);

  const locSummary=locs.filter(l=>l.aktif).map((l,i)=>{
    const days=wkEntries.reduce((a,e)=>a+(e.breakdown||[]).filter(b=>b.lokasiId===l.id).reduce((x,b)=>x+b.hari,0),0);
    const cnt=wkEntries.filter(e=>(e.breakdown||[]).some(b=>b.lokasiId===l.id)).length;
    return{...l,days,cnt,col:LOC_COLORS[i%LOC_COLORS.length]};
  }).filter(l=>l.days>0);

  return(
    <div className="fi">
      <PH title="Dashboard" sub={`Minggu ini: ${weekLabel(thisMonday)}`}/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:14}}>
        {[
          {l:"Pekerja Aktif",    v:active.length,                   s:`${workers.length} total`,              c:"#4ade80"},
          {l:"Total Gaji Minggu",v:fmt(totGaji),                    s:"minggu berjalan",                      c:"#60a5fa"},
          {l:"Total Hari Kerja", v:`${totHari} hari`,               s:`${wkEntries.length} entri`,            c:"#fb923c"},
          {l:"Sudah Dibayar",    v:`${totPaid}/${wkEntries.length}`,s:"pekerja",                              c:"#fbbf24"},
        ].map(s=>(
          <div key={s.l} className="card" style={{borderTop:`2px solid ${s.c}`}}>
            <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:".08em"}}>{s.l}</div>
            <div className="mono" style={{fontSize:17,fontWeight:600,color:s.c,margin:"6px 0 2px"}}>{s.v}</div>
            <div style={{fontSize:10,color:"#2d3748"}}>{s.s}</div>
          </div>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1.2fr 1fr",gap:12,marginBottom:12}}>
        <div className="card">
          <div style={{fontSize:13,fontWeight:700,marginBottom:14}}>Tren Total Gaji (6 Minggu)</div>
          <div style={{display:"flex",alignItems:"flex-end",gap:7,height:68}}>
            {trend.map((t,i)=>(
              <div key={i} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                <div title={fmt(t.total)} style={{width:"100%",borderRadius:"3px 3px 0 0",height:`${Math.max(t.total>0?3:0,(t.total/maxT)*58)}px`,background:i===5?"#4ade80":"#1e2535",transition:"height .4s"}}/>
                <div style={{fontSize:9,color:"#475569",fontWeight:700}}>{t.label}</div>
              </div>
            ))}
          </div>
        </div>
        <div className="card">
          <div style={{fontSize:13,fontWeight:700,marginBottom:12}}>Rincian Minggu Ini</div>
          {[
            {l:"Gaji Harian",v:fmt(wkEntries.reduce((a,e)=>a+(e.gajiHarian||0)*(e.totalHari||0),0)),c:"#e2e8f0"},
            {l:"Lembur",     v:fmt(wkEntries.reduce((a,e)=>a+(e.lembur||0),0)),                     c:"#60a5fa"},
            {l:"Bonus",      v:fmt(wkEntries.reduce((a,e)=>a+(e.bonus||0),0)),                       c:"#fbbf24"},
            {l:"Hutang",     v:`– ${fmt(totHutang)}`,                                                c:"#f87171"},
          ].map(r=>(
            <div key={r.l} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #1e2535",fontSize:12}}>
              <span style={{color:"#64748b"}}>{r.l}</span>
              <span className="mono" style={{color:r.c,fontSize:11}}>{r.v}</span>
            </div>
          ))}
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div className="card">
          <div style={{fontSize:13,fontWeight:700,marginBottom:12}}>Sebaran Lokasi Minggu Ini</div>
          {locSummary.length>0?locSummary.map(l=>(
            <div key={l.id} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #1a2030"}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <span className="loc-chip" style={{background:l.col.bg,color:l.col.text,borderColor:l.col.border}}>{l.nama}</span>
                <span style={{fontSize:11,color:"#475569"}}>{l.cnt} pekerja · {l.days}h</span>
              </div>
            </div>
          )):<Empty msg="Belum ada data minggu ini"/>}
        </div>
        <div className="card">
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
            <div style={{fontSize:13,fontWeight:700}}>Status Pembayaran</div>
            <button className="btn btn-s" style={{padding:"3px 9px",fontSize:10}} onClick={()=>setPage("gaji")}>Input →</button>
          </div>
          {wkEntries.slice(0,6).map(e=>{
            const w=workers.find(x=>x.id===e.workerId);
            return w?(
              <div key={e.workerId} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"5px 0",borderBottom:"1px solid #1a2030"}}>
                <div>
                  <div style={{fontWeight:600,fontSize:12}}>{w.nama}</div>
                  <span className={`bdg ${e.sudahDibayar?"bdg-g":"bdg-y"}`}>{e.sudahDibayar?"Dibayar":"Belum"}</span>
                </div>
                <div className="mono" style={{fontSize:11,color:"#60a5fa"}}>{fmt(e.totalGaji)}</div>
              </div>
            ):null;
          })}
          {wkEntries.length===0&&<Empty msg="Belum ada penggajian minggu ini"/>}
        </div>
      </div>
    </div>
  );
}

// ─── LOKASI ───────────────────────────────────────────────────────────────────
function Lokasi({locs,saveL,workers}){
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState({nama:"",deskripsi:"",aktif:true});
  const [err,setErr]=useState("");

  const openAdd=()=>{setEditId(null);setForm({nama:"",deskripsi:"",aktif:true});setErr("");setModal(true);};
  const openEdit=l=>{setEditId(l.id);setForm({nama:l.nama,deskripsi:l.deskripsi||"",aktif:l.aktif});setErr("");setModal(true);};
  const save=async()=>{
    if(!form.nama.trim()){setErr("Nama wajib diisi");return;}
    const updated=editId?locs.map(l=>l.id===editId?{...l,...form}:l):[...locs,{id:uid(),...form}];
    await saveL(updated);setModal(false);
  };
  const toggle=async id=>await saveL(locs.map(l=>l.id===id?{...l,aktif:!l.aktif}:l));
  const remove=async id=>{
    if(workers.some(w=>w.lokasiId===id)&&!confirm("Pekerja masih di lokasi ini. Hapus tetap?"))return;
    if(!confirm("Hapus lokasi ini?"))return;
    await saveL(locs.filter(l=>l.id!==id));
  };

  return(
    <div className="fi">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
        <PH title="Lokasi / Proyek" sub="Kelola lokasi kerja & proyek aktif"/>
        <button className="btn btn-g" onClick={openAdd}>+ Tambah Lokasi</button>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(270px,1fr))",gap:12}}>
        {locs.map((l,i)=>{
          const col=LOC_COLORS[i%LOC_COLORS.length];
          const wc=workers.filter(w=>w.lokasiId===l.id&&w.aktif).length;
          return(
            <div key={l.id} className="card" style={{borderLeft:`3px solid ${col.text}`,opacity:l.aktif?1:.5}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div>
                  <div style={{fontWeight:800,fontSize:15,color:col.text}}>{l.nama}</div>
                  <div style={{fontSize:11,color:"#475569",marginTop:2}}>{l.deskripsi||"Tidak ada deskripsi"}</div>
                </div>
                <span className={`bdg ${l.aktif?"bdg-g":"bdg-r"}`}>{l.aktif?"Aktif":"Selesai"}</span>
              </div>
              <div style={{marginTop:10,fontSize:11,color:"#64748b"}}>
                <span className="loc-chip" style={{background:col.bg,color:col.text,borderColor:col.border}}>{wc} pekerja aktif</span>
              </div>
              <div style={{display:"flex",gap:6,marginTop:12}}>
                <button className="btn btn-s" style={{padding:"4px 9px",fontSize:11}} onClick={()=>openEdit(l)}>Edit</button>
                <button className="btn btn-s" style={{padding:"4px 9px",fontSize:11}} onClick={()=>toggle(l.id)}>{l.aktif?"Selesaikan":"Aktifkan"}</button>
                <button className="btn btn-r" style={{padding:"4px 9px",fontSize:11}} onClick={()=>remove(l.id)}>Hapus</button>
              </div>
            </div>
          );
        })}
        {locs.length===0&&<div className="card" style={{gridColumn:"1/-1",padding:40,textAlign:"center"}}><Empty msg="Belum ada lokasi. Tambahkan proyek pertama."/></div>}
      </div>
      {modal&&(
        <div className="overlay" onClick={e=>{if(e.target===e.currentTarget)setModal(false)}}>
          <div className="modal">
            <div style={{fontSize:16,fontWeight:800,marginBottom:16}}>{editId?"Edit Lokasi":"Tambah Lokasi Baru"}</div>
            <div style={{display:"grid",gap:12}}>
              <FLD label="Nama Lokasi / Proyek *">
                <input value={form.nama} onChange={e=>setForm(f=>({...f,nama:e.target.value}))} placeholder="Proyek A, Gudang B, dll"/>
                {err&&<div style={{color:"#ef4444",fontSize:11,marginTop:3}}>{err}</div>}
              </FLD>
              <FLD label="Deskripsi (opsional)">
                <input value={form.deskripsi} onChange={e=>setForm(f=>({...f,deskripsi:e.target.value}))} placeholder="Alamat, keterangan, dll"/>
              </FLD>
            </div>
            <div style={{display:"flex",gap:8,marginTop:18,justifyContent:"flex-end"}}>
              <button className="btn btn-s" onClick={()=>setModal(false)}>Batal</button>
              <button className="btn btn-g" onClick={save}>{editId?"Simpan":"Tambah"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── PEKERJA ──────────────────────────────────────────────────────────────────
function Pekerja({workers,locs,saveW}){
  const EMP={nama:"",jabatan:"",gajiHarian:"",lokasiId:"",tanggalMasuk:"",aktif:true,catatan:""};
  const [modal,setModal]=useState(false);
  const [editId,setEditId]=useState(null);
  const [form,setForm]=useState(EMP);
  const [search,setSearch]=useState("");
  const [fLoc,setFLoc]=useState("all");
  const [showOff,setShowOff]=useState(false);
  const [err,setErr]=useState("");

  const openAdd=()=>{setEditId(null);setForm({...EMP,lokasiId:locs.find(l=>l.aktif)?.id||""});setErr("");setModal(true);};
  const openEdit=w=>{setEditId(w.id);setForm({...w,gajiHarian:String(w.gajiHarian)});setErr("");setModal(true);};
  const save=async()=>{
    if(!form.nama.trim()){setErr("Nama wajib diisi");return;}
    const entry={...form,gajiHarian:num(form.gajiHarian)};
    const updated=editId?workers.map(w=>w.id===editId?{...entry,id:editId}:w):[...workers,{...entry,id:uid()}];
    await saveW(updated);setModal(false);
  };
  const toggle=async id=>await saveW(workers.map(w=>w.id===id?{...w,aktif:!w.aktif}:w));
  const remove=async id=>{if(!confirm("Hapus pekerja ini?"))return;await saveW(workers.filter(w=>w.id!==id));};

  const filtered=workers
    .filter(w=>showOff?true:w.aktif)
    .filter(w=>fLoc==="all"?true:w.lokasiId===fLoc)
    .filter(w=>w.nama.toLowerCase().includes(search.toLowerCase())||(w.jabatan||"").toLowerCase().includes(search.toLowerCase()));

  return(
    <div className="fi">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:20}}>
        <PH title="Data Pekerja" sub={`${workers.filter(w=>w.aktif).length} aktif dari ${workers.length} total`}/>
        <button className="btn btn-g" onClick={openAdd}>+ Tambah Pekerja</button>
      </div>
      <div style={{display:"flex",gap:10,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <input placeholder="Cari nama / jabatan…" value={search} onChange={e=>setSearch(e.target.value)} style={{maxWidth:230}}/>
        <select value={fLoc} onChange={e=>setFLoc(e.target.value)} style={{maxWidth:170}}>
          <option value="all">Semua Lokasi</option>
          {locs.map(l=><option key={l.id} value={l.id}>{l.nama}</option>)}
        </select>
        <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:"#64748b",cursor:"pointer"}}>
          <input type="checkbox" checked={showOff} onChange={e=>setShowOff(e.target.checked)}/> Non-Aktif
        </label>
      </div>
      <div className="t-wrap">
        <table>
          <thead><tr>
            <th>#</th><th>Nama</th><th>Jabatan</th><th>Lokasi Utama</th>
            <th>Gaji / Hari</th><th>Tgl Masuk</th><th>Status</th><th>Aksi</th>
          </tr></thead>
          <tbody>
            {filtered.map((w,i)=>{
              const loc=locs.find(l=>l.id===w.lokasiId);
              const ci=locs.findIndex(l=>l.id===w.lokasiId);
              const col=ci>=0?LOC_COLORS[ci%LOC_COLORS.length]:null;
              return(
                <tr key={w.id} className="hov">
                  <td style={{color:"#2d3748",fontSize:11}}>{i+1}</td>
                  <td style={{fontWeight:700}}>{w.nama}</td>
                  <td style={{color:"#94a3b8",fontSize:12}}>{w.jabatan||"—"}</td>
                  <td>{loc&&col?<span className="loc-chip" style={{background:col.bg,color:col.text,borderColor:col.border}}>{loc.nama}</span>:<span style={{color:"#2d3748"}}>—</span>}</td>
                  <td className="mono" style={{color:"#4ade80",fontSize:12}}>{fmt(w.gajiHarian)}</td>
                  <td style={{color:"#475569",fontSize:11}}>{w.tanggalMasuk||"—"}</td>
                  <td><span className={`bdg ${w.aktif?"bdg-g":"bdg-r"}`}>{w.aktif?"Aktif":"Non-Aktif"}</span></td>
                  <td>
                    <div style={{display:"flex",gap:5}}>
                      <button className="btn btn-s" style={{padding:"3px 9px",fontSize:10}} onClick={()=>openEdit(w)}>Edit</button>
                      <button className="btn btn-s" style={{padding:"3px 9px",fontSize:10}} onClick={()=>toggle(w.id)}>{w.aktif?"Nonaktifkan":"Aktifkan"}</button>
                      <button className="btn btn-r" style={{padding:"3px 9px",fontSize:10}} onClick={()=>remove(w.id)}>Hapus</button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {filtered.length===0&&<tr><td colSpan={8}><Empty msg="Tidak ada pekerja ditemukan"/></td></tr>}
          </tbody>
        </table>
      </div>
      {modal&&(
        <div className="overlay" onClick={e=>{if(e.target===e.currentTarget)setModal(false)}}>
          <div className="modal">
            <div style={{fontSize:16,fontWeight:800,marginBottom:16}}>{editId?"Edit Pekerja":"Tambah Pekerja"}</div>
            <div style={{display:"grid",gap:12}}>
              <FLD label="Nama Lengkap *">
                <input value={form.nama} onChange={e=>setForm(f=>({...f,nama:e.target.value}))} placeholder="Nama pekerja"/>
                {err&&<div style={{color:"#ef4444",fontSize:11,marginTop:3}}>{err}</div>}
              </FLD>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <FLD label="Jabatan"><input value={form.jabatan} onChange={e=>setForm(f=>({...f,jabatan:e.target.value}))} placeholder="Teknisi, Kasir…"/></FLD>
                <FLD label="Gaji per Hari (Rp)"><input className="mono" value={form.gajiHarian} onChange={e=>setForm(f=>({...f,gajiHarian:e.target.value}))} placeholder="150000"/></FLD>
              </div>
              <FLD label="Lokasi Utama">
                <select value={form.lokasiId} onChange={e=>setForm(f=>({...f,lokasiId:e.target.value}))}>
                  <option value="">— Pilih Lokasi —</option>
                  {locs.filter(l=>l.aktif).map(l=><option key={l.id} value={l.id}>{l.nama}</option>)}
                </select>
              </FLD>
              <FLD label="Tanggal Masuk"><input type="date" value={form.tanggalMasuk} onChange={e=>setForm(f=>({...f,tanggalMasuk:e.target.value}))} style={{colorScheme:"dark"}}/></FLD>
              <FLD label="Catatan"><input value={form.catatan} onChange={e=>setForm(f=>({...f,catatan:e.target.value}))} placeholder="Opsional"/></FLD>
            </div>
            <div style={{display:"flex",gap:8,marginTop:18,justifyContent:"flex-end"}}>
              <button className="btn btn-s" onClick={()=>setModal(false)}>Batal</button>
              <button className="btn btn-g" onClick={save}>{editId?"Simpan":"Tambah"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── PENGGAJIAN ───────────────────────────────────────────────────────────────
function Penggajian({workers,locs,payroll,saveP}){
  const [weekStart,setWeekStart]=useState(thisMonday);
  const [local,setLocal]=useState({});
  const [fLoc,setFLoc]=useState("all");
  const [saved,setSaved]=useState(false);
  const [saving,setSaving]=useState(false);

  const active=workers.filter(w=>w.aktif);

  useEffect(()=>{
    const init={};
    active.forEach(w=>{
      const key=`${w.id}_${weekStart}`;
      if(payroll[key]){
        init[w.id]={
          breakdown:payroll[key].breakdown||[{lokasiId:w.lokasiId,hari:0}],
          lembur:String(payroll[key].lembur||0),
          bonus:String(payroll[key].bonus||0),
          hutang:String(payroll[key].hutang||0),
          keterangan:payroll[key].keterangan||"",
          sudahDibayar:payroll[key].sudahDibayar||false,
        };
      } else {
        init[w.id]={
          breakdown:[{lokasiId:w.lokasiId||"",hari:0}],
          lembur:"0",bonus:"0",hutang:"0",keterangan:"",sudahDibayar:false,
        };
      }
    });
    setLocal(init);setSaved(false);
  // eslint-disable-next-line
  },[weekStart,workers.length]);

  const upd=(wid,field,val)=>{setLocal(l=>({...l,[wid]:{...l[wid],[field]:val}}));setSaved(false);};
  const updBD=(wid,idx,field,val)=>{
    setLocal(l=>{
      const bd=[...(l[wid]?.breakdown||[])];
      bd[idx]={...bd[idx],[field]:val};
      return{...l,[wid]:{...l[wid],breakdown:bd}};
    });setSaved(false);
  };
  const addBD=wid=>{
    setLocal(l=>{
      const bd=[...(l[wid]?.breakdown||[]),{lokasiId:"",hari:0}];
      return{...l,[wid]:{...l[wid],breakdown:bd}};
    });
  };
  const removeBD=(wid,idx)=>{
    setLocal(l=>{
      const bd=(l[wid]?.breakdown||[]).filter((_,i)=>i!==idx);
      return{...l,[wid]:{...l[wid],breakdown:bd.length?bd:[{lokasiId:"",hari:0}]}};
    });
  };

  const totalH=wid=>(local[wid]?.breakdown||[]).reduce((a,b)=>a+num(b.hari),0);
  const calcG=(w,e)=>{
    const th=(e?.breakdown||[]).reduce((a,b)=>a+num(b.hari),0);
    return w.gajiHarian*th+num(e?.lembur)+num(e?.bonus)-num(e?.hutang);
  };

  const handleSave=async()=>{
    setSaving(true);
    const upd2={...payroll};
    active.forEach(w=>{
      const e=local[w.id]||{};
      const th=totalH(w.id);
      upd2[`${w.id}_${weekStart}`]={
        workerId:w.id,weekStart,gajiHarian:w.gajiHarian,
        breakdown:(e.breakdown||[]).map(b=>({...b,hari:num(b.hari)})),
        totalHari:th,lembur:num(e.lembur),bonus:num(e.bonus),hutang:num(e.hutang),
        keterangan:e.keterangan||"",sudahDibayar:e.sudahDibayar||false,
        totalGaji:calcG(w,e),
      };
    });
    await saveP(upd2);setSaving(false);setSaved(true);setTimeout(()=>setSaved(false),2500);
  };

  const markAll=async paid=>{
    const upd2={...payroll};
    active.forEach(w=>{
      const k=`${w.id}_${weekStart}`;
      if(upd2[k])upd2[k]={...upd2[k],sudahDibayar:paid};
    });
    setLocal(l=>{
      const nl={...l};
      active.forEach(w=>{if(nl[w.id])nl[w.id]={...nl[w.id],sudahDibayar:paid};});
      return nl;
    });
    await saveP(upd2);
  };

  const prevW=()=>setWeekStart(toISO(getMonday(addDays(new Date(weekStart+"T00:00:00"),-7))));
  const nextW=()=>setWeekStart(toISO(getMonday(addDays(new Date(weekStart+"T00:00:00"), 7))));

  const filtered=active.filter(w=>fLoc==="all"?true:w.lokasiId===fLoc);
  const totG=filtered.reduce((a,w)=>a+calcG(w,local[w.id]),0);
  const totH=filtered.reduce((a,w)=>a+totalH(w.id),0);
  const totPd=filtered.filter(w=>(local[w.id]||{}).sudahDibayar).length;

  return(
    <div className="fi">
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14,flexWrap:"wrap",gap:10}}>
        <PH title="Input Penggajian Mingguan" sub="Gaji = Gaji/Hari × Jumlah Hari Kerja"/>
        <div style={{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap"}}>
          {saved&&<span style={{color:"#4ade80",fontSize:12,fontWeight:700}}>✓ Tersimpan</span>}
          <button className="btn btn-s" style={{fontSize:11,padding:"5px 10px"}} onClick={()=>markAll(true)}>✓ Semua Dibayar</button>
          <button className="btn btn-s" style={{fontSize:11,padding:"5px 10px"}} onClick={()=>markAll(false)}>↺ Reset Status</button>
          <button className="btn btn-g" onClick={handleSave} disabled={saving}>{saving?"…":"💾 Simpan"}</button>
        </div>
      </div>

      {/* Week nav */}
      <div style={{display:"flex",gap:8,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
        <button className="btn btn-s" style={{padding:"5px 11px",fontSize:12}} onClick={prevW}>‹</button>
        <div className="card" style={{padding:"7px 16px",display:"inline-flex",alignItems:"center",gap:8,fontSize:13,fontWeight:700}}>
          <span style={{color:"#4ade80",fontSize:14}}>◈</span>
          {weekLabel(weekStart)}
          {weekStart===thisMonday&&<span className="bdg bdg-g" style={{fontSize:9}}>INI</span>}
        </div>
        <button className="btn btn-s" style={{padding:"5px 11px",fontSize:12}} onClick={nextW}>›</button>
        {weekStart!==thisMonday&&<button className="btn btn-s" style={{padding:"5px 10px",fontSize:10}} onClick={()=>setWeekStart(thisMonday)}>Minggu Ini</button>}
      </div>

      {/* Filter */}
      <div style={{display:"flex",gap:10,marginBottom:12,alignItems:"center",flexWrap:"wrap"}}>
        <select value={fLoc} onChange={e=>setFLoc(e.target.value)} style={{maxWidth:200}}>
          <option value="all">Semua Lokasi</option>
          {locs.map(l=><option key={l.id} value={l.id}>{l.nama}</option>)}
        </select>
        <span style={{fontSize:11,color:"#475569"}}>{filtered.length} pekerja · {totH} hari · {totPd}/{filtered.length} dibayar · <span className="mono" style={{color:"#4ade80"}}>{fmt(totG)}</span></span>
      </div>

      {active.length===0?(
        <div className="card" style={{textAlign:"center",padding:40}}><Empty msg="Belum ada pekerja aktif"/></div>
      ):(
        <>
          <div style={{overflowX:"auto"}}>
            <div className="t-wrap">
              <table style={{minWidth:980}}>
                <thead><tr>
                  <th style={{minWidth:140}}>Nama</th>
                  <th style={{minWidth:90}}>Gaji/Hari</th>
                  <th style={{minWidth:280}}>
                    Lokasi &amp; Hari Kerja
                    <div style={{fontSize:9,color:"#2d3748",fontWeight:500,marginTop:1}}>Tambah baris jika pindah lokasi sementara</div>
                  </th>
                  <th style={{minWidth:55,textAlign:"center"}}>Total Hari</th>
                  <th style={{minWidth:120}}>Lembur (Rp)</th>
                  <th style={{minWidth:120}}>Bonus (Rp)</th>
                  <th style={{minWidth:130}}>Hutang (Rp)</th>
                  <th style={{minWidth:130}}>Total Gaji</th>
                  <th style={{minWidth:65,textAlign:"center"}}>Dibayar</th>
                  <th style={{minWidth:100}}>Ket.</th>
                </tr></thead>
                <tbody>
                  {filtered.map(w=>{
                    const e=local[w.id]||{breakdown:[{lokasiId:w.lokasiId,hari:0}],lembur:"0",bonus:"0",hutang:"0"};
                    const th=totalH(w.id);
                    const total=calcG(w,e);
                    const loc=locs.find(l=>l.id===w.lokasiId);
                    const ci=locs.findIndex(l=>l.id===w.lokasiId);
                    const col=ci>=0?LOC_COLORS[ci%LOC_COLORS.length]:null;
                    return(
                      <tr key={w.id} className="hov" style={{verticalAlign:"top"}}>
                        <td style={{paddingTop:10}}>
                          <div style={{fontWeight:700,fontSize:12}}>{w.nama}</div>
                          <div style={{fontSize:10,color:"#475569",marginTop:1}}>{w.jabatan}</div>
                          {loc&&col&&<span className="loc-chip" style={{background:col.bg,color:col.text,borderColor:col.border,marginTop:5,fontSize:9}}>{loc.nama}</span>}
                        </td>
                        <td className="mono" style={{color:"#4ade80",fontSize:12,paddingTop:12,whiteSpace:"nowrap"}}>{fmt(w.gajiHarian)}</td>
                        {/* ── LOKASI BREAKDOWN ── */}
                        <td style={{padding:"8px 10px"}}>
                          <div style={{display:"flex",flexDirection:"column",gap:5}}>
                            {(e.breakdown||[]).map((b,idx)=>{
                              const bci=locs.findIndex(l=>l.id===b.lokasiId);
                              const bcol=bci>=0?LOC_COLORS[bci%LOC_COLORS.length]:null;
                              return(
                                <div key={idx} style={{display:"flex",alignItems:"center",gap:5}}>
                                  <select
                                    value={b.lokasiId}
                                    onChange={ev=>updBD(w.id,idx,"lokasiId",ev.target.value)}
                                    style={{fontSize:11,padding:"4px 7px",width:138,border:`1px solid ${bcol?bcol.border:"#2d3748"}`,color:bcol?bcol.text:"#94a3b8",background:"#111318"}}
                                  >
                                    <option value="">— Lokasi —</option>
                                    {locs.filter(l=>l.aktif).map(l=><option key={l.id} value={l.id}>{l.nama}</option>)}
                                  </select>
                                  <input
                                    type="number" min="0" max="7"
                                    value={b.hari}
                                    onChange={ev=>updBD(w.id,idx,"hari",Math.min(7,Math.max(0,parseInt(ev.target.value)||0)))}
                                    style={{width:42,textAlign:"center",padding:"4px 5px",fontSize:12}}
                                  />
                                  <span style={{fontSize:10,color:"#475569",flexShrink:0}}>hr</span>
                                  {(e.breakdown||[]).length>1&&(
                                    <button onClick={()=>removeBD(w.id,idx)} title="Hapus baris" style={{background:"#ef444420",border:"none",color:"#ef4444",borderRadius:4,width:18,height:18,cursor:"pointer",fontSize:12,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>×</button>
                                  )}
                                </div>
                              );
                            })}
                            <button onClick={()=>addBD(w.id)} style={{alignSelf:"flex-start",background:"transparent",border:"1px dashed #4ade8040",color:"#4ade80",borderRadius:5,padding:"2px 8px",fontSize:10,cursor:"pointer",fontFamily:"Syne,sans-serif",fontWeight:700}}>
                              + Lokasi Lain
                            </button>
                          </div>
                        </td>
                        <td style={{textAlign:"center",paddingTop:12}}>
                          <span className="mono" style={{fontSize:18,fontWeight:700,color:th>0?"#fbbf24":"#2d3748"}}>{th}</span>
                        </td>
                        <td style={{paddingTop:9}}>
                          <input className="mono" type="number" min="0" value={e.lembur} onChange={ev=>upd(w.id,"lembur",ev.target.value)} style={{width:110,fontSize:12,color:"#60a5fa"}}/>
                        </td>
                        <td style={{paddingTop:9}}>
                          <input className="mono" type="number" min="0" value={e.bonus} onChange={ev=>upd(w.id,"bonus",ev.target.value)} style={{width:110,fontSize:12,color:"#fbbf24"}}/>
                        </td>
                        <td style={{paddingTop:9}}>
                          <input className="mono" type="number" min="0" value={e.hutang} onChange={ev=>upd(w.id,"hutang",ev.target.value)} style={{width:120,fontSize:12,color:"#f87171"}}/>
                        </td>
                        <td style={{paddingTop:10}}>
                          <div className="mono" style={{fontSize:13,fontWeight:700,color:total>=0?"#4ade80":"#f87171",whiteSpace:"nowrap"}}>{fmt(total)}</div>
                          {th>0&&<div style={{fontSize:10,color:"#475569",marginTop:2}}>{fmt(w.gajiHarian)} × {th}h</div>}
                        </td>
                        <td style={{textAlign:"center",paddingTop:13}}>
                          <input type="checkbox" checked={e.sudahDibayar||false} onChange={ev=>upd(w.id,"sudahDibayar",ev.target.checked)}/>
                        </td>
                        <td style={{paddingTop:9}}>
                          <input value={e.keterangan||""} onChange={ev=>upd(w.id,"keterangan",ev.target.value)} placeholder="—" style={{fontSize:11,width:90}}/>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
          {/* Summary */}
          <div style={{display:"flex",gap:10,marginTop:12,flexWrap:"wrap"}}>
            {[
              {l:"Total Gaji",    v:fmt(totG),                                                             c:"#4ade80"},
              {l:"Total Hari",    v:`${totH}h`,                                                            c:"#e2e8f0"},
              {l:"Lembur",        v:fmt(filtered.reduce((a,w)=>a+num((local[w.id]||{}).lembur),0)),        c:"#60a5fa"},
              {l:"Bonus",         v:fmt(filtered.reduce((a,w)=>a+num((local[w.id]||{}).bonus),0)),         c:"#fbbf24"},
              {l:"Hutang",        v:fmt(filtered.reduce((a,w)=>a+num((local[w.id]||{}).hutang),0)),        c:"#f87171"},
              {l:"Dibayar",       v:`${totPd}/${filtered.length}`,                                         c:"#a78bfa"},
            ].map(s=>(
              <div key={s.l} className="card" style={{flex:"1 0 90px",padding:"10px 13px"}}>
                <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:".07em"}}>{s.l}</div>
                <div className="mono" style={{fontSize:13,fontWeight:700,color:s.c,marginTop:4}}>{s.v}</div>
              </div>
            ))}
          </div>
        </>
      )}
    </div>
  );
}

// ─── RIWAYAT ──────────────────────────────────────────────────────────────────
function Riwayat({workers,locs,payroll}){
  const allWeeks=[...new Set(Object.values(payroll).map(e=>e.weekStart).filter(Boolean))].sort().reverse();
  const [fWeek,setFWeek]=useState("all");
  const [fW,setFW]=useState("all");
  const [fL,setFL]=useState("all");
  const [pg,setPg]=useState(1);
  const PER=30;

  const entries=Object.entries(payroll)
    .map(([k,v])=>({key:k,...v}))
    .filter(e=>{
      if(!e.weekStart)return false;
      if(fWeek!=="all"&&e.weekStart!==fWeek)return false;
      if(fW!=="all"&&e.workerId!==fW)return false;
      if(fL!=="all"&&!(e.breakdown||[]).some(b=>b.lokasiId===fL))return false;
      return true;
    })
    .sort((a,b)=>b.weekStart.localeCompare(a.weekStart));

  const paged=entries.slice((pg-1)*PER,pg*PER);
  const pages=Math.max(1,Math.ceil(entries.length/PER));
  const totG=entries.reduce((a,e)=>a+(e.totalGaji||0),0);

  return(
    <div className="fi">
      <PH title="Riwayat Penggajian" sub="Histori gaji semua pekerja"/>
      <div style={{display:"flex",gap:10,marginBottom:14,flexWrap:"wrap"}}>
        <div style={{minWidth:220}}>
          <div className="label">Minggu</div>
          <select value={fWeek} onChange={e=>{setFWeek(e.target.value);setPg(1);}}>
            <option value="all">Semua Minggu</option>
            {allWeeks.map(w=><option key={w} value={w}>{weekLabel(w)}</option>)}
          </select>
        </div>
        <div style={{minWidth:180}}>
          <div className="label">Pekerja</div>
          <select value={fW} onChange={e=>{setFW(e.target.value);setPg(1);}}>
            <option value="all">Semua Pekerja</option>
            {workers.map(w=><option key={w.id} value={w.id}>{w.nama}</option>)}
          </select>
        </div>
        <div style={{minWidth:170}}>
          <div className="label">Lokasi</div>
          <select value={fL} onChange={e=>{setFL(e.target.value);setPg(1);}}>
            <option value="all">Semua Lokasi</option>
            {locs.map(l=><option key={l.id} value={l.id}>{l.nama}</option>)}
          </select>
        </div>
        <div style={{marginTop:22,fontSize:11,color:"#475569"}}>
          {entries.length} entri · <span className="mono" style={{color:"#4ade80"}}>{fmt(totG)}</span>
        </div>
      </div>
      <div className="t-wrap">
        <table>
          <thead><tr>
            <th>Minggu</th><th>Nama</th><th>Lokasi Kerja (Hari)</th>
            <th style={{textAlign:"center"}}>Total Hari</th>
            <th>Gaji/Hari</th><th>Lembur</th><th>Bonus</th><th>Hutang</th>
            <th>Total Gaji</th><th>Status</th>
          </tr></thead>
          <tbody>
            {paged.map(e=>{
              const w=workers.find(x=>x.id===e.workerId);
              return(
                <tr key={e.key} className="hov">
                  <td style={{whiteSpace:"nowrap",fontSize:11,color:"#64748b"}}>{e.weekStart?weekLabel(e.weekStart):"—"}</td>
                  <td><div style={{fontWeight:700,fontSize:12}}>{w?.nama||"—"}</div><div style={{fontSize:10,color:"#475569"}}>{w?.jabatan}</div></td>
                  <td>
                    <div style={{display:"flex",flexWrap:"wrap",gap:3}}>
                      {(e.breakdown||[]).filter(b=>b.hari>0).map((b,i)=>{
                        const l=locs.find(x=>x.id===b.lokasiId);
                        const ci=locs.findIndex(x=>x.id===b.lokasiId);
                        const c=ci>=0?LOC_COLORS[ci%LOC_COLORS.length]:null;
                        return l&&c?<span key={i} className="loc-chip" style={{background:c.bg,color:c.text,borderColor:c.border,fontSize:10}}>{l.nama} {b.hari}h</span>:null;
                      })}
                    </div>
                  </td>
                  <td style={{textAlign:"center"}} className="mono">{e.totalHari||0}</td>
                  <td className="mono" style={{fontSize:11,color:"#4ade80"}}>{fmt(e.gajiHarian)}</td>
                  <td className="mono" style={{fontSize:11,color:"#60a5fa"}}>{e.lembur?fmt(e.lembur):"—"}</td>
                  <td className="mono" style={{fontSize:11,color:"#fbbf24"}}>{e.bonus?fmt(e.bonus):"—"}</td>
                  <td className="mono" style={{fontSize:11,color:"#f87171"}}>{e.hutang?`–${fmt(e.hutang)}`:"—"}</td>
                  <td className="mono" style={{color:"#4ade80",fontWeight:700}}>{fmt(e.totalGaji)}</td>
                  <td><span className={`bdg ${e.sudahDibayar?"bdg-g":"bdg-y"}`}>{e.sudahDibayar?"Dibayar":"Belum"}</span></td>
                </tr>
              );
            })}
            {entries.length===0&&<tr><td colSpan={10}><Empty msg="Tidak ada data"/></td></tr>}
          </tbody>
        </table>
      </div>
      {pages>1&&(
        <div style={{display:"flex",gap:6,marginTop:10,justifyContent:"center",alignItems:"center"}}>
          <button className="btn btn-s" style={{padding:"4px 10px"}} onClick={()=>setPg(p=>Math.max(1,p-1))} disabled={pg===1}>‹</button>
          <span style={{fontSize:12,color:"#475569"}}>Hal {pg}/{pages}</span>
          <button className="btn btn-s" style={{padding:"4px 10px"}} onClick={()=>setPg(p=>Math.min(pages,p+1))} disabled={pg===pages}>›</button>
        </div>
      )}
    </div>
  );
}

// ─── SLIP GAJI ────────────────────────────────────────────────────────────────
function SlipGaji({workers,locs,payroll}){
  const allWeeks=[...new Set(Object.values(payroll).map(e=>e.weekStart).filter(Boolean))].sort().reverse();
  const [selW,setSelW]=useState("");
  const [selWk,setSelWk]=useState(allWeeks[0]||thisMonday);
  const worker=workers.find(w=>w.id===selW);
  const entry=selW?payroll[`${selW}_${selWk}`]:null;
  const locName=id=>locs.find(l=>l.id===id)?.nama||"—";
  const locCol=id=>{const ci=locs.findIndex(l=>l.id===id);return ci>=0?LOC_COLORS[ci%LOC_COLORS.length]:null;};

  return(
    <div className="fi">
      <div id="noprint">
        <PH title="Slip Gaji Mingguan" sub="Lihat dan cetak slip gaji"/>
        <div style={{display:"flex",gap:12,marginBottom:22,flexWrap:"wrap",alignItems:"flex-end"}}>
          <div style={{minWidth:220}}>
            <div className="label">Pekerja</div>
            <select value={selW} onChange={e=>setSelW(e.target.value)}>
              <option value="">— Pilih Pekerja —</option>
              {workers.map(w=><option key={w.id} value={w.id}>{w.nama}</option>)}
            </select>
          </div>
          <div style={{minWidth:260}}>
            <div className="label">Minggu</div>
            <select value={selWk} onChange={e=>setSelWk(e.target.value)}>
              <option value={thisMonday}>{weekLabel(thisMonday)} (Minggu ini)</option>
              {allWeeks.filter(w=>w!==thisMonday).map(w=><option key={w} value={w}>{weekLabel(w)}</option>)}
            </select>
          </div>
          {entry&&worker&&<button className="btn btn-g" onClick={()=>window.print()}>🖨️ Cetak</button>}
        </div>
        {!selW&&<div className="card" style={{textAlign:"center",padding:40}}><Empty msg="Pilih pekerja untuk melihat slip gaji"/></div>}
        {selW&&!entry&&<div className="card" style={{textAlign:"center",padding:40}}><Empty msg={`Belum ada data gaji ${worker?.nama} minggu ini`}/></div>}
      </div>

      {entry&&worker&&(
        <div id="slip-area">
          <div style={{maxWidth:560,background:"#fff",color:"#0f172a",borderRadius:14,overflow:"hidden",boxShadow:"0 20px 60px #00000080",fontFamily:"Syne,sans-serif"}}>
            <div style={{background:"#0c0f17",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontSize:20,fontWeight:800,color:"#4ade80"}}>PayTrack</div>
                <div style={{fontSize:9,color:"#475569",fontWeight:700,letterSpacing:".1em",marginTop:1}}>SLIP GAJI MINGGUAN</div>
              </div>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:13,fontWeight:700,color:"#e2e8f0"}}>{weekLabel(entry.weekStart)}</div>
                <div style={{fontSize:10,color:"#475569",marginTop:1}}>Periode Pembayaran</div>
              </div>
            </div>
            <div style={{padding:"13px 24px",background:"#f8fafc",borderBottom:"2px solid #e2e8f0"}}>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
                {[
                  {l:"Nama",v:worker.nama},
                  {l:"Jabatan",v:worker.jabatan||"—"},
                  {l:"Lokasi Utama",v:locName(worker.lokasiId)},
                ].map(f=>(
                  <div key={f.l}>
                    <div style={{fontSize:9,color:"#64748b",fontWeight:700,textTransform:"uppercase",letterSpacing:".06em"}}>{f.l}</div>
                    <div style={{fontSize:13,fontWeight:800,color:"#0f172a",marginTop:2}}>{f.v}</div>
                  </div>
                ))}
              </div>
            </div>
            <div style={{padding:"18px 24px"}}>
              {/* Breakdown lokasi */}
              <div style={{fontSize:10,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".07em",marginBottom:8}}>Rincian Hari Kerja</div>
              {(entry.breakdown||[]).filter(b=>b.hari>0).map((b,i)=>{
                const c=locCol(b.lokasiId);
                return(
                  <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid #f1f5f9"}}>
                    <span style={{display:"inline-flex",alignItems:"center",gap:6,fontSize:13}}>
                      {c&&<span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:c.text}}/>}
                      {locName(b.lokasiId)} — <b>{b.hari} hari</b>
                    </span>
                    <span style={{fontFamily:"monospace",fontWeight:700,color:"#0f172a"}}>{fmt(b.hari*(entry.gajiHarian||0))}</span>
                  </div>
                );
              })}
              <div style={{display:"flex",justifyContent:"space-between",padding:"8px 10px",background:"#f8fafc",borderRadius:6,marginTop:5}}>
                <span style={{fontWeight:700,fontSize:13}}>Subtotal {entry.totalHari} Hari × {fmt(entry.gajiHarian)}</span>
                <span style={{fontFamily:"monospace",fontWeight:800,color:"#1e40af"}}>{fmt((entry.gajiHarian||0)*(entry.totalHari||0))}</span>
              </div>
              {(entry.lembur>0||entry.bonus>0||entry.hutang>0)&&(
                <>
                  <div style={{fontSize:10,fontWeight:700,color:"#475569",textTransform:"uppercase",letterSpacing:".07em",margin:"13px 0 8px"}}>Komponen Lain</div>
                  {entry.lembur>0&&<SR label="Lembur" val={fmt(entry.lembur)} c="#1d4ed8"/>}
                  {entry.bonus>0 &&<SR label="Bonus"  val={fmt(entry.bonus)}  c="#b45309"/>}
                  {entry.hutang>0&&<SR label="Hutang / Potongan" val={`– ${fmt(entry.hutang)}`} c="#dc2626"/>}
                </>
              )}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:14,padding:"14px 16px",background:"#f0fdf4",borderRadius:10}}>
                <span style={{fontWeight:800,fontSize:15,color:"#0f172a"}}>TOTAL GAJI DITERIMA</span>
                <span style={{fontFamily:"monospace",fontSize:22,fontWeight:800,color:"#15803d"}}>{fmt(entry.totalGaji)}</span>
              </div>
              {entry.keterangan&&<div style={{marginTop:10,padding:"8px 12px",background:"#fffbeb",borderRadius:8,fontSize:12,color:"#92400e",borderLeft:"3px solid #f59e0b"}}><b>Ket:</b> {entry.keterangan}</div>}
              <div style={{marginTop:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{display:"inline-block",padding:"4px 12px",borderRadius:20,fontSize:11,fontWeight:700,background:entry.sudahDibayar?"#dcfce7":"#fef9c3",color:entry.sudahDibayar?"#166534":"#854d0e"}}>
                  {entry.sudahDibayar?"✓ Sudah Dibayar":"⏳ Belum Dibayar"}
                </span>
                <div style={{fontSize:10,color:"#94a3b8"}}>Dicetak: {fmtL(new Date())}</div>
              </div>
              <div style={{marginTop:20,borderTop:"1px dashed #e2e8f0",paddingTop:12,display:"flex",justifyContent:"space-between",fontSize:11,color:"#94a3b8"}}>
                <div>Tanda Tangan Pemberi Kerja<br/><span style={{display:"inline-block",marginTop:22,borderBottom:"1px solid #ccc",width:130}}/></div>
                <div style={{textAlign:"right"}}>Tanda Tangan Karyawan<br/><span style={{display:"inline-block",marginTop:22,borderBottom:"1px solid #ccc",width:130}}/></div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
function SR({label,val,c}){
  return(
    <div style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #f1f5f9"}}>
      <span style={{fontSize:13,color:"#475569"}}>{label}</span>
      <span style={{fontFamily:"monospace",fontSize:13,fontWeight:600,color:c}}>{val}</span>
    </div>
  );
}

// ─── EXPORT ───────────────────────────────────────────────────────────────────
function ExportPage({workers,locs,payroll}){
  const [loading,setLoading]=useState(false);
  const all=Object.values(payroll);
  const allWeeks=[...new Set(all.map(e=>e.weekStart).filter(Boolean))].sort().reverse();
  const locName=id=>locs.find(l=>l.id===id)?.nama||"—";

  const doExport=async()=>{
    setLoading(true);
    try{
      const wb=XLSX.utils.book_new();
      // Pekerja
      const ws1=XLSX.utils.aoa_to_sheet([
        ["ID","Nama","Jabatan","Lokasi Utama","Gaji per Hari (Rp)","Tgl Masuk","Status","Catatan"],
        ...workers.map(w=>[w.id,w.nama,w.jabatan||"",locName(w.lokasiId),w.gajiHarian,w.tanggalMasuk||"",w.aktif?"Aktif":"Non-Aktif",w.catatan||""])
      ]);
      ws1["!cols"]=[{wch:12},{wch:24},{wch:16},{wch:22},{wch:16},{wch:12},{wch:10},{wch:20}];
      XLSX.utils.book_append_sheet(wb,ws1,"Data Pekerja");
      // Lokasi
      const ws2=XLSX.utils.aoa_to_sheet([
        ["ID","Nama Lokasi","Deskripsi","Status","Pekerja Aktif"],
        ...locs.map(l=>[l.id,l.nama,l.deskripsi||"",l.aktif?"Aktif":"Selesai",workers.filter(w=>w.lokasiId===l.id&&w.aktif).length])
      ]);
      ws2["!cols"]=[{wch:12},{wch:26},{wch:24},{wch:10},{wch:14}];
      XLSX.utils.book_append_sheet(wb,ws2,"Data Lokasi");
      // Riwayat
      const ws3=XLSX.utils.aoa_to_sheet([
        ["Minggu","Nama","Jabatan","Lokasi Utama","Detail Lokasi & Hari","Total Hari","Gaji/Hari","Lembur","Bonus","Hutang","Total Gaji","Status","Keterangan"],
        ...all.sort((a,b)=>b.weekStart?.localeCompare(a.weekStart||"")||0).map(e=>{
          const w=workers.find(x=>x.id===e.workerId);
          const bd=(e.breakdown||[]).filter(b=>b.hari>0).map(b=>`${locName(b.lokasiId)}: ${b.hari}hr`).join(", ");
          return[e.weekStart?weekLabel(e.weekStart):"",w?.nama||"",w?.jabatan||"",locName(w?.lokasiId||""),bd,e.totalHari||0,e.gajiHarian||0,e.lembur||0,e.bonus||0,e.hutang||0,e.totalGaji||0,e.sudahDibayar?"Dibayar":"Belum",e.keterangan||""];
        })
      ]);
      ws3["!cols"]=[{wch:28},{wch:22},{wch:14},{wch:22},{wch:40},{wch:10},{wch:14},{wch:12},{wch:12},{wch:12},{wch:14},{wch:10},{wch:20}];
      XLSX.utils.book_append_sheet(wb,ws3,"Riwayat Penggajian");
      // Per minggu
      allWeeks.slice(0,20).forEach(wk=>{
        const es=all.filter(e=>e.weekStart===wk);if(!es.length)return;
        const rows=[
          [`LAPORAN GAJI — ${weekLabel(wk)}`],[],
          ["Nama","Jabatan","Detail Lokasi & Hari","Total Hari","Gaji/Hari","Lembur","Bonus","Hutang","Total Gaji","Status"],
          ...es.map(e=>{
            const w=workers.find(x=>x.id===e.workerId);
            const bd=(e.breakdown||[]).filter(b=>b.hari>0).map(b=>`${locName(b.lokasiId)}: ${b.hari}hr`).join(", ");
            return[w?.nama||"",w?.jabatan||"",bd,e.totalHari||0,e.gajiHarian||0,e.lembur||0,e.bonus||0,e.hutang||0,e.totalGaji||0,e.sudahDibayar?"Dibayar":"Belum"];
          }),
          [],
          ["TOTAL","","",es.reduce((a,e)=>a+(e.totalHari||0),0),"",es.reduce((a,e)=>a+(e.lembur||0),0),es.reduce((a,e)=>a+(e.bonus||0),0),es.reduce((a,e)=>a+(e.hutang||0),0),es.reduce((a,e)=>a+(e.totalGaji||0),0),""]
        ];
        const ws=XLSX.utils.aoa_to_sheet(rows);
        ws["!cols"]=[{wch:22},{wch:14},{wch:36},{wch:10},{wch:14},{wch:12},{wch:12},{wch:12},{wch:14},{wch:10}];
        XLSX.utils.book_append_sheet(wb,ws,`Mgg ${wk}`.slice(0,31));
      });
      XLSX.writeFile(wb,`PayTrack_${new Date().toISOString().slice(0,10)}.xlsx`);
    }finally{setLoading(false);}
  };

  return(
    <div className="fi">
      <PH title="Export Data" sub="Download ke Excel · Import ke Google Sheets"/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:20}}>
        {[
          {l:"Total Pekerja",       v:workers.length,              c:"#4ade80"},
          {l:"Total Lokasi",        v:locs.length,                 c:"#60a5fa"},
          {l:"Entri Penggajian",    v:all.length,                  c:"#fb923c"},
          {l:"Minggu Tercatat",     v:allWeeks.length,             c:"#fbbf24"},
          {l:"Total Semua Gaji",    v:fmt(all.reduce((a,e)=>a+(e.totalGaji||0),0)),c:"#a78bfa"},
          {l:"Belum Dibayar",       v:all.filter(e=>!e.sudahDibayar).length+" entri",c:"#f87171"},
        ].map(s=>(
          <div key={s.l} className="card" style={{borderTop:`2px solid ${s.c}`,padding:"13px 15px"}}>
            <div style={{fontSize:9,color:"#475569",fontWeight:700,textTransform:"uppercase",letterSpacing:".08em"}}>{s.l}</div>
            <div className="mono" style={{fontSize:15,fontWeight:700,color:s.c,marginTop:5}}>{s.v}</div>
          </div>
        ))}
      </div>
      <div className="card" style={{maxWidth:500}}>
        <div style={{fontSize:15,fontWeight:800,marginBottom:8}}>Export ke Excel (.xlsx)</div>
        <p style={{fontSize:12,color:"#475569",marginBottom:14}}>File berisi sheet:</p>
        <div style={{display:"grid",gap:7,marginBottom:16}}>
          {[
            {e:"📋",t:"Data Pekerja",         d:"Nama, lokasi utama, gaji harian"},
            {e:"📍",t:"Data Lokasi",           d:"Semua lokasi & proyek"},
            {e:"📊",t:"Riwayat Penggajian",    d:"Histori lengkap + detail lokasi per minggu"},
            {e:"📅",t:`Per Minggu (${Math.min(allWeeks.length,20)} minggu)`, d:"Rekap gaji per sheet per minggu"},
          ].map(r=>(
            <div key={r.t} style={{display:"flex",gap:10,padding:"8px 10px",background:"#111318",borderRadius:8,border:"1px solid #1e2535"}}>
              <span style={{fontSize:17}}>{r.e}</span>
              <div><div style={{fontWeight:700,fontSize:12}}>{r.t}</div><div style={{fontSize:10,color:"#475569",marginTop:1}}>{r.d}</div></div>
            </div>
          ))}
        </div>
        <button className="btn btn-g" style={{width:"100%",justifyContent:"center",padding:"11px",fontSize:13}} onClick={doExport} disabled={loading}>
          {loading?"⟳ Mengekspor…":"📥 Download Excel"}
        </button>
        <p style={{fontSize:10,color:"#2d3748",marginTop:8,textAlign:"center"}}>Google Sheets: File → Import → Upload .xlsx</p>
      </div>
    </div>
  );
}

// ─── SHARED ───────────────────────────────────────────────────────────────────
function PH({title,sub}){
  return(
    <div style={{marginBottom:18}}>
      <h1 style={{fontSize:21,fontWeight:800,letterSpacing:"-0.4px",color:"#f1f5f9"}}>{title}</h1>
      {sub&&<p style={{color:"#475569",fontSize:11,marginTop:3,fontWeight:600}}>{sub}</p>}
    </div>
  );
}
function FLD({label,children}){return <div><div className="label">{label}</div>{children}</div>;}
function Empty({msg}){return <div style={{textAlign:"center",padding:"26px 0",color:"#2d3748",fontSize:13,fontWeight:600}}>{msg}</div>;}
