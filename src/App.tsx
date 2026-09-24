import React, { useState, useEffect, useRef, useMemo } from "react";
import logoUrl from "container:///mnt/data/src/assets/edf85c4782e630a4-Document_Analytics_Logo";

const FONT_LINK = "https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=Inter:wght@400;500;600;700&display=swap";
const XLSX_CDN = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
const XLSX_CDN_FALLBACK = "https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js";
const FILE_ACCEPT = ".csv,.json,.xlsx,.xls,.txt";
const EXCEL_ACCEPT = ".xlsx,.xls";

type Row = Record<string, any>;
type User = { name: string; email: string; type: "Free" | "Guest" } | null;
type Prefs = {
  currency: string;
  dateFormat: string;
  numberFormat: string;
  autoDetect: boolean;
  maxRows: number;
  compact: boolean;
  livePreview: boolean;
};

type AIIssue = {
  id: string;
  type: string;
  count: number;
  description: string;
  example: string;
  fixLabel: string;
  severity: "low"|"med"|"high";
};

const DEMO_CSV = `Date,Category,Amount,Currency,Customer,Status
2024-01-05,Marketing,$1,240.00,USD,Avery Co.,Paid
2024-01-12,Engineering,$4,890.50,USD,Nova Labs,Paid
2024-02-03,Design,$890.00,USD,Pulse Inc,Pending
2024-02-19,Marketing,$2,100.00,USD,Avery Co.,Paid
2024-03-01,Sales,$3,450.75,USD,Orbit HQ,Paid
2024-03-11,Engineering,$5,200.00,USD,Nova Labs,Paid
2024-03-28,Design,$1,100.00,USD,Pulse Inc,Paid
2024-04-09,Marketing,$1,800.25,USD,Avery Co.,Pending
2024-04-22,Sales,$4,100.00,USD,Orbit HQ,Paid
2024-05-05,Engineering,$6,300.00,USD,Nova Labs,Paid`;

const RATES: Record<string, number> = { USD: 1, EUR: 0.92, INR: 83.5, GBP: 0.79 };
const SYM_TO_CODE: Record<string,string> = { "$":"USD", "€":"EUR", "£":"GBP", "₹":"INR" };

const MAX_FILE_SIZE = 20 * 1024 * 1024;
const LARGE_FILE_THRESHOLD = 5 * 1024 * 1024;
function formatFileSize(bytes: number): string {
  if(bytes < 1024) return `${bytes} B`;
  if(bytes < 1024*1024) return `${(bytes/1024).toFixed(1)} KB`;
  return `${(bytes/(1024*1024)).toFixed(1)} MB`;
}

function parseCSV(text: string): Row[] {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map(h => h.trim());
  return lines.slice(1).map(line => {
    let clean = line.replace(/\$\s*([\d,]+\.\d+)/g, (m) => m.replace(/,/g, ""));
    const parts = clean.split(",");
    const obj: Row = {};
    headers.forEach((h, i) => obj[h] = (parts[i] || "").trim());
    return obj;
  });
}

function detectType(values: any[]) {
  const sample = values.filter(v => v !== "" && v != null).slice(0, 20);
  if (!sample.length) return "text";
  if (sample.every(v => !isNaN(Date.parse(v)) && isNaN(Number(v)))) return "date";
  const money = sample.filter(v => typeof v === "string" && (v.includes("$") || v.includes("€") || v.includes("₹") || v.includes("£"))).length;
  if (money > sample.length * 0.5) return "money";
  if (sample.every(v => !isNaN(Number(String(v).replace(/[$,]/g,""))))) return "number";
  return "text";
}

function normalizeHeaderName(raw: string): string {
  if(!raw) return "";
  let s = String(raw).trim();
  s = s.replace(/[_-]+/g, " ");
  s = s.replace(/[^a-zA-Z0-9 ]/g, "");
  s = s.replace(/\s+/g, " ").trim();
  if(!s) return "";
  return s.split(" ").map(w=> w.charAt(0).toUpperCase()+w.slice(1).toLowerCase()).join(" ");
}

function toISODate(str: string): string | null {
  const t = String(str).trim();
  if(!t) return null;
  if(/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  let m = t.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if(m){
    const y=m[1], mo=m[2].padStart(2,'0'), d=m[3].padStart(2,'0');
    const iso=`${y}-${mo}-${d}`;
    if(!isNaN(new Date(iso).getTime())) return iso;
  }
  m = t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if(m){
    let a=parseInt(m[1]), b=parseInt(m[2]), y=m[3];
    if(y.length===2) y='20'+y;
    let day, month;
    if(a>12){ day=a; month=b; } else if(b>12){ day=b; month=a; } else { month=a; day=b; }
    if(month<1||month>12||day<1||day>31) return null;
    const iso=`${y}-${String(month).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
    if(!isNaN(new Date(iso).getTime())) return iso;
  }
  const d=new Date(t);
  if(!isNaN(d.getTime())) return d.toISOString().slice(0,10);
  return null;
}

function calcQuality(rows: Row[], cols: string[], issues: AIIssue[]): number {
  if(!rows.length) return 0;
  const total = rows.length * cols.length;
  let empty=0;
  rows.forEach(r=>cols.forEach(c=>{ if(String(r[c]??"").trim()==="") empty++; }));
  const emptyPenalty = (empty/total)*35;
  const issuePenalty = issues.reduce((a,i)=>a+ (i.severity==="high"?3: i.severity==="med"?1.5:0.6),0);
  const score = 100 - emptyPenalty - issuePenalty;
  return Math.max(52, Math.min(98, Math.round(score)));
}

export default function App() {
  const [view, setView] = useState<"Upload"|"Data"|"Analyze"|"Visuals"|"Currency"|"Settings">("Upload");
  const [settingsTab, setSettingsTab] = useState<"Profile"|"Appearance"|"Preferences"|"Data"|"About">("Profile");
  const [rows, setRows] = useState<Row[]>(() => {
    try { const s = localStorage.getItem("sheetsense_rows"); return s ? JSON.parse(s) : []; } catch { return []; }
  });
  const [originalRows, setOriginalRows] = useState<Row[]|null>(() => {
    try { const s = localStorage.getItem("sheetsense_original"); return s ? JSON.parse(s) : null; } catch { return null; }
  });
  const [user, setUser] = useState<User>(() => {
    try { const s = localStorage.getItem("sheetsense_user"); return s ? JSON.parse(s) : null; } catch { return null; }
  });
  const [prefs, setPrefs] = useState<Prefs>(() => {
    try { const s = localStorage.getItem("sheetsense_prefs"); return s ? JSON.parse(s) : { currency:"USD", dateFormat:"MM/DD/YYYY", numberFormat:"en-US", autoDetect:true, maxRows:200, compact:false, livePreview:true }; } catch { return { currency:"USD", dateFormat:"MM/DD/YYYY", numberFormat:"en-US", autoDetect:true, maxRows:200, compact:false, livePreview:true }; }
  });
  const [search, setSearch] = useState("");
  const [dragActive, setDragActive] = useState(false);
  const [authOpen, setAuthOpen] = useState(false);
  const [authMode, setAuthMode] = useState<"login"|"signup">("login");
  const [authForm, setAuthForm] = useState({ name:"", email:"", password:"" });
  const [toast, setToast] = useState<{msg:string; id:number} | null>(null);
  const [chartType, setChartType] = useState<"Bar"|"Line"|"Area"|"Pie">("Bar");
  const [yAxis, setYAxis] = useState<string>("Amount");
  const [targetCurrency, setTargetCurrency] = useState("EUR");
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const excelInputRef = useRef<HTMLInputElement>(null);

  // XLSX
  const [xlsxReady, setXlsxReady] = useState<boolean>(()=> typeof (window as any).XLSX !== "undefined");
  const [workbookMeta, setWorkbookMeta] = useState<{sheets:string[], fileName:string}|null>(null);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [wbObj, setWbObj] = useState<any>(null);
  const [lastFileSize, setLastFileSize] = useState<number>(0);
  const [uploadProgress, setUploadProgress] = useState<number>(0);
  const [isParsing, setIsParsing] = useState<boolean>(false);
  const [parsingFileName, setParsingFileName] = useState<string>("");
  const [pendingFile, setPendingFile] = useState<File|null>(null);
  const [xlsxLoading, setXlsxLoading] = useState<boolean>(()=> typeof (window as any).XLSX === "undefined");

  // AI
  const [aiIssues, setAiIssues] = useState<AIIssue[]>([]);
  const [aiToggles, setAiToggles] = useState<Record<string, boolean>>({});
  const [isFormatted, setIsFormatted] = useState<boolean>(()=> {
    try{ return !!localStorage.getItem("sheetsense_formatted"); } catch { return false; }
  });
  const [formatLog, setFormatLog] = useState<string[]>(()=>{
    try{ const s=localStorage.getItem("sheetsense_log"); return s?JSON.parse(s):[]; } catch{ return []; }
  });
  const [showRaw, setShowRaw] = useState(false);
  const [headerMapPreview, setHeaderMapPreview] = useState<{before:string, after:string}[]>([]);

  const columns = useMemo(() => rows.length ? Object.keys(rows[0]) : [], [rows]);
  const originalColumns = useMemo(()=> originalRows?.length ? Object.keys(originalRows[0]) : columns, [originalRows, columns]);
  const displayCols = showRaw && originalRows ? originalColumns : columns;
  const displayRows = showRaw && originalRows ? originalRows : rows;

  const numericCols = useMemo(() => {
    if (!rows.length) return [];
    return columns.filter(c => detectType(rows.map(r=>r[c])) === "number" || detectType(rows.map(r=>r[c]))==="money");
  }, [rows, columns]);

  useEffect(() => {
    localStorage.setItem("sheetsense_rows", JSON.stringify(rows));
  }, [rows]);
  useEffect(()=> {
    if(originalRows) localStorage.setItem("sheetsense_original", JSON.stringify(originalRows));
  }, [originalRows]);
  useEffect(()=>{
    if(isFormatted) localStorage.setItem("sheetsense_formatted","1");
    else localStorage.removeItem("sheetsense_formatted");
  }, [isFormatted]);
  useEffect(()=>{ localStorage.setItem("sheetsense_log", JSON.stringify(formatLog)); }, [formatLog]);
  useEffect(() => {
    if (user) localStorage.setItem("sheetsense_user", JSON.stringify(user));
    else localStorage.removeItem("sheetsense_user");
  }, [user]);
  useEffect(() => {
    localStorage.setItem("sheetsense_prefs", JSON.stringify(prefs));
  }, [prefs]);

  // Load SheetJS with fallback + pending queue (fixed for offline validation)
  useEffect(()=>{
    if((window as any).XLSX){
      setXlsxReady(true);
      setXlsxLoading(false);
      return;
    }
    const loadScript = (src: string, onLoad: ()=>void, onError: ()=>void) => {
      const existing = document.querySelector(`script[src="${src}"]`) as HTMLScriptElement | null;
      if(existing){
        if((window as any).XLSX){
          onLoad();
          return existing;
        }
        existing.addEventListener("load", onLoad);
        existing.addEventListener("error", onError);
        return existing;
      }
      const s=document.createElement("script");
      s.src=src; s.async=true; s.defer=true;
      s.onload=onLoad;
      s.onerror=onError;
      document.head.appendChild(s);
      return s;
    };
    const handleReady = () => {
      setXlsxReady(true);
      setXlsxLoading(false);
    };
    const handlePrimaryFail = () => {
      console.warn("Primary XLSX CDN failed, trying fallback");
      setToast({msg:"Primary Excel parser failed, trying backup…", id: Date.now()});
      setTimeout(()=>setToast(null), 2500);
      loadScript(XLSX_CDN_FALLBACK, handleReady, ()=>{
        console.warn("Fallback XLSX CDN also failed");
        setXlsxLoading(false);
        setToast({msg:"Failed to load Excel parser. Check network.", id: Date.now()});
      });
    };
    setXlsxLoading(true);
    loadScript(XLSX_CDN, handleReady, handlePrimaryFail);
  },[]);

  // When XLSX becomes ready, parse queued file
  useEffect(()=>{
    if(xlsxReady && pendingFile){
      const f = pendingFile;
      setPendingFile(null);
      // small delay to ensure state stable
      setTimeout(()=> handleFile(f), 50);
    }
  }, [xlsxReady]);

  const showToast = (msg: string) => {
    setToast({ msg, id: Date.now() });
    setTimeout(()=>setToast(null), 3200);
  };

  const analyzeRows = (rws: Row[]) => {
    if(!rws.length){ setAiIssues([]); setAiToggles({}); setHeaderMapPreview([]); return []; }
    const cols = Object.keys(rws[0]||{});
    let emptyRowsCount = 0;
    let whitespaceCount = 0;
    let dateCount = 0;
    let numberSymCount = 0;
    let currencyCombinedCount = 0;
    let boolMixedCount = 0;
    let missingNumCount = 0;
    let headerSpecial = 0;
    let duplicateHeaders = 0;
    let emptyColsList: string[] = [];

    // empty rows
    rws.forEach(r=>{
      if(Object.values(r).every(v=> String(v??"").trim()==="")) emptyRowsCount++;
    });
    // empty cols
    cols.forEach(c=>{
      if(rws.every(row=> String(row[c]??"").trim()==="")) emptyColsList.push(c);
    });
    // header special & duplicate
    const seen = new Map<string,string[]>();
    cols.forEach(c=>{
      if(/[_-]/.test(c) || /[^a-zA-Z0-9 ]/.test(c)) headerSpecial++;
      const low=c.toLowerCase().trim();
      if(!seen.has(low)) seen.set(low,[c]); else seen.get(low)!.push(c);
    });
    seen.forEach(v=>{ if(v.length>1) duplicateHeaders+= v.length-1; });

    // cell level
    let dateExample = "";
    let numberExample = "";
    let currencyExample = "";
    rws.forEach(r=>{
      cols.forEach(c=>{
        const raw = r[c];
        const s = String(raw??"");
        if(s && s!==s.trim()) whitespaceCount++;
        if(s.trim()){
          const iso = toISODate(s);
          if(iso && !/^\d{4}-\d{2}-\d{2}$/.test(s.trim())){
            dateCount++;
            if(!dateExample) dateExample=`${s.trim()} → ${iso}`;
          }
          if(/[$€£₹,]/.test(s) && !isNaN(Number(s.replace(/[$€£₹,\s]/g,"")))){
            numberSymCount++;
            if(!numberExample) numberExample=s;
          }
          if(/^\s*[$€£₹]\s*[\d,]+/.test(s) || /\d\s*(USD|EUR|INR|GBP)\s*$/i.test(s) || /^\s*[\d,]+\s*[$€£₹]\s*$/.test(s)){
            currencyCombinedCount++;
            if(!currencyExample) currencyExample=s;
          }
        } else {
          // blank check for numeric cols later
        }
      });
    });

    // bool mixed + missing numbers
    cols.forEach(c=>{
      const vals = rws.map(r=>String(r[c]??"").trim()).filter(Boolean);
      const lower = vals.map(v=>v.toLowerCase());
      const boolSet = new Set(["yes","no","y","n","true","false","1","0"]);
      const boolVals = lower.filter(v=>boolSet.has(v));
      if(boolVals.length>=2 && boolVals.length>= vals.length*0.5){
        // mixed bool
        const hasYes = boolVals.some(v=>["yes","y","true","1"].includes(v));
        const hasNo = boolVals.some(v=>["no","n","false","0"].includes(v));
        if(hasYes && hasNo) boolMixedCount++;
      }
      // numeric missing
      if(detectType(rws.map(r=>r[c]))==="number" || detectType(rws.map(r=>r[c]))==="money"){
        const blank = rws.filter(r=> String(r[c]??"").trim()==="").length;
        if(blank>0) missingNumCount+=blank;
      }
    });

    const issues: AIIssue[] = [];
    if(emptyRowsCount>0) issues.push({id:"empty_rows", type:"empty_rows", count:emptyRowsCount, description:`${emptyRowsCount} empty rows found`, example:"Blank rows will be removed", fixLabel:`Remove ${emptyRowsCount} empty rows`, severity:"low"});
    if(emptyColsList.length>0) issues.push({id:"empty_cols", type:"empty_cols", count:emptyColsList.length, description:`${emptyColsList.length} empty columns: ${emptyColsList.slice(0,3).join(", ")}`, example:"Columns with no data", fixLabel:`Remove ${emptyColsList.length} empty columns`, severity:"low"});
    if(whitespaceCount>0) issues.push({id:"whitespace", type:"whitespace", count:whitespaceCount, description:`${whitespaceCount} cells with leading/trailing spaces`, example:"'  Avery Co. ' → 'Avery Co.'", fixLabel:"Trim whitespace", severity:"med"});
    if(headerSpecial>0) issues.push({id:"header_norm", type:"header_norm", count:headerSpecial, description:`${headerSpecial} headers need normalization`, example:"customer_name → Customer Name", fixLabel:"Normalize headers (Title Case)", severity:"med"});
    if(duplicateHeaders>0) issues.push({id:"dup_headers", type:"dup_headers", count:duplicateHeaders, description:`${duplicateHeaders} duplicate header${duplicateHeaders>1?'s':''} detected`, example:"Name, Name → Name, Name 2", fixLabel:"Auto-rename duplicates", severity:"high"});
    if(dateCount>0) issues.push({id:"date_format", type:"date_format", count:dateCount, description:`${dateCount} dates in mixed formats`, example:dateExample||"01/02/2024 → 2024-01-02", fixLabel:"Convert dates to YYYY-MM-DD ISO", severity:"high"});
    if(numberSymCount>0) issues.push({id:"number_fmt", type:"number_fmt", count:numberSymCount, description:`${numberSymCount} numbers with symbols/commas`, example:numberExample?`${numberExample} → ${Number(numberExample.replace(/[^0-9.-]/g,""))}`:"$1,240.00 → 1240", fixLabel:"Strip currency symbols & commas", severity:"med"});
    if(currencyCombinedCount>0) issues.push({id:"currency_split", type:"currency_split", count:currencyCombinedCount, description:`${currencyCombinedCount} combined amount+currency values`, example:currencyExample?`${currencyExample} → Amount + Currency`:"$2400 → 2400 + USD", fixLabel:"Split Amount & Currency columns", severity:"high"});
    if(boolMixedCount>0) issues.push({id:"bool_norm", type:"bool_norm", count:boolMixedCount, description:`${boolMixedCount} columns with mixed booleans`, example:"yes / No / 1 / 0 → Yes / No", fixLabel:"Normalize yes/no true/false", severity:"low"});
    if(missingNumCount>0) issues.push({id:"smart_fill", type:"smart_fill", count:missingNumCount, description:`${missingNumCount} blank numeric cells`, example:"Fill with 0 or average", fixLabel:`Smart fill ${missingNumCount} blanks with 0`, severity:"low"});

    // header map preview
    const mapPreview: {before:string, after:string}[] = [];
    const seenNorm = new Map<string,number>();
    cols.forEach(c=>{
      let norm = normalizeHeaderName(c) || c.trim();
      if(!norm) norm=`Column`;
      const key = norm.toLowerCase();
      const cnt = seenNorm.get(key)||0;
      seenNorm.set(key,cnt+1);
      let finalName = cnt===0? norm : `${norm} ${cnt+1}`;
      if(finalName!==c) mapPreview.push({before:c, after:finalName});
    });
    setHeaderMapPreview(mapPreview.slice(0,8));

    const toggles: Record<string,boolean> = {};
    issues.forEach(i=> toggles[i.id]=true);
    setAiToggles(toggles);
    setAiIssues(issues);
    return issues;
  };

  useEffect(()=>{
    if(rows.length && !isFormatted){
      analyzeRows(rows);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rows.length]);

  const filteredRows = useMemo(() => {
    const base = displayRows;
    if (!search) return base.slice(0, prefs.maxRows);
    const q = search.toLowerCase();
    return base.filter(r => Object.values(r).join(" ").toLowerCase().includes(q)).slice(0, prefs.maxRows);
  }, [displayRows, search, prefs.maxRows]);

  const qualityScore = useMemo(()=> calcQuality(rows, columns, aiIssues), [rows, columns, aiIssues]);

  const stats = useMemo(() => {
    if (!rows.length) return null;
    const sumCol = numericCols[0];
    let sum = 0, count = 0;
    if (sumCol) {
      rows.forEach(r => {
        const v = Number(String(r[sumCol]).replace(/[^0-9.-]/g,""));
        if (!isNaN(v)) { sum += v; count++; }
      });
    }
    return {
      total: rows.length,
      numericCols: numericCols.length,
      sum,
      avg: count ? sum / count : 0,
    };
  }, [rows, numericCols]);

  // Canvas chart drawing
  useEffect(() => {
    if (view !== "Visuals" || !canvasRef.current || !rows.length) return;
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    const dpr = window.devicePixelRatio || 1;
    const rect = canvas.getBoundingClientRect();
    canvas.width = rect.width * dpr;
    canvas.height = rect.height * dpr;
    ctx.scale(dpr, dpr);
    const W = rect.width, H = rect.height;
    ctx.clearRect(0,0,W,H);
    ctx.strokeStyle = "rgba(17,17,19,0.06)";
    ctx.lineWidth = 1;
    for (let i=0;i<5;i++) {
      const y = 24 + (H-72)*(i/4);
      ctx.beginPath(); ctx.moveTo(48,y); ctx.lineTo(W-16,y); ctx.stroke();
    }
    const vals = rows.slice(0,12).map(r => {
      const raw = r[yAxis] ?? Object.values(r)[1];
      const n = Number(String(raw).replace(/[^0-9.-]/g,""));
      return isNaN(n) ? 0 : n;
    });
    const max = Math.max(...vals, 1);
    const barW = (W-72) / vals.length * 0.62;
    if (chartType === "Bar") {
      vals.forEach((v,i)=>{
        const x = 56 + i * ((W-72)/vals.length) + ((W-72)/vals.length - barW)/2;
        const h = (v/max)*(H-96);
        const y = H-40 - h;
        const grad = ctx.createLinearGradient(0,y,0,y+h);
        grad.addColorStop(0,"#6C5CE7");
        grad.addColorStop(1,"#FF9A62");
        ctx.fillStyle = grad;
        ctx.beginPath();
        const r = 10;
        ctx.moveTo(x+r,y); ctx.lineTo(x+barW-r,y); ctx.quadraticCurveTo(x+barW,y,x+barW,y+r);
        ctx.lineTo(x+barW,y+h-r); ctx.quadraticCurveTo(x+barW,y+h,x+barW-r,y+h);
        ctx.lineTo(x+r,y+h); ctx.quadraticCurveTo(x,y+h,x,y+h-r);
        ctx.lineTo(x,y+r); ctx.quadraticCurveTo(x,y,x+r,y); ctx.closePath();
        ctx.fill();
      });
    } else if (chartType === "Line" || chartType === "Area") {
      ctx.beginPath();
      vals.forEach((v,i)=>{
        const x = 56 + i * ((W-72)/(vals.length-1));
        const y = H-40 - (v/max)*(H-96);
        if (i===0) ctx.moveTo(x,y); else ctx.lineTo(x,y);
      });
      ctx.strokeStyle = "#6C5CE7";
      ctx.lineWidth = 2.8;
      ctx.lineJoin = "round";
      ctx.stroke();
      vals.forEach((v,i)=>{
        const x = 56 + i * ((W-72)/(vals.length-1));
        const y = H-40 - (v/max)*(H-96);
        ctx.beginPath(); ctx.arc(x,y,4,0,Math.PI*2); ctx.fillStyle = "#fff"; ctx.fill(); ctx.strokeStyle="#6C5CE7"; ctx.lineWidth=2; ctx.stroke();
      });
      if (chartType==="Area") {
        ctx.lineTo(56 + (vals.length-1)*((W-72)/(vals.length-1)), H-40);
        ctx.lineTo(56, H-40);
        ctx.closePath();
        const grad = ctx.createLinearGradient(0,24,0,H-40);
        grad.addColorStop(0,"rgba(108,92,231,0.28)"); grad.addColorStop(1,"rgba(255,154,98,0.02)");
        ctx.fillStyle = grad; ctx.fill();
      }
    } else if (chartType==="Pie") {
      const total = vals.reduce((a,b)=>a+b,0) || 1;
      let ang = -Math.PI/2;
      const cx = W/2, cy = H/2+8, rad = Math.min(W,H)*0.32;
      const colors = ["#6C5CE7","#FF9A62","#111113","#E9E5FF","#FFE1D0"];
      vals.forEach((v,i)=>{
        const slice = (v/total)*Math.PI*2;
        ctx.beginPath(); ctx.moveTo(cx,cy); ctx.arc(cx,cy,rad,ang,ang+slice); ctx.closePath();
        ctx.fillStyle = colors[i%colors.length]; ctx.fill();
        ang+=slice;
      });
      ctx.beginPath(); ctx.arc(cx,cy,rad*0.52,0,Math.PI*2); ctx.fillStyle="#FCFAF3"; ctx.fill();
    }
  }, [view, rows, chartType, yAxis]);

  const parseSheetFromWB = (wb: any, sheetName: string, explicitFileSize?: number) => {
    try{
      const sheet = wb.Sheets[sheetName];
      if(!sheet){ showToast("Sheet not found"); setIsParsing(false); return; }
      console.log(`[SheetSense] Parsing sheet "${sheetName}"`);
      let json = (window as any).XLSX.utils.sheet_to_json(sheet, { defval:"", raw:false, blankrows:false });
      if(!json.length){
        const arr = (window as any).XLSX.utils.sheet_to_json(sheet, { header:1, defval:"", blankrows:false }) as any[][];
        if(!arr.length){ showToast("Empty sheet"); setIsParsing(false); return; }
        let headerRowIdx = 0;
        let maxText = -1;
        arr.slice(0,5).forEach((row,i)=>{
          const textCount = row.filter(v=> typeof v==="string" && isNaN(Number(v)) && String(v).trim()!=="").length;
          if(textCount>maxText){ maxText=textCount; headerRowIdx=i; }
        });
        const headers = arr[headerRowIdx].map((h:any, idx:number)=> String(h||`Column_${idx+1}`).trim());
        json = arr.slice(headerRowIdx+1).map((row:any[])=>{
          const obj: Row = {};
          headers.forEach((h:string, j:number)=> obj[h]= row[j]??"");
          return obj;
        }).filter((r:Row)=> Object.values(r).some(v=> String(v??"").trim()!==""));
      }
      const cloned = JSON.parse(JSON.stringify(json));
      setOriginalRows(cloned);
      setRows(cloned);
      setIsFormatted(false);
      setFormatLog([]);
      setShowRaw(false);
      setSelectedSheet(sheetName);
      setTimeout(()=> analyzeRows(cloned), 100);
      const sizeStr = formatFileSize(explicitFileSize ?? lastFileSize);
      showToast(`XLSX parsed: ${cloned.length} rows from "${sheetName}" • ${sizeStr}`);
      setIsParsing(false);
      setUploadProgress(100);
      // auto-navigate to Data view for immediate feedback
      setView("Data");
      console.log(`[SheetSense] Parsed ${cloned.length} rows from ${sheetName}`);
    }catch(e){ console.error("[SheetSense] parseSheetFromWB failed", e); showToast("Failed to parse sheet"); setIsParsing(false); }
  };

  const handleFile = (file: File) => {
    console.log(`[SheetSense] handleFile: ${file.name} ${formatFileSize(file.size)} type=${file.type}`);
    if(file.size > MAX_FILE_SIZE){
      const mb = (file.size/(1024*1024)).toFixed(1);
      showToast(`File too large: ${mb} MB. Max 20MB for Excel files`);
      return;
    }
    setLastFileSize(file.size);
    setParsingFileName(file.name);
    setUploadProgress(15); // start immediately, don't wait for onprogress
    setIsParsing(true);
    const lowerName = file.name.toLowerCase();
    const isExcelByExt = lowerName.endsWith(".xlsx") || lowerName.endsWith(".xls");
    const isExcelByMime = file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" || file.type === "application/vnd.ms-excel" || (file.type === "application/octet-stream" && isExcelByExt);
    const isExcel = isExcelByExt || isExcelByMime;
    const isJson = lowerName.endsWith(".json");
    const isCsv = lowerName.endsWith(".csv") || lowerName.endsWith(".txt");

    if(isExcel){
      if(!(window as any).XLSX || !xlsxReady){
        console.log("[SheetSense] XLSX not ready, queuing file");
        setPendingFile(file);
        showToast("Loading Excel parser... please wait 2s");
        setXlsxLoading(true);
        setUploadProgress(40);
        if(!(window as any).XLSX){
          const exists = document.querySelector(`script[src="${XLSX_CDN}"]`) || document.querySelector(`script[src="${XLSX_CDN_FALLBACK}"]`);
          if(!exists){
            const s=document.createElement("script");
            s.src=XLSX_CDN; s.async=true; s.defer=true;
            s.onload=()=>{ console.log("XLSX loaded"); setXlsxReady(true); setXlsxLoading(false); setUploadProgress(60); };
            s.onerror=()=>{
              console.log("XLSX primary failed, trying fallback");
              const fb=document.createElement("script");
              fb.src=XLSX_CDN_FALLBACK; fb.async=true; fb.defer=true;
              fb.onload=()=>{ console.log("XLSX fallback loaded"); setXlsxReady(true); setXlsxLoading(false); setUploadProgress(60); };
              fb.onerror=()=>{ showToast("Failed to load Excel parser. Check internet."); setIsParsing(false); setXlsxLoading(false); };
              document.head.appendChild(fb);
            };
            document.head.appendChild(s);
          }
        }
        return;
      }
      const reader = new FileReader();
      // Don't rely on onprogress - set progress manually
      reader.onload = (e) => {
        setUploadProgress(75);
        setTimeout(()=>{
          try{
            const data = e.target?.result as ArrayBuffer;
            const wb = (window as any).XLSX.read(data, { type:"array", cellDates:true, raw:false, cellNF:false });
            setWbObj(wb);
            const sheets = wb.SheetNames as string[];
            setWorkbookMeta({ sheets, fileName:file.name });
            const first = sheets[0];
            setSelectedSheet(first);
            parseSheetFromWB(wb, first, file.size);
          }catch(err){ 
            console.error(err); 
            const msg = (err as any)?.message || "";
            if(msg.toLowerCase().includes("memory") || file.size > 15*1024*1024){
              showToast("File too large to parse. Try splitting <10MB.");
            } else {
              showToast("Failed to parse XLSX — try CSV export");
            }
            setIsParsing(false);
            setUploadProgress(0);
          }
        }, 100);
      };
      reader.onerror = () => { showToast("Failed to read file"); setIsParsing(false); setUploadProgress(0); };
      reader.readAsArrayBuffer(file);
      return;
    }
    // CSV handling via SheetJS if available
    if(isCsv && (window as any).XLSX && xlsxReady){
      const reader = new FileReader();
      reader.onprogress = (ev:any) => { if(ev.lengthComputable) setUploadProgress(Math.round((ev.loaded/ev.total)*100)); };
      reader.onload = () => {
        try{
          const text = reader.result as string;
          const wb = (window as any).XLSX.read(text, { type:"string", cellDates:true, raw:false });
          const first = wb.SheetNames[0];
          const sheet = wb.Sheets[first];
          const json = (window as any).XLSX.utils.sheet_to_json(sheet, { defval:"", raw:false });
          if(json.length){
            const cloned = JSON.parse(JSON.stringify(json));
            setOriginalRows(cloned); setRows(cloned); setIsFormatted(false); setFormatLog([]); setShowRaw(false);
            setWorkbookMeta({ sheets: wb.SheetNames, fileName:file.name });
            setWbObj(wb); setSelectedSheet(first);
            setTimeout(()=> analyzeRows(cloned),100);
            showToast(`CSV parsed via SheetJS: ${cloned.length} rows • ${formatFileSize(file.size)}`);
            setIsParsing(false); setUploadProgress(100);
            setView("Data");
            return;
          }
        }catch{}
        const text = reader.result as string;
        const parsed = parseCSV(text);
        const cloned = JSON.parse(JSON.stringify(parsed));
        setOriginalRows(cloned); setRows(parsed); setIsFormatted(false); setFormatLog([]); setShowRaw(false);
        setWorkbookMeta(null); setWbObj(null);
        setTimeout(()=> analyzeRows(cloned),100);
        showToast(`Loaded ${file.name} • ${formatFileSize(file.size)}`);
        setIsParsing(false); setUploadProgress(100);
        setView("Data");
      };
      reader.onerror = () => { showToast("Failed to read file"); setIsParsing(false); };
      reader.readAsText(file);
      return;
    }
    // json / csv fallback
    const reader = new FileReader();
    reader.onprogress = (ev:any) => { if(ev.lengthComputable) setUploadProgress(Math.round((ev.loaded/ev.total)*100)); };
    reader.onload = () => {
      const text = reader.result as string;
      try {
        if (isJson) {
          const data = JSON.parse(text);
          const arr = Array.isArray(data) ? data : [data];
          const cloned = JSON.parse(JSON.stringify(arr));
          setOriginalRows(cloned); setRows(arr); setIsFormatted(false); setFormatLog([]); setShowRaw(false);
          setWorkbookMeta(null); setWbObj(null);
          setTimeout(()=> analyzeRows(cloned),100);
        } else {
          const parsed = parseCSV(text);
          const cloned = JSON.parse(JSON.stringify(parsed));
          setOriginalRows(cloned); setRows(parsed); setIsFormatted(false); setFormatLog([]); setShowRaw(false);
          setWorkbookMeta(null); setWbObj(null);
          setTimeout(()=> analyzeRows(cloned),100);
        }
        showToast(`Loaded ${file.name} • ${formatFileSize(file.size)}`);
        setIsParsing(false); setUploadProgress(100);
        setView("Data");
      } catch { showToast("Failed to parse file"); setIsParsing(false); }
    };
    reader.onerror = () => { showToast("Failed to read file"); setIsParsing(false); };
    reader.readAsText(file);
  };

  const applyAiFixes = () => {
    if(!originalRows && !rows.length){ showToast("No data"); return; }
    const baseOriginal = originalRows ? JSON.parse(JSON.stringify(originalRows)) : JSON.parse(JSON.stringify(rows));
    let working = JSON.parse(JSON.stringify(baseOriginal)) as Row[];
    const logs: string[] = [];
    let currentCols = working.length? Object.keys(working[0]):[];

    // helper to ensure toggles default true if undefined
    const isOn = (id:string)=> aiToggles[id]!==false;

    // 1. Trim whitespace
    if(isOn("whitespace")){
      let c=0;
      working.forEach(r=>{
        Object.keys(r).forEach(k=>{
          const v=r[k];
          if(typeof v==="string" && v!==v.trim()){ r[k]=v.trim(); c++; }
        });
      });
      if(c>0) logs.push(`Trimmed ${c} cells with extra spaces`);
    }

    // 2. Header normalization + duplicate handling
    if(isOn("header_norm") || isOn("dup_headers")){
      const oldCols = currentCols;
      const normMap: Record<string,string> = {};
      const seen = new Map<string,number>();
      oldCols.forEach(orig=>{
        let norm = normalizeHeaderName(orig);
        if(!norm) norm=orig.trim()||"Column";
        const low = norm.toLowerCase();
        const cnt = seen.get(low)||0;
        seen.set(low,cnt+1);
        const finalName = cnt===0? norm : `${norm} ${cnt+1}`;
        normMap[orig]=finalName;
      });
      // dedupe if duplicates still after norm? already handled
      // Apply map
      working = working.map(r=>{
        const nr: Row = {};
        Object.keys(r).forEach(oldK=>{
          const newK = normMap[oldK]||oldK;
          nr[newK]=r[oldK];
        });
        return nr;
      });
      currentCols = working.length? Object.keys(working[0]):[];
      const changed = Object.entries(normMap).filter(([o,n])=>o!==n);
      if(changed.length) logs.push(`Normalized headers: ${changed.map(([o,n])=>`${o} → ${n}`).slice(0,4).join(", ")}${changed.length>4?"…":""}`);
    }

    // 3. Empty rows removal
    if(isOn("empty_rows")){
      const before=working.length;
      working = working.filter(r=> !Object.values(r).every(v=> String(v??"").trim()===""));
      const removed = before-working.length;
      if(removed>0) logs.push(`Removed ${removed} empty rows`);
    }

    // 4. Empty cols removal
    if(isOn("empty_cols")){
      if(working.length){
        const colsToCheck = Object.keys(working[0]);
        const emptyCols = colsToCheck.filter(c=> working.every(row=> String(row[c]??"").trim()===""));
        if(emptyCols.length){
          working = working.map(r=>{
            const nr={...r};
            emptyCols.forEach(c=> delete nr[c]);
            return nr;
          });
          logs.push(`Removed ${emptyCols.length} empty columns: ${emptyCols.join(", ")}`);
        }
      }
    }

    // 5. Date formatting
    if(isOn("date_format")){
      let conv=0;
      working.forEach(r=>{
        Object.keys(r).forEach(k=>{
          const v=r[k];
          if(typeof v==="string" && v.trim()){
            const iso = toISODate(v);
            if(iso && iso!==v.trim()){
              // only if original looked like date and iso valid and original not iso
              if(!/^\d{4}-\d{2}-\d{2}$/.test(v.trim())){
                r[k]=iso;
                conv++;
              }
            }
          }
        });
      });
      if(conv>0) logs.push(`Converted ${conv} dates to YYYY-MM-DD ISO`);
    }

    // 6. Number formatting + currency split
    const currencyColExists = () => {
      if(!working.length) return false;
      const cols = Object.keys(working[0]);
      return cols.some(c=> /currency/i.test(c));
    };
    let hasCurrencyCol = currencyColExists();
    let numCleaned=0;
    let currSplit=0;
    working.forEach(r=>{
      Object.keys(r).forEach(k=>{
        const raw = String(r[k]??"");
        if(!raw.trim()) return;
        // currency combined detection
        if(isOn("currency_split")){
          // patterns
          let amount: number | null = null;
          let currCode: string | null = null;
          const s=raw.trim();
          let m = s.match(/^\s*([$€£₹])\s*([\d,]+\.?\d*)\s*$/);
          if(m){
            amount = Number(m[2].replace(/,/g,""));
            currCode = SYM_TO_CODE[m[1]]||"USD";
          } else {
            m = s.match(/^\s*([\d,]+\.?\d*)\s*([$€£₹])\s*$/);
            if(m){ amount=Number(m[1].replace(/,/g,"")); currCode=SYM_TO_CODE[m[2]]||"USD"; }
            else {
              m = s.match(/^\s*([\d,]+\.?\d*)\s*(USD|EUR|INR|GBP)\s*$/i);
              if(m){ amount=Number(m[1].replace(/,/g,"")); currCode=m[2].toUpperCase(); }
              else {
                m = s.match(/^\s*(USD|EUR|INR|GBP)\s*([\d,]+\.?\d*)\s*$/i);
                if(m){ currCode=m[1].toUpperCase(); amount=Number(m[2].replace(/,/g,"")); }
              }
            }
          }
          if(amount!==null && currCode && !isNaN(amount)){
            // if column looks like amount/price
            if(/amount|price|total|revenue|cost/i.test(k) || isOn("currency_split")){
              r[k]=amount;
              if(!hasCurrencyCol){
                // create Currency column if not exists
                r["Currency"]=currCode;
                hasCurrencyCol=true;
              } else {
                // set existing currency column if empty
                const curCol = Object.keys(r).find(c=> /currency/i.test(c));
                if(curCol){ if(!String(r[curCol]??"").trim()) r[curCol]=currCode; }
                else r["Currency"]=currCode;
              }
              currSplit++; return;
            }
          }
        }
        // number cleaning
        if(isOn("number_fmt")){
          if(/[$€£₹,]/.test(raw)){
            const cleaned = raw.replace(/[$€£₹,\s]/g,"");
            if(cleaned && !isNaN(Number(cleaned))){
              const n=Number(cleaned);
              if(String(r[k])!==String(n)){ r[k]=n; numCleaned++; }
            }
          }
        }
      });
    });
    if(numCleaned>0) logs.push(`Cleaned ${numCleaned} numeric cells (stripped symbols/commas)`);
    if(currSplit>0) logs.push(`Split ${currSplit} amount+currency values into Amount & Currency`);

    // 7. Boolean normalization
    if(isOn("bool_norm")){
      const cols = working.length? Object.keys(working[0]):[];
      cols.forEach(c=>{
        const sample = working.map(r=>String(r[c]??"").trim().toLowerCase()).filter(Boolean);
        const boolSet = new Set(["yes","no","y","n","true","false","1","0"]);
        const boolVals = sample.filter(v=> boolSet.has(v));
        if(boolVals.length>=2 && boolVals.length>=sample.length*0.5){
          working.forEach(r=>{
            const v=String(r[c]??"").trim().toLowerCase();
            if(["yes","y","true","1"].includes(v)){ r[c]="Yes"; }
            else if(["no","n","false","0"].includes(v)){ r[c]="No"; }
          });
          logs.push(`Normalized booleans in "${c}" to Yes/No`);
        }
      });
    }

    // 8. Smart fill
    if(isOn("smart_fill")){
      const cols = working.length? Object.keys(working[0]):[];
      cols.forEach(c=>{
        const isNum = working.some(r=> !isNaN(Number(String(r[c]??"").replace(/[^0-9.-]/g,""))) && String(r[c]??"").trim()!=="");
        if(isNum){
          let blanks=0;
          working.forEach(r=>{
            if(String(r[c]??"").trim()===""){ r[c]=0; blanks++; }
          });
          if(blanks>0) logs.push(`Filled ${blanks} blanks in "${c}" with 0`);
        }
      });
    }

    setRows(working);
    setIsFormatted(true);
    setFormatLog(logs);
    setShowRaw(false);
    // re-analyze after
    setTimeout(()=> analyzeRows(working), 200);
    showToast(`AI fixed ${logs.length} issues • ${working.length} rows ready`);
  };

  const undoFormat = () => {
    if(originalRows){ setRows(JSON.parse(JSON.stringify(originalRows))); setIsFormatted(false); setFormatLog([]); setShowRaw(false); setTimeout(()=> analyzeRows(originalRows),100); showToast("Reverted to original"); }
  };

  const moneyCols = useMemo(() => columns.filter(c => /amount|price|total|revenue|cost/i.test(c) || detectType(rows.map(r=>r[c]))==="money"), [columns, rows]);

  const aiConfidence = useMemo(()=>{
    if(!aiIssues.length) return 96;
    const base = 96;
    const penalty = aiIssues.reduce((a,i)=> a+ (i.count>10?3: i.count>3?2:1),0);
    return Math.max(71, base - penalty);
  }, [aiIssues]);

  return (
    <>
      <link rel="stylesheet" href={FONT_LINK} />
      <link rel="icon" href={logoUrl} />
      {/* SheetJS CDN required for XLSX - per spec: https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js with defer */}
      <script defer src={XLSX_CDN}></script>
      <style>{`
        :root{
          --violet:#6C5CE7;
          --peach:#FF9A62;
          --base:#FCFAF3;
          --white:#FFFFFF;
          --ink:#111113;
          --line:rgba(17,17,19,0.08);
          --shadow:0 12px 40px rgba(17,17,19,0.08), 0 2px 10px rgba(17,17,19,0.06);
          --r:24px;
        }
        *{font-family:Inter,system-ui,sans-serif; box-sizing:border-box}
        h1,h2,h3,.serif{font-family:"Instrument Serif", Georgia, serif; letter-spacing:-0.02em}
        body{margin:0; background:var(--base); color:var(--ink)}
        .glass{backdrop-filter:blur(16px) saturate(1.2); -webkit-backdrop-filter:blur(16px) saturate(1.2)}
        .card{background:var(--white); border:1px solid var(--line); border-radius:var(--r); box-shadow:var(--shadow)}
        ::-webkit-scrollbar{width:6px; height:6px} ::-webkit-scrollbar-thumb{background:rgba(17,17,19,0.12); border-radius:999px}
        @keyframes in{ from{transform:translateY(8px); opacity:0} to{transform:translateY(0); opacity:1} }
      `}</style>

      <div className="min-h-screen bg-[#FCFAF3] text-[#111113] selection:bg-[#6C5CE7]/20">
        {/* Header */}
        <header className="sticky top-0 z-40 h-[64px] border-b border-black/[0.06] bg-[#FCFAF3]/80 glass flex items-center">
          <div className="w-full max-w-[1440px] mx-auto px-4 md:px-6 flex items-center justify-between">
            <div className="flex items-center gap-3">
              <div className="w-9 h-9 rounded-[10px] bg-white border border-black/10 shadow-[0_2px_8px_rgba(0,0,0,0.06)] overflow-hidden grid place-items-center">
                <img src={logoUrl} alt="SheetSense" style={{width:"38px", height:"38px", objectFit:"contain", borderRadius:"10px"}} />
              </div>
              <div className="flex items-baseline gap-2">
                <span className="font-bold tracking-tight text-[16px]" style={{fontFamily:"Inter"}}>SheetSense</span>
                <span className="hidden md:inline text-[11px] px-2 py-0.5 rounded-full bg-black/5 border border-black/10">v2.1.0</span>
                <span className="hidden md:inline text-[10px] px-2 py-0.5 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/20 text-[#6C5CE7] font-medium">{xlsxReady ? "XLSX ready" : "Loading XLSX…"}</span>
              </div>
              {user?.type==="Guest" && <span className="ml-2 text-[11px] px-2.5 py-1 rounded-full bg-[#FF9A62]/15 border border-[#FF9A62]/30 text-[#111113] font-medium">Guest mode</span>}
            </div>
            <div className="flex items-center gap-2">
              {!user ? (
                <>
                  <button onClick={()=>{setAuthMode("login"); setAuthOpen(true);}} className="h-9 px-4 rounded-full text-[13px] font-medium border border-black/10 hover:bg-black/[0.04] transition">Log in</button>
                  <button onClick={()=>{setAuthMode("signup"); setAuthOpen(true);}} className="h-9 px-4 rounded-full text-[13px] font-semibold bg-[#111113] text-white hover:bg-black transition">Start free</button>
                </>
              ) : (
                <div className="relative group">
                  <button className="flex items-center gap-2.5 h-9 pl-1 pr-3 rounded-full border border-black/10 bg-white">
                    <div className="w-7 h-7 rounded-full bg-[#6C5CE7] text-white grid place-items-center text-[12px] font-semibold">{user.name.charAt(0).toUpperCase()}</div>
                    <span className="text-[13px] font-medium hidden md:inline">{user.name}</span>
                  </button>
                  <div className="absolute right-0 top-[44px] w-[200px] card p-1.5 hidden group-hover:block">
                    <div className="px-3 py-2">
                      <div className="text-[13px] font-medium truncate">{user.name}</div>
                      <div className="text-[11px] opacity-60 truncate">{user.email}</div>
                    </div>
                    <button onClick={()=>{setView("Settings"); setSettingsTab("Profile");}} className="w-full text-left px-3 py-2 rounded-[12px] text-[13px] hover:bg-black/5">Settings</button>
                    <button onClick={()=>{setUser(null); showToast("Logged out");}} className="w-full text-left px-3 py-2 rounded-[12px] text-[13px] hover:bg-black/5">Log out</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </header>

        <div className="max-w-[1440px] mx-auto flex">
          {/* Sidebar desktop */}
          <aside className="hidden md:flex w-[280px] shrink-0 sticky top-[64px] h-[calc(100vh-64px)] flex-col p-4 gap-4">
            {/* Brand lockup */}
            <div className="card p-3.5 flex items-center gap-3">
              <div className="w-[38px] h-[38px] rounded-[10px] bg-white border border-black/10 shadow-[0_2px_10px_rgba(0,0,0,0.07)] overflow-hidden shrink-0 grid place-items-center">
                <img src={logoUrl} alt="SheetSense logo" style={{width:"38px", height:"38px", objectFit:"contain", borderRadius:"10px"}} />
              </div>
              <div className="leading-[1.1]">
                <div className="font-bold text-[15px] tracking-tight" style={{fontFamily:"Inter"}}>SheetSense</div>
                <div className="text-[11px] opacity-60 mt-0.5 font-medium">sheets → insights</div>
              </div>
              <div className="ml-auto w-2 h-2 rounded-full bg-[#6C5CE7] animate-pulse" />
            </div>

            <nav className="card p-2">
              {[
                {k:"Upload", i:"⤓", d:"Import sheet"},
                {k:"Data", i:"▤", d:`${rows.length} rows`},
                {k:"Analyze", i:"◑", d:"Insights"},
                {k:"Visuals", i:"◿", d:"Charts"},
                {k:"Currency", i:"¤", d:"Convert"},
                {k:"Settings", i:"⚙", d:"Workspace"},
              ].map(item=>(
                <button key={item.k} onClick={()=>{ if(view===item.k){ showToast(`${item.k} — already here`);} setView(item.k as any); }} className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-[14px] text-left transition ${view===item.k ? "bg-[#111113] text-white" : "hover:bg-black/[0.04] text-[#111113]/80"}`}>
                  <span className={`w-8 h-8 rounded-[10px] grid place-items-center text-[13px] ${view===item.k ? "bg-white/10" : "bg-black/[0.06]"}`}>{item.i}</span>
                  <span className="flex-1">
                    <span className="block text-[13px] font-medium leading-none">{item.k}</span>
                    <span className={`block text-[11px] mt-1 ${view===item.k ? "text-white/60" : "text-black/50"}`}>{item.d}</span>
                  </span>
                  {view===item.k && <span className="w-1.5 h-1.5 rounded-full bg-[#FF9A62]" />}
                </button>
              ))}
            </nav>

            <div className="card p-4 bg-gradient-to-br from-[#6C5CE7]/[0.08] to-[#FF9A62]/[0.12] border-[#6C5CE7]/10">
              <div className="text-[11px] font-semibold tracking-widest uppercase opacity-60">Preview</div>
              <div className="mt-3 rounded-[16px] bg-[#111113] text-white p-4 relative overflow-hidden">
                <div className="absolute -right-8 -top-8 w-24 h-24 rounded-full bg-gradient-to-br from-[#6C5CE7] to-[#FF9A62] blur-[1px] opacity-80" />
                <div className="relative flex items-center gap-2 mb-3">
                  <div className="w-8 h-8 rounded-full bg-white grid place-items-center overflow-hidden shadow-sm">
                    <img src={logoUrl} alt="logo" style={{width:"28px", height:"28px", objectFit:"contain", borderRadius:"8px"}} />
                  </div>
                  <span className="text-[11px] font-medium opacity-80">SheetSense • 20MB Excel</span>
                </div>
                <div className="relative flex justify-between text-[11px] opacity-80"><span>Rows</span><span>{rows.length || "—"}</span></div>
                <div className="relative flex justify-between text-[11px] opacity-80 mt-1"><span>Cols</span><span>{columns.length || "—"}</span></div>
                <div className="relative flex justify-between text-[11px] opacity-80 mt-1"><span>Quality</span><span className="text-[#FF9A62]">{rows.length ? `${qualityScore}%` : "—"}</span></div>
                <div className="relative mt-3 flex gap-1.5">
                  {[0.7,0.4,0.9,0.5,0.8].map((h,i)=><div key={i} className="flex-1 rounded-full bg-white/20" style={{height: `${8 + h*18}px`}} />)}
                </div>
              </div>
              <div className="mt-3 text-[12px] leading-[1.4] opacity-70">Everything runs locally. XLSX up to 20MB • CSV, JSON parsed with SheetJS. AI cleans instantly. Optimized for ~100k rows.</div>
              {isFormatted && <div className="mt-3 text-[11px] px-2.5 py-1.5 rounded-full bg-[#6C5CE7] text-white inline-flex items-center gap-1">✦ AI Formatted • {formatLog.length} fixes</div>}
            </div>

            <div className="mt-auto text-[11px] opacity-50 px-2">Everything local • v2.1.0 • 20MB Excel • {xlsxReady ? "XLSX ✓" : "XLSX loading"}</div>
          </aside>

          {/* Main */}
          <main className="flex-1 min-w-0 pb-[88px] md:pb-6">
            {/* Upload */}
            {view==="Upload" && (
              <div className="px-4 md:px-8 py-6 md:py-10 max-w-[1080px]">
                {!user && (
                  <div className="mb-6 rounded-[20px] border border-[#6C5CE7]/20 bg-gradient-to-r from-[#6C5CE7]/10 via-white to-[#FF9A62]/10 px-4 py-3 flex items-center justify-between gap-3">
                    <div className="flex items-center gap-2.5">
                      <img src={logoUrl} alt="SheetSense" style={{width:"28px", height:"28px", objectFit:"contain", borderRadius:"8px"}} className="shadow-sm border border-black/5 bg-white" />
                      <div className="text-[13px]"><span className="font-semibold">Welcome to SheetSense</span><span className="opacity-70"> — 20MB XLSX + AI cleaning. Drop any sheet.</span></div>
                    </div>
                    <button onClick={()=>{setAuthMode("signup"); setAuthOpen(true);}} className="shrink-0 h-8 px-3 rounded-full bg-[#111113] text-white text-[12px] font-medium">Start free</button>
                  </div>
                )}
                <div className="flex flex-wrap items-start justify-between gap-6">
                  <div className="max-w-[560px]">
                    <div className="inline-flex items-center gap-2 text-[11px] font-medium px-3 py-1.5 rounded-full bg-white border border-black/10 shadow-sm">
                      <span className="w-1.5 h-1.5 rounded-full bg-[#6C5CE7] animate-pulse" /> New — XLSX + AI Format
                    </div>
                    <h1 className="serif text-[40px] md:text-[56px] leading-[0.95] tracking-[-0.03em] mt-4">Turn sheets<br/>into <span className="italic font-normal text-[#6C5CE7]">insights</span></h1>
                    <p className="mt-4 text-[15px] leading-[1.6] opacity-70 max-w-[44ch]">Upload CSV, JSON, XLSX, XLS up to 20MB. SheetJS parses locally with cellDates. AI detects messy headers, dates, currency, blanks and fixes in one click. Optimized for ~100k rows.</p>
                    <div className="mt-6 flex gap-2">
                      <button onClick={()=>fileInputRef.current?.click()} className="h-10 px-5 rounded-full bg-[#111113] text-white text-[13px] font-semibold">Browse files</button>
                      <button onClick={()=>{ const parsed=parseCSV(DEMO_CSV); const cloned=JSON.parse(JSON.stringify(parsed)); setOriginalRows(cloned); setRows(parsed); setIsFormatted(false); setFormatLog([]); setWorkbookMeta(null); setWbObj(null); setTimeout(()=>analyzeRows(cloned),100); showToast("Demo loaded"); }} className="h-10 px-5 rounded-full bg-white border border-black/10 text-[13px] font-medium">Load demo CSV</button>
                    </div>
                    <div className="mt-4 flex gap-2 text-[11px]">
                      <span className="px-2.5 py-1 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/20">CSV</span>
                      <span className="px-2.5 py-1 rounded-full bg-[#FF9A62]/15 border border-[#FF9A62]/30">XLSX / XLS</span>
                      <span className="px-2.5 py-1 rounded-full bg-black/5 border border-black/10">JSON</span>
                      <span className="px-2.5 py-1 rounded-full bg-black/5 border border-black/10">20MB max • local</span>
                    </div>
                  </div>
                  <div className="w-full md:w-[360px] card p-3">
                    <div className="rounded-[16px] bg-[#111113] text-white p-4">
                      <div className="flex items-center justify-between text-[12px] opacity-80">
                        <span className="flex items-center gap-2"><span className="w-6 h-6 rounded-full bg-white grid place-items-center overflow-hidden"><img src={logoUrl} alt="logo" style={{width:"20px", height:"20px", objectFit:"contain"}} /></span>SheetSense Live</span>
                        <span className="px-2 py-0.5 rounded-full bg-white/10 text-[10px]">{xlsxReady ? "XLSX 20MB ✓" : "LOADING"}</span>
                      </div>
                      <div className="mt-4 grid grid-cols-3 gap-3">
                        <div className="rounded-[12px] bg-white/[0.06] border border-white/10 p-3"><div className="text-[11px] opacity-60">Rows</div><div className="text-[18px] font-semibold mt-1">{rows.length || 248}</div></div>
                        <div className="rounded-[12px] bg-white/[0.06] border border-white/10 p-3"><div className="text-[11px] opacity-60">Cols</div><div className="text-[18px] font-semibold mt-1">{columns.length || 6}</div></div>
                        <div className="rounded-[12px] bg-[#FF9A62] text-[#111113] p-3"><div className="text-[11px] opacity-80">Quality</div><div className="text-[18px] font-bold mt-1">{rows.length?`${qualityScore}%`:"94%"}</div></div>
                      </div>
                      <div className="mt-4 h-[54px] flex items-end gap-1">
                        {[12,22,10,28,16,24,8,20].map((h,i)=><div key={i} className="flex-1 rounded-full" style={{height:h*2, background: i%2? "rgba(108,92,231,0.9)" : "rgba(255,154,98,0.9)"}} />)}
                      </div>
                      <div className="mt-3 text-[10px] opacity-50">Optimized for up to 20MB Excel files — ~100k rows</div>
                    </div>
                  </div>
                </div>

                {/* Dropzone */}
                <div
                  onDragOver={e=>{e.preventDefault(); setDragActive(true);}}
                  onDragLeave={()=>setDragActive(false)}
                  onDrop={e=>{e.preventDefault(); setDragActive(false); const f=e.dataTransfer.files[0]; if(f) handleFile(f);}}
                  className={`mt-8 rounded-[24px] border-[1.5px] border-dashed bg-white p-8 md:p-10 text-center transition ${dragActive ? "border-[#6C5CE7] bg-[#6C5CE7]/[0.04] scale-[0.99]" : "border-black/15"}`}
                >
                  <div className="mx-auto w-12 h-12 rounded-[14px] bg-white border border-black/10 grid place-items-center shadow-sm overflow-hidden">
                    <img src={logoUrl} alt="upload" style={{width:"32px", height:"32px", objectFit:"contain", borderRadius:"8px"}} />
                  </div>
                  <div className="mt-4 text-[15px] font-medium">Drop CSV, JSON, XLSX, XLS here</div>
                  <div className="mt-1 text-[13px] opacity-60">.csv, .json, .xlsx, .xls up to 20MB • parsed locally with SheetJS {xlsxReady?"✓":"(loading…)"} • Optimized for ~100k rows</div>
                  <button onClick={()=>fileInputRef.current?.click()} className="mt-5 h-9 px-4 rounded-full border border-black/10 bg-[#111113] text-white text-[13px] font-medium">Browse</button>
                  <input ref={fileInputRef} type="file" accept=".csv,.json,.xlsx,.xls" className="hidden" onChange={e=>{ const f=e.target.files?.[0]; if(f) handleFile(f); }} />
                  {rows.length>0 && <div className="mt-4 text-[12px] opacity-60">Loaded {rows.length} rows • {columns.length} cols • {numericCols.length} numeric • {workbookMeta?.fileName||"demo"} {lastFileSize? `• ${formatFileSize(lastFileSize)}`:""} {lastFileSize? `• ~${rows.length} rows`:""}</div>}
                  {workbookMeta && workbookMeta.sheets.length>1 && (
                    <div className="mt-4 flex items-center justify-center gap-2">
                      <span className="text-[12px] opacity-60">Sheet:</span>
                      <select value={selectedSheet} onChange={e=>{ const name=e.target.value; setSelectedSheet(name); if(wbObj) parseSheetFromWB(wbObj,name); }} className="h-8 rounded-full border border-black/10 bg-[#FCFAF3] px-3 text-[12px] font-medium">
                        {workbookMeta.sheets.map(s=><option key={s} value={s}>{s}</option>)}
                      </select>
                      <span className="text-[11px] px-2 py-0.5 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/20">{workbookMeta.sheets.length} sheets</span>
                    </div>
                  )}
                  {isParsing && (
                    <div className="mt-6 max-w-[420px] mx-auto">
                      <div className="flex justify-between text-[11px] mb-1.5">
                        <span className="font-medium truncate">{parsingFileName} • {formatFileSize(lastFileSize)}</span>
                        <span className="opacity-60">{uploadProgress}%</span>
                      </div>
                      <div className="h-2 rounded-full bg-black/10 overflow-hidden">
                        <div className="h-full bg-gradient-to-r from-[#6C5CE7] to-[#FF9A62] transition-all duration-200" style={{width:`${uploadProgress}%`}} />
                      </div>
                      <div className="mt-1 text-[10px] opacity-50">{uploadProgress<100 ? "Parsing large file… SheetJS cellDates mode" : "Parsing complete"}</div>
                    </div>
                  )}
                </div>

                {/* AI Format & Clean Panel */}
                {rows.length>0 && (
                  <div className="mt-8 card overflow-hidden border-[#6C5CE7]/20 shadow-[0_12px_40px_rgba(108,92,231,0.12)]">
                    <div className="bg-gradient-to-br from-[#6C5CE7] to-[#7B6EF0] text-white p-5 flex items-start justify-between gap-4">
                      <div className="flex gap-3">
                        <div className="w-10 h-10 rounded-[14px] bg-white border border-white/20 grid place-items-center overflow-hidden shadow-sm">
                          <img src={logoUrl} alt="AI" style={{width:"32px", height:"32px", objectFit:"contain", borderRadius:"8px"}} />
                        </div>
                        <div>
                          <div className="flex items-center gap-2">
                            <span className="font-semibold text-[15px]">AI Format & Clean</span>
                            <span className="text-[10px] px-2 py-0.5 rounded-full bg-white/15 border border-white/20">SheetJS + AI</span>
                          </div>
                          <div className="text-[12px] opacity-80 mt-1 max-w-[50ch]">Auto-detected {aiIssues.length} issues • Headers, dates, currency, blanks • One-click fix, all local.</div>
                        </div>
                      </div>
                      <div className="flex flex-col items-end gap-2">
                        <div className="text-[11px] px-2.5 py-1 rounded-full bg-[#FF9A62] text-[#111113] font-semibold">AI Confidence: {aiConfidence}%</div>
                        {isFormatted && <div className="text-[11px] px-2 py-1 rounded-full bg-white/15 border border-white/20">Formatted ✓</div>}
                      </div>
                    </div>

                    {headerMapPreview.length>0 && (
                      <div className="px-5 py-3 bg-[#6C5CE7]/[0.06] border-b border-[#6C5CE7]/10 flex flex-wrap gap-2 items-center">
                        <span className="text-[11px] font-medium opacity-60">Header intelligence:</span>
                        {headerMapPreview.map((h,i)=>(
                          <span key={i} className="text-[11px] px-2.5 py-1 rounded-full bg-white border border-black/10 flex items-center gap-1.5">
                            <span className="opacity-50 line-through">{h.before}</span>
                            <span className="opacity-40">→</span>
                            <span className="font-medium text-[#6C5CE7]">{h.after}</span>
                          </span>
                        ))}
                      </div>
                    )}

                    <div className="p-5">
                      <div className="flex items-center justify-between mb-3">
                        <div className="text-[13px] font-semibold">Detected issues ({aiIssues.length})</div>
                        <div className="text-[11px] opacity-60">{rows.length} rows • {columns.length} cols • Quality {qualityScore}%</div>
                      </div>

                      {aiIssues.length===0 ? (
                        <div className="rounded-[16px] bg-[#6C5CE7]/5 border border-[#6C5CE7]/10 p-4 text-[13px]">
                          <div className="font-medium">✨ All clean — no issues detected</div>
                          <div className="opacity-60 text-[12px] mt-1">Headers normalized, dates ISO, no blanks. Ready for Analyze & Visuals.</div>
                        </div>
                      ) : (
                        <div className="grid md:grid-cols-2 gap-2.5">
                          {aiIssues.map(issue=>(
                            <label key={issue.id} className={`group flex gap-3 p-3 rounded-[14px] border text-left cursor-pointer transition ${aiToggles[issue.id] ? "bg-white border-[#6C5CE7]/25 shadow-sm" : "bg-black/[0.02] border-black/10 opacity-70"}`}>
                              <input type="checkbox" checked={!!aiToggles[issue.id]} onChange={e=> setAiToggles({...aiToggles, [issue.id]: e.target.checked})} className="mt-0.5 accent-[#6C5CE7] w-4 h-4" />
                              <span className="flex-1 min-w-0">
                                <span className="flex items-center gap-2">
                                  <span className="text-[12px] font-semibold">{issue.fixLabel}</span>
                                  <span className={`text-[10px] px-1.5 py-0.5 rounded-full border ${issue.severity==="high"?"bg-red-50 border-red-200 text-red-700": issue.severity==="med"?"bg-[#FF9A62]/10 border-[#FF9A62]/20 text-[#111113]":"bg-black/5 border-black/10"}`}>{issue.count}</span>
                                </span>
                                <span className="block text-[11px] opacity-60 mt-1 leading-[1.3]">{issue.description}</span>
                                <span className="block text-[11px] mt-1 font-mono bg-[#FCFAF3] border border-black/5 rounded-full px-2 py-0.5 w-fit">{issue.example}</span>
                              </span>
                            </label>
                          ))}
                        </div>
                      )}

                      <div className="mt-5 flex flex-wrap gap-2">
                        <button onClick={applyAiFixes} className="h-10 px-5 rounded-full bg-[#111113] text-white text-[13px] font-semibold flex items-center gap-2">
                          <span>✦</span> Apply AI Fixes ({Object.values(aiToggles).filter(Boolean).length})
                        </button>
                        <button onClick={undoFormat} disabled={!isFormatted && !originalRows} className={`h-10 px-5 rounded-full border text-[13px] font-medium ${isFormatted ? "bg-white border-black/10 hover:bg-black/5" : "bg-black/5 border-black/10 opacity-50"}`}>Undo</button>
                        <button onClick={()=>{ if(originalRows) analyzeRows(originalRows); else analyzeRows(rows); showToast("Re-analyzed"); }} className="h-10 px-4 rounded-full bg-white border border-black/10 text-[13px] font-medium">Re-analyze</button>
                        {formatLog.length>0 && <span className="h-10 px-3 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/15 text-[11px] flex items-center gap-1.5"><span className="w-1.5 h-1.5 rounded-full bg-[#6C5CE7]" />{formatLog.length} fixes applied</span>}
                      </div>

                      {formatLog.length>0 && (
                        <div className="mt-4 rounded-[14px] bg-[#111113] text-white p-3.5">
                          <div className="text-[11px] uppercase tracking-widest opacity-60">Formatting log</div>
                          <div className="mt-2 space-y-1">
                            {formatLog.map((l,i)=><div key={i} className="text-[12px] flex gap-2"><span className="text-[#FF9A62]">•</span><span className="opacity-90">{l}</span></div>)}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                <div className="mt-6 grid md:grid-cols-3 gap-3">
                  {[
                    {t:"XLSX real parsing 20MB", d:"SheetJS reads .xlsx/.xls ArrayBuffer with cellDates, optimized for up to 20MB — ~100k rows. Sheet selector for multi-sheet."},
                    {t:"Header intelligence", d:"Finds first text-heavy row as header, Title Case, removes _ - special chars, dedupes."},
                    {t:"File size up to 20MB", d:"Progress bar for >5MB files, size in meta: demo.xlsx • 14.2 MB • 3400 rows • 8 cols. Memory-safe parsing."},
                  ].map(c=>(
                    <div key={c.t} className="card p-4">
                      <div className="text-[13px] font-semibold">{c.t}</div>
                      <div className="text-[12px] opacity-60 mt-1 leading-[1.5]">{c.d}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}

            {view==="Data" && (
              <div className="px-4 md:px-8 py-6 max-w-[1180px]">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <div className="flex items-center gap-3">
                    <h2 className="serif text-[28px]">Data</h2>
                    {rows.length>0 && <span className="text-[11px] px-2.5 py-1 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/20 text-[#6C5CE7] font-medium">{qualityScore}% quality</span>}
                    {isFormatted && <span className="text-[11px] px-2.5 py-1 rounded-full bg-[#FF9A62]/15 border border-[#FF9A62]/30 font-medium">AI formatted</span>}
                  </div>
                  <div className="flex gap-2 flex-wrap">
                    {originalRows && (
                      <div className="flex items-center gap-1 p-1 rounded-full bg-black/5 border border-black/10">
                        <button onClick={()=>setShowRaw(false)} className={`h-7 px-3 rounded-full text-[12px] font-medium ${!showRaw?"bg-[#111113] text-white":"opacity-70"}`}>Formatted</button>
                        <button onClick={()=>setShowRaw(true)} className={`h-7 px-3 rounded-full text-[12px] font-medium ${showRaw?"bg-[#111113] text-white":"opacity-70"}`}>Raw</button>
                      </div>
                    )}
                    <div className="relative">
                      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search rows…" className="h-9 w-[180px] md:w-[220px] rounded-full border border-black/10 bg-white px-4 text-[13px] outline-none focus:border-[#6C5CE7]/40" />
                    </div>
                    <button onClick={()=>{ setView("Upload"); setTimeout(()=> window.scrollTo({top:0, behavior:"smooth"}),100); }} className="h-9 px-4 rounded-full bg-white border border-black/10 text-[13px] font-medium">✦ Format</button>
                    <button onClick={()=>{
                      if(!rows.length) return showToast("No data");
                      const cols = showRaw && originalRows ? Object.keys(originalRows[0]) : columns;
                      const base = showRaw && originalRows ? originalRows : rows;
                      const csv = [cols.join(","), ...base.map(r=>cols.map(c=>`"${String(r[c]??"").replace(/"/g,'""')}"`).join(","))].join("\n");
                      const blob = new Blob([csv], {type:"text/csv"}); const url = URL.createObjectURL(blob);
                      const a = document.createElement("a"); a.href=url; a.download="sheetsense_export.csv"; a.click(); URL.revokeObjectURL(url);
                      showToast("CSV exported");
                    }} className="h-9 px-4 rounded-full bg-white border border-black/10 text-[13px] font-medium">Export CSV</button>
                  </div>
                </div>
                {!rows.length ? (
                  <div className="mt-10 card p-12 text-center">
                    <div className="mx-auto w-12 h-12 rounded-[14px] bg-white border border-black/10 grid place-items-center overflow-hidden shadow-sm"><img src={logoUrl} alt="logo" style={{width:"32px", height:"32px", objectFit:"contain"}} /></div>
                    <div className="mt-3 text-[15px] font-medium">No sheet loaded</div>
                    <div className="text-[13px] opacity-60 mt-1">Go to Upload and drop CSV, JSON, XLSX, XLS (up to 20MB) or load demo. Optimized for ~100k rows.</div>
                    <button onClick={()=>setView("Upload")} className="mt-4 h-9 px-4 rounded-full bg-[#111113] text-white text-[13px]">Go to Upload</button>
                  </div>
                ) : (
                  <>
                    <div className="mt-4 flex flex-wrap gap-2 text-[11px]">
                      <span className="px-2.5 py-1 rounded-full bg-white border border-black/10">Showing {filteredRows.length} of {displayRows.length} • {showRaw ? "Raw" : "Formatted"} {lastFileSize? `• ${formatFileSize(lastFileSize)}`:""}</span>
                      <span className="px-2.5 py-1 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/15">Quality {qualityScore}% • {aiIssues.length} issues</span>
                      {workbookMeta && <span className="px-2.5 py-1 rounded-full bg-[#FF9A62]/10 border border-[#FF9A62]/20">{workbookMeta.fileName} • {lastFileSize? `${formatFileSize(lastFileSize)} • `:""}{displayRows.length} rows • {displayCols.length} cols • {selectedSheet}</span>}
                      {formatLog.length>0 && !showRaw && <span className="px-2.5 py-1 rounded-full bg-[#111113] text-white">{formatLog.length} AI fixes applied</span>}
                    </div>
                    <div className="mt-5 card overflow-hidden">
                      <div className="overflow-auto max-h-[64vh]">
                        <table className="w-full text-[13px] border-collapse">
                          <thead className="sticky top-0 z-10 bg-[#FCFAF3] border-b border-black/10">
                            <tr>
                              {displayCols.map(col=>{
                                const t = detectType(displayRows.map(r=>r[col]));
                                return <th key={col} className="text-left font-medium px-4 py-3 whitespace-nowrap">
                                  <div className="flex items-center gap-2"><span>{col}</span><span className={`text-[10px] px-1.5 py-0.5 rounded-full border ${t==="money"?"bg-[#FF9A62]/15 border-[#FF9A62]/30": t==="number"?"bg-[#6C5CE7]/10 border-[#6C5CE7]/20": t==="date"?"bg-black/5 border-black/10":"bg-white border-black/10"}`}>{t}</span></div>
                                </th>
                              })}
                            </tr>
                          </thead>
                          <tbody>
                            {filteredRows.map((r,i)=>(
                              <tr key={i} className="border-b border-black/[0.06] hover:bg-black/[0.02]">
                                {displayCols.map(c=><td key={c} className="px-4 py-2.5 whitespace-nowrap max-w-[220px] truncate">{String(r[c]??"")}</td>)}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div className="px-4 py-2.5 text-[11px] opacity-60 border-t border-black/5 flex justify-between"><span>Showing {filteredRows.length} of {displayRows.length}</span><span>Limit {prefs.maxRows} • type detection {prefs.autoDetect ? "on" : "off"} • {showRaw?"raw":"AI formatted"} {lastFileSize? `• ${formatFileSize(lastFileSize)}`:""}</span></div>
                    </div>
                  </>
                )}
              </div>
            )}

            {view==="Analyze" && (
              <div className="px-4 md:px-8 py-6 max-w-[1180px]">
                <div className="flex items-center gap-3">
                  <h2 className="serif text-[28px]">Analyze</h2>
                  {isFormatted && <span className="text-[11px] px-2.5 py-1 rounded-full bg-[#6C5CE7] text-white">AI formatted • {qualityScore}% quality</span>}
                  {lastFileSize>0 && <span className="text-[11px] px-2.5 py-1 rounded-full bg-white border border-black/10">{formatFileSize(lastFileSize)} • {rows.length} rows</span>}
                </div>
                {!rows.length ? (
                  <div className="mt-8 card p-10 text-center text-[13px] opacity-70">Load a sheet first. Supports up to 20MB Excel.</div>
                ) : (
                  <>
                    <div className="mt-5 grid grid-cols-2 md:grid-cols-4 gap-3">
                      <div className="card p-4"><div className="text-[11px] uppercase tracking-widest opacity-50">Total rows</div><div className="serif text-[28px] mt-1">{stats?.total}</div></div>
                      <div className="card p-4"><div className="text-[11px] uppercase tracking-widest opacity-50">Numeric cols</div><div className="serif text-[28px] mt-1">{stats?.numericCols}</div></div>
                      <div className="card p-4"><div className="text-[11px] uppercase tracking-widest opacity-50">Sum ({numericCols[0]})</div><div className="serif text-[28px] mt-1">{stats ? `$${stats.sum.toLocaleString(undefined,{maximumFractionDigits:0})}` : "—"}</div></div>
                      <div className="card p-4 bg-[#111113] text-white"><div className="text-[11px] uppercase tracking-widest opacity-60">Average</div><div className="serif text-[28px] mt-1 text-[#FF9A62]">{stats ? `$${stats.avg.toFixed(2)}` : "—"}</div></div>
                    </div>

                    {formatLog.length>0 && (
                      <div className="mt-6 card p-4 border-[#6C5CE7]/20 bg-gradient-to-br from-[#6C5CE7]/[0.06] to-white">
                        <div className="flex items-center gap-2">
                          <span className="w-7 h-7 rounded-full bg-white border border-black/10 grid place-items-center overflow-hidden"><img src={logoUrl} alt="logo" style={{width:"20px", height:"20px", objectFit:"contain"}} /></span>
                          <span className="text-[13px] font-semibold">AI Formatting Log</span>
                          <span className="text-[11px] px-2 py-0.5 rounded-full bg-[#FF9A62] text-[#111113] font-medium">{aiConfidence}% confidence</span>
                        </div>
                        <div className="mt-3 grid md:grid-cols-2 gap-2">
                          {formatLog.map((l,i)=>(
                            <div key={i} className="rounded-[12px] bg-white border border-black/10 px-3 py-2 text-[12px] flex gap-2"><span className="text-[#6C5CE7]">✓</span><span>{l}</span></div>
                          ))}
                        </div>
                        {originalRows && <div className="mt-3 text-[11px] opacity-60">Original: {originalRows.length} rows • Formatted: {rows.length} rows • Headers: {originalColumns.length} → {columns.length}</div>}
                      </div>
                    )}

                    <div className="mt-6 grid md:grid-cols-3 gap-3">
                      {columns.map(col=>{
                        const vals = rows.map(r=>String(r[col]??"")).filter(Boolean);
                        const uniq = Array.from(new Set(vals));
                        const top = Object.entries(vals.reduce((a,v)=>{a[v]=(a[v]||0)+1; return a;}, {} as any)).sort((a:any,b:any)=>b[1]-a[1]).slice(0,4);
                        return (
                          <div key={col} className="card p-4">
                            <div className="flex items-center justify-between"><span className="text-[13px] font-semibold">{col}</span><span className="text-[10px] px-2 py-0.5 rounded-full bg-black/5 border border-black/10">{detectType(rows.map(r=>r[col]))}</span></div>
                            <div className="mt-3 text-[11px] opacity-60">{uniq.length} unique • {vals.length} filled</div>
                            <div className="mt-3 space-y-1.5">
                              {top.map(([k,c]: any)=>(
                                <div key={k} className="flex items-center gap-2 text-[12px]"><div className="flex-1 truncate opacity-80">{k}</div><div className="text-[11px] px-1.5 py-0.5 rounded-full bg-[#6C5CE7]/10">{c}</div></div>
                              ))}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </>
                )}
              </div>
            )}

            {view==="Visuals" && (
              <div className="px-4 md:px-8 py-6 max-w-[1180px]">
                <div className="flex flex-wrap items-center justify-between gap-3">
                  <h2 className="serif text-[28px]">Visuals</h2>
                  <div className="flex gap-2">
                    <select value={chartType} onChange={e=>setChartType(e.target.value as any)} className="h-9 rounded-full border border-black/10 bg-white px-3 text-[13px]">
                      <option>Bar</option><option>Line</option><option>Area</option><option>Pie</option>
                    </select>
                    <select value={yAxis} onChange={e=>setYAxis(e.target.value)} className="h-9 rounded-full border border-black/10 bg-white px-3 text-[13px]">
                      {columns.map(c=><option key={c}>{c}</option>)}
                    </select>
                  </div>
                </div>
                {!rows.length ? (
                  <div className="mt-8 card p-10 text-center text-[13px] opacity-70">Load data to render canvas charts.</div>
                ) : (
                  <div className="mt-5 card p-3 md:p-5">
                    <div className="flex items-center justify-between px-2 py-1">
                      <div className="text-[12px] opacity-60">{chartType} • {yAxis} • {rows.length} points • gradient violet→peach {isFormatted?`• AI cleaned ${qualityScore}%`:""}</div>
                      <div className="flex gap-1.5"><span className="w-2 h-2 rounded-full bg-[#6C5CE7]" /><span className="w-2 h-2 rounded-full bg-[#FF9A62]" /></div>
                    </div>
                    <canvas ref={canvasRef} className="w-full h-[340px] md:h-[420px] rounded-[16px] bg-[#FCFAF3] border border-black/[0.06] mt-3" />
                    <div className="mt-3 grid grid-cols-3 gap-2 text-[11px]">
                      <div className="rounded-[12px] bg-[#6C5CE7]/10 border border-[#6C5CE7]/15 px-3 py-2">Canvas 2D • no libs</div>
                      <div className="rounded-[12px] bg-[#FF9A62]/10 border border-[#FF9A62]/15 px-3 py-2">Gradient bars • grid</div>
                      <div className="rounded-[12px] bg-black/5 border border-black/10 px-3 py-2">Responsive • DPR aware</div>
                    </div>
                  </div>
                )}
              </div>
            )}

            {view==="Currency" && (
              <div className="px-4 md:px-8 py-6 max-w-[900px]">
                <h2 className="serif text-[28px]">Currency</h2>
                <div className="mt-5 card p-5">
                  <div className="flex flex-wrap gap-3 items-center justify-between">
                    <div>
                      <div className="text-[13px] font-medium">Detected money columns</div>
                      <div className="mt-2 flex gap-1.5 flex-wrap">
                        {moneyCols.length ? moneyCols.map(c=><span key={c} className="text-[11px] px-2.5 py-1 rounded-full bg-[#FF9A62]/15 border border-[#FF9A62]/30">{c}</span>) : <span className="text-[12px] opacity-60">No money column detected — add Amount/Price</span>}
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="text-[12px] opacity-60">Target</span>
                      <select value={targetCurrency} onChange={e=>setTargetCurrency(e.target.value)} className="h-9 rounded-full border border-black/10 bg-white px-3 text-[13px] font-medium">
                        <option>USD</option><option>EUR</option><option>INR</option><option>GBP</option>
                      </select>
                    </div>
                  </div>
                  <div className="mt-6 border-t border-black/10 pt-4">
                    <div className="text-[11px] uppercase tracking-widest opacity-50 mb-3">Conversion preview (demo rates)</div>
                    <div className="grid gap-2">
                      {rows.slice(0,8).map((r,i)=>{
                        const col = moneyCols[0] || columns.find(c=>/amount/i.test(c)) || columns[2];
                        const raw = Number(String(r[col]??"0").replace(/[^0-9.-]/g,"")) || 0;
                        const rate = RATES[targetCurrency] / RATES[(r.Currency||r.currency||"USD").toString().slice(0,3) || "USD" as any] || RATES[targetCurrency];
                        const conv = raw * rate;
                        return (
                          <div key={i} className="flex items-center justify-between rounded-[14px] border border-black/10 bg-white px-4 py-2.5">
                            <div className="text-[13px]"><span className="opacity-60">{r.Date||r.date||`#${i+1}`}</span><span className="mx-2">•</span><span className="font-medium">{r.Category||r.category||col}</span></div>
                            <div className="flex items-center gap-3"><span className="text-[13px]">${raw.toLocaleString()}</span><span className="text-[11px] opacity-40">→</span><span className="text-[13px] font-semibold text-[#6C5CE7]">{targetCurrency} {conv.toLocaleString(undefined,{maximumFractionDigits:2})}</span></div>
                          </div>
                        );
                      })}
                    </div>
                    <div className="mt-4 text-[11px] opacity-50">Rates: USD 1 • EUR 0.92 • INR 83.5 • GBP 0.79 — demo only, runs locally.</div>
                  </div>
                </div>
              </div>
            )}

            {view==="Settings" && (
              <div className="px-4 md:px-8 py-6 max-w-[1100px] flex flex-col md:flex-row gap-4">
                <div className="md:w-[220px] shrink-0 card p-2 h-fit">
                  {(["Profile","Appearance","Preferences","Data","About"] as const).map(t=>(
                    <button key={t} onClick={()=>setSettingsTab(t)} className={`w-full text-left px-3 py-2.5 rounded-[12px] text-[13px] font-medium ${settingsTab===t?"bg-[#111113] text-white":"hover:bg-black/5"}`}>{t}</button>
                  ))}
                  <div className="mt-3 px-3 py-2 text-[11px] opacity-50">Version 2.1.0<br/>Everything local • XLSX</div>
                </div>
                <div className="flex-1 card p-5 md:p-7">
                  {settingsTab==="Profile" && (
                    <div>
                      <h3 className="serif text-[22px]">Profile</h3>
                      <div className="mt-6 flex gap-4 items-center">
                        <div className="w-14 h-14 rounded-[18px] bg-[#111113] text-white grid place-items-center text-[20px] font-semibold">{(user?.name||"G").charAt(0).toUpperCase()}</div>
                        <div><div className="text-[14px] font-semibold">{user?.name||"Guest"}</div><div className="text-[12px] opacity-60">{user?.email||"guest@sheetsense.local"}</div><div className="mt-1 inline-flex text-[10px] px-2 py-0.5 rounded-full bg-[#6C5CE7]/10 border border-[#6C5CE7]/20">{user?.type||"Guest"} • Free</div></div>
                      </div>
                      <div className="mt-6 grid md:grid-cols-2 gap-4">
                        <label className="text-[12px]">Name<input value={authForm.name || user?.name || ""} onChange={e=>setAuthForm({...authForm, name:e.target.value})} placeholder="Your name" className="mt-1 w-full h-10 rounded-[12px] border border-black/10 px-3 text-[13px]" /></label>
                        <label className="text-[12px]">Email<input value={authForm.email || user?.email || ""} onChange={e=>setAuthForm({...authForm, email:e.target.value})} placeholder="you@company.com" className="mt-1 w-full h-10 rounded-[12px] border border-black/10 px-3 text-[13px]" /></label>
                      </div>
                      <button onClick={()=>{ if(authForm.name || authForm.email){ setUser({name:authForm.name||user?.name||"User", email:authForm.email||user?.email||"user@sheetsense.local", type:"Free"}); showToast("Profile saved"); } else showToast("Enter name or email");}} className="mt-5 h-9 px-4 rounded-full bg-[#111113] text-white text-[13px] font-medium">Save changes</button>
                    </div>
                  )}
                  {settingsTab==="Appearance" && (
                    <div>
                      <h3 className="serif text-[22px]">Appearance</h3>
                      <div className="mt-6 space-y-5">
                        <div className="flex items-center justify-between"><span className="text-[13px]">Theme</span><span className="text-[12px] px-3 py-1 rounded-full bg-black/5 border border-black/10">Light only • System disabled</span></div>
                        <div className="flex items-center justify-between"><span className="text-[13px]">Accent preview</span><span className="flex gap-2"><span className="w-6 h-6 rounded-full bg-[#6C5CE7] border border-black/10" /><span className="w-6 h-6 rounded-full bg-[#FF9A62] border border-black/10" /></span></div>
                        {[
                          {k:"compact", l:"Compact mode", d:"Reduce padding and card spacing"},
                          {k:"livePreview", l:"Show live preview", d:"Mini chart in sidebar"},
                        ].map(row=>(
                          <div key={row.k} className="flex items-center justify-between">
                            <div><div className="text-[13px] font-medium">{row.l}</div><div className="text-[11px] opacity-60">{row.d}</div></div>
                            <button onClick={()=>setPrefs({...prefs, [row.k]: !(prefs as any)[row.k]})} className={`w-10 h-6 rounded-full p-0.5 transition ${ (prefs as any)[row.k] ? "bg-[#111113]" : "bg-black/15"}`}><div className={`w-5 h-5 rounded-full bg-white transition ${ (prefs as any)[row.k] ? "translate-x-4" : "translate-x-0"}`} /></button>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                  {settingsTab==="Preferences" && (
                    <div>
                      <h3 className="serif text-[22px]">Preferences</h3>
                      <div className="mt-6 grid md:grid-cols-2 gap-4">
                        <label className="text-[12px]">Default currency<select value={prefs.currency} onChange={e=>setPrefs({...prefs, currency:e.target.value})} className="mt-1 w-full h-10 rounded-[12px] border border-black/10 px-3 text-[13px]"><option>USD</option><option>EUR</option><option>INR</option><option>GBP</option></select></label>
                        <label className="text-[12px]">Date format<select value={prefs.dateFormat} onChange={e=>setPrefs({...prefs, dateFormat:e.target.value})} className="mt-1 w-full h-10 rounded-[12px] border border-black/10 px-3 text-[13px]"><option>MM/DD/YYYY</option><option>DD/MM/YYYY</option><option>ISO</option></select></label>
                        <label className="text-[12px]">Number format<select value={prefs.numberFormat} onChange={e=>setPrefs({...prefs, numberFormat:e.target.value})} className="mt-1 w-full h-10 rounded-[12px] border border-black/10 px-3 text-[13px]"><option>en-US</option><option>en-GB</option><option>de-DE</option><option>en-IN</option></select></label>
                        <div className="flex items-center justify-between rounded-[12px] border border-black/10 px-3 h-10"><span className="text-[12px]">Auto-detect types</span><button onClick={()=>setPrefs({...prefs, autoDetect:!prefs.autoDetect})} className={`w-10 h-6 rounded-full p-0.5 ${prefs.autoDetect?"bg-[#111113]":"bg-black/15"}`}><div className={`w-5 h-5 rounded-full bg-white ${prefs.autoDetect?"translate-x-4":""} transition`} /></button></div>
                      </div>
                      <div className="mt-6">
                        <div className="text-[12px] flex justify-between"><span>Max rows</span><span className="opacity-60">{prefs.maxRows}</span></div>
                        <input type="range" min={50} max={500} step={50} value={prefs.maxRows} onChange={e=>setPrefs({...prefs, maxRows: Number(e.target.value)})} className="w-full mt-2 accent-[#6C5CE7]" />
                      </div>
                    </div>
                  )}
                  {settingsTab==="Data" && (
                    <div>
                      <h3 className="serif text-[22px]">Data & Privacy</h3>
                      <div className="mt-5 rounded-[16px] border border-[#6C5CE7]/20 bg-[#6C5CE7]/[0.06] p-4 text-[13px] leading-[1.5]">Everything runs locally — SheetJS parses XLSX/CSV in browser, no server upload. Storage: sheetsense_rows, sheetsense_original, sheetsense_log.</div>
                      <div className="mt-4 grid md:grid-cols-2 gap-3 text-[12px]">
                        <div className="card p-4"><div className="opacity-60">Storage used</div><div className="mt-1 font-medium">{(new Blob([JSON.stringify(rows)]).size/1024).toFixed(1)} KB rows • {localStorage.length} keys</div></div>
                        <div className="card p-4"><div className="opacity-60">Local keys</div><div className="mt-1 font-mono text-[11px] truncate">{Object.keys(localStorage).filter(k=>k.startsWith("sheetsense")).join(", ")||"—"}</div></div>
                      </div>
                      <div className="mt-5 flex flex-wrap gap-2">
                        <button onClick={()=>{ setRows([]); setOriginalRows(null); localStorage.removeItem("sheetsense_original"); showToast("Loaded file cleared"); }} className="h-9 px-4 rounded-full border border-black/10 bg-white text-[13px]">Clear loaded file</button>
                        <button onClick={()=>{ const data = {rows, originalRows, user, prefs, formatLog}; const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"}); const url=URL.createObjectURL(blob); const a=document.createElement("a"); a.href=url; a.download="sheetsense_data.json"; a.click(); URL.revokeObjectURL(url); showToast("Exported JSON"); }} className="h-9 px-4 rounded-full border border-black/10 bg-white text-[13px]">Export all as JSON</button>
                        <button onClick={()=>{ localStorage.removeItem("sheetsense_user"); setUser(null); showToast("Auth cleared"); }} className="h-9 px-4 rounded-full border border-black/10 bg-white text-[13px]">Clear auth data</button>
                      </div>
                      <div className="mt-8 rounded-[16px] border border-red-200 bg-red-50/60 p-4">
                        <div className="text-[13px] font-semibold text-red-700">Danger Zone</div>
                        <div className="mt-3 flex gap-2">
                          <button onClick={()=>{ localStorage.removeItem("sheetsense_rows"); localStorage.removeItem("sheetsense_original"); localStorage.removeItem("sheetsense_log"); localStorage.removeItem("sheetsense_formatted"); setRows([]); setOriginalRows(null); setFormatLog([]); setIsFormatted(false); showToast("Reset to defaults"); }} className="h-9 px-4 rounded-full bg-white border border-red-200 text-red-700 text-[13px]">Reset to defaults</button>
                          <button onClick={()=>{ localStorage.clear(); showToast("Session deleted — reloading"); setTimeout(()=>location.reload(),600); }} className="h-9 px-4 rounded-full bg-[#111113] text-white text-[13px]">Delete session</button>
                        </div>
                      </div>
                    </div>
                  )}
                  {settingsTab==="About" && (
                    <div>
                      <h3 className="serif text-[22px]">About SheetSense</h3>
                      <p className="mt-3 text-[13px] leading-[1.6] opacity-70 max-w-[60ch]">SheetSense v2 now includes real XLSX parsing via SheetJS (CDN 0.18.5) and AI Format & Clean. Parse CSV/JSON/XLSX/XLS locally, detect types, clean headers, normalize dates to ISO YYYY-MM-DD, split $2400 into Amount + Currency, trim spaces, remove empties, and render canvas charts. Glass, 24px radius, violet #6C5CE7 + peach #FF9A62.</p>
                      <div className="mt-6 flex gap-3 text-[12px]">
                        {["Privacy","Terms","Help"].map(l=><button key={l} onClick={()=>showToast(`${l} — coming soon`)} className="px-3 py-1.5 rounded-full border border-black/10 bg-white">{l}</button>)}
                      </div>
                    </div>
                  )}
                  <div className="mt-10 pt-4 border-t border-black/10 flex items-center justify-between text-[11px] opacity-50"><span>Version 2.1.0 • © SheetSense • XLSX + AI</span><span>Local-first • No tracking</span></div>
                </div>
              </div>
            )}
          </main>
        </div>

        {/* Mobile tab bar */}
        <div className="md:hidden fixed bottom-0 inset-x-0 z-40 border-t border-black/10 bg-[#FCFAF3]/90 glass">
          <div className="flex justify-around py-1.5">
            {[
              {k:"Upload", icon:"⤓"},
              {k:"Data", icon:"▤"},
              {k:"Analyze", icon:"◑"},
              {k:"Visuals", icon:"◿"},
              {k:"Currency", icon:"¤"},
              {k:"Settings", icon:"⚙"},
            ].map(t=>(
              <button key={t.k} onClick={()=>setView(t.k as any)} className={`flex flex-col items-center gap-0.5 px-3 py-1 rounded-[12px] ${view===t.k?"bg-[#111113] text-white":"opacity-70"}`}>
                <span className="text-[16px]">{t.icon}</span><span className="text-[10px] font-medium">{t.k}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Auth modal */}
        {authOpen && (
          <div className="fixed inset-0 z-[60] flex items-end md:items-center justify-center p-0 md:p-6">
            <div className="absolute inset-0 bg-[#111113]/40 backdrop-blur-[6px]" onClick={()=>setAuthOpen(false)} />
            <div className="relative w-full md:w-[440px] card p-6 md:rounded-[24px] rounded-t-[24px] bg-[#FCFAF3] shadow-2xl">
              <div className="flex items-center justify-between">
                <div className="flex gap-1 p-1 rounded-full bg-black/5 border border-black/10">
                  <button onClick={()=>setAuthMode("login")} className={`h-7 px-3 rounded-full text-[12px] font-medium ${authMode==="login"?"bg-[#111113] text-white":"opacity-70"}`}>Log in</button>
                  <button onClick={()=>setAuthMode("signup")} className={`h-7 px-3 rounded-full text-[12px] font-medium ${authMode==="signup"?"bg-[#111113] text-white":"opacity-70"}`}>Sign up</button>
                </div>
                <button onClick={()=>setAuthOpen(false)} className="w-8 h-8 rounded-full bg-white border border-black/10 grid place-items-center">✕</button>
              </div>
              <h3 className="serif text-[24px] mt-5">{authMode==="login" ? "Welcome back" : "Create account"}</h3>
              <div className="mt-4 space-y-3">
                {authMode==="signup" && <input value={authForm.name} onChange={e=>setAuthForm({...authForm, name:e.target.value})} placeholder="Name" className="w-full h-11 rounded-[12px] border border-black/10 bg-white px-3 text-[13px]" />}
                <input value={authForm.email} onChange={e=>setAuthForm({...authForm, email:e.target.value})} placeholder="Email" className="w-full h-11 rounded-[12px] border border-black/10 bg-white px-3 text-[13px]" />
                <input value={authForm.password} onChange={e=>setAuthForm({...authForm, password:e.target.value})} type="password" placeholder="Password (min 6 chars)" className="w-full h-11 rounded-[12px] border border-black/10 bg-white px-3 text-[13px]" />
                {authForm.email && !/\S+@\S+\.\S+/.test(authForm.email) && <div className="text-[11px] text-red-600">Enter a valid email</div>}
                {authForm.password && authForm.password.length<6 && <div className="text-[11px] text-red-600">Password must be at least 6 characters</div>}
              </div>
              <button onClick={()=>{
                if(!/\S+@\S+\.\S+/.test(authForm.email)) return showToast("Invalid email");
                if(authForm.password.length<6) return showToast("Password too short");
                if(authMode==="signup" && !authForm.name) return showToast("Name required");
                setUser({name: authForm.name || authForm.email.split("@")[0], email: authForm.email, type:"Free"});
                setAuthOpen(false); showToast("Logged in");
              }} className="mt-4 w-full h-11 rounded-full bg-[#111113] text-white text-[13px] font-semibold">Continue</button>
              <div className="mt-3 grid grid-cols-2 gap-2">
                <button onClick={()=>showToast("Google auth mock — use email")} className="h-10 rounded-full bg-white border border-black/10 text-[12px] font-medium">Continue with Google</button>
                <button onClick={()=>{ setUser({name:"Guest", email:"guest@sheetsense.local", type:"Guest"}); setAuthOpen(false); showToast("Guest mode enabled"); }} className="h-10 rounded-full bg-[#FF9A62]/15 border border-[#FF9A62]/30 text-[12px] font-medium">Continue as Guest</button>
              </div>
              <button onClick={()=>showToast("Password reset link sent (mock)")} className="mt-3 text-[11px] underline opacity-60">Forgot password?</button>
            </div>
          </div>
        )}

        {/* Toast */}
        {toast && (
          <div className="fixed bottom-[90px] md:bottom-6 right-4 z-[70] card px-4 py-2.5 text-[13px] font-medium bg-[#111113] text-white border-black/20 shadow-xl animate-[in_0.2s_ease]">
            {toast.msg}
          </div>
        )}
      </div>
    </>
  );
}
