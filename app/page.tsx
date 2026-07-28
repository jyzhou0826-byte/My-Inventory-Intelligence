"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import seedData from "./data/dashboard.json";
import { analyzeUploads, type UploadSelection } from "./analysis";
import { exportInventoryPptx } from "./pptx-export";

let data: any = seedData;

type Tab = "Executive" | "ALL" | "FORGN" | "LP3" | "KD" | "QoQ" | "Plan";
type Item = (typeof seedData.latestItems)[number] & { partName?: string };

const tabs: { key: Tab; label: string }[] = [
  { key: "Executive", label: "Executive" }, { key: "ALL", label: "All Inventory" },
  { key: "FORGN", label: "FORGN" }, { key: "LP3", label: "LP3" },
  { key: "KD", label: "KD" }, { key: "QoQ", label: "QoQ Analysis" }, { key: "Plan", label: "Plan Stability" },
];
const riskTone: Record<string, string> = { 缺料: "red", 缺料風險: "red", 安全: "green", 積壓: "yellow", 呆滯: "orange", 無需求庫存: "blue", "無需求／無庫存": "gray" };
const fmtWan = (n: number) => `${(n / 10000).toLocaleString("zh-TW", { maximumFractionDigits: 1 })}`;
const fmtPct = (n: number) => `${(n * 100).toFixed(1)}%`;
const safeDelta = (current: number, previous: number) => previous === 0 ? null : (current - previous) / previous;
const grade = (n: number) => n >= .8 ? "優秀" : n >= .6 ? "良好" : n >= .4 ? "警示" : "危險";

function Icon({ name }: { name: "grid" | "box" | "trend" | "plan" | "settings" | "upload" }) {
  const icons = { grid: "▦", box: "◇", trend: "⌁", plan: "≋", settings: "⚙", upload: "↑" };
  return <span className="icon" aria-hidden="true">{icons[name]}</span>;
}

function QuarterBars({ values, labels }: { values: number[]; labels: string[] }) {
  const max = Math.max(...values, 1);
  return <div className="quarter-bars" aria-label="各季度總庫存金額與 QoQ 變化">{values.map((v, i) => {
    const delta = i ? safeDelta(v, values[i - 1]) : null;
    return <div key={labels[i]}><span>{delta === null ? "基期" : `${delta >= 0 ? "↑" : "↓"} ${fmtPct(Math.abs(delta))}`}</span><i style={{ height: `${Math.max(16, v / max * 100)}%` }} /><b>{labels[i]}</b></div>;
  })}</div>;
}

function Donut({ value, label, tone = "green" }: { value: number; label: string; tone?: string }) {
  return <div className={`donut ${tone}`} style={{ "--p": `${Math.max(0, Math.min(100, value * 100)) * 3.6}deg` } as React.CSSProperties}>
    <div><strong>{fmtPct(value)}</strong><span>{label}</span></div>
  </div>;
}

function DeltaBadge({ delta, increaseGood = false }: { delta: number | null; increaseGood?: boolean }) {
  if (delta === null) return <b className="delta-neutral">基準</b>;
  const good = increaseGood ? delta >= 0 : delta < 0;
  return <b className={Math.abs(delta)<.001?"delta-neutral":good?"delta-good":"delta-bad"}>{delta>=0?"+":""}{fmtPct(delta)}</b>;
}

const statusOrder = ["缺料", "缺料風險", "安全", "積壓", "呆滯", "無需求庫存", "無需求／無庫存"] as const;
const statusColors: Record<string,string> = { 缺料: "#DC2626", 缺料風險: "#F87171", 安全: "#16A34A", 積壓: "#F4B400", 呆滯: "#7C3AED", 無需求庫存: "#2563EB", "無需求／無庫存": "#94A3B8" };

function HealthExplanation() {
  return <details className="health-explanation"><summary>庫存健康度與 LP3 判定說明</summary><div><p>LP3 先檢查「生產需求與可用庫存」，再判斷庫存天數；不得以庫存水準為 0 單獨判定缺料。備註為「子件」的 LP3 件號排除分析。</p><b>缺料 KPI 僅統計：生產需求 &gt; 0，且可用庫存不足或低於缺料門檻的件號。</b><p>需求為 0 且有庫存者依庫存天數分類：超過 180 天為呆滯、超過 90 天為積壓，其餘為無需求庫存。無需求狀態不納入健康度分母。</p></div></details>;
}

function StatusDonut({ rows, group }: { rows: any[]; group: string }) {
  const stats=statusOrder.map(status=>{const matched=rows.filter(r=>r.risk===status);return{status,count:matched.length,amount:matched.reduce((s,r)=>s+r.inventory,0)}});
  const total=Math.max(rows.length,1); let cursor=0;
  const stops=stats.map(x=>{const start=cursor;cursor+=x.count/total*360;return `${statusColors[x.status]} ${start}deg ${cursor}deg`}).join(",");
  return <section className="card status-card"><div className="section-head"><div><span className="section-index">01</span><strong>{group} 庫存健康結構</strong></div><span>四類互斥 · 合計 100%</span></div><div className="status-layout"><div className="status-donut" style={{background:`conic-gradient(${stops || "#E2E8F0 0deg 360deg"})`}} aria-label={`${group} 健康狀態甜甜圈圖`}><div><b>{group}</b><strong>{rows.length}</strong><span>總件號數</span></div></div><div className="status-summary">{stats.map(x=><article key={x.status}><i style={{background:statusColors[x.status]}}/><span>{x.status}<small>{fmtWan(x.amount)} 萬元</small></span><b>{x.count} 件<em>{fmtPct(x.count/total)}</em></b></article>)}</div></div></section>;
}

function VarianceTrend() {
  const canvasRef=useRef<HTMLCanvasElement>(null); const categories=["FORGN","LP3","KD"],quarters=data.quarters;
  const series=categories.map(group=>({label:group,values:quarters.map((q:any,i:number)=>{const previous=i?quarters[i-1].groupTotals?.[group]??0:null,current=q.groupTotals?.[group]??0;return{quarter:q.quarter,previous,current,delta:previous===null?null:current-previous,rate:previous===null||previous===0?null:(current-previous)/previous}})}));
  useEffect(()=>{const canvas=canvasRef.current;if(!canvas)return;const draw=()=>{const rect=canvas.getBoundingClientRect(),ratio=window.devicePixelRatio||1,w=Math.max(rect.width,640),h=400;canvas.width=w*ratio;canvas.height=h*ratio;const c=canvas.getContext("2d");if(!c)return;c.scale(ratio,ratio);c.clearRect(0,0,w,h);const rates=series.flatMap(s=>s.values.map(v=>v.rate).filter((v):v is number=>v!==null));const bound=Math.max(.1,...rates.map(v=>Math.abs(v)))*1.2,left=72,right=32,top=44,bottom=62,cw=w-left-right,ch=h-top-bottom,y=(v:number)=>top+(bound-v)/(bound*2)*ch,x=(i:number)=>left+(quarters.length===1?cw/2:i*cw/(quarters.length-1));c.font='14px "Microsoft JhengHei", sans-serif';for(let step=-2;step<=2;step++){const value=bound*step/2,py=y(value);c.strokeStyle=step===0?"#94A3B8":"#E2E8F0";c.lineWidth=step===0?1.5:1;c.beginPath();c.moveTo(left,py);c.lineTo(w-right,py);c.stroke();c.fillStyle="#64748B";c.textAlign="right";c.fillText(`${(value*100).toFixed(0)}%`,left-12,py+5)}quarters.forEach((q:any,i:number)=>{const px=x(i);c.strokeStyle="#EEF2F7";c.beginPath();c.moveTo(px,top);c.lineTo(px,h-bottom);c.stroke();c.fillStyle="#334155";c.textAlign="center";c.fillText(q.quarter,px,h-24)});const colors=["#2563EB","#16A34A","#EA580C"];series.forEach((s,si)=>{c.strokeStyle=colors[si];c.fillStyle=colors[si];c.lineWidth=si===0?4:3;c.beginPath();let started=false;s.values.forEach((v,i)=>{if(v.rate===null)return;const px=x(i),py=y(v.rate);if(!started)c.moveTo(px,py);else c.lineTo(px,py);started=true});c.stroke();s.values.forEach((v,i)=>{if(v.rate===null)return;const px=x(i),py=y(v.rate);c.beginPath();c.arc(px,py,5,0,Math.PI*2);c.fill();c.font='700 13px "Microsoft JhengHei", sans-serif';c.textAlign="center";c.fillText(`${(v.rate*100).toFixed(1)}%`,px,py-12)});});};draw();const observer=new ResizeObserver(draw);observer.observe(canvas);return()=>observer.disconnect()},[data]);
  const details=series.flatMap(s=>s.values.slice(1).map(v=>({...v,group:s.label})));
  return <section className="card variance-card"><div className="section-head"><div><span className="section-index">02</span><strong>全般部品差異率趨勢曲線圖</strong><small>VARIANCE ANALYSIS COMPARISON</small></div><span>橫軸：所有上傳季度 · 0% 基準線</span></div>{quarters.length>1?<><div className="variance-legend category-lines">{series.map(s=><span key={s.label} className={s.label.toLowerCase()}><i/>{s.label}</span>)}</div><canvas ref={canvasRef}/><div className="variance-details">{details.map(v=><span key={`${v.quarter}-${v.group}`} title={`前季 ${fmtWan(v.previous!)} 萬；本季 ${fmtWan(v.current)} 萬；差異 ${fmtWan(v.delta!)} 萬`}>{v.quarter} · {v.group}<b>{v.rate===null?"—":fmtPct(v.rate)}</b></span>)}</div></>:<div className="empty">選擇兩個以上季度即可顯示相對前季的差異率曲線。</div>}</section>;
}

function RiskTable({ title, rows, tone }: { title: string; rows: Item[]; tone: string }) {
  return <section className="card risk-card">
    <div className="section-head"><div><span className={`dot ${tone}`} /> <strong>{title}</strong></div><span>依庫存金額</span></div>
    {rows.length ? <div className="table-wrap"><table><thead><tr><th>件號／件名</th><th>群組</th><th>庫存金額</th><th>水準</th><th>建議</th></tr></thead>
      <tbody>{rows.map((r, i) => <tr key={r.part}><td><div className="part-cell"><b className="rank">{i + 1}</b><span><strong>{r.part}</strong><small className="source-missing">{r.partName || "來源未提供"}</small></span></div></td><td>{r.rawGroup}</td><td>{fmtWan(r.inventory)} 萬</td><td>{r.level.toFixed(2)}</td><td><span className={`pill ${tone}`}>{r.action}</span></td></tr>)}</tbody></table></div>
    : <div className="empty">本期沒有符合條件的項目</div>}
  </section>;
}

function Setup({ onRun }: { onRun: (result: any) => void }) {
  const [model, setModel] = useState("RV");
  const [years, setYears] = useState([2026]);
  const [quarters, setQuarters] = useState(["Q1", "Q2", "Q3"]);
  const [files, setFiles] = useState<Record<string, string>>({ inventory: "", previous: "", current: "" });
  const [uploads, setUploads] = useState<UploadSelection>({ inventory: [], previous: null, current: null });
  const [error, setError] = useState("");
  const [running, setRunning] = useState(false);
  const invalid = false;
  const toggle = <T,>(arr: T[], value: T, setter: (v: T[]) => void) => setter(arr.includes(value) ? arr.filter(x => x !== value) : [...arr, value]);
  const choose = (key: string, list: FileList | null) => { if (!list?.length) return; const picked=Array.from(list); setFiles(x => ({ ...x, [key]: picked.map(f => f.name).join(" · ") })); setUploads(x=>({...x,[key]:key==="inventory"?picked:picked[0]})); setError(""); };
  const execute = async () => { try { setRunning(true); setError(""); const result=await analyzeUploads(model,uploads,{years,quarters}); onRun(result); } catch (reason) { setError(reason instanceof Error?reason.message:"分析失敗，請確認檔案格式。"); } finally { setRunning(false); } };
  return <main className="setup-page">
    <div className="setup-shell">
      <header className="setup-brand"><div className="brand-mark">MY</div><div><b>My Inventory Intelligence</b><span>庫存決策分析平台</span></div><span className="verified">● 正式分析模式</span></header>
      <div className="setup-grid">
        <section className="setup-copy"><span className="eyebrow">INVENTORY SPECIALIST</span><h1>把庫存資料，轉成<br /><em>可執行的決策。</em></h1><p>整合 Power BI 匯出資料與生產計畫，自動辨識庫存風險、策略備料與季度變化。</p>
          <div className="feature-row"><span><b>CSV</b>CP950 / Big5</span><span><b>XLSX</b>Power BI 匯出</span><span><b>13</b>必要欄位驗證</span></div>
        </section>
        <section className="setup-card">
          <div className="step-title"><b>01</b><div><h2>設定分析範圍</h2><p>確認車型、年度與季度</p></div></div>
          <label>車型<input value={model} onChange={e => setModel(e.target.value.toUpperCase())} placeholder="例如：RV" /></label>
          <div className="label">分析年度（可多選）</div><div className="chips">{[2025, 2026, 2027].map(y => <button className={years.includes(y) ? "active" : ""} onClick={() => toggle(years, y, setYears)} key={y}>{y}</button>)}</div>
          <div className="label">分析季度（可多選）</div><div className="chips">{["Q1", "Q2", "Q3", "Q4"].map(q => <button className={quarters.includes(q) ? "active" : ""} onClick={() => toggle(quarters, q, setQuarters)} key={q}>{q}</button>)}</div>
          <div className="divider" />
          <div className="step-title compact"><b>02</b><div><h2>資料來源</h2><p>CSV 或 XLSX</p></div></div>
          {[["inventory", "Power BI 庫存資料", true], ["previous", "前回生產計畫", false], ["current", "本回生產計畫", false]].map(([key, label, multi]) => <label className="file" key={String(key)}><input type="file" accept=".csv,.xlsx" multiple={Boolean(multi)} onChange={e => choose(String(key), e.target.files)} /><span className="file-icon"><Icon name="upload" /></span><span><b>{label}</b><small>{files[String(key)] || "選擇 CSV 或 XLSX"}</small></span><i>瀏覽</i></label>)}
          {invalid && <div className="alert">【車型資料異常】輸入車型與已載入資料 RV 不一致，請確認。</div>}
          {error && <div className="alert">{error}</div>}
          <button className="run" disabled={running || !model || !years.length || !quarters.length || !uploads.inventory.length || !uploads.previous || !uploads.current} onClick={execute}>{running?"分析中…":"執行分析"} <span>→</span></button>
        </section>
      </div>
    </div>
  </main>;
}

function Executive({ setTab }: { setTab: (t: Tab) => void }) {
  const q = data.quarters; const latest = data.latest; const totals = latest.totals;
  const invQ = q.map(x => x.totals.inventory); const coverage = totals.inventory / totals.demand;
  const previousQ = q.length > 1 ? q[q.length-2] : null;
  const qoq = previousQ?.totals.inventory ? (totals.inventory - previousQ.totals.inventory) / previousQ.totals.inventory : null;
  const previousHealthBase = previousQ ? previousQ.records - (previousQ.counts.無需求庫存 ?? 0) - (previousQ.counts["無需求／無庫存"] ?? 0) : 0;
  const previousSafeRate = previousHealthBase ? (previousQ.counts.安全 ?? 0) / previousHealthBase : null;
  const healthDelta = previousSafeRate === null ? null : latest.safeRate - previousSafeRate;
  const coverageStatus = coverage < .8 ? "缺料風險" : coverage <= 1.2 ? "正常" : coverage > 2 ? "呆滯風險" : coverage > 1.5 ? "高庫存風險" : "偏高";
  return <>
    <div className="hero-row"><div><span className="eyebrow">EXECUTIVE OVERVIEW · {data.meta.latestQuarter}</span><h1>庫存決策總覽</h1><p>以 {latest.items} 筆 {data.meta.latestQuarter} 實際明細進行即時風險診斷</p></div><div className="data-status"><span>●</span>資料已驗證<small>{data.meta.model} · {data.meta.sourceFiles.length} 份庫存檔</small></div></div>
    <div className="kpi-grid">
      <article className="kpi primary inventory-kpi"><div className="kpi-top"><span>總庫存金額</span><b className={qoq === null ? "delta-neutral" : qoq >= 0 ? "delta-bad" : "delta-good"}>{qoq === null ? "單季基準" : `${qoq >= 0 ? "+" : ""}${fmtPct(qoq)}`}</b></div><strong>{fmtWan(totals.inventory)}<small> 萬元</small></strong><QuarterBars values={invQ} labels={q.map(x=>x.quarter)} /><footer>前期：{previousQ ? `${fmtWan(previousQ.totals.inventory)} 萬元` : "無前期"} · 庫存增加視為不利</footer></article>
      <article className="kpi"><div className="kpi-top"><span>庫存健康度</span><b className={healthDelta===null?"delta-neutral":healthDelta>=0?"delta-good":"delta-bad"}>{healthDelta===null?grade(latest.safeRate):`${healthDelta>=0?"+":""}${(healthDelta*100).toFixed(1)}pp`}</b></div><strong>{fmtPct(latest.safeRate)}</strong><div className="kpi-context"><b>{latest.counts.安全}</b><span>安全 / {latest.healthBaseItems ?? latest.items} 件</span></div><footer>前期：{previousSafeRate===null?"無前期":fmtPct(previousSafeRate)} · 無需求狀態不納入分母</footer></article>
      <article className="kpi"><div className="kpi-top"><span>高風險占比</span><b className="bad">極高</b></div><strong>{fmtPct(latest.highRiskRatio)}</strong><div className="progress"><i style={{ width: fmtPct(latest.highRiskRatio) }} /></div><footer>{fmtWan(latest.highRiskInventory)} 萬元需優先處置</footer></article>
      <article className="kpi"><div className="kpi-top"><span>需求覆蓋率 <abbr title="總庫存金額 ÷ 生產需求金額；1.0x 代表庫存約可覆蓋一次目前需求。">i</abbr></span><b className={coverageStatus==="正常"?"risk-safe":coverageStatus==="偏高"?"risk-warning":"risk-danger"}>{coverageStatus}</b></div><strong>{Number.isFinite(coverage)?coverage.toFixed(2):"—"}<small> x</small></strong><div className="coverage-scale"><i style={{ left: `${Math.min(95, coverage / 3 * 100)}%` }} /></div><div className="coverage-note">1.0x 代表庫存約可覆蓋一次目前生產需求。</div><footer>總庫存 ÷ 生產需求 · 標準 0.8–1.2</footer></article>
    </div>
    <HealthExplanation />
    <div className="dashboard-grid">
      <section className="card health-card"><div className="section-head"><div><span className="section-index">01</span><strong>庫存健康分析</strong></div><button onClick={() => setTab("ALL")}>查看明細 →</button></div>
        <div className="health-layout"><div className="risk-stack">{statusOrder.map(r => <div key={r} style={{ flex: Number(latest.counts[r] ?? 0) }} className={riskTone[r]}><b>{latest.counts[r] ?? 0}</b><span>{r}</span></div>)}</div>
        <div className="legend">{Object.entries(latest.counts).map(([r,n]) => <span key={r}><i className={riskTone[r]} />{r}<b>{n} 件</b></span>)}</div></div>
        <div className="group-row">{["FORGN","LP3","KD"].map(g => { const x = (data.groups as Record<string, typeof data.groups.FORGN>)[g]; return <button key={g} onClick={() => setTab(g as Tab)}><span>{g}</span>{x ? <><b>{fmtWan(x.inventory)} 萬</b><small>平均水準 {x.avgLevel.toFixed(2)} {g === "LP3" ? "日" : "月"}</small></> : <><b>—</b><small>本期無資料</small></>}</button>})}</div>
      </section>
      <section className="card structure-card"><div className="section-head"><div><span className="section-index">02</span><strong>總庫存金額分析</strong></div><span>{data.meta.latestQuarter} · 萬元</span></div>
        <div className="structure-total"><span>總庫存金額</span><strong>{fmtWan(totals.inventory)} <small>萬元</small></strong></div>
        <QuarterBars values={invQ} labels={q.map(x=>x.quarter)} />
        <div className="ai-note"><b>管理判讀</b><p>僅以總庫存金額追蹤各季度變化，並搭配高風險金額、需求趨勢與計畫異動判斷管理優先順序。</p></div>
      </section>
    </div>
    <section className="card executive-qoq"><div className="section-head"><div><span className="section-index">03</span><strong>QoQ 季度變化</strong></div><button onClick={() => setTab("QoQ")}>完整分析 →</button></div>
      <div className="executive-qoq-grid"><div className="exec-trend">{q.map((x,i)=>{const d=i?safeDelta(x.totals.inventory,q[i-1].totals.inventory):null;return <div key={x.quarter}><span>{fmtWan(x.totals.inventory)} 萬</span><i style={{height:`${x.totals.inventory/Math.max(...invQ,1)*100}%`}}/><b>{x.quarter}</b>{i>0&&<em className={d===null?"":""+(d>=0?"up":"down")}>{d===null?"—":`${d>=0?"↑":"↓"} ${fmtPct(Math.abs(d))}`}</em>}</div>})}</div>
      <div className="qoq-insight"><span>{previousQ ? `${previousQ.quarter} → ${data.meta.latestQuarter}` : data.meta.latestQuarter}</span><strong className={qoq === null ? "" : qoq >= 0 ? "up" : "down"}>{qoq === null ? "分析基準季" : `${qoq >= 0 ? "↑" : "↓"} ${fmtPct(Math.abs(qoq))}`}</strong><p>{previousQ ? `總庫存${qoq! >= 0 ? "增加" : "減少"} ${fmtWan(Math.abs(totals.inventory-previousQ.totals.inventory))} 萬元；請依風險清單安排處置。` : "目前僅選擇一個季度；已完成當季庫存與風險分析，選擇兩季以上即可計算 QoQ。"}</p></div></div>
    </section>
    <section className="section-block"><div className="section-title"><div><span>04</span><h2>關鍵風險分析</h2></div><p>依 {data.meta.latestQuarter} 庫存金額排序 · 件名依來源資料顯示 · 單位：萬元</p></div><div className="risk-grid"><RiskTable title="TOP 3 缺料風險" rows={data.tops.缺料} tone="red" /><RiskTable title="TOP 3 積壓風險" rows={data.tops.積壓} tone="yellow" /><RiskTable title="TOP 3 呆滯風險" rows={data.tops.呆滯} tone="orange" /></div></section>
  </>;
}

function KDProjectPage() {
  const projects=[
    {key:"外開把手",type:"策略備料",question:"是否需要持續備料？目前備料是否合理？",accent:"blue"},
    {key:"370轉置差異件",type:"庫存去化",question:"剩餘庫存是否有效消化？是否存在呆滯風險？",accent:"orange"},
  ];
  return <><div className="hero-row"><div><span className="eyebrow">{data.meta.latestQuarter} · KD PROJECTS</span><h1>KD 專案庫存分析</h1><p>僅分析外開把手與 370 轉置差異件，不與一般 KD 庫存合併。</p></div><div className="data-status"><span>●</span>2 個指定專案<small>依專案定位分開管理</small></div></div>
    <section className="kd-project-grid">{projects.map(project=>{const p=data.kdProjects?.[project.key];const top=[...(p?.items??[])].sort((a,b)=>b.inventory-a.inventory).slice(0,10);return <article className={`card kd-project ${project.accent}`} key={project.key}><header><div><span>KD PROJECT</span><h2>{project.key}</h2></div><b>{project.type}</b></header>{p?.count?<><div className="kd-project-kpis"><div><span>總庫存金額</span><strong>{fmtWan(p.inventory)} 萬</strong></div><div><span>生產需求</span><strong>{fmtWan(p.demand)} 萬</strong></div><div><span>件號數</span><strong>{p.count} 件</strong></div></div><p>{project.question}</p><div className="kd-project-status">{statusOrder.filter(s=>p.counts?.[s]).map(s=><span key={s}><i className={riskTone[s]}/>{s} {p.counts[s]}</span>)}</div><div className="kd-project-items"><h3>高金額件號</h3>{top.map((item,i)=><div key={item.part}><b>{i+1}</b><span><strong>{item.part}</strong><small>{item.partName||"來源未提供件名"}</small></span><em>{fmtWan(item.inventory)} 萬</em><i className={`pill ${riskTone[item.risk]}`}>{item.risk}</i></div>)}</div></>:<div className="empty">本期無此專案資料</div>}</article>})}</section>
  </>;
}

function InventoryPage({ group }: { group: "ALL" | "FORGN" | "LP3" | "KD" }) {
  if (group === "KD") return <KDProjectPage />;
  const rows = group === "ALL" ? data.latestItems : data.latestItems.filter(x => x.group === group);
  const inventory = rows.reduce((s,x)=>s+x.inventory,0), demand = rows.reduce((s,x)=>s+x.demand,0), safe = rows.filter(x=>x.risk==="安全").length;
  const previousQuarter=data.quarters.length>1?data.quarters[data.quarters.length-2]:null;
  const previous=previousQuarter?(group==="ALL"?{inventory:previousQuarter.totals.inventory,demand:previousQuarter.totals.demand,count:previousQuarter.records,safe:previousQuarter.counts.安全}:previousQuarter.groupMetrics?.[group]):null;
  const currentCoverage=demand?inventory/demand:null, previousCoverage=previous?.demand?previous.inventory/previous.demand:null;
  const healthBase=rows.filter(x=>x.risk!=="無需求庫存"&&x.risk!=="無需求／無庫存").length;
  const currentHealth=healthBase?safe/healthBase:0, previousHealth=previous?.count?previous.safe/previous.count:null;
  const sorted = [...rows].sort((a,b)=>b.inventory-a.inventory).slice(0,20);
  return <><div className="hero-row"><div><span className="eyebrow">{data.meta.latestQuarter} · {group}</span><h1>{group === "ALL" ? "全般庫存分析" : `${group} 庫存分析`}</h1><p>{group === "LP3" ? "庫存水準單位：日" : "庫存水準單位：月"}</p></div><div className="data-status"><span>●</span>{rows.length} 筆明細<small>依 Material Group 自動分類</small></div></div>
    {!rows.length ? <section className="card no-data"><b>{group}</b><h2>本期無 {group} 分類資料</h2><p>{data.meta.latestQuarter} 原始資料的 Material Group 未包含此分類，因此不產生推估數據。</p></section> : <>
      <div className="kpi-grid compact-kpis"><article className="kpi"><span>總庫存金額</span><strong>{fmtWan(inventory)}<small> 萬元</small></strong><footer className="category-kpi-footer"><span>上季：{previous?`${fmtWan(previous.inventory)} 萬元`:"—"}</span><DeltaBadge delta={previous?safeDelta(inventory,previous.inventory):null}/></footer></article><article className="kpi"><span>生產需求金額</span><strong>{fmtWan(demand)}<small> 萬元</small></strong><footer className="category-kpi-footer"><span>上季：{previous?`${fmtWan(previous.demand)} 萬元`:"—"}</span><DeltaBadge delta={previous?safeDelta(demand,previous.demand):null}/></footer></article><article className="kpi"><span>需求覆蓋率 <abbr title="總庫存金額 ÷ 生產需求金額">i</abbr></span><strong>{currentCoverage===null?"—":currentCoverage.toFixed(2)}<small> x</small></strong><footer className="category-kpi-footer"><span>上季：{previousCoverage===null?"—":`${previousCoverage.toFixed(2)} x`}</span><DeltaBadge delta={currentCoverage!==null&&previousCoverage!==null?safeDelta(currentCoverage,previousCoverage):null}/></footer></article><article className="kpi"><span>庫存健康度</span><strong>{fmtPct(currentHealth)}<small> · {grade(currentHealth)}</small></strong><footer className="category-kpi-footer"><span>上季：{previousHealth===null?"—":fmtPct(previousHealth)}</span>{previousHealth===null?<b className="delta-neutral">基準</b>:<b className={currentHealth>=previousHealth?"delta-good":"delta-bad"}>{currentHealth-previousHealth>=0?"+":""}{((currentHealth-previousHealth)*100).toFixed(1)}pp</b>}</footer></article></div>
      {(group === "FORGN" || group === "LP3") && <StatusDonut rows={rows} group={group} />}
      {group === "ALL" && <section className="card category-health"><div className="section-head"><div><span className="section-index">01</span><strong>各分類庫存健康分析</strong></div><span>{data.meta.year} {data.meta.latestQuarter} · ALL</span></div><div className="category-health-grid">{["FORGN","LP3","KD"].map(g=>{const rs=data.latestItems.filter(x=>x.group===g);const counts={缺料:rs.filter(x=>x.risk==="缺料").length,安全:rs.filter(x=>x.risk==="安全").length,積壓:rs.filter(x=>x.risk==="積壓").length,呆滯:rs.filter(x=>x.risk==="呆滯").length};const safeRate=rs.length?counts.安全/rs.length:0;return <article key={g}><header><div><b>{g}</b><span>{g==="LP3"?"國產件 / TAX FREE":g==="FORGN"?"海外採購件":"專案物料"}</span></div><strong>{rs.length?fmtPct(safeRate):"—"}<small>{rs.length?grade(safeRate):"無資料"}</small></strong></header>{rs.length?<><div className="category-risk-bar">{Object.entries(counts).filter(([,n])=>n).map(([r,n])=><i key={r} className={riskTone[r]} style={{flex:n}} title={`${r} ${n} 件`}/>)}</div><footer>{Object.entries(counts).map(([r,n])=><span key={r}><i className={riskTone[r]}/>{r} {n}</span>)}</footer></>:<div className="category-empty">{data.meta.latestQuarter} 無 {g} 資料</div>}</article>})}</div></section>}
      {group === "ALL" && <VarianceTrend />}
      <section className="card inventory-table"><div className="section-head"><div><span className="section-index">{group === "ALL" ? "03" : "02"}</span><strong>高金額庫存明細</strong></div><span>顯示前 20 筆</span></div><div className="table-wrap"><table><thead><tr><th>件號</th><th>件名</th><th>Material Group</th><th>原因</th><th>庫存金額</th><th>需求金額</th><th>庫存水準</th><th>風險</th><th>建議措施</th></tr></thead><tbody>{sorted.map(r=><tr key={r.part}><td><b>{r.part}</b></td><td><span className="source-missing">{r.partName || "來源未提供"}</span></td><td>{r.rawGroup}</td><td>{r.reason}</td><td>{fmtWan(r.inventory)} 萬</td><td>{fmtWan(r.demand)} 萬</td><td>{r.level.toFixed(2)}</td><td><span className={`pill ${riskTone[r.risk]}`}>{r.risk}</span></td><td>{r.action}</td></tr>)}</tbody></table></div></section>
    </>}
  </>;
}

function QoQPage() {
  const max = Math.max(...data.quarters.map(q=>q.totals.inventory),1);
  const demandMax=Math.max(...data.quarters.map(q=>q.totals.demand),1);
  const points=data.quarters.map((q,i)=>`${data.quarters.length===1?50:5+i*90/(data.quarters.length-1)},${92-q.totals.demand/demandMax*78}`).join(" ");
  return <><div className="hero-row"><div><span className="eyebrow">QUARTER OVER QUARTER</span><h1>QoQ 季度變化</h1><p>{data.meta.quarters.join(" → ")} 的總庫存金額與生產需求趨勢</p></div></div>
    <section className="card quarter-structure-card"><div className="section-head"><div><span className="section-index">01</span><strong>季度總庫存金額</strong></div><span>單位：萬元</span></div><div className="quarter-structure total-only">{data.quarters.map(q=><div key={q.quarter}><b>{fmtWan(q.totals.inventory)}</b><div style={{height:`${q.totals.inventory/max*100}%`}}><i/></div><strong>{q.quarter}</strong></div>)}</div></section>
    <section className="card demand-trend-card"><div className="section-head"><div><span className="section-index">02</span><strong>季度生產需求趨勢</strong></div><span>需求指標，不屬於庫存組成</span></div><svg viewBox="0 0 100 110" preserveAspectRatio="none" aria-label="季度生產需求折線圖"><line x1="5" y1="92" x2="95" y2="92"/><polyline points={points}/>{data.quarters.map((q,i)=>{const x=data.quarters.length===1?50:5+i*90/(data.quarters.length-1),y=92-q.totals.demand/demandMax*78;return <g key={q.quarter}><circle cx={x} cy={y} r="1.6"/><text x={x} y={Math.max(8,y-5)}>{fmtWan(q.totals.demand)}</text><text x={x} y="105">{q.quarter}</text></g>})}</svg></section>
    <div className="qoq-grid">{data.quarters.slice(1).map((q,i)=>{const p=data.quarters[i]; const metrics:[[string,number,number],[string,number,number]]=[["總庫存",p.totals.inventory,q.totals.inventory],["生產需求",p.totals.demand,q.totals.demand]];return <section className="card" key={q.quarter}><div className="section-head"><strong>{p.quarter} → {q.quarter}</strong><span>增減分析</span></div>{metrics.map(([name,a,b])=>{const d=safeDelta(b,a);return <div className="variance-row" key={name}><span>{name}</span><b>{fmtWan(b-a)} 萬</b><em className={d===null?"":d>=0?"up":"down"}>{d===null?"—":`${d>=0?"↑":"↓"} ${fmtPct(Math.abs(d))}`}</em></div>})}</section>})}</div>
  </>;
}

function PlanPage() {
  const comparable=data.plans.months.filter(x=>x.variance!==null);
  const maxVariance=Math.max(...comparable.map(x=>Math.abs(x.variance||0)),1);
  const maxAdjustment=Math.max(...comparable.map(x=>Math.abs(x.materialAdjustment||0)),1);
  const totalAdjustment=comparable.reduce((s,x)=>s+(x.materialAdjustment||0),0);
  const totalImpact=comparable.reduce((s,x)=>s+(x.inventoryImpact||0),0);
  return <><div className="hero-row"><div><span className="eyebrow">PLAN STABILITY · {data.meta.model}計</span><h1>生產計畫異動分析</h1><p>前回計畫與本回計畫重疊月份差異</p></div><div className="data-status"><span>●</span>車型一致<small>兩份資料皆為 {data.meta.model}計</small></div></div>
    <div className="plan-summary"><article className="card"><span>可比較月份</span><strong>{comparable.length}<small> 個月</small></strong></article><article className="card"><span>計畫淨增減</span><strong className={comparable.reduce((s,x)=>s+(x.variance||0),0)===0?"": "risk"}>{comparable.reduce((s,x)=>s+(x.variance||0),0)}<small> 台</small></strong></article><article className="card"><span>備料調整金額</span><strong className={totalAdjustment===0?"":"risk"}>{fmtWan(totalAdjustment)}<small> 萬元</small></strong></article><article className="card"><span>庫存影響金額</span><strong className={totalImpact===0?"":"risk"}>{fmtWan(totalImpact)}<small> 萬元</small></strong></article></div>
    <section className="card plan-impact-chart"><div className="section-head"><div><span className="section-index">01</span><strong>季度計畫差異與庫存影響組合圖</strong></div><div className="chart-legend"><i className="variance"/>計畫增減台數 <i className="adjustment"/>備料調整金額</div></div><div className="combo-chart">{comparable.map(x=><div key={x.month} title={`前回計畫：${x.previous}；今回計畫：${x.current}；差異台數：${x.variance}；差異率：${x.variancePct===null?"—":fmtPct(x.variancePct)}；庫存影響：${fmtWan(x.inventoryImpact||0)} 萬元`}><span>{fmtWan(x.materialAdjustment||0)} 萬</span><div className="combo-axis"><i className={(x.variance||0)>=0?"positive":"negative"} style={{height:`${Math.abs(x.variance||0)/maxVariance*46}%`}}/><em style={{bottom:`${50+(x.materialAdjustment||0)/maxAdjustment*40}%`}}/></div><b>{x.month.slice(5)}</b></div>)}</div><p className="chart-note">正值與負值皆以紅色警示：正值代表計畫上修與追加備料；負值代表計畫下修與庫存增加風險。改善或風險降低才使用綠色。</p></section>
    <section className="card plan-definition"><b>分析說明</b><p>計畫差異分析用於判斷生產計畫變動所帶動的備料調整，以及對後續庫存金額、缺料與高庫存風險的影響。</p></section>
    <section className="card inventory-table"><div className="section-head"><div><span className="section-index">02</span><strong>計畫差異對備料與庫存影響</strong></div><span>金額單位：萬元</span></div><div className="table-wrap"><table><thead><tr><th>月份</th><th>前回／今回</th><th>增減量</th><th>差異率</th><th>備料調整</th><th>庫存影響</th><th>判讀</th></tr></thead><tbody>{data.plans.months.map(x=><tr key={x.month}><td><b>{x.month}</b></td><td>{x.previous ?? "—"} / {x.current ?? "—"}</td><td className={x.variance===0?"": "risk-text"}>{x.variance===null?"—":`${x.variance>=0?"+":""}${x.variance}`}</td><td className={x.variance===0?"": "risk-text"}>{x.variancePct===null?"—":fmtPct(x.variancePct)}</td><td>{x.materialAdjustment===null?"—":`${fmtWan(x.materialAdjustment)} 萬`}</td><td>{x.inventoryImpact===null?"—":`${fmtWan(x.inventoryImpact)} 萬`}</td><td><span className={`pill ${x.variance===0?"green":"red"}`}>{x.interpretation ?? x.status}</span></td></tr>)}</tbody></table></div></section>
  </>;
}

export default function Home() {
  const [started,setStarted]=useState(false); const [tab,setTab]=useState<Tab>("Executive"); const [period,setPeriod]=useState("2026 Q1–Q3"); const [activeData,setActiveData]=useState<any>(seedData);
  const [exporting,setExporting]=useState(false); const [exportMessage,setExportMessage]=useState("");
  data=activeData;
  const exportPptx=async()=>{try{setExporting(true);setExportMessage("");await exportInventoryPptx(data);setExportMessage("現代商務資訊圖表簡報已產生並下載。");}catch(reason){console.error(reason);setExportMessage("簡報產生失敗，請稍後再試。");}finally{setExporting(false);}};
  const content=useMemo(()=>tab==="Executive"?<Executive setTab={setTab}/>:tab==="QoQ"?<QoQPage/>:tab==="Plan"?<PlanPage/>:<InventoryPage group={tab as "ALL"|"FORGN"|"LP3"|"KD"}/>,[tab,activeData]);
  if(!started)return <Setup onRun={(result)=>{data=result;setActiveData(result);setPeriod(result.meta.quarters.join("–"));setStarted(true)}}/>;
  return <div className="app-shell"><aside><div className="side-brand"><div className="brand-mark">MY</div><div><b>My Inventory</b><span>Intelligence</span></div></div><nav>{tabs.map(t=><button key={t.key} className={tab===t.key?"active":""} onClick={()=>setTab(t.key)}><Icon name={t.key==="Plan"?"plan":t.key==="QoQ"?"trend":t.key==="Executive"?"grid":"box"}/><span>{t.label}</span>{t.key==="KD"&&!((data.groups as Record<string, unknown>).KD)&&<i>0</i>}</button>)}</nav><div className="side-note"><b>分析基準</b><span>FORGN 1.6 月</span><span>LP3 2.2 日</span></div><button className="reset" onClick={()=>setStarted(false)}><Icon name="settings"/>重新設定</button></aside>
    <main className="dashboard"><header className="topbar"><div className="crumb"><b>{data.meta.model}</b><span>/</span><em>{tab}</em></div><div><select value={period} onChange={e=>setPeriod(e.target.value)} aria-label="分析期間"><option>{data.meta.quarters.join("–")}</option><option>{data.meta.latestQuarter}</option></select><button className="print-export" onClick={()=>window.print()}>列印報表</button><button className="export pptx-export" onClick={exportPptx} disabled={exporting}>{exporting?"產生中…":"匯出 PPTX"}</button><span className="avatar">{data.meta.model}</span></div></header>{exportMessage&&<div className="export-toast" role="status">{exportMessage}<button aria-label="關閉通知" onClick={()=>setExportMessage("")}>×</button></div>}<div className="content">{content}</div><footer className="page-footer">My Inventory Intelligence · {data.meta.model} · 金額單位：萬元</footer></main>
  </div>;
}
