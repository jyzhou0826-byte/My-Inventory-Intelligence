import * as XLSX from "xlsx";

export type UploadSelection = {
  inventory: File[];
  previous: File | null;
  current: File | null;
};

export type AnalysisScope = { years: number[]; quarters: string[] };

const required = ["車型", "件號", "廠別", "Material Group", "原因細項", "策略備料金額(人工維護)", "策略備料金額(系統計算)", "生產需求金額", "總庫存金額", "總庫存量", "淨庫存金額", "庫存水準"];
const numberOf = (value: unknown) => {
  const n = Number(String(value ?? "0").replaceAll(",", ""));
  return Number.isFinite(n) ? n : 0;
};
const cleanHeader = (value: unknown) => String(value ?? "").trim().replace(/\s*的總和$/, "");

async function rowsFromFile(file: File): Promise<unknown[][]> {
  const bytes = await file.arrayBuffer();
  if (/\.csv$/i.test(file.name)) {
    const raw = new Uint8Array(bytes);
    let text = new TextDecoder("utf-8").decode(raw);
    if (text.includes("�")) text = new TextDecoder("big5").decode(raw);
    const book = XLSX.read(text, { type: "string", raw: true });
    return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header: 1, defval: null, raw: true }) as unknown[][];
  }
  const book = XLSX.read(bytes, { type: "array", raw: true, cellDates: true });
  return XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], { header: 1, defval: null, raw: true }) as unknown[][];
}

function bucket(value: unknown) {
  const text = String(value ?? "").toUpperCase().replaceAll(" ", "");
  if (text.includes("FORGN") || text.includes("海外採購")) return "FORGN";
  if (text.includes("KD")) return "KD";
  if (text.includes("LP3") || text.includes("TAXFREE") || text.includes("國產")) return "LP3";
  return "OTHER";
}

function diagnose(group: string, level: number, demand: number, qty: number) {
  if (group === "LP3") {
    if (demand <= 0 && qty <= 0) return ["無需求／無庫存", "資訊", "不納入缺料 KPI"];
    if (demand <= 0 && level > 180) return ["呆滯", "極高", "啟動去化"];
    if (demand <= 0 && level > 90) return ["積壓", "高", "檢討庫存去化"];
    if (demand <= 0 && qty > 0) return ["無需求庫存", "注意", "確認供應商在庫或子件用途"];
    if (demand > 0 && qty <= 0) return ["缺料", "極高", "立即確認供應"];
    if (demand > 0 && level < .5) return ["缺料風險", "高", "持續監控"];
    return ["安全", "正常", "安全庫存"];
  }
  if (group === "FORGN" || group === "KD") {
    if (level < .5) return ["缺料", "極高", "立即追料"];
    if (level < 1.6) return ["缺料", "高", "確認交期"];
    if (level <= 2.5) return ["安全", "正常", "持續監控"];
    if (level <= 3.5) return ["積壓", "高", "檢討庫存積壓"];
    return ["呆滯", "極高", "加速去化"];
  }
  return ["安全", "正常", "持續監控"];
}

function kdProject(row: Record<string, unknown>) {
  const text = [row["原因細項"], row["件名"], row["Material Group"], row["備註"]].map(value => String(value ?? "")).join(" ");
  if (text.includes("外開把手")) return "外開把手";
  if (text.includes("370轉置差異件") || (text.includes("370") && text.includes("轉置"))) return "370轉置差異件";
  return null;
}

function periodFromName(name: string, index: number) {
  const match = name.match(/(?:^|[_\s-])(20\d{2})[_\s-]*(Q[1-4])(?:\.|[_\s-]|$)/i);
  return { year: match ? Number(match[1]) : null, quarter: match?.[2].toUpperCase() ?? `Q${index + 1}` };
}

function findHeader(rows: unknown[][]) {
  const index = rows.findIndex(row => row.some(cell => cleanHeader(cell) === "件號") && row.some(cell => cleanHeader(cell) === "Material Group"));
  if (index < 0) throw new Error("找不到 Power BI 欄位標題列（件號、Material Group）。");
  const headers = rows[index].map(cleanHeader);
  const missing = required.filter(name => !headers.includes(name));
  if (missing.length) throw new Error(`Power BI 必要欄位缺少：${missing.join("、")}`);
  return { index, headers };
}

function monthLabel(value: unknown) {
  if (value instanceof Date) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}`;
  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}`;
  }
  const text = String(value ?? "").trim().replace(/^'/, "");
  const match = text.match(/(\d{2,4})\D+(\d{1,2})/);
  if (!match) return text;
  const year = Number(match[1]) < 100 ? Number(match[1]) + 2000 : Number(match[1]);
  return `${year}-${String(Number(match[2])).padStart(2, "0")}`;
}

function planFromRows(rows: unknown[][], model: string) {
  if (!rows.length) return { found: false, label: null, months: [] as {month:string,value:number}[] };
  const target = rows.slice(1).find(row => {
    const label = String(row[0] ?? "").trim().toUpperCase();
    return label === model || label === `${model}計` || label.startsWith(`${model}計`);
  });
  if (!target) return { found: false, label: null, months: [] as {month:string,value:number}[] };
  return { found: true, label: String(target[0]), months: rows[0].slice(1).map((header, i) => ({ month: monthLabel(header), value: numberOf(target[i + 1]) })).filter(x => x.month) };
}

export async function analyzeUploads(modelInput: string, selection: UploadSelection, scope?: AnalysisScope) {
  const model = modelInput.trim().toUpperCase();
  if (!model) throw new Error("請輸入車型。");
  if (!selection.inventory.length || !selection.previous || !selection.current) throw new Error("請完整上傳庫存資料、前回計畫與本回計畫。");
  const scopedFiles = selection.inventory.filter((file, index) => {
    if (!scope) return true;
    const period = periodFromName(file.name, index);
    return (period.year === null || scope.years.includes(period.year)) && scope.quarters.includes(period.quarter);
  });
  if (!scopedFiles.length) throw new Error(`找不到所選期間（${scope?.years.join("、")} ${scope?.quarters.join("、")}）的庫存檔案，請確認檔名包含年度與季度，例如 RV_2026_Q1.xlsx。`);
  const quarters: any[] = [];
  const allItems: any[] = [];
  const years = new Set<number>();
  for (const [index, file] of scopedFiles.entries()) {
    const rows = await rowsFromFile(file);
    const { index: headerIndex, headers } = findHeader(rows);
    const records = rows.slice(headerIndex + 1).map(values => Object.fromEntries(headers.map((h, i) => [h, values[i]]))).filter(row => row["件號"]);
    const wrongModel = records.find(row => String(row["車型"] ?? "").trim().toUpperCase() !== model);
    if (wrongModel) throw new Error(`【車型資料異常】${file.name} 內含車型 ${wrongModel["車型"]}，與輸入的 ${model} 不一致。`);
    const period = periodFromName(file.name, index);
    const quarter = period.year ? `${String(period.year).slice(-2)}/${period.quarter}` : period.quarter;
    if (period.year) years.add(period.year);
    const items = records.map(row => {
      const group = bucket(row["Material Group"]);
      const reason = String(row["原因細項"] ?? "—") || "—";
      const note = String(row["備註"] ?? "");
      if (group === "LP3" && `${reason} ${note}`.includes("子件")) return null;
      const project = group === "KD" ? kdProject(row) : null;
      if (group === "KD" && !project) return null;
      const level = numberOf(row["庫存水準"]);
      const demand = numberOf(row["生產需求金額"]);
      const qty = numberOf(row["總庫存量"]);
      const inventory = numberOf(row["總庫存金額"]);
      const net = numberOf(row["淨庫存金額"]);
      const [risk, severity, action] = diagnose(group, level, demand, qty);
      return { model, part:String(row["件號"]), partName:String(row["件名"] ?? "").trim(), plant:String(row["廠別"] ?? ""), group, rawGroup:String(row["Material Group"] ?? ""), reason, note, kdProject:project, date:String(row["日期"] ?? ""), manual:numberOf(row["策略備料金額(人工維護)"]), system:numberOf(row["策略備料金額(系統計算)"]), demand, inventory, qty, net, level, risk, severity, action, quarter };
    }).filter((item): item is NonNullable<typeof item> => item !== null);
    const totals = Object.fromEntries(["manual","system","demand","inventory","net"].map(key => [key, items.reduce((sum, row) => sum + row[key], 0)]));
    const statusNames = ["安全","缺料","缺料風險","積壓","呆滯","無需求庫存","無需求／無庫存"];
    const counts = Object.fromEntries(statusNames.map(risk => [risk, items.filter(row => row.risk === risk).length]));
    const groupTotals = Object.fromEntries(["FORGN","LP3","KD","OTHER"].map(group => [group, items.filter(row => row.group === group).reduce((sum,row)=>sum+row.inventory,0)]));
    const groupMetrics = Object.fromEntries(["FORGN","LP3","KD","OTHER"].map(group => { const groupRows=items.filter(row=>row.group===group); return [group,{count:groupRows.length,inventory:groupRows.reduce((sum,row)=>sum+row.inventory,0),demand:groupRows.reduce((sum,row)=>sum+row.demand,0),safe:groupRows.filter(row=>row.risk==="安全").length}]}));
    quarters.push({ quarter, sortKey: period.year ? period.year * 10 + Number(period.quarter.slice(1)) : index, records: items.length, totals, counts, groupTotals, groupMetrics });
    allItems.push(...items);
  }
  quarters.sort((a,b) => a.sortKey - b.sortKey);
  const latestQuarter = quarters.at(-1)?.quarter;
  const latestItems = allItems.filter(row => row.quarter === latestQuarter);
  const latestTotals = Object.fromEntries(["manual","system","demand","inventory","net"].map(key => [key, latestItems.reduce((sum,row)=>sum+row[key],0)]));
  const statusNames = ["安全","缺料","缺料風險","積壓","呆滯","無需求庫存","無需求／無庫存"];
  const latestCounts = Object.fromEntries(statusNames.map(risk => [risk, latestItems.filter(row => row.risk === risk).length]));
  const groups: Record<string, any> = {};
  for (const group of ["FORGN","LP3","KD","OTHER"]) {
    const rows = latestItems.filter(row => row.group === group); if (!rows.length) continue;
    const manual=rows.reduce((s,r)=>s+r.manual,0), system=rows.reduce((s,r)=>s+r.system,0);
    groups[group]={count:rows.length,inventory:rows.reduce((s,r)=>s+r.inventory,0),demand:rows.reduce((s,r)=>s+r.demand,0),net:rows.reduce((s,r)=>s+r.net,0),manual,system,avgLevel:rows.reduce((s,r)=>s+r.level,0)/rows.length,safe:rows.filter(r=>r.risk==="安全").length,manualRatio:manual+system?manual/(manual+system):0};
  }
  const tops = {
    缺料: latestItems.filter(row=>row.demand>0 && (row.risk==="缺料" || row.risk==="缺料風險")).sort((a,b)=>b.inventory-a.inventory).slice(0,3),
    積壓: latestItems.filter(row=>row.risk==="積壓").sort((a,b)=>b.inventory-a.inventory).slice(0,3),
    呆滯: latestItems.filter(row=>row.risk==="呆滯").sort((a,b)=>b.inventory-a.inventory).slice(0,3),
  };
  const previous = planFromRows(await rowsFromFile(selection.previous), model), current = planFromRows(await rowsFromFile(selection.current), model);
  if (!previous.found || !current.found) throw new Error(`【車型資料異常】生產計畫找不到 ${model} 或 ${model}計，請確認檔案內容。`);
  const prevMap=new Map(previous.months.map(x=>[x.month,x.value])), currMap=new Map(current.months.map(x=>[x.month,x.value]));
  const comparableCurrentUnits = [...currMap.entries()].filter(([month]) => prevMap.has(month)).reduce((sum,[,value])=>sum+value,0);
  const demandPerVehicle = comparableCurrentUnits > 0 ? latestTotals.demand / comparableCurrentUnits : 0;
  const months=[...new Set([...prevMap.keys(),...currMap.keys()])].sort().map(month=>{
    const pv=prevMap.get(month)??null,cv=currMap.get(month)??null;
    const variance=pv===null||cv===null?null:cv-pv;
    const variancePct=variance===null||!pv?null:variance/pv;
    const status=variancePct===null?"無法比較":variancePct>=.5?"重大需求增加":variancePct>=.2?"需求波動":variancePct<=-.5?"呆滯風險":variancePct<=-.2?"需求急降":"穩定";
    const materialAdjustment=variance===null?null:variance*demandPerVehicle;
    const inventoryImpact=materialAdjustment===null?null:-materialAdjustment;
    const interpretation=variance===null?"無法比較":variance>0?"追加備料需求與缺料風險":variance<0?"多餘備料與庫存增加風險":"無明顯影響";
    return{month,previous:pv,current:cv,variance,variancePct,status,materialAdjustment,inventoryImpact,projectedInventory:inventoryImpact===null?null:latestTotals.inventory+inventoryImpact,projectedNet:inventoryImpact===null?null:latestTotals.net+inventoryImpact,interpretation};
  });
  const highRisk=latestItems.filter(row=>["缺料","缺料風險","積壓","呆滯"].includes(row.risk));
  const healthBase=latestItems.filter(row=>!["無需求庫存","無需求／無庫存"].includes(row.risk));
  const safeRate=healthBase.length?healthBase.filter(row=>row.risk==="安全").length/healthBase.length:0;
  const kdProjects = Object.fromEntries(["外開把手","370轉置差異件"].map(project => {
    const rows=latestItems.filter(row=>row.group==="KD"&&row.kdProject===project);
    const counts=Object.fromEntries(statusNames.map(risk=>[risk,rows.filter(row=>row.risk===risk).length]));
    return [project,{count:rows.length,inventory:rows.reduce((sum,row)=>sum+row.inventory,0),net:rows.reduce((sum,row)=>sum+row.net,0),demand:rows.reduce((sum,row)=>sum+row.demand,0),counts,items:rows}];
  }));
  return { meta:{model,year:years.size === 1 ? [...years][0] : (scope?.years.join("、") ?? new Date().getFullYear()),quarters:quarters.map(q=>q.quarter),generated:new Date().toISOString(),sourceFiles:scopedFiles.map(f=>f.name),latestQuarter,schemaValid:true,headers:required},quarters,latest:{totals:latestTotals,counts:latestCounts,items:latestItems.length,safeRate,healthBaseItems:healthBase.length,highRiskInventory:highRisk.reduce((s,r)=>s+r.inventory,0),highRiskRatio:latestTotals.inventory?highRisk.reduce((s,r)=>s+r.inventory,0)/latestTotals.inventory:0},groups,kdProjects,tops,plans:{previous:{file:selection.previous.name,...previous},current:{file:selection.current.name,...current},months,demandPerVehicle,modelMismatch:false},latestItems};
}
