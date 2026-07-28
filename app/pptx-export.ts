const C = {
  snow: "F6F8FB", paper: "FFFFFF", ink: "0F172A", muted: "64748B", line: "DCE3EC",
  pine: "2563EB", sage: "93C5FD", mist: "E8F0FE", sky: "06B6D4",
  clay: "EF4444", ochre: "F59E0B", moss: "10B981",
};
const FONT = "Microsoft JhengHei";
const wan = (n: number) => Number((n / 10000).toFixed(1));
const pct = (n: number) => `${((Number.isFinite(n) ? n : 0) * 100).toFixed(1)}%`;

export async function exportInventoryPptx(data: any) {
  const { default: PptxGenJS } = await import("pptxgenjs");
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "My Inventory Intelligence";
  pptx.company = "My Inventory Intelligence";
  pptx.subject = "庫存決策儀表板";
  pptx.title = `${data.meta.model} 庫存決策簡報`;
  pptx.lang = "zh-TW";
  pptx.theme = { headFontFace: FONT, bodyFontFace: FONT, lang: "zh-TW" };

  const totals = data.latest.totals;
  const quarters = data.quarters ?? [];
  const previous = quarters.length > 1 ? quarters[quarters.length - 2] : null;
  const qoq = previous?.totals?.inventory
    ? (totals.inventory - previous.totals.inventory) / previous.totals.inventory : null;

  const text = (slide: any, value: string, x: number, y: number, w: number, h: number, options: any = {}) =>
    slide.addText(value, {
      x, y, w, h, fontFace: FONT, fontSize: 16, color: C.ink, margin: 0,
      breakLine: false, valign: "mid", fit: "shrink", ...options,
    });
  const page = (slide: any, eyebrow: string, title: string, number: number) => {
    slide.background = { color: C.snow };
    text(slide, eyebrow.toUpperCase(), .62, .34, 5.5, .25, {
      fontSize: 10, bold: true, color: C.pine, charSpacing: 1.6,
    });
    text(slide, title, .62, .64, 11.8, .66, { fontSize: 35, bold: true });
    slide.addShape(pptx.ShapeType.line, {
      x: .62, y: 1.32, w: 12.08, h: 0, line: { color: C.line, width: 1 },
    });
    text(slide, String(number).padStart(2, "0"), 12.05, 7.08, .62, .2, {
      fontSize: 9, color: C.muted, align: "right",
    });
  };
  const metric = (slide: any, x: number, y: number, w: number, label: string, value: string, note: string, color = C.pine) => {
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w, h: 1.28, fill: { color: C.paper }, line: { color: C.line, width: 1 }, radius: .08,
    });
    slide.addShape(pptx.ShapeType.rect, {
      x, y, w: .07, h: 1.28, fill: { color }, line: { color, transparency: 100 },
    });
    text(slide, label, x + .22, y + .12, w - .42, .22, { fontSize: 11, bold: true, color: C.muted });
    text(slide, value, x + .22, y + .39, w - .42, .42, { fontSize: 23, bold: true });
    text(slide, note, x + .22, y + .91, w - .42, .2, { fontSize: 9.5, color: C.muted });
  };
  const source = (slide: any) => text(
    slide,
    `資料來源：${data.meta.sourceFiles?.join("、") || "使用者上傳資料"}｜${data.meta.year} ${data.meta.latestQuarter}`,
    .62, 7.04, 10.5, .2, { fontSize: 8.5, color: "87928E" },
  );
  const svgData = (svg: string) => {
    const bytes = new TextEncoder().encode(svg);
    let binary = "";
    bytes.forEach(byte => { binary += String.fromCharCode(byte); });
    return `data:image/svg+xml;base64,${btoa(binary)}`;
  };
  const esc = (value: unknown) => String(value).replace(/[&<>"']/g, match => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;",
  }[match] as string));
  const lineChart = (labels: string[], values: number[]) => {
    const width = 900, height = 500, left = 86, right = 34, top = 46, bottom = 74;
    const max = Math.max(...values, 1) * 1.12;
    const x = (index: number) => labels.length === 1 ? (left + width - right) / 2 : left + index * (width - left - right) / (labels.length - 1);
    const y = (value: number) => top + (max - value) * (height - top - bottom) / max;
    const grid = Array.from({ length: 5 }, (_, index) => {
      const value = max * (4 - index) / 4, py = y(value);
      return `<line x1="${left}" y1="${py}" x2="${width-right}" y2="${py}" stroke="#DCE3EC" stroke-width="1"/>
        <text x="${left-14}" y="${py+5}" text-anchor="end" font-size="16" fill="#64748B">${Math.round(value).toLocaleString()}</text>`;
    }).join("");
    const path = values.map((value, index) => `${index ? "L" : "M"} ${x(index)} ${y(value)}`).join(" ");
    const points = values.map((value, index) => `<circle cx="${x(index)}" cy="${y(value)}" r="7" fill="#FFFFFF" stroke="#2563EB" stroke-width="4"/>
      <text x="${x(index)}" y="${y(value)-18}" text-anchor="middle" font-size="18" font-weight="700" fill="#0F172A">${value.toLocaleString()}</text>
      <text x="${x(index)}" y="${height-28}" text-anchor="middle" font-size="17" font-weight="700" fill="#334155">${esc(labels[index])}</text>`).join("");
    return svgData(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <rect width="${width}" height="${height}" fill="#FFFFFF"/>${grid}
      <path d="${path}" fill="none" stroke="#2563EB" stroke-width="6" stroke-linecap="round" stroke-linejoin="round"/>${points}</svg>`);
  };
  const groupedBars = (labels: string[], series: { name: string; values: number[]; color: string }[]) => {
    const width = 900, height = 500, left = 82, right = 28, top = 54, bottom = 92;
    const max = Math.max(...series.flatMap(item => item.values), 1) * 1.14;
    const groupW = (width - left - right) / Math.max(labels.length, 1);
    const barW = Math.min(54, groupW / (series.length + 1));
    const grid = Array.from({ length: 5 }, (_, index) => {
      const value = max * (4 - index) / 4;
      const py = top + (max - value) * (height - top - bottom) / max;
      return `<line x1="${left}" y1="${py}" x2="${width-right}" y2="${py}" stroke="#DCE3EC" stroke-width="1"/>
        <text x="${left-12}" y="${py+5}" text-anchor="end" font-size="15" fill="#64748B">${Math.round(value).toLocaleString()}</text>`;
    }).join("");
    const bars = labels.map((label, index) => {
      const center = left + groupW * index + groupW / 2;
      const blocks = series.map((item, si) => {
        const value = item.values[index] ?? 0;
        const h = value * (height - top - bottom) / max;
        const bx = center + (si - (series.length - 1) / 2) * (barW + 8) - barW / 2;
        const by = height - bottom - h;
        return `<rect x="${bx}" y="${by}" width="${barW}" height="${h}" rx="4" fill="#${item.color}"/>
          <text x="${bx + barW / 2}" y="${Math.max(top + 14, by - 10)}" text-anchor="middle" font-size="16" font-weight="700" fill="#0F172A">${value.toLocaleString()}</text>`;
      }).join("");
      return `${blocks}<text x="${center}" y="${height-55}" text-anchor="middle" font-size="17" font-weight="700" fill="#334155">${esc(label)}</text>`;
    }).join("");
    const legend = series.map((item, index) => `<rect x="${left + index * 210}" y="${height-24}" width="20" height="8" rx="4" fill="#${item.color}"/>
      <text x="${left + 30 + index * 210}" y="${height-16}" font-size="15" fill="#475569">${esc(item.name)}</text>`).join("");
    return svgData(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
      <rect width="${width}" height="${height}" fill="#FFFFFF"/>${grid}${bars}${legend}</svg>`);
  };
  const donutChart = (labels: string[], values: number[], colors: string[], centerValue = pct(data.latest.safeRate), centerLabel = "庫存健康度") => {
    const total = Math.max(values.reduce((sum, value) => sum + value, 0), 1);
    const radius = 142, circumference = 2 * Math.PI * radius;
    let offset = 0;
    const arcs = values.map((value, index) => {
      const length = value / total * circumference;
      const result = `<circle cx="250" cy="250" r="${radius}" fill="none" stroke="#${colors[index]}" stroke-width="58"
        stroke-dasharray="${length} ${circumference - length}" stroke-dashoffset="${-offset}" transform="rotate(-90 250 250)"/>`;
      offset += length;
      return result;
    }).join("");
    const legend = labels.map((label, index) => `<rect x="500" y="${110 + index * 72}" width="18" height="18" rx="4" fill="#${colors[index]}"/>
      <text x="535" y="${126 + index * 72}" font-size="20" font-weight="700" fill="#0F172A">${esc(label)}</text>
      <text x="820" y="${126 + index * 72}" text-anchor="end" font-size="20" font-weight="700" fill="#0F172A">${values[index]} 件</text>
      <text x="885" y="${126 + index * 72}" text-anchor="end" font-size="17" fill="#64748B">${(values[index] / total * 100).toFixed(1)}%</text>`).join("");
    return svgData(`<svg xmlns="http://www.w3.org/2000/svg" width="920" height="500" viewBox="0 0 920 500">
      <rect width="920" height="500" fill="#FFFFFF"/>${arcs}
      <text x="250" y="238" text-anchor="middle" font-size="46" font-weight="800" fill="#0F172A">${esc(centerValue)}</text>
      <text x="250" y="278" text-anchor="middle" font-size="20" fill="#64748B">${esc(centerLabel)}</text>${legend}</svg>`);
  };

  // 封面：深藍、亮藍與俐落幾何線條，適合正式企業提案。
  {
    const slide = pptx.addSlide();
    slide.background = { color: C.snow };
    slide.addShape(pptx.ShapeType.rect, {
      x: 8.7, y: 0, w: 4.63, h: 7.5, fill: { color: C.ink }, line: { color: C.ink, transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 9.32, y: 1.08, w: 2.92, h: .18,
      fill: { color: C.pine }, line: { color: C.pine, transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 9.32, y: 1.52, w: 2.1, h: .18,
      fill: { color: C.sky }, line: { color: C.sky, transparency: 100 },
    });
    [0, 1, 2, 3].forEach(index => slide.addShape(pptx.ShapeType.line, {
      x: 9.32, y: 4.35 + index * .42, w: 2.82 - index * .34, h: 0,
      line: { color: index === 0 ? C.pine : "475569", width: index === 0 ? 4 : 2 },
    }));
    text(slide, "MY INVENTORY INTELLIGENCE", .72, .56, 6.6, .28, {
      fontSize: 11, bold: true, color: C.pine, charSpacing: 1.8,
    });
    text(slide, `${data.meta.model} 庫存決策\n儀表板簡報`, .72, 1.58, 7.4, 1.68, {
      fontSize: 50, bold: true, breakLine: true, valign: "top",
    });
    text(slide, `${data.meta.year} ${data.meta.quarters.join("–")}｜庫存健康、關鍵風險與行動建議`, .72, 3.65, 7.3, .42, {
      fontSize: 17, color: C.muted,
    });
    text(slide, "MODERN BUSINESS · INFOGRAPHIC REPORT", .72, 6.56, 6.5, .28, {
      fontSize: 11, bold: true, color: C.pine, charSpacing: 1.2,
    });
  }

  {
    const slide = pptx.addSlide();
    page(slide, "KD project portfolio", "KD 僅追蹤兩個指定專案，分開管理", 13);
    [
      { key: "外開把手", type: "策略備料", prompt: "是否需要持續備料？備料是否合理？", color: C.pine },
      { key: "370轉置差異件", type: "庫存去化", prompt: "剩餘庫存是否有效消化？是否存在呆滯風險？", color: C.ochre },
    ].forEach((project, index) => {
      const p = data.kdProjects?.[project.key];
      const x = .62 + index * 6.18;
      slide.addShape(pptx.ShapeType.rect, { x, y: 1.62, w: 5.88, h: 4.82, fill: { color: C.paper }, line: { color: C.line, width: 1 }, radius: .08 });
      slide.addShape(pptx.ShapeType.rect, { x, y: 1.62, w: 5.88, h: .08, fill: { color: project.color }, line: { color: project.color, transparency: 100 } });
      text(slide, project.type, x + .28, 1.92, 1.6, .25, { fontSize: 11, bold: true, color: project.color });
      text(slide, project.key, x + .28, 2.28, 4.9, .45, { fontSize: 27, bold: true });
      metric(slide, x + .28, 3.05, 2.48, "總庫存金額", p ? `${wan(p.inventory).toLocaleString()} 萬元` : "本期無資料", p ? `${p.count} 件` : "未補造數字", project.color);
      metric(slide, x + 3.0, 3.05, 2.48, "生產需求金額", p ? `${wan(p.demand).toLocaleString()} 萬元` : "—", "專案需求指標", C.sky);
      text(slide, project.prompt, x + .28, 4.68, 5.22, .38, { fontSize: 15, bold: true });
      text(slide, "分析與建議", x + .28, 5.28, 1.35, .24, { fontSize: 12, bold: true, color: C.muted });
      slide.addShape(pptx.ShapeType.line, { x: x + 1.58, y: 5.46, w: 3.92, h: 0, line: { color: C.line, width: 1 } });
      slide.addShape(pptx.ShapeType.line, { x: x + .28, y: 5.88, w: 5.22, h: 0, line: { color: C.line, width: 1 } });
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Executive overview", "庫存上升，健康度仍處於危險區", 2);
    metric(slide, .62, 1.62, 3.82, "總庫存金額", `${wan(totals.inventory).toLocaleString()} 萬元`,
      qoq === null ? "目前為單季分析基準" : `QoQ ${qoq >= 0 ? "+" : ""}${pct(qoq)}`, C.clay);
    metric(slide, 4.75, 1.62, 3.82, "庫存健康度", pct(data.latest.safeRate),
      `安全 ${data.latest.counts.安全} 件／總計 ${data.latest.items} 件`, C.clay);
    metric(slide, 8.88, 1.62, 3.82, "高風險占比", pct(data.latest.highRiskRatio),
      `${wan(data.latest.highRiskInventory).toLocaleString()} 萬元需優先處置`, C.ochre);
    text(slide, "主管判讀", .62, 3.32, 2.2, .34, { fontSize: 19, bold: true });
    const insights = [
      ["01", "庫存增速", qoq === null ? "需累積兩季以上判讀 QoQ。" : `最新一季庫存${qoq >= 0 ? "增加" : "減少"} ${pct(Math.abs(qoq))}。`],
      ["02", "風險結構", `積壓 ${data.latest.counts.積壓} 件，是最大宗風險。`],
      ["03", "總庫存金額", `最新一季 ${wan(totals.inventory).toLocaleString()} 萬元，應搭配需求與高風險金額判讀。`],
    ];
    insights.forEach((item, index) => {
      const y = 3.85 + index * .78;
      text(slide, item[0], .72, y, .45, .25, { fontSize: 11, bold: true, color: C.pine });
      text(slide, item[1], 1.3, y - .02, 2.25, .3, { fontSize: 17, bold: true });
      text(slide, item[2], 3.55, y - .02, 8.7, .3, { fontSize: 14, color: C.muted });
      slide.addShape(pptx.ShapeType.line, { x: .72, y: y + .47, w: 11.5, h: 0, line: { color: C.line, width: .7 } });
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Quarterly movement", "所有上傳季度均納入庫存趨勢比較", 3);
    slide.addImage({
      data: lineChart(
        quarters.map((q: any) => `${String(data.meta.year).slice(-2)}/${q.quarter}`),
        quarters.map((q: any) => wan(q.totals.inventory)),
      ),
      x: .62, y: 1.62, w: 8.2, h: 4.8,
    });
    const first = quarters[0], latest = quarters[quarters.length - 1];
    const change = first?.totals?.inventory ? (latest.totals.inventory - first.totals.inventory) / first.totals.inventory : null;
    metric(slide, 9.18, 1.62, 3.52, "最新季度", `${wan(latest.totals.inventory).toLocaleString()} 萬元`,
      `${String(data.meta.year).slice(-2)}/${latest.quarter} 實績`, C.pine);
    metric(slide, 9.18, 3.18, 3.52, "分析期累計變化", change === null ? "—" : `${change >= 0 ? "+" : ""}${pct(change)}`,
      `${first.quarter} 至 ${latest.quarter}`, change !== null && change > 0 ? C.clay : C.moss);
    text(slide, "判讀", 9.18, 4.9, 1.2, .28, { fontSize: 17, bold: true, color: C.pine });
    text(slide, "庫存增幅應與需求及生產計畫同步檢視，避免需求放緩後形成積壓。", 9.18, 5.28, 3.42, .86, {
      fontSize: 14, color: C.muted, valign: "top",
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Inventory health", "積壓與缺料並存，需採雙軌管理", 4);
    const statuses = ["缺料", "安全", "積壓", "呆滯"];
    slide.addImage({
      data: donutChart(
        statuses,
        statuses.map(name => Number(data.latest.counts[name] ?? 0)),
        [C.clay, C.moss, C.ochre, "7C3AED"],
      ),
      x: .62, y: 1.58, w: 6.0, h: 4.85,
    });
    statuses.forEach((name, index) => {
      const y = 1.72 + index * 1.14;
      const colors: Record<string, string> = { 缺料: C.clay, 安全: C.moss, 積壓: C.ochre, 呆滯: "B85C45" };
      slide.addShape(pptx.ShapeType.rect, { x: 6.72, y, w: .08, h: .72, fill: { color: colors[name] }, line: { color: colors[name], transparency: 100 } });
      text(slide, name, 6.98, y, 1.6, .28, { fontSize: 17, bold: true });
      text(slide, `${data.latest.counts[name] ?? 0} 件`, 8.58, y, 1.5, .28, { fontSize: 17, bold: true, align: "right" });
      text(slide, pct((data.latest.counts[name] ?? 0) / Math.max(data.latest.items, 1)), 10.22, y, 1.65, .28, { fontSize: 14, color: C.muted, align: "right" });
      slide.addShape(pptx.ShapeType.line, { x: 6.98, y: y + .6, w: 4.9, h: 0, line: { color: C.line, width: .7 } });
    });
    text(slide, "健康度＝安全庫存件數 ÷ 總件數；40% 以下評級為危險。", 6.98, 6.28, 5, .3, { fontSize: 11, color: C.muted });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Material groups", "分類比較揭示庫存金額集中位置", 5);
    const names = ["FORGN", "LP3", "KD"];
    slide.addImage({
      data: groupedBars(names, [
        { name: "總庫存金額", values: names.map(name => wan(data.groups?.[name]?.inventory ?? 0)), color: C.pine },
        { name: "生產需求金額", values: names.map(name => wan(data.groups?.[name]?.demand ?? 0)), color: C.sky },
      ]),
      x: .62, y: 1.62, w: 7.7, h: 4.92,
    });
    names.forEach((name, index) => {
      const y = 1.72 + index * 1.48, group = data.groups?.[name];
      text(slide, name, 8.82, y, 1.42, .3, { fontSize: 18, bold: true, color: index === 0 ? C.pine : C.ink });
      text(slide, group ? `${wan(group.inventory).toLocaleString()} 萬元` : "本期無資料", 10.16, y, 2.14, .3, { fontSize: 17, bold: true, align: "right" });
      text(slide, group ? `${group.count} 件｜平均水準 ${group.avgLevel.toFixed(2)} ${name === "LP3" ? "日" : "月"}` : "未補造數字",
        8.82, y + .43, 3.48, .26, { fontSize: 11, color: C.muted });
      slide.addShape(pptx.ShapeType.line, { x: 8.82, y: y + .94, w: 3.48, h: 0, line: { color: C.line, width: .7 } });
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Financial KPI", "五項金額 KPI 完整呈現庫存資金結構", 6);
    const coverage = totals.demand ? totals.inventory / totals.demand : 0;
    metric(slide, .62, 1.62, 3.82, "總庫存金額", `${wan(totals.inventory).toLocaleString()} 萬元`,
      qoq === null ? "本期分析基準" : `較上季 ${qoq >= 0 ? "+" : ""}${pct(qoq)}`, C.pine);
    metric(slide, 4.75, 1.62, 3.82, "生產需求金額", `${wan(totals.demand).toLocaleString()} 萬元`,
      `需求覆蓋率 ${coverage.toFixed(2)}x`, C.sky);
    metric(slide, 8.88, 1.62, 3.82, "QoQ 變化", qoq === null ? "基準季" : `${qoq >= 0 ? "+" : ""}${pct(qoq)}`,
      previous ? `${previous.quarter} → ${data.meta.latestQuarter}` : "無前期", qoq !== null && qoq > 0 ? C.clay : C.moss);
    metric(slide, .62, 3.25, 3.82, "需求覆蓋率", `${coverage.toFixed(2)} x`,
      `總庫存金額 ÷ 生產需求金額`, C.sky);
    metric(slide, 4.75, 3.25, 3.82, "高風險庫存金額", `${wan(data.latest.highRiskInventory).toLocaleString()} 萬元`,
      `高風險占比 ${pct(data.latest.highRiskRatio)}`, C.clay);
    metric(slide, 8.88, 3.25, 3.82, "有效件號", `${data.latest.items} 件`,
      `安全 ${data.latest.counts.安全}｜缺料 ${data.latest.counts.缺料}｜積壓 ${data.latest.counts.積壓}｜呆滯 ${data.latest.counts.呆滯}`, C.moss);
    text(slide, "管理判讀", .62, 5.08, 1.6, .32, { fontSize: 20, bold: true });
    text(slide,
      `目前每 1 元生產需求由 ${coverage.toFixed(2)} 元總庫存覆蓋；高風險資金應依金額排序推動改善。`,
      .62, 5.55, 11.6, .72, { fontSize: 17, color: C.muted, valign: "top" });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Total inventory", "以總庫存金額掌握季度變化與資金風險", 7);
    metric(slide, .62, 1.62, 3.82, "總庫存金額", `${wan(totals.inventory).toLocaleString()} 萬元`, "最新季度實績", C.pine);
    metric(slide, 4.75, 1.62, 3.82, "QoQ 變化", qoq === null ? "基準季" : `${qoq >= 0 ? "+" : ""}${pct(qoq)}`, previous ? `${previous.quarter} → ${data.meta.latestQuarter}` : "無前期", qoq !== null && qoq > 0 ? C.clay : C.moss);
    metric(slide, 8.88, 1.62, 3.82, "高風險庫存", `${wan(data.latest.highRiskInventory).toLocaleString()} 萬元`, `占總庫存 ${pct(data.latest.highRiskRatio)}`, C.clay);
    text(slide, "季度總庫存趨勢", .62, 3.28, 2.5, .3, { fontSize: 20, bold: true });
    slide.addShape(pptx.ShapeType.rect, {
      x: .62, y: 3.86, w: 11.76, h: .52, fill: { color: "E2E8F0" }, line: { color: "E2E8F0", transparency: 100 },
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: .62, y: 3.86, w: 11.76, h: .52, fill: { color: C.pine }, line: { color: C.pine, transparency: 100 },
    });
    text(slide, `最新總庫存 ${wan(totals.inventory).toLocaleString()} 萬元`, .62, 4.55, 4.2, .28, { fontSize: 16, bold: true, color: C.pine });
    text(slide, "管理定義", .62, 5.25, 1.6, .28, { fontSize: 20, bold: true });
    text(slide, "本頁僅以總庫存金額呈現資金規模，不進行庫存組成拆分。", 2.12, 5.19, 9.6, .42, { fontSize: 18, bold: true, color: C.ink });
    text(slide, "策略備料金額不納入庫存決策分析。", .62, 5.86, 11.6, .45, {
      fontSize: 16, color: C.muted,
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Quarterly total inventory", "每季僅呈現總庫存金額", 8);
    const labels = quarters.map((q: any) => `${String(data.meta.year).slice(-2)}/${q.quarter}`);
    const chartMax = Math.max(...quarters.map((q: any) => q.totals.inventory), 1);
    quarters.forEach((q: any, index: number) => {
      const x = 1.0 + index * (8.0 / Math.max(quarters.length, 1));
      const height = 3.9 * q.totals.inventory / chartMax;
      text(slide, `${wan(q.totals.inventory).toLocaleString()}`, x - .18, 5.85 - height, 1.45, .24, { fontSize: 13, bold: true, align: "center" });
      slide.addShape(pptx.ShapeType.rect, { x, y: 6.12 - height, w: 1.08, h: height, fill: { color: C.pine }, line: { color: C.pine, transparency: 100 } });
      text(slide, labels[index], x - .18, 6.24, 1.45, .24, { fontSize: 13, bold: true, align: "center" });
    });
    metric(slide, 9.35, 1.62, 3.35, "最新總庫存", `${wan(totals.inventory).toLocaleString()} 萬元`, "本期實績", C.pine);
    metric(slide, 9.35, 3.18, 3.35, "QoQ 變化", qoq === null ? "基準季" : `${qoq >= 0 ? "+" : ""}${pct(qoq)}`, previous ? `${previous.quarter} → ${data.meta.latestQuarter}` : "無前期", qoq !== null && qoq > 0 ? C.clay : C.moss);
    metric(slide, 9.35, 4.74, 3.35, "高風險占比", pct(data.latest.highRiskRatio), "依風險料號金額計算", C.clay);
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Quarterly demand trend", "生產需求金額獨立呈現，不與庫存組成混用", 9);
    slide.addImage({ data: lineChart(
      quarters.map((q: any) => `${String(data.meta.year).slice(-2)}/${q.quarter}`),
      quarters.map((q: any) => wan(q.totals.demand)),
    ), x: .62, y: 1.62, w: 8.25, h: 4.85 });
    metric(slide, 9.2, 1.62, 3.5, "最新生產需求", `${wan(totals.demand).toLocaleString()} 萬元`, "需求指標", C.sky);
    text(slide, "分析與建議", 9.2, 3.25, 2.2, .3, { fontSize: 19, bold: true });
    slide.addShape(pptx.ShapeType.rect, { x: 9.2, y: 3.75, w: 3.5, h: 2.1, fill: { color: C.paper }, line: { color: C.line, width: 1 }, radius: .08 });
    text(slide, "（可編輯留白）", 9.45, 4.62, 3, .3, { fontSize: 13, color: "94A3B8", align: "center" });
    source(slide);
  }

  for (const [groupIndex, groupName] of ["FORGN", "LP3"].entries()) {
    const slide = pptx.addSlide();
    const group = data.groups?.[groupName];
    const rows = (data.latestItems ?? []).filter((item: any) => item.group === groupName);
    const groupSafe = rows.filter((item: any) => item.risk === "安全").length;
    const groupHighRisk = rows.filter((item: any) => item.risk !== "安全").reduce((sum: number, item: any) => sum + item.inventory, 0);
    page(slide, `${groupName} inventory`, group
      ? `${groupName} 庫存分析：金額、健康度與關鍵料號`
      : `${groupName} 本期無資料，保留分類監控頁`, 9 + groupIndex);
    if (!group) {
      text(slide, "本期無資料", .62, 2.25, 6.2, .8, { fontSize: 34, bold: true, color: C.muted });
      text(slide, `${data.meta.latestQuarter} 上傳資料未包含 ${groupName} 分類，系統未補造任何數字。後續季度如有資料將自動產生 KPI、健康度與風險清單。`,
        .62, 3.35, 9.7, .82, { fontSize: 18, color: C.muted, valign: "top" });
      source(slide);
      continue;
    }
    metric(slide, .62, 1.62, 2.82, "總庫存金額", `${wan(group.inventory).toLocaleString()} 萬元`,
      `占全部 ${totals.inventory ? pct(group.inventory / totals.inventory) : "—"}`, C.pine);
    metric(slide, 3.64, 1.62, 2.82, "生產需求金額", `${wan(group.demand).toLocaleString()} 萬元`,
      `覆蓋率 ${group.demand ? (group.inventory / group.demand).toFixed(2) : "—"}x`, C.sky);
    metric(slide, 6.66, 1.62, 2.82, "庫存健康度", rows.length ? pct(groupSafe / rows.length) : "—",
      `安全 ${groupSafe}／${rows.length} 件`, C.moss);
    metric(slide, 9.68, 1.62, 2.82, "高風險金額", `${wan(groupHighRisk).toLocaleString()} 萬元`,
      `平均水準 ${group.avgLevel.toFixed(2)} ${groupName === "LP3" ? "日" : "月"}`, C.clay);
    const riskCounts = ["缺料", "安全", "積壓", "呆滯"].map(risk => rows.filter((item: any) => item.risk === risk).length);
    slide.addImage({
      data: donutChart(
        ["缺料", "安全", "積壓", "呆滯"], riskCounts, [C.clay, C.moss, C.ochre, "7C3AED"],
        rows.length ? pct(groupSafe / rows.length) : "—", `${groupName} 健康度`,
      ),
      x: .62, y: 3.18, w: 6.15, h: 3.35,
    });
    const topRows = [...rows].sort((a: any, b: any) => b.inventory - a.inventory).slice(0, 4);
    text(slide, "高金額風險料號", 7.18, 3.28, 2.8, .3, { fontSize: 20, bold: true });
    topRows.forEach((item: any, index: number) => {
      const y = 3.82 + index * .62;
      text(slide, String(index + 1).padStart(2, "0"), 7.18, y, .42, .24, { fontSize: 11, bold: true, color: C.pine });
      text(slide, item.part, 7.72, y - .02, 2.1, .26, { fontSize: 15, bold: true });
      text(slide, `${item.risk}｜${wan(item.inventory).toLocaleString()} 萬`, 9.82, y - .02, 2.5, .26, {
        fontSize: 14, bold: true, align: "right", color: item.risk === "安全" ? C.moss : item.risk === "缺料" ? C.clay : C.ochre,
      });
      text(slide, item.action, 7.72, y + .27, 4.6, .22, { fontSize: 11, color: C.muted });
      slide.addShape(pptx.ShapeType.line, { x: 7.72, y: y + .52, w: 4.6, h: 0, line: { color: C.line, width: .6 } });
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Critical items", "關鍵料號依風險金額排序，直接連結處置措施", 12);
    const rows = [
      ...(data.tops?.積壓 ?? []).slice(0, 3),
      ...(data.tops?.缺料 ?? []).slice(0, 2),
      ...(data.tops?.呆滯 ?? []).slice(0, 1),
    ];
    slide.addTable([
      ["風險", "件號／件名", "分類", "庫存金額", "庫存水準", "建議措施"],
      ...rows.map((item: any) => [
        item.risk, `${item.part}\n${item.partName || "來源未提供件名"}`, item.rawGroup,
        `${wan(item.inventory).toLocaleString()} 萬`, `${Number(item.level).toFixed(2)} ${item.group === "LP3" ? "日" : "月"}`, item.action,
      ]),
    ], {
      x: .62, y: 1.62, w: 12.08, h: 4.82,
      border: { type: "solid", color: C.line, pt: .8 }, fill: C.paper, color: C.ink,
      fontFace: FONT, fontSize: 11, margin: .08, rowH: .68,
      colW: [.78, 2.22, 1.1, 1.45, 1.45, 2.4],
      valign: "middle", autoFit: false, bold: false,
      fillHeader: C.pine, colorHeader: C.paper, boldHeader: true,
    });
    text(slide, "件名未出現在來源資料時，簡報明確標示「來源未提供件名」。", .62, 6.58, 8.2, .24, { fontSize: 10, color: C.muted });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "QoQ analysis", "季度變化分別比較庫存與生產需求", 14);
    const labels = quarters.map((q: any) => `${String(data.meta.year).slice(-2)}/${q.quarter}`);
    slide.addImage({
      data: groupedBars(labels, [
        { name: "總庫存", values: quarters.map((q: any) => wan(q.totals.inventory)), color: C.pine },
        { name: "生產需求", values: quarters.map((q: any) => wan(q.totals.demand)), color: C.sky },
      ]),
      x: .62, y: 1.62, w: 7.7, h: 4.86,
    });
    text(slide, "季度差異", 8.72, 1.72, 2.2, .3, { fontSize: 20, bold: true });
    quarters.slice(1).forEach((quarter: any, index: number) => {
      const prior = quarters[index];
      const y = 2.28 + index * 1.25;
      const inventoryDelta = prior.totals.inventory ? (quarter.totals.inventory - prior.totals.inventory) / prior.totals.inventory : 0;
      const demandDelta = prior.totals.demand ? (quarter.totals.demand - prior.totals.demand) / prior.totals.demand : 0;
      text(slide, `${prior.quarter} → ${quarter.quarter}`, 8.72, y, 2.7, .28, { fontSize: 17, bold: true, color: C.pine });
      text(slide, `庫存 ${inventoryDelta >= 0 ? "↑" : "↓"} ${pct(Math.abs(inventoryDelta))}`, 8.72, y + .38, 1.9, .24, {
        fontSize: 14, bold: true, color: inventoryDelta >= 0 ? C.clay : C.moss,
      });
      text(slide, `需求 ${demandDelta >= 0 ? "↑" : "↓"} ${pct(Math.abs(demandDelta))}`, 10.55, y + .38, 1.72, .24, {
        fontSize: 14, bold: true, color: demandDelta >= 0 ? C.clay : C.moss,
      });
      text(slide, `金額增減 ${wan(quarter.totals.inventory-prior.totals.inventory).toLocaleString()} 萬元`, 8.72, y + .73, 3.55, .24, {
        fontSize: 14, color: inventoryDelta > 0 ? C.clay : C.moss,
      });
      slide.addShape(pptx.ShapeType.line, { x: 8.72, y: y + 1.04, w: 3.55, h: 0, line: { color: C.line, width: .7 } });
    });
    source(slide);
  }

  {
    const slide = pptx.addSlide();
    page(slide, "Plan variance & inventory impact", "計畫差異帶動備料調整與庫存風險", 15);
    const comparable = (data.plans?.months ?? []).filter((m: any) => m.previous != null && m.current != null);
    if (comparable.length) {
      slide.addImage({
        data: groupedBars(
          comparable.map((m: any) => m.month.replace("20", "")),
          [
            { name: "計畫增減台數", values: comparable.map((m: any) => Math.abs(m.variance ?? 0)), color: C.clay },
            { name: "備料調整金額（萬元）", values: comparable.map((m: any) => Math.abs(wan(m.materialAdjustment ?? 0))), color: C.ochre },
          ],
        ),
        x: .62, y: 1.62, w: 6.65, h: 4.8,
      });
    } else {
      text(slide, "目前沒有可比較的重疊月份", .62, 2.8, 6.65, .5, { fontSize: 22, bold: true, align: "center" });
    }
    text(slide, "管理行動", 7.82, 1.7, 2, .32, { fontSize: 20, bold: true });
    [
      ["上修", "顯示追加備料需求與缺料風險；計畫差異及備料調整均以紅色警示。"],
      ["下修", "顯示多餘備料與庫存增加風險；下降本身不視為綠色改善。"],
      ["不變", "顯示無明顯影響；僅確認為改善或風險降低時使用綠色。"],
    ].forEach((item, index) => {
      const y = 2.25 + index * 1.2;
      text(slide, item[0], 7.82, y, .82, .3, { fontSize: 14, bold: true, color: C.pine });
      text(slide, item[1], 8.83, y - .04, 3.48, .65, { fontSize: 14, valign: "top" });
      slide.addShape(pptx.ShapeType.line, { x: 7.82, y: y + .78, w: 4.48, h: 0, line: { color: C.line, width: .7 } });
    });
    text(slide, "請主管裁示", 7.82, 5.88, 1.48, .28, { fontSize: 14, bold: true, color: C.clay });
    text(slide, "計畫差異分析用於判斷備料調整，以及對後續庫存金額、缺料與高庫存風險的影響。", 9.12, 5.84, 3.18, .58, { fontSize: 12, bold: true });
    source(slide);
  }

  await pptx.writeFile({
    fileName: `${data.meta.model}_${data.meta.year}_${data.meta.quarters.join("-")}_庫存決策簡報_現代商務資訊圖表風.pptx`,
  });
}
