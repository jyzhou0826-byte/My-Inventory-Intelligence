import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  return worker.fetch(new Request("http://localhost/", { headers: { accept: "text/html" } }), {
    ASSETS: { fetch: async () => new Response("Not found", { status: 404 }) },
  }, { waitUntil() {}, passThroughOnException() {} });
}

test("server-renders the inventory analysis setup page", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);
  const html = await response.text();
  assert.match(html, /<title>RV 庫存決策儀表板<\/title>/i);
  assert.match(html, /My Inventory Intelligence/);
  assert.match(html, /正式分析模式/);
  assert.match(html, /CSV/);
  assert.match(html, /XLSX/);
  assert.match(html, /執行分析/);
  assert.doesNotMatch(html, /codex-preview|Your site is taking shape|react-loading-skeleton/i);
});

test("ships production analysis without embedded company data", async () => {
  const source = await readFile(new URL("../app/page.tsx", import.meta.url), "utf8");
  assert.match(source, /Executive/);
  assert.match(source, /QoQ Analysis/);
  assert.match(source, /Plan Stability/);
  assert.match(source, /Material Group/);
  assert.match(source, /QuarterBars/);
  assert.match(source, /各分類庫存健康分析/);
  assert.match(source, /來源未提供/);
  assert.match(source, /QoQ 季度變化/);
  assert.match(source, /1\.0x 代表/);
  assert.match(source, /analyzeUploads/);
  assert.doesNotMatch(source, /dashboard\.json|seedData/);
});

test("includes production spreadsheet parsing and validation", async () => {
  const [analysis, page] = await Promise.all([
    readFile(new URL("../app/analysis.ts", import.meta.url), "utf8"),
    readFile(new URL("../app/page.tsx", import.meta.url), "utf8"),
  ]);
  assert.match(analysis, /TextDecoder\("big5"\)/);
  assert.match(analysis, /Power BI 必要欄位缺少/);
  assert.match(analysis, /【車型資料異常】/);
  assert.match(analysis, /XLSX\.read/);
  assert.match(analysis, /diagnose/);
  assert.match(analysis, /scopedFiles/);
  assert.match(analysis, /scope\.quarters\.includes/);
  assert.doesNotMatch(page, /q\[2\]\.totals/);
  assert.match(page, /單季基準/);
  assert.match(page, /HealthExplanation/);
  assert.match(page, /StatusDonut/);
  assert.match(page, /VarianceTrend/);
  assert.match(page, /safeDelta/);
  assert.match(analysis, /groupTotals/);
  assert.match(analysis, /groupMetrics/);
  assert.match(analysis, /sortKey/);
  assert.match(analysis, /String\(period\.year\)\.slice\(-2\)/);
  assert.match(page, /categories\.map\(group=>/);
  assert.match(page, /quarters\.forEach\(\(q:any,i:number\)=>/);
  assert.match(page, /series\.flatMap\(s=>s\.values\.slice\(1\)/);
  assert.doesNotMatch(page, /首季為 0% 基準|基準 0%/);
  assert.match(page, /0% 基準線/);
  assert.match(page, /slice\(0,20\)/);
  assert.match(page, /上季：/);
  assert.doesNotMatch(page, /顯示前 30 筆/);
});

test("ships accessible KPI and risk visualization styles", async () => {
  const css = await readFile(new URL("../app/globals.css", import.meta.url), "utf8");
  assert.match(css, /Microsoft JhengHei/);
  assert.match(css, /delta-bad/);
  assert.match(css, /status-donut/);
  assert.match(css, /variance-card/);
  assert.match(css, /risk-card \.table-wrap\{overflow:hidden/);
  assert.match(css, /risk-card table\{table-layout:fixed/);
  assert.match(css, /overflow-x:hidden/);
  assert.match(css, /repeat\(3,minmax\(0,1fr\)\)/);
  assert.match(css, /risk-card th,\.risk-card td\{white-space:nowrap/);
  assert.match(css, /word-break:keep-all/);
  assert.match(css, /variance-legend\.category-lines/);
});
