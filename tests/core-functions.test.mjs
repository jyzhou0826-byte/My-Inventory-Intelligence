import assert from "node:assert/strict";
import { after, test } from "node:test";
import { mkdir, readFile, rm, stat, writeFile } from "node:fs/promises";
import { File } from "node:buffer";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import * as XLSX from "xlsx";
import ts from "typescript";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const generated = [];

async function importTypeScript(relativePath) {
  const sourcePath = path.join(root, relativePath);
  const runtimePath = sourcePath.replace(/\.ts$/, `.test-runtime-${process.pid}.mjs`);
  const source = await readFile(sourcePath, "utf8");
  const result = ts.transpileModule(source, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
    },
    fileName: sourcePath,
  });
  await writeFile(runtimePath, result.outputText, "utf8");
  generated.push(runtimePath);
  return import(`${pathToFileURL(runtimePath).href}?v=${Date.now()}`);
}

function workbookFile(name, rows) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), "Sheet1");
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return new File([bytes], name, {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

function csvFile(name, rows) {
  const csv = rows.map(row => row.map(value => {
    const text = String(value ?? "");
    return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  }).join(",")).join("\n");
  return new File([`\uFEFF${csv}`], name, { type: "text/csv;charset=utf-8" });
}

const headers = [
  "車型", "件號", "件名", "廠別", "Material Group", "原因細項", "備註", "日期",
  "策略備料金額(人工維護)", "策略備料金額(系統計算)", "生產需求金額",
  "總庫存金額", "總庫存量", "淨庫存金額", "庫存水準",
];

const q1Rows = [
  headers,
  ["RV", "F-001", "海外件", "1100", "FORGN", "", "", "2026-01-01", 0, 0, 100, 200, 2, 200, 2],
  ["RV", "L-001", "國產件", "1100", "LP3", "", "", "2026-01-01", 0, 0, 0, 50, 10, 50, 0],
];

const q2Rows = [
  headers,
  ["RV", "F-001", "海外件", "1100", "FORGN", "減產", "", "2026-04-01", 0, 0, 100, 300, 3, 300, 3],
  ["RV", "L-002", "國產件", "1100", "LP3", "", "", "2026-04-01", 0, 0, 100, 0, 0, 0, 0],
  ["RV", "L-SUB", "子件", "1100", "LP3", "", "子件", "2026-04-01", 0, 0, 100, 9999, 10, 9999, 10],
];

const planPrevious = [["車型", "2026年5月", "2026年6月"], ["RV計", 100, 120]];
const planCurrent = [["車型", "2026年5月", "2026年6月"], ["RV計", 120, 120]];

let analysisResult;

test("SheetJS parses CSV and XLSX while preserving inventory rules", async () => {
  const { analyzeUploads } = await importTypeScript("app/analysis.ts");
  analysisResult = await analyzeUploads("RV", {
    inventory: [
      csvFile("RV_2026_Q1.csv", q1Rows),
      workbookFile("RV_2026_Q2.xlsx", q2Rows),
    ],
    previous: csvFile("Previous_Plan.csv", planPrevious),
    current: workbookFile("Current_Plan.xlsx", planCurrent),
  }, { years: [2026], quarters: ["Q1", "Q2"] });

  assert.deepEqual(analysisResult.meta.quarters, ["26/Q1", "26/Q2"]);
  assert.deepEqual(analysisResult.quarters.map(quarter => quarter.records), [2, 2]);
  assert.equal(analysisResult.quarters[0].counts["無需求庫存"], 1);
  assert.equal(analysisResult.latest.counts.積壓, 1);
  assert.equal(analysisResult.latest.counts.缺料, 1);
  assert.equal(analysisResult.latest.totals.inventory, 300);
  assert.equal(analysisResult.latestItems.some(item => item.part === "L-SUB"), false);
  assert.equal(analysisResult.plans.months.find(month => month.month === "2026-05").variance, 20);
});

test("PptxGenJS produces a real editable PPTX archive", async () => {
  assert.ok(analysisResult, "analysis fixture must run first");
  const { exportInventoryPptx } = await importTypeScript("app/pptx-export.ts");
  const output = path.join(root, `outputs/core-test-${process.pid}.pptx`);
  generated.push(output);
  await mkdir(path.dirname(output), { recursive: true });
  await exportInventoryPptx(analysisResult, output);
  const info = await stat(output);
  const bytes = await readFile(output);
  assert.ok(info.size > 50_000);
  assert.equal(bytes.subarray(0, 2).toString(), "PK");
  assert.ok(bytes.includes(Buffer.from("ppt/presentation.xml")));
});

after(async () => {
  await Promise.all(generated.map(file => rm(file, { force: true })));
});
