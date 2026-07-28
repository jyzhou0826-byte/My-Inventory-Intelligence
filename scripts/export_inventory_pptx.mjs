#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import ts from "typescript";

const args = Object.fromEntries(process.argv.slice(2).reduce((pairs, value, index, all) => {
  if (value.startsWith("--")) pairs.push([value.slice(2), all[index + 1]]);
  return pairs;
}, []));

if (!args.input || !args.output) {
  console.error("Usage: node scripts/export_inventory_pptx.mjs --input dashboard.json --output report.pptx");
  process.exit(2);
}

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname), "..");
const sourcePath = path.join(root, "app", "pptx-export.ts");
const runtimePath = path.join(root, "app", ".pptx-export-runtime.mjs");

try {
  const source = await fs.readFile(sourcePath, "utf8");
  const result = ts.transpileModule(source, {
    compilerOptions: {
      target: ts.ScriptTarget.ES2022,
      module: ts.ModuleKind.ESNext,
      moduleResolution: ts.ModuleResolutionKind.Bundler,
    },
    fileName: sourcePath,
  });
  await fs.writeFile(runtimePath, result.outputText, "utf8");
  const { exportInventoryPptx } = await import(`${pathToFileURL(runtimePath).href}?v=${Date.now()}`);
  const data = JSON.parse(await fs.readFile(path.resolve(args.input), "utf8"));
  await exportInventoryPptx(data, path.resolve(args.output));
  console.log(path.resolve(args.output));
} finally {
  await fs.rm(runtimePath, { force: true });
}
