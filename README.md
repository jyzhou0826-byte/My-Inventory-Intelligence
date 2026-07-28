# My Inventory Intelligence

庫存決策分析平台，整合 Power BI 匯出資料、生產計畫、LP3／FORGN／KD 分類分析、季度趨勢與管理層 PPTX 匯出。

## 系統需求

- Node.js `>=22.13.0`
- pnpm

## 本機執行

```bash
pnpm install
pnpm run dev
```

## Production build

```bash
pnpm run build
```

## 主要功能

- CSV／XLSX Power BI 庫存資料匯入
- LP3、FORGN 與指定 KD 專案分類
- 庫存健康度、缺料、積壓與呆滯風險判定
- 季度總庫存金額趨勢分析
- PPTX 匯出採深藍管理層儀表板風格，並提供獨立 Python CLI：
  `python3 scripts/export_inventory_pptx.py --input app/data/dashboard.json --output inventory-report.pptx`
- 生產需求及 QoQ 趨勢
- 生產計畫差異、備料調整與庫存影響分析
- 管理層儀表板 PPTX 匯出

## 專案結構

- `app/`：Dashboard、分析邏輯及 PPTX 匯出
- `public/`：網站靜態資源
- `tests/`：回歸測試
- `worker/`：Sites production worker
- `.openai/hosting.json`：既有 Sites 專案設定

## 資料安全

環境變數、上傳檔案、建置輸出、暫存分析資料與本機套件快取均由 `.gitignore` 排除。
