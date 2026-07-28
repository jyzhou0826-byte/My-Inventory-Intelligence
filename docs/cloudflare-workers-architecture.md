# Cloudflare Workers 正式部署架構稽核

## 結論

- 前端不是 Streamlit；目前為 React 19 + TypeScript + Vinext，並由 Vite 8 建置。
- Worker 入口為 `worker/index.ts`，使用 Cloudflare Workers Fetch API 與 `nodejs_compat`。
- Excel／CSV 在瀏覽器端由 SheetJS (`xlsx`) 解析，上傳檔案不送往伺服器。
- PPTX 在瀏覽器端由 PptxGenJS 產生，不需 Python runtime。
- GitHub `main` 是唯一正式版本來源；Push 到 `main` 後由 GitHub Actions 執行建置、測試與 Workers 部署。

## Python 套件稽核

| 套件 | 專案 runtime 使用狀態 | Workers 處理 |
|---|---|---|
| Streamlit | 未使用 | 不需替換 |
| pandas | 未使用 | 不需替換 |
| openpyxl | 未使用 | 由 SheetJS 取代 Excel 解析 |
| python-pptx | 未使用 | 由 PptxGenJS 取代 PPTX 匯出 |

`scripts/export_inventory_pptx.py` 只是可選的本機 CLI 入口，不會進入 Worker bundle，也不是正式網站 runtime 相依套件。

## Workers 相容性清單

| 套件／能力 | 狀態 | 說明 |
|---|---|---|
| React / React DOM | 已驗證 | 由 Vinext／Vite 建置 |
| Vinext | 已驗證 Build | 產生 Cloudflare Worker 相容輸出 |
| `@cloudflare/vite-plugin` | 已驗證 Build | 本機與 Worker 環境整合 |
| SheetJS `xlsx` | 需持續回歸 | 瀏覽器端解析 CSV／XLSX；已加入合成資料測試 |
| PptxGenJS | 需持續回歸 | 瀏覽器端匯出；已加入可開啟 ZIP/PPTX 測試 |
| `next/font/google` | 已驗證 Build | 建置環境需可取得或快取字型資源 |
| Drizzle ORM / D1 | 未啟用 | 目前沒有正式 D1 binding；啟用前需補 migration 與整合測試 |
| Cloudflare Images binding | 需正式帳號驗證 | `/_vinext/image` 依賴 `IMAGES`；目前頁面 SVG／PNG 可直接由 Assets 提供 |
| Node 相容 API | 已驗證 Build | Worker 使用 `nodejs_compat`；仍需以正式帳號做 smoke test |

## 資料安全

- Git 忽略所有 `.csv`、`.xls*`、`.pptx`、`.env*`、API Key／憑證與本機輸出。
- 不再將公司資料衍生的 `dashboard.json` 放入 Repository。
- 正式網站初始頁面不含公司庫存資料；資料只在使用者瀏覽器記憶體中分析。
- GitHub Actions 只從 Repository Secrets 讀取 Cloudflare Token 與 Account ID。

### 既有 Git 歷史注意事項

`app/data/dashboard.json` 曾存在於舊版提交。從目前版本刪除檔案不會自動清除
Git 歷史；若該衍生資料被公司定義為機密，必須在取得 Repository 歷史重寫
授權後，使用 `git filter-repo` 清除所有歷史版本、強制更新遠端分支與標籤，
並要求既有 checkout 重新 clone。

## 正式部署必要條件

1. GitHub Repository Secrets：
   - `CLOUDFLARE_API_TOKEN`
   - `CLOUDFLARE_ACCOUNT_ID`
2. Cloudflare API Token 至少包含目標帳號的 Workers Scripts Edit 權限。
3. 第一次正式部署後，驗證 `/`、CSV、XLSX、多季度分析與 PPTX 下載。
4. 公司內部使用時，另外以 Cloudflare Access 保護正式網域。
