# 月銷售排行產生器

## 這個工具在做什麼

整合多個電子書平台的原始銷售 Excel 報表，產生當月銷售排行，輸出為 Excel 文件。

使用者在網頁介面上傳各平台的原始報表，按下按鈕後自動整合，可下載結果。

## 文件結構

- `app.py` — Streamlit 網頁介面（上傳區、按鈕、下載）
- `engine.py` — 核心處理邏輯（解析各平台報表、整合、輸出 Excel）
- `requirements.txt` — 依賴套件
- `CLAUDE.md` — 本文件

## 技術棧

- Python
- Streamlit（網頁介面）
- openpyxl（讀寫 .xlsx）
- xlrd（讀 .xls 舊格式）

## 支援的平台與對應格式

| 平台 | 上傳幾個檔案 | 格式 |
|------|------------|------|
| Google Play | 1 個 | *稅務試算.xlsx |
| Kobo | 3 個 | TW_BOOKWALKER / TW_GAEABOOKS / TW_KADOKAWA |
| Readmoo | 1 個 | 台灣漫讀拆帳報表 YYYY-MM.xlsx |
| 博客來 | 2 個 | YYYYMM_bookwalker.xls + YYYYMM_bookwalker1.xls |
| 亂搭 | 1 個 | YYYYMM_台灣漫讀.xlsx |
| Pubu / 誠品 / PChome | 3 個 | .xls 格式 |

## 書種目錄（選填）

可上傳 `product_fixed.xlsx`（BW 書種目錄）來填入書種欄位。
對照表在 engine.py 的 CAT_MAP，目前支援：漫畫、輕小說、文學、實用書、雜誌。

## 輸入與輸出

- 輸入：各平台原始報表 + 月份（格式 YYYYMM） + 上個月的排行 Excel（選填）
- 輸出：`{YYYYMM}銷售排行_{MMDD}.xlsx`

## 目前狀態

[請在這裡更新目前的開發進度，例如：功能已完成、正在測試、準備新增 XX 功能]

## 下一步計劃

[請在這裡記錄計劃中但還沒做的功能]

## 重要規則

- 測試用的 Excel 報表不要 commit 到 Git（裡面可能有真實銷售數據）
- 如果未來加入 API 金鑰，不可以寫在代碼裡，要用環境變數
