# 方向定案：待處理退費（帳冊式）

- 展示日期：2026-11-05
- 展示內容：`A-ledger.html`、`B-by-person.html`、`C-by-date.html` 三個方向，各附截圖（`shot-*.png`）
- 使用者選擇原話：「A」
- 定案方向：A 帳冊式。一列一筆、列高 38px、事由與日期退成灰階，紅色只留給「已退 100」；手機（≤720px）改為兩行堆疊，按鈕靠右。
- 實作位置：`index.html` 的 `.meal-notice` 樣式區塊與 `renderMeal()` 的清單 markup。
- 實作驗證：
  - `jscheck.py`（抽出 inline script 交 `node --check`）通過。
  - Playwright 量測：1280px 寬時列高 38px、表頭與資料欄位左緣對齊（學生欄 x=107）、按鈕不溢出面板；430px 寬時列高 88px、三行堆疊、無水平溢出。
  - `design_lint.py check` 無紅燈、無模板裝飾。
- 保留的既有行為：勾選批次退費（`data-pendpick`／`data-pendall`／`btnMealPayPick`）、`data-mealtake` 帶走、`data-mealpaid` 已退、展開／收起、全月標記已退。
