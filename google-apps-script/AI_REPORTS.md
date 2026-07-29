# Time Stat AI Reports — 部署 checklist

## 部署最常踩嘅坑

編輯器儲存咗 **≠** 網上 `/exec` 已更新。一定要：

1. **部署 → 管理部署**
2. 撳**而家用緊嗰個** Web app 旁邊嘅鉛筆
3. **版本：新版本**（唔好淨係儲存檔案）
4. **部署**

PWA 連嘅 URL（`config.remote.json` / app 內 default）必須係呢個部署嘅 `/exec`。

專案要有兩個檔：
- `TimeStatSync.gs`（有 `doPost` + `aiPing`）
- `TimeStatAiReports.gs`（有 `handleGetAiSettings_` 等）

喺 PWA **AI Reports** 睇歷史；Refresh 重載列表同 Settings。  
若人手 Generate 仍 `missing_state` / `unknown_action` → 呢個 `/exec` 仲係舊版本。

1. 貼上／更新（同一專案兩個檔）：
   - `TimeStatSync.gs`
   - `TimeStatAiReports.gs`
2. Script properties：
   - `GEMINI_API_KEY` =（你已 set）
   - 已有 `ALLOWED_EMAILS`、`GOOGLE_CLIENT_ID`
3. 儲存 → **管理部署 → 新版本**
4. 編輯器跑一次：`installAiReportTriggers()`（會重裝含**星期六 07:00 週報**）
5. 乾跑（可選）：
   - `testGenerateAiReportMonth()`
   - `testGenerateAiReportWeek()`（會寄信 + 寫入 `TimeStatAIReports`）

## 2. 自動排程（Asia/Hong_Kong）

- **每星期六 ~07:00** → 上一個 ISO 週（Mon–Sun），key 如 `2026-W29`
- 每月 1 號 ~03:10 → 上一個曆月
- 季首日 ~03:20 → 上一季
- 1 月 1 日 ~03:30 → 上一年  

寄去 `ALLOWED_EMAILS`；寫入 PWA **AI Reports** 歷史。  
**唔寫 Obsidian。**

每份報告嘅 DATA_JSON 會附 `comparisons[]`（週：上 3 週；月／季／年：上 2 期），AI 大綱要求有「週期對比」章節（建議 Markdown table）。

## 3. PWA

- Tab **AI Reports**：自動報告歷史（Markdown 會渲染粗體／table）
- 同一頁 **AI Settings**：Roles and Rules／Structure／Topic of the Period／Temperature → Save
- Report 頁 **AI Report**：人手 generate（唔入歷史）；loading 期間報告框會一直顯示；可 **Email me**
- Report 日期若係完整 ISO 週（Mon–Sun）會以 `week` 送出

## 4. 驗證

- Settings Save 後，人手 AI Report 語氣／章節有變
- Generate 期間報告框顯示 Generating…，完成先換內容；`**粗體**` 同 table 可見
- 自動報告只 email + 歷史，無 Obsidian 檔
- 跑完 `installAiReportTriggers()` 後，觸發器列表有 Saturday 07:00 嗰條
