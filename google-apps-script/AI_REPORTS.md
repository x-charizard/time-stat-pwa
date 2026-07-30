# Time Stat AI Reports — 部署 checklist

## 部署最常踩嘅坑

編輯器儲存咗 **≠** 網上 `/exec` 已更新。一定要：

1. **部署 → 管理部署**
2. 撳**而家用緊嗰個** Web app 旁邊嘅鉛筆
3. **版本：新版本**（唔好淨係儲存檔案）
4. **部署**

PWA 連嘅 URL（`config.remote.json` / app 內 default）必須係呢個部署嘅 `/exec`。

專案要有三個檔：
- `TimeStatSync.gs`（有 `doPost` + emotion brief hook）
- `TimeStatAiReports.gs`
- `TimeStatAiPeriodKpis.gs`（KPI／periodConfig／情緒 brief）

喺 PWA **AI Reports** 睇歷史；Refresh 重載列表同 Settings。  
若人手 Generate 仍 `missing_state` / `unknown_action` → 呢個 `/exec` 仲係舊版本。

1. 貼上／更新（同一專案三個檔）：
   - `TimeStatSync.gs`
   - `TimeStatAiReports.gs`
   - `TimeStatAiPeriodKpis.gs`
2. Script properties：
   - `GEMINI_API_KEY` =（你已 set）
   - 已有 `ALLOWED_EMAILS`、`GOOGLE_CLIENT_ID`
3. 儲存 → **管理部署 → 新版本**
4. 編輯器跑一次：`installAiReportTriggers()`（含**星期六 07:00 週報**）
5. AI Settings：**Reset defaults → Save**（載入週／月／季／年 checklist 預設）
6. 乾跑（可選）：`testGenerateAiReportWeek()` / `testGenerateAiReportMonth()`

## 2. 自動排程（Asia/Hong_Kong）

- **每星期六 ~07:00** → 上一個 ISO 週（Mon–Sun）
- 每月 1 號 ~03:10 → 上一個曆月
- 季首日 ~03:20 → 上一季
- 1 月 1 日 ~03:30 → 上一年  

寄去 `ALLOWED_EMAILS`；寫入 PWA **AI Reports** 歷史。

## 3. 期別內容（可調）

AI Settings → **Period content**（Week／Month／Quarter／Year checklist + Notes；section labels 英文 + 定義）。  
Generate 只輸出已勾選章節；KPI 喺 `DATA_JSON.kpis`（含 `passFail`／targets）。  
Diffused Mode（前稱 DMN）：Reading／Friending 按 remark 判定；Photoing 唔計運動除非 remark 注明高強度。  
週期顯示用 `periodLabel`／`weekLabel`（月-日），唔用 W29。

人手 Report：日期範圍自動對應 `week`／`month`／`quarter`／`year`（掣旁有 badge）。

連續 **3 個同類型週期**對比（本期 + 上 2 期）。

## 4. 情緒 brief（即時）

Remark 含 negative keywords（可 Settings 改，預設含 chaos／頭痛／焦慮等）→ 雲端 sync 後 **即刻 email** 過去 72h 摘要。  
同一 wake-day 最多一封（dedupe）。

## 5. 驗證

- Settings Save 後，人手週／月報章節跟 checklist
- 揀 Mon–Sun → badge 顯示 `week · YYYY-Www`
- 打含 `chaos` 嘅 remark 並 sync → 應收到 Emotion brief
- 自動報告只 email + 歷史，無 Obsidian 檔
