# Time Stat AI Reports — 部署 checklist

## 1. Apps Script 專案

1. 貼上／更新（同一專案兩個檔）：
   - `TimeStatSync.gs`
   - `TimeStatAiReports.gs`
2. Script properties：
   - `GEMINI_API_KEY` =（你已 set）
   - 已有 `ALLOWED_EMAILS`、`GOOGLE_CLIENT_ID`
3. 儲存 → **管理部署 → 新版本**
4. 編輯器跑一次：`installAiReportTriggers()`
5. 乾跑（可選）：`testGenerateAiReportMonth()`（會寄信 + 寫入 `TimeStatAIReports`）

## 2. 自動排程

- 每月 1 號 ~03:10 HKT → 上一個曆月
- 季首日 ~03:20 → 上一季
- 1 月 1 日 ~03:30 → 上一年  
寄去 `ALLOWED_EMAILS`；寫入 PWA **AI Reports** 歷史。  
**唔寫 Obsidian。**

## 3. PWA

- Tab **AI Reports**：自動報告歷史
- 同一頁 **AI Settings**：改「點同 AI 溝通」「報告大綱」「額外要求」→ Save（存雲端，自動／人手共用）
- Report 頁 **AI Report**：人手 generate（唔入歷史）；可 **Email me**

## 4. 驗證

- Settings Save 後，人手 AI Report 語氣／章節有變
- 自動報告只 email + 歷史，無 Obsidian 檔
