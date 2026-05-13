# time-stat-pwa

手機優先 **PWA**：用「每一筆活動 **開始時間**」串起時間軸；下一筆嘅開始 = 上一段嘅結束。資料存在瀏覽器本機（localStorage），可 **匯出／還原 JSON**，亦可 **匯入 Google Form export 嘅 CSV**。

## 專案位置

本機路徑（注意 `Li's` 係 Unicode 撇號）：

`/Users/xavier/Desktop/Desktop - Li's MacBook Pro/PycharmProject/time-stat-pwa`

## 本機預覽

在專案目錄：

```bash
python3 -m http.server 8765
```

瀏覽器開 `http://localhost:8765`。**不要用 `file://` 開**，否則 Service Worker 可能唔郁。

## 免費上線（之後）

將整個資料夾拖到 [Netlify Drop](https://app.netlify.com/drop) 或 GitHub Pages，用 **HTTPS** 開；手機 Safari → 分享 → **加入主畫面**。

## 備份

定期用 App 內 **「備份」→ 匯出 JSON**；檔案可放本專案 `data/`（自行複製）。

## Spec / 命名 / alias（Obsidian）

規格、對照 Notion 用字等留喺 Obsidian，唔放喺此 repo 強制同步。

## 資料格式（v2）

- `activities`（舊備份可能係 `entities`）＋ `events[].activityId`（舊備份可能係 `entityId`）。開 app 會自動遷移。

## Git

```bash
cd "/Users/xavier/Desktop/Desktop - Li's MacBook Pro/PycharmProject/time-stat-pwa"
git status
```

（若 Terminal 打唔開條路徑，用 Python `os.listdir` 複製正確資料夾名。）


## Google Sheet 當 database（TimeStatDB）

1. 喺綁定試算表嘅 Apps Script 專案貼上 google-apps-script/TimeStatSync.gs，喺 Project settings → Script properties 新增 API_TOKEN，再部署「網頁應用程式」（exec 網址唔好帶 query）。
2. PWA → Import CSV 分頁最底：填 exec 網址同 token（會存本機 localStorage：timeStatRemoteSyncBase / timeStatRemoteSyncToken）。
3. 一次搬 Form → TimeStatDB：喺 Apps Script 執行 migrateFormRowsToTimeStatDb，或喺 PWA 按「Form→DB（migrate）」。
4. 按「拉取（load）」將 TimeStatDB JSON 同步落本機；之後喺 app 內改動會喺 save 時自動 POST state 上雲。
5. 表單欄名唔標準：設 FORM_SOURCE_SHEET、TIMESTAMP_COLUMN_HEADER、ACTIVITY_COLUMN_HEADER、FORM_HEADER_ROW；除錯設 TIME_STAT_DEBUG=1。


## 自動連 Google（唔使每台機再填 Import）

優先序：Import 頁／localStorage → 同目錄 **config.remote.json**（專案已帶空嘅 config.remote.json，可改填 execUrl／token 再部署；勿將真 token commit 上公開 repo，或用净 app.js 兩個 DEFAULT）→ app.js 內 **REMOTE_SYNC_BASE_DEFAULT**／**REMOTE_SYNC_TOKEN_DEFAULT**（只自己 build 嘅副本可先填呢兩個常數）。

開網會自動 fetch config.remote.json，再向 Apps Script 拉取 TimeStatDB；電腦同電話用**同一個公開 HTTPS 網址**就唔使逐台再撳 Import。
