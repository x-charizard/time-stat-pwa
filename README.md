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

## Git

```bash
cd "/Users/xavier/Desktop/Desktop - Li's MacBook Pro/PycharmProject/time-stat-pwa"
git status
```

（若 Terminal 打唔開條路徑，用 Python `os.listdir` 複製正確資料夾名。）
