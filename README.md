# time-stat-pwa

手機優先 **PWA**：用「每一筆活動 **開始時間**」串起時間軸；下一筆嘅開始 = 上一段嘅結束。資料存在瀏覽器本機（localStorage），並可經 **Google Sign-In** 同步到 Google Sheet（TimeStatDB）。

公開站：`https://x-charizard.github.io/time-stat-pwa/`

## 本機預覽

```bash
python3 -m http.server 8765
```

開 `http://localhost:8765`（唔好用 `file://`）。本機測試 Google 登入時，OAuth Client 要加 `http://localhost:8765` 做 JavaScript origin。

## Google 登入（只允許你嘅 Gmail）

未登入唔會拉／寫雲端資料。允許名單只放喺 **Apps Script Script properties**，唔放公開前端。

### 1) Google Cloud Console

1. 建立 OAuth Client（類型：**網頁應用程式**）
2. **已授權嘅 JavaScript 來源**：
   - `https://x-charizard.github.io`
   - （本機）`http://localhost:8765`
3. 複製 **Client ID**

### 2) Apps Script（貼上 `google-apps-script/TimeStatSync.gs` 後）

Project settings → Script properties：

| Key | Value |
|-----|--------|
| `ALLOWED_EMAILS` | `xavierlichitau@gmail.com,xavierlichitau1995@gmail.com` |
| `GOOGLE_CLIENT_ID` | 上面嘅 Client ID |
| `API_TOKEN` | **換一個新密碼**（只作緊急後門；**唔好**放公開網／repo） |

然後重新部署「網頁應用程式」（新版本）。

### 3) PWA 設定 Client ID

二揀一：

- 喺 `app.js` 填 `GOOGLE_CLIENT_ID_DEFAULT = "….apps.googleusercontent.com"`
- 或 `config.remote.json`：

```json
{
  "execUrl": "",
  "googleClientId": "….apps.googleusercontent.com"
}
```

（`REMOTE_SYNC_BASE_DEFAULT` 已有 exec 網址時，可留空 `execUrl`。）

### 4) 安全收尾（重要）

舊版曾喺前端 bake `API_TOKEN`，視為**已洩露**：

1. Apps Script **換新** `API_TOKEN`
2. 確認公開 repo／Pages **冇**舊 token 字串
3. 正常路徑只認 Google **idToken**；訪客未登入／非允許名單 = `unauthorized`，無 state

## 自動連 Google Sheet

開網 → Sign in with Google（允許名單內）→ 自動 `load` TimeStatDB；之後本機 `save` 會帶 `idToken` POST 上雲。

優先序（exec／Client ID）：`app.js` 常數 → `config.remote.json` → localStorage。

## 資料格式（v2）

- `activities`（舊備份可能係 `entities`）＋ `events[].activityId`（舊備份可能係 `entityId`）。開 app 會自動遷移。

## Git / 部署

push `main` 會經 GitHub Actions 部署到 GitHub Pages。
