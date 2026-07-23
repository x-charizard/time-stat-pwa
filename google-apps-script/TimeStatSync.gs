/**
 * Time Stat PWA — Google Sheet 後端（Apps Script）
 *
 * ⚠ 重要：專案入面只可以有「一個」doPost / doGet。
 * 若左邊同時有 Code.gs + 其他 .gs，而兩邊都有 doPost，舊版會繼續回 unauthorized。
 * → 刪走／清空舊 Code.gs，只留呢份內容，儲存後「管理部署 → 新版本」。
 *
 * 用法：
 * 1) 將此檔內容貼到「綁定你張試算表」嘅 Apps Script 專案（試算表 → 擴充功能 → Apps Script）。
 * 2) Project settings → Script properties：
 *    - ALLOWED_EMAILS = xavierlichitau@gmail.com,xavierlichitau1995@gmail.com
 *    - GOOGLE_CLIENT_ID = 348329876798-nhvl3ppsckle1lv1r7u3vs2tb6pm3al0.apps.googleusercontent.com
 *    - API_TOKEN（可選緊急後門；唔好放公開前端）
 * 3) 部署 → 管理部署 → 編輯 → 版本「新版本」→ 部署（執行身分=我；存取權=任何人）。
 * 4) 編輯器跑 testAuthSetup()，Logger 應見 authApi=gis-v1 同 properties OK。
 *
 * 瀏覽器 POST：Content-Type: text/plain（body 仍然係 JSON），避免 CORS preflight。
 *
 * API（POST + text/plain，body 為 JSON）：
 * - { idToken, action:"load" }           → { ok:true, state:object, authApi:"gis-v1" }
 * - { idToken, action:"whoami" }         → { ok:true, email, authApi:"gis-v1" }
 * - { idToken, action:"migrateFormToDb" } → Form 合併寫入 TimeStatDB
 * - { idToken, state }                  → 寫入 TimeStatDB
 * 可選：{ token: API_TOKEN, ... } 作緊急後門（唔經前端）。
 *
 * 資料：工作表 "TimeStatDB"，一格 JSON；超長自動分 CHUNK。
 */

var DB_SHEET = "TimeStatDB";
var CHUNK_MARK = "TIME_STAT_CHUNKED_V1";
var CHUNK_SIZE = 45000;
var AUTH_API = "gis-v1";
var __FORM_IMPORT_DEBUG__ = null;

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function authFail_(error) {
  return jsonOut_({ ok: false, error: String(error || "unauthorized"), authApi: AUTH_API });
}

function authOkFields_(extra) {
  var o = { ok: true, authApi: AUTH_API };
  if (extra) {
    for (var k in extra) {
      if (Object.prototype.hasOwnProperty.call(extra, k)) o[k] = extra[k];
    }
  }
  return o;
}

/** 喺 Apps Script 編輯器按 ▶ 跑呢個，睇 Logger。 */
function testAuthSetup() {
  var props = PropertiesService.getScriptProperties();
  var cid = String(props.getProperty("GOOGLE_CLIENT_ID") || "").trim();
  var allowed = allowedEmailsList_();
  Logger.log("authApi=" + AUTH_API);
  Logger.log("has_verifyGoogleIdToken_=" + (typeof verifyGoogleIdToken_ === "function"));
  Logger.log("GOOGLE_CLIENT_ID_set=" + Boolean(cid) + " len=" + cid.length);
  Logger.log("ALLOWED_EMAILS=" + allowed.join(","));
  if (!cid) Logger.log("FAIL: missing GOOGLE_CLIENT_ID script property");
  if (!allowed.length) Logger.log("FAIL: missing ALLOWED_EMAILS script property");
  if (cid && allowed.length) Logger.log("OK: properties look ready — Redeploy New version if PWA still says unauthorized");
}

function allowedEmailsList_() {
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty("ALLOWED_EMAILS") || "");
  return raw
    .split(/[,;\s]+/)
    .map(function (s) {
      return String(s || "")
        .trim()
        .toLowerCase();
    })
    .filter(Boolean);
}

/** Base64url → 字串（解 JWT payload；簽名仍靠 tokeninfo 優先）。 */
function decodeJwtPayload_(jwt) {
  var parts = String(jwt || "").split(".");
  if (parts.length < 2) return null;
  var b64 = parts[1].replace(/-/g, "+").replace(/_/g, "/");
  while (b64.length % 4) b64 += "=";
  try {
    return JSON.parse(Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString());
  } catch (e) {
    return null;
  }
}

function fetchTokenInfo_(token) {
  // POST：避免超長 JWT 塞入 query string
  try {
    var res = UrlFetchApp.fetch("https://oauth2.googleapis.com/tokeninfo", {
      method: "post",
      contentType: "application/x-www-form-urlencoded",
      payload: { id_token: token },
      muteHttpExceptions: true,
      followRedirects: true,
    });
    var code = res.getResponseCode();
    var bodyText = String(res.getContentText() || "");
    if (code >= 200 && code < 300) {
      return { ok: true, info: JSON.parse(bodyText) };
    }
  } catch (ex) {
    /* fall through */
  }
  try {
    var res2 = UrlFetchApp.fetch(
      "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(token),
      { muteHttpExceptions: true, followRedirects: true }
    );
    var code2 = res2.getResponseCode();
    var body2 = String(res2.getContentText() || "");
    if (code2 >= 200 && code2 < 300) {
      return { ok: true, info: JSON.parse(body2) };
    }
    return { ok: false, error: "invalid_id_token", http: code2 };
  } catch (ex2) {
    return { ok: false, error: "tokeninfo_fetch_failed" };
  }
}

/**
 * 驗證 Google ID token（GIS credential）。
 * @returns {{ ok: boolean, email?: string, error?: string }}
 */
function verifyGoogleIdToken_(idToken) {
  var token = String(idToken || "").trim();
  if (!token) return { ok: false, error: "missing_id_token" };
  var props = PropertiesService.getScriptProperties();
  var clientId = String(props.getProperty("GOOGLE_CLIENT_ID") || "").trim();
  if (!clientId) return { ok: false, error: "missing_google_client_id" };
  var allowed = allowedEmailsList_();
  if (!allowed.length) return { ok: false, error: "missing_allowed_emails" };

  var info = null;
  var ti = fetchTokenInfo_(token);
  if (ti.ok) {
    info = ti.info;
  } else {
    // tokeninfo 失敗時：仍用 JWT payload 做 aud／email／exp 守門（個人 allowlist 用途）
    info = decodeJwtPayload_(token);
    if (!info) return { ok: false, error: ti.error || "invalid_id_token" };
  }

  var aud = String(info.aud || "").trim();
  if (aud !== clientId) return { ok: false, error: "aud_mismatch" };
  if (info.exp != null) {
    var now = Math.floor(new Date().getTime() / 1000);
    if (Number(info.exp) < now - 30) return { ok: false, error: "id_token_expired" };
  }
  var verified = String(info.email_verified || "").toLowerCase();
  // tokeninfo 通常有 email_verified；純 JWT decode 時 GIS 發嘅 token 多數係 true
  if (verified && verified !== "true" && verified !== "1") {
    return { ok: false, error: "email_not_verified" };
  }
  var email = String(info.email || "")
    .trim()
    .toLowerCase();
  if (!email) return { ok: false, error: "missing_email" };
  var hit = false;
  for (var i = 0; i < allowed.length; i++) {
    if (allowed[i] === email) {
      hit = true;
      break;
    }
  }
  if (!hit) return { ok: false, error: "email_not_allowed" };
  return { ok: true, email: email };
}

/** API_TOKEN 緊急後門（唔經公開前端）。 */
function authApiTokenOk_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  if (!expected) return false;
  return String(token || "") === String(expected);
}

/**
 * 統一授權：優先 idToken（Google 登入）；否則可選 API_TOKEN 後門。
 * @returns {{ ok: boolean, email?: string, via?: string, error?: string }}
 */
function authorizeRequest_(idToken, apiToken) {
  var id = String(idToken || "").trim();
  if (id) {
    var v = verifyGoogleIdToken_(id);
    if (v.ok) return { ok: true, email: v.email, via: "idToken" };
    return { ok: false, error: v.error || "unauthorized" };
  }
  if (authApiTokenOk_(apiToken)) return { ok: true, via: "apiToken" };
  return { ok: false, error: "unauthorized" };
}

function authFromGet_(e) {
  var idToken = e && e.parameter && e.parameter.idToken ? String(e.parameter.idToken) : "";
  var token = e && e.parameter && e.parameter.token ? String(e.parameter.token) : "";
  return authorizeRequest_(idToken, token);
}

/** @deprecated 舊名；改用 authFromGet_ / authorizeRequest_ */
function authToken_(e) {
  var a = authFromGet_(e);
  return { ok: a.ok, expected: "", token: "" };
}

function readStateFromSheet_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(DB_SHEET);
    if (!sh) return null;
    var top = String(sh.getRange(1, 1).getDisplayValue() || "");
    if (!top) return null;
    if (top === CHUNK_MARK) {
      var n = parseInt(String(sh.getRange(2, 1).getDisplayValue() || ""), 10);
      if (!n || n < 1) return null;
      var buf = "";
      for (var r = 0; r < n; r++) {
        buf += String(sh.getRange(3 + r, 1).getValue() || "");
      }
      return JSON.parse(buf);
    }
    return JSON.parse(top);
  } catch (ex) {
    return null;
  }
}

function writeStateToSheet_(obj) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DB_SHEET);
  if (!sh) sh = ss.insertSheet(DB_SHEET);
  sh.clearContents();
  var json = JSON.stringify(obj);
  if (json.length <= CHUNK_SIZE) {
    sh.getRange(1, 1).setValue(json);
    return;
  }
  sh.getRange(1, 1).setValue(CHUNK_MARK);
  var parts = [];
  for (var i = 0; i < json.length; i += CHUNK_SIZE) {
    parts.push(json.slice(i, i + CHUNK_SIZE));
  }
  sh.getRange(2, 1).setValue(parts.length);
  for (var j = 0; j < parts.length; j++) {
    sh.getRange(3 + j, 1).setValue(parts[j]);
  }
}

function mergeMigratedIntoState_(st, migrated) {
  var out = JSON.parse(JSON.stringify(st));
  if (!out.activities) out.activities = [];
  if (!out.events) out.events = [];
  if (!out.projectsRegistry) out.projectsRegistry = [];
  var nameToStId = {};
  for (var i = 0; i < out.activities.length; i++) {
    nameToStId[normalizeActivityKeyGs_(out.activities[i].name)] = out.activities[i].id;
  }
  function ensureActivityId(formLabel) {
    var nk = normalizeActivityKeyGs_(formLabel);
    if (nameToStId[nk]) return nameToStId[nk];
    var nid = uid_();
    out.activities.push({ id: nid, name: String(formLabel).trim(), aliases: [] });
    nameToStId[nk] = nid;
    return nid;
  }
  var midToName = {};
  for (var a = 0; a < migrated.activities.length; a++) {
    midToName[migrated.activities[a].id] = migrated.activities[a].name;
  }
  for (var e = 0; e < migrated.events.length; e++) {
    var me = migrated.events[e];
    var nm = midToName[me.activityId];
    if (!nm) continue;
    var newAid = ensureActivityId(nm);
    var copy = JSON.parse(JSON.stringify(me));
    copy.id = uid_();
    copy.activityId = newAid;
    out.events.push(copy);
  }
  dedupeStateEventsByImportKeyGs_(out);
  return out;
}

function readStateForClient_() {
  __FORM_IMPORT_DEBUG__ = { at: new Date().toISOString() };
  var st = readStateFromSheet_();
  var dbEv = st && typeof st === "object" && Array.isArray(st.events) ? st.events.length : 0;

  var pack = buildStateFromFormSheetPack_();
  var migrated = pack ? pack.state : null;
  var formEv = migrated && migrated.events ? migrated.events.length : 0;
  if (pack && pack.debug) {
    __FORM_IMPORT_DEBUG__ = Object.assign(__FORM_IMPORT_DEBUG__ || {}, pack.debug);
  }

  var props = PropertiesService.getScriptProperties();
  var mergeForm = String(props.getProperty("MERGE_FORM_WITH_DB") || "") === "1";

  if (formEv > 0 && dbEv === 0) {
    writeStateToSheet_(migrated);
    __FORM_IMPORT_DEBUG__.wroteDb = "form_only";
    return migrated;
  }
  if (formEv > 0 && dbEv > 0 && st) {
    var merged = mergeMigratedIntoState_(st, migrated);
    if (mergeForm) {
      writeStateToSheet_(merged);
      __FORM_IMPORT_DEBUG__.wroteDb = "merged_persisted";
    } else {
      __FORM_IMPORT_DEBUG__.wroteDb = "merged_return_only_not_written";
    }
    __FORM_IMPORT_DEBUG__.dbEventsBefore = dbEv;
    __FORM_IMPORT_DEBUG__.dbEventsAfter = merged.events.length;
    return merged;
  }
  if (st != null) return st;
  return { version: 3, activities: [], events: [], structure: [], projectsRegistry: [] };
}

function migrateFormToTimeStatDb_() {
  __FORM_IMPORT_DEBUG__ = { at: new Date().toISOString(), mode: "migrateFormToDb" };
  var pack = buildStateFromFormSheetPack_();
  var migrated = pack && pack.state ? pack.state : null;
  if (pack && pack.debug) {
    __FORM_IMPORT_DEBUG__ = Object.assign(__FORM_IMPORT_DEBUG__ || {}, pack.debug);
  }
  var formEv = migrated && migrated.events ? migrated.events.length : 0;
  if (formEv === 0) {
    return {
      ok: false,
      error: "form_zero_events",
      debug: (pack && pack.debug) || { reason: "no_pack" },
    };
  }
  var st = readStateFromSheet_();
  var dbEv = st && typeof st === "object" && Array.isArray(st.events) ? st.events.length : 0;
  var out;
  if (!st || dbEv === 0) {
    out = migrated;
  } else {
    out = mergeMigratedIntoState_(st, migrated);
  }
  if (!out.activities) out.activities = [];
  if (!out.events) out.events = [];
  if (!out.projectsRegistry) out.projectsRegistry = [];
  dedupeStateEventsByImportKeyGs_(out);
  writeStateToSheet_(out);
  __FORM_IMPORT_DEBUG__.wroteDb = "migrateFormToDb_forced";
  __FORM_IMPORT_DEBUG__.wroteEvents = out.events.length;
  return {
    ok: true,
    wroteEvents: out.events.length,
    priorDbEvents: dbEv,
    formEvents: formEv,
    debug: __FORM_IMPORT_DEBUG__,
  };
}

function migrateFormRowsToTimeStatDb() {
  var res = migrateFormToTimeStatDb_();
  Logger.log(JSON.stringify(res, null, 2));
  try {
    SpreadsheetApp.getUi().alert(
      res.ok ? "Time Stat：已寫入 TimeStatDB" : "Time Stat：未完成",
      JSON.stringify(
        res.ok
          ? { wroteEvents: res.wroteEvents, priorDbEvents: res.priorDbEvents, formEvents: res.formEvents }
          : res,
        null,
        2
      ),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log("getUi 不可用 → 睇 Logger");
  }
  return res;
}

function getFormSourceSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var custom = PropertiesService.getScriptProperties().getProperty("FORM_SOURCE_SHEET");
  if (custom) {
    var sh0 = ss.getSheetByName(String(custom).trim());
    if (sh0) return sh0;
  }
  var names = ["Form", "Form responses 1", "表單回應 1"];
  for (var i = 0; i < names.length; i++) {
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  var all = ss.getSheets();
  for (var j = 0; j < all.length; j++) {
    var nm = all[j].getName();
    var low = String(nm).toLowerCase();
    if (/^form responses/i.test(nm) || /^表單回應/.test(nm) || low === "form" || /\bform\b/i.test(nm)) {
      return all[j];
    }
  }
  return null;
}

function uid_() {
  return Utilities.getUuid();
}

function normalizeActivityKeyGs_(name) {
  return String(name || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseSlashDateTimeGs_(s) {
  var m = String(s)
    .trim()
    .match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return null;
  var p1 = parseInt(m[1], 10);
  var p2 = parseInt(m[2], 10);
  var y = parseInt(m[3], 10);
  var hh = parseInt(m[4], 10);
  var mi = parseInt(m[5], 10);
  var sec = m[6] != null ? parseInt(m[6], 10) : 0;
  function valid(dt, mo0, day) {
    if (isNaN(dt.getTime())) return null;
    if (dt.getFullYear() !== y || dt.getMonth() !== mo0 || dt.getDate() !== day) return null;
    return dt.toISOString();
  }
  if (p1 > 12) {
    return valid(new Date(y, p2 - 1, p1, hh, mi, sec), p2 - 1, p1);
  }
  if (p2 > 12) {
    return valid(new Date(y, p1 - 1, p2, hh, mi, sec), p1 - 1, p2);
  }
  var us = valid(new Date(y, p1 - 1, p2, hh, mi, sec), p1 - 1, p2);
  if (us) return us;
  return valid(new Date(y, p2 - 1, p1, hh, mi, sec), p2 - 1, p1);
}

function parseIsoLikeDateGs_(s) {
  var t = String(s).trim();
  var m = t.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (!m) return null;
  var y = parseInt(m[1], 10);
  var mo = parseInt(m[2], 10);
  var d = parseInt(m[3], 10);
  var hh = m[4] != null ? parseInt(m[4], 10) : 0;
  var mi = m[5] != null ? parseInt(m[5], 10) : 0;
  var sec = m[6] != null ? parseInt(m[6], 10) : 0;
  var dt = new Date(y, mo - 1, d, hh, mi, sec);
  if (isNaN(dt.getTime())) return null;
  return dt.toISOString();
}

function cellToIsoGs_(cell) {
  if (cell === null || cell === "") return null;
  if (Object.prototype.toString.call(cell) === "[object Date]" && !isNaN(cell.getTime())) {
    return cell.toISOString();
  }
  if (typeof cell === "number") {
    var epoch = new Date(1899, 11, 30);
    var ms = epoch.getTime() + cell * 86400000;
    var dNum = new Date(ms);
    if (!isNaN(dNum.getTime())) return dNum.toISOString();
  }
  var str = String(cell).trim();
  if (!str) return null;
  var iso = parseSlashDateTimeGs_(str);
  if (iso) return iso;
  iso = parseIsoLikeDateGs_(str);
  if (iso) return iso;
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed.toISOString();
  return null;
}

function buildHeadersFromRow_(row) {
  var nameCount = {};
  var headers = [];
  for (var i = 0; i < row.length; i++) {
    var raw = row[i];
    var base =
      raw == null || String(raw).trim() === ""
        ? "Column" + (i + 1)
        : String(raw)
            .replace(/\u3000/g, " ")
            .trim();
    var c = nameCount[base] || 0;
    nameCount[base] = c + 1;
    var key = c === 0 ? base : base + "__" + (c + 1);
    headers.push(key);
  }
  return headers;
}

function resolveTsAndActivityCols_(headers) {
  var props = PropertiesService.getScriptProperties();
  var forcedTs = props.getProperty("TIMESTAMP_COLUMN_HEADER");
  var forcedAct = props.getProperty("ACTIVITY_COLUMN_HEADER");
  var tsCol = "";
  var actCol = "";
  if (forcedTs) {
    var fts = String(forcedTs).trim();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim() === fts) {
        tsCol = headers[i];
        break;
      }
    }
  }
  if (forcedAct) {
    var fa = String(forcedAct).trim();
    for (var j = 0; j < headers.length; j++) {
      if (String(headers[j]).trim() === fa) {
        actCol = headers[j];
        break;
      }
    }
  }
  if (!tsCol) {
    for (var a = 0; a < headers.length; a++) {
      var h = String(headers[a]).trim();
      var hl = h.toLowerCase();
      if (
        hl === "timestamp" ||
        /^時間戳/.test(h) ||
        /^提交時間/.test(h) ||
        /^submitted/i.test(h)
      ) {
        tsCol = headers[a];
        break;
      }
    }
  }
  if (!actCol) {
    for (var b = 0; b < headers.length; b++) {
      var h2 = String(headers[b]).trim();
      var h2l = h2.toLowerCase();
      if (
        h2l === "activity" ||
        /^活動/.test(h2) ||
        /^what.*(activity|doing)/i.test(h2)
      ) {
        actCol = headers[b];
        break;
      }
    }
  }
  if (!tsCol || !actCol) return null;
  return { tsCol: tsCol, actCol: actCol };
}

function pickHeader_(headers, matchers) {
  for (var m = 0; m < matchers.length; m++) {
    var fn = matchers[m];
    for (var i = 0; i < headers.length; i++) {
      if (fn(headers[i])) return headers[i];
    }
  }
  return "";
}

function splitPeopleGs_(s) {
  return String(s)
    .split(/\s*[,，、]\s*/)
    .map(function (x) {
      return String(x).trim();
    })
    .filter(Boolean);
}

function projectIdByNameGs_(state, nm) {
  var t = String(nm || "").trim();
  if (!t || !state.projectsRegistry) return null;
  var tl = t.toLowerCase();
  for (var i = 0; i < state.projectsRegistry.length; i++) {
    var p = state.projectsRegistry[i];
    if (!p) continue;
    if (String(p.project || "").trim().toLowerCase() === tl) return p.projectId;
    if (String(p.projectId || "").trim().toLowerCase() === tl) return p.projectId;
  }
  return null;
}

function assignOptionalFormFieldsGs_(ev, row) {
  for (var key in row) {
    if (!Object.prototype.hasOwnProperty.call(row, key)) continue;
    var ks = String(key).trim();
    var ks2 = ks.replace(/\u3000/g, " ");
    var low2 = ks2.toLowerCase();
    if (/^place$/i.test(low2) || /^地點$/.test(ks2)) {
      if (!ev.place) ev.place = String(row[key] || "").trim();
      continue;
    }
    if (/^category$/i.test(low2) || /^類別$/.test(ks2)) {
      var cv = String(row[key] || "").trim();
      if (cv) ev.category = cv;
      continue;
    }
    if (/^remark$/i.test(low2) || /^備註$/.test(ks2) || /^notes$/i.test(low2)) {
      var rv = String(row[key] || "").trim();
      if (rv) ev.remark = ev.remark ? ev.remark + " · " + rv : rv;
    }
  }
}

function mergeImportCsvLooseFieldsGs_(ev, row, metaCols) {
  var skip = {};
  for (var i = 0; i < metaCols.length; i++) {
    if (metaCols[i]) skip[String(metaCols[i]).trim().toLowerCase()] = true;
  }
  for (var key in row) {
    if (!Object.prototype.hasOwnProperty.call(row, key)) continue;
    var kl = String(key).trim().toLowerCase();
    if (skip[kl]) continue;
    var nk = String(key).replace(/\u3000/g, " ").replace(/？/g, "?").trim();
    var nkl = nk.toLowerCase();
    if (/what\s+is\s+the\s+project/i.test(nkl)) {
      var v = String(row[key] || "").trim();
      if (!v) continue;
      if (!ev.projectsFromForm) ev.projectsFromForm = v;
      else if (ev.projectsFromForm.toLowerCase().indexOf(v.toLowerCase()) === -1) {
        ev.projectsFromForm = ev.projectsFromForm + " · " + v;
      }
      continue;
    }
    if (/^group$/i.test(nk)) {
      var gv = String(row[key] || "").trim().toLowerCase();
      if (gv === "work" || gv === "rest") {
        ev.group = gv === "work" ? "Work" : "Rest";
        ev.category = ev.group;
      }
      continue;
    }
    if (/remark|notes|comment|description|journal|memo|feedback|備註|備注|日記|心得|說明|説明|反思/i.test(nkl)) {
      var v3 = String(row[key] || "").trim();
      if (!v3) continue;
      var lowV = v3.toLowerCase();
      if (String(ev.remark || "").toLowerCase().indexOf(lowV) !== -1) continue;
      ev.remark = ev.remark ? ev.remark + " · " + v3 : v3;
    }
  }
}

function eventImportDedupeKeyGs_(ev) {
  var t = new Date(ev.start).getTime();
  if (isNaN(t)) return "__badtime:" + String(ev.id || "");
  var pl = String(ev.place || "").trim().toLowerCase();
  var rm = String(ev.remark || "").trim().toLowerCase();
  var pp = (ev.people || [])
    .map(function (p) {
      return String(p).trim().toLowerCase();
    })
    .filter(Boolean)
    .sort()
    .join(";");
  var pf = String(ev.projectsFromForm || "").trim().toLowerCase();
  return t + "|" + String(ev.activityId || "") + "|" + pl + "|" + rm + "|" + pp + "|" + pf;
}

function dedupeStateEventsByImportKeyGs_(state) {
  var arr = state.events.slice();
  arr.sort(function (a, b) {
    return new Date(a.start) - new Date(b.start);
  });
  var seen = {};
  var out = [];
  for (var i = arr.length - 1; i >= 0; i--) {
    var k = eventImportDedupeKeyGs_(arr[i]);
    if (seen[k]) continue;
    seen[k] = true;
    out.push(arr[i]);
  }
  out.reverse();
  state.events = out;
}

function buildStateFromFormSheetPack_() {
  var dbg = {
    sheet: null,
    dataRows: 0,
    tsCol: "",
    actCol: "",
    skipNoTs: 0,
    skipNoAct: 0,
    eventsParsed: 0,
    headerSample: "",
  };
  var sh = getFormSourceSheet_();
  if (!sh) {
    dbg.reason = "no_form_sheet";
    try {
      dbg.allSheetNames = SpreadsheetApp.getActiveSpreadsheet()
        .getSheets()
        .map(function (s) {
          return s.getName();
        })
        .join(" | ");
    } catch (ex2) {}
    return { state: null, debug: dbg };
  }
  dbg.sheet = sh.getName();
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    dbg.reason = "sheet_too_short";
    return { state: null, debug: dbg };
  }
  dbg.dataRows = values.length - 1;

  var props = PropertiesService.getScriptProperties();
  var forcedRow1 = parseInt(String(props.getProperty("FORM_HEADER_ROW") || ""), 10);
  var headerRowIdx = 0;
  var headers = [];
  var tsCol = "";
  var actCol = "";
  if (forcedRow1 >= 1 && forcedRow1 <= values.length) {
    headerRowIdx = forcedRow1 - 1;
    headers = buildHeadersFromRow_(values[headerRowIdx]);
    var resForced = resolveTsAndActivityCols_(headers);
    if (resForced) {
      tsCol = resForced.tsCol;
      actCol = resForced.actCol;
    }
    dbg.headerRowUsed = headerRowIdx + 1;
  } else {
    for (var tri = 0; tri <= Math.min(2, values.length - 1); tri++) {
      headers = buildHeadersFromRow_(values[tri]);
      var resTry = resolveTsAndActivityCols_(headers);
      if (resTry) {
        headerRowIdx = tri;
        tsCol = resTry.tsCol;
        actCol = resTry.actCol;
        dbg.headerRowUsed = tri + 1;
        break;
      }
    }
  }
  dbg.headerSample = headers.length ? headers.slice(0, 24).join(" | ") : "";
  function rowToObject(row) {
    var o = {};
    for (var i = 0; i < headers.length && i < row.length; i++) {
      o[headers[i]] = row[i];
    }
    return o;
  }

  dbg.tsCol = tsCol || "";
  dbg.actCol = actCol || "";
  if (!tsCol || !actCol) {
    dbg.reason = "missing_ts_or_activity_column";
    return { state: null, debug: dbg };
  }

  var placeCol = pickHeader_(headers, [
    function (h) {
      return /^Place__\d+$/.test(String(h));
    },
    function (h) {
      return String(h).indexOf("Place__") === 0;
    },
    function (h) {
      return h === "Place";
    },
    function (h) {
      return /^地點$/.test(String(h).trim());
    },
  ]);
  var catCol = pickHeader_(headers, [
    function (h) {
      return h === "Category";
    },
  ]);
  var projCol = pickHeader_(headers, [
    function (h) {
      return /^what\s+is\s+the\s+project/i.test(String(h).replace(/\u3000/g, " "));
    },
    function (h) {
      return h === "Projects";
    },
    function (h) {
      return h === "What is the Projects?";
    },
    function (h) {
      return /^What is the Projects\?__\d+$/i.test(String(h));
    },
    function (h) {
      return /^Projects__\d+$/i.test(String(h));
    },
  ]);

  var state = { version: 3, activities: [], events: [], structure: [], projectsRegistry: [] };
  var labelKeyToId = {};

  function getOrCreateActivityId(label) {
    var t = String(label || "").trim();
    if (!t) return null;
    var nk = normalizeActivityKeyGs_(t);
    if (labelKeyToId[nk]) return labelKeyToId[nk];
    var id = uid_();
    state.activities.push({ id: id, name: t, aliases: [] });
    labelKeyToId[nk] = id;
    return id;
  }

  var importKeySeen = {};
  var metaCols = [tsCol, actCol, placeCol, catCol, projCol].filter(Boolean);

  for (var r = headerRowIdx + 1; r < values.length; r++) {
    var row = rowToObject(values[r]);
    var rawTs = row[tsCol];
    var iso = cellToIsoGs_(rawTs);
    if (!iso) {
      dbg.skipNoTs++;
      continue;
    }
    var actLabel = String(row[actCol] || "").trim();
    if (!actLabel) {
      dbg.skipNoAct++;
      continue;
    }
    var aid = getOrCreateActivityId(actLabel);
    if (!aid) continue;

    var ev = { id: uid_(), start: iso, activityId: aid };
    if (placeCol && row[placeCol]) ev.place = String(row[placeCol]).trim();
    if (catCol && row[catCol]) ev.category = String(row[catCol]).trim();
    assignOptionalFormFieldsGs_(ev, row);
    mergeImportCsvLooseFieldsGs_(ev, row, metaCols);

    var w2 = row["With__2"];
    var w0 = row["With"];
    if (w2 && String(w2).trim()) ev.people = splitPeopleGs_(w2);
    else if (w0 && String(w0).trim()) ev.people = splitPeopleGs_(w0);

    if (projCol && row[projCol] != null) {
      var pv = String(row[projCol]).trim();
      if (pv) {
        if (ev.projectsFromForm) {
          if (ev.projectsFromForm.toLowerCase().indexOf(pv.toLowerCase()) === -1) {
            ev.projectsFromForm = ev.projectsFromForm + " · " + pv;
          }
        } else {
          ev.projectsFromForm = pv;
        }
        var toks = String(ev.projectsFromForm || "")
          .split(/\s*[·,，、]\s*/)
          .map(function (x) {
            return String(x).trim();
          })
          .filter(Boolean);
        var firstTok = toks[0];
        if (firstTok) {
          var pid = projectIdByNameGs_(state, firstTok);
          if (pid) ev.projectId = pid;
        }
      }
    }

    var dk = eventImportDedupeKeyGs_(ev);
    if (dk && String(dk).indexOf("__badtime:") !== 0 && importKeySeen[dk]) continue;
    if (dk && String(dk).indexOf("__badtime:") !== 0) importKeySeen[dk] = true;
    state.events.push(ev);
  }

  dbg.eventsParsed = state.events.length;
  if (!state.events.length) {
    dbg.reason = "zero_rows_after_parse";
    return { state: null, debug: dbg };
  }
  dedupeStateEventsByImportKeyGs_(state);
  dbg.eventsAfterDedupe = state.events.length;
  dbg.reason = "ok";
  return { state: state, debug: dbg };
}

function doGet(e) {
  var a = authFromGet_(e);
  if (!a.ok) return authFail_(a.error || "unauthorized");

  var action = e && e.parameter && e.parameter.action ? String(e.parameter.action) : "";
  if (action === "whoami") {
    return jsonOut_(authOkFields_({ email: a.email || "", via: a.via || "" }));
  }
  if (action === "load") {
    try {
      var st = readStateForClient_();
      var outG = authOkFields_({ state: st });
      if (String(PropertiesService.getScriptProperties().getProperty("TIME_STAT_DEBUG") || "") === "1") {
        outG._formImportDebug = __FORM_IMPORT_DEBUG__ || {};
      }
      return jsonOut_(outG);
    } catch (err) {
      return authFail_(String(err && err.message ? err.message : err));
    }
  }

  if (action === "migrateFormToDb") {
    try {
      var resGm = migrateFormToTimeStatDb_();
      var outGm = authOkFields_({ ok: resGm.ok });
      outGm.ok = resGm.ok;
      if (resGm.ok) {
        outGm.wroteEvents = resGm.wroteEvents;
        outGm.priorDbEvents = resGm.priorDbEvents;
        outGm.formEvents = resGm.formEvents;
        outGm.message = "Form merged into TimeStatDB";
      } else {
        outGm.error = resGm.error || "migrate_failed";
        outGm.debug = resGm.debug;
      }
      if (String(PropertiesService.getScriptProperties().getProperty("TIME_STAT_DEBUG") || "") === "1") {
        outGm._formImportDebug = resGm.debug || __FORM_IMPORT_DEBUG__ || {};
      }
      return jsonOut_(outGm);
    } catch (errGm) {
      return authFail_(String(errGm && errGm.message ? errGm.message : errGm));
    }
  }

  var sh0 = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var v = sh0.getRange("A1").getDisplayValue();
  return jsonOut_(authOkFields_({ a1: v }));
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (x) {
    return authFail_("bad_json");
  }
  var auth = authorizeRequest_(body && body.idToken, body && body.token);
  if (!auth.ok) return authFail_(auth.error || "unauthorized");

  if (String(body.action || "") === "whoami") {
    return jsonOut_(authOkFields_({ email: auth.email || "", via: auth.via || "" }));
  }

  if (String(body.action || "") === "load") {
    try {
      var st = readStateForClient_();
      var outP = authOkFields_({ state: st });
      if (String(PropertiesService.getScriptProperties().getProperty("TIME_STAT_DEBUG") || "") === "1") {
        outP._formImportDebug = __FORM_IMPORT_DEBUG__ || {};
      }
      return jsonOut_(outP);
    } catch (errL) {
      return authFail_(String(errL && errL.message ? errL.message : errL));
    }
  }

  if (String(body.action || "") === "migrateFormToDb") {
    try {
      var resP = migrateFormToTimeStatDb_();
      var outMg = authOkFields_({});
      outMg.ok = resP.ok;
      if (resP.ok) {
        outMg.wroteEvents = resP.wroteEvents;
        outMg.priorDbEvents = resP.priorDbEvents;
        outMg.formEvents = resP.formEvents;
        outMg.message = "Form merged into TimeStatDB";
      } else {
        outMg.error = resP.error || "migrate_failed";
        outMg.debug = resP.debug;
      }
      if (String(PropertiesService.getScriptProperties().getProperty("TIME_STAT_DEBUG") || "") === "1") {
        outMg._formImportDebug = resP.debug || __FORM_IMPORT_DEBUG__ || {};
      }
      return jsonOut_(outMg);
    } catch (errMg) {
      return authFail_(String(errMg && errMg.message ? errMg.message : errMg));
    }
  }

  if (!body.state || typeof body.state !== "object") return authFail_("missing_state");
  try {
    writeStateToSheet_(body.state);
    return jsonOut_(authOkFields_({}));
  } catch (err2) {
    return authFail_(String(err2 && err2.message ? err2.message : err2));
  }
}
