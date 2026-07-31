(function () {
  const STORAGE_KEY = "time-stat-state-v1";
  const DAY_MS = 86400000;
  const MIN_TIMELINE_MS = 30 * 60000;
  /** Report Activity filter：選項來自最近 N 個曆日（含今日）有出現過嘅紀錄。 */
  const REPORT_ACTIVITY_FILTER_ROLLING_DAYS = 31;
  const MSG_NO_RECORDS = "No records";
  const MSG_PLEASE_INPUT_DATA = "Please Input Data";
  const MSG_ACTIVITY_REQUIRED = "請輸入 Activity 名稱";
  const MSG_PLACE_REQUIRED = "請先輸入地點（必填）";
  const MSG_MANUAL_NEED_FIELDS = "請揀好 Activity、日期同時間";
  const MSG_MANUAL_INVALID_TIME = "日期／時間唔有效";
  const MSG_LOG_NOW_DONE = "已記錄（而家）";
  const MSG_MANUAL_DONE = "已加入（後補）";

  /** @typedef {{ id: string, name: string, aliases: string[] }} Activity */
  /** @typedef {{ projectId: string, project: string }} ProjectRegistryItem */
  /** @typedef {{ id: string, start: string, activityId: string, remark?: string, people?: string[], place?: string, distractionSec?: number, category?: string, group?: string, layer?: string, cat?: string, subCat?: string, structureItem?: string, project?: string, projectId?: string, objective?: string, activityQuestion?: string, achievement?: string, improveLast?: string, importantElement?: string, detailsBetter?: string, action?: string, longTermGoals?: string, shortTermGoals?: string, miniGoals?: string, groupFromForm?: string, layersFromForm?: string, projectsFromForm?: string, categoriesFromForm?: string }} Event */

  function defaultState() {
    return {
      version: 3,
      activities: [],
      events: [],
      structure: [],
      projectsRegistry: [],
    };
  }

  function uid() {
    return crypto.randomUUID ? crypto.randomUUID() : "id-" + Date.now() + "-" + Math.random().toString(36).slice(2);
  }

  /** 兼容舊版 entities / entityId 備份 */
  const REMOTE_LS_BASE_KEY = "timeStatRemoteSyncBase";
  const REMOTE_LS_CLIENT_ID_KEY = "timeStatGoogleClientId";
  const AUTH_ID_TOKEN_KEY = "timeStatGoogleIdToken";
  const AUTH_EMAIL_KEY = "timeStatGoogleEmail";

  /**
   * Apps Script …/exec（公開無妨；真正守門係 Google 登入 + ALLOWED_EMAILS）。
   * Google OAuth Web Client ID（公開；喺 Cloud Console 建立後填入）。
   * 亦可由 config.remote.json 的 execUrl／googleClientId 覆寫。
   */
  const REMOTE_SYNC_BASE_DEFAULT =
    "https://script.google.com/macros/s/AKfycbx-FXrKsSY8lAFEQJHKHAaT2PPRZCW7B01UvpVtVEWf34eFijBLdnXltCfeFNGYYS_eIA/exec";
  /** 填入你嘅 OAuth Web Client ID，例如 123456789-xxxx.apps.googleusercontent.com */
  const GOOGLE_CLIENT_ID_DEFAULT =
    "348329876798-nhvl3ppsckle1lv1r7u3vs2tb6pm3al0.apps.googleusercontent.com";

  function getRemoteSyncBase() {
    try {
      if (typeof window !== "undefined" && window.__TIME_STAT_REMOTE_BASE__) {
        const w = String(window.__TIME_STAT_REMOTE_BASE__).trim();
        if (w) return w;
      }
    } catch (e2) {}
    const baked = String(REMOTE_SYNC_BASE_DEFAULT || "").trim();
    if (baked) return baked;
    try {
      const u = localStorage.getItem(REMOTE_LS_BASE_KEY);
      if (u && String(u).trim()) return String(u).trim();
    } catch (e) {}
    return "";
  }

  function getGoogleClientId() {
    const baked = String(GOOGLE_CLIENT_ID_DEFAULT || "").trim();
    if (baked) return baked;
    try {
      const c = localStorage.getItem(REMOTE_LS_CLIENT_ID_KEY);
      if (c && String(c).trim()) return String(c).trim();
    } catch (e) {}
    try {
      if (typeof window !== "undefined" && window.__TIME_STAT_GOOGLE_CLIENT_ID__) {
        const w = String(window.__TIME_STAT_GOOGLE_CLIENT_ID__).trim();
        if (w) return w;
      }
    } catch (e2) {}
    return "";
  }

  function getGoogleIdToken() {
    try {
      return String(sessionStorage.getItem(AUTH_ID_TOKEN_KEY) || "").trim();
    } catch (e) {
      return "";
    }
  }

  function getAuthEmail() {
    try {
      return String(sessionStorage.getItem(AUTH_EMAIL_KEY) || "").trim();
    } catch (e) {
      return "";
    }
  }

  function setAuthSession(idToken, email) {
    try {
      if (idToken) sessionStorage.setItem(AUTH_ID_TOKEN_KEY, String(idToken));
      else sessionStorage.removeItem(AUTH_ID_TOKEN_KEY);
      if (email) sessionStorage.setItem(AUTH_EMAIL_KEY, String(email));
      else sessionStorage.removeItem(AUTH_EMAIL_KEY);
    } catch (e) {}
  }

  function clearAuthSession() {
    setAuthSession("", "");
  }

  /** GIS id_token 通常 ~1 小時；過期就當未登入（避免一直用 stale token 打 sync） */
  function isIdTokenExpired_(token) {
    try {
      const parts = String(token || "").split(".");
      if (parts.length < 2) return true;
      const json = atob(parts[1].replace(/-/g, "+").replace(/_/g, "/"));
      const payload = JSON.parse(json);
      const exp = Number(payload.exp);
      if (!Number.isFinite(exp)) return true;
      // 60s 緩衝，避免邊界剛好過期
      return exp * 1000 <= Date.now() + 60000;
    } catch (e) {
      return true;
    }
  }

  function getFreshGoogleIdToken_() {
    const t = getGoogleIdToken();
    if (!t) return "";
    if (isIdTokenExpired_(t)) {
      clearAuthSession();
      return "";
    }
    return t;
  }

  function isSignedIn() {
    return Boolean(getFreshGoogleIdToken_());
  }

  /** 有 exec 網址就視為啟用遠端；實際請求要已 Google 登入。 */
  function useRemoteSync() {
    return Boolean(getRemoteSyncBase());
  }

  function canRemoteSync() {
    return Boolean(getRemoteSyncBase() && getFreshGoogleIdToken_());
  }

  try {
    if (typeof window !== "undefined") {
      window.__TIME_STAT_SYNC_STATUS__ = function () {
        return {
          useRemote: useRemoteSync(),
          canSync: canRemoteSync(),
          baseLen: getRemoteSyncBase().length,
          hasIdToken: Boolean(getGoogleIdToken()),
          clientIdLen: getGoogleClientId().length,
          authEmail: getAuthEmail() || "",
        };
      };
    }
  } catch (e) {}

  function getRemotePostUrl() {
    const b = getRemoteSyncBase();
    if (!b) return "";
    try {
      const u = new URL(b);
      u.search = "";
      return u.toString();
    } catch (e) {
      return "";
    }
  }

  function remoteAuthBody(extra) {
    const body = Object.assign({}, extra || {});
    const idToken = getFreshGoogleIdToken_();
    if (idToken) body.idToken = idToken;
    return body;
  }

  function handleRemoteUnauthorized_(errMsg) {
    const m = String(errMsg || "");
    if (
      m === "unauthorized" ||
      m === "invalid_id_token" ||
      m === "id_token_expired" ||
      m === "email_not_allowed" ||
      m === "aud_mismatch" ||
      m === "email_not_verified" ||
      m === "missing_id_token"
    ) {
      clearAuthSession();
      showAuthOverlay_(
        m === "id_token_expired" || m === "invalid_id_token"
          ? "Session expired. Please sign in again."
          : "Session expired or not allowed. Please sign in again."
      );
      return true;
    }
    return false;
  }

  /** 與 loadState／遠端 hydrate 共用：將任意 object 正規化成 app state。 */
  function normalizeStateFromParsed(o) {
    try {
      if (!o || typeof o !== "object") return null;
      const activities = Array.isArray(o.activities)
        ? o.activities
        : Array.isArray(o.entities)
          ? o.entities
          : null;
      if (!activities || !Array.isArray(o.events)) return null;
      let anyIdFixed = false;
      const events = o.events.map((ev) => {
        const n = { ...ev };
        if (n.activityId == null && n.entityId != null) n.activityId = n.entityId;
        delete n.entityId;
        if (n.id == null || n.id === "") {
          n.id = uid();
          anyIdFixed = true;
        }
        return n;
      });
      const projectsRegistry = Array.isArray(o.projectsRegistry) ? o.projectsRegistry : [];
      const out = { version: o.version || 3, activities, events, structure: [], projectsRegistry };
      if (o.updatedAt != null && Number.isFinite(Number(o.updatedAt))) {
        out.updatedAt = Number(o.updatedAt);
      }
      return { out, anyIdFixed };
    } catch (e) {
      return null;
    }
  }

  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return defaultState();
      const parsed = normalizeStateFromParsed(JSON.parse(raw));
      if (!parsed) return defaultState();
      const { out, anyIdFixed } = parsed;
      if (anyIdFixed) {
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(out));
        } catch (e) {
          /* ignore quota */
        }
      }
      return out;
    } catch (e) {
      return defaultState();
    }
  }

  let state = loadState();
  let reportPresetSuppress = false;
  let timelinePointerTipAbort = null;
  /** Timeline 揀嘅日子（YYYY-MM-DD）；預設今日 */
  let timelineCenterYmd = null;
  const timelineBlockDetailMap = new WeakMap();
  let _reportFiltersCachedGen = -1;
  let _reportRenderScheduled = false;
  /** Generate／渲染 AI 報告期間：暫緩重繪大型 Report，避免主線程卡死 */
  let _aiReportBusy_ = false;
  let _reportRenderPendingWhileAi_ = false;
  let _aiReportGenTimer_ = null;
  let _aiReportGenStartedAt_ = 0;
  let _historyInferCacheGen = -1;
  const _historyInferCache = new Map();

  const CLASSIFICATION_GROUP_OPTIONS = ["Work", "Rest"];
  const CLASSIFICATION_LAYER_OPTIONS = ["Health", "Freedom", "Achievement"];
  const CLASSIFICATION_CAT_OPTIONS = [
    "Mental Health",
    "Physical Health",
    "Time",
    "Finance",
    "Time Management",
    "Financial Management",
    "Business",
    "Art",
  ];
  const CLASSIFICATION_SUB_OPTIONS = [
    "Long Term",
    "Short Term",
    "Project",
    "non-project",
    "Needs-Review",
    "Xavier Li Photography",
  ];
  let _eventsMutationGen = 0;
  let _sortedUniqueCachedGen = -1;
  let _sortedUniqueCache = null;

  function bumpEventsMutationGen() {
    _eventsMutationGen++;
  }


  /**
   * 匯入／庫內去重鍵：同一毫秒開始 + 同一 Activity + 同一 Place + 同一 Remark + 同一 With + 同一 Projects 欄
   * → 視為同一筆（重覆匯入只保留時間序最後一筆）。
   */
  function eventImportDedupeKey(ev) {
    const t = new Date(ev.start).getTime();
    if (Number.isNaN(t)) return "__badtime:" + String(ev.id || "");
    const pl = String(ev.place || "").trim().toLowerCase();
    const rm = String(ev.remark || "").trim().toLowerCase();
    const pp = (ev.people || [])
      .map((p) => String(p).trim().toLowerCase())
      .filter(Boolean)
      .sort()
      .join(";");
    const pf = String(ev.projectsFromForm || "").trim().toLowerCase();
    return `${t}|${String(ev.activityId || "")}|${pl}|${rm}|${pp}|${pf}`;
  }

  /** 依去重鍵整庫去重，時間升序後由尾掃上嚟，每鍵只留最後一筆。 */
  function dedupeStateEventsByImportKey() {
    const arr = state.events.slice();
    arr.sort((a, b) => new Date(a.start) - new Date(b.start));
    const seen = new Set();
    const out = [];
    for (let i = arr.length - 1; i >= 0; i--) {
      const k = eventImportDedupeKey(arr[i]);
      if (seen.has(k)) continue;
      seen.add(k);
      out.push(arr[i]);
    }
    out.reverse();
    state.events = out;
    bumpEventsMutationGen();
  }

  (function dedupeLoadedEventsOnce() {
    const b = state.events.length;
    dedupeStateEventsByImportKey();
    if (state.events.length < b) save();
  })();

  let _remotePushInflight = null;
  let _remotePushAgain = false;
  let _remoteSyncStatus = ""; // "", "ok", "pending", "error"

  function setRemoteSyncStatus_(status, detail) {
    _remoteSyncStatus = status || "";
    const el = document.getElementById("remoteSyncHint");
    if (!el) return;
    if (!useRemoteSync()) {
      el.textContent = "";
      return;
    }
    if (!isSignedIn()) {
      el.textContent = "Cloud: signed out";
      return;
    }
    // 成功／就緒唔顯示（避免「Cloud: synced · N events」佔位）；只顯示同步中／失敗／未登入
    if (status === "pending") el.textContent = "Cloud: syncing…";
    else if (status === "ok") el.textContent = "";
    else if (status === "error")
      el.textContent = "Cloud: sync failed" + (detail ? " · " + detail : "");
    else el.textContent = "";
  }

  async function pushRemoteStateOnce_() {
    const url = getRemotePostUrl();
    if (!url || !canRemoteSync()) return { ok: false, error: "not_ready" };
    // 確保每次寫入有單調時間戳，避免並行／多裝置舊寫覆蓋新寫
    if (state.updatedAt == null || !Number.isFinite(Number(state.updatedAt))) {
      state.updatedAt = Date.now();
    }
    const payloadState = JSON.parse(JSON.stringify(state));
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(remoteAuthBody({ state: payloadState })),
      mode: "cors",
      cache: "no-store",
    });
    const j = await r.json().catch(() => ({}));
    if (j && j.ok === false) {
      if (String(j.error || "") === "stale_write" && j.state) {
        return { ok: false, error: "stale_write", serverState: j.state };
      }
      handleRemoteUnauthorized_(j.error);
      return { ok: false, error: (j && j.error) || "push_failed" };
    }
    if (!j || j.ok !== true) return { ok: false, error: (j && j.error) || "push_failed" };
    return { ok: true };
  }

  function mergeRemoteStateIntoLocal_(remoteObj) {
    const parsed = normalizeStateFromParsed(remoteObj);
    if (!parsed) return false;
    const remote = parsed.out;
    const byKey = new Map();
    for (const ev of remote.events || []) byKey.set(eventImportDedupeKey(ev), ev);
    for (const ev of state.events || []) {
      const k = eventImportDedupeKey(ev);
      if (!byKey.has(k)) byKey.set(k, ev);
    }
    const actById = new Map();
    for (const a of remote.activities || []) actById.set(a.id, a);
    for (const a of state.activities || []) {
      if (!actById.has(a.id)) actById.set(a.id, a);
    }
    const reg = new Set(
      []
        .concat(remote.projectsRegistry || [], state.projectsRegistry || [])
        .map((x) => String(x || "").trim())
        .filter(Boolean)
    );
    const remoteTs = Number(remoteObj && remoteObj.updatedAt);
    const localTs = Number(state.updatedAt);
    state = {
      version: remote.version || state.version || 3,
      activities: Array.from(actById.values()),
      events: Array.from(byKey.values()),
      structure: [],
      projectsRegistry: Array.from(reg),
      updatedAt: Math.max(
        Date.now(),
        Number.isFinite(remoteTs) ? remoteTs : 0,
        Number.isFinite(localTs) ? localTs : 0
      ),
    };
    bumpEventsMutationGen();
    dedupeStateEventsByImportKey();
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {}
    return true;
  }

  async function pushRemoteStateQuiet() {
    if (!canRemoteSync()) return;
    _remotePushAgain = true;
    setRemoteSyncStatus_("pending");
    if (_remotePushInflight) return _remotePushInflight;
    _remotePushInflight = (async () => {
      let lastErr = "";
      try {
        while (_remotePushAgain) {
          _remotePushAgain = false;
          try {
            const res = await pushRemoteStateOnce_();
            if (res.ok) {
              lastErr = "";
              continue;
            }
            if (res.error === "stale_write" && res.serverState) {
              mergeRemoteStateIntoLocal_(res.serverState);
              state.updatedAt = Date.now();
              try {
                localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
              } catch (e) {}
              _remotePushAgain = true;
              continue;
            }
            lastErr = res.error || "push_failed";
          } catch (e) {
            lastErr = (e && e.message) || "network";
          }
        }
      } finally {
        _remotePushInflight = null;
        if (_remotePushAgain) {
          void pushRemoteStateQuiet();
          return;
        }
        if (lastErr) {
          setRemoteSyncStatus_("error", lastErr);
          try {
            toast("Cloud sync failed: " + lastErr);
          } catch (e2) {}
        } else {
          setRemoteSyncStatus_("ok");
        }
      }
    })();
    return _remotePushInflight;
  }

  function save() {
    state.structure = [];
    state.updatedAt = Date.now();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    void pushRemoteStateQuiet();
  }

  function toast(msg) {
    const el = document.getElementById("toast");
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.remove("show"), 2200);
  }

  /** 必須人手撳 OK 先關閉（電話鍵盤唔會遮住；唔會自動 fade） */
  function showBlockingAlert(message) {
    try {
      if (document.activeElement && typeof document.activeElement.blur === "function") {
        document.activeElement.blur();
      }
    } catch (e) {}
    const overlay = document.getElementById("blockingAlert");
    const msg = document.getElementById("blockingAlertMsg");
    const btn = document.getElementById("blockingAlertOk");
    if (!overlay || !msg || !btn) {
      window.alert(String(message || ""));
      return;
    }
    msg.textContent = String(message || "");
    overlay.classList.remove("hidden");
    const close = () => {
      overlay.classList.add("hidden");
      btn.removeEventListener("click", close);
    };
    btn.addEventListener("click", close);
    try {
      btn.focus();
    } catch (e2) {}
  }

  function notifyGateFailure_(gate) {
    const msg = (gate && gate.message) || "請填 Remark／Reason";
    if (gate && gate.hardBlock) showBlockingAlert(msg);
    else toast(msg);
  }

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  }

  /** 手打日期 YYYY-MM-DD，唔靠原生 date picker */
  function parseYMDStrict(s) {
    const t = String(s || "").trim();
    const m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10);
    const d = parseInt(m[3], 10);
    const dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return null;
    return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
  }

  function activityById(activities, id) {
    return activities.find((e) => e.id === id);
  }

  function activityDisplayName(id) {
    const e = activityById(state.activities, id);
    return e ? e.name : "（已刪 Activity）";
  }

  function resolveActivityByLabel(label) {
    const t = String(label || "").trim();
    if (!t) return null;
    for (const e of state.activities) {
      if (e.name === t) return e;
      if (e.aliases && e.aliases.includes(t)) return e;
    }
    return null;
  }

  function getOrCreateActivity(name) {
    const t = String(name || "").trim();
    if (!t) return null;
    const existing = resolveActivityByLabel(t);
    if (existing) return existing;
    const e = { id: uid(), name: t, aliases: [] };
    state.activities.push(e);
    return e;
  }

  function sortedEvents() {
    return [...state.events].sort((a, b) => new Date(a.start) - new Date(b.start));
  }

  function todayYmd() {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return ymdFromLocalDate(d);
  }

  function yesterdayYmd() {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - 1);
    return ymdFromLocalDate(d);
  }

  function initTimelineDatePicker() {
    timelineCenterYmd = todayYmd();
    const pick = document.getElementById("timelinePickDate");
    if (pick) pick.value = timelineCenterYmd;
  }

  function getTimelineDisplayDays() {
    const pick = timelineCenterYmd || todayYmd();
    const today = todayYmd();
    const yesterday = yesterdayYmd();
    if (pick === today || pick === yesterday) {
      const out = [];
      for (let i = 2; i >= 0; i--) {
        const d = new Date();
        d.setHours(0, 0, 0, 0);
        d.setDate(d.getDate() - i);
        out.push({ date: d, ymd: ymdFromLocalDate(d) });
      }
      return out;
    }
    const anchor = parseYMDStrict(pick);
    if (!anchor) {
      timelineCenterYmd = todayYmd();
      return getTimelineDisplayDays();
    }
    const [yy, mo, da] = anchor.split("-").map(Number);
    const out = [];
    for (let off = -1; off <= 1; off++) {
      const d = new Date(yy, mo - 1, da);
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() + off);
      out.push({ date: d, ymd: ymdFromLocalDate(d) });
    }
    return out;
  }

  /** 時間軸可視三日欄最早一日 0:00（用於篩選跨日段） */
  function timelineViewCutoffMs() {
    const cols = getTimelineDisplayDays();
    if (!cols.length) return Date.now();
    return ymdDayBounds(cols[0].ymd).start;
  }

  /** 時間軸用：三日欄任何一刻有用到嘅紀錄（包含「開始喺視窗外但段尾跌入視窗」——跨日凌晨）。 */
  function timelineEventsAscending(fullAscOpt) {
    const fullAsc = fullAscOpt || sortedEventsUniqueById();
    const cutoff = timelineViewCutoffMs();
    return fullAsc.filter((ev) => {
      const t0 = new Date(ev.start).getTime();
      if (Number.isNaN(t0)) return false;
      if (t0 >= cutoff) return true;
      const idx = fullAsc.indexOf(ev);
      if (idx < 0) return false;
      const segMs = segmentDurationMsForReport(fullAsc, idx);
      const endMs = t0 + segMs;
      return endMs > cutoff;
    });
  }

  function durationMs(ev, nextEv) {
    if (!nextEv) return null;
    return new Date(nextEv.start) - new Date(ev.start);
  }

  /** 由升序列表建立「每筆 → 時間上下一筆」 */
  function chronologicalNextById(asc) {
    const m = new Map();
    for (let i = 0; i < asc.length - 1; i++) m.set(asc[i].id, asc[i + 1]);
    return m;
  }

  /** 同一毫秒開始嘅連續區間 [lo, hi]（用於攤分時長，避免下一筆同時間 → 0 ms）。 */
  function sameStartRunBounds(list, i) {
    const t0 = new Date(list[i].start).getTime();
    let lo = i;
    while (lo > 0 && new Date(list[lo - 1].start).getTime() === t0) lo--;
    let hi = i;
    while (hi + 1 < list.length && new Date(list[hi + 1].start).getTime() === t0) hi++;
    return { lo, hi, t0 };
  }

  /**
   * 報表／匯出用 segment 長度：由本筆開始到「下一個更遲嘅 start」之間嘅總毫秒，喺同時間戳嘅多筆之間**平均攤分**；
   * 最後一段無更遲嘅下一筆時用「而家 − start」再攤分（同一套邏輯）。
   */
  function segmentDurationMsForReport(list, i) {
    const { lo, hi, t0 } = sameStartRunBounds(list, i);
    const runLen = hi - lo + 1;
    const idxInRun = i - lo;
    let nextT = null;
    for (let j = hi + 1; j < list.length; j++) {
      const nt = new Date(list[j].start).getTime();
      if (nt > t0) {
        nextT = nt;
        break;
      }
    }
    if (nextT == null) {
      const open = Math.max(0, Date.now() - t0);
      const base = Math.floor(open / runLen);
      const rem = open - base * runLen;
      return base + (idxInRun < rem ? 1 : 0);
    }
    const span = nextT - t0;
    const base = Math.floor(span / runLen);
    const rem = span - base * runLen;
    return base + (idxInRun < rem ? 1 : 0);
  }

  /** 同一 `id` 出現多次（例如重覆匯入）時只保留時間序最後一筆，避免畫面重覆。 */
  function sortedEventsUniqueById() {
    if (_sortedUniqueCache !== null && _sortedUniqueCachedGen === _eventsMutationGen) {
      return _sortedUniqueCache;
    }
    const asc = sortedEvents();
    const out = [];
    const seen = new Set();
    for (let i = asc.length - 1; i >= 0; i--) {
      const id = asc[i].id;
      if (id != null && id !== "") {
        if (seen.has(id)) continue;
        seen.add(id);
      }
      out.push(asc[i]);
    }
    out.reverse();
    _sortedUniqueCache = out;
    _sortedUniqueCachedGen = _eventsMutationGen;
    return out;
  }

  function formatDur(ms) {
    if (ms == null || ms < 0) return "—";
    const s = Math.floor(ms / 1000);
    const h = Math.floor(s / 3600);
    const m = Math.floor((s % 3600) / 60);
    const sec = s % 60;
    if (h > 0) return `${h}:${String(m).padStart(2, "0")}:${String(sec).padStart(2, "0")}`;
    return `${m}:${String(sec).padStart(2, "0")}`;
  }

  function formatDurHours(ms) {
    if (ms == null || ms < 0) return "—";
    return (ms / 3600000).toFixed(2) + " h";
  }

  function formatHmLocal(input) {
    const d = input instanceof Date ? input : new Date(input);
    return d.toLocaleTimeString("zh-Hant", { hour: "2-digit", minute: "2-digit", hour12: false });
  }

  function timelineNowStatusText() {
    if (!state.events.length) return MSG_NO_RECORDS;
    const all = sortedEventsUniqueById();
    const cur = all[all.length - 1];
    const act = activityDisplayName(cur.activityId);
    const hmStart = formatHmLocal(cur.start);
    const hmNow = formatHmLocal(new Date());
    return `${act} (${hmStart} ~ ${hmNow})`;
  }

  /** Distraction 顯示：MM:SS（0 都顯示 00:00） */
  function formatDistractionMmSs_(sec) {
    const n = Math.max(0, Math.floor(Number(sec) || 0));
    const m = Math.floor(n / 60);
    const s = n % 60;
    return String(m).padStart(2, "0") + ":" + String(s).padStart(2, "0");
  }

  /** Timeline 浮層：顯示成段 activity（與報表同一 segment 計法），跨日亦係完整一段嘅時間／分鐘。 */
  function timelineBlockDetailText(ev, fullAsc) {
    const idx = fullAsc.indexOf(ev);
    if (idx < 0) return "";
    const segMs = segmentDurationMsForReport(fullAsc, idx);
    const t0 = new Date(ev.start).getTime();
    if (Number.isNaN(t0)) return "";
    const tEnd = t0 + segMs;
    const lines = [];
    lines.push(`${formatHmLocal(ev.start)} ~ ${formatHmLocal(tEnd)}`);
    lines.push(`${Math.round(segMs / 60000)} mins`);
    lines.push(`Distraction：${formatDistractionMmSs_(ev.distractionSec)}`);
    const rm = displayRemarkForRawRecord(ev);
    if (rm) lines.push(`Remark：${rm}`);
    return lines.join("\n");
  }

  /** Google Sheet style: MM/DD/YYYY HH:mm:ss */
  function parseSheetsTimestamp(cell) {
    if (cell == null) return null;
    const s = String(cell).trim();
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})$/);
    if (!m) return null;
    const mo = parseInt(m[1], 10);
    const d = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);
    const hh = parseInt(m[4], 10);
    const mi = parseInt(m[5], 10);
    const ss = parseInt(m[6], 10);
    const dt = new Date(y, mo - 1, d, hh, mi, ss);
    if (Number.isNaN(dt.getTime())) return null;
    return dt.toISOString();
  }

  function fillMergeSelects() {
    ["mergeFrom", "mergeTo"].forEach((id) => {
      const sel = document.getElementById(id);
      if (!sel) return;
      sel.innerHTML = "";
      state.activities.forEach((e) => {
        const o = document.createElement("option");
        o.value = e.id;
        o.textContent = e.name;
        sel.appendChild(o);
      });
    });
  }

  function refreshActivityDatalist() {
    // Place／Activity 建議改用自製 ▾ Top-3；唔再用原生 datalist（電話會多一隻箭咀）
  }

  function uniqueProjectsSorted() {
    const s = new Set();
    const reg = Array.isArray(state.projectsRegistry) ? state.projectsRegistry : [];
    for (let i = 0; i < reg.length; i++) {
      const p = String(reg[i].project || "").trim();
      if (p) s.add(p);
    }
    for (let i = 0; i < state.events.length; i++) {
      const p1 = String(state.events[i].projectsFromForm || "").trim();
      if (p1) {
        s.add(p1);
        p1.split(/\s*[·,，、]\s*/).forEach((x) => {
          const t = String(x || "").trim();
          if (t) s.add(t);
        });
      }
    }
    return [...s].sort((a, b) => a.localeCompare(b, "zh-Hant"));
  }

  function refreshProjectPickers() {
    // 保留空函式，兼容舊邏輯呼叫；而家 project 改為純系統建議，唔再手填。
    return uniqueProjectsSorted();
  }

  function ymdFromLocalDate(d) {
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return `${y}-${mo}-${day}`;
  }

  function minutesSinceMidnight(ms) {
    const d = new Date(ms);
    return d.getHours() * 60 + d.getMinutes() + d.getSeconds() / 60 + d.getMilliseconds() / 60000;
  }

  function ymdDayBounds(ymd) {
    const [yy, mo, da] = ymd.split("-").map(Number);
    const start = new Date(yy, mo - 1, da, 0, 0, 0, 0).getTime();
    return { start, endEx: start + DAY_MS };
  }

  /**
   * Timeline 專用：同一毫秒開始 + 同一 Activity 只保留最後一筆（常見於重覆匯入），避免兩格完全疊住。
   * 注意：Report 唔做呢層；若要同 Report 對齊，render 路徑唔應再用呢個去重。
   */
  function dedupeTimelineByStartAndActivity(asc) {
    if (!asc || asc.length < 2) return asc || [];
    const lastIdxByKey = new Map();
    for (let i = 0; i < asc.length; i++) {
      const ev = asc[i];
      const k = `${new Date(ev.start).getTime()}\t${String(ev.activityId)}`;
      lastIdxByKey.set(k, i);
    }
    const keep = new Set(lastIdxByKey.values());
    return asc.filter((_, i) => keep.has(i));
  }

  function timelineClearRoot(root) {
    if (!root) return;
    if (typeof root.replaceChildren === "function") {
      root.replaceChildren();
    } else {
      root.innerHTML = "";
    }
  }

  function timelineMountRoot(root, inner) {
    if (!root || !inner) return;
    if (typeof root.replaceChildren === "function") {
      root.replaceChildren(inner);
    } else {
      timelineClearRoot(root);
      root.appendChild(inner);
    }
  }

  function timelineIndexById_(fullAsc) {
    const m = new Map();
    for (let i = 0; i < fullAsc.length; i++) {
      const id = fullAsc[i] && fullAsc[i].id;
      if (id != null && id !== "" && !m.has(id)) m.set(id, i);
    }
    return m;
  }

  function timelinePassesMin(ev, fullAsc, idxMapOpt) {
    const idx =
      idxMapOpt && ev && ev.id != null && ev.id !== ""
        ? idxMapOpt.has(ev.id)
          ? idxMapOpt.get(ev.id)
          : -1
        : fullAsc.indexOf(ev);
    if (idx < 0) return false;
    return segmentDurationMsForReport(fullAsc, idx) >= MIN_TIMELINE_MS;
  }

  /**
   * 將一段 clip 入某日欄。
   * ≥30 分鐘只睇「成段」長度（同 Report 可對齊嘅門檻）；
   * 唔好再要求「當日切片」都 ≥30 分鐘——否則跨日／接近午夜嘅長段會喺某日欄消失，但 Report 仍計到。
   */
  function timelineDayClip(ev, colYmd, fullAsc, idxMapOpt) {
    const idx =
      idxMapOpt && ev && ev.id != null && ev.id !== ""
        ? idxMapOpt.has(ev.id)
          ? idxMapOpt.get(ev.id)
          : -1
        : fullAsc.indexOf(ev);
    if (idx < 0) return null;
    const segAll = segmentDurationMsForReport(fullAsc, idx);
    if (segAll < MIN_TIMELINE_MS) return null;
    const evStart = new Date(ev.start);
    const t0 = evStart.getTime();
    const t1Wall = t0 + segAll;
    const isCurrent = idx === fullAsc.length - 1;
    const { start: d0, endEx: d1 } = ymdDayBounds(colYmd);
    const segmentEnd = isCurrent ? Date.now() : t1Wall;
    const vs = Math.max(t0, d0);
    const ve = Math.min(segmentEnd, d1);
    if (ve <= vs) return null;
    const seg = ve - vs;
    return { vs, ve, t1: isCurrent ? null : t1Wall, t0, seg };
  }

  function renderTimeline() {
    const root = document.getElementById("timelineCalendar");
    const empty = document.getElementById("timelineEmpty");
    const nowStatus = document.getElementById("timelineNowStatus");
    const fullAsc = sortedEventsUniqueById();
    // 同 Report：唔再做 start+activity 額外去重（否則同秒同 activity 多筆會漏）
    const asc = timelineEventsAscending(fullAsc);
    const idxMap = timelineIndexById_(fullAsc);
    if (!root || !empty) return;
    timelineClearRoot(root);
    if (timelinePointerTipAbort) {
      timelinePointerTipAbort.abort();
      timelinePointerTipAbort = null;
    }
    if (nowStatus) nowStatus.textContent = timelineNowStatusText();
    if (!asc.length) {
      empty.classList.remove("hidden");
      empty.textContent = "No records in the selected 3-day window.";
      return;
    }
    empty.classList.add("hidden");

    const columns = getTimelineDisplayDays();

    const wk = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    const inner = document.createElement("div");
    inner.className = "timeline-cal-inner";

    const headRow = document.createElement("div");
    headRow.className = "timeline-cal-head";
    const corner = document.createElement("div");
    corner.className = "timeline-cal-corner";
    corner.setAttribute("aria-hidden", "true");
    headRow.appendChild(corner);
    for (const c of columns) {
      const h = document.createElement("div");
      h.className = "timeline-cal-day-title";
      const [, mm, dd] = c.ymd.split("-");
      h.textContent = `${mm} - ${dd} (${wk[c.date.getDay()]})`;
      headRow.appendChild(h);
    }
    inner.appendChild(headRow);

    const body = document.createElement("div");
    body.className = "timeline-cal-body";

    const yAxis = document.createElement("div");
    yAxis.className = "timeline-cal-y-axis";
    for (let hr = 0; hr < 24; hr++) {
      const lab = document.createElement("div");
      lab.className = "timeline-cal-hour";
      lab.textContent = String(hr).padStart(2, "0");
      yAxis.appendChild(lab);
    }
    body.appendChild(yAxis);

    const board = document.createElement("div");
    board.className = "timeline-cal-board";

    const stripeById = new Map();
    let slot = 0;
    for (const ev of asc) {
      const ok = timelinePassesMin(ev, fullAsc, idxMap);
      if (ok) stripeById.set(ev.id, slot % 2 === 0 ? "a" : "b");
      slot += ok ? 1 : 2;
    }

    for (const c of columns) {
      const col = document.createElement("div");
      col.className = "timeline-cal-day";
      const clips = [];
      for (let ei = 0; ei < asc.length; ei++) {
        const ev = asc[ei];
        const clip = timelineDayClip(ev, c.ymd, fullAsc, idxMap);
        if (!clip) continue;
        clips.push({ ev, clip });
      }
      clips.sort((a, b) => {
        const d = a.clip.vs - b.clip.vs;
        if (d !== 0) return d;
        return a.clip.seg - b.clip.seg;
      });
      for (let ci = 0; ci < clips.length; ci++) {
        const { ev, clip } = clips[ci];
        const { vs, ve, t1, t0, seg } = clip;
        const topPct = (minutesSinceMidnight(vs) / (24 * 60)) * 100;
        let hPct = (seg / DAY_MS) * 100;
        const isCurrent = t1 == null;
        if (hPct < 0.35) hPct = 0.35;
        if (isCurrent && hPct < 1.2) hPct = 1.2;

        const blk = document.createElement("div");
        blk.className = "timeline-cal-block";
        if (isCurrent) blk.classList.add("is-current");
        blk.dataset.stripe = stripeById.get(ev.id) || "a";
        blk.style.top = `${topPct}%`;
        blk.style.height = `${hPct}%`;
        blk.style.zIndex = isCurrent ? "12" : String(3 + ci);

        const title = document.createElement("div");
        title.className = "timeline-cal-block-title";
        title.textContent = activityDisplayName(ev.activityId);
        blk.appendChild(title);

        timelineBlockDetailMap.set(blk, timelineBlockDetailText(ev, fullAsc));

        col.appendChild(blk);
      }
      board.appendChild(col);
    }

    body.appendChild(board);
    inner.appendChild(body);
    timelineMountRoot(root, inner);
    timelinePointerTipAbort = new AbortController();
    bindTimelineHoverTip(root, timelinePointerTipAbort.signal);
  }

  function getTimelineTipEl() {
    let el = document.getElementById("timelineHoverTip");
    if (el) return el;
    el = document.createElement("div");
    el.id = "timelineHoverTip";
    el.className = "timeline-hover-tip hidden";
    document.body.appendChild(el);
    return el;
  }

  function showTimelineTip(text, x, y) {
    const tip = getTimelineTipEl();
    tip.textContent = text;
    tip.classList.remove("hidden");
    tip.style.left = `${x}px`;
    tip.style.top = `${y}px`;
    requestAnimationFrame(() => {
      const rect = tip.getBoundingClientRect();
      let nx = x;
      let ny = y - 8;
      if (rect.right > window.innerWidth - 10) nx = Math.max(10, window.innerWidth - rect.width - 10);
      if (rect.left < 10) nx = 10;
      if (rect.bottom > window.innerHeight - 10) ny = Math.max(10, window.innerHeight - rect.height - 10);
      if (rect.top < 10) ny = 10;
      tip.style.left = `${nx}px`;
      tip.style.top = `${ny}px`;
    });
  }

  function hideTimelineTip() {
    const tip = document.getElementById("timelineHoverTip");
    if (tip) tip.classList.add("hidden");
  }

  function bindTimelineHoverTip(root, signal) {
    if (!root || !signal) return;
    const move = (e) => {
      const blk = e.target && e.target.closest ? e.target.closest(".timeline-cal-block") : null;
      if (!blk || !root.contains(blk)) {
        hideTimelineTip();
        return;
      }
      const detail = timelineBlockDetailMap.get(blk);
      if (!detail) return;
      showTimelineTip(detail, e.clientX, e.clientY - 10);
    };
    root.addEventListener("pointermove", move, { signal });
    root.addEventListener("pointerleave", () => hideTimelineTip(), { signal });
  }

  function parseProjectsCsvRows(data) {
    return data
      .map((row) => ({
        projectId: String(row["Project ID"] || row.projectId || "").trim(),
        project: String(row.Projects || row.project || "").trim(),
      }))
      .filter((r) => r.projectId || r.project);
  }

  function projectIdByName(name) {
    const t = reportNormLabel(name).toLowerCase();
    if (!t) return "";
    const hit = (state.projectsRegistry || []).find((r) => reportNormLabel(r.project).toLowerCase() === t);
    return hit ? hit.projectId : "";
  }

  function suggestProjectsFromText(activityLabel, remark) {
    const remarkLow = reportNormLabel(remark).toLowerCase();
    if (!remarkLow) return [];
    const reg = Array.isArray(state.projectsRegistry) ? state.projectsRegistry : [];
    const hits = [];
    const seen = new Set();
    for (let i = 0; i < reg.length; i++) {
      const p = reportNormLabel(reg[i].project);
      const pl = p.toLowerCase();
      if (!p || seen.has(pl)) continue;
      if (remarkLow.includes(pl) || pl === remarkLow) {
        seen.add(pl);
        hits.push({
          project: p,
          projectId: reportNormLabel(reg[i].projectId) || projectIdByName(p),
        });
      }
    }
    return hits.slice(0, 2);
  }

  /** Transporting：按 vault 規則用「下一筆非 Transporting activity」推斷層／類（記錄當下後驗）。 */
  function inferTransportingLayerCatFromNext(evStartIso) {
    const asc = sortedEventsUniqueById();
    const t0 = new Date(evStartIso).getTime();
    if (Number.isNaN(t0)) return { layer: "Freedom", cat: "Time" };
    const nextMap = chronologicalNextById(asc);
    let follow = null;
    for (let i = 0; i < asc.length; i++) {
      if (new Date(asc[i].start).getTime() > t0) {
        follow = asc[i];
        break;
      }
    }
    let hops = 0;
    while (follow && hops < 8) {
      const fk = normalizeActivityKey(activityDisplayName(follow.activityId));
      if (fk !== "transporting") break;
      follow = nextMap.get(follow.id) || null;
      hops++;
    }
    if (!follow) return { layer: "Freedom", cat: "Time" };
    const nb = inferRulesLayerCatExcludeTransporting(follow);
    const nl = (nb.layer || "").trim().toLowerCase();
    if (nl === "health") return { layer: "Health", cat: nb.cat || "Mental Health" };
    const nbLayerLow = (nb.layer || "").toLowerCase();
    const cat = nb.cat && nb.cat !== "Needs-Review" && nbLayerLow !== "needs-review" ? nb.cat : "Time";
    return { layer: "Freedom", cat: cat };
  }

  function buildMappingCandidates(activityLabel, remark, evStartIso) {
    const suggested = suggestProjectsFromText(activityLabel, remark);
    let inferred = inferByHistoryOrHeuristic(activityLabel);
    const actKey = normalizeActivityKey(activityLabel);
    if (actKey === "transporting" && evStartIso) {
      const tr = inferTransportingLayerCatFromNext(evStartIso);
      inferred = { group: inferred.group, layer: tr.layer, cat: tr.cat };
    }
    const text = (String(activityLabel || "") + " " + String(remark || "")).toLowerCase();
    const inferSubByRules = (layer, hasProjectSignal) => {
      const layerN = reportNormLabel(layer).toLowerCase();
      if (layerN === "health") {
        const longTermHint = /(每日|每天|習慣|routine|長期|持續|keep|daily|habit)/.test(text);
        const shortTermHint = /(臨時|即刻|急救|今晚|今日|短期|一次|暫時|急性|panic|overload)/.test(text);
        if (longTermHint && !shortTermHint) return "Long Term";
        if (shortTermHint && !longTermHint) return "Short Term";
        return "Long Term";
      }
      if (layerN === "freedom") {
        return hasProjectSignal ? "Project" : "non-project";
      }
      return hasProjectSignal ? "Project" : "non-project";
    };
    const g0 = inferred.group || "Rest";
    const l0 = inferred.layer || "Health";
    const c0 = inferred.cat || "Mental Health";
    const hasProj = suggested.length > 0;
    const sp = hasProj ? suggested[0] : null;

    const cand1 = hasProj
      ? {
          label: "Suggestion 1",
          group: g0,
          layer: "Freedom",
          cat: "Time",
          subCat: "Project",
          activity: activityLabel,
          project: sp.project,
          projectId: sp.projectId || projectIdByName(sp.project),
        }
      : {
          label: "Suggestion 1",
          group: g0,
          layer: l0,
          cat: c0,
          subCat: inferSubByRules(l0, false),
          activity: activityLabel,
          project: "",
          projectId: "",
        };

    const cand2 = {
      label: "Suggestion 2",
      group: g0,
      layer: l0,
      cat: c0,
      subCat: "non-project",
      activity: activityLabel,
      project: "",
      projectId: "",
    };

    return [cand1, cand2];
  }

  function inferByHistoryOrHeuristic(activityLabel) {
    const key = normalizeActivityKey(activityLabel);
    for (let i = state.events.length - 1; i >= 0; i--) {
      const ev = state.events[i];
      const nm = normalizeActivityKey(activityDisplayName(ev.activityId));
      if (nm === key) {
        return {
          group: ev.group || ev.category || "Rest",
          layer: ev.layer || "Health",
          cat: ev.cat || "Mental Health",
        };
      }
    }
    const mentalRest = new Set(["resting", "familying", "walking", "meditating"]);
    const physicalRest = new Set(["sleeping", "showering", "fooding"]);
    const physicalSet = new Set(["gyming", "running", "yogaing", "exercise", "workouting", "hiking", "camping"]);
    const workSet = new Set(["trading", "trading practice", "trading planning", "programming", "obsidianing", "photoing", "photography", "planning", "reviewing", "reading"]);
    if (physicalSet.has(key)) return { group: "Rest", layer: "Health", cat: "Physical Health" };
    if (physicalRest.has(key)) return { group: "Rest", layer: "Health", cat: "Physical Health" };
    if (mentalRest.has(key)) return { group: "Rest", layer: "Health", cat: "Mental Health" };
    if (workSet.has(key)) return { group: "Work", layer: "Freedom", cat: "Time" };
    return { group: "Rest", layer: "Health", cat: "Mental Health" };
  }

  /** Raw 表 Project：<strong>只</strong>顯示 CSV／表單「What is the project…」等匯入嘅 <code>projectsFromForm</code>；唔用 <code>ev.project</code>。 */
  function displayProjectForRawRecord(ev) {
    return String(ev.projectsFromForm || "").trim();
  }

  /** Raw 表／篩選顯示用：Freedom 層內部推斷用 <code>Time</code> 時，畫面統一顯示為你 vault 定義表嘅 <strong>Time Management</strong>。 */
  function normalizeCatDisplayForRaw(cat) {
    const c = String(cat || "").trim();
    if (!c) return "";
    const low = c.toLowerCase();
    if (low === "time" || c === "Time") return "Time Management";
    return c;
  }

  /** 報表篩選／按 Cat 分桶：與 Raw 顯示一致（例如 <code>Time</code> → Time Management）。 */
  function effectiveReportCatKey(s) {
    const raw = reportNormLabel(s || "");
    return normalizeCatDisplayForRaw(raw) || raw;
  }

  /** 合併多個表單／備註欄，避免 CSV 用 Notes 等非「Remark」欄時畫面空白。 */
  function displayRemarkForRawRecord(ev) {
    const bits = [];
    const add = (s) => {
      const v = String(s || "").trim();
      if (!v) return;
      const low = v.toLowerCase();
      if (bits.some((b) => b.toLowerCase() === low)) return;
      bits.push(v);
    };
    add(ev.remark);
    add(ev.activityQuestion);
    add(ev.achievement);
    add(ev.improveLast);
    add(ev.detailsBetter);
    add(ev.importantElement);
    add(ev.objective);
    return bits.length ? bits.join(" · ") : "";
  }

  function ymdHmFromEventStart(iso) {
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "—";
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, "0");
    const da = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    return `${y}-${mo}-${da} ${hh}:${mi}`;
  }

  function durationMinutesLabel(ms) {
    if (ms == null || ms < 0) return "—";
    const mins = Math.round(ms / 60000);
    if (mins === 0 && ms > 0) return "<1 min";
    return `${mins} mins`;
  }

  function eventProjectLinksProjectsRegistry(ev) {
    const reg = Array.isArray(state.projectsRegistry) ? state.projectsRegistry : [];
    const pid = String(ev.projectId || "").trim();
    if (pid) {
      for (let i = 0; i < reg.length; i++) {
        if (String(reg[i].projectId || "").trim().toLowerCase() === pid.toLowerCase()) return true;
      }
    }
    const form = String(ev.projectsFromForm || "").trim().toLowerCase();
    if (form) {
      for (let i = 0; i < reg.length; i++) {
        const p = reportNormLabel(reg[i].project).toLowerCase();
        if (!p) continue;
        if (form === p || form.includes(p) || p.includes(form)) return true;
      }
    }
    return false;
  }

  function inferHealthSubCatFromRemark(remark) {
    const t = String(remark || "").toLowerCase();
    const longHint = /(每日|每天|習慣|routine|長期|持續|habit|daily)/.test(t);
    const shortHint = /(臨時|即刻|急救|今晚|今日|短期|一次|急性)/.test(t);
    if (longHint && !shortHint) return "Long Term";
    if (shortHint && !longHint) return "Short Term";
    return "Long Term";
  }

  function inferFreedomSubCatFromRules(ev, activityLabel, remark) {
    const form = String(ev.projectsFromForm || "").trim();
    const pid = String(ev.projectId || "").trim();
    if (form || pid) {
      if (eventProjectLinksProjectsRegistry(ev)) return "Project";
      return "non-project";
    }
    const sug = suggestProjectsFromText(activityLabel, String(remark || ""));
    if (!sug.length) return "non-project";
    const probe = { projectsFromForm: sug[0].project, projectId: sug[0].projectId || "" };
    if (eventProjectLinksProjectsRegistry(probe)) return "Project";
    return "non-project";
  }

  function inferRulesLayerCatExcludeTransporting(ev) {
    const activityLabel = activityDisplayName(ev.activityId);
    const key = normalizeActivityKey(activityLabel);
    const remarkLow = String(ev.remark || "").toLowerCase();

    const fixed = {
      photoing: { layer: "Freedom", cat: "Time" },
      photography: { layer: "Freedom", cat: "Finance" },
      trading: { layer: "Freedom", cat: "Finance" },
      "trading planning": { layer: "Freedom", cat: "Finance" },
      "trading practice": { layer: "Freedom", cat: "Time" },
      museuming: { layer: "Freedom", cat: "Time" },
    };
    if (fixed[key]) return fixed[key];

    if (key === "reading") {
      if (!String(ev.remark || "").trim()) return { layer: "Freedom", cat: "Time" };
      if (/(情緒|焦慮|內耗|平靜|放鬆|心境|冥想|安心)/.test(remarkLow)) return { layer: "Health", cat: "Mental Health" };
      if (/(交易|投資|變現|理財|倉位|策略|portfolio|損益|p&l)/i.test(remarkLow)) return { layer: "Freedom", cat: "Finance" };
      if (/(學習|技能|流程|系統|效率|筆記|教學|課程|how to)/i.test(remarkLow)) return { layer: "Freedom", cat: "Time" };
      return { layer: "Freedom", cat: "Time" };
    }

    const physicalRestSet = new Set(["sleeping", "showering", "fooding"]);
    const mentalRestSet = new Set(["resting", "familying", "walking", "meditating"]);
    const physicalSet = new Set(["gyming", "running", "yogaing", "exercise", "workouting", "hiking", "camping"]);
    if (physicalSet.has(key)) return { layer: "Health", cat: "Physical Health" };
    if (physicalRestSet.has(key)) return { layer: "Health", cat: "Physical Health" };
    if (mentalRestSet.has(key)) return { layer: "Health", cat: "Mental Health" };
    return { layer: "Freedom", cat: "Time" };
  }

  /** 舊 CSV／表單已入庫嘅 Work／Rest（<code>ev.group</code>／<code>ev.category</code>）；有就用，唔再估。 */
  function normalizeStoredWorkRest(ev) {
    const tryOne = (s) => {
      const t = String(s || "").trim().toLowerCase();
      if (t === "work" || t === "上班" || t === "工作") return "Work";
      if (t === "rest" || t === "休息" || t === "休閒") return "Rest";
      return "";
    };
    return tryOne(ev.group) || tryOne(ev.category) || tryOne(ev.groupFromForm);
  }

  function inferRulesWorkRestGroup(ev) {
    const fromCsv = normalizeStoredWorkRest(ev);
    if (fromCsv) return fromCsv;
    const activityLabel = activityDisplayName(ev.activityId);
    const key = normalizeActivityKey(activityLabel);
    const remark = String(ev.remark || "").trim();
    const remarkLow = remark.toLowerCase();
    const grey = new Set([
      "reading",
      "photoing",
      "writing",
      "obsidianing",
      "diarying",
      "gaming",
      "travel planning",
    ]);
    if (grey.has(key)) {
      if (!remark) return "Rest";
      const workHit = /(輸出|決策|練習|交付|學習|專注|實作|寫稿|修圖|code|coding|debug|趕工)/i.test(remarkLow);
      const restHit = /(放鬆|休息|行山|冥想|深呼吸|陪家人|hea|chill|瞓|睡眠|度假)/i.test(remarkLow);
      if (workHit && !restHit) return "Work";
      if (restHit && !workHit) return "Rest";
      if (workHit && restHit) return "Rest";
      return "Rest";
    }
    const restSetGrey = new Set(["resting", "sleeping", "showering", "fooding", "familying", "walking", "meditating"]);
    const physicalSetGrey = new Set(["gyming", "running", "yogaing", "exercise", "workouting", "hiking", "camping"]);
    if (physicalSetGrey.has(key) || restSetGrey.has(key)) return "Rest";
    return "Work";
  }

  const WAKE_TIME_KEY = "timeStatWakeTime";
  const WORK_CAP_HOURS_KEY = "timeStatWorkCapHours";
  const TRADING_CAP_HOURS_KEY = "timeStatTradingCapHours";
  const WORK_CUTOFF_HOUR_KEY = "timeStatWorkCutoffHour";
  const TRADING_ACTIVITY_KEYS = new Set(["trading", "trading practice", "trading planning"]);
  const REVIEWING_ACTIVITY_KEYS = new Set(["reviewing"]);
  const TRANSPORTING_ACTIVITY_KEYS = new Set(["transporting"]);
  const SOCIAL_NO_TRADE_KEYS = new Set(["friending", "familying", "socialing"]);
  const NO_TRADES_BANNER_MS = 2 * 3600000;
  const REVIEWING_EMAIL_MS = 30 * 60000;
  const MSG_GET_REST_AFTER_17 = "Get REST after 17:00.";
  const MSG_NO_TRADES_TODAY = "no Trades today";

  function getWakeTimeHm_() {
    try {
      const raw = String(localStorage.getItem(WAKE_TIME_KEY) || "03:00").trim();
      const m = raw.match(/^(\d{1,2}):(\d{2})$/);
      if (!m) return { h: 3, mi: 0, label: "03:00" };
      const h = Math.min(23, Math.max(0, parseInt(m[1], 10)));
      const mi = Math.min(59, Math.max(0, parseInt(m[2], 10)));
      return { h, mi, label: String(h).padStart(2, "0") + ":" + String(mi).padStart(2, "0") };
    } catch (e) {
      return { h: 3, mi: 0, label: "03:00" };
    }
  }

  /** 舊預設 06:30 → 03:00（只遷移明顯舊預設，唔改用戶自訂） */
  (function migrateWakeDefaultTo0300_() {
    try {
      const raw = localStorage.getItem(WAKE_TIME_KEY);
      if (raw == null || String(raw).trim() === "" || String(raw).trim() === "06:30") {
        localStorage.setItem(WAKE_TIME_KEY, "03:00");
      }
    } catch (e) {}
  })();

  /** Work soft-cap（超限要填 Reasons）；預設 4h */
  function getWorkCapMs_() {
    try {
      const raw = localStorage.getItem(WORK_CAP_HOURS_KEY);
      if (raw == null || raw === "") return 4 * 3600000;
      const n = Number(raw);
      if (Number.isFinite(n) && n > 0) return n * 3600000;
    } catch (e) {}
    return 4 * 3600000;
  }

  function getTradingCapMs_() {
    try {
      const raw = localStorage.getItem(TRADING_CAP_HOURS_KEY);
      if (raw == null || raw === "") return 2 * 3600000;
      const n = Number(raw);
      if (Number.isFinite(n) && n > 0) return n * 3600000;
    } catch (e) {}
    return 2 * 3600000;
  }

  /** 本地鐘：17:00（含）至翌日 wake 之前 = 硬擋 Work 時段 */
  function getWorkCutoffHour_() {
    try {
      // 注意：Number(null)===0，唔可以當有效 cutoff（否則全日硬擋）
      const raw = localStorage.getItem(WORK_CUTOFF_HOUR_KEY);
      if (raw == null || String(raw).trim() === "") return 17;
      const n = Number(raw);
      if (Number.isFinite(n) && n >= 0 && n <= 23) return Math.floor(n);
    } catch (e) {}
    return 17;
  }

  function isInWorkHardBlockWindow_(refInput) {
    const d = refInput instanceof Date ? refInput : new Date(refInput);
    if (Number.isNaN(d.getTime())) return false;
    const mins = d.getHours() * 60 + d.getMinutes();
    const cutoffMins = getWorkCutoffHour_() * 60;
    const wake = getWakeTimeHm_();
    const wakeMins = wake.h * 60 + wake.mi;
    // [cutoff, 24:00) ∪ [00:00, wake)
    if (mins >= cutoffMins) return true;
    if (mins < wakeMins) return true;
    return false;
  }

  /** 清醒日：[當日 wake, 翌日 wake)；若 ref 喺當日 wake 之前，屬前一個清醒日。 */
  function wakeDayBounds(refInput) {
    const d = refInput instanceof Date ? new Date(refInput.getTime()) : new Date(refInput);
    if (Number.isNaN(d.getTime())) {
      const now = new Date();
      return wakeDayBounds(now);
    }
    const { h, mi } = getWakeTimeHm_();
    const wakeToday = new Date(d.getFullYear(), d.getMonth(), d.getDate(), h, mi, 0, 0);
    let start;
    if (d.getTime() < wakeToday.getTime()) {
      start = new Date(wakeToday.getTime());
      start.setDate(start.getDate() - 1);
    } else {
      start = wakeToday;
    }
    const end = new Date(start.getTime());
    end.setDate(end.getDate() + 1);
    return { startMs: start.getTime(), endMs: end.getTime(), wakeLabel: getWakeTimeHm_().label };
  }

  function isTradingActivityEv_(ev) {
    const key = normalizeActivityKey(activityDisplayName(ev.activityId));
    return TRADING_ACTIVITY_KEYS.has(key);
  }

  function eventWorkRestGroupForCap_(ev) {
    const stored = normalizeStoredWorkRest(ev);
    if (stored) return stored;
    if (ev.group === "Work" || ev.group === "Rest") return ev.group;
    // soft cap：未標 Group 嘅舊紀錄唔當 Work，避免推斷過度放大工時
    return "";
  }

  function sumWakeDayMs_(refMs, predicate) {
    const bounds = wakeDayBounds(refMs);
    const list = sortedEventsUniqueById();
    let total = 0;
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const st = new Date(ev.start).getTime();
      if (Number.isNaN(st) || st < bounds.startMs || st >= bounds.endMs) continue;
      if (!predicate(ev)) continue;
      total += segmentDurationMsForReport(list, i);
    }
    return { ms: total, bounds };
  }

  function sumWorkMsInWakeDay(refMs) {
    return sumWakeDayMs_(refMs, (ev) => eventWorkRestGroupForCap_(ev) === "Work");
  }

  function sumTradingMsInWakeDay(refMs) {
    return sumWakeDayMs_(refMs, (ev) => isTradingActivityEv_(ev));
  }

  function activityKeyOfEv_(ev) {
    return normalizeActivityKey(activityDisplayName(ev.activityId));
  }

  function sumActivityKeysMsInWakeDay(refMs, keySet) {
    return sumWakeDayMs_(refMs, (ev) => keySet.has(activityKeyOfEv_(ev)));
  }

  function wakeDayKey_(refMs) {
    const b = wakeDayBounds(refMs);
    const d = new Date(b.startMs);
    return (
      d.getFullYear() +
      "-" +
      String(d.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(d.getDate()).padStart(2, "0")
    );
  }

  function formatCapClock_(ms) {
    const totalMin = Math.max(0, Math.floor(ms / 60000));
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    return String(h).padStart(2, "0") + ":" + String(m).padStart(2, "0");
  }

  // —— Xavier Energy Model 3.0（Focus Points / FP + Work Momentum）——
  const ENERGY_FP_CAP = 1000;
  const ENERGY_FP_FLOOR = -500;
  const ENERGY_HIGH_RATE = 4.17;
  const ENERGY_MED_RATE = 2.08;
  const ENERGY_LOW_RATE = 1.39;
  const ENERGY_HIGH_KEYS = new Set([
    "trading",
    "trading practice",
    "programming",
    "timing",
    "financing",
    "web",
    "webing",
    "web development",
    "system",
    "systeming",
    "system development",
    "apping",
    "app development",
  ]);
  const ENERGY_MED_KEYS = new Set([
    "reviewing",
    "planning",
    "trading planning",
    "mind mapping",
    "mindmapping",
    "aiing",
    "ai",
    "code",
    "obsidianing",
    "notioning",
    "reading",
  ]);
  const ENERGY_LOW_KEYS = new Set([
    "photoing",
    "photography",
    "photo editing",
    "photoediting",
    "editing",
    "transporting",
  ]);
  const ENERGY_RECOVERY = {
    meditating: { rate: 3.0, cap: 90 },
    gyming: { rate: 2.0, cap: 120 },
    yogaing: { rate: 2.0, cap: 120 },
    showering: { rate: 2.0, cap: 30 },
    resting: { rate: 1.5, cap: 120 },
    walking: { rate: 1.2, cap: 90 },
    hiking: { rate: 1.0, cap: 150 },
    fooding: { rate: 0.8, cap: 60 },
  };
  /** Sleep：Recovery_Points = (min/480)^1.5 * 1000（8h = 1000） */
  const ENERGY_SLEEP_FULL_MIN = 480;
  const ENERGY_HEAT_START_MIN = 60;
  const ENERGY_HEAT_STEP_MIN = 30;
  const ENERGY_HEAT_STEP_PCT = 0.1;
  const ENERGY_RECOVER_RESET_MIN = 15;
  const ENERGY_INERTIA_BREAK_MIN = 60;
  const ENERGY_INERTIA_HIGH_MIN = 20;
  const ENERGY_INERTIA_MULT = 1.3;
  const ENERGY_MIN_DAY_CAP = 100;
  /** 下一筆距離超過呢個 → 當漏打卡／過夜，非 Sleep 最多計呢啲分鐘（避免假 System Crash） */
  const ENERGY_ORPHAN_GAP_MIN = 4 * 60;
  const ENERGY_ORPHAN_CREDIT_CAP_MIN = 90;
  /** 單一 punch 非 Sleep 最多計入 FP 嘅分鐘（含「而家仲進行中」；對齊 High 4h 上限） */
  const ENERGY_MAX_SEGMENT_NON_SLEEP_MIN = 4 * 60;
  let _energyRefreshTimer_ = null;
  let _lastEnergySnap_ = null;

  function clampEnergyFp_(fp, maxCap) {
    const cap = maxCap != null ? maxCap : ENERGY_FP_CAP;
    return Math.max(ENERGY_FP_FLOOR, Math.min(cap, fp));
  }

  function energyDayDrainScale_(maxCap) {
    const cap = Math.max(1, maxCap != null ? maxCap : ENERGY_FP_CAP);
    return ENERGY_FP_CAP / cap;
  }

  /** Heat：連續 High/Med >60 分後，每多 30 分 +10% drain */
  function energyHeatMultiplier_(workStreakMin) {
    if (workStreakMin <= ENERGY_HEAT_START_MIN) return 1;
    const steps = Math.ceil((workStreakMin - ENERGY_HEAT_START_MIN) / ENERGY_HEAT_STEP_MIN);
    return 1 + ENERGY_HEAT_STEP_PCT * steps;
  }

  /** 指數 Sleep 累積恢復總分（對 sleepTotalMin） */
  function energySleepRecoveryTotal_(sleepMin) {
    const m = Math.max(0, sleepMin);
    if (m <= 0) return 0;
    return Math.pow(m / ENERGY_SLEEP_FULL_MIN, 1.5) * ENERGY_FP_CAP;
  }

  /** 小說／放鬆閱讀 → Recover；非小說／預設 reading → Medium Drain */
  function energyReadingIsLeisure_(ev) {
    const blob = String((ev && ev.remark) || "").toLowerCase();
    if (/小說|novel|fiction|harry\s*potter|漫畫|comic|輕小說|romance|thriller/.test(blob)) {
      return true;
    }
    if (
      /非小說|non-?fiction|原子習慣|trading\s*in\s*the\s*zone|成長|教科書|textbook|technical|論文/.test(
        blob,
      )
    ) {
      return false;
    }
    return false;
  }

  /** Fooding 期間有冇標明同 AI 傾偈／Aiing */
  function energyRemarkSuggestsAiing_(ev) {
    const r = String((ev && ev.remark) || "").toLowerCase();
    if (!r) return false;
    return /aiing|\bai\b|chatgpt|claude|gemini|\bgpt\b|同\s*ai|同ai|問\s*ai|傾.*ai|chat\s*with\s*ai|cursor\b|copilot|llm/.test(
      r,
    );
  }

  /**
   * Fooding 內嵌 Aiing 分鐘：優先 distractionSec；remark 有「AI + N 分」亦可。
   * Recover 只用 duration − aiMin（真·食飯時間）。
   */
  function energyFoodingEmbeddedAiMinutes_(ev, durationMin) {
    const dur = Math.max(0, Number(durationMin) || 0);
    if (dur <= 0 || !energyRemarkSuggestsAiing_(ev)) return 0;
    let aiMin = 0;
    const dSec = Number(ev && ev.distractionSec) || 0;
    if (dSec > 0) aiMin = dSec / 60;
    const remark = String((ev && ev.remark) || "");
    const m1 = remark.match(/(?:aiing|\bai\b)[^\d]{0,16}(\d{1,3})\s*(?:m|min|mins|分鐘|分)\b/i);
    const m2 = remark.match(/(\d{1,3})\s*(?:m|min|mins|分鐘|分)[^\d]{0,16}(?:aiing|\bai\b)/i);
    if (m1) aiMin = Math.max(aiMin, parseInt(m1[1], 10));
    if (m2) aiMin = Math.max(aiMin, parseInt(m2[1], 10));
    if (aiMin <= 0) return 0;
    return Math.max(0, Math.min(dur, aiMin));
  }

  /**
   * @returns {"high"|"medium"|"low"|"recover"|"sleep"|"none"}
   */
  function energyTierOfEvent_(ev) {
    const key = activityKeyOfEv_(ev);
    if (key === "sleeping") return "sleep";
    if (key === "reading") {
      return energyReadingIsLeisure_(ev) ? "recover" : "medium";
    }
    if (ENERGY_RECOVERY[key]) return "recover";
    if (ENERGY_HIGH_KEYS.has(key)) return "high";
    if (ENERGY_MED_KEYS.has(key)) return "medium";
    if (ENERGY_LOW_KEYS.has(key)) return "low";
    return "none";
  }

  function energyIsHighActivityEv_(ev) {
    return energyTierOfEvent_(ev) === "high";
  }

  function energyIsWorkTier_(tier) {
    return tier === "high" || tier === "medium";
  }

  function createEnergyDayState_(maxCap) {
    const cap = Math.max(ENERGY_MIN_DAY_CAP, Math.min(ENERGY_FP_CAP, maxCap != null ? maxCap : ENERGY_FP_CAP));
    return {
      fp: cap,
      startFp: cap,
      maxCap: cap,
      highMin: 0,
      medMin: 0,
      lowMin: 0,
      highTradingMin: 0,
      drainedPts: 0,
      recoveredPts: 0,
      sleepTotalMin: 0,
      sleepRecoveredPts: 0,
      sleepPenaltyPts: 0,
      recoverUsed: Object.create(null),
      foodingAiMin: 0,
      workStreakMin: 0,
      recoverStreakMin: 0,
      lastActivityEndMs: null,
      pendingInertia: false,
      inertiaLeftMin: 0,
      systemTemp: "idle",
      lastHeatMult: 1,
      lastDrainMult: 1,
    };
  }

  function energyUpdateSystemTemp_(st, tier) {
    if (energyIsWorkTier_(tier)) {
      if (st.inertiaLeftMin > 0 || st.workStreakMin < ENERGY_RECOVER_RESET_MIN) {
        st.systemTemp = "cold";
      } else if (st.workStreakMin <= ENERGY_HEAT_START_MIN) {
        st.systemTemp = "optimal";
      } else {
        st.systemTemp = "overheating";
      }
      return;
    }
    if (tier === "recover" || tier === "sleep") {
      st.systemTemp = "cooldown";
      return;
    }
    st.systemTemp = "idle";
  }

  /**
   * 段開始：用 last_activity_end_time 計 break；Recover>15 重置 heat；
   * Break/Recover>60 → 下一節 High 頭 20 分 Inertia 1.3x。
   */
  function energyOnSegmentStart_(st, t0, tier) {
    let breakMin = 0;
    if (st.lastActivityEndMs != null && t0 > st.lastActivityEndMs) {
      breakMin = (t0 - st.lastActivityEndMs) / 60000;
    }
    if (breakMin > 0.01) {
      st.workStreakMin = 0;
      st.recoverStreakMin = 0;
      if (breakMin > ENERGY_INERTIA_BREAK_MIN) st.pendingInertia = true;
    }
    const prevRecover = st.recoverStreakMin;
    if (energyIsWorkTier_(tier)) {
      if (prevRecover > ENERGY_RECOVER_RESET_MIN) {
        st.workStreakMin = 0;
      }
      if (
        tier === "high" &&
        (st.pendingInertia || prevRecover > ENERGY_INERTIA_BREAK_MIN || breakMin > ENERGY_INERTIA_BREAK_MIN)
      ) {
        st.inertiaLeftMin = Math.max(st.inertiaLeftMin, ENERGY_INERTIA_HIGH_MIN);
        st.pendingInertia = false;
      }
      st.recoverStreakMin = 0;
    } else if (tier === "recover") {
      st.workStreakMin = 0;
    } else if (tier === "sleep") {
      st.workStreakMin = 0;
      st.recoverStreakMin = 0;
    } else {
      st.workStreakMin = 0;
    }
  }

  function energyApplySleepMinute_(st, step) {
    const s = Math.max(0, step);
    if (s <= 0) return;
    const prevTotal = energySleepRecoveryTotal_(st.sleepTotalMin);
    st.sleepTotalMin += s;
    const nextTotal = energySleepRecoveryTotal_(st.sleepTotalMin);
    const gain = nextTotal - prevTotal;
    if (gain > 0) {
      st.fp = clampEnergyFp_(st.fp + gain, st.maxCap);
      st.sleepRecoveredPts += gain;
      st.recoveredPts += gain;
    }
    st.workStreakMin = 0;
    st.recoverStreakMin = 0;
    energyUpdateSystemTemp_(st, "sleep");
  }

  function energyApplyRecoverMinute_(st, key, step, segmentMin) {
    const s = Math.max(0, step);
    if (s <= 0) return;
    st.workStreakMin = 0;
    st.recoverStreakMin += s;
    if (key === "reading") {
      let rate = 1.0;
      if (segmentMin < ENERGY_RECOVER_RESET_MIN) rate *= 0.5;
      const used = st.recoverUsed.leisure_reading || 0;
      const room = Math.max(0, 60 - used);
      const gain = Math.min(room, rate * s);
      st.fp = clampEnergyFp_(st.fp + gain, st.maxCap);
      st.recoverUsed.leisure_reading = used + gain;
      st.recoveredPts += gain;
      energyUpdateSystemTemp_(st, "recover");
      return;
    }
    const cfg = ENERGY_RECOVERY[key];
    if (!cfg) {
      energyUpdateSystemTemp_(st, "recover");
      return;
    }
    let rate = cfg.rate;
    if (key !== "fooding" && segmentMin < ENERGY_RECOVER_RESET_MIN) rate *= 0.5;
    const used = st.recoverUsed[key] || 0;
    const room = Math.max(0, cfg.cap - used);
    if (room > 0) {
      const gain = Math.min(room, rate * s);
      st.fp = clampEnergyFp_(st.fp + gain, st.maxCap);
      st.recoverUsed[key] = used + gain;
      st.recoveredPts += gain;
    }
    energyUpdateSystemTemp_(st, "recover");
  }

  function energyApplyDrainMinute_(st, baseRate, tier, step, isTrading) {
    const s = Math.max(0, step);
    if (s <= 0) return;
    let mult = energyDayDrainScale_(st.maxCap);
    if (energyIsWorkTier_(tier)) {
      st.workStreakMin += s;
      const heat = energyHeatMultiplier_(st.workStreakMin);
      mult *= heat;
      st.lastHeatMult = heat;
    } else {
      st.workStreakMin = 0;
      st.lastHeatMult = 1;
    }
    if (tier === "high" && st.inertiaLeftMin > 0) {
      const inertStep = Math.min(s, st.inertiaLeftMin);
      const inertFrac = inertStep / s;
      mult *= 1 + (ENERGY_INERTIA_MULT - 1) * inertFrac;
      st.inertiaLeftMin = Math.max(0, st.inertiaLeftMin - inertStep);
    }
    st.lastDrainMult = mult;
    const loss = baseRate * mult * s;
    st.fp = clampEnergyFp_(st.fp - loss, st.maxCap);
    st.drainedPts += loss;
    if (tier === "high") {
      st.highMin += s;
      if (isTrading) st.highTradingMin += s;
    } else if (tier === "medium") st.medMin += s;
    else if (tier === "low") st.lowMin += s;
    if (energyIsWorkTier_(tier)) st.recoverStreakMin = 0;
    energyUpdateSystemTemp_(st, tier);
  }

  /**
   * 計算某一 wake-day 結束時嘅 FP 快照（Model 3.0）。
   * @param {number} refMs
   * @param {{ endAtMs?: number, nowMs?: number, depth?: number }} [opts]
   */
  function calculateWakeDayEnergy_(refMs, opts) {
    const o = opts || {};
    const bounds = wakeDayBounds(refMs);
    const endAt = o.endAtMs != null ? o.endAtMs : bounds.endMs;
    const nowMs = o.nowMs != null ? o.nowMs : Date.now();
    const depth = o.depth || 0;
    const list = sortedEventsUniqueById();

    let maxCap = ENERGY_FP_CAP;
    if (depth < 12) {
      const prevSnap = calculateWakeDayEnergy_(bounds.startMs - 1, {
        depth: depth + 1,
        endAtMs: bounds.startMs,
        nowMs: bounds.startMs,
      });
      if (prevSnap.fp < -1e-9) {
        maxCap = Math.max(
          ENERGY_MIN_DAY_CAP,
          Math.min(ENERGY_FP_CAP, Math.round(ENERGY_FP_CAP + prevSnap.fp)),
        );
      }
    }

    const st = createEnergyDayState_(maxCap);
    const dayEv = [];
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const t0 = new Date(ev.start).getTime();
      if (Number.isNaN(t0) || t0 < bounds.startMs || t0 >= endAt) continue;
      dayEv.push({ ev: ev, index: i, t0: t0 });
    }
    dayEv.sort((a, b) => a.t0 - b.t0);
    const liveDay = nowMs >= bounds.startMs && nowMs < endAt;

    for (let i = 0; i < dayEv.length; i++) {
      const row = dayEv[i];
      let nextGlobal = null;
      for (let j = row.index + 1; j < list.length; j++) {
        const tj = new Date(list[j].start).getTime();
        if (!Number.isNaN(tj) && tj > row.t0) {
          nextGlobal = tj;
          break;
        }
      }
      const cred = energySegmentCreditMinutes_(
        row.ev,
        row.t0,
        nextGlobal,
        endAt,
        nowMs,
        liveDay,
      );
      if (cred.durationMin <= 0) continue;
      applyEnergySegmentByMinutes_(st, row.ev, cred.durationMin, null, row.t0);
    }

    const highMin = st.highMin;
    const highTradingMin = st.highTradingMin;
    const nonTradeHigh = highMin - highTradingMin;
    const banTrade = highMin >= 240 || (highMin >= 120 && nonTradeHigh > 0.01);
    const blockHigh = highMin >= 240;
    const systemCrash = st.fp <= ENERGY_FP_FLOOR + 1e-9;

    let level = "green";
    let message = "System Stable. High-bandwidth tasks allowed.";
    if (systemCrash) {
      level = "black";
      message = "System Crash (−500 FP). All Work hard-blocked. Stop immediately.";
    } else if (st.fp < 0) {
      level = "red";
      message = "System Overload. All focus tasks forbidden. Please sleep or rest immediately.";
    } else if (st.fp <= 300) {
      level = "orange";
      message =
        "Critical: Cognitive lock risk! Mandatory 15min Resting/Meditating recommended.";
    } else if (st.fp <= 700) {
      level = "yellow";
      message = "";
    }

    let suggest = "";
    if (st.systemTemp === "overheating") {
      suggest =
        "You are in Flow, but overheating. Take a 15-min walk to reset drain rates.";
    } else if (systemCrash || st.fp < 0) {
      suggest = "建議下一個活動：Sleeping／Resting（修復）";
    } else if (blockHigh || banTrade) {
      suggest =
        "建議下一個活動：" +
        (banTrade && !blockHigh ? "Medium／Low（禁 Trading）" : "Recover／Low（禁 High／Trading）");
    } else if (st.fp <= 300) {
      suggest = "建議下一個活動：15min+ Resting／Meditating／Walking";
    } else if (st.systemTemp === "cold" && st.inertiaLeftMin > 0) {
      suggest = "Cold start（Inertia）：頭 20 分鐘 High drain ×1.3 — 細步進入。";
    }

    return {
      fp: Math.round(st.fp * 10) / 10,
      startFp: Math.round(st.startFp * 10) / 10,
      maxCap: st.maxCap,
      level: level,
      message: message,
      suggest: suggest,
      highMin: Math.round(highMin * 10) / 10,
      medMin: Math.round(st.medMin * 10) / 10,
      lowMin: Math.round(st.lowMin * 10) / 10,
      highTradingMin: Math.round(highTradingMin * 10) / 10,
      foodingAiMin: Math.round((st.foodingAiMin || 0) * 10) / 10,
      banTrade: banTrade,
      blockHigh: blockHigh,
      systemCrash: systemCrash,
      blockAllWork: systemCrash,
      drainMult: Math.round(st.lastDrainMult * 100) / 100,
      heatMult: Math.round(st.lastHeatMult * 100) / 100,
      workStreakMin: Math.round(st.workStreakMin * 10) / 10,
      inertiaLeftMin: Math.round(st.inertiaLeftMin * 10) / 10,
      systemTemp: st.systemTemp,
      sleepTotalMin: Math.round(st.sleepTotalMin * 10) / 10,
      sleepRecoveredPts: Math.round(st.sleepRecoveredPts * 10) / 10,
      bounds: bounds,
    };
  }

  /** 公開別名（方便之後擴展／測試） */
  function calculateEnergy(refMs) {
    return calculateWakeDayEnergy_(refMs != null ? refMs : Date.now());
  }

  /** 按 Report 日期跨度揀曲線粒度 */
  function energyChartGranularity_(fromYmd, toYmd) {
    const a = new Date(fromYmd + "T12:00:00").getTime();
    const b = new Date(toYmd + "T12:00:00").getTime();
    const days = Math.max(1, Math.round((b - a) / 86400000) + 1);
    if (days <= 1) {
      return { mode: "minute", bucketMs: 5 * 60000, days: days, label: "每 5 分鐘" };
    }
    if (days <= 14) {
      return { mode: "hour", bucketMs: 3600000, days: days, label: "每小時" };
    }
    return { mode: "day", bucketMs: 86400000, days: days, label: "每日" };
  }

  function energySegmentCreditMinutes_(ev, t0, nextGlobal, endAt, nowMs, liveDay) {
    let t1 = nextGlobal != null ? nextGlobal : liveDay ? Math.max(t0, nowMs) : endAt;
    if (t1 > endAt) t1 = endAt;
    let durationMin = Math.max(0, (t1 - t0) / 60000);
    const tier = energyTierOfEvent_(ev);
    if (tier !== "sleep") {
      const gapMin =
        nextGlobal != null ? (nextGlobal - t0) / 60000 : (endAt - t0) / 60000;
      const endsAtWake = t1 >= endAt - 1;
      const orphan =
        gapMin > ENERGY_ORPHAN_GAP_MIN ||
        (endsAtWake && nextGlobal != null && nextGlobal > endAt) ||
        (nextGlobal == null && !liveDay);
      if (orphan && durationMin > ENERGY_ORPHAN_CREDIT_CAP_MIN) {
        durationMin = ENERGY_ORPHAN_CREDIT_CAP_MIN;
      }
      if (durationMin > ENERGY_MAX_SEGMENT_NON_SLEEP_MIN) {
        durationMin = ENERGY_MAX_SEGMENT_NON_SLEEP_MIN;
      }
    }
    return { durationMin: durationMin, t1: t1, tier: tier };
  }

  /**
   * 將一段活動按「分鐘」推進 st（banner／chart 共用 Model 3.0）。
   * @param {number|null} t0Ms segment start；用於 break／last_activity_end_time
   */
  function applyEnergySegmentByMinutes_(st, ev, durationMin, onMin, t0Ms) {
    const tier = energyTierOfEvent_(ev);
    const key = activityKeyOfEv_(ev);
    const n = Math.max(0, Math.floor(durationMin));
    const frac = durationMin - n;
    const aiMin =
      tier === "recover" && key === "fooding"
        ? energyFoodingEmbeddedAiMinutes_(ev, durationMin)
        : 0;
    const isTrading = isTradingActivityEv_(ev);
    const t0 = t0Ms != null ? t0Ms : 0;

    energyOnSegmentStart_(st, t0, tier === "recover" && key === "fooding" && aiMin > 0 ? "medium" : tier);

    // Fooding 前半 AI：以 medium work session 開始；轉 recover 時再 trigger recover streak
    let foodPhase = aiMin > 0 ? "ai" : "food";

    function stepOne(isFrac) {
      const step = isFrac ? frac : 1;
      if (step <= 0) return;
      if (tier === "sleep") {
        energyApplySleepMinute_(st, step);
      } else if (tier === "recover" && key === "fooding") {
        const idx = st._segMinIndex || 0;
        if (idx < aiMin) {
          if (foodPhase !== "ai") {
            energyOnSegmentStart_(st, t0 + idx * 60000, "medium");
            foodPhase = "ai";
          }
          energyApplyDrainMinute_(st, ENERGY_MED_RATE, "medium", step, false);
          st.foodingAiMin = (st.foodingAiMin || 0) + step;
        } else {
          if (foodPhase !== "food") {
            energyOnSegmentStart_(st, t0 + idx * 60000, "recover");
            foodPhase = "food";
          }
          energyApplyRecoverMinute_(st, "fooding", step, Math.max(0, durationMin - aiMin));
        }
        st._segMinIndex = idx + step;
      } else if (tier === "recover") {
        energyApplyRecoverMinute_(st, key === "reading" ? "reading" : key, step, durationMin);
      } else if (tier === "high") {
        energyApplyDrainMinute_(st, ENERGY_HIGH_RATE, "high", step, isTrading);
      } else if (tier === "medium") {
        energyApplyDrainMinute_(st, ENERGY_MED_RATE, "medium", step, false);
      } else if (tier === "low") {
        energyApplyDrainMinute_(st, ENERGY_LOW_RATE, "low", step, false);
      } else {
        energyUpdateSystemTemp_(st, "none");
      }
      if (typeof onMin === "function") onMin(st.fp, step);
    }

    st._segMinIndex = 0;
    for (let i = 0; i < n; i++) stepOne(false);
    if (frac > 0.01) stepOne(true);
    delete st._segMinIndex;

    const creditedMs = Math.max(0, durationMin) * 60000;
    st.lastActivityEndMs = t0 + creditedMs;
  }

  function pushEnergySample_(points, bucketMs, t, fp, meta) {
    const bt = Math.floor(t / bucketMs) * bucketMs;
    const last = points.length ? points[points.length - 1] : null;
    const act = meta && meta.activity ? String(meta.activity) : "";
    const remark = meta && meta.remark ? String(meta.remark) : "";
    if (last && last.t === bt) {
      last.fp = Math.round(fp * 10) / 10;
      if (act) last.activity = act;
      if (remark) last.remark = remark;
      return;
    }
    points.push({
      t: bt,
      fp: Math.round(fp * 10) / 10,
      activity: act,
      remark: remark,
    });
  }

  /**
   * Report 用：由 from～to 產生 FP 時間序列。
   * @returns {{ mode: string, label: string, points: {t:number,fp:number}[] }}
   */
  function buildFocusEnergySeries_(fromYmd, toYmd) {
    const gran = energyChartGranularity_(fromYmd, toYmd);
    const list = sortedEventsUniqueById();
    const nowMs = Date.now();
    const points = [];

    if (gran.mode === "day") {
      let cursor = new Date(fromYmd + "T12:00:00");
      const end = new Date(toYmd + "T12:00:00");
      let guard = 0;
      while (cursor.getTime() <= end.getTime() && guard < 800) {
        guard++;
        const snap = calculateWakeDayEnergy_(cursor.getTime(), { nowMs: nowMs });
        points.push({
          t: snap.bounds.startMs,
          fp: snap.fp,
        });
        cursor.setDate(cursor.getDate() + 1);
      }
      return { mode: gran.mode, label: gran.label, bucketMs: gran.bucketMs, points: points };
    }

    // minute / hour：逐個 wake-day 模擬（同 Model 3.0）
    let dayCursor = wakeDayBounds(new Date(fromYmd + "T12:00:00")).startMs;
    const rangeEnd = wakeDayBounds(new Date(toYmd + "T12:00:00")).endMs;
    let guard = 0;
    while (dayCursor < rangeEnd && guard < 40) {
      guard++;
      const bounds = wakeDayBounds(dayCursor);
      const endAt = Math.min(bounds.endMs, rangeEnd);
      const liveDay = nowMs >= bounds.startMs && nowMs < bounds.endMs;
      // 同 Model 3.0：前一日負 FP → 今日 maxCap 下調
      const prevSnap = calculateWakeDayEnergy_(bounds.startMs - 1, {
        endAtMs: bounds.startMs,
        nowMs: bounds.startMs,
      });
      let maxCap = ENERGY_FP_CAP;
      if (prevSnap.fp < -1e-9) {
        maxCap = Math.max(
          ENERGY_MIN_DAY_CAP,
          Math.min(ENERGY_FP_CAP, Math.round(ENERGY_FP_CAP + prevSnap.fp)),
        );
      }
      const st = createEnergyDayState_(maxCap);
      pushEnergySample_(points, gran.bucketMs, bounds.startMs, st.fp);

      const dayEv = [];
      for (let i = 0; i < list.length; i++) {
        const t0 = new Date(list[i].start).getTime();
        if (Number.isNaN(t0) || t0 < bounds.startMs || t0 >= endAt) continue;
        dayEv.push({ ev: list[i], index: i, t0: t0 });
      }
      dayEv.sort((a, b) => a.t0 - b.t0);

      for (let i = 0; i < dayEv.length; i++) {
        const row = dayEv[i];
        let nextGlobal = null;
        for (let j = row.index + 1; j < list.length; j++) {
          const tj = new Date(list[j].start).getTime();
          if (!Number.isNaN(tj) && tj > row.t0) {
            nextGlobal = tj;
            break;
          }
        }
        const cred = energySegmentCreditMinutes_(
          row.ev,
          row.t0,
          nextGlobal,
          endAt,
          nowMs,
          liveDay,
        );
        if (cred.durationMin <= 0) continue;
        const tipMeta =
          gran.mode === "minute" || gran.mode === "hour"
            ? {
                activity: activityDisplayName(row.ev.activityId) || "",
                remark: String(row.ev.remark || "").trim(),
              }
            : null;
        pushEnergySample_(points, gran.bucketMs, row.t0, st.fp, tipMeta);
        let elapsed = 0;
        applyEnergySegmentByMinutes_(
          st,
          row.ev,
          cred.durationMin,
          () => {
            elapsed += 1;
            const wall = row.t0 + elapsed * 60000;
            pushEnergySample_(points, gran.bucketMs, wall, st.fp, tipMeta);
          },
          row.t0,
        );
      }
      const endFpT = liveDay ? Math.min(nowMs, endAt) : endAt;
      pushEnergySample_(points, gran.bucketMs, endFpT - 1, st.fp);
      dayCursor = bounds.endMs;
    }

    return { mode: gran.mode, label: gran.label, bucketMs: gran.bucketMs, points: points };
  }

  function formatEnergyChartTick_(t, mode) {
    const d = new Date(t);
    const pad = (n) => String(n).padStart(2, "0");
    if (mode === "day") {
      return pad(d.getMonth() + 1) + "-" + pad(d.getDate());
    }
    if (mode === "hour") {
      return pad(d.getMonth() + 1) + "-" + pad(d.getDate()) + " " + pad(d.getHours()) + ":00";
    }
    return pad(d.getHours()) + ":" + pad(d.getMinutes());
  }

  const ENERGY_CHART_LAYOUT_ = {
    w: 720,
    h: 220,
    padL: 44,
    padR: 12,
    padT: 16,
    padB: 36,
  };

  /** 純 SVG 專注力曲線（唔引入 chart library） */
  function buildFocusEnergyChartHtmlFromPack_(pack, fromYmd, toYmd) {
    const pts = (pack && pack.points) || [];
    if (pts.length < 2) {
      return (
        '<div class="energy-chart-card">' +
        '<h2 class="report-h">Focus Energy</h2>' +
        '<p class="muted">Not enough points for ' +
        escapeHtml((pack && pack.label) || "") +
        " chart.</p></div>"
      );
    }
    const { w, h, padL, padR, padT, padB } = ENERGY_CHART_LAYOUT_;
    const ymin = ENERGY_FP_FLOOR;
    const ymax = ENERGY_FP_CAP;
    const t0 = pts[0].t;
    const t1 = pts[pts.length - 1].t;
    const spanT = Math.max(1, t1 - t0);
    const xOf = (t) => padL + ((t - t0) / spanT) * (w - padL - padR);
    const yOf = (fp) => {
      const c = Math.max(ymin, Math.min(ymax, fp));
      return padT + (1 - (c - ymin) / (ymax - ymin)) * (h - padT - padB);
    };
    let d = "";
    for (let i = 0; i < pts.length; i++) {
      const x = xOf(pts[i].t);
      const y = yOf(pts[i].fp);
      d += (i === 0 ? "M" : "L") + x.toFixed(1) + "," + y.toFixed(1) + " ";
    }
    const y0 = yOf(0);
    const tickIdx = [0];
    if (pts.length > 2) tickIdx.push(Math.floor((pts.length - 1) / 2));
    tickIdx.push(pts.length - 1);
    const uniq = [...new Set(tickIdx)];
    let ticks = "";
    for (let i = 0; i < uniq.length; i++) {
      const p = pts[uniq[i]];
      const x = xOf(p.t);
      ticks +=
        '<text x="' +
        x.toFixed(1) +
        '" y="' +
        (h - 8) +
        '" text-anchor="middle" class="energy-chart-tick">' +
        escapeHtml(formatEnergyChartTick_(p.t, pack.mode)) +
        "</text>";
    }
    return (
      '<div class="energy-chart-card">' +
      '<div class="energy-chart-head">' +
      '<h2 class="report-h" style="margin:0;">Focus Energy</h2>' +
      '<span class="muted energy-chart-gran">' +
      escapeHtml(pack.label) +
      " · " +
      escapeHtml(fromYmd) +
      " → " +
      escapeHtml(toYmd) +
      "</span></div>" +
      '<div class="energy-chart-svg-wrap">' +
      '<svg class="energy-chart-svg" viewBox="0 0 ' +
      w +
      " " +
      h +
      '" role="img" aria-label="Focus energy chart">' +
      '<line x1="' +
      padL +
      '" y1="' +
      y0.toFixed(1) +
      '" x2="' +
      (w - padR) +
      '" y2="' +
      y0.toFixed(1) +
      '" class="energy-chart-zero" stroke="rgba(255,255,255,0.22)" stroke-width="1" stroke-dasharray="4 4"/>' +
      '<text x="8" y="' +
      (padT + 4) +
      '" class="energy-chart-tick">1000</text>' +
      '<text x="8" y="' +
      y0.toFixed(1) +
      '" class="energy-chart-tick">0</text>' +
      '<text x="8" y="' +
      (h - padB) +
      '" class="energy-chart-tick">-500</text>' +
      '<path d="' +
      d.trim() +
      '" fill="none" stroke="#ee8326" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>' +
      '<line class="energy-chart-crosshair hidden" x1="0" y1="' +
      padT +
      '" x2="0" y2="' +
      (h - padB) +
      '" stroke="rgba(255,255,255,0.35)" stroke-width="1"/>' +
      '<circle class="energy-chart-dot hidden" r="4" fill="#ee8326" stroke="#fff" stroke-width="1.5"/>' +
      ticks +
      "</svg>" +
      '<div class="energy-chart-tooltip hidden" role="tooltip"></div>' +
      "</div></div>"
    );
  }

  function bindFocusEnergyChartTooltip_(host, pack) {
    const card = host.querySelector(".energy-chart-card");
    const wrap = host.querySelector(".energy-chart-svg-wrap");
    const svg = host.querySelector(".energy-chart-svg");
    const tip = host.querySelector(".energy-chart-tooltip");
    const cross = host.querySelector(".energy-chart-crosshair");
    const dot = host.querySelector(".energy-chart-dot");
    const pts = (pack && pack.points) || [];
    if (!card || !wrap || !svg || !tip || pts.length < 2) return;

    const { w, h, padL, padR, padT, padB } = ENERGY_CHART_LAYOUT_;
    const ymin = ENERGY_FP_FLOOR;
    const ymax = ENERGY_FP_CAP;
    const t0 = pts[0].t;
    const t1 = pts[pts.length - 1].t;
    const spanT = Math.max(1, t1 - t0);
    const xOf = (t) => padL + ((t - t0) / spanT) * (w - padL - padR);
    const yOf = (fp) => {
      const c = Math.max(ymin, Math.min(ymax, fp));
      return padT + (1 - (c - ymin) / (ymax - ymin)) * (h - padT - padB);
    };
    const showActRemark = pack.mode === "minute" || pack.mode === "hour";

    const hide = () => {
      tip.classList.add("hidden");
      if (cross) cross.classList.add("hidden");
      if (dot) dot.classList.add("hidden");
    };

    const onMove = (ev) => {
      const clientX = ev.clientX != null ? ev.clientX : 0;
      const rect = svg.getBoundingClientRect();
      const cardRect = card.getBoundingClientRect();
      if (!rect.width) return;
      const xSvg = ((clientX - rect.left) / rect.width) * w;
      let best = 0;
      let bestDx = Infinity;
      for (let i = 0; i < pts.length; i++) {
        const dx = Math.abs(xOf(pts[i].t) - xSvg);
        if (dx < bestDx) {
          bestDx = dx;
          best = i;
        }
      }
      const p = pts[best];
      const x = xOf(p.t);
      const y = yOf(p.fp);

      let html =
        '<div class="energy-chart-tip-line">' +
        escapeHtml(formatEnergyChartTick_(p.t, pack.mode)) +
        " · " +
        escapeHtml(String(p.fp)) +
        " FP</div>";
      if (showActRemark) {
        const act = String(p.activity || "").trim() || "—";
        const rm = String(p.remark || "").trim();
        html +=
          '<div class="energy-chart-tip-line energy-chart-tip-act">' +
          escapeHtml(act) +
          "</div>";
        if (rm) {
          const rmShort = rm.length > 120 ? rm.slice(0, 120) + "…" : rm;
          html +=
            '<div class="energy-chart-tip-line energy-chart-tip-remark">' +
            escapeHtml(rmShort) +
            "</div>";
        }
      }
      tip.innerHTML = html;
      tip.classList.remove("hidden");

      // FP≈1000 近頂：tooltip 改顯示喺點下方，避免被裁切
      const yPx = (y / h) * rect.height;
      const showBelow = y < padT + 36 || yPx < 48;
      tip.classList.toggle("energy-chart-tooltip--below", showBelow);
      const tipX = clientX - cardRect.left;
      const tipY = rect.top - cardRect.top + yPx;
      tip.style.left = Math.max(12, Math.min(cardRect.width - 12, tipX)) + "px";
      tip.style.top = tipY + "px";

      if (cross) {
        cross.setAttribute("x1", String(x));
        cross.setAttribute("x2", String(x));
        cross.classList.remove("hidden");
      }
      if (dot) {
        dot.setAttribute("cx", String(x));
        dot.setAttribute("cy", String(y));
        dot.classList.remove("hidden");
      }
    };

    wrap.onmousemove = onMove;
    wrap.onmouseleave = hide;
    wrap.ontouchstart = (ev) => {
      if (ev.touches && ev.touches[0]) {
        onMove(ev.touches[0]);
      }
    };
  }

  function mountFocusEnergyChart_(fromYmd, toYmd) {
    const host = document.getElementById("reportEnergyChart");
    if (!host) return;
    if (!fromYmd || !toYmd) {
      host.innerHTML = "";
      return;
    }
    let pack;
    try {
      pack = buildFocusEnergySeries_(fromYmd, toYmd);
    } catch (e) {
      host.innerHTML =
        '<div class="energy-chart-card"><p class="muted">Focus energy chart unavailable.</p></div>';
      return;
    }
    host.innerHTML = buildFocusEnergyChartHtmlFromPack_(pack, fromYmd, toYmd);
    bindFocusEnergyChartTooltip_(host, pack);
  }

  function energyTempLabel_(temp) {
    if (temp === "cold") return "Cold";
    if (temp === "optimal") return "Optimal";
    if (temp === "overheating") return "Overheating";
    if (temp === "cooldown") return "Cooldown";
    return "Idle";
  }

  function refreshEnergyBanner_() {
    const banner = document.getElementById("energyBanner");
    if (!banner) return;
    // 防止 Generate 中斷後 ai-report-busy 殘留，連 Report／chart 都睇唔到
    if (
      document.documentElement.classList.contains("ai-report-busy") &&
      !document.getElementById("btnAiReportGenerate")?.disabled
    ) {
      document.documentElement.classList.remove("ai-report-busy");
    }
    const snap = calculateEnergy(Date.now());
    _lastEnergySnap_ = snap;
    banner.classList.remove(
      "hidden",
      "energy-banner--green",
      "energy-banner--yellow",
      "energy-banner--orange",
      "energy-banner--red",
      "energy-banner--black",
    );
    banner.classList.add("energy-banner--" + snap.level);

    const fill = document.getElementById("energyBarFill");
    const msg = document.getElementById("energyBannerMsg");
    const sug = document.getElementById("energyBannerSuggest");
    const meta = document.getElementById("energyBannerMeta");
    const tempEl = document.getElementById("energyBannerTemp");
    const maxCap = snap.maxCap != null ? snap.maxCap : ENERGY_FP_CAP;
    const pct = Math.max(
      0,
      Math.min(100, ((snap.fp - ENERGY_FP_FLOOR) / (maxCap - ENERGY_FP_FLOOR)) * 100),
    );
    if (fill) fill.style.width = pct.toFixed(1) + "%";
    if (msg) {
      msg.textContent = snap.message || "";
      msg.classList.toggle("hidden", !snap.message);
    }
    if (sug) {
      sug.textContent = snap.suggest || "";
      sug.classList.toggle("hidden", !snap.suggest);
    }
    if (tempEl) {
      const t = snap.systemTemp || "idle";
      tempEl.textContent = "Temp: " + energyTempLabel_(t);
      tempEl.classList.remove(
        "energy-temp--cold",
        "energy-temp--optimal",
        "energy-temp--overheating",
        "energy-temp--cooldown",
        "energy-temp--idle",
      );
      tempEl.classList.add("energy-temp--" + t);
      tempEl.classList.toggle("hidden", false);
    }
    if (meta) {
      const bits = [];
      if (maxCap < ENERGY_FP_CAP) bits.push("Max Cap " + maxCap);
      if (snap.workStreakMin > 0) bits.push("Work streak " + Math.round(snap.workStreakMin) + "m");
      if (snap.heatMult > 1) bits.push("Heat ×" + snap.heatMult);
      if (snap.inertiaLeftMin > 0) bits.push("Inertia " + Math.round(snap.inertiaLeftMin) + "m");
      meta.textContent = bits.join(" · ");
      meta.classList.toggle("hidden", !bits.length);
    }

    // Work／Sleep 進行中：較密刷新（Heat／Inertia／Sleep 曲線）
    const list = sortedEventsUniqueById();
    const last = list.length ? list[list.length - 1] : null;
    const lastTier = last ? energyTierOfEvent_(last) : "none";
    const needTick =
      lastTier === "sleep" || lastTier === "high" || lastTier === "medium" || lastTier === "recover";
    if (needTick) {
      if (!_energyRefreshTimer_) {
        _energyRefreshTimer_ = setInterval(() => {
          refreshEnergyBanner_();
          refreshSoftCapBanner();
        }, 30000);
      }
    } else if (_energyRefreshTimer_) {
      clearInterval(_energyRefreshTimer_);
      _energyRefreshTimer_ = null;
    }
  }

  /**
   * 超限後仍可入 Work／Trading，但 Remark 必填（當 Reason）。
   * 17:00 起至翌日 wake：硬性唔准入 Work。
   * Energy：−500 硬擋所有 Work；4h High 硬擋再入 High；Trading 禁令跟 2h／4h High 規則。
   * @returns {{ ok: boolean, message?: string, needsReason?: boolean, workOver?: boolean, tradingOver?: boolean, hardBlock?: boolean }}
   */
  function softCapGateForEvent(ev) {
    const refMs = new Date(ev.start).getTime();
    if (Number.isNaN(refMs)) return { ok: true };
    const remark = String(ev.remark || "").trim();
    const isWork = eventWorkRestGroupForCap_(ev) === "Work";
    const isTrading = isTradingActivityEv_(ev);
    const tier = energyTierOfEvent_(ev);
    const energy = calculateWakeDayEnergy_(refMs);

    if (isWork && isInWorkHardBlockWindow_(refMs)) {
      return {
        ok: false,
        hardBlock: true,
        message: MSG_GET_REST_AFTER_17,
        needsReason: false,
      };
    }
    if (isWork && energy.blockAllWork) {
      return {
        ok: false,
        hardBlock: true,
        message: "FP at floor (−500). All Work hard-blocked. Rest / sleep first.",
        needsReason: false,
      };
    }
    if (energy.blockHigh && tier === "high") {
      return {
        ok: false,
        hardBlock: true,
        message:
          "High Drain ≥ 4h today. No more High work (incl. Trading). Switch to Recover / Low.",
        needsReason: false,
      };
    }
    if (energy.banTrade && isTrading) {
      return {
        ok: false,
        hardBlock: true,
        message:
          "Trading forbidden: High Drain ≥ 2h (non-trading High) or ≥ 4h total High. Cognitive overload.",
        needsReason: false,
      };
    }

    const workPack = sumWorkMsInWakeDay(refMs);
    const tradePack = sumTradingMsInWakeDay(refMs);
    const workCap = getWorkCapMs_();
    const tradeCap = getTradingCapMs_();
    const workOver = workPack.ms > workCap;
    const tradingOver = tradePack.ms > tradeCap;
    const needsReason = (isWork && workOver) || (isTrading && tradingOver);
    if (!needsReason) return { ok: true, needsReason: false, workOver, tradingOver };
    if (remark) return { ok: true, needsReason: true, workOver, tradingOver };
    const bits = [];
    if (isWork && workOver) {
      bits.push("Today's Work: " + formatCapClock_(workPack.ms) + " — please fill Reasons in Remark.");
    }
    if (isTrading && tradingOver) {
      bits.push("Today's Trading: " + formatCapClock_(tradePack.ms) + " — please fill Reasons in Remark.");
    }
    return { ok: false, message: bits.join(" "), needsReason: true, workOver, tradingOver };
  }

  function applyReasonPrefixIfNeeded_(ev, needsReason) {
    if (!needsReason || !ev) return;
    const r = String(ev.remark || "").trim();
    if (!r) return;
    if (/^reason\s*:/i.test(r)) ev.remark = r;
    else ev.remark = "Reason: " + r;
  }

  function syncRemarkFieldLabels_(needReasons) {
    const title = needReasons ? "Reasons" : "Remark";
    const qLab = document.querySelector("label[for=quickRemark]");
    const mLab = document.querySelector("label[for=manualRemark]");
    if (qLab) qLab.textContent = title;
    if (mLab) mLab.textContent = title;
    const q = document.getElementById("quickRemark");
    const m = document.getElementById("manualRemark");
    if (q) {
      q.placeholder = "";
      q.removeAttribute("placeholder");
    }
    if (m) {
      m.placeholder = "";
      m.removeAttribute("placeholder");
    }
  }

  function refreshSoftCapBanner() {
    refreshEnergyBanner_();
    const el = document.getElementById("softCapBanner");
    if (!el) return;
    const now = Date.now();
    const workPack = sumWorkMsInWakeDay(now);
    const tradePack = sumTradingMsInWakeDay(now);
    const transportPack = sumActivityKeysMsInWakeDay(now, TRANSPORTING_ACTIVITY_KEYS);
    const socialPack = sumActivityKeysMsInWakeDay(now, SOCIAL_NO_TRADE_KEYS);
    const workCap = getWorkCapMs_();
    const tradeCap = getTradingCapMs_();
    const workOver = workPack.ms > workCap;
    const tradingOver = tradePack.ms > tradeCap;
    const noTradesToday =
      transportPack.ms > NO_TRADES_BANNER_MS || socialPack.ms > NO_TRADES_BANNER_MS;
    const energy = _lastEnergySnap_ || calculateEnergy(now);
    // 17:00 唔做常駐 banner；只喺嘗試入 Work 時 modal + 硬擋
    // 舊 soft-cap 字條：Work／Trading 鐘數超標、Transport／Social → no Trades；Energy 禁令另見主 Banner
    syncRemarkFieldLabels_(workOver || tradingOver);
    const showSoft =
      workOver || tradingOver || noTradesToday || energy.banTrade || energy.blockHigh || energy.systemCrash;
    if (!showSoft) {
      el.classList.add("hidden");
      el.classList.remove("soft-cap-banner--warn", "soft-cap-banner--danger");
      el.textContent = "";
    } else {
      const lines = [];
      if (energy.systemCrash) lines.push("Energy crash: Work hard-blocked");
      else if (energy.blockHigh) lines.push("High Drain ≥4h: High/Trading blocked");
      else if (energy.banTrade) lines.push("Trading forbidden (High Drain rule)");
      if (workOver) lines.push("Today's Work: " + formatCapClock_(workPack.ms));
      if (tradingOver) lines.push("Today's Trading: " + formatCapClock_(tradePack.ms));
      if (noTradesToday) lines.push(MSG_NO_TRADES_TODAY);
      el.textContent = lines.join(" · ");
      el.classList.remove("hidden", "soft-cap-banner--warn");
      el.classList.add("soft-cap-banner--danger");
    }
    // Email 唔依賴 banner 顯示（例如净 Reviewing 超 30m）
    void maybeNotifyCapEmails_();
  }

  let _capEmailInflight = false;
  async function maybeNotifyCapEmails_() {
    if (!canRemoteSync() || _capEmailInflight) return;
    const url = getRemotePostUrl();
    if (!url) return;
    const now = Date.now();
    const wakeDayKey = wakeDayKey_(now);
    const reviewing = sumActivityKeysMsInWakeDay(now, REVIEWING_ACTIVITY_KEYS);
    const trading = sumTradingMsInWakeDay(now);
    const jobs = [];
    if (reviewing.ms > REVIEWING_EMAIL_MS) {
      jobs.push({
        rule: "reviewing",
        durationMs: reviewing.ms,
        thresholdLabel: "30 minutes",
      });
    }
    if (trading.ms > getTradingCapMs_()) {
      jobs.push({
        rule: "trading",
        durationMs: trading.ms,
        thresholdLabel: "2 hours",
      });
    }
    if (!jobs.length) return;
    _capEmailInflight = true;
    try {
      for (let i = 0; i < jobs.length; i++) {
        const job = jobs[i];
        try {
          const r = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(
              remoteAuthBody({
                action: "notifyCap",
                rule: job.rule,
                wakeDayKey,
                durationMs: job.durationMs,
                thresholdLabel: job.thresholdLabel,
              })
            ),
            mode: "cors",
            cache: "no-store",
          });
          const j = await r.json().catch(() => ({}));
          // Cap email 係 best-effort：token 過期／失敗唔好鎖成個 app
          if (j && j.ok === false) {
            const err = String(j.error || "");
            if (
              err === "id_token_expired" ||
              err === "invalid_id_token" ||
              err === "missing_id_token" ||
              err === "unauthorized"
            ) {
              break;
            }
          }
        } catch (eJob) {}
      }
    } finally {
      _capEmailInflight = false;
    }
  }

  /**
   * Time Stat mapping rules（對應 vault：`Time Stat mapping rules.md`）——僅用於 Report「Raw records」顯示，唔改寫入庫 event。
   */
  function inferTimeStatMappingForRaw(ev, nextEv) {
    const act = activityDisplayName(ev.activityId);
    const key = normalizeActivityKey(act);
    const group = inferRulesWorkRestGroup(ev);
    let layer;
    let cat;
    let subCat;

    if (key === "transporting") {
      const asc = sortedEventsUniqueById();
      const nextMap = chronologicalNextById(asc);
      let follow = nextEv;
      let hops = 0;
      while (follow && hops < 8) {
        const fk = normalizeActivityKey(activityDisplayName(follow.activityId));
        if (fk !== "transporting") break;
        follow = nextMap.get(follow.id) || null;
        hops++;
      }
      if (!follow) {
        layer = "Freedom";
        cat = "Time";
      } else {
        const nb = inferRulesLayerCatExcludeTransporting(follow);
        const nl = (nb.layer || "").trim().toLowerCase();
        if (nl === "health") {
          layer = "Health";
          cat = nb.cat || "Mental Health";
        } else {
          layer = "Freedom";
          const nbLayerLow = (nb.layer || "").toLowerCase();
          cat =
            nb.cat && nb.cat !== "Needs-Review" && nbLayerLow !== "needs-review" ? nb.cat : "Time";
        }
      }
      subCat =
        layer === "Freedom"
          ? inferFreedomSubCatFromRules(ev, act, ev.remark || "")
          : inferHealthSubCatFromRemark(ev.remark || "");
      return { group, layer, cat, subCat };
    }

    const base = inferRulesLayerCatExcludeTransporting(ev);
    layer = base.layer;
    cat = base.cat;
    if (layer === "Freedom") subCat = inferFreedomSubCatFromRules(ev, act, ev.remark || "");
    else if (layer === "Health") subCat = inferHealthSubCatFromRemark(ev.remark || "");
    else subCat = "Needs-Review";

    return { group, layer, cat, subCat };
  }

  /** 報表篩選／匯總／Keyword 與 Raw 表共用：同一套 <code>inferTimeStatMappingForRaw</code>；Cat 經 <code>normalizeCatDisplayForRaw</code>（例如 <code>Time</code> → Time Management）。 */
  function reportInferredMapping(ev, list) {
    const nextMap = chronologicalNextById(list);
    const nextEv = nextMap.get(ev.id) || null;
    const m = inferTimeStatMappingForRaw(ev, nextEv);
    const catRaw = String(m.cat || "").trim();
    const catDisp = normalizeCatDisplayForRaw(catRaw) || catRaw;
    return {
      group: reportNormLabel(m.group),
      layer: reportNormLabel(m.layer),
      cat: reportNormLabel(catDisp),
      subCat: reportNormLabel(m.subCat),
    };
  }

  let pendingApproval = null;

  function clearApprovalPanel() {
    pendingApproval = null;
    const card = document.getElementById("mappingApprovalCard");
    const list = document.getElementById("mappingApprovalList");
    const meta = document.getElementById("mappingApprovalMeta");
    if (list) list.innerHTML = "";
    if (meta) meta.textContent = "";
    if (card) card.classList.add("hidden");
  }

  function clearQuickLogForm() {
    ["quickPlace", "quickActivity", "quickPeople", "quickRemark"].forEach((id) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.value = "";
      el.dataset.userEdited = "0";
      delete el.dataset.autoSuggestedValue;
    });
    refreshQuickAutoSuggestions();
  }

  function clearManualLogForm() {
    ["manualPlace", "manualActivity", "manualPeople", "manualRemark"].forEach((id) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.value = "";
      el.dataset.userEdited = "0";
      delete el.dataset.autoSuggestedValue;
    });
    initManualDateTime();
    refreshManualAutoSuggestions();
  }

  function scheduleRenderReport() {
    if (_aiReportBusy_) {
      _reportRenderPendingWhileAi_ = true;
      return;
    }
    if (_reportRenderScheduled) return;
    _reportRenderScheduled = true;
    requestAnimationFrame(() => {
      _reportRenderScheduled = false;
      if (_aiReportBusy_) {
        _reportRenderPendingWhileAi_ = true;
        return;
      }
      renderReport();
    });
  }

  function yieldToMain_() {
    return new Promise((resolve) => {
      if (typeof requestAnimationFrame === "function") {
        requestAnimationFrame(() => setTimeout(resolve, 0));
      } else {
        setTimeout(resolve, 0);
      }
    });
  }

  function setAiReportBusy_(on) {
    _aiReportBusy_ = !!on;
    document.documentElement.classList.toggle("ai-report-busy", _aiReportBusy_);
    if (!_aiReportBusy_ && _reportRenderPendingWhileAi_) {
      _reportRenderPendingWhileAi_ = false;
      scheduleRenderReport();
    }
  }

  function stopAiReportGenTimer_() {
    if (_aiReportGenTimer_) {
      clearInterval(_aiReportGenTimer_);
      _aiReportGenTimer_ = null;
    }
    _aiReportGenStartedAt_ = 0;
  }

  function startAiReportGenTimer_(el) {
    stopAiReportGenTimer_();
    _aiReportGenStartedAt_ = Date.now();
    const tick = () => {
      if (!el || !el.classList.contains("loading")) {
        stopAiReportGenTimer_();
        return;
      }
      const sec = Math.max(0, Math.floor((Date.now() - _aiReportGenStartedAt_) / 1000));
      const p = el.querySelector(".ai-report-loading");
      if (p) {
        p.innerHTML =
          '<span class="ai-report-loading-spin" aria-hidden="true"></span>' +
          "Generating… " +
          sec +
          "s";
      }
    };
    tick();
    _aiReportGenTimer_ = setInterval(tick, 1000);
  }

  function scrollPageToTopInstant() {
    try {
      const html = document.documentElement;
      const prev = html.style.scrollBehavior;
      html.style.scrollBehavior = "auto";
      window.scrollTo(0, 0);
      html.scrollTop = 0;
      if (document.body) document.body.scrollTop = 0;
      html.style.scrollBehavior = prev || "";
    } catch (e) {
      try {
        window.scrollTo(0, 0);
      } catch (e2) {}
    }
  }

  /**
   * Distraction 掛喺「而家進行中」= 時間上上一筆 activity（唔跟新入庫嗰筆）。
   * @returns {{ applied: boolean, target?: object, sec: number }}
   */
  function applyDistractionToOngoingPrevious_(newEv) {
    const sec = currentDistractionSec_();
    if (sec <= 0) return { applied: false, sec: 0 };
    const list = sortedEventsUniqueById();
    const tNew = new Date(newEv && newEv.start).getTime();
    if (Number.isNaN(tNew)) return { applied: false, sec };
    let prev = null;
    for (let i = 0; i < list.length; i++) {
      const t = new Date(list[i].start).getTime();
      if (Number.isNaN(t)) continue;
      if (t <= tNew) prev = list[i];
      else break;
    }
    if (!prev) return { applied: false, sec };
    const target = state.events.find((e) => e && e.id === prev.id) || prev;
    target.distractionSec = (Number(target.distractionSec) || 0) + sec;
    return { applied: true, target, sec };
  }

  /**
   * 連續同一個 activity：合入時間上前一筆（保留較早 start；remark 用 ", " 串）。
   * @returns {{ merged: boolean, ev: object }}
   */
  function mergeIntoPreviousIfSameActivity_(ev) {
    const list = sortedEventsUniqueById();
    const tNew = new Date(ev.start).getTime();
    if (Number.isNaN(tNew)) return { merged: false, ev };
    let prev = null;
    for (let i = 0; i < list.length; i++) {
      const t = new Date(list[i].start).getTime();
      if (Number.isNaN(t)) continue;
      if (t <= tNew) prev = list[i];
      else break;
    }
    if (!prev) return { merged: false, ev };
    const kPrev = activityKeyOfEv_(prev);
    const kNew = activityKeyOfEv_(ev);
    if (!kPrev || !kNew || kPrev !== kNew) return { merged: false, ev };

    const target = state.events.find((e) => e && e.id === prev.id) || prev;
    const parts = [];
    const r1 = String(target.remark || "").trim();
    const r2 = String(ev.remark || "").trim();
    if (r1) parts.push(r1);
    if (r2) parts.push(r2);
    if (parts.length) target.remark = parts.join(", ");
    else delete target.remark;

    if (!String(target.place || "").trim() && ev.place) target.place = ev.place;
    if ((!target.people || !target.people.length) && Array.isArray(ev.people) && ev.people.length) {
      target.people = ev.people.slice();
    }
    if (!target.group && ev.group) target.group = ev.group;
    if (!target.layer && ev.layer) target.layer = ev.layer;
    if (!target.cat && ev.cat) target.cat = ev.cat;
    if (!target.subCat && ev.subCat) target.subCat = ev.subCat;
    if (!target.projectId && ev.projectId) target.projectId = ev.projectId;
    if (!String(target.projectsFromForm || "").trim() && ev.projectsFromForm) {
      target.projectsFromForm = ev.projectsFromForm;
    }
    // distraction 已喺 push 前掛去上一筆；唔再跟新筆合併

    return { merged: true, ev: target };
  }

  function pushEventAndRefresh(ev, msg, opts) {
    const silent = opts && opts.silent;
    const formSource = opts && opts.formSource;
    // 新筆唔帶 distraction；累計掛去上一筆（進行中）activity
    if (ev) delete ev.distractionSec;
    const dist = applyDistractionToOngoingPrevious_(ev);
    if (dist.sec > 0 && dist.applied) resetDistractionWatch_();
    else if (dist.sec > 0 && !dist.applied) {
      toast("Distraction kept — no previous activity to attach.");
    } else {
      resetDistractionWatch_();
    }
    const merged = mergeIntoPreviousIfSameActivity_(ev);
    if (!merged.merged) state.events.push(ev);
    const savedEv = merged.ev;
    bumpEventsMutationGen();
    save();
    refreshActivityDatalist();
    renderTimeline();
    refreshSoftCapBanner();
    if (silent) {
      if (formSource === "manual") clearManualLogForm();
      else if (formSource === "quick") clearQuickLogForm();
      // 等高度收合完先即時回頂；避免 smooth scroll + 內容縮短喺 iOS 造成半邊黑屏
      requestAnimationFrame(() => {
        scrollPageToTopInstant();
        requestAnimationFrame(scrollPageToTopInstant);
      });
    } else {
      refreshQuickAutoSuggestions();
      refreshManualAutoSuggestions();
      updateLastSavedHint(savedEv);
      toast(`${msg}${merged.merged ? " · merged" : ""} · 總筆數 ${state.events.length}`);
    }
    requestAnimationFrame(() => {
      fillMergeSelects();
      scheduleRenderReport();
    });
  }

  function updateLastSavedHint(ev) {
    const el = document.getElementById("lastSavedHint");
    if (!el) return;
    if (!ev) {
      if (!state.events.length) {
        el.textContent = "尚未入庫新紀錄。";
        return;
      }
      const u = sortedEventsUniqueById();
      const latest = u[u.length - 1];
      if (!latest) {
        el.textContent = "尚未入庫新紀錄。";
        return;
      }
      ev = latest;
    }
    const when = new Date(ev.start).toLocaleString("zh-Hant", { hour12: false });
    const act = activityDisplayName(ev.activityId);
    el.textContent = `最近入庫：${when} · ${act}（累計 ${state.events.length} 筆）`;
  }

  function showApprovalPanel(payload) {
    pendingApproval = payload;
    const card = document.getElementById("mappingApprovalCard");
    const list = document.getElementById("mappingApprovalList");
    const meta = document.getElementById("mappingApprovalMeta");
    if (!card || !list || !meta) return;
    const place = String(payload.ev.place || "").trim() || "—";
    const withStr =
      payload.ev.people && payload.ev.people.length ? payload.ev.people.join(", ") : "—";
    const remarkDisp = String(payload.remark || payload.ev.remark || "").trim() || "—";
    meta.innerHTML =
      `<div class="mapping-context">` +
      `<div>Place</div><div>${escapeHtml(place)}</div>` +
      `<div>With</div><div>${escapeHtml(withStr)}</div>` +
      `<div>Remark</div><div>${escapeHtml(remarkDisp)}</div>` +
      `</div>`;
    list.innerHTML = "";
    const projectOpts = uniqueProjectsSorted();
    const mkSel = (id, options, selected, allowBlank) => {
      let h = `<select id="${id}" class="mapping-edit-select">`;
      if (allowBlank) h += `<option value="">blank</option>`;
      for (let i = 0; i < options.length; i++) {
        const v = options[i];
        const sel = v === selected ? " selected" : "";
        h += `<option value="${escapeHtml(v)}"${sel}>${escapeHtml(v)}</option>`;
      }
      h += `</select>`;
      return h;
    };
    payload.candidates.forEach((c) => {
      const row = document.createElement("div");
      row.className = "mapping-suggestion";
      const editableId = uid();
      const catSelected = normalizeCatDisplayForRaw(c.cat) || c.cat || CLASSIFICATION_CAT_OPTIONS[0];
      const projOpts = reportUniqueSorted([...projectOpts, c.project].filter(Boolean));
      row.innerHTML =
        `<h4 class="mapping-suggestion-title">${escapeHtml(c.label || "Suggestion")}</h4>` +
        `<div class="mapping-grid">` +
        `<div>Group</div><div>${mkSel(`map-g-${editableId}`, CLASSIFICATION_GROUP_OPTIONS, c.group || CLASSIFICATION_GROUP_OPTIONS[0], false)}</div>` +
        `<div>Layers</div><div>${mkSel(`map-l-${editableId}`, CLASSIFICATION_LAYER_OPTIONS, c.layer || CLASSIFICATION_LAYER_OPTIONS[0], false)}</div>` +
        `<div>Cat</div><div>${mkSel(`map-c-${editableId}`, CLASSIFICATION_CAT_OPTIONS, catSelected, false)}</div>` +
        `<div>Sub Cat</div><div>${mkSel(`map-s-${editableId}`, CLASSIFICATION_SUB_OPTIONS, c.subCat || CLASSIFICATION_SUB_OPTIONS[0], false)}</div>` +
        `<div>Activity</div><div>${escapeHtml(c.activity || "—")}</div>` +
        `<div>Project</div><div>${mkSel(`map-p-${editableId}`, projOpts, c.project || "", true)}</div>` +
        `</div>`;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "primary";
      btn.style.marginTop = "8px";
      btn.textContent = "Confirm";
      btn.addEventListener("click", () => {
        try {
          const ev = { ...pendingApproval.ev };
          const gEl = document.getElementById(`map-g-${editableId}`);
          const lEl = document.getElementById(`map-l-${editableId}`);
          const cEl = document.getElementById(`map-c-${editableId}`);
          const sEl = document.getElementById(`map-s-${editableId}`);
          const pEl = document.getElementById(`map-p-${editableId}`);
          const gVal = gEl ? reportNormLabel(gEl.value) : c.group || "";
          const lVal = lEl ? reportNormLabel(lEl.value) : c.layer || "";
          const catVal = cEl ? reportNormLabel(cEl.value) : c.cat || "";
          const sVal = sEl ? reportNormLabel(sEl.value) : c.subCat || "";
          const pVal = pEl ? reportNormLabel(pEl.value) : c.project || "";
          ev.group = gVal || undefined;
          ev.layer = lVal || undefined;
          ev.cat = catVal || undefined;
          ev.subCat = sVal || undefined;
          if (pVal) {
            ev.projectsFromForm = pVal;
            ev.projectId = projectIdByName(pVal) || c.projectId || undefined;
          } else {
            delete ev.projectsFromForm;
            delete ev.projectId;
          }
          delete ev.project;
          if (ev.group === "Work" || ev.group === "Rest") ev.category = ev.group;
          const formSrc = pendingApproval.formSource || "quick";
          const remarkEl =
            formSrc === "manual"
              ? document.getElementById("manualRemark")
              : document.getElementById("quickRemark");
          if (remarkEl && String(remarkEl.value || "").trim()) {
            ev.remark = String(remarkEl.value).trim();
          }
          const gate = softCapGateForEvent(ev);
          if (!gate.ok) {
            notifyGateFailure_(gate);
            return;
          }
          applyReasonPrefixIfNeeded_(ev, gate.needsReason);
          const msg = pendingApproval.doneMsg;
          const formSource = pendingApproval.formSource;
          clearApprovalPanel();
          pushEventAndRefresh(ev, msg, { silent: true, formSource: formSource });
        } catch (err) {
          toast("入庫失敗：" + (err && err.message ? err.message : "未知錯誤"));
        }
      });
      row.appendChild(btn);
      list.appendChild(row);
    });
    card.classList.remove("hidden");
    card.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  // —— Distraction stopwatch（入庫成功先清零；失敗／取消保留）——
  let _distractAccumMs = 0;
  let _distractRunning = false;
  let _distractStartedAt = 0;
  let _distractTickTimer = null;

  function formatDistractClock_(ms) {
    const totalSec = Math.max(0, Math.floor(ms / 1000));
    const m = Math.floor(totalSec / 60);
    const s = totalSec % 60;
    return String(m).padStart(2, "0") + ":" + String(s).padStart(2, "0");
  }

  function currentDistractionMs_() {
    let ms = _distractAccumMs;
    if (_distractRunning && _distractStartedAt) ms += Date.now() - _distractStartedAt;
    return ms;
  }

  function currentDistractionSec_() {
    return Math.floor(currentDistractionMs_() / 1000);
  }

  function refreshDistractionDisplay_() {
    const el = document.getElementById("distractionDisplay");
    if (el) el.textContent = formatDistractClock_(currentDistractionMs_());
    const startBtn = document.getElementById("btnDistractStart");
    const stopBtn = document.getElementById("btnDistractStop");
    if (startBtn) startBtn.disabled = _distractRunning;
    if (stopBtn) stopBtn.disabled = !_distractRunning;
  }

  function stopDistractTick_() {
    if (_distractTickTimer) {
      clearInterval(_distractTickTimer);
      _distractTickTimer = null;
    }
  }

  function startDistractTick_() {
    stopDistractTick_();
    _distractTickTimer = setInterval(refreshDistractionDisplay_, 250);
  }

  function resetDistractionWatch_() {
    stopDistractTick_();
    _distractAccumMs = 0;
    _distractRunning = false;
    _distractStartedAt = 0;
    refreshDistractionDisplay_();
  }

  function bindDistractionWatch_() {
    const startBtn = document.getElementById("btnDistractStart");
    const stopBtn = document.getElementById("btnDistractStop");
    const resetBtn = document.getElementById("btnDistractReset");
    if (!startBtn || startBtn.dataset.bound) return;
    startBtn.dataset.bound = "1";
    startBtn.addEventListener("click", () => {
      if (_distractRunning) return;
      _distractRunning = true;
      _distractStartedAt = Date.now();
      startDistractTick_();
      refreshDistractionDisplay_();
    });
    stopBtn.addEventListener("click", () => {
      if (!_distractRunning) return;
      _distractAccumMs += Date.now() - _distractStartedAt;
      _distractRunning = false;
      _distractStartedAt = 0;
      stopDistractTick_();
      refreshDistractionDisplay_();
    });
    resetBtn.addEventListener("click", () => {
      resetDistractionWatch_();
    });
    refreshDistractionDisplay_();
  }

  document.getElementById("btnLogNow").addEventListener("click", () => {
    const nowD = new Date();
    const quickActivityEl = document.getElementById("quickActivity");
    const quickPlaceEl = document.getElementById("quickPlace");
    const quickPeopleEl = document.getElementById("quickPeople");
    const quickRemarkEl = document.getElementById("quickRemark");
    const rawLabel = quickActivityEl.value.trim();
    const rawPlace = quickPlaceEl.value.trim();
    const rawPeople = quickPeopleEl.value.trim();
    const rawRemark = quickRemarkEl.value.trim();
    if (allBlank([rawLabel, rawPlace, rawPeople, rawRemark])) {
      toast(MSG_PLEASE_INPUT_DATA);
      return;
    }

    applyAutoSuggestion("quickPlace", mostLikelyPlaceByTime(nowD.getHours(), nowD.getMinutes()));
    const label = quickActivityEl.value.trim();
    const place = quickPlaceEl.value.trim();
    const act = getOrCreateActivity(label);
    if (!act) {
      toast(MSG_ACTIVITY_REQUIRED);
      return;
    }
    if (!place) {
      toast(MSG_PLACE_REQUIRED);
      return;
    }
    const ev = {
      id: uid(),
      start: new Date().toISOString(),
      activityId: act.id,
      people: splitPeople(quickPeopleEl.value),
    };
    const remark = quickRemarkEl.value.trim();
    if (remark) ev.remark = remark;
    if (place) ev.place = place;
    handleEventClassificationFlow({
      ev: ev,
      activityLabel: label,
      remark: remark,
      doneMsg: MSG_LOG_NOW_DONE,
      formSource: "quick",
    });
  });

  document.getElementById("btnManual").addEventListener("click", () => {
    const manualActivityEl = document.getElementById("manualActivity");
    const manualPlaceEl = document.getElementById("manualPlace");
    const manualPeopleEl = document.getElementById("manualPeople");
    const manualRemarkEl = document.getElementById("manualRemark");
    const label = manualActivityEl.value.trim();
    const place = manualPlaceEl.value.trim();
    const rawPeople = manualPeopleEl.value.trim();
    const rawRemark = manualRemarkEl.value.trim();
    if (allBlank([label, place, rawPeople, rawRemark])) {
      toast(MSG_PLEASE_INPUT_DATA);
      return;
    }
    const act = getOrCreateActivity(label);
    const dateNorm = document.getElementById("manualDateSelected").value.trim();
    const hourStr = document.getElementById("manualHourSel").value;
    const minuteStr = document.getElementById("manualMinuteSel").value;
    if (!act || !dateNorm || hourStr === "" || minuteStr === "") {
      toast(MSG_MANUAL_NEED_FIELDS);
      return;
    }
    if (!place) {
      toast(MSG_PLACE_REQUIRED);
      return;
    }
    const d = new Date(`${dateNorm}T${hourStr}:${minuteStr}:00`);
    if (Number.isNaN(d.getTime())) {
      toast(MSG_MANUAL_INVALID_TIME);
      return;
    }
    const ev = {
      id: uid(),
      start: d.toISOString(),
      activityId: act.id,
      people: splitPeople(manualPeopleEl.value),
    };
    const remark = manualRemarkEl.value.trim();
    if (remark) ev.remark = remark;
    if (place) ev.place = place;
    handleEventClassificationFlow({
      ev: ev,
      activityLabel: label,
      remark: remark,
      doneMsg: MSG_MANUAL_DONE,
      formSource: "manual",
    });
  });

  function splitPeople(s) {
    return String(s || "")
      .split(/[,，、]/)
      .map((x) => x.trim())
      .filter(Boolean);
  }

  function allBlank(values) {
    return values.every((v) => !String(v || "").trim());
  }

  function handleEventClassificationFlow(params) {
    const ev = params.ev;
    const activityLabel = params.activityLabel;
    const remark = params.remark;
    const doneMsg = params.doneMsg;
    const formSource = params.formSource || "quick";
    if (remark) ev.remark = String(remark).trim();
    // 入庫前先用推斷／已有 group 檢查；Confirm 改 Group 後會再檢一次
    const probe = { ...ev };
    if (!probe.group) {
      const cands = buildMappingCandidates(activityLabel, remark, ev.start);
      if (cands[0] && cands[0].group) probe.group = cands[0].group;
    }
    const early = softCapGateForEvent(probe);
    if (!early.ok) {
      notifyGateFailure_(early);
      const ta =
        formSource === "manual"
          ? document.getElementById("manualRemark")
          : document.getElementById("quickRemark");
      if (!early.hardBlock && ta) {
        try {
          ta.focus();
        } catch (e) {}
      }
      return;
    }
    const candidates = buildMappingCandidates(activityLabel, remark, ev.start);
    if (!candidates.length) {
      applyReasonPrefixIfNeeded_(ev, early.needsReason);
      pushEventAndRefresh(ev, doneMsg, { silent: true, formSource: formSource });
      return;
    }
    showApprovalPanel({
      ev: ev,
      activityLabel: activityLabel,
      remark: remark,
      candidates: candidates,
      doneMsg: doneMsg,
      formSource: formSource,
    });
  }

  /**
   * 過濾極少出現／疑似錯字：次數 < minCount 唔入榜；
   * 若短字串係另一個更長、更常用字串嘅 prefix（Ubu ⊂ Ubud），剔走短嗰個。
   * @param {{ sortByScore?: boolean, minCount?: number }} [opts]
   */
  function finalizeSuggestRanking_(scoredEntries, countMap, limit, opts) {
    const lim = Math.max(1, limit || 3);
    const minCount = opts && opts.minCount != null ? opts.minCount : 2;
    const sortByScore = !!(opts && opts.sortByScore);
    let entries = scoredEntries.slice().sort((a, b) => {
      if (sortByScore) {
        if (b[1] !== a[1]) return b[1] - a[1];
        const ca = countMap.get(a[0]) || 0;
        const cb = countMap.get(b[0]) || 0;
        if (cb !== ca) return cb - ca;
      } else {
        const ca = countMap.get(a[0]) || 0;
        const cb = countMap.get(b[0]) || 0;
        if (cb !== ca) return cb - ca;
        if (b[1] !== a[1]) return b[1] - a[1];
      }
      return a[0].localeCompare(b[0], "zh-Hant");
    });
    const solid = entries.filter((e) => (countMap.get(e[0]) || 0) >= minCount);
    if (solid.length) entries = solid;
    const labels = entries.map((e) => e[0]);
    const filtered = labels.filter((p) => {
      const pl = p.toLowerCase();
      const longer = labels.find(
        (q) =>
          q !== p &&
          q.toLowerCase().startsWith(pl) &&
          q.length > p.length &&
          (countMap.get(q) || 0) >= (countMap.get(p) || 0)
      );
      return !longer;
    });
    return (filtered.length ? filtered : labels).slice(0, lim);
  }

  /**
   * Place 排名：頻率優先，再計近七日／時段。
   * 舊版「上次 +12」會把單次錯字（如 Ubu）推上第一。
   */
  function rankPlacesByLikelihood(hour, minute, limit) {
    const lim = Math.max(1, limit || 3);
    const now = Date.now();
    const monthAgo = now - 30 * DAY_MS;
    const weekAgo = now - 7 * DAY_MS;
    const score = new Map();
    const count = new Map();
    let latestPlace = "";
    let latestT = -1;
    const targetMins = hour * 60 + minute;
    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      const p = String(ev.place || "").trim();
      if (!p) continue;
      const t = new Date(ev.start).getTime();
      if (Number.isNaN(t) || t < monthAgo) continue;
      count.set(p, (count.get(p) || 0) + 1);
      if (t > latestT) {
        latestT = t;
        latestPlace = p;
      }
      const d = new Date(t);
      const diff = Math.abs(d.getHours() * 60 + d.getMinutes() - targetMins);
      const daysAgo = (now - t) / DAY_MS;
      let w = 1;
      if (t >= weekAgo) w += 5;
      if (daysAgo <= 1) w += 5;
      else if (daysAgo <= 3) w += 3;
      else if (daysAgo <= 7) w += 2;
      if (diff === 0) w += 2;
      else if (diff <= 5) w += 1;
      else if (diff <= 30) w += 0.5;
      score.set(p, (score.get(p) || 0) + w);
    }
    // 只有出現 ≥2 次嘅「上次地點」先加分，避免單次錯字霸榜
    if (latestPlace && (count.get(latestPlace) || 0) >= 2) {
      score.set(latestPlace, (score.get(latestPlace) || 0) + 6);
    }
    return finalizeSuggestRanking_([...score.entries()], count, lim);
  }

  function mostLikelyPlaceByTime(hour, minute) {
    return rankPlacesByLikelihood(hour, minute, 1)[0] || "";
  }

  function clockDiffMins_(aMins, bMins) {
    const diff = Math.abs(aMins - bMins);
    return Math.min(diff, 1440 - diff);
  }

  function timeProximityWeight_(clockDiff) {
    if (clockDiff === 0) return 24;
    if (clockDiff <= 5) return 18;
    if (clockDiff <= 15) return 12;
    if (clockDiff <= 30) return 7;
    if (clockDiff <= 60) return 3;
    if (clockDiff <= 90) return 1;
    return 0;
  }

  /**
   * Activity 排名：以「目標時：分」為主；
   * 加分來源 = 近幾日同時段 + 過去數星期同一個 weekday 同時段。
   * 分數優先（唔再用總頻率壓過時段命中）。
   */
  function rankActivitiesByLikelihood(hour, minute, limit, refDate) {
    const lim = Math.max(1, limit || 3);
    const ref =
      refDate instanceof Date && !Number.isNaN(refDate.getTime()) ? refDate : new Date();
    const refMs = ref.getTime();
    const targetDow = ref.getDay();
    const targetMins = hour * 60 + minute;
    const recentCut = refMs - 7 * DAY_MS;
    const dowLookback = refMs - 56 * DAY_MS; // ~8 個同一個 weekday
    const score = new Map();
    const hitCount = new Map();
    const rawCount = new Map();

    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      const a = String(activityDisplayName(ev.activityId) || "").trim();
      if (!a || a === "（已刪 Activity）") continue;
      const t = new Date(ev.start).getTime();
      if (Number.isNaN(t) || t > refMs || t < dowLookback) continue;
      rawCount.set(a, (rawCount.get(a) || 0) + 1);

      const d = new Date(t);
      const clockDiff = clockDiffMins_(d.getHours() * 60 + d.getMinutes(), targetMins);
      const prox = timeProximityWeight_(clockDiff);
      if (prox <= 0) continue;

      const daysAgo = (refMs - t) / DAY_MS;
      const sameDow = d.getDay() === targetDow;
      let w = 0;

      // 近 7 日（任何 weekday）：同時段
      if (t >= recentCut) {
        const dayDecay = daysAgo <= 1 ? 1.25 : daysAgo <= 3 ? 1 : 0.75;
        w += prox * dayDecay;
      }

      // 過去數星期同一個 weekday：同時段（含今日／近七日，可疊加）
      if (sameDow) {
        const weeksAgo = daysAgo / 7;
        const weekDecay =
          weeksAgo <= 1 ? 1.15 : weeksAgo <= 2 ? 1 : weeksAgo <= 4 ? 0.8 : 0.55;
        w += prox * 1.35 * weekDecay;
      }

      if (w <= 0) continue;
      hitCount.set(a, (hitCount.get(a) || 0) + 1);
      score.set(a, (score.get(a) || 0) + w);
    }

    // 分數優先；rawCount 只用作剔錯字（Ubu 類）
    return finalizeSuggestRanking_([...score.entries()], rawCount, lim, {
      sortByScore: true,
      minCount: 1,
    });
  }

  function mostLikelyActivityByTime(hour, minute, refDate) {
    return rankActivitiesByLikelihood(hour, minute, 1, refDate)[0] || "";
  }

  function applyAutoSuggestion(inputId, suggestedValue) {
    const inp = document.getElementById(inputId);
    if (!inp) return;
    const s = String(suggestedValue || "").trim();
    if (!s) return;
    const edited = inp.dataset.userEdited === "1";
    if (edited) return;
    inp.value = s;
    inp.dataset.autoSuggestedValue = s;
    inp.dataset.userEdited = "0";
  }

  function closeAllSuggestMenus_() {
    document.querySelectorAll(".suggest-menu").forEach((m) => m.classList.add("hidden"));
  }

  function openSuggestMenu_(inputId, options) {
    const menu = document.getElementById(inputId + "Menu");
    const inp = document.getElementById(inputId);
    if (!menu || !inp) return;
    closeAllSuggestMenus_();
    const opts = (options || []).map((x) => String(x || "").trim()).filter(Boolean);
    menu.innerHTML = "";
    if (!opts.length) {
      const empty = document.createElement("div");
      empty.className = "suggest-menu-empty muted";
      empty.textContent = "No suggestions yet";
      menu.appendChild(empty);
    } else {
      opts.forEach((label, idx) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "suggest-menu-item";
        btn.textContent = idx + 1 + ". " + label;
        btn.addEventListener("mousedown", (e) => e.preventDefault());
        btn.addEventListener("click", () => {
          inp.value = label;
          inp.dataset.userEdited = "1";
          inp.dataset.autoSuggestedValue = label;
          menu.classList.add("hidden");
        });
        menu.appendChild(btn);
      });
    }
    menu.classList.remove("hidden");
  }

  function bindSuggestChevron_(inputId, getTopOptions) {
    const btn = document.querySelector('[data-suggest-for="' + inputId + '"]');
    const menu = document.getElementById(inputId + "Menu");
    if (!btn || !menu) return;
    btn.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (!menu.classList.contains("hidden")) {
        menu.classList.add("hidden");
        return;
      }
      openSuggestMenu_(inputId, getTopOptions());
    });
  }

  function bindSmartInput(inputId, getSuggestedValue) {
    const inp = document.getElementById(inputId);
    if (!inp) return;
    inp.addEventListener("focus", () => {
      closeAllSuggestMenus_();
      const auto = String(inp.dataset.autoSuggestedValue || "").trim();
      const cur = String(inp.value || "").trim();
      const edited = inp.dataset.userEdited === "1";
      if (!edited && auto && cur === auto) {
        inp.value = "";
      }
    });
    inp.addEventListener("input", () => {
      inp.dataset.userEdited = inp.value.trim() ? "1" : "0";
    });
    inp.addEventListener("blur", () => {
      const cur = String(inp.value || "").trim();
      if (cur) {
        inp.dataset.userEdited = "1";
        return;
      }
      inp.dataset.userEdited = "0";
      const s = String(getSuggestedValue() || "").trim();
      if (s) {
        inp.value = s;
        inp.dataset.autoSuggestedValue = s;
      }
    });
  }

  function manualSelectedHM() {
    const h = parseInt(document.getElementById("manualHourSel")?.value || "", 10);
    const m = parseInt(document.getElementById("manualMinuteSel")?.value || "", 10);
    if (Number.isFinite(h) && Number.isFinite(m)) return { h, m };
    const d = new Date();
    return { h: d.getHours(), m: d.getMinutes() };
  }

  /** Manual 表單目標時間（含日期 → weekday）；缺欄就用而家 */
  function manualSelectedRefDate() {
    const dateNorm = String(document.getElementById("manualDateSelected")?.value || "").trim();
    const hm = manualSelectedHM();
    if (dateNorm) {
      const d = new Date(
        `${dateNorm}T${String(hm.h).padStart(2, "0")}:${String(hm.m).padStart(2, "0")}:00`
      );
      if (!Number.isNaN(d.getTime())) return d;
    }
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate(), hm.h, hm.m, 0, 0);
  }

  function refreshQuickAutoSuggestions() {
    const d = new Date();
    applyAutoSuggestion("quickPlace", mostLikelyPlaceByTime(d.getHours(), d.getMinutes()));
    applyAutoSuggestion("quickActivity", mostLikelyActivityByTime(d.getHours(), d.getMinutes(), d));
  }

  function refreshManualAutoSuggestions() {
    const hm = manualSelectedHM();
    const ref = manualSelectedRefDate();
    applyAutoSuggestion("manualPlace", mostLikelyPlaceByTime(hm.h, hm.m));
    applyAutoSuggestion("manualActivity", mostLikelyActivityByTime(hm.h, hm.m, ref));
  }

  function bindPlaceAutoSuggest() {
    bindSmartInput("quickPlace", () => {
      const d = new Date();
      return mostLikelyPlaceByTime(d.getHours(), d.getMinutes());
    });
    bindSmartInput("manualPlace", () => {
      const hm = manualSelectedHM();
      return mostLikelyPlaceByTime(hm.h, hm.m);
    });
    bindSmartInput("quickActivity", () => {
      const d = new Date();
      return mostLikelyActivityByTime(d.getHours(), d.getMinutes(), d);
    });
    bindSmartInput("manualActivity", () => {
      const hm = manualSelectedHM();
      return mostLikelyActivityByTime(hm.h, hm.m, manualSelectedRefDate());
    });
    bindSuggestChevron_("quickPlace", () => {
      const d = new Date();
      return rankPlacesByLikelihood(d.getHours(), d.getMinutes(), 3);
    });
    bindSuggestChevron_("manualPlace", () => {
      const hm = manualSelectedHM();
      return rankPlacesByLikelihood(hm.h, hm.m, 3);
    });
    bindSuggestChevron_("quickActivity", () => {
      const d = new Date();
      return rankActivitiesByLikelihood(d.getHours(), d.getMinutes(), 3, d);
    });
    bindSuggestChevron_("manualActivity", () => {
      const hm = manualSelectedHM();
      return rankActivitiesByLikelihood(hm.h, hm.m, 3, manualSelectedRefDate());
    });
    document.addEventListener("click", (e) => {
      const t = e.target;
      if (t && t.closest && t.closest(".suggest-field")) return;
      closeAllSuggestMenus_();
    });
  }

  function normalizeActivityKey(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }

  function assignOptionalFormFields(ev, row) {
    const t = (k) => String(row[k] || "").trim();
    const u = (k, field) => {
      const v = t(k);
      if (v) ev[field] = v;
    };
    u("What is your Ultimate Objective?", "objective");
    u("What is your Activity?", "activityQuestion");
    u("What is the Group?", "groupFromForm");
    u("What is the Layers?", "layersFromForm");
    const projBits = [];
    const projSeen = new Set();
    const addProj = (v) => {
      const s = String(v || "").trim();
      if (!s) return;
      const low = s.toLowerCase();
      if (projSeen.has(low)) return;
      projSeen.add(low);
      projBits.push(s);
    };
    for (const key of Object.keys(row)) {
      const ks = String(key).replace(/\u3000/g, " ").replace(/？/g, "?").trim();
      const low = ks.toLowerCase();
      if (/what\s+is\s+the\s+project/i.test(low)) {
        addProj(row[key]);
      }
    }
    if (projBits.length) ev.projectsFromForm = projBits.join(" · ");
    const sheetProj = t("Projects") || t("Project");
    if (sheetProj) {
      if (!ev.projectsFromForm) ev.projectsFromForm = sheetProj;
      else if (!ev.projectsFromForm.toLowerCase().includes(sheetProj.toLowerCase())) {
        ev.projectsFromForm = ev.projectsFromForm + " · " + sheetProj;
      }
    }
    u("What is the Categories", "categoriesFromForm");
    u("What did you achieve in the last time block?", "achievement");
    u("What can you Improve in the last time block?", "improveLast");
    u("The most important element to achieve your objective", "importantElement");
    u("How can you do better in the details?", "detailsBetter");
    u("Action", "action");
    u("Long Term Goals (Year)", "longTermGoals");
    u("Short Term Goals (Months)", "shortTermGoals");
    u("Mini Goals (Weeks)", "miniGoals");
    const r0 = t("Remark");
    const r1 = t("Remark__2");
    const remarkPieces = [r0, r1].filter(Boolean);
    for (const key of Object.keys(row)) {
      const ks = String(key).trim();
      const low = ks.toLowerCase();
      if (low === "remark" || /^remark__/.test(low)) continue;
      if (
        /^notes?$/i.test(low) ||
        /^comments?$/i.test(low) ||
        /^description$/i.test(low) ||
        /^memo$/i.test(low) ||
        /^journal$/i.test(low) ||
        /^feedback$/i.test(low) ||
        /^備註$/.test(ks) ||
        /^備注$/.test(ks) ||
        /^說明$/.test(ks) ||
        /^説明$/.test(ks) ||
        /^日記$/.test(ks) ||
        /^心得$/.test(ks)
      ) {
        const v = String(row[key] || "").trim();
        if (v) remarkPieces.push(v);
      }
    }
    const rm = remarkPieces.join(" · ");
    if (rm) ev.remark = ev.remark ? ev.remark + " · " + rm : rm;
  }

  /** 匯入：避開 Timestamp／Activity 等主欄後，再掃「What is the project」變體、Group、備註類欄。 */
  function mergeImportCsvLooseFields(ev, row, metaCols) {
    const skip = new Set((metaCols || []).filter(Boolean).map((c) => String(c).trim().toLowerCase()));
    for (const key of Object.keys(row)) {
      const kl = String(key).trim().toLowerCase();
      if (skip.has(kl)) continue;
      const nk = String(key).replace(/\u3000/g, " ").replace(/？/g, "?").trim();
      const nkl = nk.toLowerCase();
      if (/what\s+is\s+the\s+project/i.test(nkl)) {
        const v = String(row[key] || "").trim();
        if (!v) continue;
        if (!ev.projectsFromForm) ev.projectsFromForm = v;
        else if (!ev.projectsFromForm.toLowerCase().includes(v.toLowerCase())) {
          ev.projectsFromForm = ev.projectsFromForm + " · " + v;
        }
        continue;
      }
      if (/^group$/i.test(nk)) {
        const v = String(row[key] || "").trim().toLowerCase();
        if (v === "work" || v === "rest") {
          ev.group = v === "work" ? "Work" : "Rest";
          ev.category = ev.group;
        }
        continue;
      }
      if (/remark|notes|comment|description|journal|memo|feedback|備註|備注|日記|心得|說明|説明|反思/i.test(nkl)) {
        const v = String(row[key] || "").trim();
        if (!v) continue;
        const lowV = v.toLowerCase();
        if (String(ev.remark || "").toLowerCase().includes(lowV)) continue;
        ev.remark = ev.remark ? ev.remark + " · " + v : v;
      }
    }
  }

  function renderActivityList() {
    const root = document.getElementById("activityCards");
    if (!root) return;
    root.innerHTML = "";
    state.activities.forEach((e) => {
      const card = document.createElement("div");
      card.className = "card activity-list-row";
      const aliasesStr = (e.aliases || []).join(", ");
      card.innerHTML =
        `<label>名稱（改名會把舊名加入 alias）</label>` +
        `<div class="row"><input type="text" data-id="${e.id}" class="activity-name" value="${escapeHtml(e.name)}" />` +
        `<button type="button" class="danger fixed" data-del="${e.id}">刪</button></div>` +
        `<p class="muted" style="margin:8px 0 0;">Aliases：${escapeHtml(aliasesStr || "—")}</p>`;
      root.appendChild(card);
    });
    root.querySelectorAll(".activity-name").forEach((inp) => {
      inp.addEventListener("change", () => {
        const id = inp.getAttribute("data-id");
        const ent = activityById(state.activities, id);
        if (!ent) return;
        const nv = inp.value.trim();
        if (!nv) {
          inp.value = ent.name;
          return;
        }
        if (nv !== ent.name) {
          if (!ent.aliases) ent.aliases = [];
          if (!ent.aliases.includes(ent.name)) ent.aliases.push(ent.name);
          ent.name = nv;
          save();
          refreshActivityDatalist();
          fillMergeSelects();
          renderTimeline();
          renderReport();
          renderActivityList();
          toast("已改名並保留 alias");
        }
      });
    });
    root.querySelectorAll("button[data-del]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const id = btn.getAttribute("data-del");
        const has = state.events.some((ev) => ev.activityId === id);
        if (has && !confirm("仍有紀錄用緊此 Activity；刪除後會顯示「已刪 Activity」。繼續？")) return;
        if (!has && !confirm("確定刪除此 Activity？")) return;
        state.activities = state.activities.filter((e) => e.id !== id);
        save();
        refreshActivityDatalist();
        fillMergeSelects();
        renderActivityList();
        renderTimeline();
        renderReport();
        toast("已刪除 Activity");
      });
    });
  }

  (function bindActivityAdd() {
    const btn = document.getElementById("btnAddActivity");
    const inp = document.getElementById("newActivityName");
    if (!btn || !inp) return;
    btn.addEventListener("click", () => {
      const name = inp.value.trim();
      if (!name) {
        toast("輸入名稱");
        return;
      }
      if (resolveActivityByLabel(name)) {
        toast("已有同名／alias");
        return;
      }
      state.activities.push({ id: uid(), name, aliases: [] });
      inp.value = "";
      save();
      refreshActivityDatalist();
      fillMergeSelects();
      renderActivityList();
      toast("已新增");
    });
  })();

  (function bindActivityMerge() {
    const btn = document.getElementById("btnMerge");
    const fromSel = document.getElementById("mergeFrom");
    const toSel = document.getElementById("mergeTo");
    if (!btn || !fromSel || !toSel) return;
    btn.addEventListener("click", () => {
      const fromId = fromSel.value;
      const toId = toSel.value;
      if (!fromId || !toId || fromId === toId) {
        toast("揀兩個唔同 Activity");
        return;
      }
      const fromE = activityById(state.activities, fromId);
      const toE = activityById(state.activities, toId);
      if (!fromE || !toE) return;
      if (!confirm(`將「${fromE.name}」合併入「${toE.name}」？所有紀錄會指去後者，前者會刪除。`)) return;
      if (!toE.aliases) toE.aliases = [];
      if (!toE.aliases.includes(fromE.name)) toE.aliases.push(fromE.name);
      (fromE.aliases || []).forEach((a) => {
        if (!toE.aliases.includes(a)) toE.aliases.push(a);
      });
      state.events.forEach((ev) => {
        if (ev.activityId === fromId) ev.activityId = toId;
      });
      state.activities = state.activities.filter((e) => e.id !== fromId);
      save();
      refreshActivityDatalist();
      fillMergeSelects();
      renderActivityList();
      renderTimeline();
      renderReport();
      toast("已合併");
    });
  })();

  let reportPeopleSearchTimer = null;
  let reportKeywordSearchTimer = null;

  function buildReportMappingContext(list) {
    const nextMap = chronologicalNextById(list);
    const inferred = new Map();
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const nextEv = nextMap.get(ev.id) || null;
      const m = inferTimeStatMappingForRaw(ev, nextEv);
      const catRaw = String(m.cat || "").trim();
      const catDisp = normalizeCatDisplayForRaw(catRaw) || catRaw;
      inferred.set(ev.id, {
        group: reportNormLabel(m.group),
        layer: reportNormLabel(m.layer),
        cat: reportNormLabel(catDisp),
        subCat: reportNormLabel(m.subCat),
      });
    }
    return { nextMap, inferred };
  }

  function refreshReportFilterSelectsIfNeeded() {
    if (_reportFiltersCachedGen === _eventsMutationGen) return;
    _reportFiltersCachedGen = _eventsMutationGen;
    refreshReportFilterSelects();
  }

  function reportNormLabel(s) {
    return String(s || "").trim();
  }

  function reportUniqueSorted(arr) {
    return [...new Set(arr.map((x) => reportNormLabel(x)).filter(Boolean))].sort((a, b) => a.localeCompare(b, "zh-Hant"));
  }

  function populateReportSelect(sel, list, preferred) {
    if (!sel) return;
    const pref = reportNormLabel(preferred);
    const cur = pref && list.includes(pref) ? pref : "";
    sel.innerHTML = "";
    const o0 = document.createElement("option");
    o0.value = "";
    o0.textContent = "（全部）";
    sel.appendChild(o0);
    for (let i = 0; i < list.length; i++) {
      const v = list[i];
      const o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      sel.appendChild(o);
    }
    if (cur) sel.value = cur;
  }

  function normalizeSubCatChoice(s) {
    const t = reportNormLabel(s);
    const low = t.toLowerCase();
    if (low === "group" || low === "grouped") return "Grouped";
    if (low === "individual") return "Individual";
    if (low === "projects") return "Project";
    return t;
  }

  function labelsFromStructureItem(item) {
    const raw = reportNormLabel(item);
    const low = raw.toLowerCase();
    if (low === "individual") return ["Individual"];
    if (low === "group" || low === "grouped") return ["Grouped"];
    return [];
  }

  function structureRowMatchesSubFilter(r, subWant) {
    const w = normalizeSubCatChoice(subWant);
    if (!w) return true;
    if (normalizeSubCatChoice(r.subCat) === w) return true;
    const xs = labelsFromStructureItem(r.item);
    for (let i = 0; i < xs.length; i++) {
      if (xs[i] === w) return true;
    }
    return false;
  }

  function eventMatchesSubFilter(ev, subWant) {
    const w = normalizeSubCatChoice(subWant);
    if (!w) return true;
    if (normalizeSubCatChoice(ev.subCat) === w) return true;
    const xs = labelsFromStructureItem(ev.structureItem);
    for (let i = 0; i < xs.length; i++) {
      if (xs[i] === w) return true;
    }
    return false;
  }

  /** 篩選 Sub：與 Raw 一致，用 mapping 推斷嘅 Sub（唔再用 Structure <code>ev.subCat</code>）。 */
  function eventMatchesSubFilterInferred(ev, subWant, list) {
    const w = normalizeSubCatChoice(subWant);
    if (!w) return true;
    return normalizeSubCatChoice(reportInferredMapping(ev, list).subCat) === w;
  }

  /** Project 篩選：只認 <code>projectsFromForm</code>（拆段 + 全串）。 */
  function eventProjectTokensForFilter(ev) {
    const tokens = new Set();
    const add = (t) => {
      const x = reportNormLabel(t);
      if (!x) return;
      tokens.add(x);
      x.split(/\s*[·,，、]\s*/).forEach((p) => {
        const q = reportNormLabel(p);
        if (q) tokens.add(q);
      });
    };
    add(ev.projectsFromForm || "");
    return tokens;
  }

  function eventMatchesProjectReportFilter(ev, wantList) {
    if (!wantList || !wantList.length) return true;
    const tokens = eventProjectTokensForFilter(ev);
    if (!tokens.size) return false;
    return wantList.some((w) => {
      const ww = reportNormLabel(w);
      return ww && tokens.has(ww);
    });
  }

  function subsForReport() {
    const list = sortedEventsUniqueById();
    const a = ["Long Term", "Short Term", "Project", "non-project", "Needs-Review", "Grouped", "Individual"];
    const struct = Array.isArray(state.structure) ? state.structure : [];
    for (let i = 0; i < struct.length; i++) {
      const r = struct[i];
      if (r.subCat) a.push(r.subCat);
      const xs = labelsFromStructureItem(r.item);
      for (let j = 0; j < xs.length; j++) a.push(xs[j]);
    }
    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      if (ev.subCat) a.push(ev.subCat);
      const xs = labelsFromStructureItem(ev.structureItem);
      for (let j = 0; j < xs.length; j++) a.push(xs[j]);
    }
    for (let i = 0; i < list.length; i++) {
      const s = reportInferredMapping(list[i], list).subCat;
      if (s) a.push(s);
    }
    return reportUniqueSorted(a);
  }

  function subsForReportCats(catArr) {
    const list = sortedEventsUniqueById();
    const base = ["Long Term", "Short Term", "Project", "non-project", "Needs-Review", "Grouped", "Individual"];
    if (!catArr || !catArr.length) {
      const a = [...base];
      for (let i = 0; i < list.length; i++) {
        const s = reportInferredMapping(list[i], list).subCat;
        if (s) a.push(s);
      }
      const struct = Array.isArray(state.structure) ? state.structure : [];
      for (let i = 0; i < struct.length; i++) {
        const r = struct[i];
        if (r.subCat) a.push(r.subCat);
        const xs = labelsFromStructureItem(r.item);
        for (let j = 0; j < xs.length; j++) a.push(xs[j]);
      }
      for (let i = 0; i < state.events.length; i++) {
        const ev = state.events[i];
        if (ev.subCat) a.push(ev.subCat);
        const xs = labelsFromStructureItem(ev.structureItem);
        for (let j = 0; j < xs.length; j++) a.push(xs[j]);
      }
      return reportUniqueSorted(a);
    }
    const catKeys = catArr.map((c) => effectiveReportCatKey(c));
    const a = [...base];
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const inf = reportInferredMapping(ev, list);
      const ck = effectiveReportCatKey(inf.cat);
      if (!catKeys.some((c) => c === ck)) continue;
      if (inf.subCat) a.push(inf.subCat);
    }
    const struct = Array.isArray(state.structure) ? state.structure : [];
    for (let i = 0; i < struct.length; i++) {
      const r = struct[i];
      if (!catKeys.some((c) => c === effectiveReportCatKey(r.cat))) continue;
      if (r.subCat) a.push(r.subCat);
      const xs = labelsFromStructureItem(r.item);
      for (let j = 0; j < xs.length; j++) a.push(xs[j]);
    }
    return reportUniqueSorted(a);
  }

  function projectsForReportCatsSubs(catArr, subArr) {
    const list = sortedEventsUniqueById();
    const cats = catArr && catArr.length ? catArr.map(reportNormLabel) : null;
    const subs = subArr && subArr.length ? subArr : null;
    const struct = Array.isArray(state.structure) ? state.structure : [];
    const a = [];
    for (let i = 0; i < struct.length; i++) {
      const r = struct[i];
      if (cats && !cats.some((c) => effectiveReportCatKey(c) === effectiveReportCatKey(r.cat))) continue;
      if (subs && !subs.some((s) => structureRowMatchesSubFilter(r, s))) continue;
      if (r.project) a.push(r.project);
    }
    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      if (cats && !cats.some((c) => effectiveReportCatKey(c) === effectiveReportCatKey(reportInferredMapping(ev, list).cat)))
        continue;
      if (subs && !subs.some((s) => eventMatchesSubFilterInferred(ev, s, list))) continue;
      const pf = String(ev.projectsFromForm || "").trim();
      if (pf) {
        pf.split(/\s*·\s*/)
          .map((x) => x.trim())
          .filter(Boolean)
          .forEach((x) => a.push(x));
      }
    }
    return reportUniqueSorted(a);
  }

  function layersForReport() {
    const list = sortedEventsUniqueById();
    const a = ["Health", "Freedom", "Achievement", "Needs-Review"];
    const struct = Array.isArray(state.structure) ? state.structure : [];
    for (let i = 0; i < struct.length; i++) {
      if (struct[i].layers) a.push(struct[i].layers);
    }
    for (let i = 0; i < state.events.length; i++) {
      if (state.events[i].layer) a.push(state.events[i].layer);
    }
    for (let i = 0; i < list.length; i++) {
      const ly = reportInferredMapping(list[i], list).layer;
      if (ly) a.push(ly);
    }
    return reportUniqueSorted(a);
  }

  function groupsForReport() {
    const list = sortedEventsUniqueById();
    const a = ["Work", "Rest"];
    const struct = Array.isArray(state.structure) ? state.structure : [];
    for (let i = 0; i < struct.length; i++) {
      if (struct[i].group) a.push(struct[i].group);
    }
    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      const g = ev.group || ev.category;
      if (g) a.push(g);
    }
    for (let i = 0; i < list.length; i++) {
      const g = reportInferredMapping(list[i], list).group;
      if (g) a.push(g);
    }
    return reportUniqueSorted(a);
  }

  function readCheckedValuesFromMsBox(boxId) {
    const box = document.getElementById(boxId);
    const s = new Set();
    if (!box) return s;
    box.querySelectorAll('input[type="checkbox"][data-report-ms]').forEach((cb) => {
      if (cb.checked) s.add(cb.value);
    });
    return s;
  }

  function renderReportMultiPick(boxId, options) {
    const box = document.getElementById(boxId);
    if (!box) return;
    const prev = readCheckedValuesFromMsBox(boxId);
    box.innerHTML = "";
    for (let i = 0; i < options.length; i++) {
      const v = options[i];
      const lab = document.createElement("label");
      lab.className = "report-ms-row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.dataset.reportMs = "1";
      cb.value = v;
      cb.checked = prev.has(v);
      const span = document.createElement("span");
      span.textContent = v;
      lab.appendChild(cb);
      lab.appendChild(span);
      box.appendChild(lab);
    }
  }

  function readMultiSet(boxId, extraId) {
    const set = new Set();
    const box = document.getElementById(boxId);
    if (box) {
      box.querySelectorAll('input[type="checkbox"][data-report-ms]:checked').forEach((cb) => {
        const v = reportNormLabel(cb.value);
        if (v) set.add(v);
      });
    }
    const ex = document.getElementById(extraId);
    if (ex && ex.value) {
      String(ex.value)
        .split(/[,，、]/)
        .map((x) => reportNormLabel(x))
        .filter(Boolean)
        .forEach((v) => set.add(v));
    }
    return [...set];
  }

  function catsForReport() {
    const list = sortedEventsUniqueById();
    const a = ["Mental Health", "Physical Health", "Time Management", "Finance", "Needs-Review"];
    const struct = Array.isArray(state.structure) ? state.structure : [];
    for (let i = 0; i < struct.length; i++) {
      if (struct[i].cat) a.push(struct[i].cat);
    }
    for (let i = 0; i < state.events.length; i++) {
      if (state.events[i].cat) a.push(state.events[i].cat);
    }
    for (let i = 0; i < list.length; i++) {
      const c = reportInferredMapping(list[i], list).cat;
      if (c) a.push(c);
    }
    return reportUniqueSorted(a);
  }

  /** Report 側欄 Activity 多選：最近 <code>REPORT_ACTIVITY_FILTER_ROLLING_DAYS</code> 日內有 <code>start</code> 嘅事件所涉及嘅名稱。 */
  function activitiesForReportFilterRollingDays() {
    const to = new Date();
    to.setHours(0, 0, 0, 0);
    const from = new Date(to);
    from.setDate(from.getDate() - (REPORT_ACTIVITY_FILTER_ROLLING_DAYS - 1));
    const fromYmd = ymdFromLocalDate(from);
    const toYmd = ymdFromLocalDate(to);
    const t0 = new Date(fromYmd + "T00:00:00").getTime();
    const t1 = new Date(toYmd + "T23:59:59.999").getTime();
    if (Number.isNaN(t0) || Number.isNaN(t1)) return [];
    const list = sortedEventsUniqueById();
    const acc = [];
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const st = new Date(ev.start).getTime();
      if (Number.isNaN(st)) continue;
      if (st < t0 || st > t1) continue;
      const nm = reportNormLabel(activityDisplayName(ev.activityId));
      if (nm && nm !== "（已刪 Activity）") acc.push(nm);
    }
    return reportUniqueSorted(acc);
  }

  function refreshReportFilterSelects() {
    renderReportMultiPick("reportFilterGroupBox", groupsForReport());
    renderReportMultiPick("reportFilterLayerBox", layersForReport());

    renderReportMultiPick("reportFilterCatBox", catsForReport());
    const catNow = readMultiSet("reportFilterCatBox", "reportFilterCatExtra");

    renderReportMultiPick("reportFilterSubBox", subsForReportCats(catNow));
    const subNow = readMultiSet("reportFilterSubBox", "reportFilterSubExtra");

    renderReportMultiPick("reportFilterProjectBox", projectsForReportCatsSubs(catNow, subNow));
    renderReportMultiPick("reportFilterActivityBox", activitiesForReportFilterRollingDays());
  }

  function reportPeopleTokens(query) {
    return String(query || "")
      .split(/[,，、]/)
      .map((x) => x.trim().toLowerCase())
      .filter(Boolean);
  }

  /** Activity / Remark keyword tokens: split on spaces or commas; every token must match (AND). */
  function reportKeywordTokens(query) {
    return String(query || "")
      .split(/[,，、\s]+/)
      .map((x) => x.trim().toLowerCase())
      .filter(Boolean);
  }

  function reportKeywordActive(f) {
    return String((f && f.keywordQuery) || "").trim().length > 0;
  }

  function normalizeForKeywordLoose(s) {
    try {
      return String(s || "")
        .normalize("NFKC")
        .replace(/[\u200B-\u200D\uFEFF]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toLowerCase();
    } catch {
      return String(s || "")
        .replace(/[\u200B-\u200D\uFEFF]/g, "")
        .trim()
        .toLowerCase();
    }
  }

  /** All searchable text from one event (for report keyword). Optional <code>list</code>：加入推斷維度（同篩選／Raw）。 */
  function eventKeywordSearchHaystack(ev, listForInfer) {
    const parts = [];
    const push = (v) => {
      if (v == null) return;
      if (Array.isArray(v)) {
        for (let i = 0; i < v.length; i++) push(v[i]);
        return;
      }
      const t = String(v).trim();
      if (t) parts.push(t);
    };
    push(ev.group);
    push(ev.category);
    push(ev.layer);
    push(ev.cat);
    push(ev.subCat);
    push(ev.structureItem);
    push(ev.projectId);
    push(ev.projectsFromForm);
    const tryRegistryIds = [];
    if (String(ev.projectId || "").trim()) tryRegistryIds.push(String(ev.projectId).trim());
    const pfMain = String(ev.projectsFromForm || "").trim();
    pfMain.split(/\s*[·,，、]\s*/).forEach((seg) => {
      const s = String(seg || "").trim();
      if (s && /^[0-9a-f-]{36}$/i.test(s)) tryRegistryIds.push(s);
    });
    const seenPid = new Set();
    for (let t = 0; t < tryRegistryIds.length; t++) {
      const pidRaw = tryRegistryIds[t];
      const low = pidRaw.toLowerCase();
      if (seenPid.has(low)) continue;
      seenPid.add(low);
      if (!Array.isArray(state.projectsRegistry)) continue;
      for (let i = 0; i < state.projectsRegistry.length; i++) {
        const r = state.projectsRegistry[i];
        if (String(r.projectId || "").trim().toLowerCase() === low) {
          push(r.project);
          break;
        }
      }
    }
    push(ev.place);
    push(ev.remark);
    push(ev.start);
    const ent = activityById(state.activities, ev.activityId);
    if (ent) {
      push(ent.name);
      push(ent.aliases);
    } else {
      push(activityDisplayName(ev.activityId));
    }
    push(ev.people);
    push(ev.groupFromForm);
    push(ev.layersFromForm);
    push(ev.categoriesFromForm);
    push(ev.objective);
    push(ev.activityQuestion);
    push(ev.achievement);
    push(ev.improveLast);
    push(ev.importantElement);
    push(ev.detailsBetter);
    push(ev.action);
    push(ev.longTermGoals);
    push(ev.shortTermGoals);
    push(ev.miniGoals);
    for (const k of Object.keys(ev)) {
      if (
        k === "id" ||
        k === "activityId" ||
        k === "start" ||
        k === "people"
      )
        continue;
      const val = ev[k];
      if (typeof val === "string") push(val);
    }

    const pform = String(ev.projectsFromForm || "").trim();
    const pformLow = pform.toLowerCase();
    if (Array.isArray(state.projectsRegistry) && pformLow.length >= 2) {
      for (let ri = 0; ri < state.projectsRegistry.length; ri++) {
        const rp = reportNormLabel(state.projectsRegistry[ri].project);
        if (!rp || rp.length < 2) continue;
        const rl = rp.toLowerCase();
        let hit = false;
        if (pformLow.includes(rl)) hit = true;
        else if (pformLow.length >= 4 && rl.length >= 4 && rl.includes(pformLow)) hit = true;
        if (hit) push(rp);
      }
    }

    if (Array.isArray(listForInfer) && listForInfer.length) {
      const inf = reportInferredMapping(ev, listForInfer);
      push(inf.group);
      push(inf.layer);
      push(inf.cat);
      push(inf.subCat);
    }

    return parts.join(" ");
  }

  /** Search: tokens (spaces / commas), AND, case-insensitive substring match on joined row text. */
  function eventMatchesKeywordSearch(ev, f, list) {
    const q = String(f.keywordQuery || "").trim();
    if (!q) return true;
    const hayJoined = eventKeywordSearchHaystack(ev, list);
    const tokens = reportKeywordTokens(f.keywordQuery);
    if (!tokens.length) return true;
    const hay = normalizeForKeywordLoose(hayJoined);
    return tokens.every((tok) => hay.includes(normalizeForKeywordLoose(tok)));
  }

  /** For empty-report hints: filters other than search (AND). */
  function reportNonKeywordFilterSummaryText(f) {
    const bits = [];
    if (f.groups && f.groups.length) bits.push("Group: " + f.groups.join(", "));
    if (f.layers && f.layers.length) bits.push("Layers: " + f.layers.join(", "));
    if (f.cats && f.cats.length) bits.push("Category: " + f.cats.join(", "));
    if (f.subCats && f.subCats.length) bits.push("Sub Category: " + f.subCats.join(", "));
    if (f.projects && f.projects.length) bits.push("Project: " + f.projects.join(", "));
    if (f.activities && f.activities.length) bits.push("Activity: " + f.activities.join(", "));
    const ptoks = reportPeopleTokens(f.peopleQuery);
    if (ptoks.length) bits.push("With: " + ptoks.join(", "));
    return bits.join(" · ");
  }

  function eventMatchesPeopleSearch(ev, query) {
    const tokens = reportPeopleTokens(query);
    if (!tokens.length) return true;
    const people = (ev.people || []).map((p) => String(p).trim().toLowerCase()).filter(Boolean);
    if (!people.length) return false;
    return tokens.some((tok) => people.some((p) => p.includes(tok) || tok.includes(p)));
  }

  function readReportFilters() {
    const gpeo = document.getElementById("reportPeopleSearch");
    const kwEl = document.getElementById("reportKeywordSearch");
    return {
      groups: [...readCheckedValuesFromMsBox("reportFilterGroupBox")],
      layers: [...readCheckedValuesFromMsBox("reportFilterLayerBox")],
      cats: [...readCheckedValuesFromMsBox("reportFilterCatBox")],
      subCats: [...readCheckedValuesFromMsBox("reportFilterSubBox")],
      projects: [...readCheckedValuesFromMsBox("reportFilterProjectBox")],
      activities: [...readCheckedValuesFromMsBox("reportFilterActivityBox")],
      peopleQuery: (gpeo && gpeo.value) || "",
      keywordQuery: (kwEl && kwEl.value) || "",
      keywordMode: "loose",
    };
  }

  function eventMatchesReportFilters(ev, f, list) {
    const inf = reportInferredMapping(ev, list);
    if (f.groups && f.groups.length) {
      const gv = reportNormLabel(inf.group);
      const ok = f.groups.some((s) => reportNormLabel(s) === gv);
      if (!ok) return false;
    }
    if (f.layers && f.layers.length) {
      const lv = reportNormLabel(inf.layer);
      const ok = f.layers.some((s) => reportNormLabel(s) === lv);
      if (!ok) return false;
    }
    if (f.cats && f.cats.length) {
      const raw = reportNormLabel(inf.cat);
      const cv = (normalizeCatDisplayForRaw(raw) || raw).toLowerCase();
      const ok = f.cats.some((s) => {
        const sn = reportNormLabel(s);
        const snNorm = (normalizeCatDisplayForRaw(sn) || sn).toLowerCase();
        return snNorm === cv || sn.toLowerCase() === cv;
      });
      if (!ok) return false;
    }
    if (f.subCats && f.subCats.length) {
      const ok = f.subCats.some((s) => eventMatchesSubFilterInferred(ev, s, list));
      if (!ok) return false;
    }
    if (f.projects && f.projects.length) {
      if (!eventMatchesProjectReportFilter(ev, f.projects)) return false;
    }
    if (f.activities && f.activities.length) {
      const actName = reportNormLabel(activityDisplayName(ev.activityId));
      const ok = f.activities.some((s) => reportNormLabel(s) === actName);
      if (!ok) return false;
    }
    if (!eventMatchesPeopleSearch(ev, f.peopleQuery)) return false;
    if (!eventMatchesKeywordSearch(ev, f, list)) return false;
    return true;
  }

  function reportHasAnyFilter(f) {
    return !!(
      (f.groups && f.groups.length) ||
      (f.layers && f.layers.length) ||
      (f.cats && f.cats.length) ||
      (f.subCats && f.subCats.length) ||
      (f.projects && f.projects.length) ||
      (f.activities && f.activities.length) ||
      reportPeopleTokens(f.peopleQuery).length ||
      reportKeywordActive(f)
    );
  }

  function compareModeAnchorYmd(preset, toYmd, fromYmd) {
    if (preset === "cmp_weeks") return ymdFromLocalDate(new Date());
    return toYmd || fromYmd || ymdFromLocalDate(new Date());
  }

  function aggregateReportForRange(fromYmd, toYmd, list, f, showByDay, ctxOpt) {
    if (!fromYmd || !toYmd) return null;
    const t0 = new Date(fromYmd + "T00:00:00").getTime();
    const t1 = new Date(toYmd + "T23:59:59.999").getTime();
    if (Number.isNaN(t0) || Number.isNaN(t1) || t0 > t1) return null;
    const ctx = ctxOpt || buildReportMappingContext(list);
    const rawSegmentRows = [];
    const byEnt = {};
    const byGroup = {};
    const byLayer = {};
    const byProject = {};
    const byCatDim = {};
    const bySubDim = {};
    const byDay = {};
    const byPerson = {};
    let segmentsInRange = 0;
    let segmentsKept = 0;
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const st = new Date(ev.start).getTime();
      if (st < t0 || st > t1) continue;
      const ms = segmentDurationMsForReport(list, i);
      segmentsInRange++;
      if (!eventMatchesReportFilters(ev, f, list)) continue;
      segmentsKept++;
      rawSegmentRows.push({ ev, ms });
      byEnt[ev.activityId] = (byEnt[ev.activityId] || 0) + ms;
      const inf = ctx.inferred.get(ev.id) || reportInferredMapping(ev, list);
      const g = inf.group || "\uff08\u672a\u6a19 Group\uff09";
      const ly = inf.layer || "\u2014";
      byGroup[g] = (byGroup[g] || 0) + ms;
      byLayer[ly] = (byLayer[ly] || 0) + ms;
      const pj = reportNormLabel(ev.projectsFromForm || "") || "blank";
      const cj = inf.cat || "\uff08\u7a7a\uff0f\u672a\u914d Cat\uff09";
      const sj = inf.subCat || "\uff08\u7a7a\uff0f\u672a\u914d Sub\uff09";
      byProject[pj] = (byProject[pj] || 0) + ms;
      byCatDim[cj] = (byCatDim[cj] || 0) + ms;
      bySubDim[sj] = (bySubDim[sj] || 0) + ms;
      {
        const ymd = ymdFromLocalDate(new Date(ev.start));
        byDay[ymd] = (byDay[ymd] || 0) + ms;
      }
      const ppl = ev.people || [];
      for (let pi = 0; pi < ppl.length; pi++) {
        const nm = reportNormLabel(ppl[pi]);
        if (!nm) continue;
        byPerson[nm] = (byPerson[nm] || 0) + ms;
      }
    }
    let totalKept = 0;
    for (const k of Object.keys(byEnt)) totalKept += byEnt[k];
    return {
      t0,
      t1,
      rawSegmentRows,
      byEnt,
      byGroup,
      byLayer,
      byProject,
      byCatDim,
      bySubDim,
      byDay,
      byPerson,
      segmentsInRange,
      segmentsKept,
      totalKept,
    };
  }

  function csvEscapeReportCell(val) {
    const s = String(val ?? "");
    if (/[",\r\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  const REPORT_DATA_CSV_HEADERS = [
    "Start",
    "Duration",
    "Distraction",
    "Group",
    "Layers",
    "Cat",
    "Sub Cat",
    "Activity",
    "Project",
    "Place",
    "Remark",
    "With",
  ];

  function reportDataCsvRowCells(ev, ms, nextGlobal) {
    const startStr = ymdHmFromEventStart(ev.start);
    const durStr = durationMinutesLabel(ms);
    const distractStr = formatDistractionMmSs_(ev.distractionSec);
    const nextEv = nextGlobal.get(ev.id) || null;
    const mapped = inferTimeStatMappingForRaw(ev, nextEv);
    const g = String(mapped.group || "").trim() || "\u2014";
    const ly = String(mapped.layer || "").trim() || "\u2014";
    const cj = normalizeCatDisplayForRaw(String(mapped.cat || "").trim()) || "\u2014";
    const sj = String(mapped.subCat || "").trim() || "\u2014";
    const act = activityDisplayName(ev.activityId);
    const pj = displayProjectForRawRecord(ev).trim() || "\u2014";
    const place = String(ev.place || "").trim() || "\u2014";
    const remark = displayRemarkForRawRecord(ev).trim() || "\u2014";
    const withStr = ev.people && ev.people.length ? ev.people.join(", ") : "\u2014";
    return [startStr, durStr, distractStr, g, ly, cj, sj, act, pj, place, remark, withStr];
  }

  function buildReportDataCsvText(list, fromYmd, toYmd, f, showByDay) {
    const agg = aggregateReportForRange(fromYmd, toYmd, list, f, showByDay);
    if (!agg || agg.segmentsKept === 0) return null;
    const sortedRaw = [...agg.rawSegmentRows].sort((a, b) => new Date(a.ev.start) - new Date(b.ev.start));
    const nextGlobal = chronologicalNextById(list);
    const rows = [REPORT_DATA_CSV_HEADERS.map(csvEscapeReportCell).join(",")];
    for (let ri = 0; ri < sortedRaw.length; ri++) {
      const ev = sortedRaw[ri].ev;
      const ms = sortedRaw[ri].ms;
      rows.push(reportDataCsvRowCells(ev, ms, nextGlobal).map(csvEscapeReportCell).join(","));
    }
    return { text: "\ufeff" + rows.join("\r\n") + "\r\n", rowCount: sortedRaw.length };
  }

  function buildReportCompareResultsCsvText(list, preset, f, reportUnitMode) {
    const fromEl = document.getElementById("reportFromStr");
    const toEl = document.getElementById("reportToStr");
    const toY = parseYMDStrict(toEl && toEl.value);
    const fromY = parseYMDStrict(fromEl && fromEl.value);
    const anchorYmd = compareModeAnchorYmd(preset, toY, fromY);
    const slices = buildReportComparisonSlices(preset, anchorYmd);
    if (!slices || !slices.ranges || slices.ranges.length !== 3) return null;

    const ag = [
      aggregateReportForRange(slices.ranges[0].from, slices.ranges[0].to, list, f, false),
      aggregateReportForRange(slices.ranges[1].from, slices.ranges[1].to, list, f, false),
      aggregateReportForRange(slices.ranges[2].from, slices.ranges[2].to, list, f, false),
    ];
    if (ag[0] === null || ag[1] === null || ag[2] === null) return null;
    if (ag[0].segmentsInRange + ag[1].segmentsInRange + ag[2].segmentsInRange === 0) return null;
    if (ag[0].segmentsKept + ag[1].segmentsKept + ag[2].segmentsKept === 0) return null;

    const agDisp = [ag[2], ag[1], ag[0]];
    const labelsDisp = [slices.labels[2], slices.labels[1], slices.labels[0]];
    const colTotals = [agDisp[0].totalKept, agDisp[1].totalKept, agDisp[2].totalKept];
    const fmtCell = (ms, colIdx) =>
      colTotals[colIdx] ? ((ms / colTotals[colIdx]) * 100).toFixed(1) + "%" : "\u2014";
    const fmtH = (ms) => (ms / 3600000).toFixed(2);
    const fmtCmp = (ms, colIdx) => (reportUnitMode === "pct" ? fmtCell(ms, colIdx) : fmtH(ms));

    const out = [];
    out.push(["Export", "Compare Results"].map(csvEscapeReportCell).join(","));
    out.push(["Preset", preset].map(csvEscapeReportCell).join(","));
    out.push(["Display Unit", reportUnitMode === "pct" ? "%" : "Hours"].map(csvEscapeReportCell).join(","));

    const pushSection = (sectionTitle, pick) => {
      const m0 = pick(agDisp[0]);
      const m1 = pick(agDisp[1]);
      const m2 = pick(agDisp[2]);
      const keys = new Set([...Object.keys(m0), ...Object.keys(m1), ...Object.keys(m2)]);
      const arr = [...keys].map((k) => {
        const a = m0[k] || 0;
        const b = m1[k] || 0;
        const c = m2[k] || 0;
        return { k, ms: [a, b, c], sum: a + b + c };
      });
      arr.sort((x, y) => y.sum - x.sum);
      out.push("");
      out.push(csvEscapeReportCell(sectionTitle));
      out.push(["Item", ...labelsDisp].map(csvEscapeReportCell).join(","));
      if (!arr.length) {
        out.push(["(None)", "\u2014", "\u2014", "\u2014"].map(csvEscapeReportCell).join(","));
        return 1;
      }
      let n = 0;
      for (let i = 0; i < arr.length; i++) {
        const row = arr[i];
        const cells = [row.k, ...row.ms.map((ms, j) => fmtCmp(ms, j))];
        out.push(cells.map(csvEscapeReportCell).join(","));
        n++;
      }
      return n;
    };

    let dataRows = 0;
    dataRows += pushSection("Group", (a) => a.byGroup);
    dataRows += pushSection("Layers", (a) => a.byLayer);
    dataRows += pushSection("Category", (a) => a.byCatDim);
    dataRows += pushSection("Sub Category", (a) => a.bySubDim);

    const actMap = new Map();
    for (let j = 0; j < 3; j++) {
      Object.entries(agDisp[j].byEnt).forEach(([eid, ms]) => {
        if (!actMap.has(eid)) actMap.set(eid, [0, 0, 0]);
        actMap.get(eid)[j] = ms;
      });
    }
    const actRows = [...actMap.entries()]
      .map(([eid, msco]) => ({ eid, msco, sum: msco[0] + msco[1] + msco[2] }))
      .sort((a, b) => b.sum - a.sum);
    out.push("");
    out.push(csvEscapeReportCell("Activity"));
    out.push(["Item", ...labelsDisp].map(csvEscapeReportCell).join(","));
    if (!actRows.length) {
      out.push(["(None)", "\u2014", "\u2014", "\u2014"].map(csvEscapeReportCell).join(","));
      dataRows += 1;
    } else {
      for (let i = 0; i < actRows.length; i++) {
        const row = actRows[i];
        const name = activityDisplayName(row.eid);
        const cells = [name, ...row.msco.map((ms, j) => fmtCmp(ms, j))];
        out.push(cells.map(csvEscapeReportCell).join(","));
        dataRows++;
      }
    }

    dataRows += pushSection("Project", (a) => a.byProject);
    dataRows += pushSection("People", (a) => a.byPerson);

    return {
      text: "\ufeff" + out.join("\r\n") + "\r\n",
      rowCount: dataRows,
      fileStem: `${slices.ranges[2].from}_${slices.ranges[0].to}`,
    };
  }

  function runReportDataExport() {
    const presetEl = document.getElementById("reportPeriodPreset");
    const preset = (presetEl && presetEl.value) || "custom";
    const f = readReportFilters();
    const list = sortedEventsUniqueById();
    const showByDayEl = document.getElementById("reportShowByDay");
    const showByDay = !!(showByDayEl && showByDayEl.checked);
    let pack = null;
    let downloadName = "";
    if (reportComparePresetActive(preset)) {
      const reportUnitMode = readReportUnitMode();
      pack = buildReportCompareResultsCsvText(list, preset, f, reportUnitMode);
      if (!pack) {
        toast("No Compare Results To Export For These Windows And Filters.");
        return;
      }
      downloadName = `time-stat-compare-results-${preset}-${pack.fileStem}.csv`;
    } else {
      const from = parseYMDStrict(document.getElementById("reportFromStr").value);
      const to = parseYMDStrict(document.getElementById("reportToStr").value);
      if (!from || !to) {
        toast("Set The Date Range First.");
        return;
      }
      const t0 = new Date(from + "T00:00:00").getTime();
      const t1 = new Date(to + "T23:59:59.999").getTime();
      if (Number.isNaN(t0) || Number.isNaN(t1) || t0 > t1) {
        toast("Start Date Must Be On Or Before End Date.");
        return;
      }
      pack = buildReportDataCsvText(list, from, to, f, showByDay);
      if (!pack) {
        toast("No Data Rows To Export For This Range And Filters.");
        return;
      }
      downloadName = `time-stat-data-${from}_to_${to}.csv`;
    }
    const blob = new Blob([pack.text], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = downloadName;
    a.click();
    URL.revokeObjectURL(a.href);
    toast(
      reportComparePresetActive(preset)
        ? `Downloaded Compare Results (${pack.rowCount} Rows)`
        : `Downloaded ${pack.rowCount} Rows`,
    );
  }

  function buildReportComparisonSlices(preset, anchorYmd) {
    const m = String(anchorYmd || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return null;
    const y = parseInt(m[1], 10);
    const mo = parseInt(m[2], 10) - 1;
    const d = parseInt(m[3], 10);
    const pad = (n) => String(n).padStart(2, "0");
    const ymd = (yy, mm, dd) => `${yy}-${pad(mm)}-${pad(dd)}`;
    const lastDay = (yy, m0) => new Date(yy, m0 + 1, 0).getDate();

    if (preset === "cmp_years") {
      const y3 = y;
      const y2 = y - 1;
      const y1 = y - 2;
      return {
        labels: [String(y3), String(y2), String(y1)],
        ranges: [
          { from: `${y3}-01-01`, to: `${y3}-12-31` },
          { from: `${y2}-01-01`, to: `${y2}-12-31` },
          { from: `${y1}-01-01`, to: `${y1}-12-31` },
        ],
      };
    }
    if (preset === "cmp_months") {
      const decM = (yy, m0) => {
        if (m0 > 0) return { yy, m0: m0 - 1 };
        return { yy: yy - 1, m0: 11 };
      };
      const b3 = { yy: y, m0: mo };
      const b2 = decM(b3.yy, b3.m0);
      const b1 = decM(b2.yy, b2.m0);
      const bounds = (bb) => {
        const ld = lastDay(bb.yy, bb.m0);
        return { from: ymd(bb.yy, bb.m0 + 1, 1), to: ymd(bb.yy, bb.m0 + 1, ld) };
      };
      const r1 = bounds(b1);
      const r2 = bounds(b2);
      const r3 = bounds(b3);
      const lab = (bb) => `${bb.yy}-${pad(bb.m0 + 1)}`;
      return { labels: [lab(b3), lab(b2), lab(b1)], ranges: [r3, r2, r1] };
    }
    if (preset === "cmp_quarters") {
      const qOf = (m0) => Math.floor(m0 / 3);
      const decQ = (yy, q) => {
        if (q > 0) return { yy, q: q - 1 };
        return { yy: yy - 1, q: 3 };
      };
      const qb = (yy, q) => {
        const m0 = q * 3;
        const ld = lastDay(yy, m0 + 2);
        return { from: ymd(yy, m0 + 1, 1), to: ymd(yy, m0 + 3, ld) };
      };
      const q = qOf(mo);
      const b3 = { yy: y, q };
      const b2 = decQ(b3.yy, b3.q);
      const b1 = decQ(b2.yy, b2.q);
      const r1 = qb(b1.yy, b1.q);
      const r2 = qb(b2.yy, b2.q);
      const r3 = qb(b3.yy, b3.q);
      const lab = (bb) => `${bb.yy} Q${bb.q + 1}`;
      return { labels: [lab(b3), lab(b2), lab(b1)], ranges: [r3, r2, r1] };
    }
    if (preset === "cmp_weeks") {
      const dow = new Date(y, mo, d).getDay();
      const diffToMon = (dow + 6) % 7;
      const mon3 = new Date(y, mo, d - diffToMon);
      const mon2 = new Date(mon3);
      mon2.setDate(mon3.getDate() - 7);
      const mon1 = new Date(mon2);
      mon1.setDate(mon2.getDate() - 7);
      const pack = (mon) => {
        const sun = new Date(mon);
        sun.setDate(mon.getDate() + 6);
        return { from: ymdFromLocalDate(mon), to: ymdFromLocalDate(sun) };
      };
      const p1 = pack(mon1);
      const p2 = pack(mon2);
      const p3 = pack(mon3);
      const fmtShort = (fr, to) => `${fr.slice(5)} \u2192 ${to.slice(5)}`;
      return {
        labels: [fmtShort(p3.from, p3.to), fmtShort(p2.from, p2.to), fmtShort(p1.from, p1.to)],
        ranges: [p3, p2, p1],
      };
    }
    return null;
  }

  function reportComparePresetActive(preset) {
    return ["cmp_years", "cmp_quarters", "cmp_months", "cmp_weeks"].indexOf(preset || "") >= 0;
  }

  function applyReportPeriodPreset() {
    const presetEl = document.getElementById("reportPeriodPreset");
    const fromEl = document.getElementById("reportFromStr");
    const toEl = document.getElementById("reportToStr");
    if (!presetEl || !fromEl || !toEl) return;
    const preset = presetEl.value || "custom";
    if (preset === "custom") return;
    if (!reportComparePresetActive(preset)) return;
    const anchorYmd = compareModeAnchorYmd(
      preset,
      parseYMDStrict(toEl.value),
      parseYMDStrict(fromEl.value),
    );
    const slices = buildReportComparisonSlices(preset, anchorYmd);
    if (!slices || !slices.ranges || slices.ranges.length !== 3) return;
    reportPresetSuppress = true;
    try {
      // Date pickers = focus period（最新一期）；對比表仍用三期，Raw Data 跟呢個範圍
      fromEl.value = slices.ranges[0].from;
      toEl.value = slices.ranges[0].to;
    } finally {
      queueMicrotask(() => {
        reportPresetSuppress = false;
      });
    }
  }

  function readReportUnitMode() {
    try {
      const v = localStorage.getItem("timeStatReportUnit");
      if (v === "pct" || v === "hours") return v;
    } catch (e) {}
    return "hours";
  }

  function setReportUnitMode(mode) {
    try {
      if (mode === "pct" || mode === "hours") localStorage.setItem("timeStatReportUnit", mode);
    } catch (e) {}
    syncReportUnitToggleButtons();
  }

  function syncReportUnitToggleButtons() {
    const wrap = document.getElementById("reportUnitToggle");
    if (!wrap) return;
    const mode = readReportUnitMode();
    wrap.querySelectorAll(".report-unit-btn").forEach((b) => {
      b.classList.toggle("active", b.getAttribute("data-unit") === mode);
    });
  }

  function renderReport() {
    refreshReportFilterSelectsIfNeeded();
    const presetElR = document.getElementById("reportPeriodPreset");
    const preset = (presetElR && presetElR.value) || "custom";
    if (!reportPresetSuppress && preset !== "custom") applyReportPeriodPreset();

    const from = parseYMDStrict(document.getElementById("reportFromStr").value);
    const to = parseYMDStrict(document.getElementById("reportToStr").value);
    const box = document.getElementById("reportSummary");
    const showByDayEl = document.getElementById("reportShowByDay");
    const showByDay = !!(showByDayEl && showByDayEl.checked);
    if (!box) return;
    try {
    if (!from || !to) {
      mountFocusEnergyChart_("", "");
      box.innerHTML = `<p class="muted">Please Set The <strong>Date Range</strong> (Start ～ End).</p>`;
      return;
    }
    const t0 = new Date(from + "T00:00:00").getTime();
    const t1 = new Date(to + "T23:59:59.999").getTime();
    if (t0 > t1) {
      mountFocusEnergyChart_("", "");
      box.innerHTML = `<p class="muted">Start Date Must Be On Or Before End Date.</p>`;
      return;
    }
    // Chart 獨立 host，唔放喺 #reportSummary（免被 AI busy／innerHTML 清走）
    mountFocusEnergyChart_(from, to);
    const f = readReportFilters();
    const list = sortedEventsUniqueById();
    const reportCtx = buildReportMappingContext(list);
    const reportUnitMode = readReportUnitMode();

    const renderFilterMismatch = () => {
      const kw = String(f.keywordQuery || "").trim();
      const nf = reportNonKeywordFilterSummaryText(f);
      const kwHint =
        kw.length > 0 ? `<p class="muted" style="margin-top:10px;">Search: ${escapeHtml(kw)}</p>` : "";
      const other =
        nf.length > 0 ? `<p class="muted" style="margin-top:8px;">${escapeHtml(nf)}</p>` : "";
      return (
        '<p class="muted">There Are Records In Range, But <strong>None Match Your Filters</strong>.</p>' +
        kwHint +
        other
      );
    };

    if (reportComparePresetActive(preset)) {
      const anchorYmd = compareModeAnchorYmd(preset, to, from);
      const slices = buildReportComparisonSlices(preset, anchorYmd);
      if (!slices) {
        box.innerHTML = `<p class="muted">Could Not Build This Multi-Window View.</p>`;
        return;
      }
      reportPresetSuppress = true;
      try {
        // Focus period only（唔好擴到三期全距，否則 Raw Data 會睇到較新／較舊季度）
        document.getElementById("reportFromStr").value = slices.ranges[0].from;
        document.getElementById("reportToStr").value = slices.ranges[0].to;
      } finally {
        queueMicrotask(() => {
          reportPresetSuppress = false;
        });
      }
      const ag = [
        aggregateReportForRange(slices.ranges[0].from, slices.ranges[0].to, list, f, false, reportCtx),
        aggregateReportForRange(slices.ranges[1].from, slices.ranges[1].to, list, f, false, reportCtx),
        aggregateReportForRange(slices.ranges[2].from, slices.ranges[2].to, list, f, false, reportCtx),
      ];
      if (ag[0] === null || ag[1] === null || ag[2] === null) {
        box.innerHTML = `<p class="muted">Invalid Time Windows.</p>`;
        return;
      }
      if (ag[0].segmentsInRange + ag[1].segmentsInRange + ag[2].segmentsInRange === 0) {
        box.innerHTML = `<p class="muted">No Billable Segments In The Last Three Periods.</p>`;
        return;
      }
      if (ag[0].segmentsKept + ag[1].segmentsKept + ag[2].segmentsKept === 0) {
        box.innerHTML = renderFilterMismatch();
        return;
      }

      const agDisp = [ag[2], ag[1], ag[0]];
      const labelsDisp = [slices.labels[2], slices.labels[1], slices.labels[0]];
      const colTotals = [agDisp[0].totalKept, agDisp[1].totalKept, agDisp[2].totalKept];
      const fmtCell = (ms, colIdx) =>
        colTotals[colIdx] ? ((ms / colTotals[colIdx]) * 100).toFixed(1) + "%" : "\u2014";
      const fmtH = (ms) => (ms / 3600000).toFixed(2);
      const fmtCmp = (ms, colIdx) => (reportUnitMode === "pct" ? fmtCell(ms, colIdx) : fmtH(ms));
      const thCols = labelsDisp
        .map((lb) => `<th class="mono report-cmp-col report-cmp-col-date">${escapeHtml(lb)}</th>`)
        .join("");

      const tblCmp = (title, pick) => {
        const m0 = pick(agDisp[0]);
        const m1 = pick(agDisp[1]);
        const m2 = pick(agDisp[2]);
        const keys = new Set([...Object.keys(m0), ...Object.keys(m1), ...Object.keys(m2)]);
        const arr = [...keys].map((k) => {
          const a = m0[k] || 0;
          const b = m1[k] || 0;
          const c = m2[k] || 0;
          return { k, ms: [a, b, c], sum: a + b + c };
        });
        arr.sort((x, y) => y.sum - x.sum);
        if (!arr.length) return `<h2 class="report-h">${escapeHtml(title)}</h2><p class="muted">(None)</p>`;
        let h =
          `<h2 class="report-h">${escapeHtml(title)}</h2><table class="report-cmp-table"><thead><tr><th>Item</th>${thCols}</tr></thead><tbody>`;
        for (let i = 0; i < arr.length; i++) {
          const row = arr[i];
          h += `<tr><td>${escapeHtml(row.k)}</td>`;
          for (let j = 0; j < 3; j++) {
            h += `<td class="mono">${fmtCmp(row.ms[j], j)}</td>`;
          }
          h += `</tr>`;
        }
        h += `</tbody></table>`;
        return h;
      };

      let html = "";
      if (reportHasAnyFilter(f)) {
        const bits = [];
        if (f.groups && f.groups.length) bits.push("Group: " + f.groups.join(" / "));
        if (f.layers && f.layers.length) bits.push("Layers: " + f.layers.join(" / "));
        if (f.cats && f.cats.length) bits.push("Category: " + f.cats.join(" / "));
        if (f.subCats && f.subCats.length) bits.push("Sub Category: " + f.subCats.join(" / "));
        if (f.projects && f.projects.length) bits.push("Project: " + f.projects.join(" / "));
        if (f.activities && f.activities.length) bits.push("Activity: " + f.activities.join(" / "));
        const ptoks = reportPeopleTokens(f.peopleQuery);
        if (ptoks.length) bits.push("People: " + ptoks.join(" / "));
        const kwTrim = String(f.keywordQuery || "").trim();
        if (kwTrim) {
          const kwtoks = reportKeywordTokens(f.keywordQuery);
          if (kwtoks.length) bits.push("Search: " + kwtoks.join(" And "));
          else bits.push("Search: " + kwTrim);
        }
        html += `<p class="muted" style="margin:0 0 12px;">${escapeHtml(bits.join(" · "))}</p>`;
      }
      html += tblCmp("Group", (a) => a.byGroup);
      html += tblCmp("Layers", (a) => a.byLayer);
      html += tblCmp("Category", (a) => a.byCatDim);
      html += tblCmp("Sub Category", (a) => a.bySubDim);

      const actMap = new Map();
      for (let j = 0; j < 3; j++) {
        Object.entries(agDisp[j].byEnt).forEach(([eid, ms]) => {
          if (!actMap.has(eid)) actMap.set(eid, [0, 0, 0]);
          actMap.get(eid)[j] = ms;
        });
      }
      const actRows = [...actMap.entries()]
        .map(([eid, msco]) => ({ eid, msco, sum: msco[0] + msco[1] + msco[2] }))
        .sort((a, b) => b.sum - a.sum);
      html += `<h2 class="report-h">Activity</h2><table class="report-cmp-table"><thead><tr><th>Activity</th>${thCols}</tr></thead><tbody>`;
      for (let i = 0; i < actRows.length; i++) {
        const row = actRows[i];
        html += `<tr><td>${escapeHtml(activityDisplayName(row.eid))}</td>`;
        for (let j = 0; j < 3; j++) {
          html += `<td class="mono">${fmtCmp(row.msco[j], j)}</td>`;
        }
        html += `</tr>`;
      }
      html += `</tbody></table>`;
      html += tblCmp("Project", (a) => a.byProject);
      html += tblCmp("People", (a) => a.byPerson);

      const focusFrom = slices.ranges[0].from;
      const focusTo = slices.ranges[0].to;
      html +=
        `<p class="muted" style="margin:16px 0 8px;">Raw Data · focus period ${escapeHtml(focusFrom)} ～ ${escapeHtml(focusTo)}（comparison columns stay 3 periods）</p>` +
        buildReportRawDataHtml_(ag[0].rawSegmentRows, reportCtx);

      mountFocusEnergyChart_(focusFrom, focusTo);
      box.innerHTML = html;
      bindReportRawDayFilter_();
      return;
    }

    const agg = aggregateReportForRange(from, to, list, f, showByDay, reportCtx);
    if (agg.segmentsInRange === 0) {
      box.innerHTML = `<p class="muted">No Billable Segments In This Range.</p>`;
      return;
    }
    if (agg.segmentsKept === 0) {
      const kw = String(f.keywordQuery || "").trim();
      const nf = reportNonKeywordFilterSummaryText(f);
      const kwHint =
        kw.length > 0
          ? `<p class="muted" style="margin-top:10px;">Search: ${escapeHtml(kw)}</p>`
          : "";
      const other =
        nf.length > 0
          ? `<p class="muted" style="margin-top:8px;">${escapeHtml(nf)}</p>`
          : `<p class="muted" style="margin-top:8px;">No Group, Layers, Category, Sub Category, Project, Activity, Or People Filters Are Set — Only Search Narrows Results.</p>`;
      box.innerHTML =
        `<p class="muted">There Are Records With Duration In This Range, But <strong>None Match Your Filters</strong>.</p>` +
        kwHint +
        other;
      return;
    }

    const { byEnt, byGroup, byLayer, byProject, byCatDim, bySubDim, byPerson, rawSegmentRows } = agg;
    const rows = Object.entries(byEnt).sort((a, b) => b[1] - a[1]);
    const aggTotalMs = agg.totalKept || 0;
    const fmtAggCell = (ms) => {
      if (reportUnitMode === "pct") {
        return aggTotalMs ? ((ms / aggTotalMs) * 100).toFixed(1) + "%" : "\u2014";
      }
      return (ms / 3600000).toFixed(2);
    };
    const unitTh = reportUnitMode === "pct" ? "%" : "Hours";
    const tbl = (title, map) => {
      const r = Object.entries(map).sort((a, b) => b[1] - a[1]);
      if (!r.length) return `<h2 class="report-h">${title}</h2><p class="muted">(None)</p>`;
      let h = `<h2 class="report-h">${title}</h2><table><thead><tr><th>Item</th><th>${unitTh}</th></tr></thead><tbody>`;
      for (let i = 0; i < r.length; i++) {
        const k = r[i][0];
        const ms = r[i][1];
        h += `<tr><td>${escapeHtml(k)}</td><td class="mono">${fmtAggCell(ms)}</td></tr>`;
      }
      h += `</tbody></table>`;
      return h;
    };
    let html = "";
    if (reportHasAnyFilter(f)) {
      const bits = [];
      if (f.groups && f.groups.length) bits.push("Group: " + f.groups.join(" / "));
      if (f.layers && f.layers.length) bits.push("Layers: " + f.layers.join(" / "));
      if (f.cats && f.cats.length) bits.push("Category: " + f.cats.join(" / "));
      if (f.subCats && f.subCats.length) bits.push("Sub Category: " + f.subCats.join(" / "));
      if (f.projects && f.projects.length) bits.push("Project: " + f.projects.join(" / "));
      if (f.activities && f.activities.length) bits.push("Activity: " + f.activities.join(" / "));
      const ptoks = reportPeopleTokens(f.peopleQuery);
      if (ptoks.length) bits.push("People: " + ptoks.join(" / "));
      const kwTrim = String(f.keywordQuery || "").trim();
      if (kwTrim) {
        const kwtoks = reportKeywordTokens(f.keywordQuery);
        if (kwtoks.length) bits.push("Search: " + kwtoks.join(" And "));
        else bits.push("Search: " + kwTrim);
      }
      html += `<p class="muted" style="margin:0 0 12px;">${escapeHtml(bits.join(" · "))}</p>`;
    }
    html += tbl("Group", byGroup);
    html += tbl("Layers", byLayer);
    html += tbl("Category", byCatDim);
    html += tbl("Sub Category", bySubDim);
    html += `<h2 class="report-h">Activity</h2><table><thead><tr><th>Activity</th><th>${unitTh}</th></tr></thead><tbody>`;
    for (let i = 0; i < rows.length; i++) {
      const eid = rows[i][0];
      const ms = rows[i][1];
      html += `<tr><td>${escapeHtml(activityDisplayName(eid))}</td><td class="mono">${fmtAggCell(ms)}</td></tr>`;
    }
    html += `</tbody></table>`;
    html += tbl("Project", byProject);
    html += tbl("People", byPerson);
    html += buildReportRawDataHtml_(rawSegmentRows, reportCtx);

    box.innerHTML = html;
    bindReportRawDayFilter_();
    } finally {
      syncReportUnitToggleButtons();
    }
  }

  let _reportRawDayFilter_ = "";

  function buildReportRawDataHtml_(rawSegmentRows, reportCtx) {
    const sortedRaw = [...(rawSegmentRows || [])].sort(
      (a, b) => new Date(a.ev.start) - new Date(b.ev.start),
    );
    const rawDayKeys = [];
    const rawDaySeen = Object.create(null);
    for (let ri0 = 0; ri0 < sortedRaw.length; ri0++) {
      const y = ymdFromLocalDate(new Date(sortedRaw[ri0].ev.start));
      if (!rawDaySeen[y]) {
        rawDaySeen[y] = 1;
        rawDayKeys.push(y);
      }
    }
    rawDayKeys.sort();
    if (_reportRawDayFilter_ && !rawDaySeen[_reportRawDayFilter_]) {
      _reportRawDayFilter_ = "";
    }

    let html = `<h2 class="report-h" id="reportRawDataHeading">Data</h2>`;
    html +=
      `<div class="report-raw-toolbar">` +
      `<label for="reportRawDayFilter" class="report-raw-day-label">Day</label>` +
      `<select id="reportRawDayFilter" class="input-xl report-raw-day-select" aria-label="Filter raw data by day">` +
      `<option value="">All days</option>`;
    for (let di = 0; di < rawDayKeys.length; di++) {
      const d = rawDayKeys[di];
      const sel = d === _reportRawDayFilter_ ? " selected" : "";
      html += `<option value="${escapeHtml(d)}"${sel}>${escapeHtml(d)}</option>`;
    }
    html += `</select></div>`;

    html += `<div class="report-records-wrap" id="reportRecordsWrap"><table class="report-records-table"><thead><tr><th>Start</th><th>Duration</th><th>Distraction</th><th>Group</th><th>Layers</th><th>Cat</th><th>Sub Cat</th><th>Activity</th><th>Project</th><th>Place</th><th>Remark</th><th>With</th></tr></thead><tbody>`;
    for (let ri = 0; ri < sortedRaw.length; ri++) {
      const row = sortedRaw[ri];
      const ev = row.ev;
      const ms = row.ms;
      const startStr = ymdHmFromEventStart(ev.start);
      const rowYmd = ymdFromLocalDate(new Date(ev.start));
      const durStr = durationMinutesLabel(ms);
      const distractStr = formatDistractionMmSs_(ev.distractionSec);
      const mapped = reportCtx.inferred.get(ev.id);
      const g = mapped ? String(mapped.group || "").trim() || "\u2014" : "\u2014";
      const ly = mapped ? String(mapped.layer || "").trim() || "\u2014" : "\u2014";
      const cj = mapped ? String(mapped.cat || "").trim() || "\u2014" : "\u2014";
      const sj = mapped ? String(mapped.subCat || "").trim() || "\u2014" : "\u2014";
      const act = activityDisplayName(ev.activityId);
      const pj = displayProjectForRawRecord(ev).trim() || "blank";
      const place = String(ev.place || "").trim() || "\u2014";
      const remark = displayRemarkForRawRecord(ev).trim() || "\u2014";
      const withStr = ev.people && ev.people.length ? ev.people.join(", ") : "\u2014";
      const hide =
        _reportRawDayFilter_ && rowYmd !== _reportRawDayFilter_ ? " hidden" : "";
      html +=
        `<tr data-ymd="${escapeHtml(rowYmd)}"${hide}><td class="mono">${escapeHtml(startStr)}</td><td class="mono">${escapeHtml(durStr)}</td>` +
        `<td class="mono">${escapeHtml(distractStr)}</td>` +
        `<td>${escapeHtml(g)}</td><td>${escapeHtml(ly)}</td><td>${escapeHtml(cj)}</td><td>${escapeHtml(sj)}</td>` +
        `<td>${escapeHtml(act)}</td><td>${escapeHtml(pj)}</td><td>${escapeHtml(place)}</td><td class="remark-cell">${escapeHtml(remark)}</td><td>${escapeHtml(withStr)}</td></tr>`;
    }
    html += `</tbody></table></div>`;
    return html;
  }

  function bindReportRawDayFilter_() {
    const sel = document.getElementById("reportRawDayFilter");
    if (!sel || sel.dataset.bound === "1") return;
    sel.dataset.bound = "1";
    sel.addEventListener("change", () => {
      applyReportRawDayFilter_(sel.value || "", { scroll: !!sel.value });
    });
  }

  function applyReportRawDayFilter_(ymd, opts) {
    const o = opts || {};
    const next = String(ymd || "").trim();
    _reportRawDayFilter_ = next;
    const sel = document.getElementById("reportRawDayFilter");
    if (sel && sel.value !== next) {
      const hasOpt = !next || [...sel.options].some((op) => op.value === next);
      if (hasOpt) sel.value = next;
      else {
        toast("No raw rows for " + next);
        _reportRawDayFilter_ = "";
        sel.value = "";
        return;
      }
    }
    const wrap = document.getElementById("reportRecordsWrap");
    if (!wrap) return;
    wrap.querySelectorAll("tbody tr[data-ymd]").forEach((tr) => {
      const rowYmd = tr.getAttribute("data-ymd") || "";
      tr.hidden = !!(next && rowYmd !== next);
      tr.classList.remove("report-raw-day-hl");
    });
    if (o.scroll && next) scrollReportRawToYmd_(next);
  }

  function jumpReportRawToYmd_(ymd) {
    const target = String(ymd || "").trim();
    if (!target) return;
    const wrap = document.getElementById("reportRecordsWrap");
    if (!wrap) {
      toast("Open Report Data first.");
      return;
    }
    applyReportRawDayFilter_(target, { scroll: true });
  }

  function scrollReportRawToYmd_(ymd) {
    const wrap = document.getElementById("reportRecordsWrap");
    const box = document.getElementById("reportSummary");
    if (!wrap) return;
    wrap.querySelectorAll("tr.report-raw-day-hl").forEach((tr) => tr.classList.remove("report-raw-day-hl"));
    const row = wrap.querySelector('tr[data-ymd="' + ymd.replace(/"/g, "") + '"]:not([hidden])');
    if (!row) {
      toast("No raw rows for " + ymd);
      return;
    }
    row.classList.add("report-raw-day-hl");
    const heading = document.getElementById("reportRawDataHeading");
    if (heading && typeof heading.scrollIntoView === "function") {
      heading.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
    const results = document.querySelector(".report-results");
    if (results && box && results.contains(box)) {
      const top =
        wrap.getBoundingClientRect().top - results.getBoundingClientRect().top + results.scrollTop - 8;
      results.scrollTo({ top: Math.max(0, top), behavior: "smooth" });
    }
    const rowTop = row.offsetTop - (wrap.querySelector("thead")?.offsetHeight || 0) - 4;
    wrap.scrollTo({ top: Math.max(0, rowTop), behavior: "smooth" });
  }

  function csv(s) {
    const t = String(s).replace(/"/g, '""');
    if (/[",\n\r]/.test(t)) return '"' + t + '"';
    return t;
  }

  document.querySelectorAll(".tabs button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tab = btn.getAttribute("data-tab");
      hideTimelineTip();
      document.querySelectorAll(".tabs button").forEach((b) => {
        b.classList.toggle("active", b === btn);
      });
      document.querySelectorAll("[data-panel]").forEach((p) => {
        p.classList.toggle("hidden", p.getAttribute("data-panel") !== tab);
      });
      if (tab === "report") renderReport();
      if (tab === "timeline") renderTimeline();
      if (tab === "ai-reports") {
        void loadAiReportsList_();
        void loadAiSettingsForm_();
      }
      refreshSoftCapBanner();
    });
  });

  let _lastManualAiReport_ = null;

  function resetAiFollowUpUi_() {
    const wrap = document.getElementById("aiReportFollowUp");
    const thread = document.getElementById("aiReportFollowUpThread");
    const input = document.getElementById("aiReportFollowUpInput");
    if (wrap) wrap.classList.add("hidden");
    if (thread) thread.innerHTML = "";
    if (input) {
      input.value = "";
      input.disabled = false;
    }
    const ask = document.getElementById("btnAiReportFollowUpAsk");
    if (ask) ask.disabled = false;
  }

  function showAiFollowUpUi_() {
    const wrap = document.getElementById("aiReportFollowUp");
    if (wrap) wrap.classList.remove("hidden");
  }

  function appendAiFollowUpMessage_(role, markdown, opts) {
    const thread = document.getElementById("aiReportFollowUpThread");
    if (!thread) return;
    const o = opts || {};
    const row = document.createElement("div");
    row.className = "ai-followup-msg " + (role === "user" ? "user" : "assistant");
    const roleEl = document.createElement("span");
    roleEl.className = "ai-followup-role";
    roleEl.textContent = role === "user" ? "You" : "AI";
    row.appendChild(roleEl);
    const body = document.createElement("div");
    body.className = "ai-report-md";
    if (o.loading) {
      body.innerHTML = '<p class="ai-report-loading">Thinking…</p>';
    } else if (o.error) {
      body.innerHTML = '<p class="ai-report-loading">' + escapeHtml(String(o.error)) + "</p>";
    } else if (role === "user") {
      body.innerHTML = "<p>" + escapeHtml(String(markdown || "")).replace(/\n/g, "<br>") + "</p>";
    } else {
      body.innerHTML = renderAiMarkdownToHtml_(markdown || "", { linkDates: true });
      bindAiManualDateJumps_(body);
    }
    row.appendChild(body);
    thread.appendChild(row);
    thread.scrollTop = thread.scrollHeight;
    return row;
  }

  async function askAiReportFollowUp_() {
    if (!canRemoteSync()) {
      toast("請先 Google 登入。");
      return;
    }
    if (!_lastManualAiReport_ || !_lastManualAiReport_.markdown) {
      toast("請先 Generate AI Report。");
      return;
    }
    const input = document.getElementById("aiReportFollowUpInput");
    const askBtn = document.getElementById("btnAiReportFollowUpAsk");
    const question = String(input && input.value ? input.value : "").trim();
    if (!question) {
      toast("請輸入問題。");
      return;
    }
    if (!Array.isArray(_lastManualAiReport_.history)) _lastManualAiReport_.history = [];
    appendAiFollowUpMessage_("user", question);
    if (input) {
      input.value = "";
      input.disabled = true;
    }
    if (askBtn) askBtn.disabled = true;
    setAiReportBusy_(true);
    const thinkingRow = appendAiFollowUpMessage_("assistant", "", { loading: true });
    try {
      const j = await postAiAction_({
        action: "askAiReportFollowUp",
        periodType: _lastManualAiReport_.periodType,
        periodKey: _lastManualAiReport_.periodKey,
        question: question,
        reportMarkdown: _lastManualAiReport_.markdown,
        history: _lastManualAiReport_.history.slice(-8),
      });
      const answer = j.markdown || "";
      _lastManualAiReport_.history.push({ role: "user", content: question });
      _lastManualAiReport_.history.push({ role: "assistant", content: answer });
      if (thinkingRow && thinkingRow.parentNode) thinkingRow.parentNode.removeChild(thinkingRow);
      appendAiFollowUpMessage_("assistant", answer);
      const fb =
        j.fallback || j.tier === "free-lite"
          ? " · fallback free lite"
          : j.model
            ? " · " + j.model + (j.thinkingLevel ? "@" + j.thinkingLevel : "")
            : "";
      toast("Answered" + fb);
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      if (thinkingRow && thinkingRow.parentNode) thinkingRow.parentNode.removeChild(thinkingRow);
      appendAiFollowUpMessage_("assistant", "", { error: msg });
      toast("追問失敗：" + msg);
    } finally {
      setAiReportBusy_(false);
      if (input) input.disabled = false;
      if (askBtn) askBtn.disabled = false;
      if (input) input.focus();
    }
  }

  function splitAiMdTableRow_(line) {
    let s = String(line || "").trim();
    if (s.startsWith("|")) s = s.slice(1);
    if (s.endsWith("|")) s = s.slice(0, -1);
    return s.split("|").map((c) => c.trim());
  }

  function formatAiMdInline_(text, opts) {
    const o = opts || {};
    let s = escapeHtml(text);
    s = s.replace(/`([^`\n]+)`/g, "<code>$1</code>");
    s = s.replace(/\*\*([^*\n]+)\*\*/g, "<strong>$1</strong>");
    s = s.replace(/__([^_\n]+)__/g, "<strong>$1</strong>");
    s = s.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, "$1<em>$2</em>");
    if (o.linkDates) s = linkifyAiReportDates_(s);
    return s;
  }

  function resolveAiJumpYmd_(year, month, day) {
    const mo = String(month).padStart(2, "0");
    const da = String(day).padStart(2, "0");
    const yFrom = String(document.getElementById("reportFromStr")?.value || "").slice(0, 4);
    const yTo = String(document.getElementById("reportToStr")?.value || "").slice(0, 4);
    let y = String(year || yFrom || yTo || new Date().getFullYear());
    if (!year && yFrom && yTo && yFrom !== yTo) {
      const candFrom = yFrom + "-" + mo + "-" + da;
      const candTo = yTo + "-" + mo + "-" + da;
      const from = String(document.getElementById("reportFromStr")?.value || "");
      const to = String(document.getElementById("reportToStr")?.value || "");
      if (candFrom >= from && candFrom <= to) y = yFrom;
      else if (candTo >= from && candTo <= to) y = yTo;
    }
    return y + "-" + mo + "-" + da;
  }

  function aiDateJumpBtn_(ymd, label) {
    return (
      '<span role="button" tabindex="0" class="ai-date-jump" data-jump-ymd="' +
      escapeHtml(ymd) +
      '">' +
      escapeHtml(label) +
      "</span>"
    );
  }

  function linkifyAiReportDates_(escapedText) {
    let s = String(escapedText || "");
    // 日期＋可選時間整段一齊（唔好淨包日期令時間跌落下一行）
    s = s.replace(
      /\b(\d{4})-(\d{2})-(\d{2})(?:([ T])(\d{2}:\d{2}(?::\d{2})?))?\b/g,
      (m, y, mo, d) => {
        return aiDateJumpBtn_(y + "-" + mo + "-" + d, m);
      },
    );
    s = s.replace(/(\d{1,2})月(\d{1,2})日(?:\s*(\d{1,2}:\d{2}(?::\d{2})?))?/g, (m, mo, d) => {
      return aiDateJumpBtn_(resolveAiJumpYmd_("", mo, d), m);
    });
    // MM-DD（避免吃到 YYYY-MM-DD 尾段：前面唔係 digit 或 -）
    s = s.replace(/(^|[^0-9-])(\d{2})-(\d{2})(?![0-9:])/g, (m, pre, mo, d) => {
      return pre + aiDateJumpBtn_(resolveAiJumpYmd_("", mo, d), mo + "-" + d);
    });
    return s;
  }

  /** Lightweight Markdown → safe HTML（粗體、標題、列表、GFM table） */
  function renderAiMarkdownToHtml_(md, opts) {
    const inlineOpts = { linkDates: !!(opts && opts.linkDates) };
    const raw = String(md || "").replace(/\r\n/g, "\n");
    const lines = raw.split("\n");
    const html = [];
    let i = 0;
    let inUl = false;
    let inOl = false;

    function closeLists() {
      if (inUl) {
        html.push("</ul>");
        inUl = false;
      }
      if (inOl) {
        html.push("</ol>");
        inOl = false;
      }
    }

    function isTableSep_(line) {
      return /^\s*\|?[\s|:/-]+\|[\s|:|/-]*$/.test(String(line || ""));
    }

    while (i < lines.length) {
      const line = lines[i];
      if (/^\s*\|/.test(line) && i + 1 < lines.length && isTableSep_(lines[i + 1])) {
        closeLists();
        const rows = [];
        while (i < lines.length && /^\s*\|/.test(lines[i])) {
          rows.push(lines[i]);
          i++;
        }
        html.push('<table class="ai-md-table"><thead><tr>');
        splitAiMdTableRow_(rows[0]).forEach((h) => {
          html.push("<th>" + formatAiMdInline_(h, inlineOpts) + "</th>");
        });
        html.push("</tr></thead><tbody>");
        for (let r = 1; r < rows.length; r++) {
          if (isTableSep_(rows[r])) continue;
          html.push("<tr>");
          splitAiMdTableRow_(rows[r]).forEach((c) => {
            html.push("<td>" + formatAiMdInline_(c, inlineOpts) + "</td>");
          });
          html.push("</tr>");
        }
        html.push("</tbody></table>");
        continue;
      }

      const hm = line.match(/^(#{1,3})\s+(.+)$/);
      if (hm) {
        closeLists();
        const level = hm[1].length;
        html.push("<h" + level + ">" + formatAiMdInline_(hm[2], inlineOpts) + "</h" + level + ">");
        i++;
        continue;
      }

      if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
        closeLists();
        html.push("<hr>");
        i++;
        continue;
      }

      const ulm = line.match(/^\s*[-*+]\s+(.+)$/);
      if (ulm) {
        if (inOl) {
          html.push("</ol>");
          inOl = false;
        }
        if (!inUl) {
          html.push("<ul>");
          inUl = true;
        }
        html.push("<li>" + formatAiMdInline_(ulm[1], inlineOpts) + "</li>");
        i++;
        continue;
      }

      const olm = line.match(/^\s*\d+\.\s+(.+)$/);
      if (olm) {
        if (inUl) {
          html.push("</ul>");
          inUl = false;
        }
        if (!inOl) {
          html.push("<ol>");
          inOl = true;
        }
        html.push("<li>" + formatAiMdInline_(olm[1], inlineOpts) + "</li>");
        i++;
        continue;
      }

      if (/^\s*$/.test(line)) {
        closeLists();
        i++;
        continue;
      }

      closeLists();
      const paras = [];
      while (
        i < lines.length &&
        !/^\s*$/.test(lines[i]) &&
        !/^\s*\|/.test(lines[i]) &&
        !/^#{1,3}\s/.test(lines[i]) &&
        !/^\s*[-*+]\s/.test(lines[i]) &&
        !/^\s*\d+\.\s/.test(lines[i]) &&
        !/^(-{3,}|\*{3,}|_{3,})\s*$/.test(lines[i])
      ) {
        paras.push(lines[i]);
        i++;
      }
      html.push("<p>" + formatAiMdInline_(paras.join(" "), inlineOpts) + "</p>");
    }
    closeLists();
    return html.join("\n") || '<p class="muted">（空白報告）</p>';
  }

  function bindAiManualDateJumps_(el) {
    if (!el || el.dataset.dateJumpBound === "1") return;
    el.dataset.dateJumpBound = "1";
    el.addEventListener("click", (ev) => {
      const btn = ev.target && ev.target.closest ? ev.target.closest("[data-jump-ymd]") : null;
      if (!btn || !el.contains(btn)) return;
      ev.preventDefault();
      jumpReportRawToYmd_(btn.getAttribute("data-jump-ymd") || "");
    });
    el.addEventListener("keydown", (ev) => {
      if (ev.key !== "Enter" && ev.key !== " ") return;
      const btn = ev.target && ev.target.closest ? ev.target.closest("[data-jump-ymd]") : null;
      if (!btn || !el.contains(btn)) return;
      ev.preventDefault();
      jumpReportRawToYmd_(btn.getAttribute("data-jump-ymd") || "");
    });
  }

  function setAiReportBodyHtml_(el, markdown, opts) {
    if (!el) return;
    const o = opts || {};
    if (o.loading) {
      el.classList.add("loading");
      el.innerHTML =
        '<p class="ai-report-loading"><span class="ai-report-loading-spin" aria-hidden="true"></span>Generating…</p>';
      startAiReportGenTimer_(el);
      return;
    }
    stopAiReportGenTimer_();
    el.classList.remove("loading");
    if (o.error) {
      el.innerHTML = '<p class="ai-report-loading">' + escapeHtml(String(o.error)) + "</p>";
      return;
    }
    el.innerHTML = renderAiMarkdownToHtml_(markdown, { linkDates: !!o.linkDates });
    if (o.linkDates) bindAiManualDateJumps_(el);
  }

  /** 先上屏純 Markdown，再喺 idle 加日期連結（避免一次卡死主線程） */
  async function paintAiReportMarkdownSmooth_(el, markdown) {
    if (!el) return;
    stopAiReportGenTimer_();
    el.classList.remove("loading");
    await yieldToMain_();
    el.innerHTML = renderAiMarkdownToHtml_(markdown, { linkDates: false });
    await yieldToMain_();
    const enhance = () => {
      if (!el.isConnected) return;
      el.innerHTML = renderAiMarkdownToHtml_(markdown, { linkDates: true });
      bindAiManualDateJumps_(el);
    };
    if (typeof requestIdleCallback === "function") {
      requestIdleCallback(enhance, { timeout: 600 });
    } else {
      setTimeout(enhance, 50);
    }
  }

  function isoWeekKeyFromDateLocal_(d) {
    const date = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    const dayNr = (date.getDay() + 6) % 7;
    date.setDate(date.getDate() - dayNr + 3);
    const weekYear = date.getFullYear();
    const week1 = new Date(weekYear, 0, 4);
    const day1 = (week1.getDay() + 6) % 7;
    week1.setDate(week1.getDate() - day1);
    const weekNo = 1 + Math.round((date.getTime() - week1.getTime()) / 604800000);
    return weekYear + "-W" + String(weekNo).padStart(2, "0");
  }

  /** Display week as MM-DD → MM-DD (never W29). Internal key stays ISO week. */
  function formatPeriodKeyForDisplay_(periodType, periodKey) {
    const type = String(periodType || "").toLowerCase();
    const key = String(periodKey || "").trim();
    if (type === "week") {
      const m = key.match(/^(\d{4})-W(\d{2})$/i);
      if (!m) return key;
      const y = parseInt(m[1], 10);
      const w = parseInt(m[2], 10);
      const week1 = new Date(y, 0, 4);
      const day1 = (week1.getDay() + 6) % 7;
      week1.setDate(week1.getDate() - day1);
      const mon = new Date(week1.getFullYear(), week1.getMonth(), week1.getDate() + (w - 1) * 7);
      const sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
      const pad = (n) => String(n).padStart(2, "0");
      const fmt = (d) => pad(d.getMonth() + 1) + "-" + pad(d.getDate());
      return fmt(mon) + " \u2192 " + fmt(sun);
    }
    return key;
  }

  function periodKeyFromReportDates_() {
    const from = String(document.getElementById("reportFromStr")?.value || "").trim();
    const to = String(document.getElementById("reportToStr")?.value || "").trim();
    if (!from || !to) return null;
    const fm = from.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const tm = to.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!fm || !tm) return { periodType: "custom", periodKey: from + "/" + to };
    // exact calendar month
    if (fm[1] === tm[1] && fm[2] === tm[2] && fm[3] === "01") {
      const last = new Date(parseInt(fm[1], 10), parseInt(fm[2], 10), 0).getDate();
      if (parseInt(tm[3], 10) === last) {
        return { periodType: "month", periodKey: fm[1] + "-" + fm[2] };
      }
    }
    // exact calendar year
    if (from === fm[1] + "-01-01" && to === fm[1] + "-12-31" && fm[1] === tm[1]) {
      return { periodType: "year", periodKey: fm[1] };
    }
    // quarters
    const quarters = [
      ["01-01", "03-31", "Q1"],
      ["04-01", "06-30", "Q2"],
      ["07-01", "09-30", "Q3"],
      ["10-01", "12-31", "Q4"],
    ];
    for (let i = 0; i < quarters.length; i++) {
      const q = quarters[i];
      if (from === fm[1] + "-" + q[0] && to === fm[1] + "-" + q[1] && fm[1] === tm[1]) {
        return { periodType: "quarter", periodKey: fm[1] + "-" + q[2] };
      }
    }
    // ISO week Mon–Sun
    const fromD = new Date(parseInt(fm[1], 10), parseInt(fm[2], 10) - 1, parseInt(fm[3], 10));
    const toD = new Date(parseInt(tm[1], 10), parseInt(tm[2], 10) - 1, parseInt(tm[3], 10));
    const diffDays = Math.round((toD.getTime() - fromD.getTime()) / 86400000);
    if (diffDays === 6 && fromD.getDay() === 1 && toD.getDay() === 0) {
      return { periodType: "week", periodKey: isoWeekKeyFromDateLocal_(fromD) };
    }
    return { periodType: "custom", periodKey: from + "/" + to };
  }

  async function postAiAction_(payload) {
    const url = getRemotePostUrl();
    if (!url || !canRemoteSync()) throw new Error("not_signed_in");
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(remoteAuthBody(payload)),
      mode: "cors",
      cache: "no-store",
    });
    const j = await r.json().catch(() => ({}));
    if (!j || j.ok === false) {
      const err = (j && j.error) || "request_failed";
      if (handleRemoteUnauthorized_(err)) throw new Error(err);
      // 舊部署常見：唔識 action → missing_state / unknown_action
      if (err === "missing_state" || String(err).indexOf("unknown_action:") === 0) {
        throw new Error(
          "呢個 /exec 仲係舊部署（server: " +
            err +
            "）。請喺 Apps Script「部署→管理部署」編輯而家呢個 Web app，版本選「新版本」再部署。URL: " +
            url.slice(0, 60) +
            "…"
        );
      }
      if (
        err === "missing_TimeStatAiReports_gs" ||
        err === "handleGetAiSettings_ is not defined" ||
        /handle\w+Ai\w+ is not defined/.test(String(err))
      ) {
        throw new Error(
          "Apps Script 欠 AI 報告檔（server: " +
            err +
            "）。同一專案要有三個檔：TimeStatSync.gs、TimeStatAiReports.gs、TimeStatAiPeriodKpis.gs，然後「管理部署→新版本」。"
        );
      }
      throw new Error(err);
    }
    return j;
  }

  async function generateManualAiReport_(wantEmail) {
    if (!canRemoteSync()) {
      toast("請先 Google 登入先可以 Generate AI Report。");
      return;
    }
    const fromEl = document.getElementById("reportFromStr");
    const toEl = document.getElementById("reportToStr");
    const fromSnap = String(fromEl && fromEl.value ? fromEl.value : "").trim();
    const toSnap = String(toEl && toEl.value ? toEl.value : "").trim();
    const pk = periodKeyFromReportDates_();
    if (!pk) {
      toast("請先揀 Report 日期範圍。");
      return;
    }
    const btn = document.getElementById("btnAiReportGenerate");
    const box = document.getElementById("aiReportManualBox");
    const body = document.getElementById("aiReportManualBody");
    if (btn) btn.disabled = true;
    setAiReportBusy_(true);
    if (box) box.classList.remove("hidden");
    resetAiFollowUpUi_();
    setAiReportBodyHtml_(body, "", { loading: true });
    if (box && typeof box.scrollIntoView === "function") {
      try {
        box.scrollIntoView({ behavior: "smooth", block: "nearest" });
      } catch (_) {
        box.scrollIntoView(true);
      }
    }
    try {
      await yieldToMain_();
      const j = await postAiAction_({
        action: "generateAiReport",
        periodType: pk.periodType,
        periodKey: pk.periodKey,
        email: !!wantEmail,
      });
      _lastManualAiReport_ = {
        markdown: j.markdown || "",
        subject: j.subject || "",
        periodType: j.periodType,
        periodKey: j.periodKey,
        history: [],
      };
      // Generate 期間若有 sync 改咗 date pickers，還原返用戶揀嘅範圍（只喺真係被改過先重繪報表）
      let datesChanged = false;
      if (fromEl && toEl && fromSnap && toSnap) {
        if (fromEl.value !== fromSnap || toEl.value !== toSnap) {
          const pr = document.getElementById("reportPeriodPreset");
          if (pr) pr.value = "custom";
          reportPresetSuppress = true;
          try {
            fromEl.value = fromSnap;
            toEl.value = toSnap;
          } finally {
            queueMicrotask(() => {
              reportPresetSuppress = false;
            });
          }
          datesChanged = true;
        }
      }
      if (box) box.classList.remove("hidden");
      await paintAiReportMarkdownSmooth_(body, _lastManualAiReport_.markdown);
      showAiFollowUpUi_();
      updateAiReportPeriodBadge_();
      if (datesChanged) {
        await yieldToMain_();
        renderReport();
      }
      const modelBit = j.model ? String(j.model) : "";
      const thinkBit = j.thinkingLevel ? "@" + j.thinkingLevel : "";
      const fb =
        j.fallback || j.tier === "free-lite"
          ? " · fallback free lite（Pro 無 quota／唔支援）"
          : modelBit
            ? " · " + modelBit + thinkBit
            : "";
      toast(
        (wantEmail ? "AI Report 已寄出（唔入歷史）" : "AI Report 已生成（唔入歷史）") + fb,
      );
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      setAiReportBodyHtml_(body, "", { error: "AI Report 失敗：" + msg });
      resetAiFollowUpUi_();
      toast("AI Report 失敗：" + msg);
    } finally {
      stopAiReportGenTimer_();
      setAiReportBusy_(false);
      if (btn) btn.disabled = false;
    }
  }

  async function loadAiReportsList_() {
    const listEl = document.getElementById("aiReportsList");
    const detail = document.getElementById("aiReportDetail");
    if (detail) detail.classList.add("hidden");
    if (listEl) listEl.classList.remove("hidden");
    if (!listEl) return;
    if (!canRemoteSync()) {
      listEl.innerHTML = '<p class="muted">請先 Google 登入。</p>';
      return;
    }
    listEl.innerHTML = '<p class="muted">Loading…</p>';
    try {
      const j = await postAiAction_({ action: "listAiReports", limit: 80 });
      const reports = Array.isArray(j.reports) ? j.reports : [];
      if (!reports.length) {
        listEl.innerHTML = '<p class="muted">No automatic AI reports yet.</p>';
        return;
      }
      listEl.innerHTML = "";
      reports.forEach((rep) => {
        const row = document.createElement("button");
        row.type = "button";
        row.className = "ai-reports-item";
        row.textContent =
          (rep.periodType || "") +
          " " +
          formatPeriodKeyForDisplay_(rep.periodType, rep.periodKey) +
          " · " +
          String(rep.generatedAt || "").slice(0, 19).replace("T", " ");
        row.addEventListener("click", () => void openAiReportDetail_(rep.id));
        listEl.appendChild(row);
      });
    } catch (e) {
      listEl.innerHTML =
        '<p class="muted">載入失敗：' + escapeHtml(e && e.message ? e.message : String(e)) + "</p>";
    }
  }

  async function openAiReportDetail_(id) {
    const listEl = document.getElementById("aiReportsList");
    const detail = document.getElementById("aiReportDetail");
    const title = document.getElementById("aiReportDetailTitle");
    const body = document.getElementById("aiReportDetailBody");
    if (!detail || !body) return;
    try {
      const j = await postAiAction_({ action: "getAiReport", id: id });
      const rep = j.report || {};
      if (listEl) listEl.classList.add("hidden");
      detail.classList.remove("hidden");
      if (title) {
        const label = formatPeriodKeyForDisplay_(rep.periodType, rep.periodKey);
        let subj = String(rep.subject || "");
        if (rep.periodKey && label && subj.indexOf(rep.periodKey) >= 0) {
          subj = subj.split(String(rep.periodKey)).join(label);
        }
        title.textContent = subj || ((rep.periodType || "") + " " + label) || String(id);
      }
      setAiReportBodyHtml_(body, rep.markdown || "");
    } catch (e) {
      toast("開啟報告失敗：" + (e && e.message ? e.message : String(e)));
    }
  }

  async function loadAiSettingsForm_() {
    const sys = document.getElementById("aiSettingSystem");
    const extra = document.getElementById("aiSettingExtra");
    const temp = document.getElementById("aiSettingTemp");
    if (!sys) return;
    if (!canRemoteSync()) {
      sys.value = "";
      if (extra) extra.value = "";
      return;
    }
    try {
      const j = await postAiAction_({ action: "getAiSettings" });
      const s = j.settings || {};
      sys.value = s.systemInstruction || "";
      if (extra) extra.value = s.extraInstructions || "";
      if (temp) temp.value = s.temperature != null ? String(s.temperature) : "0.3";
      window.__AI_SETTINGS_DEFAULTS__ = j.defaults || null;
      window.__AI_SECTION_LABELS__ = j.sectionLabels || {};
      renderAiPeriodPanels_(s.periodConfig || (j.defaults && j.defaults.periodConfig) || {});
      const neg = document.getElementById("aiEmotionNeg");
      const pos = document.getElementById("aiEmotionPos");
      const ek = s.emotionKeywords || {};
      if (neg) neg.value = (ek.negative || []).join(", ");
      if (pos) pos.value = (ek.positive || []).join(", ");
      bindAiPeriodTabs_();
    } catch (e) {
      toast("載入 AI Settings 失敗：" + (e && e.message ? e.message : String(e)));
    }
  }

  function parseKeywordList_(s) {
    return String(s || "")
      .split(/[,，\n]+/)
      .map((x) => x.trim())
      .filter(Boolean);
  }

  function renderAiPeriodPanels_(periodConfig) {
    const root = document.getElementById("aiPeriodPanels");
    if (!root) return;
    const labels = window.__AI_SECTION_LABELS__ || {};
    const periods = ["week", "month", "quarter", "year"];
    root.innerHTML = "";
    periods.forEach((pt, idx) => {
      const pc = (periodConfig && periodConfig[pt]) || { sections: {}, notes: "" };
      const sec = pc.sections || {};
      const lab = labels[pt] || {};
      const panel = document.createElement("div");
      panel.className = "ai-period-panel" + (idx === 0 ? "" : " hidden");
      panel.dataset.period = pt;
      const keys = Object.keys(lab).length ? Object.keys(lab) : Object.keys(sec);
      keys.forEach((k) => {
        const row = document.createElement("label");
        row.className = "ai-period-check";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.dataset.period = pt;
        cb.dataset.section = k;
        cb.checked = sec[k] !== false;
        const span = document.createElement("span");
        span.textContent = lab[k] || k;
        row.appendChild(cb);
        row.appendChild(span);
        panel.appendChild(row);
      });
      const noteLab = document.createElement("label");
      noteLab.className = "ai-period-note-label";
      noteLab.textContent = "Notes (" + pt + ")";
      noteLab.htmlFor = "aiPeriodNotes_" + pt;
      const ta = document.createElement("textarea");
      ta.id = "aiPeriodNotes_" + pt;
      ta.className = "ai-setting-ta";
      ta.rows = 3;
      ta.value = pc.notes || "";
      panel.appendChild(noteLab);
      panel.appendChild(ta);
      root.appendChild(panel);
    });
  }

  function readAiPeriodConfigFromForm_() {
    const out = { week: {}, month: {}, quarter: {}, year: {} };
    ["week", "month", "quarter", "year"].forEach((pt) => {
      const sections = {};
      document.querySelectorAll('input[type="checkbox"][data-period="' + pt + '"]').forEach((cb) => {
        sections[cb.dataset.section] = !!cb.checked;
      });
      const ta = document.getElementById("aiPeriodNotes_" + pt);
      out[pt] = { sections: sections, notes: ta ? ta.value : "" };
    });
    return out;
  }

  function bindAiPeriodTabs_() {
    const tabs = document.querySelectorAll(".ai-period-tab");
    tabs.forEach((btn) => {
      if (btn.dataset.bound) return;
      btn.dataset.bound = "1";
      btn.addEventListener("click", () => {
        const pt = btn.dataset.period;
        tabs.forEach((b) => b.classList.toggle("active", b === btn));
        document.querySelectorAll(".ai-period-panel").forEach((p) => {
          p.classList.toggle("hidden", p.dataset.period !== pt);
        });
      });
    });
  }

  function updateAiReportPeriodBadge_() {
    const badge = document.getElementById("aiReportPeriodBadge");
    if (!badge) return;
    const pk = periodKeyFromReportDates_();
    if (!pk) {
      badge.textContent = "";
      return;
    }
    badge.textContent = pk.periodType + " · " + formatPeriodKeyForDisplay_(pk.periodType, pk.periodKey);
  }

  async function saveAiSettingsForm_() {
    if (!canRemoteSync()) {
      toast("請先 Google 登入。");
      return;
    }
    const sys = document.getElementById("aiSettingSystem");
    const extra = document.getElementById("aiSettingExtra");
    const temp = document.getElementById("aiSettingTemp");
    const neg = document.getElementById("aiEmotionNeg");
    const pos = document.getElementById("aiEmotionPos");
    try {
      await postAiAction_({
        action: "saveAiSettings",
        settings: {
          systemInstruction: sys ? sys.value : "",
          extraInstructions: extra ? extra.value : "",
          temperature: temp ? Number(temp.value) : 0.3,
          periodConfig: readAiPeriodConfigFromForm_(),
          emotionKeywords: {
            negative: parseKeywordList_(neg ? neg.value : ""),
            positive: parseKeywordList_(pos ? pos.value : ""),
          },
        },
      });
      toast("AI Settings 已儲存（自動／人手報告都會用）。");
    } catch (e) {
      toast("儲存失敗：" + (e && e.message ? e.message : String(e)));
    }
  }

  function resetAiSettingsFormToDefaults_() {
    const d = window.__AI_SETTINGS_DEFAULTS__;
    if (!d) {
      void loadAiSettingsForm_().then(() => {
        const d2 = window.__AI_SETTINGS_DEFAULTS__;
        if (!d2) {
          toast("未有預設；請先 Refresh／登入。");
          return;
        }
        applyAiDefaultsToForm_(d2);
      });
      return;
    }
    applyAiDefaultsToForm_(d);
  }

  function applyAiDefaultsToForm_(d) {
    const sys = document.getElementById("aiSettingSystem");
    const extra = document.getElementById("aiSettingExtra");
    const temp = document.getElementById("aiSettingTemp");
    if (sys) sys.value = d.systemInstruction || "";
    if (extra) extra.value = d.extraInstructions || "";
    if (temp) temp.value = d.temperature != null ? String(d.temperature) : "0.3";
    renderAiPeriodPanels_(d.periodConfig || {});
    bindAiPeriodTabs_();
    const neg = document.getElementById("aiEmotionNeg");
    const pos = document.getElementById("aiEmotionPos");
    const ek = d.emotionKeywords || {};
    if (neg) neg.value = (ek.negative || []).join(", ");
    if (pos) pos.value = (ek.positive || []).join(", ");
    toast("已填入預設（未儲存；記得 Save）。");
  }

  (function bindAiReportUi_() {
    ["reportFromStr", "reportToStr"].forEach((id) => {
      const el = document.getElementById(id);
      if (!el || el.dataset.aiBadgeBound) return;
      el.dataset.aiBadgeBound = "1";
      el.addEventListener("change", () => updateAiReportPeriodBadge_());
      el.addEventListener("input", () => updateAiReportPeriodBadge_());
    });
    updateAiReportPeriodBadge_();
    const gen = document.getElementById("btnAiReportGenerate");
    if (gen && !gen.dataset.bound) {
      gen.dataset.bound = "1";
      gen.addEventListener("click", () => void generateManualAiReport_(false));
    }
    const em = document.getElementById("btnAiReportEmail");
    if (em && !em.dataset.bound) {
      em.dataset.bound = "1";
      em.addEventListener("click", () => void generateManualAiReport_(true));
    }
    const close = document.getElementById("btnAiReportClose");
    if (close && !close.dataset.bound) {
      close.dataset.bound = "1";
      close.addEventListener("click", () => {
        const box = document.getElementById("aiReportManualBox");
        if (box) box.classList.add("hidden");
        resetAiFollowUpUi_();
      });
    }
    const ask = document.getElementById("btnAiReportFollowUpAsk");
    if (ask && !ask.dataset.bound) {
      ask.dataset.bound = "1";
      ask.addEventListener("click", () => void askAiReportFollowUp_());
    }
    const followInput = document.getElementById("aiReportFollowUpInput");
    if (followInput && !followInput.dataset.bound) {
      followInput.dataset.bound = "1";
      followInput.addEventListener("keydown", (ev) => {
        if (ev.key === "Enter" && (ev.metaKey || ev.ctrlKey)) {
          ev.preventDefault();
          void askAiReportFollowUp_();
        }
      });
    }
    const refresh = document.getElementById("btnAiReportsRefresh");
    if (refresh && !refresh.dataset.bound) {
      refresh.dataset.bound = "1";
      refresh.addEventListener("click", () => {
        void loadAiReportsList_();
        void loadAiSettingsForm_();
      });
    }
    const back = document.getElementById("btnAiReportBack");
    if (back && !back.dataset.bound) {
      back.dataset.bound = "1";
      back.addEventListener("click", () => {
        const detail = document.getElementById("aiReportDetail");
        const listEl = document.getElementById("aiReportsList");
        if (detail) detail.classList.add("hidden");
        if (listEl) listEl.classList.remove("hidden");
      });
    }
    const saveBtn = document.getElementById("btnAiSettingsSave");
    if (saveBtn && !saveBtn.dataset.bound) {
      saveBtn.dataset.bound = "1";
      saveBtn.addEventListener("click", () => void saveAiSettingsForm_());
    }
    const resetBtn = document.getElementById("btnAiSettingsReset");
    if (resetBtn && !resetBtn.dataset.bound) {
      resetBtn.dataset.bound = "1";
      resetBtn.addEventListener("click", () => resetAiSettingsFormToDefaults_());
    }
  })();

  (function bindTimelineDatePicker() {
    const pick = document.getElementById("timelinePickDate");
    if (pick) {
      pick.addEventListener("change", () => {
        const v = pick.value.trim();
        const norm = v ? parseYMDStrict(v) : null;
        timelineCenterYmd = norm || todayYmd();
        if (v && !norm) pick.value = timelineCenterYmd;
        renderTimeline();
      });
    }
  })();

  /** Report 預設日期：曆月頭～曆月尾（`syncReportDatesFromEvents` 用）。 */
  function monthBoundsYMD(d) {
    const dt = d instanceof Date ? d : new Date();
    const y = dt.getFullYear();
    const m = dt.getMonth();
    const pad = (n) => String(n).padStart(2, "0");
    const lastDay = new Date(y, m + 1, 0).getDate();
    return {
      from: `${y}-${pad(m + 1)}-01`,
      to: `${y}-${pad(m + 1)}-${pad(lastDay)}`,
    };
  }

  function syncReportDatesFromEvents() {
    const fromEl = document.getElementById("reportFromStr");
    const toEl = document.getElementById("reportToStr");
    if (!fromEl || !toEl) return;
    // 已有範圍就唔好覆蓋（否則遠端 sync 會把用戶揀嘅 Apr–Jun 改返做今個月）
    if (String(fromEl.value || "").trim() && String(toEl.value || "").trim()) {
      updateAiReportPeriodBadge_();
      return;
    }
    const b = monthBoundsYMD(new Date());
    fromEl.value = b.from;
    toEl.value = b.to;
    updateAiReportPeriodBadge_();
  }

  /** When leaving multi-window presets for Custom: keep current From～To（通常係 focus period）. */
  function setReportRangeToCurrentCalendarMonth() {
    // no-op 保留：舊呼叫位唔好再強制跳去今個月
  }

  const MANUAL_DATE_CHIP_DAYS = 3;

  /** Manual date summary text: YYYY-MM-DD */
  function manualDateSummaryText(ymd) {
    const t = String(ymd || "").trim();
    if (!t) return "—";
    return t;
  }

  function updateManualDateSummary() {
    const hidden = document.getElementById("manualDateSelected");
    const sum = document.getElementById("manualDateSummary");
    if (!sum) return;
    const v = hidden && hidden.value;
    sum.textContent = manualDateSummaryText(v);
  }

  function updateManualTimeSummary() {
    const sum = document.getElementById("manualTimeSummary");
    const h = document.getElementById("manualHourSel");
    const m = document.getElementById("manualMinuteSel");
    if (!sum || !h || !m) return;
    const hh = h.value || "00";
    const mm = m.value || "00";
    sum.textContent = `${hh}:${mm}`;
    refreshManualAutoSuggestions();
  }

  /** 後補日期：最近 3 日（收埋喺 details，撳先見） */
  function renderManualDateChips() {
    const wrap = document.getElementById("manualDateWrap");
    const hidden = document.getElementById("manualDateSelected");
    const det = document.getElementById("manualDateDetails");
    if (!wrap || !hidden) return;
    wrap.innerHTML = "";
    let firstVal = "";
    for (let i = 0; i < MANUAL_DATE_CHIP_DAYS; i++) {
      const d = new Date();
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() - i);
      const value = ymdFromLocalDate(d);
      if (i === 0) firstVal = value;
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "picker-chip" + (i === 0 ? " is-active" : "");
      btn.dataset.dateValue = value;
      btn.textContent = value;
      btn.addEventListener("click", () => {
        wrap.querySelectorAll(".picker-chip").forEach((b) => b.classList.remove("is-active"));
        btn.classList.add("is-active");
        hidden.value = value;
        updateManualDateSummary();
        if (det) det.open = false;
      });
      wrap.appendChild(btn);
    }
    hidden.value = firstVal;
    updateManualDateSummary();
  }

  /** 後補時間：頁內碌轉盤（scroll-snap、三循環，00 再上＝23／59） */
  const WHEEL_ITEM_H = 44;
  const WHEEL_VIEW_H = 220;
  const WHEEL_REPEAT = 3;

  function wheelPadPx() {
    return (WHEEL_VIEW_H - WHEEL_ITEM_H) / 2;
  }

  function buildCycleWheelList(listEl, modulus) {
    listEl.innerHTML = "";
    const pad = wheelPadPx();
    listEl.style.paddingTop = `${pad}px`;
    listEl.style.paddingBottom = `${pad}px`;
    for (let r = 0; r < WHEEL_REPEAT; r++) {
      for (let i = 0; i < modulus; i++) {
        const div = document.createElement("div");
        div.className = "manual-wheel-item";
        const s = String(i).padStart(2, "0");
        div.textContent = s;
        div.dataset.value = s;
        listEl.appendChild(div);
      }
    }
  }

  function readWheelValue(viewport, modulus) {
    const itemH = WHEEL_ITEM_H;
    let idx = Math.round(viewport.scrollTop / itemH);
    return ((idx % modulus) + modulus) % modulus;
  }

  function wrapWheelIfNeeded(viewport, modulus) {
    const itemH = WHEEL_ITEM_H;
    let idx = Math.round(viewport.scrollTop / itemH);
    if (idx < modulus) {
      viewport.scrollTop = (idx + modulus) * itemH;
    } else if (idx >= modulus * 2) {
      viewport.scrollTop = (idx - modulus) * itemH;
    }
  }

  function finalizeWheelScroll(viewport, hiddenInput, modulus, allowWrap) {
    if (allowWrap) wrapWheelIfNeeded(viewport, modulus);
    const v = readWheelValue(viewport, modulus);
    hiddenInput.value = String(v).padStart(2, "0");
    updateManualTimeSummary();
  }

  function attachWheel(viewport, hiddenInput, modulus) {
    let debounceT = null;
    let touching = false;
    let wrapping = false;

    function settle(forceWrap) {
      if (wrapping) return;
      clearTimeout(debounceT);
      debounceT = setTimeout(() => {
        if (touching && !forceWrap) return;
        wrapping = true;
        finalizeWheelScroll(viewport, hiddenInput, modulus, true);
        wrapping = false;
      }, forceWrap ? 16 : 140);
    }

    viewport.addEventListener(
      "scroll",
      () => {
        if (wrapping) return;
        // 轉動中只更新顯示，唔跳 wrap，避免打斷慣性
        const v = readWheelValue(viewport, modulus);
        hiddenInput.value = String(v).padStart(2, "0");
        updateManualTimeSummary();
        settle(false);
      },
      { passive: true },
    );
    viewport.addEventListener(
      "touchstart",
      () => {
        touching = true;
        clearTimeout(debounceT);
      },
      { passive: true },
    );
    viewport.addEventListener(
      "touchend",
      () => {
        touching = false;
        settle(true);
      },
      { passive: true },
    );
    viewport.addEventListener(
      "touchcancel",
      () => {
        touching = false;
        settle(true);
      },
      { passive: true },
    );
    viewport.addEventListener("scrollend", () => {
      touching = false;
      settle(true);
    });
  }

  function setWheelToValue(viewport, hiddenInput, modulus, valNum) {
    const v = Math.max(0, Math.min(modulus - 1, Math.floor(valNum)));
    const idx = modulus + v;
    viewport.scrollTop = idx * WHEEL_ITEM_H;
    hiddenInput.value = String(v).padStart(2, "0");
    requestAnimationFrame(() => {
      requestAnimationFrame(() => finalizeWheelScroll(viewport, hiddenInput, modulus, true));
    });
  }

  let manualWheelsReady = false;
  function initManualWheelsOnce() {
    if (manualWheelsReady) return;
    const hList = document.getElementById("manualHourWheelList");
    const mList = document.getElementById("manualMinuteWheelList");
    const hVp = document.getElementById("manualHourWheelViewport");
    const mVp = document.getElementById("manualMinuteWheelViewport");
    const hHid = document.getElementById("manualHourSel");
    const mHid = document.getElementById("manualMinuteSel");
    if (!hList || !mList || !hVp || !mVp || !hHid || !mHid) return;
    buildCycleWheelList(hList, 24);
    buildCycleWheelList(mList, 60);
    attachWheel(hVp, hHid, 24);
    attachWheel(mVp, mHid, 60);
    manualWheelsReady = true;
  }

  function initManualDateTime() {
    renderManualDateChips();
    initManualWheelsOnce();
    const hVp = document.getElementById("manualHourWheelViewport");
    const mVp = document.getElementById("manualMinuteWheelViewport");
    const hHid = document.getElementById("manualHourSel");
    const mHid = document.getElementById("manualMinuteSel");
    if (!hVp || !mVp || !hHid || !mHid) return;
    const d = new Date();
    setWheelToValue(hVp, hHid, 24, d.getHours());
    setWheelToValue(mVp, mHid, 60, d.getMinutes());
    updateManualTimeSummary();
  }


  async function runRemoteHydrate() {
    if (!canRemoteSync()) return;
    const r = await fetch(getRemotePostUrl(), {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(remoteAuthBody({ action: "load" })),
      cache: "no-store",
    });
    const j = await r.json().catch(() => ({}));
    if (j && j._formImportDebug) {
      try {
        console.info("[TimeStat] form import debug", j._formImportDebug);
      } catch (e) {}
    }
    if (!j || !j.ok) {
      const err = (j && j.error) || "load_failed";
      const api = j && j.authApi ? String(j.authApi) : "";
      if (handleRemoteUnauthorized_(err)) {
        if (err === "unauthorized" && !api) {
          throw new Error(
            "unauthorized — Apps Script 仍係舊版（無 authApi）。請刪舊 Code.gs、只留 TimeStatSync.gs、管理部署→新版本。"
          );
        }
        throw new Error(err + (api ? " [" + api + "]" : ""));
      }
      throw new Error(err + (api ? " [" + api + "]" : ""));
    }
    if (!Object.prototype.hasOwnProperty.call(j, "state")) {
      toast("Google 端未回傳 state（請確認 Apps Script 已部署 doPost + action=load）。");
      return;
    }
    if (j.state === null) {
      if (state.events.length > 0) await pushRemoteStateQuiet();
      return;
    }
    const parsed = normalizeStateFromParsed(j.state);
    if (!parsed) {
      toast("遠端資料格式錯誤，已保留本機版本。");
      return;
    }
    const prevEvCount = state.events.length;
    const nextEvCount = (parsed.out.events || []).length;
    if (nextEvCount === 0 && prevEvCount > 0) {
      toast("遠端 0 筆紀錄，已保留本機 " + prevEvCount + " 筆，並寫上雲端。");
      state.updatedAt = Date.now();
      await pushRemoteStateQuiet();
      return;
    }
    // 兩邊都有資料：永遠做 union merge，避免「雲端較多但缺本機幾筆」時覆蓋掉 log
    if (prevEvCount > 0 && nextEvCount > 0) {
      const remoteOnly = nextEvCount;
      mergeRemoteStateIntoLocal_(j.state);
      state.updatedAt = Date.now();
      if (state.events.length > remoteOnly) {
        toast(
          "已合併本機＋雲端：" +
            state.events.length +
            " 筆（雲端原 " +
            remoteOnly +
            "，本機有多出嘅會寫返上）。"
        );
      }
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      } catch (e) {}
      await pushRemoteStateQuiet();
      refreshActivityDatalist();
      refreshProjectPickers();
      fillMergeSelects();
      renderActivityList();
      renderTimeline();
      syncReportDatesFromEvents();
      renderReport();
      initManualDateTime();
      refreshManualAutoSuggestions();
      updateLastSavedHint();
      updateAuthChrome_();
      refreshSoftCapBanner();
      return;
    }
    state = parsed.out;
    if (state.updatedAt == null) state.updatedAt = Date.now();
    bumpEventsMutationGen();
    dedupeStateEventsByImportKey();
    state.structure = [];
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {}
    // 登入後再 push 一次，確保本機有而雲端缺嘅筆寫返上（排隊、唔並行）
    void pushRemoteStateQuiet();
    refreshActivityDatalist();
    refreshProjectPickers();
    fillMergeSelects();
    renderActivityList();
    renderTimeline();
    syncReportDatesFromEvents();
    renderReport();
    initManualDateTime();
    refreshManualAutoSuggestions();
    updateLastSavedHint();
    updateAuthChrome_();
    refreshSoftCapBanner();
  }

  function showAuthOverlay_(message) {
    const overlay = document.getElementById("authOverlay");
    const msg = document.getElementById("authOverlayMsg");
    if (msg && message) msg.textContent = String(message);
    if (overlay) overlay.classList.remove("hidden");
    document.body.classList.add("auth-locked");
    updateAuthChrome_();
    renderGoogleSignInButton_();
  }

  function hideAuthOverlay_() {
    const overlay = document.getElementById("authOverlay");
    if (overlay) overlay.classList.add("hidden");
    document.body.classList.remove("auth-locked");
    updateAuthChrome_();
  }

  function updateAuthChrome_() {
    const emailEl = document.getElementById("authEmailLabel");
    const btnOut = document.getElementById("btnSignOut");
    const email = getAuthEmail();
    if (emailEl) emailEl.textContent = email || "";
    if (btnOut) btnOut.classList.toggle("hidden", !isSignedIn());
    setRemoteSyncStatus_(_remoteSyncStatus || (isSignedIn() ? "ok" : ""));
  }

  function parseJwtEmail_(credential) {
    try {
      const parts = String(credential || "").split(".");
      if (parts.length < 2) return "";
      const json = atob(parts[1].replace(/-/g, "+").replace(/_/g, "/"));
      const payload = JSON.parse(json);
      return String(payload.email || "").trim();
    } catch (e) {
      return "";
    }
  }

  function renderGoogleSignInButton_() {
    const host = document.getElementById("googleSignInBtn");
    if (!host) return;
    host.innerHTML = "";
    const clientId = getGoogleClientId();
    if (!clientId) {
      host.innerHTML =
        '<p class="muted" style="margin:0;">Missing Google Client ID. Set GOOGLE_CLIENT_ID_DEFAULT in app.js or config.remote.json → googleClientId.</p>';
      return;
    }
    if (typeof google === "undefined" || !google.accounts || !google.accounts.id) {
      host.innerHTML = '<p class="muted" style="margin:0;">Loading Google Sign-In…</p>';
      return;
    }
    google.accounts.id.initialize({
      client_id: clientId,
      callback: onGoogleCredentialResponse_,
      auto_select: false,
      cancel_on_tap_outside: true,
    });
    google.accounts.id.renderButton(host, {
      theme: "outline",
      size: "large",
      width: 280,
      text: "signin_with",
      shape: "rectangular",
    });
  }

  async function onGoogleCredentialResponse_(response) {
    const cred = response && response.credential ? String(response.credential) : "";
    if (!cred) {
      toast("Google sign-in failed.");
      return;
    }
    const email = parseJwtEmail_(cred);
    setAuthSession(cred, email);
    const msg = document.getElementById("authOverlayMsg");
    if (msg) msg.textContent = "Signed in" + (email ? " as " + email : "") + ". Loading…";
    try {
      await runRemoteHydrate();
      hideAuthOverlay_();
      toast("Signed in" + (email ? " · " + email : ""));
    } catch (e) {
      clearAuthSession();
      const err = e && e.message ? e.message : String(e);
      showAuthOverlay_(
        err === "email_not_allowed" || String(err).indexOf("email_not_allowed") === 0
          ? "This Google account is not allowed."
          : String(err).indexOf("unauthorized") === 0
            ? String(err)
            : "Sign-in ok but sync failed: " + err,
      );
    }
  }

  function signOut_() {
    clearAuthSession();
    try {
      if (typeof google !== "undefined" && google.accounts && google.accounts.id) {
        google.accounts.id.disableAutoSelect();
      }
    } catch (e) {}
    showAuthOverlay_("Signed out. Sign in with an allowed Google account to load your data.");
  }

  async function loadRemoteSyncDefaultsFromJson_() {
    if (typeof fetch === "undefined") return;
    try {
      const r = await fetch("config.remote.json", { cache: "no-store" });
      if (!r.ok) return;
      const j = await r.json().catch(() => null);
      if (!j || typeof j !== "object") return;
      const bu = j.execUrl != null ? String(j.execUrl).trim() : "";
      const cid = j.googleClientId != null ? String(j.googleClientId).trim() : "";
      // config.remote.json 可覆寫內建 default（換咗部署 URL 時只改 json 都得）
      if (bu && typeof window !== "undefined") {
        window.__TIME_STAT_REMOTE_BASE__ = bu;
      }
      if (cid && typeof window !== "undefined") {
        window.__TIME_STAT_GOOGLE_CLIENT_ID__ = cid;
      }
    } catch (e) {
      /* 無檔案／離線：無視 */
    }
  }

  function bootAuthAndRemote_() {
    void (async () => {
      try {
        await loadRemoteSyncDefaultsFromJson_();
      } catch (e) {}

      const btnOut = document.getElementById("btnSignOut");
      if (btnOut && !btnOut.dataset.bound) {
        btnOut.dataset.bound = "1";
        btnOut.addEventListener("click", () => signOut_());
      }

      if (!useRemoteSync()) {
        hideAuthOverlay_();
        updateAuthChrome_();
        return;
      }

      if (!getGoogleClientId()) {
        showAuthOverlay_("Configure Google Client ID (app.js GOOGLE_CLIENT_ID_DEFAULT or config.remote.json).");
        return;
      }

      if (isSignedIn()) {
        hideAuthOverlay_();
        try {
          await runRemoteHydrate();
        } catch (e) {
          const err = String((e && e.message) || e || "");
          const bare = err.replace(/\s*\[gis-v1\]\s*$/i, "").trim();
          if (handleRemoteUnauthorized_(bare) || /id_token_expired/i.test(err)) {
            if (!document.body.classList.contains("auth-locked")) {
              clearAuthSession();
              showAuthOverlay_("Session expired. Please sign in again.");
            }
            return;
          }
          showAuthOverlay_("Could not load data: " + err);
        }
        return;
      }

      showAuthOverlay_("Sign in with Google to load your Time Stat data.");
    })();
  }

  // Wait for GIS script; then boot.
  (function waitGisThenBoot() {
    let tries = 0;
    function tick() {
      tries++;
      if ((typeof google !== "undefined" && google.accounts && google.accounts.id) || tries > 40) {
        bootAuthAndRemote_();
        return;
      }
      setTimeout(tick, 100);
    }
    tick();
  })();

  try {
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden" && canRemoteSync()) {
        void pushRemoteStateQuiet();
      }
    });
    window.addEventListener("pagehide", () => {
      if (canRemoteSync()) void pushRemoteStateQuiet();
    });
  } catch (eVis) {}

  refreshActivityDatalist();
  refreshProjectPickers();
  bindPlaceAutoSuggest();
  bindDistractionWatch_();
  refreshQuickAutoSuggestions();
  fillMergeSelects();
  renderActivityList();
  initTimelineDatePicker();
  renderTimeline();
  syncReportDatesFromEvents();
  renderReport();
  initManualDateTime();
  refreshManualAutoSuggestions();
  updateLastSavedHint();
  refreshSoftCapBanner();
  const cancelBtn = document.getElementById("btnMappingCancel");
  if (cancelBtn) cancelBtn.addEventListener("click", clearApprovalPanel);


  (function initReportSearchField() {
    const el = document.getElementById("reportKeywordSearch");
    if (!el || el.dataset.searchUnlock) return;
    el.dataset.searchUnlock = "1";
    const refreshRo = () => {
      if (!String(el.value || "").trim()) {
        el.readOnly = true;
        el.placeholder = "Search";
      }
    };
    el.addEventListener("pointerdown", () => {
      el.readOnly = false;
    });
    el.addEventListener("focus", () => {
      el.readOnly = false;
    });
    el.addEventListener("blur", () => {
      refreshRo();
    });
    refreshRo();
  })();

  (function bindReportFilters() {
    const msBoxes = [
      "reportFilterGroupBox",
      "reportFilterLayerBox",
      "reportFilterCatBox",
      "reportFilterSubBox",
      "reportFilterProjectBox",
      "reportFilterActivityBox",
    ];
    msBoxes.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.addEventListener("change", () => renderReport());
    });
    const pe = document.getElementById("reportPeopleSearch");
    if (pe) {
      pe.addEventListener("input", () => {
        clearTimeout(reportPeopleSearchTimer);
        reportPeopleSearchTimer = setTimeout(() => renderReport(), 280);
      });
    }
    const kwIn = document.getElementById("reportKeywordSearch");
    if (kwIn) {
      kwIn.addEventListener("input", () => {
        clearTimeout(reportKeywordSearchTimer);
        reportKeywordSearchTimer = setTimeout(() => renderReport(), 280);
      });
    }
    const presetEl = document.getElementById("reportPeriodPreset");
    if (presetEl) {
      presetEl.dataset.prevPreset = presetEl.value || "custom";
      presetEl.addEventListener("change", () => {
        const prev = presetEl.dataset.prevPreset || "custom";
        const next = presetEl.value || "custom";
        presetEl.dataset.prevPreset = next;
        if (next === "custom" && reportComparePresetActive(prev)) {
          // 保留而家 From～To（compare focus），唔好跳去今個月
        }
        applyReportPeriodPreset();
        renderReport();
      });
    }
    ["reportFromStr", "reportToStr"].forEach((id) => {
      const el = document.getElementById(id);
      if (!el) return;
      el.addEventListener("change", () => {
        if (reportPresetSuppress) return;
        const pr = document.getElementById("reportPeriodPreset");
        if (pr) pr.value = "custom";
        renderReport();
      });
    });
  })();

  (function bindReportUnitToggle() {
    const wrap = document.getElementById("reportUnitToggle");
    if (!wrap) return;
    wrap.addEventListener("click", (e) => {
      const btn = e.target.closest(".report-unit-btn");
      if (!btn || !wrap.contains(btn)) return;
      const u = btn.getAttribute("data-unit");
      if (u !== "pct" && u !== "hours") return;
      setReportUnitMode(u);
      renderReport();
    });
    syncReportUnitToggleButtons();
  })();

  (function bindReportDataExport() {
    const btn = document.getElementById("btnReportExport");
    if (!btn) return;
    btn.addEventListener("click", () => runReportDataExport());
  })();

  // Structure CSV 已移除；舊版 preview 元素唔再綁定。

  const skipServiceWorker =
    location.protocol === "file:" ||
    location.port === "8765" ||
    location.hostname === "localhost" ||
    location.hostname === "127.0.0.1" ||
    location.hostname === "[::1]";
  if ("serviceWorker" in navigator && !skipServiceWorker) {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
})();
