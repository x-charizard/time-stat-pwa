(function () {
  const STORAGE_KEY = "time-stat-state-v1";
  const REQUIRED_PROJECTS_CSV_NAME = "Time stat V2 - Projects.csv";
  const DAY_MS = 86400000;
  const MIN_TIMELINE_MS = 30 * 60000;
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
  /** @typedef {{ id: string, start: string, activityId: string, remark?: string, people?: string[], place?: string, category?: string, group?: string, layer?: string, cat?: string, subCat?: string, structureItem?: string, project?: string, projectId?: string, objective?: string, activityQuestion?: string, achievement?: string, improveLast?: string, importantElement?: string, detailsBetter?: string, action?: string, longTermGoals?: string, shortTermGoals?: string, miniGoals?: string, groupFromForm?: string, layersFromForm?: string, projectsFromForm?: string, categoriesFromForm?: string }} Event */

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
  function loadState() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return defaultState();
      const o = JSON.parse(raw);
      const activities = Array.isArray(o.activities)
        ? o.activities
        : Array.isArray(o.entities)
          ? o.entities
          : null;
      if (!activities || !Array.isArray(o.events)) return defaultState();
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
      if (anyIdFixed) {
        try {
          localStorage.setItem(STORAGE_KEY, JSON.stringify(out));
        } catch {
          /* ignore quota */
        }
      }
      return out;
    } catch {
      return defaultState();
    }
  }

  let state = loadState();
  let reportPresetSuppress = false;
  let timelinePointerTipAbort = null;

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
  }

  (function dedupeLoadedEventsOnce() {
    const b = state.events.length;
    dedupeStateEventsByImportKey();
    if (state.events.length < b) save();
  })();

  function save() {
    state.structure = [];
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function toast(msg) {
    const el = document.getElementById("toast");
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(toast._t);
    toast._t = setTimeout(() => el.classList.remove("show"), 2200);
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

  /** 時間軸：最近三個曆日（今日、昨日、前日）由當地 0:00 起計 */
  function timelineThreeDayCutoffMs() {
    const t = new Date();
    t.setHours(0, 0, 0, 0);
    t.setDate(t.getDate() - 2);
    return t.getTime();
  }

  /** 時間軸用：窗口內、開始時間升序（計時長用） */
  function timelineEventsAscending() {
    const cutoff = timelineThreeDayCutoffMs();
    return sortedEventsUniqueById().filter((ev) => new Date(ev.start).getTime() >= cutoff);
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

  function timelineTipText(vs, t1, t0, seg) {
    const tStr = formatHmLocal(vs);
    const mins = Math.round((t1 != null ? t1 - t0 : seg) / 60000);
    return `${tStr} · ${mins} mins`;
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
    const dl = document.getElementById("activitySuggest");
    if (!dl) return;
    dl.innerHTML = "";
    const labels = new Set();
    state.activities.forEach((e) => labels.add(e.name));
    [...labels]
      .sort((a, b) => a.localeCompare(b, "zh-Hant"))
      .forEach((name) => {
        const opt = document.createElement("option");
        opt.value = name;
        dl.appendChild(opt);
      });
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

  /** Timeline 專用：同一毫秒開始 + 同一 Activity 只保留最後一筆（常見於重覆匯入），避免兩格完全疊住。 */
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

  function timelinePassesMin(ev, fullAsc) {
    const idx = fullAsc.indexOf(ev);
    if (idx < 0) return false;
    return segmentDurationMsForReport(fullAsc, idx) >= MIN_TIMELINE_MS;
  }

  function timelineDayClip(ev, colYmd, fullAsc) {
    const idx = fullAsc.indexOf(ev);
    if (idx < 0) return null;
    const segAll = segmentDurationMsForReport(fullAsc, idx);
    if (segAll < MIN_TIMELINE_MS) return null;
    const evStart = new Date(ev.start);
    const t0 = evStart.getTime();
    const t1Wall = t0 + segAll;
    const isCurrent = idx === fullAsc.length - 1;
    const { start: d0, endEx: d1 } = ymdDayBounds(colYmd);
    const vs = Math.max(t0, d0);
    let ve;
    if (isCurrent) {
      if (ymdFromLocalDate(evStart) !== colYmd) return null;
      ve = Math.min(Date.now(), d1);
    } else {
      ve = Math.min(t1Wall, d1);
    }
    if (ve <= vs) return null;
    const seg = ve - vs;
    if (seg < MIN_TIMELINE_MS) return null;
    return { vs, ve, t1: isCurrent ? null : t1Wall, t0, seg };
  }

  function renderTimeline() {
    const root = document.getElementById("timelineCalendar");
    const empty = document.getElementById("timelineEmpty");
    const nowStatus = document.getElementById("timelineNowStatus");
    const fullAsc = sortedEventsUniqueById();
    const asc = dedupeTimelineByStartAndActivity(timelineEventsAscending());
    if (!root || !empty) return;
    timelineClearRoot(root);
    if (timelinePointerTipAbort) {
      timelinePointerTipAbort.abort();
      timelinePointerTipAbort = null;
    }
    if (nowStatus) nowStatus.textContent = timelineNowStatusText();
    if (!asc.length) {
      empty.classList.remove("hidden");
      return;
    }
    empty.classList.add("hidden");

    const columns = [];
    for (let i = 2; i >= 0; i--) {
      const d = new Date();
      d.setHours(0, 0, 0, 0);
      d.setDate(d.getDate() - i);
      columns.push({ date: d, ymd: ymdFromLocalDate(d) });
    }

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
      const ok = timelinePassesMin(ev, fullAsc);
      if (ok) stripeById.set(ev.id, slot % 2 === 0 ? "a" : "b");
      slot += ok ? 1 : 2;
    }

    for (const c of columns) {
      const col = document.createElement("div");
      col.className = "timeline-cal-day";
      for (const ev of asc) {
        // 跨午夜嘅同一筆會喺「開始日」同「下一日」各有一段，睇落似重複；只喺 **start 所屬曆日** 嗰欄畫（喺該日 clip 到午夜）。
        if (ymdFromLocalDate(new Date(ev.start)) !== c.ymd) continue;
        const clip = timelineDayClip(ev, c.ymd, fullAsc);
        if (!clip) continue;
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
        blk.style.zIndex = isCurrent ? "9" : "1";

        const title = document.createElement("div");
        title.className = "timeline-cal-block-title";
        title.textContent = activityDisplayName(ev.activityId);
        blk.appendChild(title);

        const meta = document.createElement("div");
        meta.className = "timeline-cal-block-meta";
        const tip = timelineTipText(vs, t1, t0, seg);
        meta.textContent = tip;
        blk.appendChild(meta);
        blk.dataset.tip = tip;

        col.appendChild(blk);
      }
      board.appendChild(col);
    }

    body.appendChild(board);
    inner.appendChild(body);
    timelineMountRoot(root, inner);
    timelinePointerTipAbort = new AbortController();
    bindTimelinePointerTip(root, timelinePointerTipAbort.signal);
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
    tip.style.left = `${x}px`;
    tip.style.top = `${y}px`;
    tip.classList.remove("hidden");
  }

  function hideTimelineTip() {
    const tip = document.getElementById("timelineHoverTip");
    if (tip) tip.classList.add("hidden");
  }

  function bindTimelinePointerTip(root, signal) {
    if (!root || !signal) return;
    root.addEventListener(
      "mousemove",
      (e) => {
        const blk = e.target && e.target.closest ? e.target.closest(".timeline-cal-block") : null;
        if (!blk || !blk.dataset.tip) {
          hideTimelineTip();
          return;
        }
        showTimelineTip(blk.dataset.tip, e.clientX, e.clientY - 6);
      },
      { signal }
    );
    root.addEventListener("mouseleave", hideTimelineTip, { signal });
    root.addEventListener(
      "touchstart",
      (e) => {
        const t = e.touches && e.touches[0];
        if (!t) return;
        const target = document.elementFromPoint(t.clientX, t.clientY);
        const blk = target && target.closest ? target.closest(".timeline-cal-block") : null;
        if (!blk || !blk.dataset.tip) {
          hideTimelineTip();
          return;
        }
        showTimelineTip(blk.dataset.tip, t.clientX, t.clientY - 10);
      },
      { passive: true, signal }
    );
    root.addEventListener(
      "touchmove",
      (e) => {
        const t = e.touches && e.touches[0];
        if (!t) return;
        const target = document.elementFromPoint(t.clientX, t.clientY);
        const blk = target && target.closest ? target.closest(".timeline-cal-block") : null;
        if (!blk || !blk.dataset.tip) {
          hideTimelineTip();
          return;
        }
        showTimelineTip(blk.dataset.tip, t.clientX, t.clientY - 10);
      },
      { passive: true, signal }
    );
    root.addEventListener("touchend", hideTimelineTip, { passive: true, signal });
    root.addEventListener("touchcancel", hideTimelineTip, { passive: true, signal });
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
    const text = reportNormLabel(activityLabel + " " + (remark || "")).toLowerCase();
    if (!text) return [];
    const reg = Array.isArray(state.projectsRegistry) ? state.projectsRegistry : [];
    const hits = [];
    for (let i = 0; i < reg.length; i++) {
      const p = reportNormLabel(reg[i].project);
      const pid = reportNormLabel(reg[i].projectId);
      if (!p) continue;
      const pl = p.toLowerCase();
      if (text.includes(pl) || pl.includes(text)) {
        hits.push({ project: p, projectId: pid });
      }
    }
    return hits.slice(0, 2);
  }

  function buildMappingCandidates(activityLabel, remark) {
    const out = [];
    const suggested = suggestProjectsFromText(activityLabel, remark);
    const inferred = inferByHistoryOrHeuristic(activityLabel);
    const text = (String(activityLabel || "") + " " + String(remark || "")).toLowerCase();
    const looksProject = /(project|deadline|milestone|namaste|indonesia|15m|time stat|dantian|順流道)/.test(text);
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
      return hasProjectSignal ? "Project" : "";
    };
    const g0 = inferred.group;
    const l0 = inferred.layer;
    const c0 = inferred.cat;
    if (suggested.length) {
      const sp = suggested[0];
      out.push({
        label: "建議 1（由 Activity/Remark 推斷）",
        group: g0,
        layer: "Freedom",
        cat: "",
        subCat: inferSubByRules("Freedom", true),
        activity: activityLabel,
        project: sp.project,
        projectId: sp.projectId || projectIdByName(sp.project),
      });
      out.push({
        label: "建議 2（同 activity 但非 project）",
        group: g0,
        layer: l0,
        cat: c0,
        subCat: inferSubByRules(l0, false),
        activity: activityLabel,
        project: "",
        projectId: "",
      });
    } else {
      out.push({
        label: looksProject ? "建議 1（由 Remark 推斷）" : "建議 1（歷史/預設）",
        group: g0,
        layer: l0,
        cat: c0,
        subCat: inferSubByRules(l0, looksProject),
        activity: activityLabel,
        project: "",
        projectId: "",
      });
      out.push({
        label: "建議 2（保守）",
        group: g0,
        layer: l0,
        cat: c0,
        subCat: inferSubByRules(l0, false),
        activity: activityLabel,
        project: "",
        projectId: "",
      });
    }
    const dedup = [];
    const seen = new Set();
    for (let i = 0; i < out.length; i++) {
      const c = out[i];
      const key = [c.group, c.layer, c.cat, c.subCat, c.activity, c.project, c.projectId].join("|");
      if (seen.has(key)) continue;
      seen.add(key);
      dedup.push(c);
    }
    return dedup.slice(0, 3);
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
    const mentalRest = new Set(["resting", "fooding", "familying", "walking", "meditating"]);
    const physicalRest = new Set(["sleeping", "showering"]);
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
      return "Needs-Review";
    }
    const sug = suggestProjectsFromText(activityLabel, String(remark || ""));
    if (!sug.length) return "non-project";
    if (sug.length === 1) return "Project";
    return "Needs-Review";
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

    const physicalRestSet = new Set(["sleeping", "showering"]);
    const mentalRestSet = new Set(["resting", "fooding", "familying", "walking", "meditating"]);
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

  function pushEventAndRefresh(ev, msg) {
    state.events.push(ev);
    save();
    refreshActivityDatalist();
    refreshProjectPickers();
    fillMergeSelects();
    renderTimeline();
    renderReport();
    refreshQuickAutoSuggestions();
    refreshManualAutoSuggestions();
    updateLastSavedHint(ev);
    toast(`${msg} · 總筆數 ${state.events.length}`);
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
    meta.textContent = "";
    list.innerHTML = "";
    payload.candidates.forEach((c) => {
      const row = document.createElement("div");
      row.className = "mapping-suggestion";
      const editableId = uid();
      const groupOpts = reportUniqueSorted([
        ...groupsForReport(),
        "Work",
        "Rest",
      ]);
      const layerOpts = reportUniqueSorted([
        ...layersForReport(),
        "Health",
        "Freedom",
        "Achievement",
      ]);
      const catOpts = reportUniqueSorted([
        ...catsForReport(),
        "Mental Health",
        "Physical Health",
        "Time",
        "Finance",
        "Time Management",
        "Financial Management",
      ]);
      const subOpts = reportUniqueSorted([
        ...subsForReport(),
        "Long Term",
        "Short Term",
        "Project",
        "non-project",
      ]);
      const projectOpts = uniqueProjectsSorted();
      const mkSel = (id, options, selected, withEmpty) => {
        let h = `<select id="${id}" class="mapping-edit-select">`;
        if (withEmpty) h += `<option value="">blank</option>`;
        for (let i = 0; i < options.length; i++) {
          const v = options[i];
          const sel = v === selected ? ' selected' : "";
          h += `<option value="${escapeHtml(v)}"${sel}>${escapeHtml(v)}</option>`;
        }
        h += `</select>`;
        return h;
      };
      row.innerHTML =
        `<div class="mapping-grid">` +
        `<div>Group</div><div>${mkSel(`map-g-${editableId}`, groupOpts, c.group || "", true)}</div>` +
        `<div>Layers</div><div>${mkSel(`map-l-${editableId}`, layerOpts, c.layer || "", true)}</div>` +
        `<div>Cat</div><div>${mkSel(`map-c-${editableId}`, catOpts, c.cat || "", true)}</div>` +
        `<div>Sub Cat</div><div>${mkSel(`map-s-${editableId}`, subOpts, c.subCat || "", true)}</div>` +
        `<div>Activity</div><div>${escapeHtml(c.activity || "—")}</div>` +
        `<div>Project</div><div>${mkSel(`map-p-${editableId}`, projectOpts, c.project || "", true)}</div>` +
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
          const msg = pendingApproval.doneMsg;
          clearApprovalPanel();
          pushEventAndRefresh(ev, msg);
        } catch (err) {
          toast("入庫失敗：" + (err && err.message ? err.message : "未知錯誤"));
        }
      });
      row.appendChild(btn);
      list.appendChild(row);
    });
    card.classList.remove("hidden");
    setTimeout(() => {
      card.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 0);
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
      onDirectSave: function () {
        initManualDateTime();
      },
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
    const onDirectSave = params.onDirectSave;
    const candidates = buildMappingCandidates(activityLabel, remark);
    if (!candidates.length) {
      pushEventAndRefresh(ev, doneMsg);
      if (typeof onDirectSave === "function") onDirectSave();
      return;
    }
    showApprovalPanel({
      ev: ev,
      activityLabel: activityLabel,
      candidates: candidates,
      doneMsg: doneMsg,
    });
  }

  function mostLikelyPlaceByTime(hour, minute) {
    const now = Date.now();
    const monthAgo = now - 30 * DAY_MS;
    const score = new Map();
    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      const p = String(ev.place || "").trim();
      if (!p) continue;
      const t = new Date(ev.start).getTime();
      if (Number.isNaN(t) || t < monthAgo) continue;
      const d = new Date(t);
      const eh = d.getHours();
      const em = d.getMinutes();
      const diff = Math.abs(eh * 60 + em - (hour * 60 + minute));
      let w = 0;
      if (diff === 0) w = 5;
      else if (eh === hour && diff <= 5) w = 3;
      else if (eh === hour) w = 2;
      else if (diff <= 30) w = 1;
      if (!w) continue;
      score.set(p, (score.get(p) || 0) + w);
    }
    let best = "";
    let bestScore = -1;
    score.forEach((v, k) => {
      if (v > bestScore) {
        bestScore = v;
        best = k;
      }
    });
    return best;
  }

  function mostLikelyActivityByTime(hour, minute) {
    const now = Date.now();
    const monthAgo = now - 30 * DAY_MS;
    const score = new Map();
    for (let i = 0; i < state.events.length; i++) {
      const ev = state.events[i];
      const a = String(activityDisplayName(ev.activityId) || "").trim();
      if (!a || a === "（已刪 Activity）") continue;
      const t = new Date(ev.start).getTime();
      if (Number.isNaN(t) || t < monthAgo) continue;
      const d = new Date(t);
      const eh = d.getHours();
      const em = d.getMinutes();
      const diff = Math.abs(eh * 60 + em - (hour * 60 + minute));
      let w = 0;
      if (diff === 0) w = 5;
      else if (eh === hour && diff <= 5) w = 3;
      else if (eh === hour) w = 2;
      else if (diff <= 30) w = 1;
      if (!w) continue;
      score.set(a, (score.get(a) || 0) + w);
    }
    let best = "";
    let bestScore = -1;
    score.forEach((v, k) => {
      if (v > bestScore) {
        bestScore = v;
        best = k;
      }
    });
    return best;
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

  function bindSmartInput(inputId, getSuggestedValue) {
    const inp = document.getElementById(inputId);
    if (!inp) return;
    inp.addEventListener("focus", () => {
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

  function refreshQuickAutoSuggestions() {
    const d = new Date();
    applyAutoSuggestion("quickPlace", mostLikelyPlaceByTime(d.getHours(), d.getMinutes()));
    applyAutoSuggestion("quickActivity", mostLikelyActivityByTime(d.getHours(), d.getMinutes()));
  }

  function refreshManualAutoSuggestions() {
    const hm = manualSelectedHM();
    applyAutoSuggestion("manualPlace", mostLikelyPlaceByTime(hm.h, hm.m));
    applyAutoSuggestion("manualActivity", mostLikelyActivityByTime(hm.h, hm.m));
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
      return mostLikelyActivityByTime(d.getHours(), d.getMinutes());
    });
    bindSmartInput("manualActivity", () => {
      const hm = manualSelectedHM();
      return mostLikelyActivityByTime(hm.h, hm.m);
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
  const REPORT_RAW_RECORD_CAP = 400;

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

  function refreshReportFilterSelects() {
    renderReportMultiPick("reportFilterGroupBox", groupsForReport());
    renderReportMultiPick("reportFilterLayerBox", layersForReport());

    renderReportMultiPick("reportFilterCatBox", catsForReport());
    const catNow = readMultiSet("reportFilterCatBox", "reportFilterCatExtra");

    renderReportMultiPick("reportFilterSubBox", subsForReportCats(catNow));
    const subNow = readMultiSet("reportFilterSubBox", "reportFilterSubExtra");

    renderReportMultiPick("reportFilterProjectBox", projectsForReportCatsSubs(catNow, subNow));
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

  /**
   * Keyword: loose = tokens (space/comma), AND, case-insensitive, substring in any field.
   * strict = full trimmed phrase, case-sensitive, substring in concatenated row text.
   */
  function eventMatchesKeywordSearch(ev, f, list) {
    const q = String(f.keywordQuery || "").trim();
    if (!q) return true;
    const mode = f.keywordMode === "strict" ? "strict" : "loose";
    const hayJoined = eventKeywordSearchHaystack(ev, list);
    if (mode === "strict") {
      const hq = q.replace(/[\u200B-\u200D\uFEFF]/g, "");
      const ht = hayJoined.replace(/[\u200B-\u200D\uFEFF]/g, "");
      return ht.includes(hq);
    }
    const tokens = reportKeywordTokens(f.keywordQuery);
    if (!tokens.length) return true;
    const hay = normalizeForKeywordLoose(hayJoined);
    return tokens.every((tok) => hay.includes(normalizeForKeywordLoose(tok)));
  }

  /** For empty-report hints: non-keyword AND filters currently applied. */
  function reportNonKeywordFilterSummaryText(f) {
    const bits = [];
    if (f.groups && f.groups.length) bits.push("Group: " + f.groups.join(", "));
    if (f.layers && f.layers.length) bits.push("Layers: " + f.layers.join(", "));
    if (f.cats && f.cats.length) bits.push("Cat: " + f.cats.join(", "));
    if (f.subCats && f.subCats.length) bits.push("Sub: " + f.subCats.join(", "));
    if (f.projects && f.projects.length) bits.push("Project: " + f.projects.join(", "));
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
    const kwModeEl = document.getElementById("reportKeywordMode");
    return {
      groups: [...readCheckedValuesFromMsBox("reportFilterGroupBox")],
      layers: [...readCheckedValuesFromMsBox("reportFilterLayerBox")],
      cats: [...readCheckedValuesFromMsBox("reportFilterCatBox")],
      subCats: [...readCheckedValuesFromMsBox("reportFilterSubBox")],
      projects: [...readCheckedValuesFromMsBox("reportFilterProjectBox")],
      peopleQuery: (gpeo && gpeo.value) || "",
      keywordQuery: (kwEl && kwEl.value) || "",
      keywordMode: (kwModeEl && kwModeEl.value) || "loose",
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
      reportPeopleTokens(f.peopleQuery).length ||
      reportKeywordActive(f)
    );
  }

  function compareModeAnchorYmd(preset, toYmd, fromYmd) {
    return toYmd || fromYmd || ymdFromLocalDate(new Date());
  }

  function aggregateReportForRange(fromYmd, toYmd, list, f, showByDay) {
    if (!fromYmd || !toYmd) return null;
    const t0 = new Date(fromYmd + "T00:00:00").getTime();
    const t1 = new Date(toYmd + "T23:59:59.999").getTime();
    if (Number.isNaN(t0) || Number.isNaN(t1) || t0 > t1) return null;
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
      const inf = reportInferredMapping(ev, list);
      const g = inf.group || "\uff08\u672a\u6a19 Group\uff09";
      const ly = inf.layer || "\u2014";
      byGroup[g] = (byGroup[g] || 0) + ms;
      byLayer[ly] = (byLayer[ly] || 0) + ms;
      const pj = reportNormLabel(ev.projectsFromForm || "") || "\uff08\u7a7a\uff0f\u672a\u914d Project\uff09";
      const cj = inf.cat || "\uff08\u7a7a\uff0f\u672a\u914d Cat\uff09";
      const sj = inf.subCat || "\uff08\u7a7a\uff0f\u672a\u914d Sub\uff09";
      byProject[pj] = (byProject[pj] || 0) + ms;
      byCatDim[cj] = (byCatDim[cj] || 0) + ms;
      bySubDim[sj] = (bySubDim[sj] || 0) + ms;
      if (showByDay) {
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
        labels: [String(y1), String(y2), String(y3)],
        ranges: [
          { from: `${y1}-01-01`, to: `${y1}-12-31` },
          { from: `${y2}-01-01`, to: `${y2}-12-31` },
          { from: `${y3}-01-01`, to: `${y3}-12-31` },
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
      return { labels: [lab(b1), lab(b2), lab(b3)], ranges: [r1, r2, r3] };
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
      return { labels: [lab(b1), lab(b2), lab(b3)], ranges: [r1, r2, r3] };
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
        labels: [fmtShort(p1.from, p1.to), fmtShort(p2.from, p2.to), fmtShort(p3.from, p3.to)],
        ranges: [p1, p2, p3],
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
    const anchorYmd = parseYMDStrict(toEl.value) || parseYMDStrict(fromEl.value) || ymdFromLocalDate(new Date());
    const slices = buildReportComparisonSlices(preset, anchorYmd);
    if (!slices || !slices.ranges || slices.ranges.length !== 3) return;
    reportPresetSuppress = true;
    try {
      fromEl.value = slices.ranges[0].from;
      toEl.value = slices.ranges[2].to;
    } finally {
      queueMicrotask(() => {
        reportPresetSuppress = false;
      });
    }
  }

  function renderReport() {
    refreshReportFilterSelects();
    const presetElR = document.getElementById("reportPeriodPreset");
    const preset = (presetElR && presetElR.value) || "custom";
    if (!reportPresetSuppress && preset !== "custom") applyReportPeriodPreset();

    const from = parseYMDStrict(document.getElementById("reportFromStr").value);
    const to = parseYMDStrict(document.getElementById("reportToStr").value);
    const box = document.getElementById("reportSummary");
    const showByDayEl = document.getElementById("reportShowByDay");
    const showByDay = !!(showByDayEl && showByDayEl.checked);
    if (!box) return;
    if (!from || !to) {
      box.innerHTML = `<p class="muted">\u8acb\u7528<strong>\u9802\u90e8</strong>\u63b1\u597d <strong>\u7531\uff0f\u81f3</strong> \u65e5\u671f\u3002</p>`;
      return;
    }
    const t0 = new Date(from + "T00:00:00").getTime();
    const t1 = new Date(to + "T23:59:59.999").getTime();
    if (t0 > t1) {
      box.innerHTML = `<p class="muted">\u300c\u7531\u300d\u8981\u65e9\u904e\u6216\u7b49\u65bc\u300c\u81f3\u300d</p>`;
      return;
    }
    const f = readReportFilters();
    const list = sortedEventsUniqueById();

    const renderFilterMismatch = () => {
      const kw = String(f.keywordQuery || "").trim();
      const nf = reportNonKeywordFilterSummaryText(f);
      const kwHint =
        kw.length > 0 ? `<p class="muted" style="margin-top:10px;">${escapeHtml(kw)}</p>` : "";
      const other =
        nf.length > 0
          ? `<p class="muted" style="margin-top:8px;">Other active filters (AND with search): ${escapeHtml(nf)}</p>`
          : "";
      const tail =
        nf.length > 0
          ? `<p class="muted" style="margin-top:10px;">\u53ef\u8a66\u62c9\u95ca\u65e5\u671f\u7bc4\u570d\u6216\u653e\u5bec\u7be9\u9078\u3002</p>`
          : "";
      return (
        '<p class="muted">\u7bc4\u570d\u5167\u6709\u7d00\u9304\uff0c\u4f46<strong>\u7121\u4e00\u7b46\u7b26\u5408\u800c\u5bb6\u7be9\u9078</strong>\u3002</p>' +
        kwHint +
        other +
        tail
      );
    };

    if (reportComparePresetActive(preset)) {
      const anchorYmd = compareModeAnchorYmd(preset, to, from);
      const slices = buildReportComparisonSlices(preset, anchorYmd);
      if (!slices) {
        box.innerHTML = `<p class="muted">\u7121\u6cd5\u5efa\u7acb\u6bd4\u8f03\u8996\u7a97\u3002</p>`;
        return;
      }
      reportPresetSuppress = true;
      try {
        document.getElementById("reportFromStr").value = slices.ranges[0].from;
        document.getElementById("reportToStr").value = slices.ranges[2].to;
      } finally {
        queueMicrotask(() => {
          reportPresetSuppress = false;
        });
      }
      const ag = [
        aggregateReportForRange(slices.ranges[0].from, slices.ranges[0].to, list, f, false),
        aggregateReportForRange(slices.ranges[1].from, slices.ranges[1].to, list, f, false),
        aggregateReportForRange(slices.ranges[2].from, slices.ranges[2].to, list, f, false),
      ];
      if (ag[0] === null || ag[1] === null || ag[2] === null) {
        box.innerHTML = `<p class="muted">\u6bd4\u8f03\u5340\u9593\u7121\u6548\u3002</p>`;
        return;
      }
      if (ag[0].segmentsInRange + ag[1].segmentsInRange + ag[2].segmentsInRange === 0) {
        const n = state.events.length;
        box.innerHTML =
          `<p class="muted">\u6700\u8fd1\u4e09\u500b\u6bd4\u8f03\u8996\u7a97\u5167\u6c92\u6709\u53ef\u8a08\u6642\u9577\u7684\u7d00\u9304\u3002</p>` +
          `<ul class="muted" style="margin:10px 0 0;padding-left:1.2em;">` +
          `<li>\u82e5\u5c1a\u672a\u532f\u5165\uff1a<strong>\u532f\u5165 CSV</strong>\u3002</li>` +
          `<li>\u6bcf\u4e00\u65e5<strong>\u6700\u5f8c\u4e00\u689d</strong>\u7d00\u9304\u518a\u300c\u4e0b\u4e00\u7b46\u300d\uff0c\u5514\u8a08\u6642\u9577\uff08\u76ee\u524d\u5171 ${n} \u689d\uff09\u3002</li>` +
          `</ul>`;
        return;
      }
      if (ag[0].segmentsKept + ag[1].segmentsKept + ag[2].segmentsKept === 0) {
        box.innerHTML = renderFilterMismatch();
        return;
      }

      const colTotals = [ag[0].totalKept, ag[1].totalKept, ag[2].totalKept];
      const fmtCell = (ms, colIdx) =>
        colTotals[colIdx] ? ((ms / colTotals[colIdx]) * 100).toFixed(1) + "%" : "\u2014";
      const fmtH = (ms) => (ms / 3600000).toFixed(2);
      const thCols = slices.labels.map((lb) => `<th class="mono report-cmp-col">${escapeHtml(lb)}</th>`).join("");

      const tblCmp = (title, pick) => {
        const m0 = pick(ag[0]);
        const m1 = pick(ag[1]);
        const m2 = pick(ag[2]);
        const keys = new Set([...Object.keys(m0), ...Object.keys(m1), ...Object.keys(m2)]);
        const arr = [...keys].map((k) => {
          const a = m0[k] || 0;
          const b = m1[k] || 0;
          const c = m2[k] || 0;
          return { k, ms: [a, b, c], sum: a + b + c };
        });
        arr.sort((x, y) => y.sum - x.sum);
        if (!arr.length) return `<h2 class="report-h">${escapeHtml(title)}</h2><p class="muted">\uff08\u7121\uff09</p>`;
        let h =
          `<h2 class="report-h">${escapeHtml(title)}</h2><table class="report-cmp-table"><thead><tr><th>\u9805\u76ee</th>${thCols}</tr></thead><tbody>`;
        for (let i = 0; i < arr.length; i++) {
          const row = arr[i];
          h += `<tr><td>${escapeHtml(row.k)}</td>`;
          for (let j = 0; j < 3; j++) {
            h += `<td class="mono">${fmtH(row.ms[j])}<span class="muted" style="font-size:0.78em;"> (${fmtCell(row.ms[j], j)})</span></td>`;
          }
          h += `</tr>`;
        }
        h += `</tbody></table>`;
        return h;
      };

      let html = "";
      if (reportHasAnyFilter(f)) {
        const bits = [];
        if (f.groups && f.groups.length) bits.push("Group\u2208\u300c" + f.groups.join("\uff0f") + "\u300d");
        if (f.layers && f.layers.length) bits.push("Layers\u2208\u300c" + f.layers.join("\uff0f") + "\u300d");
        if (f.cats && f.cats.length) bits.push("Cat\u2208\u300c" + f.cats.join("\uff0f") + "\u300d");
        if (f.subCats && f.subCats.length) bits.push("Sub\u2208\u300c" + f.subCats.join("\uff0f") + "\u300d");
        if (f.projects && f.projects.length) bits.push("Project\u2208\u300c" + f.projects.join("\uff0f") + "\u300d");
        const ptoks = reportPeopleTokens(f.peopleQuery);
        if (ptoks.length) bits.push("\u4eba\u7269\u220b\u300c" + ptoks.join("\uff0f") + "\u300d");
        const kwTrim = String(f.keywordQuery || "").trim();
        if (kwTrim) {
          if (f.keywordMode === "strict") {
            bits.push("Keyword\uff08strict\uff09:\u300c" + kwTrim + "\u300d");
          } else {
            const kwtoks = reportKeywordTokens(f.keywordQuery);
            if (kwtoks.length) bits.push("Keyword\uff08loose, AND\uff09\u220b\u300c" + kwtoks.join("\uff0f") + "\u300d");
          }
        }
        html += `<p class="muted" style="margin:0 0 12px;">\u5df2\u5957\u7528\u7be9\u9078\uff08AND\uff09\uff1a${escapeHtml(bits.join(" \u00b7 "))}</p>`;
      }
      html += tblCmp("\u6309 Group", (a) => a.byGroup);
      html += tblCmp("\u6309 Layers", (a) => a.byLayer);
      html += tblCmp("\u6309 Cat", (a) => a.byCatDim);
      html += tblCmp("\u6309 Sub Cat", (a) => a.bySubDim);

      const actMap = new Map();
      for (let j = 0; j < 3; j++) {
        Object.entries(ag[j].byEnt).forEach(([eid, ms]) => {
          if (!actMap.has(eid)) actMap.set(eid, [0, 0, 0]);
          actMap.get(eid)[j] = ms;
        });
      }
      const actRows = [...actMap.entries()]
        .map(([eid, msco]) => ({ eid, msco, sum: msco[0] + msco[1] + msco[2] }))
        .sort((a, b) => b.sum - a.sum);
      html += `<h2 class="report-h">\u6309 Activity</h2><table class="report-cmp-table"><thead><tr><th>Activity</th>${thCols}</tr></thead><tbody>`;
      for (let i = 0; i < actRows.length; i++) {
        const row = actRows[i];
        html += `<tr><td>${escapeHtml(activityDisplayName(row.eid))}</td>`;
        for (let j = 0; j < 3; j++) {
          html += `<td class="mono">${fmtH(row.msco[j])}<span class="muted" style="font-size:0.78em;"> (${fmtCell(row.msco[j], j)})</span></td>`;
        }
        html += `</tr>`;
      }
      html += `</tbody></table>`;
      html += tblCmp("\u6309 Project", (a) => a.byProject);
      html += tblCmp("\u6309\u4eba\u7269\uff08\u6709\u586b\u300c\u540c\u908a\u500b\u4e00\u9f4a\u300d\uff09", (a) => a.byPerson);

      box.innerHTML = html;
      return;
    }

    const agg = aggregateReportForRange(from, to, list, f, showByDay);
    if (agg.segmentsInRange === 0) {
      const n = state.events.length;
      box.innerHTML =
        `<p class="muted">\u5462\u500b\u7bc4\u570d\u5167\u6c92\u6709\u53ef\u8a08\u6642\u9577\u7684\u7d00\u9304\u3002</p>` +
        `<ul class="muted" style="margin:10px 0 0;padding-left:1.2em;">` +
        `<li>\u82e5\u679c\u555f\u555f\u532f\u5165\u820a CSV\uff1a\u8acb\u5c07<strong>\u300c\u7531\uff0f\u81f3\u300d</strong>\u62c9\u5230\u5305\u4f4f\u6240\u8b02\u8cc7\u6599\u7684\u65e5\u671f\uff08\u9810\u8a2d\u6703\u8ddf\u4f4f\u4f60\u7d00\u9304\u7684\u6700\u65e9\uff0f\u6700\u5c3e\u4e00\u65e5\uff09\u3002</li>` +
        `<li>\u672c\u6a5a\u5c1a\u672a\u532f\u5165\uff1a<strong>\u532f\u5165 CSV</strong>\u5f8c\u5148\u6703\u6709\u561b\u3002</li>` +
        `<li>\u6bcf\u4e00\u65e5<strong>\u6700\u5f8c\u4e00\u689d</strong>\u7d00\u9304\u518a\u300c\u4e0b\u4e00\u7b46\u300d\uff0c\u5514\u6703\u8a08\u5165\u6642\u9577\uff08\u76ee\u524d\u5171 ${n} \u689d\u7d00\u9304\uff09\u3002</li>` +
        `</ul>`;
      return;
    }
    if (agg.segmentsKept === 0) {
      const kw = String(f.keywordQuery || "").trim();
      const nf = reportNonKeywordFilterSummaryText(f);
      const kwHint =
        kw.length > 0
          ? `<p class="muted" style="margin-top:10px;">Keyword: ${escapeHtml(kw)}${
              f.keywordMode === "strict" ? " (strict phrase, case-sensitive)" : " (loose tokens, AND, ignore case)"
            }</p>`
          : "";
      const other =
        nf.length > 0
          ? `<p class="muted" style="margin-top:8px;">Other active filters (AND with keyword): ${escapeHtml(nf)}</p>`
          : `<p class="muted" style="margin-top:8px;">\u76ee\u524d<strong>\u7121</strong>\u5257\u9078 Group\uff0fLayers\uff0fCat\uff0fSub\uff0fProject\uff0fWith \u2014\u2014 \u5373\u4fc2\u6de8\u4fc2 keyword \u55ae\u6536\u7a84\u7d50\u679c\u3002</p>`;
      const tailKw =
        kw.length > 0 && !nf.length
          ? `<ul class="muted" style="margin:10px 0 0;padding-left:1.2em;">` +
            `<li>\u672c\u6a5a\u7d00\u9304\u5165\u9762\u53ef\u80fd<strong>\u518a\u4efb\u4f55\u6b04\u4f4d</strong>\u5305\u542b\u4f60\u6253\u7684\u5b57\uff08\u4f8b\u5982 project \u672a\u5beb\u5165\u3001\u62fc\u6cd5\u5514\u540c\uff09\u3002</li>` +
            `<li>\u8acb\u78ba\u8a8d\u5df2\u7528 <strong>Import CSV</strong> \u63b1\u597d <strong>Projects</strong>\uff0f<strong>What is the project</strong> \u6b04\u5c0d\u61c9\uff0c\u540c\u57cb\u5df2\u532f\u5165 <strong>Time stat V2 - Projects.csv</strong>\u3002</li>` +
            `<li>\u82e5\u4f60\u5257\u5de6\u908a <strong>Project</strong> \u5257\u9078\u6846\uff0c\u8981\u540c Raw \u8868\u300cProject\u300d\u6b04<strong>\u5b8c\u5168\u4e00\u81f4</strong>\uff08\u6216\u62c6\u6bb5\u4e4b\u4e00\uff09\u5148\u6703\u8a08\u5165\u3002</li>` +
            `</ul>`
          : `<p class="muted" style="margin-top:10px;">\u53ef\u8a66\u62c9\u95ca\u65e5\u671f\u7bc4\u570d\uff0c\u6216\u653e\u5bec\u5de6\u908a\u7be9\u9078\u3002</p>`;
      box.innerHTML =
        `<p class="muted">\u7bc4\u570d\u5167\u6709\u8a08\u5230\u6642\u9577\u7684\u7d00\u9304\uff0c\u4f46<strong>\u5514\u7b26\u5408\u800c\u5bb6\u7be9\u9078</strong>\u3002</p>` + kwHint + other + tailKw;
      return;
    }

    const { byEnt, byGroup, byLayer, byProject, byCatDim, bySubDim, byDay, byPerson, rawSegmentRows } = agg;
    const rows = Object.entries(byEnt).sort((a, b) => b[1] - a[1]);
    let total = 0;
    for (let ri = 0; ri < rows.length; ri++) total += rows[ri][1];
    const tbl = (title, map) => {
      const r = Object.entries(map).sort((a, b) => b[1] - a[1]);
      if (!r.length) return `<h2 class="report-h">${title}</h2><p class="muted">\uff08\u7121\uff09</p>`;
      let h = `<h2 class="report-h">${title}</h2><table><thead><tr><th>\u9805\u76ee</th><th>\u5c0f\u6642</th></tr></thead><tbody>`;
      for (let i = 0; i < r.length; i++) {
        const k = r[i][0];
        const ms = r[i][1];
        h += `<tr><td>${escapeHtml(k)}</td><td class="mono">${(ms / 3600000).toFixed(2)}</td></tr>`;
      }
      h += `</tbody></table>`;
      return h;
    };
    let html = "";
    if (reportHasAnyFilter(f)) {
      const bits = [];
      if (f.groups && f.groups.length) bits.push("Group\u2208\u300c" + f.groups.join("\uff0f") + "\u300d");
      if (f.layers && f.layers.length) bits.push("Layers\u2208\u300c" + f.layers.join("\uff0f") + "\u300d");
      if (f.cats && f.cats.length) bits.push("Cat\u2208\u300c" + f.cats.join("\uff0f") + "\u300d");
      if (f.subCats && f.subCats.length) bits.push("Sub\u2208\u300c" + f.subCats.join("\uff0f") + "\u300d");
      if (f.projects && f.projects.length) bits.push("Project\u2208\u300c" + f.projects.join("\uff0f") + "\u300d");
      const ptoks = reportPeopleTokens(f.peopleQuery);
      if (ptoks.length) bits.push("\u4eba\u7269\u220b\u300c" + ptoks.join("\uff0f") + "\u300d");
      const kwTrim = String(f.keywordQuery || "").trim();
      if (kwTrim) {
        if (f.keywordMode === "strict") {
          bits.push("Keyword\uff08strict phrase, case-sensitive\uff09:\u300c" + kwTrim + "\u300d");
        } else {
          const kwtoks = reportKeywordTokens(f.keywordQuery);
          if (kwtoks.length) bits.push("Keyword\uff08loose, AND\uff09\u220b\u300c" + kwtoks.join("\uff0f") + "\u300d");
        }
      }
      html += `<p class="muted" style="margin:0 0 12px;">\u5df2\u5957\u7528\u7be9\u9078\uff08AND\uff09\uff1a${escapeHtml(bits.join(" \u00b7 "))}</p>`;
    }
    html += tbl("\u6309 Group", byGroup);
    html += tbl("\u6309 Layers", byLayer);
    html += tbl("\u6309 Cat", byCatDim);
    html += tbl("\u6309 Sub Cat", bySubDim);
    html += `<h2 class="report-h">\u6309 Activity</h2><table><thead><tr><th>Activity</th><th>\u5c0f\u6642</th></tr></thead><tbody>`;
    for (let i = 0; i < rows.length; i++) {
      const eid = rows[i][0];
      const ms = rows[i][1];
      html += `<tr><td>${escapeHtml(activityDisplayName(eid))}</td><td class="mono">${(ms / 3600000).toFixed(2)}</td></tr>`;
    }
    html += `</tbody></table>`;
    html += tbl("\u6309 Project", byProject);
    if (showByDay) {
      const dayKeys = Object.keys(byDay).sort();
      if (dayKeys.length) {
        html += `<h2 class="report-h">\u6bcf\u65e5\u5c0f\u8a08</h2><table><thead><tr><th>\u65e5\u671f</th><th>\u5c0f\u6642</th></tr></thead><tbody>`;
        for (let di = 0; di < dayKeys.length; di++) {
          const d = dayKeys[di];
          const ms = byDay[d];
          html += `<tr><td class="mono">${escapeHtml(d)}</td><td class="mono">${(ms / 3600000).toFixed(2)}</td></tr>`;
        }
        html += `</tbody></table>`;
      }
    }
    html += tbl("\u6309\u4eba\u7269\uff08\u6709\u586b\u300c\u540c\u908a\u500b\u4e00\u9f4a\u300d\uff09", byPerson);

    const cap = REPORT_RAW_RECORD_CAP;
    const sortedRaw = [...rawSegmentRows].sort((a, b) => new Date(a.ev.start) - new Date(b.ev.start));
    const sliceRaw = sortedRaw.slice(0, cap);
    const nextGlobal = chronologicalNextById(list);
    html += `<h2 class="report-h">Raw records\uff08\u542b Remark\uff09</h2><p class="muted" style="margin:0 0 8px;">\u8207\u4e0a\u9762\u540c\u4e00\u7be9\u9078\u540c\u65e5\u671f\u7bc4\u570d\uff1b\u6bcf\u884c\u4e00\u6bb5\u8a08\u6642\u3002<strong>Group</strong>\uff1aCSV\uff0f\u8868\u55ae\u6709 Work\uff0fRest \u5c31\u7528\u5165\u5eab\u503c\uff1b\u7121\u5247\u6309\u300aTime Stat mapping rules\u300b\u4f30\uff08\u53ea\u6703\u4fc2 Work \u6216 Rest\uff09\u3002<strong>Layers\uff0fCat\uff0fSub</strong>\uff1a\u6309\u540c\u4e00\u5957 rules \u63a8\u65b7\uff1bFreedom \u4e0b Cat \u756b\u9762\u986f\u793a <strong>Time Management</strong>\u3002<strong>Project</strong>\uff1a<strong>\u53ea</strong>\u986f\u793a CSV\uff0f\u8868\u55ae\u300cWhat is the project\u2026\u300d\u7b49\u532f\u5165\u7684 <code>projectsFromForm</code>\uff08\u540c\u4f60\u63b1\u7684 Projects \u6b04\u6703\u5408\u4f75\u5165\u5462\u500b\u6b04\u4f4d\uff09\uff0c<strong>\u5514\u7528</strong> Structure\u3001\u4ea6<strong>\u5514\u7528</strong> <code>ev.project</code>\u3002<strong>Remark</strong>\uff1a\u5408\u4f75 Remark\u3001Notes\u3001Description \u7b49\u3002\u5217\u8868<strong>\u9806\u6642\u5e8f</strong>\uff08\u65e9\u2192\u9072\uff09\u3002<strong>\u540c\u4e00\u79d2\u958b\u59cb\u7684\u591a\u7b46</strong>\u6703\u5c07\u4e2d\u9593\u6642\u9577<strong>\u5e73\u5747\u6524\u5206</strong>\uff1b<strong>\u540c\u4e00 id \u91cd\u8986\u5165\u5eab</strong>\u6642\u53ea\u4fdd\u7559\u6642\u9593\u5e8f<strong>\u6700\u5f8c\u4e00\u7b46</strong>\u8a08\u5165\u5831\u8868\uff0f\u532f\u51fa\uff0f\u6642\u9593\u8ef8\u3002</p>`;
    html += `<div class="report-records-wrap"><table class="report-records-table"><thead><tr><th>Start</th><th>Duration</th><th>Group</th><th>Layers</th><th>Cat</th><th>Sub Cat</th><th>Activity</th><th>Project</th><th>Place</th><th>Remark</th><th>With</th></tr></thead><tbody>`;
    for (let ri = 0; ri < sliceRaw.length; ri++) {
      const row = sliceRaw[ri];
      const ev = row.ev;
      const ms = row.ms;
      const startStr = ymdHmFromEventStart(ev.start);
      const durStr = durationMinutesLabel(ms);
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
      html +=
        `<tr><td class="mono">${escapeHtml(startStr)}</td><td class="mono">${escapeHtml(durStr)}</td>` +
        `<td>${escapeHtml(g)}</td><td>${escapeHtml(ly)}</td><td>${escapeHtml(cj)}</td><td>${escapeHtml(sj)}</td>` +
        `<td>${escapeHtml(act)}</td><td>${escapeHtml(pj)}</td><td>${escapeHtml(place)}</td><td class="remark-cell">${escapeHtml(remark)}</td><td>${escapeHtml(withStr)}</td></tr>`;
    }
    html += `</tbody></table></div>`;
    if (sortedRaw.length > cap) {
      html += `<p class="muted" style="margin-top:6px;">\u53ea\u986f\u793a\u524d ${cap} \u6bb5\uff08\u5171 ${sortedRaw.length} \u6bb5\uff1b\u9806\u6642\u5e8f\uff09\u3002</p>`;
    }

    html += `<p class="muted" style="margin-top:10px;">\u5408\u8a08\uff1a<strong style="color:var(--text);">${(total / 3600000).toFixed(2)}</strong> \u5c0f\u6642\uff08\u50c5\u8a08\u6709\u300c\u4e0b\u4e00\u7b46\u300d\u7684\u5340\u9593\uff1b\u5df2\u5957\u7528\u4e0a\u9762\u7be9\u9078\uff09</p>`;
    box.innerHTML = html;
  }

  document.getElementById("btnReport").addEventListener("click", () => renderReport());

  document.getElementById("btnExportJson").addEventListener("click", () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "time-stat-backup-" + new Date().toISOString().slice(0, 10) + ".json";
    a.click();
    URL.revokeObjectURL(a.href);
    toast("已下載 JSON");
  });

  document.getElementById("btnImportJson").addEventListener("change", function () {
    const f = this.files && this.files[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = () => {
      try {
        const o = JSON.parse(r.result);
        const activities = Array.isArray(o.activities)
          ? o.activities
          : Array.isArray(o.entities)
            ? o.entities
            : null;
        if (!activities || !Array.isArray(o.events)) throw new Error("格式唔似備份");
        const events = o.events.map((ev) => {
          const n = { ...ev };
          if (n.activityId == null && n.entityId != null) n.activityId = n.entityId;
          delete n.entityId;
          if (n.id == null || n.id === "") n.id = uid();
          return n;
        });
        state = {
          version: o.version || 3,
          activities,
          events,
          structure: [],
          projectsRegistry: Array.isArray(o.projectsRegistry) ? o.projectsRegistry : [],
        };
        dedupeStateEventsByImportKey();
        save();
        refreshActivityDatalist();
        fillMergeSelects();
        renderActivityList();
        renderTimeline();
        syncReportDatesFromEvents();
        renderReport();
        toast("已還原備份");
      } catch (e) {
        toast("匯入失敗：" + (e.message || ""));
      }
      this.value = "";
    };
    r.readAsText(f);
  });

  document.getElementById("btnExportCsv").addEventListener("click", () => {
    const list = sortedEventsUniqueById();
    const lines = [
      "start_iso,activity,place,category,group,layer,cat,subCat,project,project_id,people,remark,objective,activityQuestion,groupFromForm,layersFromForm,projectsFromForm,categoriesFromForm,achievement,improveLast,importantElement,detailsBetter,action,longTermGoals,shortTermGoals,miniGoals,duration_to_next_sec",
    ];
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const ms = segmentDurationMsForReport(list, i);
      const sec = Math.round(ms / 1000);
      lines.push(
        [
          ev.start,
          csv(activityDisplayName(ev.activityId)),
          csv(ev.place || ""),
          csv(ev.category || ""),
          csv(ev.group || ""),
          csv(ev.layer || ""),
          csv(ev.cat || ""),
          csv(ev.subCat || ""),
          csv(ev.projectsFromForm || ""),
          csv(ev.projectId || ""),
          csv((ev.people || []).join(";")),
          csv(ev.remark || ""),
          csv(ev.objective || ""),
          csv(ev.activityQuestion || ""),
          csv(ev.groupFromForm || ""),
          csv(ev.layersFromForm || ""),
          csv(ev.projectsFromForm || ""),
          csv(ev.categoriesFromForm || ""),
          csv(ev.achievement || ""),
          csv(ev.improveLast || ""),
          csv(ev.importantElement || ""),
          csv(ev.detailsBetter || ""),
          csv(ev.action || ""),
          csv(ev.longTermGoals || ""),
          csv(ev.shortTermGoals || ""),
          csv(ev.miniGoals || ""),
          sec,
        ].join(",")
      );
    }
    const blob = new Blob(["\ufeff" + lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "time-stat-export.csv";
    a.click();
    URL.revokeObjectURL(a.href);
    toast("已匯出 CSV");
  });

  function csv(s) {
    const t = String(s).replace(/"/g, '""');
    if (/[",\n\r]/.test(t)) return '"' + t + '"';
    return t;
  }

  /** ---------- CSV import (Papa) ---------- */
  let lastParsed = null;

  document.getElementById("csvFile").addEventListener("change", function () {
    const f = this.files && this.files[0];
    if (!f || typeof Papa === "undefined") {
      if (typeof Papa === "undefined") toast("Papa Parse 未載入");
      return;
    }
    const seenH = {};
    function transformHeader(h) {
      const base = h == null || h === "" ? "Column" : String(h);
      seenH[base] = (seenH[base] || 0) + 1;
      if (seenH[base] === 1) return base;
      return base + "__" + seenH[base];
    }
    Papa.parse(f, {
      header: true,
      transformHeader,
      skipEmptyLines: false,
      complete: (res) => {
        lastParsed = res;
        const headers = res.meta.fields || [];
        const selTs = document.getElementById("mapTimestamp");
        const selAct = document.getElementById("mapActivity");
        const selPlace = document.getElementById("mapPlace");
        const selCat = document.getElementById("mapCategory");
        const selProj = document.getElementById("mapProjects");
        [selTs, selAct, selPlace, selCat, selProj].forEach((sel) => {
          if (!sel) return;
          sel.innerHTML = '<option value="">—</option>';
          headers.forEach((h) => {
            const o = document.createElement("option");
            o.value = h;
            o.textContent = h;
            sel.appendChild(o);
          });
        });
        function pick(pred) {
          for (let i = 0; i < headers.length; i++) {
            if (pred(headers[i])) return headers[i];
          }
          return "";
        }
        selTs.value = pick((h) => h === "Timestamp") || pick((h) => /timestamp/i.test(h)) || "";
        selAct.value =
          pick((h) => h === "Category__2") ||
          pick((h) => h === "Activities") ||
          pick((h) => /^activity$/i.test(h)) ||
          "";
        selPlace.value =
          pick((h) => /^Place__\d+$/.test(h)) ||
          pick((h) => h.startsWith("Place__")) ||
          pick((h) => h === "Place") ||
          "";
        selCat.value = pick((h) => h === "Category") || "";
        if (selProj) {
          selProj.value =
            pick((h) => /^what\s+is\s+the\s+project/i.test(String(h).replace(/\u3000/g, " "))) ||
            pick((h) => h === "Projects") ||
            pick((h) => h === "What is the Projects?") ||
            pick((h) => /^What is the Projects\?__\d+$/i.test(h)) ||
            pick((h) => /^Projects__\d+$/i.test(h)) ||
            "";
        }

        document.getElementById("importPreview").textContent =
          `讀取 ${res.data.length} 行；請確認欄位對應再按「匯入」。`;
        toast("CSV 已解析");
      },
      error: (err) => toast("CSV 錯：" + err.message),
    });
    this.value = "";
  });

  (function () {
    const projectsInp = document.getElementById("projectsCsvFile");
    if (!projectsInp) return;
    projectsInp.addEventListener("change", function () {
      const f = this.files && this.files[0];
      if (!f || typeof Papa === "undefined") {
        if (typeof Papa === "undefined") toast("Papa Parse 未載入");
        return;
      }
      if (f.name !== REQUIRED_PROJECTS_CSV_NAME) {
        toast(`Projects 只接受：${REQUIRED_PROJECTS_CSV_NAME}`);
        this.value = "";
        return;
      }
      Papa.parse(f, {
        header: true,
        skipEmptyLines: "greedy",
        complete: (res) => {
          const rows = parseProjectsCsvRows(res.data || []);
          state.projectsRegistry = rows;
          save();
          refreshProjectPickers();
          const prev = document.getElementById("projectsImportPreview");
          if (prev) prev.textContent = `Projects：${rows.length} 行（已寫入本機）`;
          toast("已匯入 Projects");
        },
        error: (err) => toast("Projects CSV 錯：" + err.message),
      });
      this.value = "";
    });
  })();

  document.getElementById("btnImportCsv").addEventListener("click", () => {
    if (!lastParsed || !lastParsed.data) {
      toast("請先揀 CSV 檔");
      return;
    }
    const tsCol = document.getElementById("mapTimestamp").value;
    const actCol = document.getElementById("mapActivity").value;
    if (!tsCol || !actCol) {
      toast("揀 Timestamp 同 Activity 欄");
      return;
    }
    const placeCol = document.getElementById("mapPlace").value;
    const catCol = document.getElementById("mapCategory").value;
    const projCol = (document.getElementById("mapProjects") && document.getElementById("mapProjects").value) || "";
    let n = 0;
    let skip = 0;
    let dupSkip = 0;
    const rows = lastParsed.data;
    const importKeySeen = new Set();
    for (let ei = 0; ei < state.events.length; ei++) {
      const k0 = eventImportDedupeKey(state.events[ei]);
      if (k0 && !String(k0).startsWith("__badtime:")) importKeySeen.add(k0);
    }
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rawTs = row[tsCol];
      const iso = parseSheetsTimestamp(rawTs);
      if (!iso) {
        skip++;
        continue;
      }
      const actLabel = String(row[actCol] || "").trim();
      if (!actLabel) {
        skip++;
        continue;
      }
      let ent = getOrCreateActivity(actLabel);
      if (!ent) continue;
      const ev = {
        id: uid(),
        start: iso,
        activityId: ent.id,
      };
      if (placeCol && row[placeCol]) ev.place = String(row[placeCol]).trim();
      if (catCol && row[catCol]) ev.category = String(row[catCol]).trim();
      assignOptionalFormFields(ev, row);
      mergeImportCsvLooseFields(ev, row, [tsCol, actCol, placeCol, catCol, projCol]);
      const w2 = row["With__2"];
      const w0 = row["With"];
      if (w2 && String(w2).trim()) ev.people = splitPeople(w2);
      else if (w0 && String(w0).trim()) ev.people = splitPeople(w0);
      if (projCol && row[projCol] != null) {
        const pv = String(row[projCol]).trim();
        if (pv) {
          if (ev.projectsFromForm) {
            if (!ev.projectsFromForm.toLowerCase().includes(pv.toLowerCase())) {
              ev.projectsFromForm = ev.projectsFromForm + " · " + pv;
            }
          } else {
            ev.projectsFromForm = pv;
          }
          const toks = String(ev.projectsFromForm || "")
            .split(/\s*[·,，、]\s*/)
            .map((x) => x.trim())
            .filter(Boolean);
          const firstTok = toks[0];
          if (firstTok) {
            const pid = projectIdByName(firstTok);
            if (pid) ev.projectId = pid;
            else delete ev.projectId;
          }
        }
      }
      const dk = eventImportDedupeKey(ev);
      if (dk && !String(dk).startsWith("__badtime:") && importKeySeen.has(dk)) {
        dupSkip++;
        continue;
      }
      if (dk && !String(dk).startsWith("__badtime:")) importKeySeen.add(dk);
      state.events.push(ev);
      n++;
    }
    const afterPush = state.events.length;
    dedupeStateEventsByImportKey();
    const afterDedup = state.events.length;
    const removedByDedup = afterPush - afterDedup;
    save();
    refreshActivityDatalist();
    fillMergeSelects();
    renderActivityList();
    renderTimeline();
    syncReportDatesFromEvents();
    renderReport();
    toast(
      `已處理 ${rows.length} 行：新加入 ${n} 筆，略過格式 ${skip} 行，匯入階段重複 ${dupSkip} 筆；整庫合併再刪 ${removedByDedup} 筆，現共 ${afterDedup} 筆紀錄。`
    );
  });

  document.querySelectorAll(".tabs button").forEach((btn) => {
    btn.addEventListener("click", () => {
      const tab = btn.getAttribute("data-tab");
      document.querySelectorAll(".tabs button").forEach((b) => {
        b.classList.toggle("active", b === btn);
      });
      document.querySelectorAll("[data-panel]").forEach((p) => {
        p.classList.toggle("hidden", p.getAttribute("data-panel") !== tab);
      });
      if (tab === "report") renderReport();
    });
  });

  /** 無紀錄時：最近 7 日；有紀錄時：覆蓋資料最早～最尾一日（避免舊 CSV 跌出預設範圍） */
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
    const b = monthBoundsYMD(new Date());
    fromEl.value = b.from;
    toEl.value = b.to;
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

  function finalizeWheelScroll(viewport, hiddenInput, modulus) {
    const itemH = WHEEL_ITEM_H;
    let idx = Math.round(viewport.scrollTop / itemH);
    if (idx < modulus) viewport.scrollTop += modulus * itemH;
    else if (idx >= modulus * 2) viewport.scrollTop -= modulus * itemH;
    idx = Math.round(viewport.scrollTop / itemH);
    const v = ((idx % modulus) + modulus) % modulus;
    hiddenInput.value = String(v).padStart(2, "0");
    updateManualTimeSummary();
  }

  function attachWheel(viewport, hiddenInput, modulus) {
    let debounceT;
    function schedule() {
      clearTimeout(debounceT);
      debounceT = setTimeout(() => finalizeWheelScroll(viewport, hiddenInput, modulus), 90);
    }
    viewport.addEventListener("scroll", schedule, { passive: true });
    viewport.addEventListener("touchend", () => setTimeout(() => finalizeWheelScroll(viewport, hiddenInput, modulus), 150));
    viewport.addEventListener("scrollend", () => finalizeWheelScroll(viewport, hiddenInput, modulus));
  }

  function setWheelToValue(viewport, hiddenInput, modulus, valNum) {
    const v = Math.max(0, Math.min(modulus - 1, Math.floor(valNum)));
    const idx = modulus + v;
    viewport.scrollTop = idx * WHEEL_ITEM_H;
    hiddenInput.value = String(v).padStart(2, "0");
    requestAnimationFrame(() => {
      requestAnimationFrame(() => finalizeWheelScroll(viewport, hiddenInput, modulus));
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

  refreshActivityDatalist();
  refreshProjectPickers();
  bindPlaceAutoSuggest();
  refreshQuickAutoSuggestions();
  fillMergeSelects();
  renderActivityList();
  renderTimeline();
  syncReportDatesFromEvents();
  renderReport();
  initManualDateTime();
  refreshManualAutoSuggestions();
  updateLastSavedHint();
  const cancelBtn = document.getElementById("btnMappingCancel");
  if (cancelBtn) cancelBtn.addEventListener("click", clearApprovalPanel);

  (function bindReportFilters() {
    const msBoxes = [
      "reportFilterGroupBox",
      "reportFilterLayerBox",
      "reportFilterCatBox",
      "reportFilterSubBox",
      "reportFilterProjectBox",
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
    const kwMode = document.getElementById("reportKeywordMode");
    if (kwMode) kwMode.addEventListener("change", () => renderReport());
    const cb = document.getElementById("reportShowByDay");
    if (cb) cb.addEventListener("change", () => renderReport());

    const presetEl = document.getElementById("reportPeriodPreset");
    if (presetEl) {
      presetEl.addEventListener("change", () => {
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
