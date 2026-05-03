(function () {
  const STORAGE_KEY = "time-stat-state-v1";

  /** @typedef {{ id: string, name: string, aliases: string[] }} Activity */
  /** @typedef {{ id: string, start: string, activityId: string, remark?: string, people?: string[], place?: string, category?: string }} Event */

  function defaultState() {
    return {
      version: 1,
      activities: [],
      events: [],
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
      const events = o.events.map((ev) => {
        const n = { ...ev };
        if (n.activityId == null && n.entityId != null) n.activityId = n.entityId;
        delete n.entityId;
        return n;
      });
      return { version: o.version || 1, activities, events };
    } catch {
      return defaultState();
    }
  }

  let state = loadState();

  function save() {
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

  function durationMs(ev, nextEv) {
    if (!nextEv) return null;
    return new Date(nextEv.start) - new Date(ev.start);
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
    state.activities.forEach((e) => {
      const opt = document.createElement("option");
      opt.value = e.name;
      dl.appendChild(opt);
    });
  }

  function renderTimeline() {
    const tbody = document.getElementById("timelineBody");
    const empty = document.getElementById("timelineEmpty");
    const list = sortedEvents();
    tbody.innerHTML = "";
    if (list.length === 0) {
      empty.classList.remove("hidden");
      return;
    }
    empty.classList.add("hidden");
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const next = list[i + 1];
      const ms = durationMs(ev, next);
      const tr = document.createElement("tr");
      tr.innerHTML =
        `<td class="mono">${escapeHtml(new Date(ev.start).toLocaleString("zh-Hant"))}</td>` +
        `<td>${escapeHtml(activityDisplayName(ev.activityId))}</td>` +
        `<td class="mono">${ms != null ? formatDur(ms) : "—"}</td>` +
        `<td class="muted">${escapeHtml([ev.place, ev.category].filter(Boolean).join(" · ") || "—")}</td>` +
        `<td class="muted">${escapeHtml((ev.people || []).join(", ") || "—")}</td>` +
        `<td class="muted">${escapeHtml(ev.remark || "—")}</td>`;
      const tdAct = document.createElement("td");
      tdAct.className = "mono";
      const bEdit = document.createElement("button");
      bEdit.type = "button";
      bEdit.className = "ghost";
      bEdit.textContent = "改時間";
      bEdit.addEventListener("click", () => openEditDialog(ev));
      const bDel = document.createElement("button");
      bDel.type = "button";
      bDel.className = "danger";
      bDel.textContent = "刪";
      bDel.style.marginTop = "6px";
      bDel.addEventListener("click", () => {
        if (!confirm("刪除此筆？")) return;
        state.events = state.events.filter((x) => x.id !== ev.id);
        save();
        renderTimeline();
        renderReport();
        toast("已刪除");
      });
      tdAct.appendChild(bEdit);
      tdAct.appendChild(bDel);
      tr.appendChild(tdAct);
      tbody.appendChild(tr);
    }
  }

  function openEditDialog(ev) {
    const iso = ev.start.slice(0, 16);
    const input = prompt("修改開始時間（格式：YYYY-MM-DDTHH:MM，本地）", iso);
    if (input == null) return;
    const d = new Date(input);
    if (Number.isNaN(d.getTime())) {
      toast("時間格式唔啱");
      return;
    }
    ev.start = d.toISOString();
    save();
    renderTimeline();
    renderReport();
    toast("已更新");
  }

  function setDatetimeLocalNow(el) {
    if (!el) return;
    const d = new Date();
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
    el.value = d.toISOString().slice(0, 16);
  }

  document.getElementById("btnLogNow").addEventListener("click", () => {
    const label = document.getElementById("quickActivity").value.trim();
    const act = getOrCreateActivity(label);
    if (!act) {
      toast("請輸入 Activity 名稱");
      return;
    }
    const dt = document.getElementById("quickStart").value;
    if (!dt) {
      toast("請填開始時間");
      return;
    }
    const d = new Date(dt);
    if (Number.isNaN(d.getTime())) {
      toast("開始時間唔有效");
      return;
    }
    state.events.push({
      id: uid(),
      start: d.toISOString(),
      activityId: act.id,
      remark: document.getElementById("quickRemark").value.trim() || undefined,
      people: splitPeople(document.getElementById("quickPeople").value),
      place: document.getElementById("quickPlace").value.trim() || undefined,
      category: document.getElementById("quickCategory").value.trim() || undefined,
    });
    save();
    refreshActivityDatalist();
    fillMergeSelects();
    renderTimeline();
    renderReport();
    setDatetimeLocalNow(document.getElementById("quickStart"));
    toast("已記錄");
  });

  document.getElementById("btnManual").addEventListener("click", () => {
    const label = document.getElementById("manualActivity").value.trim();
    const act = getOrCreateActivity(label);
    const dt = document.getElementById("manualStart").value;
    if (!act || !dt) {
      toast("輸入 Activity 名稱 + 開始時間");
      return;
    }
    const d = new Date(dt);
    if (Number.isNaN(d.getTime())) {
      toast("時間唔有效");
      return;
    }
    state.events.push({
      id: uid(),
      start: d.toISOString(),
      activityId: act.id,
      remark: document.getElementById("manualRemark").value.trim() || undefined,
      people: splitPeople(document.getElementById("manualPeople").value),
      place: document.getElementById("manualPlace").value.trim() || undefined,
      category: document.getElementById("manualCategory").value.trim() || undefined,
    });
    save();
    refreshActivityDatalist();
    fillMergeSelects();
    renderTimeline();
    renderReport();
    toast("已加入（後補）");
  });

  function splitPeople(s) {
    return String(s || "")
      .split(/[,，、]/)
      .map((x) => x.trim())
      .filter(Boolean);
  }

  function renderActivityList() {
    const root = document.getElementById("activityCards");
    root.innerHTML = "";
    state.activities.forEach((e) => {
      const card = document.createElement("div");
      card.className = "card";
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

  document.getElementById("btnAddActivity").addEventListener("click", () => {
    const name = document.getElementById("newActivityName").value.trim();
    if (!name) {
      toast("輸入名稱");
      return;
    }
    if (resolveActivityByLabel(name)) {
      toast("已有同名／alias");
      return;
    }
    state.activities.push({ id: uid(), name, aliases: [] });
    document.getElementById("newActivityName").value = "";
    save();
    refreshActivityDatalist();
    fillMergeSelects();
    renderActivityList();
    toast("已新增");
  });

  document.getElementById("btnMerge").addEventListener("click", () => {
    const fromId = document.getElementById("mergeFrom").value;
    const toId = document.getElementById("mergeTo").value;
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

  function renderReport() {
    const from = document.getElementById("reportFrom").value;
    const to = document.getElementById("reportTo").value;
    const box = document.getElementById("reportSummary");
    if (!from || !to) {
      box.innerHTML = `<p class="muted">揀日期範圍</p>`;
      return;
    }
    const t0 = new Date(from + "T00:00:00").getTime();
    const t1 = new Date(to + "T23:59:59.999").getTime();
    const list = sortedEvents();
    const byEnt = {};
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const st = new Date(ev.start).getTime();
      if (st < t0 || st > t1) continue;
      const next = list[i + 1];
      const ms = durationMs(ev, next);
      if (ms == null) continue;
      byEnt[ev.activityId] = (byEnt[ev.activityId] || 0) + ms;
    }
    const rows = Object.entries(byEnt).sort((a, b) => b[1] - a[1]);
    if (rows.length === 0) {
      const n = state.events.length;
      box.innerHTML =
        `<p class="muted">呢個範圍內沒有可計時長嘅紀錄。</p>` +
        `<ul class="muted" style="margin:10px 0 0;padding-left:1.2em;">` +
        `<li>若果啱啱匯入舊 CSV：請將<strong>「由／至」</strong>拉到包住所謂資料嘅日期（預設會跟住你紀錄嘅最早／最尾一日）。</li>` +
        `<li>本機尚未匯入：<strong>匯入 CSV</strong> 後先會有嘢。</li>` +
        `<li>每一日<strong>最後一條</strong>紀錄冇「下一筆」，唔會計入時長（目前共 ${n} 條紀錄）。</li>` +
        `</ul>`;
      return;
    }
    let html = `<table><thead><tr><th>Activity</th><th>小時</th></tr></thead><tbody>`;
    let total = 0;
    rows.forEach(([eid, ms]) => {
      total += ms;
      html += `<tr><td>${escapeHtml(activityDisplayName(eid))}</td><td class="mono">${(ms / 3600000).toFixed(2)}</td></tr>`;
    });
    html += `</tbody></table><p class="muted" style="margin-top:10px;">合計：<strong style="color:var(--text);">${(total / 3600000).toFixed(2)}</strong> 小時（僅計有「下一筆」嘅區間）</p>`;
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
          return n;
        });
        state = { version: o.version || 1, activities, events };
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
    const list = sortedEvents();
    const lines = ["start_iso,activity,place,category,people,remark,duration_to_next_sec"];
    for (let i = 0; i < list.length; i++) {
      const ev = list[i];
      const next = list[i + 1];
      const ms = durationMs(ev, next);
      const sec = ms != null ? Math.round(ms / 1000) : "";
      lines.push(
        [
          ev.start,
          csv(activityDisplayName(ev.activityId)),
          csv(ev.place || ""),
          csv(ev.category || ""),
          csv((ev.people || []).join(";")),
          csv(ev.remark || ""),
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
        [selTs, selAct, selPlace, selCat].forEach((sel) => {
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
        selAct.value = pick((h) => h === "Activities") || pick((h) => /^activity$/i.test(h)) || "";
        selPlace.value =
          pick((h) => /^Place__\d+$/.test(h)) ||
          pick((h) => h.startsWith("Place__")) ||
          pick((h) => h === "Place") ||
          "";
        selCat.value = pick((h) => h === "Category") || pick((h) => h === "Category__2") || "";

        document.getElementById("importPreview").textContent =
          `讀取 ${res.data.length} 行；請確認欄位對應再按「匯入」。`;
        toast("CSV 已解析");
      },
      error: (err) => toast("CSV 錯：" + err.message),
    });
    this.value = "";
  });

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
    let n = 0;
    let skip = 0;
    const rows = lastParsed.data;
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
      state.events.push(ev);
      n++;
    }
    save();
    refreshActivityDatalist();
    fillMergeSelects();
    renderActivityList();
    renderTimeline();
    syncReportDatesFromEvents();
    renderReport();
    toast(`已匯入 ${n} 筆（略過 ${skip} 行）`);
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
      if (tab === "activities") renderActivityList();
    });
  });

  /** 無紀錄時：最近 7 日；有紀錄時：覆蓋資料最早～最尾一日（避免舊 CSV 跌出預設範圍） */
  function syncReportDatesFromEvents() {
    const list = sortedEvents();
    const fromEl = document.getElementById("reportFrom");
    const toEl = document.getElementById("reportTo");
    if (list.length === 0) {
      const to = new Date();
      const from = new Date();
      from.setDate(from.getDate() - 6);
      toEl.value = to.toISOString().slice(0, 10);
      fromEl.value = from.toISOString().slice(0, 10);
      return;
    }
    fromEl.value = list[0].start.slice(0, 10);
    toEl.value = list[list.length - 1].start.slice(0, 10);
  }

  function initLogDatetimeInputs() {
    setDatetimeLocalNow(document.getElementById("quickStart"));
    setDatetimeLocalNow(document.getElementById("manualStart"));
  }

  refreshActivityDatalist();
  fillMergeSelects();
  renderActivityList();
  renderTimeline();
  syncReportDatesFromEvents();
  renderReport();
  initLogDatetimeInputs();

  if (window.matchMedia("(display-mode: standalone)").matches) {
    document.getElementById("installHint").classList.remove("visible");
  } else if (/iPhone|iPad|iPod/i.test(navigator.userAgent)) {
    document.getElementById("installHint").classList.add("visible");
  }

  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    navigator.serviceWorker.register("sw.js").catch(() => {});
  }
})();
