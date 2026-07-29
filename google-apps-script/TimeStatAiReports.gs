/**
 * Time Stat — Gemini 週／月／季／年 AI 報告
 *
 * 同 TimeStatSync.gs 放喺同一個 Apps Script 專案（多個 .gs 會自動合併）。
 * Script properties 加：GEMINI_API_KEY
 *
 * 部署後跑一次：installAiReportTriggers()
 * 乾跑：testGenerateAiReportMonth()
 */

var AI_REPORTS_SHEET = "TimeStatAIReports";
var GEMINI_MODEL = "gemini-3.6-flash";
var AI_WAKE_H = 3;
var AI_WAKE_MI = 0;
var AI_WORK_CAP_MS = 4 * 3600000;
var AI_TRADE_CAP_MS = 2 * 3600000;
var AI_REVIEW_ALERT_MS = 30 * 60000;
var AI_NO_TRADES_MS = 2 * 3600000;
var AI_TRADING_KEYS = { trading: 1, "trading practice": 1, "trading planning": 1 };
var AI_REVIEW_KEYS = { reviewing: 1 };
var AI_TRANSPORT_KEYS = { transporting: 1 };
var AI_SOCIAL_KEYS = { friending: 1, familying: 1, socialing: 1 };

function aiNormKey_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function aiActivityName_(state, activityId) {
  var acts = (state && state.activities) || [];
  for (var i = 0; i < acts.length; i++) {
    if (acts[i] && acts[i].id === activityId) return String(acts[i].name || "").trim();
  }
  return "";
}

function aiParseStartMs_(iso) {
  var t = new Date(iso).getTime();
  return isNaN(t) ? NaN : t;
}

function aiYmdLocal_(ms) {
  var d = new Date(ms);
  var y = d.getFullYear();
  var m = ("0" + (d.getMonth() + 1)).slice(-2);
  var day = ("0" + d.getDate()).slice(-2);
  return y + "-" + m + "-" + day;
}

function aiWakeDayStartMs_(refMs) {
  var d = new Date(refMs);
  var wakeToday = new Date(d.getFullYear(), d.getMonth(), d.getDate(), AI_WAKE_H, AI_WAKE_MI, 0, 0);
  if (refMs < wakeToday.getTime()) {
    wakeToday.setDate(wakeToday.getDate() - 1);
  }
  return wakeToday.getTime();
}

function aiIsoWeekKeyFromDate_(d) {
  var date = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  var dayNr = (date.getDay() + 6) % 7; // Mon=0
  date.setDate(date.getDate() - dayNr + 3); // Thursday
  var weekYear = date.getFullYear();
  var week1 = new Date(weekYear, 0, 4);
  var day1 = (week1.getDay() + 6) % 7;
  week1.setDate(week1.getDate() - day1); // Monday of ISO week 1
  var weekNo = 1 + Math.round((date.getTime() - week1.getTime()) / 604800000);
  return weekYear + "-W" + ("0" + weekNo).slice(-2);
}

function aiMondayOfIsoWeek_(weekKey) {
  var m = String(weekKey || "").match(/^(\d{4})-W(\d{2})$/i);
  if (!m) throw new Error("bad_period_key_week");
  var y = parseInt(m[1], 10);
  var w = parseInt(m[2], 10);
  var week1 = new Date(y, 0, 4);
  var day1 = (week1.getDay() + 6) % 7;
  week1.setDate(week1.getDate() - day1);
  return new Date(week1.getFullYear(), week1.getMonth(), week1.getDate() + (w - 1) * 7);
}

function aiPeriodRange_(periodType, periodKey) {
  var type = String(periodType || "").toLowerCase();
  var key = String(periodKey || "").trim();
  var fromYmd = "";
  var toYmd = "";
  if (type === "month") {
    // YYYY-MM
    var ym = key.match(/^(\d{4})-(\d{2})$/);
    if (!ym) throw new Error("bad_period_key_month");
    var y = parseInt(ym[1], 10);
    var mo = parseInt(ym[2], 10) - 1;
    fromYmd = key + "-01";
    var last = new Date(y, mo + 1, 0).getDate();
    toYmd = key + "-" + ("0" + last).slice(-2);
  } else if (type === "quarter") {
    // YYYY-Qn
    var q = key.match(/^(\d{4})-Q([1-4])$/i);
    if (!q) throw new Error("bad_period_key_quarter");
    var yq = parseInt(q[1], 10);
    var qi = parseInt(q[2], 10);
    var startMo = (qi - 1) * 3;
    fromYmd = yq + "-" + ("0" + (startMo + 1)).slice(-2) + "-01";
    var endMo = startMo + 2;
    var lastQ = new Date(yq, endMo + 1, 0).getDate();
    toYmd = yq + "-" + ("0" + (endMo + 1)).slice(-2) + "-" + ("0" + lastQ).slice(-2);
  } else if (type === "year") {
    var yy = key.match(/^(\d{4})$/);
    if (!yy) throw new Error("bad_period_key_year");
    fromYmd = key + "-01-01";
    toYmd = key + "-12-31";
  } else if (type === "week") {
    // ISO week YYYY-Www (Mon–Sun)
    var mon = aiMondayOfIsoWeek_(key);
    fromYmd = aiYmdLocal_(mon.getTime());
    var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
    toYmd = aiYmdLocal_(sun.getTime());
  } else if (type === "custom") {
    var parts = key.split("/");
    if (parts.length !== 2) throw new Error("bad_period_key_custom");
    fromYmd = parts[0];
    toYmd = parts[1];
  } else {
    throw new Error("bad_period_type");
  }
  var fromMs = new Date(fromYmd + "T00:00:00").getTime();
  var toMs = new Date(toYmd + "T23:59:59.999").getTime();
  return { fromYmd: fromYmd, toYmd: toYmd, fromMs: fromMs, toMs: toMs, periodType: type, periodKey: key };
}

/** delta = -1 → 上一個同類型週期 key */
function aiShiftPeriodKey_(periodType, periodKey, delta) {
  var type = String(periodType || "").toLowerCase();
  var key = String(periodKey || "").trim();
  var dlt = Number(delta) || 0;
  if (type === "month") {
    var ym = key.match(/^(\d{4})-(\d{2})$/);
    if (!ym) return null;
    var dm = new Date(parseInt(ym[1], 10), parseInt(ym[2], 10) - 1 + dlt, 1);
    return dm.getFullYear() + "-" + ("0" + (dm.getMonth() + 1)).slice(-2);
  }
  if (type === "quarter") {
    var q = key.match(/^(\d{4})-Q([1-4])$/i);
    if (!q) return null;
    var total = parseInt(q[1], 10) * 4 + (parseInt(q[2], 10) - 1) + dlt;
    var ny = Math.floor(total / 4);
    var nq = total - ny * 4 + 1;
    return ny + "-Q" + nq;
  }
  if (type === "year") {
    var yy = key.match(/^(\d{4})$/);
    if (!yy) return null;
    return String(parseInt(yy[1], 10) + dlt);
  }
  if (type === "week") {
    try {
      var mon = aiMondayOfIsoWeek_(key);
      mon.setDate(mon.getDate() + dlt * 7);
      return aiIsoWeekKeyFromDate_(mon);
    } catch (e) {
      return null;
    }
  }
  return null;
}

function attachAiComparisons_(state, stats) {
  if (!stats) return stats;
  var type = String(stats.periodType || "").toLowerCase();
  if (type === "custom") {
    stats.comparisons = [];
    stats.comparisonNote = "自訂範圍唔附帶連續週期對比。";
    return stats;
  }
  var count = type === "week" ? 3 : 2;
  var comparisons = [];
  for (var i = 1; i <= count; i++) {
    var pk = aiShiftPeriodKey_(type, stats.periodKey, -i);
    if (!pk) continue;
    try {
      var s = aggregatePeriodStatsForAi_(state, type, pk);
      comparisons.push({
        periodKey: s.periodKey,
        range: s.range,
        totals: s.totals,
        focusMetrics: s.focusMetrics,
        byGroup: s.byGroup,
        wakeDayFlags: s.wakeDayFlags,
      });
    } catch (e) {}
  }
  stats.comparisons = comparisons;
  stats.comparisonNote =
    "comparisons[] 係同類型連續上一／幾個週期摘要。請寫「週期對比」章節（趨勢／差異），只能用呢啲數字，唔好虛構。";
  return stats;
}

function aiSortedEvents_(state) {
  var events = ((state && state.events) || []).slice();
  events.sort(function (a, b) {
    return aiParseStartMs_(a.start) - aiParseStartMs_(b.start);
  });
  return events;
}

function aiSegmentMs_(list, i, nowMs) {
  var t0 = aiParseStartMs_(list[i].start);
  if (isNaN(t0)) return 0;
  var t1 = NaN;
  for (var j = i + 1; j < list.length; j++) {
    var tj = aiParseStartMs_(list[j].start);
    if (!isNaN(tj) && tj > t0) {
      t1 = tj;
      break;
    }
  }
  if (isNaN(t1)) t1 = nowMs;
  return Math.max(0, t1 - t0);
}

function aiHours_(ms) {
  return Math.round((ms / 3600000) * 10) / 10;
}

function aggregatePeriodStatsForAi_(state, periodType, periodKey) {
  var range = aiPeriodRange_(periodType, periodKey);
  var list = aiSortedEvents_(state);
  var nowMs = Date.now();
  var byGroup = {};
  var byActivity = {};
  var focus = {
    tradingHours: 0,
    tradingPracticeHours: 0,
    tradingPlanningHours: 0,
    reviewingHours: 0,
    workGroupHours: 0,
    transportingHours: 0,
    socialFamilyFriendHours: 0,
  };
  var totalMs = 0;
  var distractMs = 0;
  var sampleRemarks = [];
  var distractByAct = {};

  for (var i = 0; i < list.length; i++) {
    var ev = list[i];
    var t0 = aiParseStartMs_(ev.start);
    if (isNaN(t0) || t0 < range.fromMs || t0 > range.toMs) continue;
    var seg = aiSegmentMs_(list, i, nowMs);
    totalMs += seg;
    var name = aiActivityName_(state, ev.activityId) || "(unknown)";
    var key = aiNormKey_(name);
    byActivity[name] = (byActivity[name] || 0) + seg;
    var g = String(ev.group || ev.category || "").trim() || "Unlabeled";
    byGroup[g] = (byGroup[g] || 0) + seg;
    if (g === "Work") focus.workGroupHours += seg;
    if (key === "trading") focus.tradingHours += seg;
    if (key === "trading practice") focus.tradingPracticeHours += seg;
    if (key === "trading planning") focus.tradingPlanningHours += seg;
    if (AI_REVIEW_KEYS[key]) focus.reviewingHours += seg;
    if (AI_TRANSPORT_KEYS[key]) focus.transportingHours += seg;
    if (AI_SOCIAL_KEYS[key]) focus.socialFamilyFriendHours += seg;
    var dsec = Number(ev.distractionSec) || 0;
    if (dsec > 0) {
      distractMs += dsec * 1000;
      distractByAct[name] = (distractByAct[name] || 0) + dsec;
    }
    var rm = String(ev.remark || "").trim();
    if (rm && sampleRemarks.length < 12) sampleRemarks.push(rm.slice(0, 120));
  }

  // wake-day flags inside calendar range
  var flags = {
    daysWorkOverCap: 0,
    daysTradingOverCap: 0,
    daysReviewingOver30m: 0,
    daysNoTradesBanner: 0,
  };
  var dayCursor = aiWakeDayStartMs_(range.fromMs);
  var endBound = range.toMs;
  var guard = 0;
  while (dayCursor <= endBound && guard < 400) {
    guard++;
    var dayEnd = dayCursor + 86400000;
    var wMs = 0;
    var tMs = 0;
    var rMs = 0;
    var trMs = 0;
    var soMs = 0;
    for (var j = 0; j < list.length; j++) {
      var st = aiParseStartMs_(list[j].start);
      if (isNaN(st) || st < dayCursor || st >= dayEnd) continue;
      var seg2 = aiSegmentMs_(list, j, nowMs);
      var nm = aiNormKey_(aiActivityName_(state, list[j].activityId));
      var gg = String(list[j].group || list[j].category || "").trim();
      if (gg === "Work") wMs += seg2;
      if (AI_TRADING_KEYS[nm]) tMs += seg2;
      if (AI_REVIEW_KEYS[nm]) rMs += seg2;
      if (AI_TRANSPORT_KEYS[nm]) trMs += seg2;
      if (AI_SOCIAL_KEYS[nm]) soMs += seg2;
    }
    if (wMs > AI_WORK_CAP_MS) flags.daysWorkOverCap++;
    if (tMs > AI_TRADE_CAP_MS) flags.daysTradingOverCap++;
    if (rMs > AI_REVIEW_ALERT_MS) flags.daysReviewingOver30m++;
    if (trMs > AI_NO_TRADES_MS || soMs > AI_NO_TRADES_MS) flags.daysNoTradesBanner++;
    dayCursor = dayEnd;
  }

  function topMap(obj, n, isMs) {
    var arr = [];
    for (var k in obj) {
      if (!Object.prototype.hasOwnProperty.call(obj, k)) continue;
      arr.push({ name: k, hours: isMs ? aiHours_(obj[k]) : obj[k] });
    }
    arr.sort(function (a, b) {
      return b.hours - a.hours;
    });
    return arr.slice(0, n || 12);
  }

  var distractTop = [];
  for (var dn in distractByAct) {
    if (!Object.prototype.hasOwnProperty.call(distractByAct, dn)) continue;
    distractTop.push({ name: dn, distractionMinutes: Math.round(distractByAct[dn] / 60) });
  }
  distractTop.sort(function (a, b) {
    return b.distractionMinutes - a.distractionMinutes;
  });

  return {
    person: "Xavier",
    periodType: range.periodType,
    periodKey: range.periodKey,
    range: { from: range.fromYmd, to: range.toYmd },
    totals: {
      loggedHours: aiHours_(totalMs),
      distractionHours: aiHours_(distractMs),
    },
    byGroup: (function () {
      var o = {};
      for (var gk in byGroup) {
        if (Object.prototype.hasOwnProperty.call(byGroup, gk)) o[gk] = aiHours_(byGroup[gk]);
      }
      return o;
    })(),
    byActivityTop: topMap(byActivity, 15, true),
    focusMetrics: {
      tradingHours: aiHours_(focus.tradingHours),
      tradingPracticeHours: aiHours_(focus.tradingPracticeHours),
      tradingPlanningHours: aiHours_(focus.tradingPlanningHours),
      reviewingHours: aiHours_(focus.reviewingHours),
      workGroupHours: aiHours_(focus.workGroupHours),
      transportingHours: aiHours_(focus.transportingHours),
      socialFamilyFriendHours: aiHours_(focus.socialFamilyFriendHours),
    },
    capContext: {
      workSoftCapHoursPerWakeDay: 4,
      tradingSoftCapHoursPerWakeDay: 2,
      reviewingAlertMinutesPerWakeDay: 30,
      noTradesIfTransportOrSocialOverHours: 2,
      workHardBlockAfterLocalHour: 17,
      wakeTime: "03:00",
    },
    wakeDayFlags: flags,
    distractionTopActivities: distractTop.slice(0, 8),
    sampleRemarks: sampleRemarks,
  };
}

function aiDefaultSystemInstruction_() {
  return [
    "你是 Xavier 的 Time Stat／交易生活節奏分析員，不是財經喊單員。",
    "使用繁體中文；可夾英文專有詞（Trading、MOC、Prop Firm）。",
    "框架：hypothesis → evidence（只能用提供的數字）→ review／下期實驗。",
    "禁止預測市場升跌；禁止虛構未提供的數據。",
    "對齊價值：過程質素、樣本、期望值、少／小／慢；僅在數據支持時點出鬆懈／資訊過載／唔跟 checklist 跡象。",
    "輸出純 Markdown。",
  ].join("\n");
}

function aiDefaultReportOutline_() {
  return [
    "報告必須包含以下章節（可用繁中標題）：",
    "1. 執行摘要",
    "2. 週期對比（對照 DATA_JSON.comparisons；用 Markdown table 列出本期 vs 上期重點指標）",
    "3. 時間配置儀表板",
    "4. 交易相關時間質素",
    "5. 恢復與干擾",
    "6. 規則遵守分數卡",
    "7. 下期 3 個可執行實驗（細、可量度）",
    "週報偏本週節奏；月報偏操作；季報加趨勢；年報加主題回顧。",
    "表格請用 GFM：| col | … | 下一行 | --- |。粗體用 **文字**。",
  ].join("\n");
}

/** 可由 PWA「AI Settings」覆寫；存 Script Property AI_REPORT_PROMPT_CONFIG */
function getAiPromptConfig_() {
  var props = PropertiesService.getScriptProperties();
  var raw = String(props.getProperty("AI_REPORT_PROMPT_CONFIG") || "").trim();
  var cfg = {
    systemInstruction: aiDefaultSystemInstruction_(),
    reportOutline: aiDefaultReportOutline_(),
    extraInstructions: "",
    temperature: 0.3,
  };
  if (!raw) return cfg;
  try {
    var j = JSON.parse(raw);
    if (j && typeof j === "object") {
      if (typeof j.systemInstruction === "string" && j.systemInstruction.trim()) {
        cfg.systemInstruction = j.systemInstruction.trim();
      }
      if (typeof j.reportOutline === "string" && j.reportOutline.trim()) {
        cfg.reportOutline = j.reportOutline.trim();
      }
      if (typeof j.extraInstructions === "string") {
        cfg.extraInstructions = j.extraInstructions.trim();
      }
      var t = Number(j.temperature);
      if (isFinite(t) && t >= 0 && t <= 1) cfg.temperature = t;
    }
  } catch (e) {}
  return cfg;
}

function saveAiPromptConfig_(cfg) {
  var next = {
    systemInstruction: String((cfg && cfg.systemInstruction) || aiDefaultSystemInstruction_()).trim(),
    reportOutline: String((cfg && cfg.reportOutline) || aiDefaultReportOutline_()).trim(),
    extraInstructions: String((cfg && cfg.extraInstructions) || "").trim(),
    temperature: 0.3,
  };
  var t = Number(cfg && cfg.temperature);
  if (isFinite(t) && t >= 0 && t <= 1) next.temperature = t;
  PropertiesService.getScriptProperties().setProperty("AI_REPORT_PROMPT_CONFIG", JSON.stringify(next));
  return next;
}

function aiSystemInstruction_() {
  var cfg = getAiPromptConfig_();
  var parts = [cfg.systemInstruction, cfg.reportOutline];
  if (cfg.extraInstructions) parts.push(cfg.extraInstructions);
  return parts.join("\n\n");
}

function aiUserPrompt_(stats) {
  var cfg = getAiPromptConfig_();
  var type = stats.periodType;
  var title =
    type === "year"
      ? stats.periodKey + " 年 Time Stat 專業報告"
      : type === "quarter"
        ? stats.periodKey + " Time Stat 專業報告"
        : type === "month"
          ? stats.periodKey + " 月 Time Stat 專業報告"
          : type === "week"
            ? stats.periodKey + " 週 Time Stat 專業報告"
            : "Time Stat 報告 " + stats.range.from + "～" + stats.range.to;
  var spanNote =
    type === "year"
      ? "期別：年報。"
      : type === "quarter"
        ? "期別：季報。"
        : type === "month"
          ? "期別：月報。"
          : type === "week"
            ? "期別：週報（ISO 週 Mon–Sun）。"
            : "期別：自訂範圍。";
  var extra = cfg.extraInstructions ? "\n\n用戶額外要求：\n" + cfg.extraInstructions : "";
  var cmpNote =
    stats.comparisons && stats.comparisons.length
      ? "\nDATA_JSON.comparisons 已附連續上期摘要，必須有「週期對比」章節（建議用 table）。"
      : "";
  return (
    "請根據以下 JSON 撰寫「" +
    title +
    "」。\n" +
    spanNote +
    cmpNote +
    "\n請嚴格遵守 system 入面嘅溝通方式同報告大綱。" +
    extra +
    "\n\nDATA_JSON:\n" +
    JSON.stringify(stats)
  );
}

/**
 * Gemini Interactions API；失敗則 fallback generateContent。
 */
function callGeminiForAiReport_(stats) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = String(props.getProperty("GEMINI_API_KEY") || "").trim();
  if (!apiKey) throw new Error("missing_GEMINI_API_KEY");

  var cfg = getAiPromptConfig_();
  var system = aiSystemInstruction_();
  var user = aiUserPrompt_(stats);
  var temperature = cfg.temperature;

  // Try Interactions API
  try {
    var urlI = "https://generativelanguage.googleapis.com/v1beta/interactions";
    var bodyI = {
      model: GEMINI_MODEL,
      system_instruction: system,
      input: user,
      generation_config: { temperature: temperature },
    };
    var resI = UrlFetchApp.fetch(urlI, {
      method: "post",
      contentType: "application/json",
      headers: { "x-goog-api-key": apiKey },
      payload: JSON.stringify(bodyI),
      muteHttpExceptions: true,
    });
    var codeI = resI.getResponseCode();
    var textI = resI.getContentText();
    if (codeI >= 200 && codeI < 300) {
      var jI = JSON.parse(textI);
      var out =
        jI.output_text ||
        (jI.output && jI.output.text) ||
        extractGeminiTextFallback_(jI);
      if (out) return { markdown: String(out), model: GEMINI_MODEL, via: "interactions" };
    }
  } catch (eI) {}

  // Fallback: generateContent
  var urlG =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    GEMINI_MODEL +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);
  var bodyG = {
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: "user", parts: [{ text: user }] }],
    generationConfig: { temperature: temperature },
  };
  var resG = UrlFetchApp.fetch(urlG, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(bodyG),
    muteHttpExceptions: true,
  });
  var codeG = resG.getResponseCode();
  var textG = resG.getContentText();
  if (codeG < 200 || codeG >= 300) {
    throw new Error("gemini_http_" + codeG + ":" + String(textG).slice(0, 300));
  }
  var jG = JSON.parse(textG);
  var outG = extractGeminiTextFallback_(jG);
  if (!outG) throw new Error("gemini_empty_output");
  return { markdown: String(outG), model: GEMINI_MODEL, via: "generateContent" };
}

function extractGeminiTextFallback_(j) {
  if (!j) return "";
  if (j.output_text) return j.output_text;
  try {
    var cands = j.candidates || [];
    if (cands[0] && cands[0].content && cands[0].content.parts) {
      var parts = cands[0].content.parts;
      var bits = [];
      for (var i = 0; i < parts.length; i++) {
        if (parts[i].text) bits.push(parts[i].text);
      }
      return bits.join("\n");
    }
  } catch (e) {}
  return "";
}

function ensureAiReportsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(AI_REPORTS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AI_REPORTS_SHEET);
    sh
      .getRange(1, 1, 1, 9)
      .setValues([
        [
          "id",
          "periodType",
          "periodKey",
          "generatedAt",
          "subject",
          "markdown",
          "model",
          "persist",
          "syncedToVault",
        ],
      ]);
  }
  return sh;
}

function appendAiReportRow_(row) {
  var sh = ensureAiReportsSheet_();
  sh.appendRow([
    row.id,
    row.periodType,
    row.periodKey,
    row.generatedAt,
    row.subject,
    row.markdown,
    row.model,
    row.persist ? "1" : "0",
    row.syncedToVault ? "1" : "0",
  ]);
}

function listAiReportsFromSheet_(limit) {
  var sh = ensureAiReportsSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var out = [];
  for (var i = values.length - 1; i >= 1; i--) {
    var r = values[i];
    if (String(r[7]) !== "1") continue; // only persist=true history
    out.push({
      id: String(r[0] || ""),
      periodType: String(r[1] || ""),
      periodKey: String(r[2] || ""),
      generatedAt: String(r[3] || ""),
      subject: String(r[4] || ""),
      model: String(r[6] || ""),
      syncedToVault: String(r[8]) === "1",
    });
    if (out.length >= (limit || 100)) break;
  }
  return out;
}

function getAiReportFromSheet_(id) {
  var want = String(id || "");
  var sh = ensureAiReportsSheet_();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === want) {
      return {
        id: String(values[i][0] || ""),
        periodType: String(values[i][1] || ""),
        periodKey: String(values[i][2] || ""),
        generatedAt: String(values[i][3] || ""),
        subject: String(values[i][4] || ""),
        markdown: String(values[i][5] || ""),
        model: String(values[i][6] || ""),
        persist: String(values[i][7]) === "1",
        syncedToVault: String(values[i][8]) === "1",
      };
    }
  }
  return null;
}

function listAiReportsForVaultSync_() {
  var sh = ensureAiReportsSheet_();
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][7]) !== "1") continue;
    out.push({
      id: String(values[i][0] || ""),
      periodType: String(values[i][1] || ""),
      periodKey: String(values[i][2] || ""),
      generatedAt: String(values[i][3] || ""),
      subject: String(values[i][4] || ""),
      markdown: String(values[i][5] || ""),
      model: String(values[i][6] || ""),
      syncedToVault: String(values[i][8]) === "1",
      rowIndex: i + 1,
    });
  }
  return out;
}

function markAiReportSynced_(id) {
  var sh = ensureAiReportsSheet_();
  var values = sh.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.getRange(i + 1, 9).setValue("1");
      return true;
    }
  }
  return false;
}

function mailAiReportToAllowed_(subject, markdown) {
  var emails = allowedEmailsList_();
  if (!emails || !emails.length) throw new Error("no_allowed_emails");
  var body = String(markdown || "");
  // Gmail plain text; keep markdown readable
  for (var i = 0; i < emails.length; i++) {
    MailApp.sendEmail({
      to: emails[i],
      subject: subject,
      body: body,
    });
  }
  return emails.slice();
}

/**
 * @param {{persist?: boolean, email?: boolean, skipDedupe?: boolean}} opts
 */
function generatePeriodAiReport_(periodType, periodKey, opts) {
  opts = opts || {};
  var persist = opts.persist === true;
  var doEmail = opts.email !== false; // default true
  var skipDedupe = opts.skipDedupe === true;

  var props = PropertiesService.getScriptProperties();
  var dedupeKey = "aiReport:" + periodType + ":" + periodKey;
  if (persist && !skipDedupe && props.getProperty(dedupeKey)) {
    return { ok: true, already: true, periodType: periodType, periodKey: periodKey };
  }

  var state = readStateFromSheet_();
  if (!state || !state.events) throw new Error("no_state");
  var stats = aggregatePeriodStatsForAi_(state, periodType, periodKey);
  attachAiComparisons_(state, stats);
  var gem = callGeminiForAiReport_(stats);
  var subject = "[Time Stat AI] " + stats.periodType + " " + stats.periodKey;
  var id = Utilities.getUuid();
  var generatedAt = new Date().toISOString();

  if (doEmail) mailAiReportToAllowed_(subject, gem.markdown);

  if (persist) {
    appendAiReportRow_({
      id: id,
      periodType: stats.periodType,
      periodKey: stats.periodKey,
      generatedAt: generatedAt,
      subject: subject,
      markdown: gem.markdown,
      model: gem.model + "/" + gem.via,
      persist: true,
      syncedToVault: false,
    });
    props.setProperty(dedupeKey, String(Date.now()));
  }

  return {
    ok: true,
    id: persist ? id : "",
    periodType: stats.periodType,
    periodKey: stats.periodKey,
    subject: subject,
    markdown: gem.markdown,
    model: gem.model,
    via: gem.via,
    persist: persist,
    emailed: doEmail,
    stats: stats,
  };
}

function previousMonthKey_() {
  var d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() - 1);
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2);
}

function previousQuarterKey_() {
  var d = new Date();
  var q = Math.floor(d.getMonth() / 3); // 0-3 current
  var y = d.getFullYear();
  if (q === 0) {
    y -= 1;
    q = 4;
  }
  return y + "-Q" + q;
}

function previousYearKey_() {
  return String(new Date().getFullYear() - 1);
}

/** 上一個已完結 ISO 週（Mon–Sun）；星期六朝早排程用 */
function previousWeekKey_() {
  var d = new Date();
  var dayNr = (d.getDay() + 6) % 7; // Mon=0
  d.setDate(d.getDate() - dayNr - 7); // Monday of previous ISO week
  return aiIsoWeekKeyFromDate_(d);
}

function runScheduledAiReportMonth_() {
  var d = new Date();
  // 每月 1 號先跑上一個曆月（其餘日子靠 dedupe 短路）
  if (d.getDate() !== 1) return;
  generatePeriodAiReport_("month", previousMonthKey_(), { persist: true, email: true });
}

function runScheduledAiReportQuarter_() {
  var d = new Date();
  // only run on first day of Jan/Apr/Jul/Oct
  if (d.getDate() !== 1) return;
  var m = d.getMonth();
  if (m !== 0 && m !== 3 && m !== 6 && m !== 9) return;
  generatePeriodAiReport_("quarter", previousQuarterKey_(), { persist: true, email: true });
}

function runScheduledAiReportYear_() {
  var d = new Date();
  if (d.getDate() !== 1 || d.getMonth() !== 0) return;
  generatePeriodAiReport_("year", previousYearKey_(), { persist: true, email: true });
}

function runScheduledAiReportWeek_() {
  // Trigger 本身已限定星期六；呢度直接出上一個 ISO 週
  generatePeriodAiReport_("week", previousWeekKey_(), { persist: true, email: true });
}

/** 部署後跑一次，安裝 Asia/Hong_Kong 時間觸發 */
function installAiReportTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (
      fn === "runScheduledAiReportMonth_" ||
      fn === "runScheduledAiReportQuarter_" ||
      fn === "runScheduledAiReportYear_" ||
      fn === "runScheduledAiReportWeek_"
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("runScheduledAiReportMonth_")
    .timeBased()
    .atHour(3)
    .nearMinute(10)
    .everyDays(1)
    .inTimezone("Asia/Hong_Kong")
    .create();
  // daily check; handlers no-op unless calendar boundary
  ScriptApp.newTrigger("runScheduledAiReportQuarter_")
    .timeBased()
    .atHour(3)
    .nearMinute(20)
    .everyDays(1)
    .inTimezone("Asia/Hong_Kong")
    .create();
  ScriptApp.newTrigger("runScheduledAiReportYear_")
    .timeBased()
    .atHour(3)
    .nearMinute(30)
    .everyDays(1)
    .inTimezone("Asia/Hong_Kong")
    .create();
  ScriptApp.newTrigger("runScheduledAiReportWeek_")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SATURDAY)
    .atHour(7)
    .nearMinute(0)
    .inTimezone("Asia/Hong_Kong")
    .create();
  Logger.log(
    "AI report triggers installed (daily 03:10/20/30 HKT month/quarter/year; Saturday 07:00 HKT week)."
  );
}

function testGenerateAiReportMonth() {
  var key = previousMonthKey_();
  var r = generatePeriodAiReport_("month", key, { persist: true, email: true, skipDedupe: true });
  Logger.log(JSON.stringify({ ok: r.ok, periodKey: r.periodKey, via: r.via, len: (r.markdown || "").length }));
}

function testGenerateAiReportWeek() {
  var key = previousWeekKey_();
  var r = generatePeriodAiReport_("week", key, { persist: true, email: true, skipDedupe: true });
  Logger.log(
    JSON.stringify({
      ok: r.ok,
      periodKey: r.periodKey,
      via: r.via,
      comparisons: (r.stats && r.stats.comparisons && r.stats.comparisons.length) || 0,
      len: (r.markdown || "").length,
    })
  );
}

function handleListAiReports_(body) {
  var lim = Number(body && body.limit != null ? body.limit : 50);
  return jsonOut_(authOkFields_({ reports: listAiReportsFromSheet_(lim) }));
}

function handleGetAiReport_(body) {
  var row = getAiReportFromSheet_(body && body.id);
  if (!row) return authFail_("not_found");
  return jsonOut_(authOkFields_({ report: row }));
}

function handleGenerateAiReport_(body) {
  var periodType = String(body && body.periodType ? body.periodType : "month");
  var periodKey = String(body && body.periodKey ? body.periodKey : "");
  if (!periodKey && body && body.fromYmd && body.toYmd) {
    periodType = "custom";
    periodKey = String(body.fromYmd) + "/" + String(body.toYmd);
  }
  if (!periodKey) return authFail_("missing_period_key");
  var wantEmail = body && body.email === true;
  // 人手：強制 persist false（唔入歷史；亦唔寫 Obsidian）
  var r = generatePeriodAiReport_(periodType, periodKey, {
    persist: false,
    email: wantEmail,
    skipDedupe: true,
  });
  return jsonOut_(
    authOkFields_({
      markdown: r.markdown,
      subject: r.subject,
      model: r.model,
      via: r.via,
      periodType: r.periodType,
      periodKey: r.periodKey,
      persist: false,
      emailed: !!wantEmail,
    })
  );
}

function handleGetAiSettings_(body) {
  var cfg = getAiPromptConfig_();
  return jsonOut_(
    authOkFields_({
      settings: cfg,
      defaults: {
        systemInstruction: aiDefaultSystemInstruction_(),
        reportOutline: aiDefaultReportOutline_(),
        extraInstructions: "",
        temperature: 0.3,
      },
    })
  );
}

function handleSaveAiSettings_(body) {
  var s = (body && body.settings) || body || {};
  var saved = saveAiPromptConfig_({
    systemInstruction: s.systemInstruction,
    reportOutline: s.reportOutline,
    extraInstructions: s.extraInstructions,
    temperature: s.temperature,
  });
  return jsonOut_(authOkFields_({ settings: saved, saved: true }));
}

function handleListAiReportsForVault_(body) {
  return jsonOut_(authOkFields_({ reports: listAiReportsForVaultSync_() }));
}

function handleMarkAiReportSynced_(body) {
  var ok = markAiReportSynced_(body && body.id);
  if (!ok) return authFail_("not_found");
  return jsonOut_(authOkFields_({ synced: true, id: String(body.id) }));
}
