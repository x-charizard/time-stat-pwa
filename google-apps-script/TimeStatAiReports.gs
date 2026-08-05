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
/** 優先：普通 Flash；忙碌／唔支援 → 3.5 Flash → Flash-Lite（唔再用 Pro／已停用 2.5） */
var GEMINI_MODEL_PRIMARY = "gemini-3.6-flash";
var GEMINI_MODEL_FLASH_FALLBACK = "gemini-3.5-flash";
var GEMINI_MODEL_FREE_LITE = "gemini-3.5-flash-lite";
var GEMINI_MODEL = GEMINI_MODEL_PRIMARY;
var GEMINI_TRANSIENT_RETRIES = 2;
var GEMINI_TRANSIENT_SLEEP_MS = 1500;
/** Lite 熔斷（epoch ms）：只喺 Lite 自己撞 429／RPD 時開；報告開始會清舊誤判 */
var GEMINI_LITE_CIRCUIT_PROP = "GEMINI_LITE_CIRCUIT_UNTIL";
var GEMINI_LITE_CIRCUIT_MS = 30 * 60000;
/** CacheService 上限 21600s（6h）；同一 remark 唔重複燒 semantic token */
var GEMINI_SEMANTIC_CACHE_TTL_SEC = 21600;
var AI_WAKE_H = 3;
var AI_WAKE_MI = 0;
var AI_WORK_CAP_MS = 4 * 3600000;
var AI_TRADE_CAP_MS = 2 * 3600000;
var AI_REVIEW_ALERT_MS = 30 * 60000;
var AI_NO_TRADES_MS = 2 * 3600000;
/** 社交活躍日：當日 Friending＋Familying＋Socialing 合計 > 2h */
var AI_SOCIAL_ACTIVE_MS = AI_NO_TRADES_MS;
var AI_TRADING_KEYS = { trading: 1, "trading practice": 1, "trading planning": 1 };
var AI_REVIEW_KEYS = { reviewing: 1 };
var AI_TRANSPORT_KEYS = { transporting: 1 };
var AI_SOCIAL_KEYS = { friending: 1, familying: 1, socialing: 1 };
/**
 * Diffused Mode（前稱 DMN）：唔使高度用腦／專注嘅恢復緩衝。
 * Reading／Friending 要睇 remark（見 aiIsDiffusedModeActivity_）。
 */
var AI_DIFFUSED_BASE_KEYS = {
  meditating: 1,
  walking: 1,
  resting: 1,
  gyming: 1,
  showering: 1,
  fooding: 1,
  sleeping: 1,
  running: 1,
  yogaing: 1,
  hiking: 1,
  camping: 1,
  exercise: 1,
  workouting: 1,
};
/** @deprecated 用 AI_DIFFUSED_BASE_KEYS；保留別名以免舊碼漏改 */
var AI_DMN_KEYS = AI_DIFFUSED_BASE_KEYS;
var AI_WORK_VACATION_MS = 2 * 3600000;
var AI_WORK_IDEAL_MAX_MS = 4 * 3600000;
var AI_WORK_OVERLOAD_MAX_MS = 6 * 3600000;
var AI_REVIEW_IDEAL_MIN_MS = 15 * 60000;
var AI_REVIEW_IDEAL_MAX_MS = 30 * 60000;
var AI_COGNITIVE_LOCK_MS = 60 * 60000;
var AI_DIFFUSED_GAP_MIN_MS = 15 * 60000;
var AI_DMN_GAP_MIN_MS = AI_DIFFUSED_GAP_MIN_MS;
/** 單一活動開放時段（至下一打卡）>2h → 疑似未打卡休息／睡眠，唔當認知鎖死 */
var AI_OPEN_SEGMENT_SUSPECT_MS = 2 * 3600000;

function aiEventTextBlob_(ev) {
  if (!ev) return "";
  return [
    ev.remark,
    ev.activityQuestion,
    ev.achievement,
    ev.improveLast,
    ev.detailsBetter,
    ev.importantElement,
    ev.objective,
    ev.projectsFromForm,
    ev.project,
  ]
    .map(function (x) {
      return String(x || "");
    })
    .join(" ");
}

/** 小說／故事類 → Diffused Mode */
function aiReadingIsNovel_(blob) {
  var t = String(blob || "").toLowerCase();
  return /(小說|fiction|novel|harry\s*potter|哈利波特|manga|漫畫|comic|故事書|言情|科幻|fantasy|推理小說|輕小說)/i.test(
    t,
  );
}

/** 成長／用腦類 → 唔係 Diffused Mode */
function aiReadingIsFocusedStudy_(blob) {
  var t = String(blob || "").toLowerCase();
  return /(原子習慣|atomic\s*habits|trading\s*in\s*the\s*zone|deep\s*work|非小說|non[\s-]?fiction|textbook|教材|self[\s-]?help|成長|筆記書|how\s*to|心理學|策略書|投資|business|學習|流程|系統)/i.test(
    t,
  );
}

/**
 * Diffused Mode：主要睇係唔係要集中用腦。
 * Reading：小說類=是；成長／用腦類=否；唔清=否。
 * Friending：預設=是；深談／會議等=否。
 */
function aiIsDiffusedModeActivity_(tKey, ev) {
  var blob = aiEventTextBlob_(ev);
  if (tKey === "reading") {
    if (aiReadingIsNovel_(blob)) return true;
    if (aiReadingIsFocusedStudy_(blob)) return false;
    return false;
  }
  if (tKey === "friending") {
    if (
      /(深談|談判|諮詢|會議|meeting|interview|面試|輔導|工作討論|1\s*on\s*1|one[\s-]?on[\s-]?one)/i.test(
        blob,
      )
    ) {
      return false;
    }
    return true;
  }
  return !!AI_DIFFUSED_BASE_KEYS[tKey];
}

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

/** MM-DD（顯示用；唔用 W29） */
function aiMdLabelFromYmd_(ymd) {
  var s = String(ymd || "");
  return s.length >= 10 ? s.slice(5, 10) : s;
}

function aiWeekLabelFromKey_(weekKey) {
  try {
    var mon = aiMondayOfIsoWeek_(weekKey);
    var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
    return aiMdLabelFromYmd_(aiYmdLocal_(mon.getTime())) + " → " + aiMdLabelFromYmd_(aiYmdLocal_(sun.getTime()));
  } catch (e) {
    return String(weekKey || "");
  }
}

/** 報告標題／對比用顯示名（週用日期範圍） */
function aiPeriodDisplayLabel_(periodType, periodKey) {
  var type = String(periodType || "").toLowerCase();
  var key = String(periodKey || "").trim();
  if (type === "week") return aiWeekLabelFromKey_(key);
  if (type === "custom") {
    var parts = key.split("/");
    if (parts.length === 2) return aiMdLabelFromYmd_(parts[0]) + " → " + aiMdLabelFromYmd_(parts[1]);
  }
  return key;
}

/**
 * 統一 AI report／email 標題：yyyy-mm-dd Monthly|Weekly|Quarterly|Yearly Time Stat Report
 * 日期用期終日（range.toYmd）。
 */
function aiReportDocumentTitle_(periodType, range) {
  var type = String(periodType || "").toLowerCase();
  var ymd = range && range.toYmd ? String(range.toYmd) : "";
  if (!ymd && range && range.fromYmd) ymd = String(range.fromYmd);
  var kind =
    type === "week"
      ? "Weekly"
      : type === "month"
        ? "Monthly"
        : type === "quarter"
          ? "Quarterly"
          : type === "year"
            ? "Yearly"
            : "Custom";
  return (ymd ? ymd + " " : "") + kind + " Time Stat Report";
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
  // checklist 關閉 comparisons → 唔算上兩期（減 token／時間）
  if (stats.enabledSections && stats.enabledSections.comparisons === false) {
    stats.comparisons = [];
    stats.comparisonNote = "enabledSections.comparisons=false，今次唔附連續週期對比。";
    return stats;
  }
  var count = 2; // 本期 + 上 2 期 = 連續 3 個週期對比
  var comparisons = [];
  for (var i = 1; i <= count; i++) {
    var pk = aiShiftPeriodKey_(type, stats.periodKey, -i);
    if (!pk) continue;
    try {
      var s = aggregatePeriodStatsForAi_(state, type, pk);
      enrichStatsWithPeriodKpis_(state, s);
      comparisons.push({
        periodKey: s.periodKey,
        periodLabel: s.periodLabel || aiPeriodDisplayLabel_(type, s.periodKey),
        range: s.range,
        totals: s.totals,
        trueFocus: s.trueFocus
          ? { totalHours: s.trueFocus.totalHours, workGroupHours: s.trueFocus.workGroupHours }
          : null,
        focusMetrics: s.focusMetrics,
        byGroup: s.byGroup,
        wakeDayFlags: s.wakeDayFlags,
        processAuditsSummary: s.processAudits ? s.processAudits.summary : null,
        kpisSummary: s.kpis ? s.kpis.summary : null,
        passFail: s.kpis ? s.kpis.passFail : null,
        weeklyPerformance: s.weeklyPerformance || null,
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
  // 同一 id 只留最後一筆
  var seenId = {};
  var byId = [];
  for (var i = events.length - 1; i >= 0; i--) {
    var id = events[i] && events[i].id != null ? String(events[i].id) : "";
    if (id) {
      if (seenId[id]) continue;
      seenId[id] = 1;
    }
    byId.push(events[i]);
  }
  byId.reverse();
  // 匯入指紋去重（同時間／活動／地點／備註等）— 避免重覆入帳令 loggedHours 爆燈
  if (typeof eventImportDedupeKeyGs_ === "function") {
    var seenK = {};
    var out = [];
    for (var j = byId.length - 1; j >= 0; j--) {
      var k = eventImportDedupeKeyGs_(byId[j]);
      if (seenK[k]) continue;
      seenK[k] = 1;
      out.push(byId[j]);
    }
    out.reverse();
    return out;
  }
  return byId;
}

/** 同 start timestamp 嘅連續筆 → 攤分時長（對齊 PWA segmentDurationMsForReport） */
function aiSameStartRunBounds_(list, i) {
  var t0 = aiParseStartMs_(list[i].start);
  var lo = i;
  while (lo > 0 && aiParseStartMs_(list[lo - 1].start) === t0) lo--;
  var hi = i;
  while (hi + 1 < list.length && aiParseStartMs_(list[hi + 1].start) === t0) hi++;
  return { lo: lo, hi: hi, t0: t0 };
}

function aiSegmentMs_(list, i, nowMs) {
  var bounds = aiSameStartRunBounds_(list, i);
  var t0 = bounds.t0;
  if (isNaN(t0)) return 0;
  var runLen = bounds.hi - bounds.lo + 1;
  var idxInRun = i - bounds.lo;
  var t1 = NaN;
  for (var j = bounds.hi + 1; j < list.length; j++) {
    var tj = aiParseStartMs_(list[j].start);
    if (!isNaN(tj) && tj > t0) {
      t1 = tj;
      break;
    }
  }
  if (isNaN(t1)) t1 = nowMs;
  var span = Math.max(0, t1 - t0);
  var base = Math.floor(span / runLen);
  var rem = span - base * runLen;
  return base + (idxInRun < rem ? 1 : 0);
}

function aiDaysInRangeInclusive_(fromYmd, toYmd) {
  var a = new Date(String(fromYmd) + "T00:00:00").getTime();
  var b = new Date(String(toYmd) + "T00:00:00").getTime();
  if (isNaN(a) || isNaN(b) || b < a) return 1;
  return Math.round((b - a) / 86400000) + 1;
}

function aiHours_(ms) {
  return Math.round((ms / 3600000) * 10) / 10;
}

function aiMinutes_(ms) {
  return Math.round(ms / 60000);
}

function aiPeopleList_(ev) {
  var p = ev && ev.people;
  if (!p) return [];
  if (Object.prototype.toString.call(p) === "[object Array]") {
    return p
      .map(function (x) {
        return String(x || "").trim();
      })
      .filter(Boolean);
  }
  var s = String(p).trim();
  return s ? [s] : [];
}

function aiEventReasonBits_(ev, activityName) {
  var place = String((ev && ev.place) || "").trim();
  var people = aiPeopleList_(ev);
  var remark = String((ev && ev.remark) || "").trim();
  return {
    activity: activityName || "",
    place: place,
    people: people,
    remark: remark.slice(0, 160),
  };
}

/** @deprecated Work 小時分層已由 DF 日型取代；保留兼容舊 KPI（已無 critical） */
function aiWorkLoadTier_(workMs) {
  if (workMs < AI_WORK_VACATION_MS) return "vacation";
  if (workMs <= AI_WORK_IDEAL_MAX_MS) return "idealFocus";
  return "overload";
}

function aiReviewTier_(reviewMs) {
  if (reviewMs < AI_REVIEW_IDEAL_MIN_MS) return "lack";
  if (reviewMs <= AI_REVIEW_IDEAL_MAX_MS) return "ideal";
  return "excessive";
}

function aiWeekdayName_(ms) {
  var names = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  return names[new Date(ms).getDay()] || "";
}

function aiIsWeekdayMonFri_(ms) {
  var d = new Date(ms).getDay();
  return d >= 1 && d <= 5;
}

/** 下一個 wake 邊界（03:00），嚴格晚於 ms */
function aiNextWakeBoundaryAfter_(ms) {
  var d = new Date(ms);
  var wake = new Date(d.getFullYear(), d.getMonth(), d.getDate(), AI_WAKE_H, AI_WAKE_MI, 0, 0);
  if (wake.getTime() <= ms) wake.setDate(wake.getDate() + 1);
  return wake.getTime();
}

function aiCrossesWakeBoundary_(t0, t1) {
  if (!(t1 > t0)) return false;
  return aiNextWakeBoundaryAfter_(t0) < t1;
}

function aiLocalHour_(ms) {
  return new Date(ms).getHours();
}

/** 22:00–06:59：易同瞓覺重疊／未打卡 Sleeping 嘅時段 */
function aiInSleepBandHour_(h) {
  return h >= 22 || h <= 6;
}

function aiCrossesLocalMidnight_(t0, t1) {
  if (!(t1 > t0)) return false;
  var a = new Date(t0);
  var b = new Date(t1 - 1);
  return (
    a.getFullYear() !== b.getFullYear() ||
    a.getMonth() !== b.getMonth() ||
    a.getDate() !== b.getDate()
  );
}

/** 時段有冇踏入睡眠帶（抽樣） */
function aiSegmentOverlapsSleepBand_(t0, t1) {
  if (!(t1 > t0)) return false;
  if (aiInSleepBandHour_(aiLocalHour_(t0)) || aiInSleepBandHour_(aiLocalHour_(t1 - 1))) return true;
  var step = 20 * 60000;
  for (var t = t0 + step; t < t1; t += step) {
    if (aiInSleepBandHour_(aiLocalHour_(t))) return true;
    if (t - t0 > 36 * 3600000) break;
  }
  return false;
}

function aiLocalDateTimeStr_(ms) {
  var d = new Date(ms);
  var y = d.getFullYear();
  var m = ("0" + (d.getMonth() + 1)).slice(-2);
  var day = ("0" + d.getDate()).slice(-2);
  var hh = ("0" + d.getHours()).slice(-2);
  var mi = ("0" + d.getMinutes()).slice(-2);
  return y + "-" + m + "-" + day + " " + hh + ":" + mi;
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
  var trueFocusMs = 0;
  var trueFocusWorkMs = 0;
  var trueFocusByActivity = {};
  var remarksForReview = [];
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
    var dMs = dsec > 0 ? dsec * 1000 : 0;
    if (dMs > 0) {
      distractMs += dMs;
      distractByAct[name] = (distractByAct[name] || 0) + dsec;
    }
    var netMs = Math.max(0, seg - dMs);
    trueFocusMs += netMs;
    trueFocusByActivity[name] = (trueFocusByActivity[name] || 0) + netMs;
    if (g === "Work") trueFocusWorkMs += netMs;
    var rm = String(ev.remark || "").trim();
    if (rm) {
      var remarkRow = {
        ymd: aiYmdLocal_(t0),
        weekday: aiWeekdayName_(t0),
        activity: name,
        group: g,
        remark: rm.slice(0, 240),
        grossMinutes: aiMinutes_(seg),
        distractionMinutes: Math.round(dsec / 60),
        trueFocusMinutes: aiMinutes_(netMs),
      };
      remarksForReview.push(remarkRow);
      if (sampleRemarks.length < 12) sampleRemarks.push(rm.slice(0, 120));
    }
  }

  // —— Process audits（wake-day + timeline）——
  var workLoadDays = [];
  var reviewDays = [];
  var highTradingDays = [];
  var noTradesDays = [];
  var socialDaySet = {}; // ymd -> true
  var tierCounts = { vacation: 0, hea: 0, idealFocus: 0, overload: 0 };
  var reviewCounts = { ideal: 0, lack: 0, excessive: 0 };
  var weekBuckets = {}; // weekKey -> weekly performance accumulator
  var flags = {
    daysWorkOverCap: 0,
    daysTradingOverCap: 0,
    daysReviewingOver30m: 0,
    daysNoTradesBanner: 0,
  };

  var energyDays =
    typeof aiEnergyReplayDayMetrics_ === "function"
      ? aiEnergyReplayDayMetrics_(state, range.fromMs, range.toMs, 14)
      : { byYmd: {}, days: [] };
  var energyByYmd = energyDays.byYmd || {};
  var startQualityCounts = {
    goodStart: 0,
    moderateStart: 0,
    badStart: 0,
    terribleStart: 0,
  };
  var fatigueSwitchTotals = { ideal: 0, good: 0, poor: 0, high: 0, medium: 0 };
  var fatigueDropSum = 0;
  var fatigueDropEvents = 0;
  var heaDays = [];

  var dayCursor = aiWakeDayStartMs_(range.fromMs);
  var endBound = range.toMs;
  var guard = 0;
  while (dayCursor <= endBound && guard < 400) {
    guard++;
    var dayEnd = dayCursor + 86400000;
    var dayYmd = aiYmdLocal_(dayCursor);
    var wMs = 0;
    var tMs = 0;
    var rMs = 0;
    var trMs = 0;
    var soMs = 0;
    var reasonBits = [];
    for (var j = 0; j < list.length; j++) {
      var st = aiParseStartMs_(list[j].start);
      if (isNaN(st) || st < dayCursor || st >= dayEnd) continue;
      var seg2 = aiSegmentMs_(list, j, nowMs);
      var actName = aiActivityName_(state, list[j].activityId) || "(unknown)";
      var nm = aiNormKey_(actName);
      var gg = String(list[j].group || list[j].category || "").trim();
      if (gg === "Work") wMs += seg2;
      if (AI_TRADING_KEYS[nm]) tMs += seg2;
      if (AI_REVIEW_KEYS[nm]) rMs += seg2;
      if (AI_TRANSPORT_KEYS[nm]) {
        trMs += seg2;
        reasonBits.push(aiEventReasonBits_(list[j], actName));
      }
      if (AI_SOCIAL_KEYS[nm]) {
        soMs += seg2;
        reasonBits.push(aiEventReasonBits_(list[j], actName));
      }
    }
    // Social Battery：Friending＋Familying＋Socialing 合計 >2h 先計消耗一日
    var hasSocial = soMs > AI_SOCIAL_ACTIVE_MS;
    if (hasSocial) socialDaySet[dayYmd] = true;

    var eDay = energyByYmd[dayYmd] || null;
    var wTier = eDay && eDay.dayType ? eDay.dayType : aiWorkLoadTier_(wMs);
    if (!tierCounts[wTier]) tierCounts[wTier] = 0;
    tierCounts[wTier]++;
    if (eDay && eDay.startQuality && startQualityCounts[eDay.startQuality] != null) {
      startQualityCounts[eDay.startQuality]++;
    }
    if (eDay && eDay.fatigueSwitchCounts) {
      fatigueSwitchTotals.ideal += eDay.fatigueSwitchCounts.ideal || 0;
      fatigueSwitchTotals.good += eDay.fatigueSwitchCounts.good || 0;
      fatigueSwitchTotals.poor += eDay.fatigueSwitchCounts.poor || 0;
      fatigueSwitchTotals.high += eDay.fatigueSwitchCounts.high || 0;
      fatigueSwitchTotals.medium += eDay.fatigueSwitchCounts.medium || 0;
    }
    if (eDay) {
      fatigueDropSum += Number(eDay.fatigueTotalDrop) || 0;
      fatigueDropEvents += Number(eDay.fatigueDropEventCount) || 0;
    }
    var dayRow = {
      ymd: dayYmd,
      weekday: aiWeekdayName_(dayCursor),
      workHours: aiHours_(wMs),
      tier: wTier,
      wakeDf: eDay ? eDay.wakeDf : null,
      endDf: eDay ? eDay.endDf : null,
      dfDrained: eDay ? eDay.dfDrained : null,
      startQuality: eDay ? eDay.startQuality : null,
      fatigueSwitchCounts: eDay ? eDay.fatigueSwitchCounts : null,
      fatigueTotalDrop: eDay ? eDay.fatigueTotalDrop : 0,
    };
    if (wTier === "overload") {
      dayRow.warning = "Overload：結束 DF < 0 或當日 DF 扣減 > 1000";
    } else if (wTier === "hea") {
      heaDays.push(dayRow);
    }
    workLoadDays.push(dayRow);

    var rTier = aiReviewTier_(rMs);
    reviewCounts[rTier]++;
    reviewDays.push({
      ymd: dayYmd,
      weekday: aiWeekdayName_(dayCursor),
      reviewingMinutes: aiMinutes_(rMs),
      tier: rTier,
    });

    if (tMs > AI_TRADE_CAP_MS) {
      highTradingDays.push({
        ymd: dayYmd,
        weekday: aiWeekdayName_(dayCursor),
        tradingHours: aiHours_(tMs),
      });
    }

    // 舊 flag：任一單項超標（兼容）
    if (wMs > AI_WORK_CAP_MS) flags.daysWorkOverCap++;
    if (tMs > AI_TRADE_CAP_MS) flags.daysTradingOverCap++;
    if (rMs > AI_REVIEW_ALERT_MS) flags.daysReviewingOver30m++;
    if (trMs > AI_NO_TRADES_MS || soMs > AI_NO_TRADES_MS) flags.daysNoTradesBanner++;

    // 新規則：Mon–Fri 且 Transporting 或 Social/Family/Friend 其中一個 >2h
    if (aiIsWeekdayMonFri_(dayCursor) && (trMs > AI_NO_TRADES_MS || soMs > AI_NO_TRADES_MS)) {
      noTradesDays.push({
        ymd: dayYmd,
        weekday: aiWeekdayName_(dayCursor),
        transportingHours: aiHours_(trMs),
        socialFamilyFriendHours: aiHours_(soMs),
        trigger:
          trMs > AI_NO_TRADES_MS && soMs > AI_NO_TRADES_MS
            ? "both"
            : trMs > AI_NO_TRADES_MS
              ? "transporting"
              : "social",
        reasons: reasonBits.slice(0, 12),
      });
    }

    // 週表現累積（季報／對比用）
    var wkKeyDay = aiIsoWeekKeyFromDate_(new Date(dayCursor));
    if (!weekBuckets[wkKeyDay]) {
      weekBuckets[wkKeyDay] = {
        weekKey: wkKeyDay,
        workMs: 0,
        tradingMs: 0,
        reviewingMs: 0,
        distractMs: 0,
        trueFocusWorkMs: 0,
        loggedMs: 0,
        days: 0,
        socialDays: 0,
        tiers: { vacation: 0, hea: 0, idealFocus: 0, overload: 0 },
        reviewTiers: { ideal: 0, lack: 0, excessive: 0 },
        noTradesDays: 0,
        highTradingDays: 0,
      };
    }
    var wb = weekBuckets[wkKeyDay];
    wb.days++;
    wb.workMs += wMs;
    wb.tradingMs += tMs;
    wb.reviewingMs += rMs;
    if (wb.tiers[wTier] == null) wb.tiers[wTier] = 0;
    wb.tiers[wTier]++;
    wb.reviewTiers[rTier]++;
    if (hasSocial) wb.socialDays++;
    if (tMs > AI_TRADE_CAP_MS) wb.highTradingDays++;
    if (aiIsWeekdayMonFri_(dayCursor) && (trMs > AI_NO_TRADES_MS || soMs > AI_NO_TRADES_MS)) {
      wb.noTradesDays++;
    }

    dayCursor = dayEnd;
  }

  // 補每週 trueFocus／distraction（按事件）
  for (var wi = 0; wi < list.length; wi++) {
    var wev = list[wi];
    var wt0 = aiParseStartMs_(wev.start);
    if (isNaN(wt0) || wt0 < range.fromMs || wt0 > range.toMs) continue;
    var wseg = aiSegmentMs_(list, wi, nowMs);
    var wdMs = (Number(wev.distractionSec) || 0) * 1000;
    var wnet = Math.max(0, wseg - wdMs);
    var wkk = aiIsoWeekKeyFromDate_(new Date(wt0));
    if (!weekBuckets[wkk]) continue;
    weekBuckets[wkk].loggedMs += wseg;
    weekBuckets[wkk].distractMs += wdMs;
    var wg = String(wev.group || wev.category || "").trim();
    if (wg === "Work") weekBuckets[wkk].trueFocusWorkMs += wnet;
  }

  var weeklyPerformance = [];
  for (var wbk in weekBuckets) {
    if (!Object.prototype.hasOwnProperty.call(weekBuckets, wbk)) continue;
    var row = weekBuckets[wbk];
    weeklyPerformance.push({
      weekKey: row.weekKey,
      weekLabel: aiWeekLabelFromKey_(row.weekKey),
      daysInBucket: row.days,
      workHours: aiHours_(row.workMs),
      trueFocusWorkHours: aiHours_(row.trueFocusWorkMs),
      distractionHours: aiHours_(row.distractMs),
      loggedHours: aiHours_(row.loggedMs),
      tradingHours: aiHours_(row.tradingMs),
      reviewingHours: aiHours_(row.reviewingMs),
      socialDays: row.socialDays,
      workLoadTierDays: row.tiers,
      reviewingTierDays: row.reviewTiers,
      noTradesDays: row.noTradesDays,
      highTradingDays: row.highTradingDays,
    });
  }
  weeklyPerformance.sort(function (a, b) {
    return a.weekKey < b.weekKey ? -1 : a.weekKey > b.weekKey ? 1 : 0;
  });

  // 認知節奏：Focused Work（High／Medium）結束後 Fatigue_Factor
  var fatigueSwitchSamples = [];
  var fatigueDropSamples = [];
  for (var fe = 0; fe < (energyDays.days || []).length; fe++) {
    var fed = energyDays.days[fe];
    if (!fed) continue;
    if (fed.highWorkEndFatigue) {
      for (var fh = 0; fh < fed.highWorkEndFatigue.length; fh++) {
        var ff = fed.highWorkEndFatigue[fh];
        var fatN = typeof ff === "number" ? ff : Number(ff && ff.fatigueFactor);
        var grade = ff && ff.grade ? ff.grade : aiEClassifyFatigueSwitch_(fatN);
        fatigueSwitchSamples.push({
          ymd: fed.ymd,
          fatigueFactor: fatN,
          grade: grade,
          workTier: (ff && ff.workTier) || "high",
          activity: (ff && ff.activity) || "",
        });
      }
    }
    if (fed.fatigueDropEvents && fed.fatigueDropEvents.length) {
      for (var fd = 0; fd < fed.fatigueDropEvents.length; fd++) {
        var drop = fed.fatigueDropEvents[fd];
        fatigueDropSamples.push({
          ymd: fed.ymd,
          kind: drop.kind || "recover",
          from: drop.from,
          to: drop.to,
          drop: drop.drop,
        });
      }
    }
  }

  // Social Battery：按 ISO 週統計社交活躍天（當日 >2h 先計 1 日；目標 ≤3）
  var socialByWeek = {};
  for (var sy in socialDaySet) {
    if (!Object.prototype.hasOwnProperty.call(socialDaySet, sy)) continue;
    var parts = sy.split("-");
    var sd = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    var wk = aiIsoWeekKeyFromDate_(sd);
    if (!socialByWeek[wk]) socialByWeek[wk] = [];
    socialByWeek[wk].push(sy);
  }
  var socialWeekRows = [];
  for (var wkKey in socialByWeek) {
    if (!Object.prototype.hasOwnProperty.call(socialByWeek, wkKey)) continue;
    var dates = socialByWeek[wkKey].slice().sort();
    socialWeekRows.push({
      weekKey: wkKey,
      weekLabel: aiWeekLabelFromKey_(wkKey),
      socialDays: dates.length,
      targetMaxDays: 3,
      overTarget: dates.length > 3,
      dates: dates,
    });
  }
  socialWeekRows.sort(function (a, b) {
    return a.weekKey < b.weekKey ? -1 : a.weekKey > b.weekKey ? 1 : 0;
  });
  var totalSocialDays = 0;
  for (var sdi in socialDaySet) {
    if (Object.prototype.hasOwnProperty.call(socialDaySet, sdi)) totalSocialDays++;
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

  var processAudits = {
    definitions: {
      workLoadTiers:
        "放假日: DF扣減<300 且結束DF>500；Hea日: 結束DF>0 且 DF扣減<700；理想專注日: 結束DF>0 且 DF扣減≥700；Overload: 結束DF<0 或 DF扣減>1000（已取消 Critical）。Start: Good=醒時DF=上限；Moderate≥700；Bad 500<wake<700；Terrible≤500",
      reviewingAudit: "理想Review: 15–30m；缺乏Review: <15m；過度Review: >30m",
      rhythmInterleaving:
        "認知節奏與 Diffused／Focused Mode 切換（強制獨立章節）：Focused=High／Medium Work 結束時 Fatigue_Factor（≤1.4 理想；1.4–1.6 良好；≥1.6 差劣）；另用 fatigueReduction 量度 Sleep／Recover 令 Fatigue 減少幾多。樣本=0 都要寫明。已取消 OCD／DMN。",
      boundaryFlags:
        "高頻交易練習日: Trading相關>2h；No-Trades(Mon–Fri): Transporting或Social/Family/Friend其中一個>2h",
      socialBattery:
        "每週社交活躍天數目標≤3；當日 Friending＋Familying＋Socialing 合計>2h 先計消耗 1 日（少過 2h 唔扣當日 Social Battery）",
    },
    summary: {
      workLoadTiers: {
        vacationDays: tierCounts.vacation || 0,
        heaDays: tierCounts.hea || 0,
        idealFocusDays: tierCounts.idealFocus || 0,
        overloadDays: tierCounts.overload || 0,
        goodStartDays: startQualityCounts.goodStart,
        moderateStartDays: startQualityCounts.moderateStart,
        badStartDays: startQualityCounts.badStart,
        terribleStartDays: startQualityCounts.terribleStart,
      },
      reviewingAudit: {
        idealDays: reviewCounts.ideal,
        lackDays: reviewCounts.lack,
        excessiveDays: reviewCounts.excessive,
      },
      rhythmInterleaving: {
        fatigueSwitchIdeal: fatigueSwitchTotals.ideal,
        fatigueSwitchGood: fatigueSwitchTotals.good,
        fatigueSwitchPoor: fatigueSwitchTotals.poor,
        fatigueTotalDropSum: Math.round(fatigueDropSum * 100) / 100,
        fatigueDropEventCount: fatigueDropEvents,
      },
      boundaryFlags: {
        highTradingPracticeDayCount: highTradingDays.length,
        noTradesBannerDayCount: noTradesDays.length,
      },
      socialBattery: {
        totalSocialDaysInPeriod: totalSocialDays,
        weeksOverTarget: socialWeekRows.filter(function (r) {
          return r.overTarget;
        }).length,
      },
    },
    workLoadTiers: {
      counts: {
        vacation: tierCounts.vacation || 0,
        hea: tierCounts.hea || 0,
        idealFocus: tierCounts.idealFocus || 0,
        overload: tierCounts.overload || 0,
      },
      startQualityCounts: startQualityCounts,
      overloadDays: workLoadDays.filter(function (d) {
        return d.tier === "overload";
      }),
      vacationDays: workLoadDays.filter(function (d) {
        return d.tier === "vacation";
      }),
      heaDays: workLoadDays.filter(function (d) {
        return d.tier === "hea";
      }),
      idealFocusDays: workLoadDays.filter(function (d) {
        return d.tier === "idealFocus";
      }),
      days: workLoadDays.slice(0, 120),
    },
    reviewingAudit: {
      counts: {
        ideal: reviewCounts.ideal,
        lack: reviewCounts.lack,
        excessive: reviewCounts.excessive,
      },
      excessiveDays: reviewDays.filter(function (d) {
        return d.tier === "excessive";
      }),
      lackDays: reviewDays
        .filter(function (d) {
          return d.tier === "lack";
        })
        .slice(0, 90),
      nonIdealDays: reviewDays
        .filter(function (d) {
          return d.tier !== "ideal";
        })
        .slice(0, 90),
    },
    rhythmInterleaving: {
      titleMustUse: "認知節奏與 Diffused／Focused Mode 切換",
      mustWriteEvenIfZero: true,
      definition:
        "Focused Mode = High／Medium Work 段落結束時量 Fatigue_Factor；之後入 Diffused（Recover／Social／Sleep）睇 Fatigue 有冇被減少。評級：≤1.4 理想；1.4–1.6 良好；≥1.6 差劣。",
      highWorkKeys: ["trading", "trading practice", "programming", "timing", "financing", "webing", "systeming", "apping"],
      mediumWorkKeys: ["reviewing", "planning", "aiing", "photo editing", "journaling", "Xavier Li Photography", "obsidianing", "reading(非小說)", "notioning"],
      restPhotoing: "普通 Photoing = Rest；Xavier Li Photography = Work M",
      fatigueSwitchCounts: fatigueSwitchTotals,
      fatigueSwitchSamples: fatigueSwitchSamples.slice(0, 60),
      fatigueReduction: {
        totalDropSum: Math.round(fatigueDropSum * 100) / 100,
        eventCount: fatigueDropEvents,
        samples: fatigueDropSamples.slice(0, 40),
        note: "量度 Fatigue_Factor 被 Sleep reset／Recover scrub 減少幾多（drop 越大越有效）",
      },
      fatigueTotalDropSum: Math.round(fatigueDropSum * 100) / 100,
      fatigueDropEventCount: fatigueDropEvents,
      grades: {
        ideal: "Fatigue_Factor ≤ 1.4（Focused／High‧Medium Work 結束後 → Diffused 切換理想）",
        good: "1.4 < Fatigue_Factor < 1.6（切換良好）",
        poor: "Fatigue_Factor ≥ 1.6（切換差劣）",
      },
    },
    boundaryFlags: {
      highTradingPracticeDays: highTradingDays,
      noTradesBanner: noTradesDays,
    },
    socialBattery: {
      targetMaxSocialDaysPerWeek: 3,
      rule: "當日 Friending+Familying+Socialing >2h 先計 1 日消耗；≤2h 唔扣當日 Social Battery",
      totalSocialDaysInPeriod: totalSocialDays,
      byWeek: socialWeekRows,
    },
  };

  var pType = range.periodType;
  var reportLens =
    pType === "week"
      ? "week_df_energy_audits"
      : pType === "month"
        ? "month_day_count_audits"
        : pType === "quarter"
          ? "quarter_weekly_performance"
          : pType === "year"
            ? "year_themes_with_weekly_monthly_trends"
            : "custom";

  // 週報：詳細 remarks；月／季：精簡
  var remarksOut =
    pType === "week"
      ? remarksForReview.slice(0, 100)
      : remarksForReview.slice(0, 24);

  return {
    person: "Xavier",
    periodType: range.periodType,
    periodKey: range.periodKey,
    periodLabel: aiPeriodDisplayLabel_(range.periodType, range.periodKey),
    reportTitle: aiReportDocumentTitle_(range.periodType, range),
    reportLens: reportLens,
    reportLensNote:
      pType === "week"
        ? "週報必須跟 Energy／DF checklist：用 processAudits 寫放假日／Hea／理想專注／Overload、Start 品質、Social Battery（>2h 先計 1 日）。必須有獨立標題「## 認知節奏與 Diffused／Focused Mode 切換」（rhythmInterleaving：Fatigue 理想／良好／差劣＋fatigueReduction）。禁止 OCD／DMN／Critical。trueFocus／remarks 係輔助。"
        : pType === "month"
          ? "月報重點：processAudits DF 日型／Start／Social。必須有「## 認知節奏與 Diffused／Focused Mode 切換」。Overload／No-Trades 要點名日期。禁止 OCD／DMN／Critical。"
          : pType === "quarter"
            ? "季報重點：weeklyPerformance 每週表現對比（趨勢／起伏）。日數細節次要。"
            : "跟大綱；可用 weeklyPerformance 同 processAudits.summary。",
    termGlossary: {
      放假日_Vacation: "當日 DF 扣減 < 300 且結束 DF > 500",
      Hea日: "結束 DF > 0 且當日 DF 扣減 < 700",
      理想專注日_IdealFocus: "結束 DF > 0 且當日 DF 扣減 ≥ 700",
      Overload: "結束 DF < 0，或當日 DF 扣減 > 1000",
      GoodStart: "Wake up 時 DF = 上限（1000）",
      ModerateStart: "Wake up DF ≥ 700（且未達 Good）",
      BadStart: "500 < Wake up DF < 700",
      TerribleStart: "Wake up DF ≤ 500",
      理想Review: "當日 Reviewing 15–30 分鐘（高效總結）",
      缺乏Review: "當日 Reviewing < 15 分鐘（可能遺漏系統修正）",
      過度Review: "當日 Reviewing > 30 分鐘（判定為無效重複／反芻風險）",
      FatigueSwitchIdeal: "Focused（High／Medium）Work 結束後 Fatigue_Factor ≤ 1.4 → Diffused／Focused 切換理想",
      FatigueSwitchGood: "結束後 1.4 < Fatigue_Factor < 1.6 → 切換良好",
      FatigueSwitchPoor: "結束後 Fatigue_Factor ≥ 1.6 → 切換差劣",
      FatigueReduction: "Sleep reset／Recover scrub 令 Fatigue_Factor 下降嘅幅度（processAudits.rhythmInterleaving.fatigueReduction）",
      trueFocus: "真正專注時間 = max(0, 活動時長 − distraction)",
      高頻交易練習日: "當日 Trading 相關活動合計 > 2 小時",
      NoTradesBanner: "週一至週五：當日 Transporting 或 Social／Family／Friend 其中一個 > 2 小時 → 交易禁令提示",
      SocialBattery:
        "每週有「社交活躍日」嘅天數；當日 Friending＋Familying＋Socialing 合計 > 2 小時先計消耗 1 日（少過 2h 唔扣當日）；目標 ≤ 3 天",
      DiffusedMode:
        "唔使高度用腦／專注嘅恢復緩衝。切換品質改睇 High Work 後 Fatigue_Factor（已取消 OCD 鎖死／DMN 間隔檢查）。",
      運動日:
        "gyming／hiking／yogaing／running 等合計 ≥ 30 分鐘；Photoing 暫時唔計，除非 remark 注明高強度／重裝／長途拍攝等",
      loggedHours:
        "期內打卡段落合計時長（已去重 id／匯入指紋；同 timestamp 多筆會攤分）。理論上限見 totals.loggedHoursCeiling24h（日數×24）；若仍接近或超過上限，先質疑資料品質，唔好當真實活躍時數",
      loggedHoursCeiling24h: "daysInRange × 24；人唔可能全日打卡超過呢個數（除非重覆入帳／計算錯誤）",
    },
    range: { from: range.fromYmd, to: range.toYmd },
    totals: (function () {
      var daysInRange = aiDaysInRangeInclusive_(range.fromYmd, range.toYmd);
      var ceiling = daysInRange * 24;
      var loggedH = aiHours_(totalMs);
      return {
        loggedHours: loggedH,
        daysInRange: daysInRange,
        loggedHoursCeiling24h: ceiling,
        loggedHoursOverCeiling: loggedH > ceiling,
        distractionHours: aiHours_(distractMs),
        trueFocusHours: aiHours_(trueFocusMs),
        trueFocusWorkHours: aiHours_(trueFocusWorkMs),
      };
    })(),
    trueFocus: {
      formula: "trueFocusMinutes = max(0, activityMinutes - distractionMinutes)",
      totalHours: aiHours_(trueFocusMs),
      workGroupHours: aiHours_(trueFocusWorkMs),
      byActivityTop: topMap(trueFocusByActivity, 15, true),
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
      noTradesIfTransportOrSocialOverHoursMonFri: 2,
      workHardBlockAfterLocalHour: 17,
      wakeTime: "03:00",
    },
    wakeDayFlags: flags,
    processAudits: processAudits,
    weeklyPerformance: weeklyPerformance,
    remarksForReview: remarksOut,
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
    "若 totals.loggedHoursOverCeiling 為 true，或 loggedHours ≥ loggedHoursCeiling24h：必須指出可能係重覆入帳／計算問題，唔好把超上限 loggedHours 當真實活躍時數。",
    "對齊價值：過程質素、樣本、期望值、少／小／慢；僅在數據支持時點出鬆懈／資訊過載／唔跟 checklist 跡象。",
    "術語必須帶定義：每次首次使用專有術語（如理想專注日、放假日、超負荷預警日、理想／缺乏／過度 Review、Fatigue 切換、trueFocus、No-Trades、Diffused Mode 等；唔好再寫 OCD 鎖死、DMN 間隔檢查或 Critical 日），必須緊接括號或一句簡短定義；定義只可用 DATA_JSON.termGlossary／processAudits.definitions，唔好自創門檻。",
    "日期／週期顯示：唔好用 W29／ISO week 編號；用月-日範圍（例如 07-14 → 07-20）或 DATA_JSON 入面嘅 weekLabel／periodLabel／range。",
    "輸出純 Markdown。",
  ].join("\n");
}

function aiDefaultReportOutline_() {
  return [
    "按 DATA_JSON.reportLens 選擇寫法（唔好用錯期別模板）：",
    "",
    "【week_df_energy_audits｜週報】必須跟 DF／Energy checklist（定義只可用 termGlossary／processAudits.definitions）：",
    "1. 執行摘要（本週 DF 日型＋Start＋Fatigue 切換一句過）",
    "2. DF 日型日數：放假日（DF扣減<300 且結束DF>500）／Hea（結束DF>0 且扣減<700）／理想專注（結束DF>0 且扣減≥700）／Overload（結束DF<0 或扣減>1000）；點名 Overload 日期（已取消 Critical）",
    "3. Start 品質：Good（醒時DF=上限）／Moderate（≥700）／Bad（500<wake<700）／Terrible（≤500）",
    "4. ## 認知節奏與 Diffused／Focused Mode 切換（強制；標題必須用呢句）：寫 processAudits.rhythmInterleaving——Focused（High／Medium Work）結束 Fatigue 次數（理想≤1.4／良好／差劣≥1.6）；引用 fatigueSwitchSamples；用 fatigueReduction 講 Fatigue 被減少幾多；樣本=0 都要寫「本週無 Focused Work 結束樣本」",
    "5. Social Battery：Friending＋Familying＋Socialing >2h 先計消耗 1 日（≤2h 唔扣）",
    "6. trueFocus／remarks（輔助，唔好蓋過上面）",
    "7. 下週 3 個細實驗",
    "",
    "【month_day_count_audits｜月報】重點係「有幾多日」出現各類日子：",
    "1. 執行摘要（用日數講）",
    "2. 週期對比",
    "3. Work Load Tiers 日數（Vacation／Hea／Ideal／Overload＋Start 品質；Overload 點名日期；已取消 Critical）",
    "4. Reviewing Audit 日數（Ideal／Lack／Excessive）",
    "5. ## 認知節奏與 Diffused／Focused Mode 切換（強制獨立章節）：Fatigue 切換理想／良好／差劣次數＋fatigueReduction；禁止 OCD／DMN",
    "6. Boundary Flags 日數（高頻 Trading、No-Trades 日期＋原因）",
    "7. Social Battery（>2h 先計 1 日；每週社交天數）",
    "8. 下期 3 個實驗",
    "",
    "【quarter_weekly_performance｜季報】重點係每週表現：",
    "1. 執行摘要（整季趨勢）",
    "2. 每週表現表（用 weeklyPerformance：trueFocusWork、Work 時數、distraction、tier 日數、No-Trades 等）",
    "3. 哪幾週最好／最差、可能原因（只可用數據）",
    "4. 季內主題回顧（短）",
    "5. 下季 3 個實驗",
    "",
    "年報：主題回顧 + 可用 weeklyPerformance／月度日數摘要。",
    "術語：首次出現必須附定義（見 termGlossary）；可用「理想專注日（結束 DF>0 且 DF 扣減≥700）」呢種寫法。",
    "只能用 DATA_JSON；表格用 GFM；粗體用 **文字**。",
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
    periodConfig: aiDefaultPeriodConfig_(),
    emotionKeywords: aiDefaultEmotionKeywords_(),
  };
  if (!raw) return cfg;
  try {
    var j = JSON.parse(raw);
    if (j && typeof j === "object") {
      if (typeof j.systemInstruction === "string" && j.systemInstruction.trim()) {
        cfg.systemInstruction = j.systemInstruction.trim();
      }
      if (typeof j.reportOutline === "string" && j.reportOutline.trim()) {
        var savedOutline = j.reportOutline.trim();
        // 舊週報模板會叫模型「唔好做日數盤點」→ 蓋過 DF checklist；自動換成新預設
        if (/week_true_focus_remarks|唔好做成日數盤點/i.test(savedOutline)) {
          cfg.reportOutline = aiDefaultReportOutline_();
        } else {
          cfg.reportOutline = savedOutline;
        }
      }
      if (typeof j.extraInstructions === "string") {
        cfg.extraInstructions = j.extraInstructions.trim();
      }
      var t = Number(j.temperature);
      if (isFinite(t) && t >= 0 && t <= 1) cfg.temperature = t;
      if (j.periodConfig && typeof j.periodConfig === "object") {
        var defPc = aiDefaultPeriodConfig_();
        var deprecatedSec = {
          cognitiveLock: 1,
          switchFail: 1,
          switchFailDays: 1,
          workOver4h: 1,
          workOver6h: 1,
        };
        ["week", "month", "quarter", "year"].forEach(function (pt) {
          var src = j.periodConfig[pt] || {};
          var defSec = (defPc[pt] && defPc[pt].sections) || {};
          var mergedSec = {};
          for (var sk in defSec) {
            if (Object.prototype.hasOwnProperty.call(defSec, sk)) {
              mergedSec[sk] = src.sections && src.sections[sk] === false ? false : true;
              if (src.sections && typeof src.sections[sk] === "boolean") {
                mergedSec[sk] = src.sections[sk];
              }
            }
          }
          // allow extra keys from client（但丟棄已廢棄嘅 OCD／DMN 章節）
          if (src.sections) {
            for (var ek in src.sections) {
              if (
                Object.prototype.hasOwnProperty.call(src.sections, ek) &&
                mergedSec[ek] === undefined &&
                !deprecatedSec[ek]
              ) {
                mergedSec[ek] = !!src.sections[ek];
              }
            }
          }
          cfg.periodConfig[pt] = {
            sections: mergedSec,
            notes: typeof src.notes === "string" ? src.notes : "",
          };
        });
      }
      if (j.emotionKeywords && typeof j.emotionKeywords === "object") {
        if (Object.prototype.toString.call(j.emotionKeywords.negative) === "[object Array]") {
          cfg.emotionKeywords.negative = j.emotionKeywords.negative.map(String);
        }
        if (Object.prototype.toString.call(j.emotionKeywords.positive) === "[object Array]") {
          cfg.emotionKeywords.positive = j.emotionKeywords.positive.map(String);
        }
      }
    }
  } catch (e) {}
  return cfg;
}

function saveAiPromptConfig_(cfg) {
  var cur = getAiPromptConfig_();
  var next = {
    systemInstruction: String((cfg && cfg.systemInstruction) || cur.systemInstruction || aiDefaultSystemInstruction_()).trim(),
    reportOutline: String((cfg && cfg.reportOutline) || cur.reportOutline || aiDefaultReportOutline_()).trim(),
    extraInstructions: String(
      cfg && cfg.extraInstructions != null ? cfg.extraInstructions : cur.extraInstructions || ""
    ).trim(),
    temperature: 0.3,
    periodConfig: (cfg && cfg.periodConfig) || cur.periodConfig || aiDefaultPeriodConfig_(),
    emotionKeywords: (cfg && cfg.emotionKeywords) || cur.emotionKeywords || aiDefaultEmotionKeywords_(),
  };
  var t = Number(cfg && cfg.temperature);
  if (isFinite(t) && t >= 0 && t <= 1) next.temperature = t;
  // re-merge periodConfig through getter logic by round-trip
  PropertiesService.getScriptProperties().setProperty("AI_REPORT_PROMPT_CONFIG", JSON.stringify(next));
  return getAiPromptConfig_();
}

function aiSystemInstruction_() {
  var cfg = getAiPromptConfig_();
  var parts = [cfg.systemInstruction];
  if (cfg.extraInstructions) parts.push("Topic of the Period／額外：\n" + cfg.extraInstructions);
  return parts.join("\n\n");
}

function aiUserPrompt_(stats) {
  var cfg = getAiPromptConfig_();
  var type = String(stats.periodType || "custom").toLowerCase();
  var title =
    stats.reportTitle ||
    aiReportDocumentTitle_(type, stats.range || {}) ||
    "Time Stat Report";
  var pcfg = (cfg.periodConfig && cfg.periodConfig[type]) || {
    sections: aiDefaultPeriodSections_(type),
    notes: "",
  };
  var outline = aiPeriodOutlineFromSections_(type, pcfg.sections || {}, pcfg.notes || "");
  var lens = String(stats.reportLens || "");
  var globalOutline = String(cfg.reportOutline || aiDefaultReportOutline_() || "");
  // 若 Settings 仍存舊週報模板，強制用預設（避免 week_true_focus 蓋過 DF checklist）
  if (/week_true_focus_remarks|唔好做成日數盤點/i.test(globalOutline)) {
    globalOutline = aiDefaultReportOutline_();
  }
  var lensBlock = "";
  if (globalOutline) {
    var marker =
      type === "week"
        ? "【week_df_energy_audits"
        : type === "month"
          ? "【month_day_count_audits"
          : type === "quarter"
            ? "【quarter_weekly_performance"
            : "";
    if (marker && globalOutline.indexOf(marker) >= 0) {
      lensBlock = "\n\n期別模板（必須跟）：\n" + globalOutline;
    } else {
      lensBlock = "\n\n全期別大綱參考：\n" + globalOutline;
    }
  }
  var lensNote = stats.reportLensNote ? "\n" + stats.reportLensNote : "";
  var cmpNote =
    stats.comparisons && stats.comparisons.length
      ? "\nDATA_JSON.comparisons 已附上 2 個同類型週期（合共連續 3 期對比）。"
      : "";
  var glossaryNote =
    "\n專有術語首次出現必須附定義（DATA_JSON.termGlossary）。負面情緒必引用 negativeRemarks[].context72h。";
  var kpiNote =
    "\nDATA_JSON.kpis.passFail 同 kpis.targets 係合格門檻；enabledSections 決定寫邊啲章節。週／月報必須引用 processAudits.summary（workLoadTiers／rhythmInterleaving／socialBattery）。";
  var energyNote =
    type === "week" || type === "month"
      ? "\nEnergy checklist（強制）：放假日／Hea／理想專注／Overload、Good–Terrible Start、Social>2h 先計日。\n必須有獨立 Markdown 標題「## 認知節奏與 Diffused／Focused Mode 切換」：用 DATA_JSON.processAudits.rhythmInterleaving（fatigueSwitchCounts／fatigueSwitchSamples／fatigueReduction／grades）；寫理想≤1.4、良好、差劣≥1.6；量度 Fatigue 被減少幾多；counts 全 0 都要寫明原因。禁止 OCD／DMN／Critical。"
      : "";
  var requiredNote =
    type === "week" || type === "month"
      ? "\nrequiredSections=" +
        JSON.stringify([
          "DF日型",
          "Start品質",
          "認知節奏與 Diffused／Focused Mode 切換",
          "Social Battery",
        ])
      : "";
  return (
    "請根據以下 JSON 撰寫「" +
    title +
    "」（reportLens=" +
    lens +
    "）。\n\n報告大綱（章節）：\n" +
    outline +
    lensBlock +
    lensNote +
    energyNote +
    requiredNote +
    cmpNote +
    glossaryNote +
    kpiNote +
    "\n請嚴格遵守 system 溝通方式。\n\nDATA_JSON:\n" +
    JSON.stringify(stats)
  );
}

/**
 * 模型鏈：3.6 Flash → 3.5 Flash → Flash-Lite（唔用 Pro；唔再用已停用 2.5）
 */
function aiGeminiModelChain_() {
  return [
    { model: GEMINI_MODEL_PRIMARY, thinkingLevel: "medium", tier: "flash" },
    { model: GEMINI_MODEL_FLASH_FALLBACK, thinkingLevel: "medium", tier: "flash-fallback" },
    { model: GEMINI_MODEL_FREE_LITE, thinkingLevel: "minimal", tier: "free-lite" },
  ];
}

/** 模型下架／唔存在（404）— 換下一檔，唔好當無額度 */
function aiIsGeminiModelMissing_(code, text) {
  var c = Number(code) || 0;
  var t = String(text || "").toLowerCase();
  if (c === 404) return true;
  return /no longer available|not found|is not found|not supported for|unknown model|invalid model/i.test(
    t,
  );
}

/** 可換下一檔嘅「唔可用」（含硬額度；唔再把純 404 當 quota 訊息） */
function aiIsGeminiQuotaOrUnavailable_(code, text) {
  if (aiIsGeminiModelMissing_(code, text)) return false;
  var c = Number(code) || 0;
  var t = String(text || "").toLowerCase();
  if (c === 429) return true;
  if (c === 403) return true;
  return /resource_exhausted|quota|rate.?limit|rpd|requests per day|exceeded your current quota|permission.?denied|billing|free.?tier|limit:\s*0/i.test(
    t,
  );
}

/** 真正 RPM／RPD／quota */
function aiIsGeminiHardQuota_(code, text) {
  if (aiIsGeminiModelMissing_(code, text)) return false;
  var c = Number(code) || 0;
  var t = String(text || "").toLowerCase();
  if (c === 429) return true;
  return /resource_exhausted|quota|rate.?limit|rpd|requests per day|exceeded your current quota|limit:\s*0/i.test(
    t,
  );
}

/** 503／高負載暫時不可用（應重試或換模型，唔好當硬額度熔斷） */
function aiIsGeminiTransient_(code, text) {
  if (aiIsGeminiModelMissing_(code, text)) return false;
  var c = Number(code) || 0;
  var t = String(text || "").toLowerCase();
  if (c === 503 || c === 502 || c === 504) return true;
  return /high demand|try again later|unavailable|temporarily|overloaded|deadline exceeded/i.test(t);
}

function aiGeminiLiteCircuitUntil_() {
  var props = PropertiesService.getScriptProperties();
  return Number(props.getProperty(GEMINI_LITE_CIRCUIT_PROP) || 0) || 0;
}

function aiGeminiLiteCircuitOpen_() {
  return Date.now() < aiGeminiLiteCircuitUntil_();
}

function aiGeminiTripLiteCircuit_(ms) {
  var until = Date.now() + (ms != null ? ms : GEMINI_LITE_CIRCUIT_MS);
  PropertiesService.getScriptProperties().setProperty(GEMINI_LITE_CIRCUIT_PROP, String(until));
  return until;
}

function aiGeminiClearLiteCircuit_() {
  PropertiesService.getScriptProperties().deleteProperty(GEMINI_LITE_CIRCUIT_PROP);
}

/**
 * 單次呼叫：預設只打 generateContent（減 RPM／避免 3.6 雙重打 Interactions）。
 * opts.useInteractions=true 先試 Interactions。
 * 503／high demand 會短重試。
 * @returns {{ok:boolean, markdown?:string, via?:string, code?:number, body?:string, quota?:boolean, hardQuota?:boolean, transient?:boolean}}
 */
function callGeminiOnce_(apiKey, model, thinkingLevel, temperature, system, user, opts) {
  opts = opts || {};
  var think = String(thinkingLevel || "high").toLowerCase();

  if (opts.useInteractions) {
    var urlI = "https://generativelanguage.googleapis.com/v1beta/interactions";
    var bodyI = {
      model: model,
      system_instruction: system,
      input: user,
      generation_config: {
        temperature: temperature,
        thinking_level: think,
      },
    };
    try {
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
        if (out) return { ok: true, markdown: String(out), via: "interactions", code: codeI };
      }
      if (aiIsGeminiHardQuota_(codeI, textI)) {
        return {
          ok: false,
          quota: true,
          hardQuota: true,
          code: codeI,
          body: String(textI).slice(0, 400),
        };
      }
      if (aiIsGeminiQuotaOrUnavailable_(codeI, textI) && !aiIsGeminiTransient_(codeI, textI)) {
        return { ok: false, quota: true, code: codeI, body: String(textI).slice(0, 400) };
      }
    } catch (eI) {}
  }

  var urlG =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    model +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);
  var bodyG = {
    systemInstruction: { parts: [{ text: system }] },
    contents: [{ role: "user", parts: [{ text: user }] }],
    generationConfig: {
      temperature: temperature,
      thinkingConfig: { thinkingLevel: think.toUpperCase() },
    },
  };

  var maxAttempts = 1 + (Number(GEMINI_TRANSIENT_RETRIES) || 0);
  var lastCode = 0;
  var lastBody = "";
  for (var attempt = 0; attempt < maxAttempts; attempt++) {
    if (attempt > 0) {
      try {
        Utilities.sleep(GEMINI_TRANSIENT_SLEEP_MS * attempt);
      } catch (eSleep) {}
    }
    var resG = UrlFetchApp.fetch(urlG, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(bodyG),
      muteHttpExceptions: true,
    });
    var codeG = resG.getResponseCode();
    var textG = resG.getContentText();
    lastCode = codeG;
    lastBody = String(textG).slice(0, 400);
    if (codeG >= 200 && codeG < 300) {
      var jG = JSON.parse(textG);
      var outG = extractGeminiTextFallback_(jG);
      if (outG) return { ok: true, markdown: String(outG), via: "generateContent", code: codeG };
      return { ok: false, code: codeG, body: "empty_output" };
    }
    if (aiIsGeminiModelMissing_(codeG, textG)) {
      return {
        ok: false,
        modelMissing: true,
        code: codeG,
        body: lastBody,
      };
    }
    if (aiIsGeminiHardQuota_(codeG, textG)) {
      return {
        ok: false,
        quota: true,
        hardQuota: true,
        code: codeG,
        body: lastBody,
      };
    }
    if (aiIsGeminiTransient_(codeG, textG)) {
      if (attempt < maxAttempts - 1) continue;
      return {
        ok: false,
        transient: true,
        code: codeG,
        body: lastBody,
      };
    }
    break;
  }
  return {
    ok: false,
    modelMissing: aiIsGeminiModelMissing_(lastCode, lastBody),
    quota: aiIsGeminiQuotaOrUnavailable_(lastCode, lastBody),
    hardQuota: aiIsGeminiHardQuota_(lastCode, lastBody),
    transient: aiIsGeminiTransient_(lastCode, lastBody),
    code: lastCode,
    body: lastBody,
  };
}

function aiSemanticCacheKey_(activity, remark) {
  var raw = String(activity || "").trim().toLowerCase() + "\n" + String(remark || "").trim().toLowerCase();
  var dig = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var b64 = Utilities.base64EncodeWebSafe(dig);
  return "esem_" + String(b64).replace(/=+$/g, "").slice(0, 80);
}

/**
 * Flash 優先；503／模型下架 → 換下一檔。
 * 每次報告清舊 Lite 熔斷（避免誤判 404 鎖死）；只喺 Lite 自己 429 先重新熔斷 30 分鐘。
 */
function callGeminiWithMessages_(system, user) {
  var props = PropertiesService.getScriptProperties();
  var apiKey = String(props.getProperty("GEMINI_API_KEY") || "").trim();
  if (!apiKey) throw new Error("missing_GEMINI_API_KEY");

  // 清走舊熔斷（先前把 404／免費額當硬限會鎖死已付費 key）
  aiGeminiClearLiteCircuit_();

  var cfg = getAiPromptConfig_();
  var temperature = cfg.temperature;
  var chain = aiGeminiModelChain_();
  var last = null;
  var attempted = [];

  for (var i = 0; i < chain.length; i++) {
    var step = chain[i];
    if (step.tier === "free-lite" && aiGeminiLiteCircuitOpen_()) {
      attempted.push(step.model + "@circuit-open");
      break;
    }
    attempted.push(step.model + "@" + step.thinkingLevel);
    var r = callGeminiOnce_(apiKey, step.model, step.thinkingLevel, temperature, system, user, {});
    last = r;
    if (r && r.ok && r.markdown) {
      return {
        markdown: r.markdown,
        model: step.model,
        thinkingLevel: step.thinkingLevel,
        tier: step.tier,
        via: r.via,
        fallback: i > 0,
        attempted: attempted,
      };
    }
    // 只熔斷「Lite 自己」嘅硬額度；Flash 429 仍試下一檔（付費 key 唔應鎖死整條鏈）
    if (r && r.hardQuota && step.tier === "free-lite") {
      aiGeminiTripLiteCircuit_();
    }
    // 模型下架／503／quota／4xx／5xx → 試下一檔
    if (!(r && (r.modelMissing || r.quota || r.transient || r.hardQuota || r.code >= 400))) break;
  }

  var detail = last && last.body ? String(last.body).slice(0, 240) : "unknown";
  var code = last && last.code ? last.code : 0;
  if (last && last.modelMissing) {
    throw new Error(
      "gemini_model_unavailable: 模型鏈全部唔可用（可能下架）。attempted=" +
        attempted.join(" → ") +
        " http_" +
        code +
        ":" +
        detail,
    );
  }
  if (last && (last.quota || last.hardQuota)) {
    throw new Error(
      "gemini_quota_exhausted: Flash／Flash-Lite 撞 RPM／RPD（付費項目請喺 AI Studio 確認 Billing 綁定同一 API key）。attempted=" +
        attempted.join(" → ") +
        " http_" +
        code +
        ":" +
        detail,
    );
  }
  throw new Error("gemini_http_" + code + ":" + detail + " attempted=" + attempted.join(" → "));
}

function callGeminiForAiReport_(stats) {
  return callGeminiWithMessages_(aiSystemInstruction_(), aiUserPrompt_(stats));
}

/**
 * 人手報告追問：帶原報告 + DATA_JSON + 對話歷史答問題。
 */
function aiFollowUpSystem_() {
  var base = aiSystemInstruction_();
  return [
    base,
    "",
    "你而家係「報告追問」模式：用戶已有一份 Time Stat AI 報告，會再問跟進問題。",
    "必須用 DATA_JSON 入面嘅數字／日期答；報告正文只作語境，唔好同 DATA_JSON 矛盾。",
    "若 DATA_JSON 無足夠證據，清楚講「數據唔夠／未提供」，唔好虛構。",
    "輸出純 Markdown；精簡、可直接執行嘅洞察優先。",
  ].join("\n");
}

function aiFollowUpUserPrompt_(stats, reportMarkdown, history, question) {
  var histLines = [];
  var hist = history || [];
  for (var i = 0; i < hist.length; i++) {
    var turn = hist[i] || {};
    var role = String(turn.role || "") === "assistant" ? "助理" : "用戶";
    var content = String(turn.content || "").trim();
    if (!content) continue;
    histLines.push(role + "：\n" + content.slice(0, 6000));
  }
  var histBlock = histLines.length ? "\n\n對話歷史：\n" + histLines.join("\n\n") : "";
  var reportSlice = String(reportMarkdown || "").slice(0, 60000);
  return (
    "期別：" +
    String(stats.periodType || "") +
    " · " +
    String(stats.periodLabel || stats.periodKey || "") +
    "（range " +
    ((stats.range && stats.range.from) || "") +
    "～" +
    ((stats.range && stats.range.to) || "") +
    "）\n\n" +
    "原報告（Markdown）：\n" +
    reportSlice +
    histBlock +
    "\n\n用戶新問題：\n" +
    String(question || "").trim() +
    "\n\n請用繁體中文答。必要時引用 DATA_JSON 具體數字／日期。\n\nDATA_JSON:\n" +
    JSON.stringify(stats)
  );
}

function callGeminiForAiFollowUp_(stats, reportMarkdown, history, question) {
  return callGeminiWithMessages_(
    aiFollowUpSystem_(),
    aiFollowUpUserPrompt_(stats, reportMarkdown, history, question),
  );
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

function aiEscapeHtmlEmail_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function aiMdInlineEmail_(text) {
  var s = aiEscapeHtmlEmail_(text);
  s = s.replace(/`([^`\n]+)`/g, "<code style=\"font-family:ui-monospace,Menlo,Consolas,monospace;font-size:0.9em;background:#f4f4f5;padding:1px 4px;border-radius:3px;\">$1</code>");
  s = s.replace(/\*\*([^*\n]+)\*\*/g, "<strong>$1</strong>");
  s = s.replace(/__([^_\n]+)__/g, "<strong>$1</strong>");
  s = s.replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, "$1<em>$2</em>");
  return s;
}

function aiSplitMdTableRowEmail_(line) {
  var s = String(line || "").trim();
  if (s.charAt(0) === "|") s = s.slice(1);
  if (s.charAt(s.length - 1) === "|") s = s.slice(0, -1);
  return s.split("|").map(function (c) {
    return c.trim();
  });
}

function aiIsMdTableSepEmail_(line) {
  return /^\s*\|?[\s|:/-]+\|[\s|:|/-]*$/.test(String(line || ""));
}

/**
 * Lightweight Markdown → email-safe HTML（標題／列表／粗體／GFM table）
 */
function aiMarkdownToEmailHtml_(md) {
  var raw = String(md || "").replace(/\r\n/g, "\n");
  var lines = raw.split("\n");
  var html = [];
  var i = 0;
  var inUl = false;
  var inOl = false;

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

  while (i < lines.length) {
    var line = lines[i];
    if (/^\s*\|/.test(line) && i + 1 < lines.length && aiIsMdTableSepEmail_(lines[i + 1])) {
      closeLists();
      var rows = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        rows.push(lines[i]);
        i++;
      }
      html.push(
        '<table style="border-collapse:collapse;width:100%;max-width:100%;font-size:14px;margin:12px 0;" cellpadding="0" cellspacing="0">',
      );
      html.push("<thead><tr>");
      aiSplitMdTableRowEmail_(rows[0]).forEach(function (h) {
        html.push(
          '<th style="border:1px solid #ccc;padding:6px 10px;text-align:left;background:#f4f4f5;">' +
            aiMdInlineEmail_(h) +
            "</th>",
        );
      });
      html.push("</tr></thead><tbody>");
      for (var r = 1; r < rows.length; r++) {
        if (aiIsMdTableSepEmail_(rows[r])) continue;
        html.push("<tr>");
        aiSplitMdTableRowEmail_(rows[r]).forEach(function (c) {
          html.push(
            '<td style="border:1px solid #ccc;padding:6px 10px;text-align:left;vertical-align:top;">' +
              aiMdInlineEmail_(c) +
              "</td>",
          );
        });
        html.push("</tr>");
      }
      html.push("</tbody></table>");
      continue;
    }

    var hm = line.match(/^(#{1,3})\s+(.+)$/);
    if (hm) {
      closeLists();
      var level = hm[1].length;
      var sizes = { 1: "22px", 2: "18px", 3: "16px" };
      html.push(
        "<h" +
          level +
          ' style="margin:1.1em 0 0.4em;line-height:1.3;font-size:' +
          sizes[level] +
          ';">' +
          aiMdInlineEmail_(hm[2]) +
          "</h" +
          level +
          ">",
      );
      i++;
      continue;
    }

    if (/^(-{3,}|\*{3,}|_{3,})\s*$/.test(line)) {
      closeLists();
      html.push('<hr style="border:0;border-top:1px solid #ddd;margin:1em 0;">');
      i++;
      continue;
    }

    var ulm = line.match(/^\s*[-*+]\s+(.+)$/);
    if (ulm) {
      if (inOl) {
        html.push("</ol>");
        inOl = false;
      }
      if (!inUl) {
        html.push('<ul style="margin:0.4em 0 0.8em;padding-left:1.4em;">');
        inUl = true;
      }
      html.push("<li style=\"margin:0.2em 0;\">" + aiMdInlineEmail_(ulm[1]) + "</li>");
      i++;
      continue;
    }

    var olm = line.match(/^\s*\d+\.\s+(.+)$/);
    if (olm) {
      if (inUl) {
        html.push("</ul>");
        inUl = false;
      }
      if (!inOl) {
        html.push('<ol style="margin:0.4em 0 0.8em;padding-left:1.4em;">');
        inOl = true;
      }
      html.push("<li style=\"margin:0.2em 0;\">" + aiMdInlineEmail_(olm[1]) + "</li>");
      i++;
      continue;
    }

    if (/^\s*$/.test(line)) {
      closeLists();
      i++;
      continue;
    }

    closeLists();
    var paras = [];
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
    html.push('<p style="margin:0.55em 0;line-height:1.55;">' + aiMdInlineEmail_(paras.join(" ")) + "</p>");
  }
  closeLists();

  return (
    '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif;font-size:15px;color:#111;line-height:1.55;max-width:720px;">' +
    (html.join("\n") || "<p>（空白報告）</p>") +
    "</div>"
  );
}

function mailAiReportToAllowed_(subject, markdown) {
  var emails = allowedEmailsList_();
  if (!emails || !emails.length) throw new Error("no_allowed_emails");
  var body = String(markdown || "");
  var htmlBody = aiMarkdownToEmailHtml_(body);
  for (var i = 0; i < emails.length; i++) {
    MailApp.sendEmail({
      to: emails[i],
      subject: subject,
      body: body,
      htmlBody: htmlBody,
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
  enrichStatsWithPeriodKpis_(state, stats);
  attachAiComparisons_(state, stats);
  var gem = callGeminiForAiReport_(stats);
  var subject =
    stats.reportTitle ||
    aiReportDocumentTitle_(stats.periodType, stats.range || {}) ||
    "[Time Stat AI] " + stats.periodType + " " + (stats.periodLabel || stats.periodKey);
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
      model: gem.model + "/" + (gem.thinkingLevel || "") + "/" + gem.via + (gem.fallback ? "/fallback" : ""),
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
    thinkingLevel: gem.thinkingLevel || "",
    tier: gem.tier || "",
    fallback: !!gem.fallback,
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
  if (typeof installEmotionTriggerQueue === "function") {
    installEmotionTriggerQueue();
  }
  Logger.log(
    "AI report triggers installed (daily 03:10/20/30 HKT month/quarter/year; Saturday 07:00 HKT week; emotion queue 08:00 HKT)."
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
      thinkingLevel: r.thinkingLevel,
      tier: r.tier,
      fallback: r.fallback,
      periodType: r.periodType,
      periodKey: r.periodKey,
      persist: false,
      emailed: !!wantEmail,
    })
  );
}

/**
 * 人手報告追問：{ periodType, periodKey, question, reportMarkdown?, history?[{role,content}] }
 */
function handleAskAiReportFollowUp_(body) {
  var periodType = String(body && body.periodType ? body.periodType : "custom");
  var periodKey = String(body && body.periodKey ? body.periodKey : "");
  if (!periodKey && body && body.fromYmd && body.toYmd) {
    periodType = "custom";
    periodKey = String(body.fromYmd) + "/" + String(body.toYmd);
  }
  if (!periodKey) return authFail_("missing_period_key");
  var question = String(body && body.question ? body.question : "").trim();
  if (!question) return authFail_("missing_question");
  if (question.length > 4000) question = question.slice(0, 4000);

  var reportMarkdown = String(body && body.reportMarkdown ? body.reportMarkdown : "");
  var historyIn = body && body.history && Object.prototype.toString.call(body.history) === "[object Array]"
    ? body.history
    : [];
  var history = [];
  for (var i = 0; i < historyIn.length && history.length < 8; i++) {
    var t = historyIn[i] || {};
    var role = String(t.role || "") === "assistant" ? "assistant" : "user";
    var content = String(t.content || "").trim();
    if (!content) continue;
    history.push({ role: role, content: content.slice(0, 8000) });
  }

  var state = readStateFromSheet_();
  if (!state || !state.events) throw new Error("no_state");
  var stats = aggregatePeriodStatsForAi_(state, periodType, periodKey);
  enrichStatsWithPeriodKpis_(state, stats);
  attachAiComparisons_(state, stats);

  var gem = callGeminiForAiFollowUp_(stats, reportMarkdown, history, question);
  return jsonOut_(
    authOkFields_({
      markdown: gem.markdown,
      model: gem.model,
      via: gem.via,
      thinkingLevel: gem.thinkingLevel,
      tier: gem.tier,
      fallback: gem.fallback,
      periodType: stats.periodType,
      periodKey: stats.periodKey,
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
        periodConfig: aiDefaultPeriodConfig_(),
        emotionKeywords: aiDefaultEmotionKeywords_(),
      },
      sectionLabels: aiPeriodSectionLabels_(),
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
    periodConfig: s.periodConfig,
    emotionKeywords: s.emotionKeywords,
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


/**
 * Energy Model 5.0 — Semantic Lite（只用 flash-lite generateContent；跳過 Interactions）
 * body: { activity, remark }
 * → { score, sleep_base, is_fragmented, model, tier }
 */
function handleAnalyzeEnergySemantic_(body) {
  var activity = String((body && body.activity) || "").trim();
  var remark = String((body && body.remark) || "").trim();
  if (!activity && !remark) {
    return jsonOut_(
      authOkFields_({
        score: 0,
        sleep_base: 1,
        is_fragmented: false,
        skipped: true,
      }),
    );
  }

  if (aiGeminiLiteCircuitOpen_()) {
    return jsonOut_(
      authOkFields_({
        score: 0,
        sleep_base: 1,
        is_fragmented: false,
        skipped: true,
        reason: "lite_circuit",
        circuitUntil: aiGeminiLiteCircuitUntil_(),
      }),
    );
  }

  var cacheKey = aiSemanticCacheKey_(activity, remark);
  try {
    var cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      var hit = JSON.parse(cached);
      return jsonOut_(
        authOkFields_({
          score: hit.score,
          sleep_base: hit.sleep_base,
          is_fragmented: !!hit.is_fragmented,
          model: hit.model || GEMINI_MODEL_FREE_LITE,
          tier: "free-lite",
          via: "cache",
          cached: true,
        }),
      );
    }
  } catch (eCache) {}

  var props = PropertiesService.getScriptProperties();
  var apiKey = String(props.getProperty("GEMINI_API_KEY") || "").trim();
  if (!apiKey) return authFail_("missing_GEMINI_API_KEY");

  var system =
    "You are a mental-energy semantic scorer for a personal time log. " +
    "Return ONLY compact JSON (no markdown) with keys: " +
    "score (number -2..+2), sleep_base (number 0.7..1.3), is_fragmented (boolean). " +
    "score: emotional/cognitive drain vs recovery for THIS activity+remark. " +
    "-2 extreme drain/anger; 0 neutral; +2 extreme joy/recovery. " +
    "If activity is Sleeping (or sleep): set sleep_base for quality " +
    "(disturbed/nightmare/mosquitoes/woke up → ~0.7 and is_fragmented true; " +
    "deep sleep/rested → ~1.3 and is_fragmented false; default 1.0). " +
    "For non-sleep, sleep_base=1.0 and is_fragmented=false unless remark clearly about sleep quality.";
  var user =
    "Activity: " +
    activity +
    "\nRemark: " +
    (remark || "(empty)") +
    "\nJSON:";

  var model = GEMINI_MODEL_FREE_LITE;
  var r = callGeminiOnce_(apiKey, model, "minimal", 0.2, system, user, {
    skipInteractions: true,
  });
  if (!(r && r.ok && r.markdown)) {
    var detail = r && r.body ? String(r.body).slice(0, 200) : "empty";
    var code = r && r.code ? r.code : 0;
    if (r && (r.hardQuota || r.quota)) {
      if (r.hardQuota) aiGeminiTripLiteCircuit_();
      return authFail_("gemini_quota_exhausted:lite:" + detail);
    }
    return authFail_("gemini_http_" + code + ":" + detail);
  }
  var raw = String(r.markdown || "").trim();
  var m = raw.match(/\{[\s\S]*\}/);
  if (!m) return authFail_("semantic_bad_json");
  var obj;
  try {
    obj = JSON.parse(m[0]);
  } catch (eParse) {
    return authFail_("semantic_bad_json");
  }
  var score = Number(obj.score);
  if (!isFinite(score)) score = 0;
  score = Math.max(-2, Math.min(2, score));
  var sleepBase = Number(obj.sleep_base);
  if (!isFinite(sleepBase)) sleepBase = 1;
  sleepBase = Math.max(0.7, Math.min(1.3, sleepBase));
  var frag = obj.is_fragmented === true || obj.is_fragmented === "true" || obj.is_fragmented === 1;
  var payload = {
    score: Math.round(score * 100) / 100,
    sleep_base: Math.round(sleepBase * 100) / 100,
    is_fragmented: !!frag,
    model: model,
  };
  try {
    CacheService.getScriptCache().put(
      cacheKey,
      JSON.stringify(payload),
      GEMINI_SEMANTIC_CACHE_TTL_SEC,
    );
  } catch (ePut) {}
  return jsonOut_(
    authOkFields_({
      score: payload.score,
      sleep_base: payload.sleep_base,
      is_fragmented: payload.is_fragmented,
      model: model,
      tier: "free-lite",
      via: r.via || "generateContent",
      cached: false,
    }),
  );
}
