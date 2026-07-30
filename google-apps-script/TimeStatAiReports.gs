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

function aiWorkLoadTier_(workMs) {
  if (workMs < AI_WORK_VACATION_MS) return "vacation";
  if (workMs <= AI_WORK_IDEAL_MAX_MS) return "idealFocus";
  if (workMs < AI_WORK_OVERLOAD_MAX_MS) return "overload";
  return "critical";
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
  var tierCounts = { vacation: 0, idealFocus: 0, overload: 0, critical: 0 };
  var reviewCounts = { ideal: 0, lack: 0, excessive: 0 };
  var weekBuckets = {}; // weekKey -> weekly performance accumulator
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
    var dayYmd = aiYmdLocal_(dayCursor);
    var wMs = 0;
    var tMs = 0;
    var rMs = 0;
    var trMs = 0;
    var soMs = 0;
    var reasonBits = [];
    var hasSocial = false;
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
        hasSocial = true;
        reasonBits.push(aiEventReasonBits_(list[j], actName));
      }
    }
    if (hasSocial) socialDaySet[dayYmd] = true;

    var wTier = aiWorkLoadTier_(wMs);
    tierCounts[wTier]++;
    var dayRow = {
      ymd: dayYmd,
      weekday: aiWeekdayName_(dayCursor),
      workHours: aiHours_(wMs),
      tier: wTier,
    };
    if (wTier === "critical") {
      dayRow.warning = "Critical：Work ≥6h，嚴重影響隔日表現";
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
        tiers: { vacation: 0, idealFocus: 0, overload: 0, critical: 0 },
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

  // Rhythm & Interleaving：連續 Work 塊 + Diffused Mode 間隔
  // 注意：未打卡 Sleeping 時，上一個活動會「吞」去下一打卡前嘅空窗（尤其隔夜）→ 唔好當成長時間專注
  var timeline = [];
  var suspectedUnloggedBreaks = [];
  for (var ti = 0; ti < list.length; ti++) {
    var tev = list[ti];
    var t0b = aiParseStartMs_(tev.start);
    if (isNaN(t0b) || t0b < range.fromMs || t0b > range.toMs) continue;
    var tSeg = aiSegmentMs_(list, ti, nowMs);
    var tName = aiActivityName_(state, tev.activityId) || "(unknown)";
    var tKey = aiNormKey_(tName);
    var tGroup = String(tev.group || tev.category || "").trim();
    var tEnd = t0b + tSeg;
    var crossesWake = aiCrossesWakeBoundary_(t0b, tEnd);
    var crossesMidnight = aiCrossesLocalMidnight_(t0b, tEnd);
    var overlapsSleep = aiSegmentOverlapsSleepBand_(t0b, tEnd);
    var longOpen = tSeg > AI_OPEN_SEGMENT_SUSPECT_MS;
    var impliedBreak = false;
    var lockMs = tSeg;
    var suspectReason = "";
    if (tGroup === "Work" && (crossesWake || crossesMidnight || overlapsSleep || longOpen)) {
      impliedBreak = true;
      if (crossesWake) {
        suspectReason = "跨越 wake(03:00) 且中間無打卡 → 疑似睡眠（唔計認知鎖死）";
      } else if (crossesMidnight || overlapsSleep) {
        suspectReason = "踏入／跨越睡眠帶（22:00–06:59）或跨午夜 → 疑似睡眠（唔計認知鎖死）";
      } else {
        suspectReason = "單一開放時段 >2 小時 → 疑似未打卡休息／睡眠（唔計認知鎖死）";
      }
      suspectedUnloggedBreaks.push({
        ymd: aiYmdLocal_(t0b),
        activity: tName,
        startLocal: aiLocalDateTimeStr_(t0b),
        rawEndLocal: aiLocalDateTimeStr_(tEnd),
        rawMinutes: aiMinutes_(tSeg),
        reason: suspectReason,
      });
    }
    timeline.push({
      startMs: t0b,
      endMs: tEnd,
      ms: tSeg,
      lockMs: lockMs,
      name: tName,
      key: tKey,
      isWork: tGroup === "Work",
      isDiffused: aiIsDiffusedModeActivity_(tKey, tev),
      countsForLock: tGroup === "Work" && !impliedBreak,
      impliedBreak: impliedBreak,
    });
  }

  var cognitiveLocks = [];
  var switchFails = [];
  var idx = 0;
  while (idx < timeline.length) {
    if (!timeline[idx].countsForLock) {
      idx++;
      continue;
    }
    var blockStart = idx;
    var blockMs = 0;
    var blockNames = {};
    while (idx < timeline.length && timeline[idx].countsForLock) {
      blockMs += timeline[idx].lockMs;
      blockNames[timeline[idx].name] = 1;
      idx++;
    }
    var blockEndIdx = idx - 1;
    var blockYmd = aiYmdLocal_(timeline[blockStart].startMs);
    // 必須被「真實打卡」嘅非 Work／Diffused Mode 結束；開放結尾或疑似睡眠斷點唔算認知鎖死
    var closer = idx < timeline.length ? timeline[idx] : null;
    var closedByLoggedBreak =
      closer &&
      !closer.impliedBreak &&
      !closer.countsForLock &&
      (closer.isDiffused || !closer.isWork);
    if (blockMs > AI_COGNITIVE_LOCK_MS && closedByLoggedBreak) {
      var nameList = [];
      for (var bn in blockNames) {
        if (Object.prototype.hasOwnProperty.call(blockNames, bn)) nameList.push(bn);
      }
      var lockEndMs = timeline[blockStart].startMs + blockMs;
      cognitiveLocks.push({
        ymd: blockYmd,
        startLocal: aiLocalDateTimeStr_(timeline[blockStart].startMs),
        endLocal: aiLocalDateTimeStr_(lockEndMs),
        durationMinutes: aiMinutes_(blockMs),
        activities: nameList,
        closedBy: closer.name,
        flag: "認知鎖死",
      });
    }
    // 下一段可信 Work 之前嘅間隔
    var nextWork = -1;
    for (var k = idx; k < timeline.length; k++) {
      if (timeline[k].countsForLock) {
        nextWork = k;
        break;
      }
    }
    if (nextWork < 0) break;
    var diffusedMs = 0;
    var betweenActs = [];
    var hasImpliedBreak = false;
    for (var m = idx; m < nextWork; m++) {
      if (timeline[m].isDiffused || timeline[m].impliedBreak) {
        diffusedMs += timeline[m].ms;
        if (timeline[m].impliedBreak) hasImpliedBreak = true;
      }
      betweenActs.push({
        name: timeline[m].impliedBreak ? "(疑似未打卡休息／睡眠)" : timeline[m].name,
        minutes: aiMinutes_(timeline[m].ms),
        isDiffused: timeline[m].isDiffused || timeline[m].impliedBreak,
        impliedBreak: !!timeline[m].impliedBreak,
      });
    }
    if (!hasImpliedBreak && diffusedMs < AI_DIFFUSED_GAP_MIN_MS) {
      switchFails.push({
        ymd: aiYmdLocal_(timeline[nextWork].startMs),
        afterWorkEndLocal: aiLocalDateTimeStr_(timeline[blockEndIdx].endMs),
        beforeNextWorkStartLocal: aiLocalDateTimeStr_(timeline[nextWork].startMs),
        diffusedMinutesBetween: aiMinutes_(diffusedMs),
        requiredDiffusedMinutes: 15,
        betweenActivities: betweenActs.slice(0, 10),
        flag: "切換失靈",
      });
    }
    idx = nextWork;
  }

  // Social Battery：按 ISO 週統計有 Social 活動嘅天數（目標 ≤3）
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
        "放假日(Vacation): Work<2h；理想專注日(Ideal Focus): 2–4h；超負荷預警日(Overload): 4h<Work<6h；臨界崩潰日(Critical): Work≥6h",
      reviewingAudit: "理想Review: 15–30m；缺乏Review: <15m；過度Review: >30m",
      rhythmInterleaving:
        "認知鎖死: 連續可信Work>60m 且由真實休息／非Work打卡結束；22:00–06:59／跨午夜／跨wake／開放>2h → 疑似睡眠唔計。切換失靈: 兩段可信Work之間 Diffused Mode <15m",
      boundaryFlags:
        "高頻交易練習日: Trading相關>2h；No-Trades(Mon–Fri): Transporting或Social/Family/Friend其中一個>2h",
      socialBattery: "每週有Social活動天數目標≤3（計天唔計時）",
    },
    summary: {
      workLoadTiers: {
        vacationDays: tierCounts.vacation,
        idealFocusDays: tierCounts.idealFocus,
        overloadDays: tierCounts.overload,
        criticalDays: tierCounts.critical,
      },
      reviewingAudit: {
        idealDays: reviewCounts.ideal,
        lackDays: reviewCounts.lack,
        excessiveDays: reviewCounts.excessive,
      },
      rhythmInterleaving: {
        cognitiveLockCount: cognitiveLocks.length,
        switchFailCount: switchFails.length,
        suspectedUnloggedBreakCount: suspectedUnloggedBreaks.length,
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
        vacation: tierCounts.vacation,
        idealFocus: tierCounts.idealFocus,
        overload: tierCounts.overload,
        critical: tierCounts.critical,
      },
      criticalDays: workLoadDays.filter(function (d) {
        return d.tier === "critical";
      }),
      overloadDays: workLoadDays.filter(function (d) {
        return d.tier === "overload";
      }),
      nonIdealDays: workLoadDays
        .filter(function (d) {
          return d.tier !== "idealFocus";
        })
        .slice(0, 90),
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
      cognitiveLockCount: cognitiveLocks.length,
      switchFailCount: switchFails.length,
      suspectedUnloggedBreakCount: suspectedUnloggedBreaks.length,
      cognitiveLocks: cognitiveLocks.slice(0, 40),
      switchFails: switchFails.slice(0, 40),
      suspectedUnloggedBreaks: suspectedUnloggedBreaks.slice(0, 40),
    },
    boundaryFlags: {
      highTradingPracticeDays: highTradingDays,
      noTradesBanner: noTradesDays,
    },
    socialBattery: {
      targetMaxSocialDaysPerWeek: 3,
      totalSocialDaysInPeriod: totalSocialDays,
      byWeek: socialWeekRows,
    },
  };

  var pType = range.periodType;
  var reportLens =
    pType === "week"
      ? "week_true_focus_remarks"
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
    reportLens: reportLens,
    reportLensNote:
      pType === "week"
        ? "週報重點：trueFocus（活動時間−distraction）＋ remarks 內容／pattern／值得留意之處。processAudits 只作輔助，唔好變成日數盤點。"
        : pType === "month"
          ? "月報重點：processAudits 各規則出現咗幾多日（日數盤點）。Critical／No-Trades／認知鎖死要點名日期。"
          : pType === "quarter"
            ? "季報重點：weeklyPerformance 每週表現對比（趨勢／起伏）。日數細節次要。"
            : "跟大綱；可用 weeklyPerformance 同 processAudits.summary。",
    termGlossary: {
      放假日_Vacation: "每日 Work Group < 2 小時",
      理想專注日_IdealFocus: "每日 Work Group 介乎 2–4 小時（含 2 同 4）",
      超負荷預警日_Overload: "每日 Work Group > 4 且 < 6 小時",
      臨界崩潰日_Critical: "每日 Work Group ≥ 6 小時（嚴重影響隔日表現）",
      理想Review: "當日 Reviewing 15–30 分鐘（高效總結）",
      缺乏Review: "當日 Reviewing < 15 分鐘（可能遺漏系統修正）",
      過度Review: "當日 Reviewing > 30 分鐘（判定為無效重複／反芻風險）",
      認知鎖死: "連續可信 Work >60 分鐘，且由真實打卡嘅休息／非 Work 結束；踏入 22:00–06:59、跨午夜／wake、或開放空窗 >2h（疑似未打卡睡眠）一律唔計",
      切換失靈: "兩段可信 Work 之間，Diffused Mode 活動合計 < 15 分鐘（Meditating／Walking／Resting／Gyming／Showering／Fooding；小說類 Reading、一般 Friending 可計入；成長類 Reading 唔計）",
      trueFocus: "真正專注時間 = max(0, 活動時長 − distraction)",
      高頻交易練習日: "當日 Trading 相關活動合計 > 2 小時",
      NoTradesBanner: "週一至週五：當日 Transporting 或 Social／Family／Friend 其中一個 > 2 小時 → 交易禁令提示",
      SocialBattery: "每週有 Social 活動嘅天數；目標 ≤ 3 天（計天唔計時）",
      DiffusedMode:
        "唔使高度用腦／專注嘅恢復緩衝（前稱 DMN）。判定原則：需唔需要集中精神。Reading：小說類（如 Harry Potter）=是；成長／用腦類（如原子習慣、Trading in the Zone）=否。Friending：一般社交=是；深談／會議=否。",
      運動日:
        "gyming／hiking／yogaing／running 等合計 ≥ 30 分鐘；Photoing 暫時唔計，除非 remark 注明高強度／重裝／長途拍攝等",
    },
    range: { from: range.fromYmd, to: range.toYmd },
    totals: {
      loggedHours: aiHours_(totalMs),
      distractionHours: aiHours_(distractMs),
      trueFocusHours: aiHours_(trueFocusMs),
      trueFocusWorkHours: aiHours_(trueFocusWorkMs),
    },
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
    "對齊價值：過程質素、樣本、期望值、少／小／慢；僅在數據支持時點出鬆懈／資訊過載／唔跟 checklist 跡象。",
    "術語必須帶定義：每次首次使用專有術語（如理想專注日、放假日、超負荷預警日、臨界崩潰日、理想／缺乏／過度 Review、認知鎖死、切換失靈、trueFocus、No-Trades、Diffused Mode 等），必須緊接括號或一句簡短定義；定義只可用 DATA_JSON.termGlossary／processAudits.definitions，唔好自創門檻。",
    "日期／週期顯示：唔好用 W29／ISO week 編號；用月-日範圍（例如 07-14 → 07-20）或 DATA_JSON 入面嘅 weekLabel／periodLabel／range。",
    "輸出純 Markdown。",
  ].join("\n");
}

function aiDefaultReportOutline_() {
  return [
    "按 DATA_JSON.reportLens 選擇寫法（唔好用錯期別模板）：",
    "",
    "【month_day_count_audits｜月報】重點係「有幾多日」出現各類日子：",
    "1. 執行摘要（用日數講）",
    "2. 週期對比",
    "3. Work Load Tiers 日數（Vacation／Ideal／Overload／Critical；Critical 點名日期）",
    "4. Reviewing Audit 日數（Ideal／Lack／Excessive）",
    "5. Rhythm／Diffused Mode（認知鎖死／切換失靈次數；疑似未打卡休息只作附註）",
    "6. Boundary Flags 日數（高頻 Trading、No-Trades 日期＋原因）",
    "7. Social Battery（每週社交天數）",
    "8. 下期 3 個實驗",
    "",
    "【week_true_focus_remarks｜週報】重點係真正專注同 remark，唔好做成日數盤點報告：",
    "1. 執行摘要（本週 trueFocus／distraction）",
    "2. 真正專注時間（trueFocus = 活動時間 − distraction；按活動 top；Work 淨專注）",
    "3. Remarks 內容回顧：寫咗啲咩、重複 pattern、情緒／決策／流程線索",
    "4. 值得留意嘅信號（只基於 remarks + trueFocus／distraction；可輕提 1–2 個 audit 紅旗）",
    "5. 下週 3 個細實驗",
    "",
    "【quarter_weekly_performance｜季報】重點係每週表現：",
    "1. 執行摘要（整季趨勢）",
    "2. 每週表現表（用 weeklyPerformance：trueFocusWork、Work 時數、distraction、tier 日數、No-Trades 等）",
    "3. 哪幾週最好／最差、可能原因（只可用數據）",
    "4. 季內主題回顧（短）",
    "5. 下季 3 個實驗",
    "",
    "年報：主題回顧 + 可用 weeklyPerformance／月度日數摘要。",
    "術語：首次出現必須附定義（見 termGlossary）；可用「理想專注日（Work Group 每日 2–4 小時）」呢種寫法。",
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
        cfg.reportOutline = j.reportOutline.trim();
      }
      if (typeof j.extraInstructions === "string") {
        cfg.extraInstructions = j.extraInstructions.trim();
      }
      var t = Number(j.temperature);
      if (isFinite(t) && t >= 0 && t <= 1) cfg.temperature = t;
      if (j.periodConfig && typeof j.periodConfig === "object") {
        var defPc = aiDefaultPeriodConfig_();
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
          // allow extra keys from client
          if (src.sections) {
            for (var ek in src.sections) {
              if (Object.prototype.hasOwnProperty.call(src.sections, ek) && mergedSec[ek] === undefined) {
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
  var display =
    stats.periodLabel || aiPeriodDisplayLabel_(type, stats.periodKey) || stats.periodKey;
  var title =
    type === "year"
      ? display + " 年 Time Stat 專業報告"
      : type === "quarter"
        ? display + " Time Stat 專業報告"
        : type === "month"
          ? display + " 月 Time Stat 專業報告"
          : type === "week"
            ? display + " 週 Time Stat 專業報告"
            : "Time Stat 報告 " + stats.range.from + "～" + stats.range.to;
  var pcfg = (cfg.periodConfig && cfg.periodConfig[type]) || {
    sections: aiDefaultPeriodSections_(type),
    notes: "",
  };
  var outline = aiPeriodOutlineFromSections_(type, pcfg.sections || {}, pcfg.notes || "");
  var lensNote = stats.reportLensNote ? "\n" + stats.reportLensNote : "";
  var cmpNote =
    stats.comparisons && stats.comparisons.length
      ? "\nDATA_JSON.comparisons 已附上 2 個同類型週期（合共連續 3 期對比）。"
      : "";
  var glossaryNote =
    "\n專有術語首次出現必須附定義（DATA_JSON.termGlossary）。負面情緒必引用 negativeRemarks[].context72h。";
  var kpiNote =
    "\nDATA_JSON.kpis.passFail 同 kpis.targets 係合格門檻；enabledSections 決定寫邊啲章節。";
  return (
    "請根據以下 JSON 撰寫「" +
    title +
    "」。\n\n報告大綱：\n" +
    outline +
    lensNote +
    cmpNote +
    glossaryNote +
    kpiNote +
    "\n請嚴格遵守 system 溝通方式。\n\nDATA_JSON:\n" +
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
  enrichStatsWithPeriodKpis_(state, stats);
  attachAiComparisons_(state, stats);
  var gem = callGeminiForAiReport_(stats);
  var subject = "[Time Stat AI] " + stats.periodType + " " + (stats.periodLabel || stats.periodKey);
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
