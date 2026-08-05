/**
 * Time Stat AI — 週／月／季／年 KPI、periodConfig、情緒 brief
 * 同 TimeStatAiReports.gs 一齊部署。
 */

var AI_SLEEP_KEY = "sleeping";
var AI_MEDITATE_KEY = "meditating";
var AI_EXERCISE_KEYS = {
  gyming: 1,
  hiking: 1,
  yogaing: 1,
  running: 1,
  workouting: 1,
  exercise: 1,
  camping: 1,
};
var AI_SLEEP_SHORT_MS = 6 * 3600000;
var AI_SLEEP_LONG_MS = 10 * 3600000;
var AI_EXERCISE_MIN_MS = 30 * 60000;

/** Photoing 暫時唔計高強度／運動，除非 remark 另有注明 */
function aiPhotoingCountsAsExercise_(ev) {
  var blob = "";
  if (typeof aiEventTextBlob_ === "function") blob = aiEventTextBlob_(ev);
  else blob = String((ev && ev.remark) || "");
  blob = blob.toLowerCase();
  return /(高強度|高强|strenuous|intense\s*shoot|heavy\s*gear|長途拍攝|重裝|體力消耗|hiking\s*shoot)/i.test(
    blob,
  );
}

function aiIsExerciseActivity_(nm, ev) {
  if (AI_EXERCISE_KEYS[nm]) return true;
  if (nm === "photoing" || nm === "photography") return aiPhotoingCountsAsExercise_(ev);
  return false;
}

function aiDefaultEmotionKeywords_() {
  return {
    negative: [
      "chaos",
      "choas",
      "頭痛",
      "焦慮",
      "崩潰",
      "內耗",
      "panic",
      "overwhelm",
      "煩",
      "暴躁",
      "失眠",
      "抑鬱",
      "sad",
      "angry",
      "anxiety",
    ],
    positive: [
      "happy",
      "enjoy",
      "開心",
      "平靜",
      "滿足",
      "grateful",
      "感恩",
      "放鬆",
      "calm",
      "joy",
      "好心情",
      "輕鬆",
    ],
  };
}

function aiDefaultPeriodSections_(periodType) {
  var t = String(periodType || "").toLowerCase();
  if (t === "week") {
    // 順序：DF／Energy checklist 優先，trueFocus／remarks 輔助
    return {
      vacationDays: true,
      heaDays: true,
      idealFocusDays: true,
      overloadDays: true,
      startQuality: true,
      fatigueSwitch: true,
      socialDays: true,
      trueFocus: true,
      negativeRemarks72h: true,
      positiveRemarks: true,
      sleepAnomalies: true,
      meditateDays: true,
      exerciseDays: true,
      weekTheme: true,
      comparisons: true,
    };
  }
  if (t === "month") {
    return {
      fatigueSwitch: true,
      overSocialWeeks: true,
      chaosStreak: true,
      sleepAnomalies: true,
      meditateKeep: true,
      exerciseGaps: true,
      dayTypeCounts: true,
      startQuality: true,
      tradingOver2h: true,
      positiveEmotionDays: true,
      emotionUpsDowns: true,
      monthTheme: true,
      comparisons: true,
    };
  }
  if (t === "quarter") {
    return {
      weeklyPassFail: true,
      patternsWeekday: true,
      placeHabits: true,
      rhythmRegularity: true,
      severeIssues: true,
      habitChanges: true,
      emotionFactors: true,
      comparisons: true,
    };
  }
  if (t === "year") {
    return {
      monthlyWorkRest: true,
      passFailTrend: true,
      placeEmotion: true,
      changePoints: true,
      monthStories: true,
      comparisons: true,
    };
  }
  return { comparisons: true, notes: true };
}

function aiDefaultPeriodConfig_() {
  return {
    week: { sections: aiDefaultPeriodSections_("week"), notes: "" },
    month: { sections: aiDefaultPeriodSections_("month"), notes: "" },
    quarter: { sections: aiDefaultPeriodSections_("quarter"), notes: "" },
    year: { sections: aiDefaultPeriodSections_("year"), notes: "" },
  };
}

function aiPeriodSectionLabels_() {
  return {
    week: {
      fatigueSwitch:
        "Fatigue switch after High Work — ≤1.4 ideal; 1.4–1.6 good; ≥1.6 poor (also report Fatigue reduction from recover/sleep)",
      socialDays:
        "Social battery days — Friending+Familying+Socialing total >2h that day (≤2h does not consume; ideal <3 / week)",
      negativeRemarks72h:
        "Negative emotion remarks — list hits + what happened in the prior 72h",
      positiveRemarks: "Positive emotion remarks — highlight matching remarks",
      vacationDays: "Vacation days — DF drain <300 and end DF >500 (ideal 1–2 / week)",
      heaDays: "Hea days — end DF >0 and DF drain <700",
      idealFocusDays: "Ideal focus days — end DF >0 and DF drain ≥700",
      overloadDays: "Overload days — end DF <0 or DF drain >1000",
      startQuality: "Wake DF start quality — Good(=cap) / Moderate(≥700) / Bad(500<wake<700) / Terrible(≤500)",
      sleepAnomalies:
        "Sleep anomalies — Sleeping <6h or >10h when logged (max consecutive ≤2 days)",
      meditateDays: "Meditate days — Meditating >0 (ideal ≥6 / week)",
      exerciseDays:
        "Exercise days — gyming/hiking/yogaing/running etc. ≥30m (ideal 2–5 / week)",
      weekTheme: "Week theme — infer from Topic + remarks/KPIs",
      trueFocus: "True focus — activity time minus distraction",
      comparisons: "3-week comparison — current + prior 2 weeks",
    },
    month: {
      fatigueSwitch:
        "Fatigue switch after High Work — count ideal/good/poor; report how much Fatigue_Factor was reduced",
      overSocialWeeks: "Over-social weeks — socialDays (>2h) >3; must not be two weeks in a row",
      chaosStreak: "Chaos streak — negative-emotion days must not run ≥3 consecutive",
      sleepAnomalies: "Sleep anomalies — short/long sleep must not be two consecutive days",
      meditateKeep: "Meditate keep — whether Meditating days stay consistent",
      exerciseGaps: "Exercise gaps — spacing between ≥30m exercise days / recovery",
      dayTypeCounts:
        "DF day types — vacation / hea / idealFocus / overload day counts (Critical removed)",
      startQuality: "Wake DF start quality — Good / Moderate / Bad / Terrible day counts",
      tradingOver2h: "Trading >2h days — must not be two consecutive days",
      positiveEmotionDays: "Positive emotion days — days with positive keyword hits",
      emotionUpsDowns: "Emotion ups and downs — daily positive/negative sequence",
      monthTheme: "Month theme — infer from Topic + remarks/KPIs",
      comparisons: "3-month comparison — current + prior 2 months",
    },
    quarter: {
      weeklyPassFail: "Weekly pass/fail — which weeks meet targets and why",
      patternsWeekday: "Weekday patterns — recurring Mon–Sun habits/risks",
      placeHabits: "Place & lifestyle — geography/place frequency effects",
      rhythmRegularity: "Rhythm regularity — steady vs chaotic week-to-week rhythm",
      severeIssues: "Severe issues — biggest red flags in the quarter",
      habitChanges: "Habit changes — activities that appeared or disappeared",
      emotionFactors: "Emotion factors — keywords co-occurring with activity/place",
      comparisons: "3-quarter comparison — current + prior 2 quarters",
    },
    year: {
      monthlyWorkRest: "Monthly Work / Rest — hours and balance by month",
      passFailTrend: "Pass/fail trend — qualifying vs non-qualifying days over months",
      placeEmotion: "Place & emotion — geography vs mood/lifestyle",
      changePoints: "Change points — which month the pattern shifted",
      monthStories: "Month stories — what happened each month (from remarks/topic)",
      comparisons: "3-year comparison — current + prior 2 years",
    },
  };
}

function aiMatchKeywords_(text, keywords) {
  var s = String(text || "").toLowerCase();
  if (!s) return [];
  var hits = [];
  for (var i = 0; i < keywords.length; i++) {
    var k = String(keywords[i] || "").trim();
    if (!k) continue;
    if (s.indexOf(k.toLowerCase()) >= 0) hits.push(k);
  }
  return hits;
}

function aiConsecutiveTrueStreaks_(boolArr) {
  var max = 0;
  var cur = 0;
  var runs = [];
  var runStart = -1;
  for (var i = 0; i < boolArr.length; i++) {
    if (boolArr[i]) {
      if (cur === 0) runStart = i;
      cur++;
      if (cur > max) max = cur;
    } else {
      if (cur > 0) runs.push({ startIndex: runStart, length: cur });
      cur = 0;
      runStart = -1;
    }
  }
  if (cur > 0) runs.push({ startIndex: runStart, length: cur });
  return { maxStreak: max, runs: runs };
}

function aiHasConsecutiveTrue_(boolArr, n) {
  return aiConsecutiveTrueStreaks_(boolArr).maxStreak >= n;
}

/**
 * 按 wake-day 掃 state，產出日級 KPI 序列（睡眠／運動／情緒等）
 */
function aiBuildDailySeries_(state, range) {
  var list = aiSortedEvents_(state);
  var nowMs = Date.now();
  var kw = (getAiPromptConfig_().emotionKeywords) || aiDefaultEmotionKeywords_();
  var days = [];
  var dayCursor = aiWakeDayStartMs_(range.fromMs);
  var endBound = range.toMs;
  var guard = 0;
  while (dayCursor <= endBound && guard < 400) {
    guard++;
    var dayEnd = dayCursor + 86400000;
    var dayYmd = aiYmdLocal_(dayCursor);
    var sleepMs = 0;
    var meditateMs = 0;
    var exerciseMs = 0;
    var workMs = 0;
    var restMs = 0;
    var tradeMs = 0;
    var socialMs = 0;
    var negHits = [];
    var posHits = [];
    var places = {};
    for (var j = 0; j < list.length; j++) {
      var st = aiParseStartMs_(list[j].start);
      if (isNaN(st) || st < dayCursor || st >= dayEnd) continue;
      var seg = aiSegmentMs_(list, j, nowMs);
      var actName = aiActivityName_(state, list[j].activityId) || "(unknown)";
      var nm = aiNormKey_(actName);
      var gg = String(list[j].group || list[j].category || "").trim();
      if (gg === "Work") workMs += seg;
      if (gg === "Rest") restMs += seg;
      if (AI_TRADING_KEYS[nm]) tradeMs += seg;
      if (AI_SOCIAL_KEYS[nm]) socialMs += seg;
      if (nm === AI_SLEEP_KEY) sleepMs += seg;
      if (nm === AI_MEDITATE_KEY) meditateMs += seg;
      if (AI_EXERCISE_KEYS[nm]) exerciseMs += seg;
      if (nm === "photoing" || nm === "photography") {
        if (aiPhotoingCountsAsExercise_(list[j])) exerciseMs += seg;
      }
      var pl = String(list[j].place || "").trim();
      if (pl) places[pl] = (places[pl] || 0) + 1;
      var rm = String(list[j].remark || "").trim();
      if (rm) {
        var nh = aiMatchKeywords_(rm, kw.negative || []);
        var ph = aiMatchKeywords_(rm, kw.positive || []);
        if (nh.length) {
          negHits.push({
            activity: actName,
            remark: rm.slice(0, 200),
            keywords: nh,
            startLocal: aiLocalDateTimeStr_(st),
            startMs: st,
          });
        }
        if (ph.length) {
          posHits.push({
            activity: actName,
            remark: rm.slice(0, 200),
            keywords: ph,
            startLocal: aiLocalDateTimeStr_(st),
          });
        }
      }
    }
    days.push({
      ymd: dayYmd,
      weekday: aiWeekdayName_(dayCursor),
      workHours: aiHours_(workMs),
      restHours: aiHours_(restMs),
      tradingHours: aiHours_(tradeMs),
      sleepHours: aiHours_(sleepMs),
      sleepShort: sleepMs > 0 && sleepMs < AI_SLEEP_SHORT_MS,
      sleepLong: sleepMs > AI_SLEEP_LONG_MS,
      sleepAnomaly: (sleepMs > 0 && sleepMs < AI_SLEEP_SHORT_MS) || sleepMs > AI_SLEEP_LONG_MS,
      // 無 Sleeping 打卡唔當異常（避免全部 false positive）；只計有記錄嘅異常
      meditate: meditateMs > 0,
      exerciseOver30m: exerciseMs >= AI_EXERCISE_MIN_MS,
      exerciseMinutes: aiMinutes_(exerciseMs),
      vacation: workMs < AI_WORK_VACATION_MS,
      overWork: workMs > AI_WORK_IDEAL_MAX_MS,
      workOver6h: workMs >= AI_WORK_OVERLOAD_MAX_MS,
      tradingOver2h: tradeMs > AI_TRADE_CAP_MS,
      social: socialMs > (typeof AI_SOCIAL_ACTIVE_MS !== "undefined" ? AI_SOCIAL_ACTIVE_MS : 2 * 3600000),
      socialHours: aiHours_(socialMs),
      negativeRemarks: negHits,
      positiveRemarks: posHits,
      hasNegative: negHits.length > 0,
      hasPositive: posHits.length > 0,
      places: places,
    });
    dayCursor = dayEnd;
  }

  // 用 Energy DF 日型覆寫 vacation／overWork（若重播可用）
  if (typeof aiEnergyReplayDayMetrics_ === "function") {
    try {
      var em = aiEnergyReplayDayMetrics_(state, range.fromMs, range.toMs, 14);
      var by = (em && em.byYmd) || {};
      for (var di = 0; di < days.length; di++) {
        var er = by[days[di].ymd];
        if (!er || er.noData) continue;
        days[di].dayType = er.dayType;
        days[di].startQuality = er.startQuality;
        days[di].wakeDf = er.wakeDf;
        days[di].endDf = er.endDf;
        days[di].dfDrained = er.dfDrained;
        days[di].vacation = er.dayType === "vacation";
        days[di].hea = er.dayType === "hea";
        days[di].idealFocus = er.dayType === "idealFocus";
        days[di].overWork = er.dayType === "overload";
        days[di].workOver6h = er.dayType === "overload";
      }
    } catch (eEn) {}
  }
  return days;
}

function aiContext72hBefore_(state, centerMs) {
  var list = aiSortedEvents_(state);
  var nowMs = Date.now();
  var from = centerMs - 72 * 3600000;
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var st = aiParseStartMs_(list[i].start);
    if (isNaN(st) || st < from || st > centerMs) continue;
    var seg = aiSegmentMs_(list, i, nowMs);
    var name = aiActivityName_(state, list[i].activityId) || "(unknown)";
    out.push({
      startLocal: aiLocalDateTimeStr_(st),
      activity: name,
      group: String(list[i].group || list[i].category || "").trim(),
      minutes: aiMinutes_(seg),
      place: String(list[i].place || "").trim(),
      people: aiPeopleList_(list[i]),
      remark: String(list[i].remark || "").trim().slice(0, 200),
    });
  }
  return out.slice(-40);
}

/**
 * 未勾選章節：刪走對應重列表，減 token；summary／passFail／targets 一律保留。
 */
function aiFilterEnabledKpis_(obj, sections) {
  if (!obj || !sections || typeof sections !== "object") return obj;
  function off(k) {
    return sections[k] === false;
  }
  if (off("negativeRemarks72h") && off("chaosStreak") && off("emotionFactors") && off("emotionUpsDowns")) {
    obj.negativeRemarks = [];
  }
  if (off("positiveRemarks") && off("positiveEmotionDays") && off("emotionUpsDowns")) {
    obj.positiveRemarks = [];
  }
  if (off("emotionUpsDowns") && off("emotionFactors")) {
    obj.emotionUpsDowns = [];
  }
  if (off("exerciseGaps") && off("exerciseDays")) {
    obj.exerciseGaps = [];
  }
  if (off("patternsWeekday")) {
    obj.byWeekday = {};
  }
  if (off("placeHabits") && off("placeEmotion")) {
    obj.placeCounts = {};
  }
  if (off("weeklyPassFail") && off("rhythmRegularity")) {
    obj.weeklyEval = [];
  }
  if (off("habitChanges")) {
    obj.habitChanges = null;
  }
  if (off("emotionFactors")) {
    obj.emotionFactors = null;
  }
  if (off("monthlyWorkRest") && off("passFailTrend") && off("changePoints") && off("monthStories")) {
    obj.monthlySeries = [];
  }
  if (off("comparisons")) {
    /* comparisons 喺 stats 頂層，唔喺 kpis */
  }
  return obj;
}

/** 活動按 ISO 週出現集合 → 新增／消失 */
function aiHabitChangesFromDaily_(daily, state, range) {
  var list = aiSortedEvents_(state);
  var nowMs = Date.now();
  var byWeek = {};
  for (var i = 0; i < list.length; i++) {
    var st = aiParseStartMs_(list[i].start);
    if (isNaN(st) || st < range.fromMs || st > range.toMs) continue;
    var wk = typeof aiIsoWeekKeyFromDate_ === "function" ? aiIsoWeekKeyFromDate_(new Date(st)) : aiYmdLocal_(st).slice(0, 7);
    if (!byWeek[wk]) byWeek[wk] = {};
    var nm = aiNormKey_(aiActivityName_(state, list[i].activityId) || "");
    if (nm) byWeek[wk][nm] = 1;
  }
  var keys = Object.keys(byWeek).sort();
  var appeared = [];
  var disappeared = [];
  for (var w = 1; w < keys.length; w++) {
    var prev = byWeek[keys[w - 1]];
    var cur = byWeek[keys[w]];
    for (var a in cur) {
      if (Object.prototype.hasOwnProperty.call(cur, a) && !prev[a]) {
        appeared.push({ weekKey: keys[w], activity: a, afterWeek: keys[w - 1] });
      }
    }
    for (var b in prev) {
      if (Object.prototype.hasOwnProperty.call(prev, b) && !cur[b]) {
        disappeared.push({ weekKey: keys[w], activity: b, afterWeek: keys[w - 1] });
      }
    }
  }
  return {
    weekKeys: keys,
    appeared: appeared.slice(0, 40),
    disappeared: disappeared.slice(0, 40),
  };
}

/** 負面／正面 keyword × activity／place 共現 */
function aiEmotionFactorsFromRemarks_(negAll, posAll) {
  function accum(rows) {
    var byAct = {};
    var byPlace = {};
    var byKw = {};
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var act = String(r.activity || "(unknown)");
      byAct[act] = (byAct[act] || 0) + 1;
      var places = {};
      if (r.context72h) {
        for (var c = 0; c < r.context72h.length; c++) {
          var pl = String(r.context72h[c].place || "").trim();
          if (pl) places[pl] = 1;
        }
      }
      for (var p in places) {
        if (Object.prototype.hasOwnProperty.call(places, p)) byPlace[p] = (byPlace[p] || 0) + 1;
      }
      var kws = r.keywords || [];
      for (var k = 0; k < kws.length; k++) {
        var kw = String(kws[k]);
        byKw[kw] = (byKw[kw] || 0) + 1;
      }
    }
    function top(map, n) {
      var arr = [];
      for (var key in map) {
        if (Object.prototype.hasOwnProperty.call(map, key)) arr.push({ key: key, count: map[key] });
      }
      arr.sort(function (x, y) {
        return y.count - x.count;
      });
      return arr.slice(0, n || 8);
    }
    return { byActivity: top(byAct, 10), byPlace: top(byPlace, 10), byKeyword: top(byKw, 12) };
  }
  return { negative: accum(negAll || []), positive: accum(posAll || []) };
}

/** 年／季：按曆月 Work／Rest 時數 + 粗略合格 */
function aiMonthlySeriesFromDaily_(daily) {
  var months = {};
  for (var i = 0; i < daily.length; i++) {
    var ym = String(daily[i].ymd || "").slice(0, 7);
    if (!ym || ym.length < 7) continue;
    if (!months[ym]) {
      months[ym] = {
        monthKey: ym,
        days: 0,
        workHours: 0,
        restHours: 0,
        overWorkDays: 0,
        workOver6hDays: 0,
        negativeDays: 0,
        positiveDays: 0,
        socialDays: 0,
        meditateDays: 0,
        exerciseDays: 0,
        vacationDays: 0,
      };
    }
    var m = months[ym];
    m.days++;
    m.workHours += Number(daily[i].workHours) || 0;
    m.restHours += Number(daily[i].restHours) || 0;
    if (daily[i].overWork) m.overWorkDays++;
    if (daily[i].workOver6h) m.workOver6hDays++;
    if (daily[i].hasNegative) m.negativeDays++;
    if (daily[i].hasPositive) m.positiveDays++;
    if (daily[i].social) m.socialDays++;
    if (daily[i].meditate) m.meditateDays++;
    if (daily[i].exerciseOver30m) m.exerciseDays++;
    if (daily[i].vacation) m.vacationDays++;
  }
  var keys = Object.keys(months).sort();
  var out = [];
  for (var j = 0; j < keys.length; j++) {
    var row = months[keys[j]];
    row.workHours = Math.round(row.workHours * 10) / 10;
    row.restHours = Math.round(row.restHours * 10) / 10;
    row.avgWorkHoursPerDay = row.days ? Math.round((row.workHours / row.days) * 10) / 10 : 0;
    var fails = [];
    if (row.workOver6hDays >= 2) fails.push("work6hDays");
    if (row.negativeDays >= 3) fails.push("chaosDays");
    if (row.overWorkDays > Math.ceil(row.days * 0.35)) fails.push("overWorkHeavy");
    row.pass = fails.length === 0;
    row.fails = fails;
    out.push(row);
  }
  return out;
}

/** 找出 overWork／chaos 比率跳升嘅月份 */
function aiChangePointsFromMonthly_(monthly) {
  var points = [];
  for (var i = 1; i < monthly.length; i++) {
    var a = monthly[i - 1];
    var b = monthly[i];
    var owA = a.days ? a.overWorkDays / a.days : 0;
    var owB = b.days ? b.overWorkDays / b.days : 0;
    var negA = a.days ? a.negativeDays / a.days : 0;
    var negB = b.days ? b.negativeDays / b.days : 0;
    if (owB - owA >= 0.2) {
      points.push({
        monthKey: b.monthKey,
        type: "overWorkSpike",
        fromRate: Math.round(owA * 100) / 100,
        toRate: Math.round(owB * 100) / 100,
      });
    }
    if (negB - negA >= 0.2) {
      points.push({
        monthKey: b.monthKey,
        type: "chaosSpike",
        fromRate: Math.round(negA * 100) / 100,
        toRate: Math.round(negB * 100) / 100,
      });
    }
    if (a.pass && !b.pass) {
      points.push({ monthKey: b.monthKey, type: "passToFail", fails: b.fails });
    }
  }
  return points.slice(0, 12);
}

/** 週序列節奏：workHours／socialDays 變異（愈細愈規律） */
function aiRhythmRegularityFromWeekly_(weeklyEval) {
  if (!weeklyEval || weeklyEval.length < 2) {
    return { weekCount: weeklyEval ? weeklyEval.length : 0, note: "樣本不足" };
  }
  var works = [];
  var socials = [];
  for (var i = 0; i < weeklyEval.length; i++) {
    works.push(Number(weeklyEval[i].workHours) || 0);
    socials.push(Number(weeklyEval[i].socialDays) || 0);
  }
  function stats(arr) {
    var n = arr.length;
    var sum = 0;
    for (var i = 0; i < n; i++) sum += arr[i];
    var mean = sum / n;
    var v = 0;
    for (var j = 0; j < n; j++) v += (arr[j] - mean) * (arr[j] - mean);
    var variance = v / n;
    return {
      mean: Math.round(mean * 10) / 10,
      stdev: Math.round(Math.sqrt(variance) * 10) / 10,
    };
  }
  var w = stats(works);
  var s = stats(socials);
  var passN = 0;
  for (var p = 0; p < weeklyEval.length; p++) {
    if (weeklyEval[p].pass) passN++;
  }
  return {
    weekCount: weeklyEval.length,
    workHours: w,
    socialDays: s,
    weeksPassed: passN,
    regularityHint:
      w.stdev <= 4 && s.stdev <= 1.2 ? "relatively_steady" : w.stdev >= 8 || s.stdev >= 2 ? "chaotic" : "mixed",
  };
}

/**
 * @param {object} stats aggregatePeriodStatsForAi_ 結果（已有 processAudits）
 */
function enrichStatsWithPeriodKpis_(state, stats) {
  if (!stats || !stats.range) return stats;
  var range = {
    fromMs: new Date(stats.range.from + "T00:00:00").getTime(),
    toMs: new Date(stats.range.to + "T23:59:59.999").getTime(),
    periodType: stats.periodType,
    periodKey: stats.periodKey,
  };
  var cfg = getAiPromptConfig_();
  var pType = String(stats.periodType || "custom").toLowerCase();
  var pcfg = (cfg.periodConfig && cfg.periodConfig[pType]) || {
    sections: aiDefaultPeriodSections_(pType),
    notes: "",
  };
  var sections = pcfg.sections || aiDefaultPeriodSections_(pType);
  var daily = aiBuildDailySeries_(state, range);
  var audits = stats.processAudits || {};
  var rhythm = audits.rhythmInterleaving || {};
  var socialBat = audits.socialBattery || {};

  // attach 72h context to negative remarks
  var negAll = [];
  var posAll = [];
  for (var d = 0; d < daily.length; d++) {
    for (var n = 0; n < daily[d].negativeRemarks.length; n++) {
      var row = daily[d].negativeRemarks[n];
      row.ymd = daily[d].ymd;
      row.context72h = aiContext72hBefore_(state, row.startMs);
      negAll.push(row);
    }
    for (var p = 0; p < daily[d].positiveRemarks.length; p++) {
      var prow = daily[d].positiveRemarks[p];
      prow.ymd = daily[d].ymd;
      posAll.push(prow);
    }
  }

  var vacationDays = daily.filter(function (x) {
    return x.vacation;
  });
  var overWorkDays = daily.filter(function (x) {
    return x.overWork;
  });
  var work6Days = daily.filter(function (x) {
    return x.workOver6h;
  });
  var tradeOverDays = daily.filter(function (x) {
    return x.tradingOver2h;
  });
  var sleepShort = daily.filter(function (x) {
    return x.sleepShort;
  });
  var sleepLong = daily.filter(function (x) {
    return x.sleepLong;
  });
  var sleepAnom = daily.filter(function (x) {
    return x.sleepAnomaly;
  });
  var meditateDays = daily.filter(function (x) {
    return x.meditate;
  });
  var exerciseDays = daily.filter(function (x) {
    return x.exerciseOver30m;
  });
  var socialDays = daily.filter(function (x) {
    return x.social;
  });
  var posDays = daily.filter(function (x) {
    return x.hasPositive;
  });
  var negDays = daily.filter(function (x) {
    return x.hasNegative;
  });

  var overWorkFlags = daily.map(function (x) {
    return !!x.overWork;
  });
  var work6Flags = daily.map(function (x) {
    return !!x.workOver6h;
  });
  var tradeFlags = daily.map(function (x) {
    return !!x.tradingOver2h;
  });
  var sleepFlags = daily.map(function (x) {
    return !!x.sleepAnomaly;
  });
  var negFlags = daily.map(function (x) {
    return !!x.hasNegative;
  });

  var fatigueSwitchCounts = rhythm.fatigueSwitchCounts || {
    ideal: 0,
    good: 0,
    poor: 0,
  };
  var fatigueSwitchPoor = Number(fatigueSwitchCounts.poor || 0);
  var fatigueTotalDropSum = Number(rhythm.fatigueTotalDropSum || 0);

  // exercise gaps (days between exercise days)
  var exGaps = [];
  for (var ei = 1; ei < exerciseDays.length; ei++) {
    var t0 = new Date(exerciseDays[ei - 1].ymd + "T12:00:00").getTime();
    var t1 = new Date(exerciseDays[ei].ymd + "T12:00:00").getTime();
    exGaps.push({
      from: exerciseDays[ei - 1].ymd,
      to: exerciseDays[ei].ymd,
      gapDays: Math.round((t1 - t0) / 86400000),
    });
  }

  var weekRows = socialBat.byWeek || stats.weeklyPerformance || [];
  var overSocialWeekKeys = [];
  for (var wi = 0; wi < weekRows.length; wi++) {
    var sd = weekRows[wi].socialDays != null ? weekRows[wi].socialDays : 0;
    if (sd > 3) overSocialWeekKeys.push(weekRows[wi].weekKey || weekRows[wi].weekKey);
  }
  // consecutive over-social weeks
  var overSocialConsecutive = false;
  for (var wj = 1; wj < weekRows.length; wj++) {
    var a = (weekRows[wj - 1].socialDays != null ? weekRows[wj - 1].socialDays : 0) > 3;
    var b = (weekRows[wj].socialDays != null ? weekRows[wj].socialDays : 0) > 3;
    if (a && b) overSocialConsecutive = true;
  }

  var passFail = { ok: true, fails: [], warns: [] };
  function fail(code, msg) {
    passFail.ok = false;
    passFail.fails.push({ code: code, message: msg });
  }
  function warn(code, msg) {
    passFail.warns.push({ code: code, message: msg });
  }

  if (pType === "week") {
    if (fatigueSwitchPoor >= 3) {
      fail("fatigueSwitchPoor", "Fatigue 切換差劣 " + fatigueSwitchPoor + " 次（理想 <3）");
    }
    if (socialDays.length >= 3) fail("socialDays", "Social Battery 日 " + socialDays.length + "（理想 <3）");
    if (vacationDays.length < 1 || vacationDays.length > 2) {
      warn("vacationDays", "放假日 " + vacationDays.length + "（理想 1–2）");
    }
    if (overWorkDays.length > 2) fail("overWorkDays", "Overload 日 " + overWorkDays.length + "（理想 ≤2）");
    if (aiHasConsecutiveTrue_(overWorkFlags, 2)) fail("overWorkStreak", "Overload 連續 ≥2 日");
    if (aiConsecutiveTrueStreaks_(sleepFlags).maxStreak > 2) {
      fail("sleepStreak", "睡眠異常連續 >2 日");
    }
    if (meditateDays.length < 6) warn("meditateDays", "Meditate 日 " + meditateDays.length + "（理想 ≥6）");
    if (exerciseDays.length < 2 || exerciseDays.length > 5) {
      warn("exerciseDays", "運動≥30m 日 " + exerciseDays.length + "（理想 2–5）");
    }
  } else if (pType === "month") {
    if (fatigueSwitchPoor > 8) {
      fail("fatigueSwitchPoor", "Fatigue 切換差劣 " + fatigueSwitchPoor + " 次（理想 ≤8／月）");
    }
    if (overSocialConsecutive) fail("overSocialWeeks", "連續兩週 over social（socialDays>3）");
    if (aiHasConsecutiveTrue_(negFlags, 3)) fail("chaosStreak", "負面情緒／chaos 連續 ≥3 日");
    if (aiHasConsecutiveTrue_(sleepFlags, 2)) fail("sleepStreak", "睡眠異常連續 ≥2 日");
    if (aiHasConsecutiveTrue_(work6Flags, 2)) fail("work6Streak", "Overload 連續 ≥2 日");
    if (aiHasConsecutiveTrue_(tradeFlags, 2)) fail("tradeStreak", "Trading>2h 連續 ≥2 日");
  }

  // weekday pattern + places for quarter/year
  var byWeekday = {};
  var placeCounts = {};
  for (var di = 0; di < daily.length; di++) {
    var wd = daily[di].weekday || "?";
    if (!byWeekday[wd]) byWeekday[wd] = { days: 0, overWork: 0, negative: 0, social: 0 };
    byWeekday[wd].days++;
    if (daily[di].overWork) byWeekday[wd].overWork++;
    if (daily[di].hasNegative) byWeekday[wd].negative++;
    if (daily[di].social) byWeekday[wd].social++;
    for (var pk in daily[di].places) {
      if (Object.prototype.hasOwnProperty.call(daily[di].places, pk)) {
        placeCounts[pk] = (placeCounts[pk] || 0) + daily[di].places[pk];
      }
    }
  }

  // weekly pass/fail for quarter (reuse weeklyPerformance + daily)
  var weeklyEval = [];
  var wp = stats.weeklyPerformance || [];
  for (var q = 0; q < wp.length; q++) {
    var wrow = wp[q];
    var failsW = [];
    if ((wrow.socialDays || 0) >= 3) failsW.push("socialDays>=3");
    var tier = wrow.workLoadTierDays || wrow.workLoadTierDays || {};
    // approximate overwork from overload if present
    var ow = tier.overload || 0;
    if (ow > 2) failsW.push("overWorkDays>2");
    weeklyEval.push({
      weekKey: wrow.weekKey,
      weekLabel: wrow.weekLabel || (typeof aiWeekLabelFromKey_ === "function" ? aiWeekLabelFromKey_(wrow.weekKey) : wrow.weekKey),
      pass: failsW.length === 0,
      fails: failsW,
      socialDays: wrow.socialDays,
      trueFocusWorkHours: wrow.trueFocusWorkHours,
      workHours: wrow.workHours,
      distractionHours: wrow.distractionHours,
    });
  }

  var habitChanges = null;
  var emotionFactors = null;
  var monthlySeries = [];
  var changePoints = [];
  var rhythmRegularity = null;
  var meditateKeep = null;
  if (pType === "quarter" || pType === "year" || sections.habitChanges) {
    habitChanges = aiHabitChangesFromDaily_(daily, state, range);
  }
  if (pType === "quarter" || pType === "year" || sections.emotionFactors) {
    emotionFactors = aiEmotionFactorsFromRemarks_(negAll, posAll);
  }
  if (pType === "month" || pType === "quarter" || pType === "year" || sections.monthlyWorkRest) {
    monthlySeries = aiMonthlySeriesFromDaily_(daily);
    changePoints = aiChangePointsFromMonthly_(monthlySeries);
  }
  if (pType === "quarter" || sections.rhythmRegularity || sections.weeklyPassFail) {
    rhythmRegularity = aiRhythmRegularityFromWeekly_(weeklyEval);
  }
  if (pType === "month" || sections.meditateKeep) {
    var medFlags = daily.map(function (x) {
      return !!x.meditate;
    });
    var medStreak = aiConsecutiveTrueStreaks_(medFlags);
    meditateKeep = {
      meditateDays: meditateDays.length,
      dayCount: daily.length,
      coverage: daily.length ? Math.round((meditateDays.length / daily.length) * 100) / 100 : 0,
      maxStreak: medStreak.maxStreak,
      kept: meditateDays.length >= Math.max(1, Math.floor(daily.length * 0.6)),
    };
  }

  var kpis = {
    periodType: pType,
    periodKey: stats.periodKey,
    enabledSections: sections,
    periodNotes: String(pcfg.notes || ""),
    targets: {
      week: {
        fatigueSwitchPoorMax: 2,
        socialDaysMax: 2,
        vacationDaysIdeal: "1-2",
        overloadDaysMax: 2,
        overloadNoConsecutive: true,
        sleepAnomalyMaxConsecutive: 2,
        meditateDaysMin: 6,
        exerciseDaysIdeal: "2-5",
      },
      month: {
        fatigueSwitchPoorMax: 8,
        overSocialNoTwoConsecutiveWeeks: true,
        chaosMaxConsecutiveDays: 2,
        sleepAnomalyMaxConsecutive: 1,
        overloadNoConsecutive: true,
        tradingOver2hNoConsecutive: true,
      },
    },
    summary: {
      fatigueSwitchCounts: fatigueSwitchCounts,
      fatigueSwitchPoor: fatigueSwitchPoor,
      fatigueTotalDropSum: fatigueTotalDropSum,
      socialDays: socialDays.length,
      vacationDays: vacationDays.length,
      overloadDays: overWorkDays.length,
      workOver6hDays: work6Days.length,
      tradingOver2hDays: tradeOverDays.length,
      sleepShortDays: sleepShort.length,
      sleepLongDays: sleepLong.length,
      sleepAnomalyDays: sleepAnom.length,
      sleepAnomalyMaxStreak: aiConsecutiveTrueStreaks_(sleepFlags).maxStreak,
      meditateDays: meditateDays.length,
      exerciseOver30mDays: exerciseDays.length,
      positiveEmotionDays: posDays.length,
      negativeEmotionDays: negDays.length,
      chaosMaxStreak: aiConsecutiveTrueStreaks_(negFlags).maxStreak,
      overloadMaxStreak: aiConsecutiveTrueStreaks_(overWorkFlags).maxStreak,
      work6hMaxStreak: aiConsecutiveTrueStreaks_(work6Flags).maxStreak,
      tradingOver2hMaxStreak: aiConsecutiveTrueStreaks_(tradeFlags).maxStreak,
      overSocialConsecutiveWeeks: overSocialConsecutive,
    },
    lists: {
      vacationDays: vacationDays.map(function (x) {
        return x.ymd;
      }),
      overWorkDays: overWorkDays.map(function (x) {
        return {
          ymd: x.ymd,
          dayType: x.dayType || null,
          endDf: x.endDf != null ? x.endDf : null,
          dfDrained: x.dfDrained != null ? x.dfDrained : null,
          workHours: x.workHours,
        };
      }),
      workOver6hDays: work6Days.map(function (x) {
        return { ymd: x.ymd, workHours: x.workHours };
      }),
      tradingOver2hDays: tradeOverDays.map(function (x) {
        return { ymd: x.ymd, tradingHours: x.tradingHours };
      }),
      sleepShortDays: sleepShort.map(function (x) {
        return { ymd: x.ymd, sleepHours: x.sleepHours };
      }),
      sleepLongDays: sleepLong.map(function (x) {
        return { ymd: x.ymd, sleepHours: x.sleepHours };
      }),
      meditateDays: meditateDays.map(function (x) {
        return x.ymd;
      }),
      exerciseDays: exerciseDays.map(function (x) {
        return { ymd: x.ymd, minutes: x.exerciseMinutes };
      }),
      socialDays: socialDays.map(function (x) {
        return x.ymd;
      }),
      positiveEmotionDays: posDays.map(function (x) {
        return x.ymd;
      }),
      negativeEmotionDays: negDays.map(function (x) {
        return x.ymd;
      }),
    },
    negativeRemarks: negAll.slice(0, 30),
    positiveRemarks: posAll.slice(0, 30),
    exerciseGaps: exGaps.slice(0, 20),
    byWeekday: byWeekday,
    placeCounts: placeCounts,
    weeklyEval: weeklyEval,
    rhythmRegularity: rhythmRegularity,
    habitChanges: habitChanges,
    emotionFactors: emotionFactors,
    monthlySeries: monthlySeries,
    changePoints: changePoints,
    meditateKeep: meditateKeep,
    emotionUpsDowns: daily.map(function (x) {
      return {
        ymd: x.ymd,
        positive: x.hasPositive,
        negative: x.hasNegative,
        posCount: x.positiveRemarks.length,
        negCount: x.negativeRemarks.length,
      };
    }),
    passFail: passFail,
    dailyLite: daily.map(function (x) {
      return {
        ymd: x.ymd,
        weekday: x.weekday,
        workHours: x.workHours,
        restHours: x.restHours,
        sleepHours: x.sleepHours,
        overWork: x.overWork,
        workOver6h: x.workOver6h,
        social: x.social,
        meditate: x.meditate,
        exerciseOver30m: x.exerciseOver30m,
        hasNegative: x.hasNegative,
        hasPositive: x.hasPositive,
      };
    }),
  };

  stats.kpis = aiFilterEnabledKpis_(kpis, sections);
  stats.enabledSections = sections;
  stats.periodNotes = String(pcfg.notes || "");
  stats.reportLens =
    pType === "week"
      ? "week_df_energy_audits"
      : pType === "month"
        ? "month_day_count_audits"
        : pType === "quarter"
          ? "quarter_weekly_trends"
          : pType === "year"
            ? "year_monthly_trends"
            : "custom_kpis";
  stats.reportLensNote =
    pType === "week" || pType === "month"
      ? "只寫 enabledSections=true 嘅章節。週／月報必須以 processAudits（DF 日型、Start、Fatigue 切換、Social Battery）為主；trueFocus／remarks 輔助。禁止 OCD 鎖死／DMN 間隔。合格門檻用 kpis.passFail／kpis.targets。負面情緒必帶 context72h。"
      : "只寫 enabledSections=true 嘅章節。合格門檻用 kpis.passFail／kpis.targets，唔好自創。負面情緒必帶 context72h。術語首次出現附 termGlossary 定義。";
  // extend glossary
  stats.termGlossary = stats.termGlossary || {};
  stats.termGlossary["睡眠不足"] = "當日 Sleeping 合計 < 6 小時（有打卡先計）";
  stats.termGlossary["睡眠過龍"] = "當日 Sleeping 合計 > 10 小時";
  stats.termGlossary["OverWork"] = "當日 Work Group > 4 小時";
  stats.termGlossary["運動日"] =
    "gyming／hiking／yogaing／running 等合計 ≥ 30 分鐘；Photoing 暫時唔計，除非 remark 注明高強度";
  stats.termGlossary["habitChanges"] = "相鄰兩週活動出現／消失（appeared／disappeared）";
  stats.termGlossary["emotionFactors"] = "情緒 keyword 與 activity／place 共現次數";
  stats.termGlossary["rhythmRegularity"] =
    "週序列 workHours／socialDays 標準差；relatively_steady／mixed／chaotic";
  return stats;
}

function aiPeriodOutlineFromSections_(periodType, sections, notes) {
  var labels = aiPeriodSectionLabels_()[periodType] || {};
  var lines = ["期別：" + periodType + "。只輸出以下已啟用章節（順序可調整，唔好加未啟用章節）："];
  var n = 1;
  for (var k in sections) {
    if (!Object.prototype.hasOwnProperty.call(sections, k)) continue;
    if (!sections[k]) continue;
    lines.push(n + ". " + (labels[k] || k) + " （sectionKey=" + k + "）");
    n++;
  }
  lines.push(n + ". 下期可執行實驗（細、可量度）");
  if (notes && String(notes).trim()) {
    lines.push("補充要求：\n" + String(notes).trim());
  }
  lines.push("首次術語附定義；表格用 GFM；只用 DATA_JSON。");
  return lines.join("\n");
}

var EMOTION_QUEUE_PROP_ = "emotionTriggerQueueV1";

function readEmotionQueue_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(EMOTION_QUEUE_PROP_);
  if (!raw) return [];
  try {
    var arr = JSON.parse(raw);
    return Object.prototype.toString.call(arr) === "[object Array]" ? arr : [];
  } catch (e) {
    return [];
  }
}

function writeEmotionQueue_(arr) {
  PropertiesService.getScriptProperties().setProperty(
    EMOTION_QUEUE_PROP_,
    JSON.stringify(arr || [])
  );
}

/**
 * 負面情緒：唔即刻寄，入 queue；第二日 08:00 用過去 72h 打 AI 分析。
 */
function queueEmotionTrigger_(state, ev, keywordsHit) {
  if (!ev || !keywordsHit || !keywordsHit.length) return { queued: false };
  var t0 = aiParseStartMs_(ev.start);
  if (isNaN(t0)) return { queued: false, error: "bad_start" };
  var wakeYmd = aiYmdLocal_(aiWakeDayStartMs_(t0));
  var eid = String(ev.id || ev.start || "");
  var q = readEmotionQueue_();
  for (var i = 0; i < q.length; i++) {
    if (String(q[i].eventId) === eid && !q[i].sentAt) {
      return { queued: false, already: true, wakeYmd: wakeYmd };
    }
  }
  q.push({
    eventId: eid,
    startMs: t0,
    wakeYmd: wakeYmd,
    keywords: keywordsHit.slice(),
    remark: String(ev.remark || "").trim(),
    activityId: ev.activityId,
    queuedAt: Date.now(),
    sentAt: null,
  });
  // 只保留最近 60 日
  var cut = Date.now() - 60 * 86400000;
  q = q.filter(function (row) {
    return Number(row.queuedAt) >= cut || !row.sentAt;
  });
  writeEmotionQueue_(q);
  return { queued: true, wakeYmd: wakeYmd };
}

/** 舊名相容：改為入 queue。 */
function maybeSendEmotionBrief_(state, ev, keywordsHit) {
  var r = queueEmotionTrigger_(state, ev, keywordsHit);
  return { sent: false, queued: !!(r && r.queued), wakeYmd: r && r.wakeYmd, already: r && r.already };
}

/**
 * 08:00 HKT：處理「昨日或更早」入隊、未寄出嘅負面情緒 → AI 分析 72h 成因。
 */
function processEmotionTriggerQueue_() {
  var state = typeof readStateFromSheet_ === "function" ? readStateFromSheet_() : null;
  if (!state) return { processed: 0, error: "no_state" };
  var todayWake = aiYmdLocal_(aiWakeDayStartMs_(Date.now()));
  var q = readEmotionQueue_();
  var processed = 0;
  var errors = [];
  for (var i = 0; i < q.length; i++) {
    var row = q[i];
    if (row.sentAt) continue;
    // 第二日先處理：觸發當日 wakeYmd < 今日
    if (!row.wakeYmd || String(row.wakeYmd) >= String(todayWake)) continue;
    try {
      var r = sendEmotionTriggerAiReport_(state, row);
      if (r && r.ok) {
        row.sentAt = Date.now();
        processed++;
      } else {
        errors.push(String(r && r.error ? r.error : "send_failed"));
      }
    } catch (e) {
      errors.push(String(e && e.message ? e.message : e));
    }
  }
  writeEmotionQueue_(q);
  return { processed: processed, errors: errors.slice(0, 5) };
}

function sendEmotionTriggerAiReport_(state, row) {
  var t0 = Number(row.startMs);
  if (isNaN(t0)) return { ok: false, error: "bad_start" };
  var act = aiActivityName_(state, row.activityId) || "(unknown)";
  var ctx = aiContext72hBefore_(state, t0);
  var ctxLines = [];
  for (var i = 0; i < ctx.length; i++) {
    var c = ctx[i];
    ctxLines.push(
      "- " +
        c.startLocal +
        " · " +
        c.activity +
        " (" +
        c.minutes +
        "m)" +
        (c.place ? " @ " + c.place : "") +
        (c.remark ? " — " + c.remark : "")
    );
  }
  if (!ctxLines.length) ctxLines.push("- （72h 內無其他事件）");

  var system =
    "你係時間／情緒分析助理。根據用戶活動日誌，分析負面情緒可能觸發原因。" +
    "用繁體中文（可夾粵語口語）。假設→證據→可執行建議；唔好保證市場結果。" +
    "只根據提供嘅 DATA，唔好虛構未出現嘅事件。";
  var user =
    "觸發字眼：" +
    (row.keywords || []).join(", ") +
    "\n事件時間：" +
    aiLocalDateTimeStr_(t0) +
    "\n活動：" +
    act +
    "\nRemark：" +
    String(row.remark || "") +
    "\n\n過去 72 小時活動摘要：\n" +
    ctxLines.join("\n") +
    "\n\n請輸出：\n1) 可能觸發原因（按可能性）\n2) 同 72h 事件嘅對應證據\n3) 今日可執行嘅細行動（少、小、慢）";

  var md = "";
  if (typeof callGeminiWithMessages_ === "function") {
    var ai = callGeminiWithMessages_(system, user);
    md = (ai && ai.markdown) || "";
  }
  if (!md) {
    // fallback：無 Gemini 都寄摘要
    md =
      "# 負面情緒延遲報告（無 AI）\n\n" +
      "觸發：" +
      (row.keywords || []).join(", ") +
      "\n\n## 72h 摘要\n" +
      ctxLines.join("\n");
  }

  var subject = "[Time Stat] Emotion trigger analysis — " + String(row.wakeYmd || "");
  var body =
    "Queued emotion trigger (next-day 08:00)\n\n" +
    "Wake day: " +
    row.wakeYmd +
    "\nKeywords: " +
    (row.keywords || []).join(", ") +
    "\nEvent: " +
    aiLocalDateTimeStr_(t0) +
    " · " +
    act +
    "\nRemark: " +
    String(row.remark || "") +
    "\n\n" +
    md;

  if (typeof mailAiReportToAllowed_ === "function") {
    mailAiReportToAllowed_(subject, body);
  } else {
    MailApp.sendEmail({
      to: (allowedEmailsList_() || [])[0] || Session.getEffectiveUser().getEmail(),
      subject: subject,
      body: body,
    });
  }
  return { ok: true };
}

/** 部署後跑一次：每日 08:00 HKT 處理情緒 queue */
function installEmotionTriggerQueue() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runScheduledEmotionTriggerQueue_") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger("runScheduledEmotionTriggerQueue_")
    .timeBased()
    .atHour(8)
    .nearMinute(0)
    .everyDays(1)
    .inTimezone("Asia/Hong_Kong")
    .create();
  Logger.log("Emotion trigger queue: daily 08:00 Asia/Hong_Kong");
}

function runScheduledEmotionTriggerQueue_() {
  return processEmotionTriggerQueue_();
}

/**
 * 對比 prev／next state，掃新增或 remark 變更嘅負面情緒 → 入 queue（唔即刻寄）。
 */
function scanIncomingForEmotionBriefs_(prevState, nextState) {
  if (!nextState || !nextState.events) return { scanned: 0, queued: 0 };
  var kw = ((getAiPromptConfig_() || {}).emotionKeywords) || aiDefaultEmotionKeywords_();
  var neg = kw.negative || [];
  var prevMap = {};
  var prevEv = (prevState && prevState.events) || [];
  for (var i = 0; i < prevEv.length; i++) {
    var id = String(prevEv[i].id || prevEv[i].start || i);
    prevMap[id] = String(prevEv[i].remark || "");
  }
  var queued = 0;
  var scanned = 0;
  var nextEv = nextState.events || [];
  for (var j = 0; j < nextEv.length; j++) {
    var ev = nextEv[j];
    var eid = String(ev.id || ev.start || j);
    var rm = String(ev.remark || "").trim();
    if (!rm) continue;
    var prevRm = prevMap[eid];
    if (prevRm != null && String(prevRm).trim() === rm) continue; // unchanged
    var hits = aiMatchKeywords_(rm, neg);
    if (!hits.length) continue;
    scanned++;
    var r = queueEmotionTrigger_(nextState, ev, hits);
    if (r && r.queued) queued++;
  }
  return { scanned: scanned, queued: queued, sent: 0 };
}
