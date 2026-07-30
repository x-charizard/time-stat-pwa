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
    return {
      cognitiveLock: true,
      switchFail: true,
      socialDays: true,
      negativeRemarks72h: true,
      positiveRemarks: true,
      vacationDays: true,
      overWorkDays: true,
      sleepAnomalies: true,
      meditateDays: true,
      exerciseDays: true,
      weekTheme: true,
      trueFocus: true,
      comparisons: true,
    };
  }
  if (t === "month") {
    return {
      switchFailDays: true,
      overSocialWeeks: true,
      chaosStreak: true,
      sleepAnomalies: true,
      meditateKeep: true,
      exerciseGaps: true,
      workOver4h: true,
      workOver6h: true,
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
      cognitiveLock:
        "Cognitive Lock — continuous trusted Work >60m closed by a real rest/non-Work log (ideal <3 / week)",
      switchFail:
        "Switch Fail — Diffused Mode between two trusted Work blocks <15m (ideal <3 / week)",
      socialDays: "Social days — Friending/Familying/Socialing present (ideal <3 / week)",
      negativeRemarks72h:
        "Negative emotion remarks — list hits + what happened in the prior 72h",
      positiveRemarks: "Positive emotion remarks — highlight matching remarks",
      vacationDays: "Vacation days — Work Group <2h (ideal 1–2 / week)",
      overWorkDays: "Over-work days — Work Group >4h (ideal ≤2 / week, not consecutive)",
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
      switchFailDays: "Switch-fail days — days with Diffused Mode gap <15m between Work (ideal ≤4 / month)",
      overSocialWeeks: "Over-social weeks — socialDays >3; must not be two weeks in a row",
      chaosStreak: "Chaos streak — negative-emotion days must not run ≥3 consecutive",
      sleepAnomalies: "Sleep anomalies — short/long sleep must not be two consecutive days",
      meditateKeep: "Meditate keep — whether Meditating days stay consistent",
      exerciseGaps: "Exercise gaps — spacing between ≥30m exercise days / recovery",
      workOver4h: "Work >4h days — count and distribution",
      workOver6h: "Work ≥6h days — must not be two consecutive days",
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
    var tradeMs = 0;
    var hasSocial = false;
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
      if (AI_TRADING_KEYS[nm]) tradeMs += seg;
      if (AI_SOCIAL_KEYS[nm]) hasSocial = true;
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
      social: hasSocial,
      negativeRemarks: negHits,
      positiveRemarks: posHits,
      hasNegative: negHits.length > 0,
      hasPositive: posHits.length > 0,
      places: places,
    });
    dayCursor = dayEnd;
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

function aiFilterEnabledKpis_(obj, sections) {
  if (!sections || typeof sections !== "object") return obj;
  // keep structural keys always
  var keep = {
    periodType: obj.periodType,
    periodKey: obj.periodKey,
    targets: obj.targets,
    passFail: obj.passFail,
    enabledSections: obj.enabledSections,
    daily: obj.daily,
    summary: obj.summary,
  };
  return obj; // full kpis for AI evidence; prompt tells which sections to write
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

  var cognitiveLockCount = Number(rhythm.cognitiveLockCount || 0);
  var switchFailCount = Number(rhythm.switchFailCount || 0);
  // 切換失靈「日」：用 switchFails 嘅 ymd unique
  var switchFailYmds = {};
  var sFails = rhythm.switchFails || [];
  for (var sf = 0; sf < sFails.length; sf++) {
    if (sFails[sf].ymd) switchFailYmds[sFails[sf].ymd] = 1;
  }
  var switchFailDayCount = 0;
  for (var sy in switchFailYmds) {
    if (Object.prototype.hasOwnProperty.call(switchFailYmds, sy)) switchFailDayCount++;
  }

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
    if (cognitiveLockCount >= 3) fail("cognitiveLock", "認知鎖死 " + cognitiveLockCount + " 次（理想 <3）");
    if (switchFailCount >= 3) fail("switchFail", "切換失靈 " + switchFailCount + " 次（理想 <3）");
    if (socialDays.length >= 3) fail("socialDays", "Social 日 " + socialDays.length + "（理想 <3）");
    if (vacationDays.length < 1 || vacationDays.length > 2) {
      warn("vacationDays", "放假日 " + vacationDays.length + "（理想 1–2）");
    }
    if (overWorkDays.length > 2) fail("overWorkDays", "Over work 日 " + overWorkDays.length + "（理想 ≤2）");
    if (aiHasConsecutiveTrue_(overWorkFlags, 2)) fail("overWorkStreak", "Over work 連續 ≥2 日");
    if (aiConsecutiveTrueStreaks_(sleepFlags).maxStreak > 2) {
      fail("sleepStreak", "睡眠異常連續 >2 日");
    }
    if (meditateDays.length < 6) warn("meditateDays", "Meditate 日 " + meditateDays.length + "（理想 ≥6）");
    if (exerciseDays.length < 2 || exerciseDays.length > 5) {
      warn("exerciseDays", "運動≥30m 日 " + exerciseDays.length + "（理想 2–5）");
    }
  } else if (pType === "month") {
    if (switchFailDayCount > 4) fail("switchFailDays", "切換失靈日 " + switchFailDayCount + "（理想 ≤4）");
    if (overSocialConsecutive) fail("overSocialWeeks", "連續兩週 over social（socialDays>3）");
    if (aiHasConsecutiveTrue_(negFlags, 3)) fail("chaosStreak", "負面情緒／chaos 連續 ≥3 日");
    if (aiHasConsecutiveTrue_(sleepFlags, 2)) fail("sleepStreak", "睡眠異常連續 ≥2 日");
    if (aiHasConsecutiveTrue_(work6Flags, 2)) fail("work6Streak", "Work≥6h 連續 ≥2 日");
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
    // approximate overwork from overload+critical if present
    var ow =
      (tier.overload || 0) + (tier.critical || 0);
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

  var kpis = {
    periodType: pType,
    periodKey: stats.periodKey,
    enabledSections: sections,
    periodNotes: String(pcfg.notes || ""),
    targets: {
      week: {
        cognitiveLockMax: 2,
        switchFailMax: 2,
        socialDaysMax: 2,
        vacationDaysIdeal: "1-2",
        overWorkDaysMax: 2,
        overWorkNoConsecutive: true,
        sleepAnomalyMaxConsecutive: 2,
        meditateDaysMin: 6,
        exerciseDaysIdeal: "2-5",
      },
      month: {
        switchFailDaysMax: 4,
        overSocialNoTwoConsecutiveWeeks: true,
        chaosMaxConsecutiveDays: 2,
        sleepAnomalyMaxConsecutive: 1,
        work6hNoConsecutive: true,
        tradingOver2hNoConsecutive: true,
      },
    },
    summary: {
      cognitiveLockCount: cognitiveLockCount,
      switchFailCount: switchFailCount,
      switchFailDayCount: switchFailDayCount,
      socialDays: socialDays.length,
      vacationDays: vacationDays.length,
      overWorkDays: overWorkDays.length,
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
      overWorkMaxStreak: aiConsecutiveTrueStreaks_(overWorkFlags).maxStreak,
      work6hMaxStreak: aiConsecutiveTrueStreaks_(work6Flags).maxStreak,
      tradingOver2hMaxStreak: aiConsecutiveTrueStreaks_(tradeFlags).maxStreak,
      overSocialConsecutiveWeeks: overSocialConsecutive,
    },
    lists: {
      vacationDays: vacationDays.map(function (x) {
        return x.ymd;
      }),
      overWorkDays: overWorkDays.map(function (x) {
        return { ymd: x.ymd, workHours: x.workHours };
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
      ? "week_kpis_checklist"
      : pType === "month"
        ? "month_kpis_checklist"
        : pType === "quarter"
          ? "quarter_weekly_trends"
          : pType === "year"
            ? "year_monthly_trends"
            : "custom_kpis";
  stats.reportLensNote =
    "只寫 enabledSections=true 嘅章節。合格門檻用 kpis.passFail／kpis.targets，唔好自創。負面情緒必帶 context72h。術語首次出現附 termGlossary 定義。";
  // extend glossary
  stats.termGlossary = stats.termGlossary || {};
  stats.termGlossary["睡眠不足"] = "當日 Sleeping 合計 < 6 小時（有打卡先計）";
  stats.termGlossary["睡眠過龍"] = "當日 Sleeping 合計 > 10 小時";
  stats.termGlossary["OverWork"] = "當日 Work Group > 4 小時";
  stats.termGlossary["運動日"] =
    "gyming／hiking／yogaing／running 等合計 ≥ 30 分鐘；Photoing 暫時唔計，除非 remark 注明高強度";
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

/** 情緒 brief：同一 wake-day 負面類最多一封 */
function maybeSendEmotionBrief_(state, ev, keywordsHit) {
  if (!ev || !keywordsHit || !keywordsHit.length) return { sent: false };
  var t0 = aiParseStartMs_(ev.start);
  if (isNaN(t0)) return { sent: false, error: "bad_start" };
  var wakeYmd = aiYmdLocal_(aiWakeDayStartMs_(t0));
  var props = PropertiesService.getScriptProperties();
  var dedupeKey = "emotionBrief:" + wakeYmd + ":neg";
  if (props.getProperty(dedupeKey)) {
    return { sent: false, already: true, wakeYmd: wakeYmd };
  }
  var act = aiActivityName_(state, ev.activityId) || "(unknown)";
  var ctx = aiContext72hBefore_(state, t0);
  var lines = [];
  lines.push("[Time Stat] Emotion brief — " + wakeYmd);
  lines.push("觸發字眼：" + keywordsHit.join(", "));
  lines.push("事件：" + aiLocalDateTimeStr_(t0) + " · " + act);
  lines.push("Remark：" + String(ev.remark || "").trim());
  lines.push("");
  lines.push("過去 72 小時摘要：");
  for (var i = 0; i < ctx.length; i++) {
    var c = ctx[i];
    lines.push(
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
  if (!ctx.length) lines.push("- （72h 內無其他事件）");
  try {
    mailAiReportToAllowed_("[Time Stat] Emotion brief " + wakeYmd, lines.join("\n"));
    props.setProperty(dedupeKey, String(Date.now()));
    return { sent: true, wakeYmd: wakeYmd };
  } catch (e) {
    return { sent: false, error: String(e && e.message ? e.message : e) };
  }
}

/**
 * 對比 prev／next state，掃新增或 remark 變更嘅負面情緒。
 */
function scanIncomingForEmotionBriefs_(prevState, nextState) {
  if (!nextState || !nextState.events) return { scanned: 0, sent: 0 };
  var kw = ((getAiPromptConfig_() || {}).emotionKeywords) || aiDefaultEmotionKeywords_();
  var neg = kw.negative || [];
  var prevMap = {};
  var prevEv = (prevState && prevState.events) || [];
  for (var i = 0; i < prevEv.length; i++) {
    var id = String(prevEv[i].id || prevEv[i].start || i);
    prevMap[id] = String(prevEv[i].remark || "");
  }
  var sent = 0;
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
    var r = maybeSendEmotionBrief_(nextState, ev, hits);
    if (r && r.sent) sent++;
  }
  return { scanned: scanned, sent: sent };
}
