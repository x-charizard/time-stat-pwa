/**
 * Time Stat AI — Energy Model 7.2 精簡重播（俾日型／Fatigue 切換 KPI 用）
 * 同 PWA app.js 常數對齊；語義 score 用 remark heuristic（唔打 live API）。
 */

var AI_E_SF_CAP = 1000;
var AI_E_DF_CAP = 1000;
var AI_E_FLOOR = -500;
var AI_E_HIGH = 4.17;
var AI_E_MED = 2.5;
var AI_E_LOW = AI_E_HIGH * 0.05;
var AI_E_SF_RECOVER = 5.0;
var AI_E_FATIGUE_PER_H = 0.15;
var AI_E_SLEEP_FATIGUE_RESET = 360;
var AI_E_SLEEP_DF_FULL = 480;
var AI_E_SLEEP_SF_FULL = 180;
var AI_E_SLEEP_P = 1.5;
var AI_E_SLEEP_P_FRAG = 1.1;
var AI_E_RECOVER_MIN = 15;
var AI_E_SOCIAL_EFF = 120;
var AI_E_SOCIAL_SAT = 240;
var AI_E_SOCIAL_DF_BASE = 0.25;
var AI_E_SOCIAL_DF_K = Math.LN2 / 60;
var AI_E_MIN_DF_CAP = 100;
var AI_E_ORPHAN_GAP = 4 * 60;
var AI_E_ORPHAN_CAP = 90;
var AI_E_MAX_NON_SLEEP = 4 * 60;

var AI_E_HIGH_KEYS = {
  trading: 1,
  "trading practice": 1,
  programming: 1,
  timing: 1,
  financing: 1,
  web: 1,
  webing: 1,
  "web development": 1,
  system: 1,
  systeming: 1,
  "system development": 1,
  apping: 1,
  "app development": 1,
};
var AI_E_MED_KEYS = {
  reviewing: 1,
  planning: 1,
  "trading planning": 1,
  "mind mapping": 1,
  mindmapping: 1,
  code: 1,
  obsidianing: 1,
  notioning: 1,
  reading: 1,
  photoing: 1,
  photography: 1,
  "photo editing": 1,
  photoediting: 1,
};
var AI_E_LOW_KEYS = {
  editing: 1,
  transporting: 1,
};
var AI_E_RECOVER_KEYS = {
  meditating: 1,
  gyming: 1,
  yogaing: 1,
  showering: 1,
  resting: 1,
  walking: 1,
  hiking: 1,
  fooding: 1,
};
var AI_E_SOCIAL_KEYS = { friending: 1, familying: 1, socialing: 1 };

function aiEClampSf_(sf) {
  return Math.max(AI_E_FLOOR, Math.min(AI_E_SF_CAP, sf));
}
function aiEClampDf_(df, cap) {
  var c = cap != null ? cap : AI_E_DF_CAP;
  return Math.max(AI_E_FLOOR, Math.min(c, df));
}
function aiEFatigue_(workMin) {
  return 1 + (Math.max(0, workMin) / 60) * AI_E_FATIGUE_PER_H;
}
function aiEFatigueTier_(workMin, tier) {
  var lin = aiEFatigue_(workMin);
  return tier === "medium" ? Math.pow(lin, 1.5) : lin;
}
function aiESfDebtMult_(df) {
  if (df > 1e-9) return 1;
  return 5 + Math.floor(Math.abs(df) / 100);
}
function aiESleepCurve_(m, power, base, fullMin, cap) {
  var mm = Math.max(0, m);
  if (mm <= 0) return 0;
  var p = power != null ? power : AI_E_SLEEP_P;
  var b = base != null ? base : 1;
  var full = fullMin > 0 ? fullMin : AI_E_SLEEP_DF_FULL;
  var c = cap != null ? cap : AI_E_DF_CAP;
  return Math.pow(mm / full, p) * c * b;
}
function aiESentimentRecover_(score) {
  var s = Number(score);
  if (!isFinite(s)) return 1;
  return Math.max(0.4, Math.min(1.8, 1 + s * 0.25));
}
function aiESentimentDrain_(score) {
  var s = Number(score);
  if (!isFinite(s)) return 1;
  return Math.max(0.4, Math.min(1.8, 1 - s * 0.25));
}

function aiEHeuristicSem_(ev) {
  var blob = String((ev && ev.remark) || "").toLowerCase();
  var score = 0;
  var sleep_base = 1;
  var is_fragmented = false;
  if (/(累|exhausted|tired|burnout|焦慮|anxiety|sad|哭|崩潰|地獄|痛|病|belly|sick|肚瀉|肚痛)/i.test(blob)) {
    score = -1.5;
  } else if (/(開心|grateful|平靜|calm|energized|充能|好瞓|rested)/i.test(blob)) {
    score = 1.2;
  }
  if (/(斷續|fragment|醒咗|insomnia|失眠|淺睡)/i.test(blob)) {
    is_fragmented = true;
    sleep_base = 0.7;
  } else if (/(深睡|deep\s*sleep|好瞓|solid\s*sleep)/i.test(blob)) {
    sleep_base = 1.3;
  }
  return { score: score, sleep_base: sleep_base, is_fragmented: is_fragmented };
}

function aiEParseDisruptMin_(ev) {
  var remark = String((ev && ev.remark) || "");
  var m = remark.match(/(?:for|約|大概|近)\s*(\d{1,3})\s*(?:m|min|mins|minutes|分鐘|分)\b/i);
  if (!m) m = remark.match(/(\d{1,3})\s*(?:m|min|mins|minutes|分鐘|分)\b/i);
  if (!m) return 0;
  var n = parseInt(m[1], 10);
  return isFinite(n) && n > 0 ? Math.min(24 * 60, n) : 0;
}

function aiESleepRecoverMult_(sem, durMin, ev) {
  var scoreMult = aiESentimentRecover_(sem && sem.score);
  var dur = Math.max(0, Number(durMin) || 0);
  var bad = aiEParseDisruptMin_(ev);
  var mult;
  if (bad > 0 && dur > 0) {
    var w = Math.min(1, bad / dur);
    mult = (1 - w) * 1 + w * scoreMult;
  } else {
    mult = scoreMult;
  }
  return Math.max(0.7, Math.min(1.8, mult));
}

function aiETier_(ev, actKey) {
  var key = String(actKey || "").toLowerCase();
  var blob = [
    String((ev && ev.remark) || ""),
    String((ev && ev.group) || ""),
    String((ev && ev.sub) || ""),
    String((ev && ev.category) || ""),
    String((ev && ev.project) || ""),
  ]
    .join(" ")
    .toLowerCase();

  if (key === "sleeping") return "sleep";
  if (AI_E_SOCIAL_KEYS[key]) return "social";

  // Xavier Li Photography（activity 或 sub／project）→ Work Medium
  if (
    key === "photoing" ||
    key === "photography" ||
    /xavier\s*li\s*photography/.test(blob)
  ) {
    return "medium";
  }

  // Aiing：Remark 決定 Rest(recover) 定 Work(medium)
  if (key === "aiing" || key === "ai") {
    if (
      /(hea|傾偈|傾計|聊天|吹水|玩|娛樂|輕鬆|休息|rest|chill|casual|meme|閒聊)/i.test(blob)
    ) {
      return "recover";
    }
    // 預設／有工作字 → Work M
    return "medium";
  }

  if (key === "reading") {
    if (/(小說|novel|fiction|harry\s*potter|漫畫|comic)/i.test(blob)) return "recover";
    return "medium";
  }
  if (AI_E_RECOVER_KEYS[key]) return "recover";
  if (AI_E_HIGH_KEYS[key]) return "high";
  if (AI_E_MED_KEYS[key]) return "medium";
  if (AI_E_LOW_KEYS[key]) return "low";
  return "none";
}

function aiEClassifyDayType_(wakeDf, endDf, dfDrained) {
  var drain = Math.max(0, Number(dfDrained) || 0);
  var end = Number(endDf);
  if (!isFinite(end)) end = 0;
  // 已取消 Critical；Overload：結束 DF < 0 或 DF 扣減 > 1000
  if (end < 0 || drain > 1000) return "overload";
  if (drain < 300 && end > 500) return "vacation";
  if (end > 0 && drain >= 700) return "idealFocus";
  if (end > 0 && drain < 700) return "hea";
  return "hea";
}

function aiEClassifyStart_(wakeDf) {
  var w = Number(wakeDf);
  if (!isFinite(w)) return "unknown";
  if (w >= AI_E_DF_CAP - 0.5) return "goodStart";
  if (w >= 700) return "moderateStart";
  if (w > 500) return "badStart";
  return "terribleStart";
}

function aiENewState_() {
  return {
    sf: AI_E_SF_CAP,
    df: AI_E_DF_CAP,
    dfMaxCap: AI_E_DF_CAP,
    dailyWorkMin: 0,
    socialMin: 0,
    sleepTotalMin: 0,
    sleepPower: AI_E_SLEEP_P,
    sleepBase: 1,
    sessionDrainMult: 1,
    sessionRecoverMult: 1,
    lastTier: "none",
    pendingWakeDfCap: false,
    fatigueResetThisSleep: false,
    recoverScrubAcc: 0,
    wakeKey: "",
    dayDfDrained: 0,
    dayWakeDf: AI_E_DF_CAP,
    dayEndDf: AI_E_DF_CAP,
    dayHighEndFatigues: [],
    dayFatigueDrops: [],
    _dayYmd: "",
    _wakeDfPending: false,
  };
}

function aiEWakeKey_(ms) {
  if (typeof aiWakeDayStartMs_ === "function") return aiYmdLocal_(aiWakeDayStartMs_(ms));
  var d = new Date(ms);
  var wake = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 3, 0, 0, 0);
  if (ms < wake.getTime()) wake = new Date(wake.getTime() - 86400000);
  return aiYmdLocal_(wake.getTime());
}

function aiEApplyWakeCap_(st) {
  var startDf = st.df;
  // Good／Terrible Start：用「真正醒嚟」嗰刻 DF（訓完回復後），唔好用 03:00 跨日中途值
  if (st._wakeDfPending || st.dayWakeDf == null || !isFinite(Number(st.dayWakeDf))) {
    st.dayWakeDf = startDf;
    st._wakeDfPending = false;
  }
  if (startDf < AI_E_DF_CAP - 1e-9) {
    if (startDf > 0) {
      st.dfMaxCap = Math.max(AI_E_MIN_DF_CAP, Math.min(AI_E_DF_CAP, Math.round(startDf)));
    } else {
      st.dfMaxCap = Math.max(
        AI_E_MIN_DF_CAP,
        Math.min(AI_E_DF_CAP, Math.round(AI_E_DF_CAP + startDf))
      );
    }
  } else {
    st.dfMaxCap = AI_E_DF_CAP;
  }
  st.df = aiEClampDf_(st.df, st.dfMaxCap);
  st.pendingWakeDfCap = false;
}

function aiESyncWake_(st, t0, skipCap) {
  var wk = aiEWakeKey_(t0);
  if (st.wakeKey === wk) return;
  if (st.wakeKey) {
    if (skipCap) {
      st.dfMaxCap = AI_E_DF_CAP;
      st.pendingWakeDfCap = true;
      st.df = aiEClampDf_(st.df, AI_E_DF_CAP);
    } else {
      aiEApplyWakeCap_(st);
    }
    st.socialMin = 0;
  }
  st.wakeKey = wk;
  st.dayDfDrained = 0;
  st.dayHighEndFatigues = [];
  st.dayFatigueDrops = [];
  st._dayYmd = wk;
  if (skipCap) {
    // 瞓緊過 03:00：未醒，唔好當 Start
    st.dayWakeDf = null;
    st._wakeDfPending = true;
  } else {
    st.dayWakeDf = st.df;
    st._wakeDfPending = false;
  }
}

function aiENoteDfDrain_(st, loss) {
  if (loss > 0) st.dayDfDrained += loss;
}

function aiEApplySleepMin_(st, step) {
  var s = Math.max(0, step);
  if (s <= 0) return;
  var prev = st.sleepTotalMin;
  var prevDf = aiESleepCurve_(prev, st.sleepPower, st.sleepBase, AI_E_SLEEP_DF_FULL, AI_E_DF_CAP);
  var prevSf = aiESleepCurve_(prev, st.sleepPower, st.sleepBase, AI_E_SLEEP_SF_FULL, AI_E_SF_CAP);
  st.sleepTotalMin += s;
  var nextDf = aiESleepCurve_(st.sleepTotalMin, st.sleepPower, st.sleepBase, AI_E_SLEEP_DF_FULL, AI_E_DF_CAP);
  var nextSf = aiESleepCurve_(st.sleepTotalMin, st.sleepPower, st.sleepBase, AI_E_SLEEP_SF_FULL, AI_E_SF_CAP);
  var dfGain = (nextDf - prevDf) * st.sessionRecoverMult;
  var sfGain = (nextSf - prevSf) * st.sessionRecoverMult;
  if (dfGain > 0) st.df = aiEClampDf_(st.df + dfGain, AI_E_DF_CAP);
  if (sfGain > 0) st.sf = aiEClampSf_(st.sf + sfGain);
  if (st.sleepTotalMin >= AI_E_SLEEP_FATIGUE_RESET && !st.fatigueResetThisSleep) {
    var beforeF = aiEFatigue_(st.dailyWorkMin);
    st.dailyWorkMin = 0;
    st.fatigueResetThisSleep = true;
    var afterF = aiEFatigue_(st.dailyWorkMin);
    if (beforeF > afterF + 1e-9) {
      st.dayFatigueDrops.push({
        kind: "sleepReset",
        from: Math.round(beforeF * 100) / 100,
        to: Math.round(afterF * 100) / 100,
        drop: Math.round((beforeF - afterF) * 100) / 100,
      });
    }
  }
}

function aiEApplyRecoverMin_(st, step, segFocus) {
  var s = Math.max(0, step);
  if (s <= 0) return;
  var rate = AI_E_SF_RECOVER * st.sessionRecoverMult;
  if (segFocus < AI_E_RECOVER_MIN) rate *= 0.5;
  st.sf = aiEClampSf_(st.sf + rate * s);
  if (segFocus >= AI_E_RECOVER_MIN) {
    var beforeF = aiEFatigue_(st.dailyWorkMin);
    st.recoverScrubAcc = (st.recoverScrubAcc || 0) + s;
    while (st.recoverScrubAcc >= 2 - 1e-9) {
      st.recoverScrubAcc -= 2;
      st.dailyWorkMin = Math.max(0, st.dailyWorkMin - 1);
    }
    var afterF = aiEFatigue_(st.dailyWorkMin);
    if (beforeF > afterF + 1e-9) {
      st.dayFatigueDrops.push({
        kind: "scrub",
        from: Math.round(beforeF * 100) / 100,
        to: Math.round(afterF * 100) / 100,
        drop: Math.round((beforeF - afterF) * 100) / 100,
      });
    }
  }
}

function aiEApplySocialDfOver_(st, before, step) {
  var left = Math.max(0, step);
  var cursor = Math.max(0, before);
  while (left > 1e-9) {
    if (cursor + 1e-9 < AI_E_SOCIAL_EFF) {
      var skip = Math.min(left, AI_E_SOCIAL_EFF - cursor);
      cursor += skip;
      left -= skip;
      continue;
    }
    var over = cursor - AI_E_SOCIAL_EFF;
    var rate = AI_E_SOCIAL_DF_BASE * Math.exp(AI_E_SOCIAL_DF_K * over);
    var slice = Math.min(left, 1);
    var loss = rate * slice;
    st.df = aiEClampDf_(st.df - loss, st.dfMaxCap);
    aiENoteDfDrain_(st, loss);
    cursor += slice;
    left -= slice;
  }
}

function aiEApplySocial_(st, step, score) {
  var s = Math.max(0, step);
  if (s <= 0) return;
  var scoreN = Number(score);
  if (!isFinite(scoreN)) scoreN = 0;
  scoreN = Math.max(-2, Math.min(2, scoreN));
  var before = st.socialMin;
  if (scoreN < 0) {
    var abs = Math.abs(scoreN);
    var sfLoss = AI_E_HIGH * abs * s;
    var dfLoss = 0.5 * abs * s;
    st.sf = aiEClampSf_(st.sf - sfLoss);
    st.df = aiEClampDf_(st.df - dfLoss, st.dfMaxCap);
    aiENoteDfDrain_(st, dfLoss);
    st.socialMin += s;
    aiEApplySocialDfOver_(st, before, s);
    return;
  }
  st.socialMin += s;
  if (scoreN > 0) {
    var mult = 1 + scoreN;
    var left = s;
    var cursor = before;
    while (left > 1e-9) {
      var rate = 0;
      var slice = left;
      if (cursor < AI_E_SOCIAL_EFF) {
        rate = 5 * mult;
        slice = Math.min(left, AI_E_SOCIAL_EFF - cursor);
      } else if (cursor < AI_E_SOCIAL_SAT) {
        rate = 1 * mult;
        slice = Math.min(left, AI_E_SOCIAL_SAT - cursor);
      } else {
        rate = -2;
        slice = left;
      }
      if (rate > 0) st.sf = aiEClampSf_(st.sf + rate * slice);
      else if (rate < 0) st.sf = aiEClampSf_(st.sf + rate * slice);
      cursor += slice;
      left -= slice;
    }
  }
  aiEApplySocialDfOver_(st, before, s);
}

function aiEApplyWork_(st, tier, step) {
  var s = Math.max(0, step);
  if (s <= 0) return;
  var base = tier === "high" ? AI_E_HIGH : tier === "medium" ? AI_E_MED : tier === "low" ? AI_E_LOW : 0;
  if (base <= 0) return;
  var fat = aiEFatigueTier_(st.dailyWorkMin, tier);
  st.dailyWorkMin += s;
  var sent = st.sessionDrainMult;
  var normalLoss = base * fat * sent * s;
  var isHigh = tier === "high";
  if (st.df > 1e-9) {
    if (st.df >= normalLoss) {
      st.df = aiEClampDf_(st.df - normalLoss, st.dfMaxCap);
      st.sf = aiEClampSf_(st.sf - normalLoss);
      aiENoteDfDrain_(st, normalLoss);
    } else {
      var used = st.df;
      var remRatio = (normalLoss - used) / normalLoss;
      st.df = 0;
      aiENoteDfDrain_(st, used);
      var debtMult = aiESfDebtMult_(0);
      st.sf = aiEClampSf_(st.sf - used - base * debtMult * sent * s * remRatio);
      if (isHigh) {
        var neg = normalLoss - used;
        st.df = aiEClampDf_(-neg, st.dfMaxCap);
        aiENoteDfDrain_(st, neg);
      }
    }
  } else {
    var dm = aiESfDebtMult_(st.df);
    st.sf = aiEClampSf_(st.sf - base * dm * sent * s);
    if (isHigh) {
      st.df = aiEClampDf_(st.df - normalLoss, st.dfMaxCap);
      aiENoteDfDrain_(st, normalLoss);
    }
  }
}

function aiECreditMin_(ev, tier, t0, nextMs, endAt, nowMs) {
  var t1 = nextMs != null ? nextMs : nowMs;
  if (endAt != null && t1 > endAt) t1 = endAt;
  var durationMin = Math.max(0, (t1 - t0) / 60000);
  if (tier !== "sleep") {
    var gapMin = nextMs != null ? (nextMs - t0) / 60000 : durationMin;
    var orphan = gapMin > AI_E_ORPHAN_GAP || nextMs == null;
    if (orphan && durationMin > AI_E_ORPHAN_CAP) durationMin = AI_E_ORPHAN_CAP;
    if (durationMin > AI_E_MAX_NON_SLEEP) durationMin = AI_E_MAX_NON_SLEEP;
  }
  return durationMin;
}

function aiEFocusMin_(ev, durationMin) {
  var distract = Math.max(0, (Number(ev && ev.distractionSec) || 0) / 60);
  return Math.max(0, durationMin - distract);
}

function aiEApplySegment_(st, ev, actKey, durationMin, t0) {
  var tier = aiETier_(ev, actKey);
  var sem = aiEHeuristicSem_(ev);
  var focus = aiEFocusMin_(ev, durationMin);
  st.sessionDrainMult = aiESentimentDrain_(sem.score);
  st.sessionRecoverMult =
    tier === "sleep" ? aiESleepRecoverMult_(sem, durationMin, ev) : aiESentimentRecover_(sem.score);
  if (tier === "sleep") {
    st.sleepPower = sem.is_fragmented ? AI_E_SLEEP_P_FRAG : AI_E_SLEEP_P;
    st.sleepBase = sem.sleep_base != null ? sem.sleep_base : 1;
    if (st.lastTier !== "sleep") {
      st.sleepTotalMin = 0;
      st.fatigueResetThisSleep = false;
    }
  } else if (st.lastTier === "sleep" || st.pendingWakeDfCap) {
    aiEApplyWakeCap_(st);
  }
  var skipCap = tier === "sleep";
  aiESyncWake_(st, t0, skipCap);
  if (tier === "recover" || tier === "social") st.recoverScrubAcc = 0;
  st.lastTier = tier;

  var step = 5;
  var elapsed = 0;
  var left = durationMin;
  var wasHigh = false;
  while (left > 1e-9) {
    var s = Math.min(step, left);
    aiESyncWake_(st, t0 + elapsed * 60000, skipCap);
    if (tier === "sleep") {
      aiEApplySleepMin_(st, s);
    } else if (tier === "recover") {
      if (elapsed < focus) aiEApplyRecoverMin_(st, Math.min(s, focus - elapsed), focus);
    } else if (tier === "social") {
      if (elapsed < focus) aiEApplySocial_(st, Math.min(s, focus - elapsed), sem.score);
    } else if (tier === "high" || tier === "medium" || tier === "low") {
      aiEApplyWork_(st, tier, s);
      if (tier === "high") wasHigh = true;
    }
    elapsed += s;
    left -= s;
  }
  if (wasHigh) {
    st.dayHighEndFatigues.push(Math.round(aiEFatigue_(st.dailyWorkMin) * 100) / 100);
  }
  st.dayEndDf = st.df;
}

function aiEClassifyFatigueSwitch_(fat) {
  var f = Number(fat);
  if (!isFinite(f)) return "unknown";
  if (f <= 1.4) return "ideal";
  if (f < 1.6) return "good";
  return "poor";
}

/**
 * 重播能量，輸出每個清醒日嘅 DF／Start／Fatigue 切換指標。
 * lookbackDays：期前多播幾日以穩定 wake DF。
 */
function aiEnergyReplayDayMetrics_(state, fromMs, toMs, lookbackDays) {
  var list = ((state && state.events) || []).slice().sort(function (a, b) {
    return aiParseStartMs_(a.start) - aiParseStartMs_(b.start);
  });
  var nowMs = Date.now();
  var lb = (lookbackDays != null ? lookbackDays : 14) * 86400000;
  var startReplay = fromMs - lb;
  var st = aiENewState_();
  var byYmd = {};
  var lastSnapYmd = "";

  function flushDay_(ymd) {
    if (!ymd || byYmd[ymd]) return;
    var highEnds = st.dayHighEndFatigues || [];
    var switchCounts = { ideal: 0, good: 0, poor: 0 };
    for (var i = 0; i < highEnds.length; i++) {
      var c = aiEClassifyFatigueSwitch_(highEnds[i]);
      if (switchCounts[c] != null) switchCounts[c]++;
    }
    var drops = st.dayFatigueDrops || [];
    var totalDrop = 0;
    for (var j = 0; j < drops.length; j++) totalDrop += Number(drops[j].drop) || 0;
    var wakeDf =
      st.dayWakeDf == null || !isFinite(Number(st.dayWakeDf))
        ? null
        : Math.round(st.dayWakeDf * 10) / 10;
    var endDf = Math.round(st.dayEndDf * 10) / 10;
    var drain = Math.round(st.dayDfDrained * 10) / 10;
    byYmd[ymd] = {
      ymd: ymd,
      wakeDf: wakeDf,
      endDf: endDf,
      dfDrained: drain,
      dayType: aiEClassifyDayType_(wakeDf, endDf, drain),
      startQuality: wakeDf == null ? "unknown" : aiEClassifyStart_(wakeDf),
      highWorkEndFatigue: highEnds.slice(),
      fatigueSwitchCounts: switchCounts,
      fatigueDropEvents: drops.slice(0, 20),
      fatigueTotalDrop: Math.round(totalDrop * 100) / 100,
      fatigueDropEventCount: drops.length,
    };
  }

  for (var i = 0; i < list.length; i++) {
    var ev = list[i];
    var t0 = aiParseStartMs_(ev.start);
    if (isNaN(t0) || t0 < startReplay) continue;
    if (t0 > toMs) break;
    var nextMs = null;
    for (var j = i + 1; j < list.length; j++) {
      var tj = aiParseStartMs_(list[j].start);
      if (!isNaN(tj) && tj > t0) {
        nextMs = tj;
        break;
      }
    }
    var actName =
      typeof aiActivityName_ === "function" ? aiActivityName_(state, ev.activityId) : String(ev.activityId || "");
    var actKey = typeof aiNormKey_ === "function" ? aiNormKey_(actName) : String(actName || "").toLowerCase();
    var tier = aiETier_(ev, actKey);
    var endAt = t0 > toMs ? toMs : Math.min(toMs + 1, nextMs != null ? nextMs : nowMs);
    // credit through min(next, now) but snapshot days in range
    var durationMin = aiECreditMin_(ev, tier, t0, nextMs, null, nowMs);
    if (durationMin <= 0) continue;

    // 若跨日，分段以便每日 snapshot
    var remaining = durationMin;
    var cursorT = t0;
    while (remaining > 1e-9) {
      var wk = aiEWakeKey_(cursorT);
      var wakeEnd =
        typeof aiWakeDayStartMs_ === "function"
          ? aiWakeDayStartMs_(cursorT) + 86400000
          : cursorT + 86400000;
      var minsToWakeEnd = Math.max(0.001, (wakeEnd - cursorT) / 60000);
      var slice = Math.min(remaining, minsToWakeEnd);
      if (st._dayYmd && st._dayYmd !== wk) flushDay_(st._dayYmd);
      aiEApplySegment_(st, ev, actKey, slice, cursorT);
      lastSnapYmd = st._dayYmd || wk;
      cursorT += slice * 60000;
      remaining -= slice;
    }
  }
  if (lastSnapYmd) flushDay_(lastSnapYmd);

  // 確保期內每日都有 row（無活動日用最後狀態粗略）
  var dayCursor =
    typeof aiWakeDayStartMs_ === "function" ? aiWakeDayStartMs_(fromMs) : fromMs;
  var guard = 0;
  while (dayCursor <= toMs && guard < 400) {
    guard++;
    var ymd = aiYmdLocal_(dayCursor);
    if (!byYmd[ymd]) {
      byYmd[ymd] = {
        ymd: ymd,
        wakeDf: null,
        endDf: null,
        dfDrained: 0,
        dayType: "vacation",
        startQuality: "unknown",
        highWorkEndFatigue: [],
        fatigueSwitchCounts: { ideal: 0, good: 0, poor: 0 },
        fatigueDropEvents: [],
        fatigueTotalDrop: 0,
        fatigueDropEventCount: 0,
        noData: true,
      };
    }
    dayCursor += 86400000;
  }

  var days = [];
  for (var k in byYmd) {
    if (Object.prototype.hasOwnProperty.call(byYmd, k)) {
      var row = byYmd[k];
      var rowMs = new Date(k + "T12:00:00").getTime();
      if (rowMs >= fromMs - 86400000 && rowMs <= toMs + 86400000) days.push(row);
    }
  }
  days.sort(function (a, b) {
    return a.ymd < b.ymd ? -1 : a.ymd > b.ymd ? 1 : 0;
  });
  return { byYmd: byYmd, days: days };
}
