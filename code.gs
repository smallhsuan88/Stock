/** =========================
 *  TW MVP - Stable Version (with prev fields)
 *  - Source: TWSE STOCK_DAY (monthly)
 *  - Output: Raw_TW + Signals_TW + Log
 *  - Indicators: MA10/MA20/MA60, VOL_MA20(prev), RSI14, ADX14, High40(prev), Breakout
 *  - Breakout: close > high40_prev AND volume > vol_ma20(prev)
 * ========================= **/

const SHEET_CONFIG  = 'Config_TW';
const SHEET_RAW     = 'Raw_TW';
const SHEET_SIGNALS = 'Signals_TW';
const SHEET_LOG     = 'Log';

const LOOKBACK_MONTHS = 6;   // 若要 MA60 覆蓋率更高，可改 9 或 12
const HIGH_N   = 40;
const VOL_MA_N = 20;
const TW_MA10_N   = 10;
const TW_MA20_N   = 20;
const TW_MA60_N   = 60;
const RESLINE_EPS = 0.005;
const NECKLINE_TOLERANCE = 0.005;
const NECKLINE_EPS = 0.005;

const SLEEP_MS_EACH_TICKER = 250; // 避免被節流

const SIGNAL_HEADERS_BASE = [
  'market','ticker','date','close','volume',
  'ma10','ma20','ma60','ma10_slope','ma20_slope','ma60_slope',
  'vol_ma20','vol_ratio',
  'rsi14','rsi14_prev',
  'adx14','adx14_prev',
  'high40','high40_prev',
  'breakout',
  'break_resline60',
  'ma_bull_stack',
  'wk_kd_golden',
  'neckline60',
  'break_neckline60',
  'bottom_lead_hits',
  'bottom_lead_hits2',
  'bottom_lead_count',
  'bottom_lead2_count',
  'position_cap',
  'position_cap2'
];

const SIGNAL_HEADERS_WEEKLY = [
  'ma10w',
  'above_ma10w',
  'wk_macd_golden',
  'wk_ma10w_2w_up'
];

/** 第一次用：補齊表頭（不會刪你資料，只會在空表時加表頭；若表不存在會自動建立） */
function initSheets_TW() {
  const ss = SpreadsheetApp.getActive();

  const raw = getOrCreateSheet_(ss, SHEET_RAW);
  const sig = getOrCreateSheet_(ss, SHEET_SIGNALS);
  const log = getOrCreateSheet_(ss, SHEET_LOG);

  if (raw.getLastRow() === 0) {
    raw.appendRow(['ticker','date','open','high','low','close','volume']);
  }
  if (sig.getLastRow() === 0) {
    sig.appendRow(buildSignalHeaders_([]));
  }
  if (log.getLastRow() === 0) {
    log.appendRow(['time','level','ticker','message']);
  }
}

/** 每天收盤後跑 */
function runTW_MVP() {
  const ss = SpreadsheetApp.getActive();

  const cfg = getOrCreateSheet_(ss, SHEET_CONFIG);
  const raw = getOrCreateSheet_(ss, SHEET_RAW);
  const sig = getOrCreateSheet_(ss, SHEET_SIGNALS);
  getOrCreateSheet_(ss, SHEET_LOG); // 確保 Log 存在

  const tickers = readTickers_(cfg);
  if (tickers.length === 0) {
    return log_(ss, 'WARN', '', 'No active tickers in Config_TW');
  }

  // 重新產出 Signals（MVP 穩）
  const existingHeader = sig.getLastRow() > 0
    ? sig.getRange(1, 1, 1, sig.getLastColumn()).getValues()[0].filter(v => String(v).trim() !== '')
    : [];
  const sigHeader = buildSignalHeaders_(existingHeader);
  sig.clearContents();
  sig.appendRow(sigHeader);

  const started = new Date();
  log_(ss, 'INFO', '', `Run started. tickers=${tickers.length}, months=${LOOKBACK_MONTHS}`);

  tickers.forEach((ticker) => {
    try {
      const rows = fetchTWSEStockDayMonths_(ticker, LOOKBACK_MONTHS); // asc by date
      if (!rows.length) {
        log_(ss, 'WARN', ticker, 'No rows fetched (maybe ETF/OTC or API empty)');
        return;
      }

      // Raw：先刪 ticker 再寫
      replaceRawForTicker_(raw, ticker, rows);

      const calc = calcSignalsFromRows_(rows);

      sig.appendRow(buildSignalRow_(sigHeader, calc, ticker));

      log_(ss, 'INFO', ticker,
        `OK ${calc.date} breakout=${calc.breakout} ` +
        `vol_ratio=${num_(calc.vol_ratio,2)} rsi=${num_(calc.rsi14,1)} adx=${num_(calc.adx14,1)}`
      );
      if (calc.infoLogs && calc.infoLogs.length) {
        calc.infoLogs.forEach(msg => log_(ss, 'INFO', ticker, msg));
      }
    } catch (e) {
      log_(ss, 'ERROR', ticker, String(e && e.stack ? e.stack : e));
    } finally {
      Utilities.sleep(SLEEP_MS_EACH_TICKER);
    }
  });

  const secs = Math.round((new Date().getTime() - started.getTime()) / 1000);
  log_(ss, 'INFO', '', `Run finished. seconds=${secs}`);
}

/** 讀取 Config_TW：A=ticker, B=active(TRUE/FALSE) */
function readTickers_(cfgSheet) {
  const values = cfgSheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const t = String(values[i][0] || '').trim();
    if (!t) continue;
    const active = values[i][1];
    // 若 B 欄空白，視為啟用；若明確 FALSE 才跳過
    if (active === false) continue;
    out.push(t);
  }
  return out;
}

/** 抓 TWSE 月資料：https://www.twse.com.tw/exchangeReport/STOCK_DAY */
function fetchTWSEStockDayMonths_(ticker, monthsBack) {
  const now = new Date();
  let all = [];

  for (let m = 0; m < monthsBack; m++) {
    const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
    const y = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const url = `https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=${y}${mm}01&stockNo=${ticker}`;

    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    const code = resp.getResponseCode();
    if (code !== 200) continue;

    const text = resp.getContentText();
    let json;
    try {
      json = JSON.parse(text);
    } catch (_) {
      continue;
    }

    if (!json || !json.data || !Array.isArray(json.data)) continue;

    json.data.forEach(arr => {
      const dateISO = rocToISO_(arr[0]);
      if (!dateISO) return;

      const volume = safeVolume_(arr);
      const open  = parseNum_(arr[3]);
      const high  = parseNum_(arr[4]);
      const low   = parseNum_(arr[5]);
      const close = parseNum_(arr[6]);

      if (!isFinite(close) || close <= 0) return;

      all.push({ date: dateISO, open, high, low, close, volume });
    });
  }

  // 排序 + 去重（同日只留一筆）
  all.sort((a,b) => a.date.localeCompare(b.date));
  const dedup = [];
  const seen = new Set();
  for (const r of all) {
    if (seen.has(r.date)) continue;
    seen.add(r.date);
    dedup.push(r);
  }
  return dedup;
}

/** volume 更安全：優先用 arr[1]（成交股數），若異常再 fallback */
function safeVolume_(arr) {
  const v1 = parseNum_(arr[1]);
  if (v1 > 0) return v1;
  for (let i = 0; i < arr.length; i++) {
    const v = parseNum_(arr[i]);
    if (v > 1000) return v;
  }
  return 0;
}

/** Raw：刪掉該 ticker 舊資料，再寫入新資料 */
function replaceRawForTicker_(rawSheet, ticker, rows) {
  const lastRow = rawSheet.getLastRow();
  if (lastRow >= 2) {
    const range = rawSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // ticker col
    const toDelete = [];
    for (let i = 0; i < range.length; i++) {
      if (String(range[i][0]).trim() === ticker) toDelete.push(i + 2);
    }
    for (let i = toDelete.length - 1; i >= 0; i--) {
      rawSheet.deleteRow(toDelete[i]);
    }
  }

  const out = rows.map(r => [ticker, r.date, r.open, r.high, r.low, r.close, r.volume]);
  rawSheet.getRange(rawSheet.getLastRow() + 1, 1, out.length, 7).setValues(out);
}

/** 計算最新一日 signals（含 prev 欄位） */
function calcSignalsFromRows_(rows) {
  const n = rows.length;
  const last = rows[n - 1];
  const infoLogs = [];

  const closes = rows.map(r => r.close);
  const vols   = rows.map(r => r.volume);
  const highs  = rows.map(r => r.high);
  const lows   = rows.map(r => r.low);

  // 今日均線（含今天收盤）
  const ma10 = sma_(closes, TW_MA10_N);
  const ma20 = sma_(closes, TW_MA20_N);
  const ma60 = sma_(closes, TW_MA60_N);

  // 昨日均線（不含今天，用於 slope）
  const closesPrev = closes.slice(0, -1);
  const highsPrev  = highs.slice(0, -1);
  const lowsPrev   = lows.slice(0, -1);
  const volsPrev   = vols.slice(0, -1);

  const ma10_prev = sma_(closesPrev, TW_MA10_N);
  const ma20_prev = sma_(closesPrev, TW_MA20_N);
  const ma60_prev = sma_(closesPrev, TW_MA60_N);

  // 量均：用「前 20 日均量」當基準（不含今天）
  const vol_ma20 = sma_(volsPrev, VOL_MA_N);
  const vol_ratio = (isFiniteNum_(vol_ma20) && vol_ma20 > 0) ? (last.volume / vol_ma20) : '';

  // high40：顯示用（含今天）
  const high40 = Math.max(...highs.slice(Math.max(0, n - HIGH_N)));

  // high40_prev：突破基準（不含今天）
  const nPrev = highsPrev.length;
  const high40_prev = (nPrev >= 1)
    ? Math.max(...highsPrev.slice(Math.max(0, nPrev - HIGH_N)))
    : '';

  // RSI / ADX（今日）
  const rsi14 = rsiWilder_(closes, 14);
  const adx14 = adxWilder_(highs, lows, closes, 14);

  // RSI / ADX（昨日）
  const rsi14_prev = rsiWilder_(closesPrev, 14);
  const adx14_prev = adxWilder_(highsPrev, lowsPrev, closesPrev, 14);

  // Breakout（用 high40_prev + vol_ma20(prev)）
  const breakout =
    (isFiniteNum_(high40_prev) ? (last.close > high40_prev) : false) &&
    (isFiniteNum_(vol_ma20) ? (last.volume > vol_ma20) : false);

  // ma_bull_stack：close > ma10 > ma20 且 close > ma20
  let ma_bull_stack = '';
  if (isFiniteNum_(ma10) && isFiniteNum_(ma20)) {
    ma_bull_stack = (last.close > ma10 && last.close > ma20 && ma10 > ma20) ? 'TRUE' : 'FALSE';
  }

  // break_resline60：Swing High 壓力線突破
  const resline = calcBreakResline60_(highs, last.close, n, infoLogs);
  const break_resline60 = resline.status;

  // neckline60 / break_neckline60：近 60 日最多碰觸價位 + 突破
  const neckline = calcNeckline60_(closes, last.close, n, infoLogs);
  const neckline60 = neckline.neckline;
  const break_neckline60 = neckline.breakStatus;

  const weeks = toWeeklyBars_(rows);
  const wCloses = weeks.map(w => w.close);

  // 週 KD 金叉
  const kd = calcWeeklyKDG_(weeks, infoLogs);
  const wk_kd_golden = kd.status;

  // 週 MA10
  let ma10w = '';
  let ma10w_prev = '';
  if (wCloses.length >= 10) {
    ma10w = sma_(wCloses, 10);
    ma10w_prev = wCloses.length >= 11 ? sma_(wCloses.slice(0, -1), 10) : '';
  }

  const above_ma10w = (isFiniteNum_(ma10w) && weeks.length)
    ? (weeks[weeks.length - 1].close > ma10w ? 'TRUE' : 'FALSE')
    : '';

  let wk_ma10w_2w_up = '';
  if (isFiniteNum_(ma10w) && isFiniteNum_(ma10w_prev) && weeks.length >= 2) {
    const lastWeekClose = weeks[weeks.length - 1].close;
    const prevWeekClose = weeks[weeks.length - 2].close;
    wk_ma10w_2w_up = (lastWeekClose > ma10w && prevWeekClose > ma10w_prev) ? 'TRUE' : 'FALSE';
  }

  const macd = calcWeeklyMacd_(wCloses, infoLogs);
  const wk_macd_golden = macd.status;

  // 底部領先訊號彙總
  const conds = [];
  const hits1to5 = [];

  if (resline.available) {
    conds.push(resline.hit);
    if (resline.hit) hits1to5.push('1');
  }
  if (ma_bull_stack !== '') {
    conds.push(ma_bull_stack === 'TRUE');
    if (ma_bull_stack === 'TRUE') hits1to5.push('2');
  }
  if (kd.available) {
    conds.push(kd.hit);
    if (kd.hit) hits1to5.push('3');
  }
  if (neckline.available) {
    conds.push(neckline.hit);
    if (neckline.hit) hits1to5.push('4');
  }
  const cond5Available = isFiniteNum_(ma60);
  if (cond5Available) {
    const cond5 = last.close > ma60;
    conds.push(cond5);
    if (cond5) hits1to5.push('5');
  } else {
    infoLogs.push('ma60 unavailable for cond5 (need >= 60 closes)');
  }

  const weeklyHits = [];
  const weeklyAvailability = [];
  const pushWeekly = (flag, label) => {
    if (flag === 'TRUE') weeklyHits.push(label);
    if (flag === 'TRUE' || flag === 'FALSE') weeklyAvailability.push(true);
  };
  pushWeekly(above_ma10w, '6');
  pushWeekly(wk_kd_golden, '7');
  pushWeekly(wk_macd_golden, '8');
  pushWeekly(wk_ma10w_2w_up, '9');

  const hasWeeklyAvailability = weeklyAvailability.length > 0;
  const bottom_lead_hits2 = hasWeeklyAvailability ? (weeklyHits.length ? weeklyHits.join(',') : '') : '';
  const bottom_lead2_count = hasWeeklyAvailability ? String(weeklyHits.length) : '';

  const allHits = hits1to5.concat(weeklyHits);
  const anyCondAvailable = conds.length > 0 || hasWeeklyAvailability;
  const bottom_lead_hits = allHits.length > 0 ? allHits.join(',') : '';
  const bottom_lead_count = anyCondAvailable ? String(allHits.length) : '';

  let position_cap = '';
  if (conds.length > 0) {
    if (hits1to5.length >= 5) position_cap = '40%';
    else if (hits1to5.length >= 3) position_cap = '20%';
  }

  let position_cap2 = '';
  if (hasWeeklyAvailability) {
    if (weeklyHits.length >= 4) position_cap2 = '40%';
    else if (weeklyHits.length >= 2) position_cap2 = '20%';
  }

  return {
    date: last.date,
    close: last.close,
    volume: last.volume,

    ma10,
    ma20,
    ma60,
    ma10_slope: (isFiniteNum_(ma10) && isFiniteNum_(ma10_prev)) ? (ma10 - ma10_prev) : '',
    ma20_slope: (isFiniteNum_(ma20) && isFiniteNum_(ma20_prev)) ? (ma20 - ma20_prev) : '',
    ma60_slope: (isFiniteNum_(ma60) && isFiniteNum_(ma60_prev)) ? (ma60 - ma60_prev) : '',

    vol_ma20,
    vol_ratio,

    rsi14,
    rsi14_prev,

    adx14,
    adx14_prev,

    high40,
    high40_prev,

    breakout,

    break_resline60,
    ma_bull_stack,
    wk_kd_golden,
    neckline60,
    break_neckline60,
    bottom_lead_hits,
    bottom_lead_count,
    bottom_lead_hits2,
    bottom_lead2_count,
    position_cap,
    position_cap2,

    ma10w,
    above_ma10w,
    wk_macd_golden,
    wk_ma10w_2w_up,

    infoLogs
  };
}

/** ===== Indicators ===== */

function sma_(arr, win) {
  if (!arr || arr.length < win) return '';
  let sum = 0;
  for (let i = arr.length - win; i < arr.length; i++) sum += Number(arr[i]) || 0;
  return sum / win;
}

function rsiWilder_(closes, period) {
  if (!closes || closes.length <= period) return '';
  let gains = 0, losses = 0;

  for (let i = 1; i <= period; i++) {
    const chg = closes[i] - closes[i - 1];
    if (chg >= 0) gains += chg;
    else losses += -chg;
  }

  let avgGain = gains / period;
  let avgLoss = losses / period;

  for (let i = period + 1; i < closes.length; i++) {
    const chg = closes[i] - closes[i - 1];
    const gain = chg > 0 ? chg : 0;
    const loss = chg < 0 ? -chg : 0;
    avgGain = (avgGain * (period - 1) + gain) / period;
    avgLoss = (avgLoss * (period - 1) + loss) / period;
  }

  if (avgLoss === 0) return 100;
  const rs = avgGain / avgLoss;
  return 100 - (100 / (1 + rs));
}

function adxWilder_(highs, lows, closes, period) {
  const n = closes.length;
  if (!highs || !lows || !closes || n <= period + 1) return '';

  const tr = [];
  const pdm = [];
  const ndm = [];

  for (let i = 1; i < n; i++) {
    const upMove = highs[i] - highs[i - 1];
    const downMove = lows[i - 1] - lows[i];

    pdm.push((upMove > downMove && upMove > 0) ? upMove : 0);
    ndm.push((downMove > upMove && downMove > 0) ? downMove : 0);

    const tr1 = highs[i] - lows[i];
    const tr2 = Math.abs(highs[i] - closes[i - 1]);
    const tr3 = Math.abs(lows[i] - closes[i - 1]);
    tr.push(Math.max(tr1, tr2, tr3));
  }

  let smTR = 0, smPDM = 0, smNDM = 0;
  for (let i = 0; i < period; i++) {
    smTR += tr[i];
    smPDM += pdm[i];
    smNDM += ndm[i];
  }

  const dxArr = [];
  for (let i = period; i < tr.length; i++) {
    if (i > period) {
      smTR  = smTR  - (smTR  / period) + tr[i];
      smPDM = smPDM - (smPDM / period) + pdm[i];
      smNDM = smNDM - (smNDM / period) + ndm[i];
    }

    const pdi = smTR === 0 ? 0 : (100 * (smPDM / smTR));
    const ndi = smTR === 0 ? 0 : (100 * (smNDM / smTR));
    const dx  = (pdi + ndi) === 0 ? 0 : (100 * Math.abs(pdi - ndi) / (pdi + ndi));
    dxArr.push(dx);
  }

  if (dxArr.length < period) return '';

  let adx = 0;
  for (let i = 0; i < period; i++) adx += dxArr[i];
  adx = adx / period;

  for (let i = period; i < dxArr.length; i++) {
    adx = (adx * (period - 1) + dxArr[i]) / period;
  }

  return adx;
}

function calcMacdSeries_(closes, fast, slow, signal) {
  const n = closes.length;
  const macdLine = new Array(n).fill('');
  const signalLine = new Array(n).fill('');
  if (!closes || n < slow) return { macdLine, signalLine };

  const kFast = 2 / (fast + 1);
  const kSlow = 2 / (slow + 1);
  const kSignal = 2 / (signal + 1);

  let emaFast = null;
  let emaSlow = null;
  let signalEma = null;
  const macdHistory = [];

  for (let i = 0; i < n; i++) {
    const price = closes[i];

    if (i === fast - 1) {
      emaFast = average_(closes.slice(0, fast));
    } else if (i >= fast && emaFast !== null) {
      emaFast = price * kFast + emaFast * (1 - kFast);
    }

    if (i === slow - 1) {
      emaSlow = average_(closes.slice(0, slow));
    } else if (i >= slow && emaSlow !== null) {
      emaSlow = price * kSlow + emaSlow * (1 - kSlow);
    }

    if (emaFast !== null && emaSlow !== null) {
      const macd = emaFast - emaSlow;
      macdLine[i] = macd;
      macdHistory.push(macd);

      if (macdHistory.length === signal) {
        signalEma = average_(macdHistory.slice(macdHistory.length - signal));
      } else if (macdHistory.length > signal && signalEma !== null) {
        signalEma = macd * kSignal + signalEma * (1 - kSignal);
      }

      if (signalEma !== null) signalLine[i] = signalEma;
    }
  }

  return { macdLine, signalLine };
}

function calcBreakResline60_(highs, lastClose, n, infoLogs) {
  if (n < TW_MA60_N) {
    infoLogs.push('break_resline60 skipped: need >= 60 days');
    return { status: '', available: false, hit: false };
  }

  const start = Math.max(0, n - TW_MA60_N);
  const end = n - 1; // 不含今天
  const swingIdx = [];
  for (let i = start + 1; i < end; i++) {
    if (i + 1 >= highs.length) break;
    const hi = highs[i];
    if (hi > highs[i - 1] && hi > highs[i + 1]) {
      swingIdx.push(i);
    }
  }

  if (swingIdx.length < 2) {
    infoLogs.push('break_resline60 skipped: swing highs < 2 in last 60 days');
    return { status: '', available: false, hit: false };
  }

  const t2 = swingIdx[swingIdx.length - 1];
  const t1 = swingIdx[swingIdx.length - 2];
  const p2 = highs[t2];
  const p1 = highs[t1];

  if (t2 === t1) {
    infoLogs.push('break_resline60 skipped: swing high indices identical');
    return { status: '', available: false, hit: false };
  }

  const tNow = n - 1;
  const y = p1 + (p2 - p1) * (tNow - t1) / (t2 - t1);
  const hit = lastClose > y * (1 + RESLINE_EPS);
  return { status: hit ? 'TRUE' : 'FALSE', available: true, hit };
}

function calcNeckline60_(closes, lastClose, n, infoLogs) {
  if (n < TW_MA60_N) {
    infoLogs.push('neckline60 skipped: need >= 60 days');
    return { neckline: '', breakStatus: '', available: false, hit: false };
  }
  const start = Math.max(0, n - TW_MA60_N);
  const prices = closes.slice(start, n - 1); // 建議不含今天

  const clusters = [];
  prices.forEach(price => {
    if (!isFiniteNum_(price) || price <= 0) return;
    let matched = false;
    for (const c of clusters) {
      if (Math.abs(price - c.center) / c.center <= NECKLINE_TOLERANCE) {
        c.count += 1;
        c.sum += price;
        c.center = c.sum / c.count;
        matched = true;
        break;
      }
    }
    if (!matched) {
      clusters.push({ center: price, count: 1, sum: price });
    }
  });

  if (!clusters.length) {
    infoLogs.push('neckline60 skipped: clusters not formed');
    return { neckline: '', breakStatus: '', available: false, hit: false };
  }

  let best = clusters[0];
  for (let i = 1; i < clusters.length; i++) {
    if (clusters[i].count > best.count) best = clusters[i];
  }

  const neckline = best.center;
  const hit = lastClose > neckline * (1 + NECKLINE_EPS);

  return {
    neckline,
    breakStatus: hit ? 'TRUE' : 'FALSE',
    available: true,
    hit
  };
}

function toWeeklyBars_(rows) {
  if (!rows || !rows.length) return [];
  const weeks = [];
  rows.forEach(r => {
    const key = weekKey_(r.date);
    if (!weeks.length || weeks[weeks.length - 1].key !== key) {
      weeks.push({
        key,
        open: r.open,
        high: r.high,
        low: r.low,
        close: r.close
      });
    } else {
      const w = weeks[weeks.length - 1];
      w.high = Math.max(w.high, r.high);
      w.low = Math.min(w.low, r.low);
      w.close = r.close;
    }
  });
  return weeks;
}

function calcWeeklyKDG_(weeks, infoLogs) {
  if (!weeks || weeks.length === 0) return { status: '', available: false, hit: false };

  if (weeks.length < 12) {
    infoLogs.push(`wk_kd_golden skipped: weeks < 12 (have ${weeks.length})`);
    return { status: '', available: false, hit: false };
  }

  let K = 50;
  let D = 50;
  let prevK = null;
  let prevD = null;

  for (let i = 0; i < weeks.length; i++) {
    if (i < 8) continue; // 9 週才開始 RSV
    const window = weeks.slice(Math.max(0, i - 8), i + 1);
    const highs = window.map(w => w.high);
    const lows = window.map(w => w.low);
    const close = weeks[i].close;

    const highestHigh = Math.max(...highs);
    const lowestLow = Math.min(...lows);
    if (highestHigh === lowestLow) continue;
    const RSV = (close - lowestLow) / (highestHigh - lowestLow) * 100;
    K = (2 / 3) * K + (1 / 3) * RSV;
    D = (2 / 3) * D + (1 / 3) * K;

    if (i === weeks.length - 1) break;
    prevK = K;
    prevD = D;
  }

  if (prevK === null || prevD === null) {
    infoLogs.push('wk_kd_golden skipped: insufficient RSV history');
    return { status: '', available: false, hit: false };
  }

  const golden = (K > D) && (prevK <= prevD);
  return { status: golden ? 'TRUE' : 'FALSE', available: true, hit: golden };
}

function calcWeeklyMacd_(wCloses, infoLogs) {
  if (!wCloses || wCloses.length === 0) return { status: '', available: false, hit: false };
  if (wCloses.length < 35) {
    infoLogs.push(`wk_macd_golden skipped: weeks < 35 (have ${wCloses.length})`);
    return { status: '', available: false, hit: false };
  }

  const { macdLine, signalLine } = calcMacdSeries_(wCloses, 12, 26, 9);
  const lastIdx = findPrevFinitePairIndex_(macdLine, signalLine, macdLine.length);
  if (lastIdx < 0) {
    infoLogs.push('wk_macd_golden skipped: macd/signal unavailable');
    return { status: '', available: false, hit: false };
  }
  const prevIdx = findPrevFinitePairIndex_(macdLine, signalLine, lastIdx);
  if (prevIdx < 0) {
    infoLogs.push('wk_macd_golden skipped: prev macd/signal unavailable');
    return { status: '', available: false, hit: false };
  }

  const hit = macdLine[lastIdx] > signalLine[lastIdx] && macdLine[prevIdx] <= signalLine[prevIdx];
  return { status: hit ? 'TRUE' : 'FALSE', available: true, hit };
}

function weekKey_(dateStr) {
  const dt = new Date(`${dateStr}T00:00:00+08:00`);
  const day = dt.getDay(); // 0=Sun
  const diff = (day + 6) % 7; // Monday as start
  dt.setDate(dt.getDate() - diff);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** ===== Utils ===== */

function parseNum_(s) {
  const x = String(s ?? '').replace(/,/g,'').trim();
  const n = Number(x);
  return isFinite(n) ? n : 0;
}

function rocToISO_(roc) {
  if (!roc) return '';
  const parts = String(roc).split('/');
  if (parts.length !== 3) return '';
  const y = Number(parts[0]) + 1911;
  const m = parts[1].padStart(2,'0');
  const d = parts[2].padStart(2,'0');
  return `${y}-${m}-${d}`;
}

function isFiniteNum_(x) {
  return typeof x === 'number' && isFinite(x);
}

function num_(x, dp) {
  if (!isFiniteNum_(x)) return '';
  return x.toFixed(dp);
}

function average_(arr) {
  if (!arr || !arr.length) return null;
  const sum = arr.reduce((s, v) => s + (Number(v) || 0), 0);
  return sum / arr.length;
}

function findPrevFinitePairIndex_(arr1, arr2, startIdx) {
  for (let i = Math.min(startIdx - 1, arr1.length - 1); i >= 0; i--) {
    if (isFiniteNum_(arr1[i]) && isFiniteNum_(arr2[i])) return i;
  }
  return -1;
}

function buildSignalHeaders_(existingHeader) {
  const headerSet = new Set();
  const finalHeaders = [];

  (existingHeader || []).forEach(h => {
    const name = String(h || '').trim();
    if (!name) return;
    if (headerSet.has(name)) return;
    headerSet.add(name);
    finalHeaders.push(name);
  });

  const addField = (field) => {
    if (!headerSet.has(field)) {
      headerSet.add(field);
      finalHeaders.push(field);
    }
  };

  SIGNAL_HEADERS_BASE.forEach(addField);
  SIGNAL_HEADERS_WEEKLY.forEach(addField);

  return finalHeaders;
}

function buildSignalRow_(header, calc, ticker) {
  const values = {
    market: 'TW',
    ticker,
    date: calc.date,
    close: calc.close,
    volume: calc.volume,

    ma10: calc.ma10,
    ma20: calc.ma20,
    ma60: calc.ma60,
    ma10_slope: calc.ma10_slope,
    ma20_slope: calc.ma20_slope,
    ma60_slope: calc.ma60_slope,

    vol_ma20: calc.vol_ma20,
    vol_ratio: calc.vol_ratio,

    rsi14: calc.rsi14,
    rsi14_prev: calc.rsi14_prev,

    adx14: calc.adx14,
    adx14_prev: calc.adx14_prev,

    high40: calc.high40,
    high40_prev: calc.high40_prev,

    breakout: calc.breakout,

    break_resline60: calc.break_resline60,
    ma_bull_stack: calc.ma_bull_stack,
    wk_kd_golden: calc.wk_kd_golden,
    neckline60: calc.neckline60,
    break_neckline60: calc.break_neckline60,
    bottom_lead_hits: calc.bottom_lead_hits,
    bottom_lead_count: calc.bottom_lead_count,
    bottom_lead_hits2: calc.bottom_lead_hits2,
    bottom_lead2_count: calc.bottom_lead2_count,
    position_cap: calc.position_cap,
    position_cap2: calc.position_cap2,

    ma10w: calc.ma10w,
    above_ma10w: calc.above_ma10w,
    wk_macd_golden: calc.wk_macd_golden,
    wk_ma10w_2w_up: calc.wk_ma10w_2w_up
  };

  return header.map(h => Object.prototype.hasOwnProperty.call(values, h) ? values[h] : '');
}

function log_(ss, level, ticker, message) {
  const sheet = getOrCreateSheet_(ss, SHEET_LOG);
  sheet.appendRow([new Date(), level, ticker, message]);
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
