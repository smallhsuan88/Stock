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

const SLEEP_MS_EACH_TICKER = 250; // 避免被節流

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
    sig.appendRow([
      'market','ticker','date','close','volume',
      'ma10','ma20','ma60','ma10_slope','ma20_slope','ma60_slope',
      'vol_ma20','vol_ratio',
      'rsi14','rsi14_prev',
      'adx14','adx14_prev',
      'high40','high40_prev',
      'breakout'
    ]);
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
  sig.clearContents();
  sig.appendRow([
    'market','ticker','date','close','volume',
    'ma10','ma20','ma60','ma10_slope','ma20_slope','ma60_slope',
    'vol_ma20','vol_ratio',
    'rsi14','rsi14_prev',
    'adx14','adx14_prev',
    'high40','high40_prev',
    'breakout'
  ]);

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

      sig.appendRow([
        'TW',
        ticker,
        calc.date,
        calc.close,
        calc.volume,

        calc.ma10,
        calc.ma20,
        calc.ma60,
        calc.ma10_slope,
        calc.ma20_slope,
        calc.ma60_slope,

        calc.vol_ma20,
        calc.vol_ratio,

        calc.rsi14,
        calc.rsi14_prev,

        calc.adx14,
        calc.adx14_prev,

        calc.high40,
        calc.high40_prev,

        calc.breakout
      ]);

      log_(ss, 'INFO', ticker,
        `OK ${calc.date} breakout=${calc.breakout} ` +
        `vol_ratio=${num_(calc.vol_ratio,2)} rsi=${num_(calc.rsi14,1)} adx=${num_(calc.adx14,1)}`
      );
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

    breakout
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

function log_(ss, level, ticker, message) {
  const sheet = getOrCreateSheet_(ss, SHEET_LOG);
  sheet.appendRow([new Date(), level, ticker, message]);
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}
