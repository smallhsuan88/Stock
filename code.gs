/** =========================
 * TW MVP - Turbo Version (Incremental Fetch)
 * - Source: TWSE STOCK_DAY (monthly)
 * - Optimization: Fetches only missing months if data exists
 * - Rule #1 (Pressure50): State-based (Close > Pressure)
 * - Rule #4 (Neckline): State-based (Close > Neckline)
 * - Rule #7 (MA Align): Strict MA10 > MA20 > MA60
 * - Rule #8 (MACD): Golden Cross + Histogram Growing
 * - Rule #11 (Month KD): State-based (K > D AND D < 60)
 * ========================= **/

const SHEET_CONFIG  = 'Config_TW';
const SHEET_RAW     = 'Raw_TW';
const SHEET_SIGNALS = 'Signals_TW';
const SHEET_LOG     = 'Log';

const LOOKBACK_MONTHS = 24; // 兩年資料，用於月 KD 收斂

const HIGH_N    = 40;
const PRESSURE_N = 50;
const VOL_MA_N = 20;
const TW_MA10_N   = 10;
const TW_MA20_N   = 20;
const TW_MA60_N   = 60;
const RESLINE_EPS = 0.005;
const NECKLINE_TOLERANCE = 0.005;
const NECKLINE_EPS = 0.005;

const SLEEP_MS_EACH_TICKER = 100; // 有快取後可稍微縮短

const TIME_BUDGET_MS = 330000; 
const CHUNK_MIN_LEFT_MS = 15000; 
const PROP_KEY = 'TW_MVP_PROGRESS';
const PROP_TRIGGER_AT = 'TW_MVP_NEXT_TRIGGER_AT';

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

const SIGNAL_HEADERS_MONTHLY = [
  'm_price_vol_up',
  'm_kd_low_golden',
  'm_2m_no_lower_low',
  'bottom_lead_hits3',
  'bottom_lead3_count',
  'position_cap3',
  'exit_action',
  'exit_note'
];

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

function runTW_MVP() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  const logBuffer = [];
  const signalsRows = [];
  
  // 這裡改用 Map 儲存新抓到的所有資料 (Key: Ticker)，用於最後回寫 Raw
  const processedRawData = new Map(); 
  
  let raw = null;
  let sig = null;
  let logSheet = null;
  let rawHeader = null;
  let sigHeader = null;

  try {
    const cfg = getOrCreateSheet_(ss, SHEET_CONFIG);
    raw = getOrCreateSheet_(ss, SHEET_RAW);
    sig = getOrCreateSheet_(ss, SHEET_SIGNALS);
    logSheet = getOrCreateSheet_(ss, SHEET_LOG);

    const tickers = readTickers_(cfg);
    if (tickers.length === 0) {
      log_(logBuffer, 'WARN', '', 'No active tickers in Config_TW');
      flushLogs_(logSheet, logBuffer);
      return;
    }

    // 1. 預先讀取 Raw_TW 現有資料建立 Cache (加速的核心)
    const rawDataMap = loadRawDataToMap_(raw);

    const tickersKey = tickers.join(',');
    let progress = parseProgress_(props.getProperty(PROP_KEY));
    const nowMs = Date.now();
    const runStartMs = nowMs;
    let runId = progress.runId;
    let next = progress.next || 0;
    let startedAt = progress.startedAt || nowMs;

    if (!runId || progress.tickersKey !== tickersKey || next >= tickers.length) {
      runId = Utilities.getUuid();
      next = 0;
      startedAt = nowMs;
    }

    props.setProperty(PROP_KEY, JSON.stringify({ next, runId, startedAt, tickersKey }));

    const rawValues = raw.getLastRow() > 0 ? raw.getDataRange().getValues() : [];
    rawHeader = rawValues.length ? rawValues[0] : ['ticker','date','open','high','low','close','volume'];

    const existingSigHeader = sig.getLastRow() > 0
      ? sig.getRange(1, 1, 1, sig.getLastColumn()).getValues()[0].filter(v => String(v).trim() !== '')
      : [];
    sigHeader = buildSignalHeaders_(existingSigHeader);

    // 若是新的一輪 (next=0)，清空 Signals 表
    if (next === 0) {
      sig.clearContents();
      sig.getRange(1, 1, 1, sigHeader.length).setValues([sigHeader]);
    } else if (sig.getLastRow() === 0) {
      sig.getRange(1, 1, 1, sigHeader.length).setValues([sigHeader]);
    }

    log_(logBuffer, 'INFO', '', `Run started. RunID=${runId} Tickers=${tickers.length}, StartIdx=${next}`);

    for (let i = next; i < tickers.length; i++) {
      const ticker = tickers[i];
      try {
        // 從 Cache 拿舊資料
        const existingRows = rawDataMap.get(ticker) || [];
        
        // 智慧抓取：只抓缺少的月份
        const { rows: mergedRows, fetchedCount } = fetchSmartTWSEData_(ticker, existingRows, LOOKBACK_MONTHS);
        
        if (!mergedRows.length) {
          log_(logBuffer, 'WARN', ticker, 'No data (ETF/OTC/API Empty)');
        } else {
          // 保存合併後的資料，稍後寫回 Raw
          processedRawData.set(ticker, mergedRows);

          // 計算指標
          const calc = calcSignalsFromRows_(mergedRows);
          signalsRows.push(buildSignalRow_(sigHeader, calc, ticker));

          log_(logBuffer, 'INFO', ticker,
            `OK ${calc.date} (Fetched ${fetchedCount}m) ` +
            `MA60=${num_(calc.ma60,1)} Breakout=${calc.breakout}`
          );
          if (calc.infoLogs && calc.infoLogs.length) {
            calc.infoLogs.forEach(msg => log_(logBuffer, 'INFO', ticker, msg));
          }
        }
      } catch (e) {
        log_(logBuffer, 'ERROR', ticker, String(e && e.stack ? e.stack : e));
      } finally {
        Utilities.sleep(SLEEP_MS_EACH_TICKER);
      }

      // Time Budget Check
      if (Date.now() - runStartMs > TIME_BUDGET_MS - CHUNK_MIN_LEFT_MS) {
        props.setProperty(PROP_KEY, JSON.stringify({
          next: i + 1,
          runId,
          startedAt,
          tickersKey
        }));
        scheduleNextRun_(props);
        log_(logBuffer, 'INFO', '', `Paused & scheduled next run. next=${i + 1}`);
        // 暫停前先存檔
        flushLogs_(logSheet, logBuffer);
        flushSignals_(sig, signalsRows, sigHeader);
        flushSmartRaw_(raw, rawHeader, rawDataMap, processedRawData); // 關鍵：混合寫入
        return;
      }
    }

    // 跑完所有股票
    const secs = Math.round((Date.now() - runStartMs) / 1000);
    log_(logBuffer, 'INFO', '', `Run finished. Total seconds=${secs}`);
    props.deleteProperty(PROP_KEY);
    props.deleteProperty(PROP_TRIGGER_AT);

  } catch (e) {
    log_(logBuffer, 'ERROR', '', String(e && e.stack ? e.stack : e));
  } finally {
    if (sig && sigHeader) flushSignals_(sig, signalsRows, sigHeader);
    if (logSheet) flushLogs_(logSheet, logBuffer);
    if (raw && rawHeader) {
      // 這裡需要用特殊的 flush 邏輯：保留沒處理到的 ticker，更新已處理的 ticker
      flushSmartRaw_(raw, rawHeader, rawDataMap, processedRawData);
    }
    lock.releaseLock();
  }
}

/** 讀取 Raw Sheet 全部資料到 Map<Ticker, Rows[]> */
function loadRawDataToMap_(sheet) {
  const map = new Map();
  if (sheet.getLastRow() <= 1) return map;
  
  const values = sheet.getDataRange().getValues();
  // row[0]=ticker, row[1]=date (yyyy-MM-dd), ...
  // headers: ticker,date,open,high,low,close,volume
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const t = String(row[0]).trim();
    if (!t) continue;
    
    if (!map.has(t)) map.set(t, []);
    // 轉為物件格式方便操作，保留原始 row 結構也行，這裡轉物件比較安全
    map.get(t).push({
      date: formatDateISO_(row[1]),
      open: Number(row[2]),
      high: Number(row[3]),
      low: Number(row[4]),
      close: Number(row[5]),
      volume: Number(row[6])
    });
  }
  
  // 確保每個 Ticker 內的日期是排序的
  for (const [k, rows] of map) {
    rows.sort((a,b) => a.date.localeCompare(b.date));
  }
  return map;
}

/** 智慧抓取：比較舊資料，只抓缺少的月份 */
function fetchSmartTWSEData_(ticker, existingRows, lookbackMonths) {
  const now = new Date();
  const neededStart = new Date(now.getFullYear(), now.getMonth() - lookbackMonths, 1);
  const neededStartStr = Utilities.formatDate(neededStart, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  let fetchStartDate = neededStart;
  
  // 檢查現有資料最後一筆
  if (existingRows.length > 0) {
    const lastRow = existingRows[existingRows.length - 1];
    const lastDate = new Date(lastRow.date);
    
    // 如果舊資料最後一天已經是這個月或之後，可能只需要更新這個月
    // 為了保險起見，我們從「最後一筆資料的月份」開始抓，確保該月資料完整
    // 例如最後一筆是 2024-05-15，我們重抓 2024-05 整月，補齊後面日期
    fetchStartDate = new Date(lastDate.getFullYear(), lastDate.getMonth(), 1);
    
    // 如果計算出的 fetchStartDate 比 neededStart 還早，就用 neededStart (不需要那麼早的)
    if (fetchStartDate < neededStart) fetchStartDate = neededStart;
  }

  // 計算需要抓取的月份列表
  const monthsToFetch = [];
  let curr = new Date(fetchStartDate);
  const end = new Date(now.getFullYear(), now.getMonth(), 1); // 到本月1號
  
  while (curr <= end) {
    monthsToFetch.push({
      y: curr.getFullYear(),
      m: String(curr.getMonth() + 1).padStart(2, '0')
    });
    // 下個月
    curr = new Date(curr.getFullYear(), curr.getMonth() + 1, 1);
  }
  
  // 實際抓取
  let newFetchedRows = [];
  for (const ym of monthsToFetch) {
    const data = fetchTWSEMonthData_(ticker, ym.y, ym.m);
    if (data) {
       data.forEach(arr => {
        const dateISO = rocToISO_(arr[0]);
        if (!dateISO) return;
        const volume = safeVolume_(arr);
        const close = parseNum_(arr[6]);
        if (isFinite(close) && close > 0) {
          newFetchedRows.push({
            date: dateISO,
            open: parseNum_(arr[3]),
            high: parseNum_(arr[4]),
            low: parseNum_(arr[5]),
            close: close,
            volume: volume
          });
        }
      });
    }
    // 避免太快
    Utilities.sleep(150); 
  }

  // 合併邏輯：
  // 1. 過濾掉 existingRows 中太舊的 (比 lookback 舊的)
  // 2. 過濾掉 existingRows 中日期 >= newFetchedRows 最小日期的 (因為新抓的比較準)
  // 3. 組合
  
  const minNewDate = newFetchedRows.length > 0 ? newFetchedRows[0].date : '9999-99-99';
  
  const keepOld = existingRows.filter(r => r.date >= neededStartStr && r.date < minNewDate);
  const combined = keepOld.concat(newFetchedRows);
  
  // 再次排序 + 去重 (以防萬一)
  combined.sort((a,b) => a.date.localeCompare(b.date));
  
  const dedup = [];
  const seen = new Set();
  for (const r of combined) {
    if (seen.has(r.date)) continue;
    seen.add(r.date);
    dedup.push(r);
  }

  return { rows: dedup, fetchedCount: monthsToFetch.length };
}


/** 聰明回寫：保留未處理的 Ticker，更新已處理的 Ticker */
function flushSmartRaw_(sheet, header, originalMap, processedMap) {
  // processedMap 包含這次執行有更新到的 ticker
  // originalMap 包含所有原始資料
  
  // 我們需要將 processedMap 的內容更新回 originalMap
  for (const [ticker, rows] of processedMap) {
    originalMap.set(ticker, rows);
  }
  
  // 轉成 Array 準備寫入
  const output = [];
  // 根據 Config 的順序或 Map 的順序都行，這裡簡單遍歷 Map
  // 為了美觀，可以稍微排個序
  const sortedTickers = Array.from(originalMap.keys()).sort();
  
  for (const t of sortedTickers) {
    const rows = originalMap.get(t);
    rows.forEach(r => {
      output.push([t, r.date, r.open, r.high, r.low, r.close, r.volume]);
    });
  }
  
  if (output.length === 0) return;

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  
  // 批次寫入
  const batchSize = 5000;
  for (let i = 0; i < output.length; i += batchSize) {
    const batch = output.slice(i, i + batchSize);
    sheet.getRange(2 + i, 1, batch.length, header.length).setValues(batch);
  }
}

function formatDateISO_(dateObj) {
  if (dateObj instanceof Date) {
    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(dateObj);
}

// --- 以下維持原本的輔助函數與指標計算邏輯 (Rules) ---

function readTickers_(cfgSheet) {
  const values = cfgSheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const t = String(values[i][0] || '').trim();
    if (!t) continue;
    const active = values[i][1];
    if (active === false) continue;
    out.push(t);
  }
  return out;
}

function fetchTWSEMonthData_(ticker, year, mm) {
  const url = `https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=${year}${mm}01&stockNo=${ticker}`;
  const backoffMs = [300, 800];

  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      const code = resp.getResponseCode();
      if (code !== 200) throw new Error(`HTTP ${code}`);

      const text = resp.getContentText();
      const json = JSON.parse(text);
      if (!json || !json.data || !Array.isArray(json.data)) {
        // 有些月份可能真的沒資料或休市
        return []; 
      }
      if (json.stat && json.stat !== 'OK') {
        // stat 不 OK 可能是查詢過於頻繁或無資料
        return [];
      }
      return json.data;
    } catch (e) {
      if (attempt < backoffMs.length) {
        Utilities.sleep(backoffMs[attempt]);
      }
    }
  }
  return null;
}

function safeVolume_(arr) {
  const v1 = parseNum_(arr[1]);
  if (v1 > 0) return v1;
  for (let i = 0; i < arr.length; i++) {
    const v = parseNum_(arr[i]);
    if (v > 1000) return v;
  }
  return 0;
}

/** 計算 signals */
function calcSignalsFromRows_(rows) {
  const n = rows.length;
  const last = rows[n - 1];
  const infoLogs = [];

  const closes = rows.map(r => r.close);
  const vols   = rows.map(r => r.volume);
  const highs  = rows.map(r => r.high);
  const lows   = rows.map(r => r.low);

  // 今日均線
  const ma10 = sma_(closes, TW_MA10_N);
  const ma20 = sma_(closes, TW_MA20_N);
  const ma60 = sma_(closes, TW_MA60_N);
  const pressure50Series = rollingAvg_(highs, PRESSURE_N);
  const pressure50 = pressure50Series.length ? pressure50Series[n - 1] : '';

  // 昨日均線
  const closesPrev = closes.slice(0, -1);
  const highsPrev  = highs.slice(0, -1);
  const lowsPrev   = lows.slice(0, -1);
  const volsPrev   = vols.slice(0, -1);

  const ma10_prev = sma_(closesPrev, TW_MA10_N);
  const ma20_prev = sma_(closesPrev, TW_MA20_N);
  const ma60_prev = sma_(closesPrev, TW_MA60_N);
  const ma60_prev2 = sma_(closes.slice(0, -2), TW_MA60_N);
  const pressure50_prev = pressure50Series.length > 1 ? pressure50Series[n - 2] : '';

  const vol_ma20 = sma_(volsPrev, VOL_MA_N);
  const vol_ratio = (isFiniteNum_(vol_ma20) && vol_ma20 > 0) ? (last.volume / vol_ma20) : '';

  const high40 = Math.max(...highs.slice(Math.max(0, n - HIGH_N)));

  const nPrev = highsPrev.length;
  const high40_prev = (nPrev >= 1)
    ? Math.max(...highsPrev.slice(Math.max(0, nPrev - HIGH_N)))
    : '';

  const rsi14 = rsiWilder_(closes, 14);
  const adx14 = adxWilder_(highs, lows, closes, 14);

  const rsi14_prev = rsiWilder_(closesPrev, 14);
  const adx14_prev = adxWilder_(highsPrev, lowsPrev, closesPrev, 14);

  const breakout =
    (isFiniteNum_(high40_prev) ? (last.close > high40_prev) : false) &&
    (isFiniteNum_(vol_ma20) ? (last.volume > vol_ma20) : false);

  let ma_bull_stack = '';
  if (isFiniteNum_(ma10) && isFiniteNum_(ma20)) {
    // 規則 #2: 僅檢查 MA10 > MA20
    ma_bull_stack = (last.close > ma10 && last.close > ma20 && ma10 > ma20) ? 'TRUE' : 'FALSE';
  }

  const resline = calcBreakResline60_(highs, last.close, n, infoLogs);
  const break_resline60 = resline.status;

  const neckline = calcNeckline60_(closes, last.close, n, infoLogs);
  const neckline60 = neckline.neckline;
  const break_neckline60 = neckline.breakStatus;

  const weeks = toWeeklyBars_(rows);
  const wCloses = weeks.map(w => w.close);

  const kd = calcWeeklyKDG_(weeks, infoLogs);
  const wk_kd_golden = kd.status;

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
    // 規則 #9: 站上兩週 且 10週均線翻揚
    wk_ma10w_2w_up = (lastWeekClose > ma10w && prevWeekClose > ma10w_prev && ma10w > ma10w_prev) ? 'TRUE' : 'FALSE';
  }

  // 規則 #8: 週 MACD 金叉 且 柱狀體擴大
  const macd = calcWeeklyMacd_(wCloses, infoLogs);
  const wk_macd_golden = macd.status;

  const months = toMonthlyBars_(rows);

  let m_price_vol_up = '';
  if (months.length >= 2) {
    const lastMonth = months[months.length - 1];
    const prevMonth = months[months.length - 2];
    const priceUp = lastMonth.close > prevMonth.close;
    const volUp = lastMonth.volume > prevMonth.volume;
    m_price_vol_up = (priceUp && volUp) ? 'TRUE' : 'FALSE';
  }

  // 規則 #11: 月 KD (State-based)
  const mkd = calcMonthlyKDLowGolden_(months, infoLogs);
  const m_kd_low_golden = mkd.status;

  let m_2m_no_lower_low = '';
  if (months.length >= 3) {
    const M = months[months.length - 1];
    const M1 = months[months.length - 2];
    const M2 = months[months.length - 3];
    const condA = M.low >= M1.low;
    const condB = M1.low >= M2.low;
    m_2m_no_lower_low = (condA && condB) ? 'TRUE' : 'FALSE';
  }

  const m_2m_no_break_low = m_2m_no_lower_low;

  const midHits = [];
  const monthlyStatuses = [m_price_vol_up, m_kd_low_golden, m_2m_no_lower_low];
  const monthlyAvailable = monthlyStatuses.some(s => s !== '');
  if (m_price_vol_up === 'TRUE') midHits.push('1');
  if (m_kd_low_golden === 'TRUE') midHits.push('2');
  if (m_2m_no_lower_low === 'TRUE') midHits.push('3');

  const mid_lead_hits_m = monthlyAvailable ? (midHits.length ? midHits.join(',') : '') : '';
  const mid_lead_m_count = monthlyAvailable ? String(midHits.length) : '';
  const mid_position_cap = (monthlyAvailable && midHits.length >= 3) ? '20%' : '';

  const monthlyHits = [];
  if (m_price_vol_up === 'TRUE') monthlyHits.push('10');
  if (m_kd_low_golden === 'TRUE') monthlyHits.push('11');
  if (m_2m_no_lower_low === 'TRUE') monthlyHits.push('12');

  const bottom_lead_hits3 = monthlyAvailable ? (monthlyHits.length ? monthlyHits.join(',') : '') : '';
  const bottom_lead3_count = monthlyAvailable ? String(monthlyHits.length) : '';
  const position_cap3 = (monthlyAvailable && monthlyHits.length === 3) ? '20%' : '';

  let exit_action = '';
  let exit_note = '';
  if (isFiniteNum_(ma60) && last.close < ma60) {
    exit_action = 'SELL_ALL';
    exit_note = '跌破60日均，全賣出';
  } else if (isFiniteNum_(ma20) && last.close < ma20) {
    exit_action = 'SELL_HALF';
    exit_note = '跌破20日均，賣一半';
  }

  // 底部領先訊號彙總
  const conds = [];
  const hits1to5 = [];

  const prevClose = n >= 2 ? closes[n - 2] : '';
  const cond1Available = isFiniteNum_(pressure50);
  if (cond1Available) {
    // 規則 #1: State (Close > Pressure)
    const cond1 = last.close > pressure50;
    conds.push(cond1);
    if (cond1) hits1to5.push('1');
  } else {
    infoLogs.push('pressure50 unavailable for cond1 (need >= 50 highs)');
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
  const cond5Status = evalCond5Status_(last.close, ma60, ma60_prev, ma60_prev2);
  if (cond5Status !== '') {
    const cond5 = cond5Status === 'TRUE';
    conds.push(cond5);
    if (cond5) hits1to5.push('5');
  } else {
    infoLogs.push('ma60 unavailable for cond5 (need >= 62 closes)');
  }

  const weeklyHits = [];
  const weeklyAvailability = [];
  const pushWeekly = (flag, label) => {
    if (flag === 'TRUE') weeklyHits.push(label);
    if (flag === 'TRUE' || flag === 'FALSE') weeklyAvailability.push(true);
  };
  pushWeekly(above_ma10w, '6');
  
  // 規則 #7 Check
  const cond7Status = evalCond7Status_(last.close, ma10, ma20, ma60);
  pushWeekly(cond7Status, '7');
  
  pushWeekly(wk_macd_golden, '8');
  pushWeekly(wk_ma10w_2w_up, '9');

  const hasWeeklyAvailability = weeklyAvailability.length > 0;
  const bottom_lead_hits2 = hasWeeklyAvailability ? (weeklyHits.length ? weeklyHits.join(',') : '') : '';
  const bottom_lead2_count = hasWeeklyAvailability ? String(weeklyHits.length) : '';

  const allHits = hits1to5.concat(weeklyHits);
  const anyCondAvailable = conds.length > 0 || hasWeeklyAvailability;
  const bottom_lead_hits = allHits.length > 0 ? allHits.join(',') : '';
  const bottom_lead_count = anyCondAvailable ? String(allHits.length) : '';

  const mergedSet = new Set();
  const mergeHits = (src) => {
    if (!src) return;
    String(src).split(',').map(s => s.trim()).filter(Boolean).forEach(v => mergedSet.add(v));
  };
  mergeHits(bottom_lead_hits);
  mergeHits(bottom_lead_hits2);
  mergeHits(bottom_lead_hits3);

  const mergedHits = Array.from(mergedSet).filter(Boolean).sort((a, b) => Number(a) - Number(b));
  const mergedBottomLeadHits = mergedHits.length ? mergedHits.join(',') : '';

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
    bottom_lead_hits: mergedBottomLeadHits, 
    bottom_lead_count,
    bottom_lead_hits2,
    bottom_lead2_count,
    position_cap,
    position_cap2,

    ma10w,
    above_ma10w,
    wk_macd_golden,
    wk_ma10w_2w_up,

    m_price_vol_up,
    m_kd_low_golden,
    m_2m_no_break_low,
    m_2m_no_lower_low,
    mid_lead_hits_m,
    mid_lead_m_count,
    mid_position_cap,
    bottom_lead_hits3,
    bottom_lead3_count,
    position_cap3,
    exit_action,
    exit_note,

    infoLogs
  };
}

function rollingAvg_(arr, win) {
  if (!arr || arr.length < win) return new Array(arr ? arr.length : 0).fill('');
  const out = new Array(arr.length).fill('');
  let sum = 0;
  for (let i = 0; i < arr.length; i++) {
    sum += Number(arr[i]) || 0;
    if (i >= win) sum -= Number(arr[i - win]) || 0;
    if (i >= win - 1) out[i] = sum / win;
  }
  return out;
}

function sma_(arr, win) {
  if (!arr || arr.length < win) return '';
  let sum = 0;
  for (let i = arr.length - win; i < arr.length; i++) sum += Number(arr[i]) || 0;
  return sum / win;
}

function evalCond5Status_(close, ma60, ma60_prev1, ma60_prev2) {
  if (!isFiniteNum_(close) || !isFiniteNum_(ma60) || !isFiniteNum_(ma60_prev1) || !isFiniteNum_(ma60_prev2)) return '';
  return (close > ma60 && ma60 > ma60_prev1 && ma60_prev1 > ma60_prev2) ? 'TRUE' : 'FALSE';
}

function evalCond7Status_(close, ma10, ma20, ma60) {
  if (!isFiniteNum_(close) || !isFiniteNum_(ma10) || !isFiniteNum_(ma20) || !isFiniteNum_(ma60)) return '';
  
  const c = Number(close);
  const m10 = Number(ma10);
  const m20 = Number(ma20);
  const m60 = Number(ma60);

  const priceCond = c >= (m60 * 1.03);
  const alignCond = (m10 > m20) && (m20 > m60);

  return (priceCond && alignCond) ? 'TRUE' : 'FALSE';
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
  const end = n - 1; 
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
  const prices = closes.slice(start, n - 1); 

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

function toMonthlyBars_(rows) {
  if (!rows || !rows.length) return [];
  const months = [];
  rows.forEach(r => {
    const key = r.date.slice(0, 7);
    if (!months.length || months[months.length - 1].key !== key) {
      months.push({
        key,
        open: r.open,
        high: r.high,
        low: r.low,
        close: r.close,
        volume: r.volume
      });
    } else {
      const m = months[months.length - 1];
      m.high = Math.max(m.high, r.high);
      m.low = Math.min(m.low, r.low);
      m.close = r.close;
      m.volume += r.volume;
    }
  });
  return months;
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
    if (i < 8) continue; 
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

  const isGoldenCross = macdLine[lastIdx] > signalLine[lastIdx] && macdLine[prevIdx] <= signalLine[prevIdx];
  
  const currHist = macdLine[lastIdx] - signalLine[lastIdx];
  const prevHist = macdLine[prevIdx] - signalLine[prevIdx];
  const isPositive = currHist > 0;
  const isGrowing = currHist > prevHist;

  const hit = isGoldenCross && isPositive && isGrowing;
  return { status: hit ? 'TRUE' : 'FALSE', available: true, hit };
}

// 規則 #11: State (K > D AND D < 60)
function calcMonthlyKDLowGolden_(months, infoLogs) {
  if (!months || months.length === 0) return { status: '', available: false, hit: false };

  if (months.length < 15) {
    infoLogs.push(`m_kd_low_golden skipped: insufficient history (have ${months.length})`);
    return { status: '', available: false, hit: false };
  }

  let K = 50;
  let D = 50;

  for (let i = 0; i < months.length; i++) {
    if (i < 8) continue; 
    const window = months.slice(Math.max(0, i - 8), i + 1);
    const highs = window.map(m => m.high);
    const lows = window.map(m => m.low);
    const close = months[i].close;

    const highestHigh = Math.max(...highs);
    const lowestLow = Math.min(...lows);
    
    let RSV = 50;
    if (highestHigh !== lowestLow) {
      RSV = (close - lowestLow) / (highestHigh - lowestLow) * 100;
    }

    K = (2 / 3) * K + (1 / 3) * RSV;
    D = (2 / 3) * D + (1 / 3) * K;
  }

  const goldenState = (K > D);
  const lowBand = (D < 60); 
  
  const hit = goldenState && lowBand;
  return { status: hit ? 'TRUE' : 'FALSE', available: true, hit };
}

function weekKey_(dateStr) {
  const dt = new Date(`${dateStr}T00:00:00+08:00`);
  const day = dt.getDay(); 
  const diff = (day + 6) % 7; 
  dt.setDate(dt.getDate() - diff);
  return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

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
  SIGNAL_HEADERS_MONTHLY.forEach(addField);

  const placeAfter = (field, afterField) => {
    if (!headerSet.has(afterField)) return;
    if (!headerSet.has(field)) {
      headerSet.add(field);
      finalHeaders.push(field);
    }
    const currentIdx = finalHeaders.indexOf(field);
    if (currentIdx === -1) return;
    const afterIdx = finalHeaders.indexOf(afterField);
    if (afterIdx === -1) return;
    if (currentIdx === afterIdx + 1) return;
    finalHeaders.splice(currentIdx, 1);
    finalHeaders.splice(afterIdx + 1, 0, field);
  };

  placeAfter('position_cap3', 'position_cap2');
  placeAfter('exit_action', 'position_cap3');
  placeAfter('exit_note', 'exit_action');

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
    wk_ma10w_2w_up: calc.wk_ma10w_2w_up,

    m_price_vol_up: calc.m_price_vol_up,
    m_kd_low_golden: calc.m_kd_low_golden,
    m_2m_no_break_low: calc.m_2m_no_break_low,
    m_2m_no_lower_low: calc.m_2m_no_lower_low,
    mid_lead_hits_m: calc.mid_lead_hits_m,
    mid_lead_m_count: calc.mid_lead_m_count,
    mid_position_cap: calc.mid_position_cap,
    bottom_lead_hits3: calc.bottom_lead_hits3,
    bottom_lead3_count: calc.bottom_lead3_count,
    position_cap3: calc.position_cap3,
    exit_action: calc.exit_action,
    exit_note: calc.exit_note
  };

  return header.map(h => Object.prototype.hasOwnProperty.call(values, h) ? values[h] : '');
}

function log_(buffer, level, ticker, message) {
  buffer.push([new Date(), level, ticker, message]);
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function flushSignals_(sigSheet, signalsRows, sigHeader) {
  if (!signalsRows.length) return;
  const rowStart = sigSheet.getLastRow() + 1;
  sigSheet.getRange(rowStart, 1, signalsRows.length, sigHeader.length).setValues(signalsRows);
  signalsRows.length = 0;
}

function flushLogs_(logSheet, logBuffer) {
  if (!logBuffer.length) return;
  const rowStart = logSheet.getLastRow() + 1;
  logSheet.getRange(rowStart, 1, logBuffer.length, 4).setValues(logBuffer);
  logBuffer.length = 0;
}

function scheduleNextRun_(props) {
  const now = Date.now();
  const nextAtRaw = props.getProperty(PROP_TRIGGER_AT);
  const nextAt = nextAtRaw ? Number(nextAtRaw) : 0;
  if (nextAt && nextAt > now) return;

  ScriptApp.newTrigger('runTW_MVP')
    .timeBased()
    .after(60 * 1000)
    .create();
  props.setProperty(PROP_TRIGGER_AT, String(now + 60 * 1000));
}

function parseProgress_(raw) {
  if (!raw) return { next: 0, runId: '', startedAt: 0, tickersKey: '' };
  try {
    const parsed = JSON.parse(raw);
    return {
      next: Number(parsed.next || 0),
      runId: parsed.runId || '',
      startedAt: Number(parsed.startedAt || 0),
      tickersKey: parsed.tickersKey || ''
    };
  } catch (e) {
    return { next: 0, runId: '', startedAt: 0, tickersKey: '' };
  }
}
