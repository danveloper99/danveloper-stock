/**
 * DANVELOPER 選股模擬系統 - GAS 後端 v2.0
 * ─────────────────────────────────────────
 * 功能：
 *   1. Web App API（資料儲存 / 讀取 / Email 通知）
 *   2. 全市場自動掃描（接力 Trigger，每天盤後自動執行）
 *   3. 掃描完成後 Email 通知
 *
 * 部署步驟：
 *   1. Extensions > Apps Script > 貼上此程式碼
 *   2. 執行 setupAll()（只需一次，建立工作表 + 設定每日觸發器）
 *   3. Deploy > New Deployment > Web App
 *      Execute as: Me | Who has access: Anyone
 *   4. 複製 Web App URL 填入網頁設定
 *
 * Sheets 工作表：
 *   State          - 網頁 state JSON
 *   Trades         - 交易紀錄
 *   Snapshots      - 每日資產快照
 *   Recommendations- 選股推薦（網頁同步用）
 *   ScanResults    - GAS 掃描結果（主要儲存）
 *   ScanProgress   - 掃描進度（接力用）
 *   Notifications  - Email 寄送紀錄
 */

// ══════════════════════════════════════════════
// 設定區（修改這裡）
// ══════════════════════════════════════════════
const CONFIG = {
  FINMIND_TOKEN: '',        // ← 填入你的 FinMind Token
  GEMINI_KEY: '',           // ← 填入你的 Gemini API Key（啟用 AI 分析）
  NOTIFY_EMAIL: '',         // ← 填入通知 Email（留空則寄給 GAS 帳號）
  BATCH_SIZE: 50,           // 每批處理幾檔（不要超過 80，避免逾時）
  SCAN_INTERVAL_MIN: 6,     // 每批間隔分鐘（6分鐘 = GAS 上限前安全觸發）
  SCAN_TRIGGER_HOUR: 14,    // 每天幾點開始掃描（14 = 下午2點，盤後）
  SCAN_TRIGGER_MIN: 30,     // 幾分開始
};

const SHEET_NAMES = {
  STATE: 'State',
  TRADES: 'Trades',
  SNAPSHOTS: 'Snapshots',
  SCAN: 'ScanResults',
  PROGRESS: 'ScanProgress',
  NOTIF: 'Notifications'
};

// ══════════════════════════════════════════════
// HTTP HANDLERS
// ══════════════════════════════════════════════

function doGet(e) {
  const action = e.parameter.action || '';

  // 無 action → 回傳掃描結果（供網頁讀取今日推薦）
  if (!action) {
    const results = loadScanResults();
    return jsonResponse({ status: 'ok', data: results });
  }

  try {
    if (action === 'getState') return jsonResponse({ status: 'ok', data: loadStateFromSheet() });
    if (action === 'getTrades') return jsonResponse({ status: 'ok', data: loadTrades() });
    if (action === 'getSnapshots') return jsonResponse({ status: 'ok', data: loadSnapshots() });
    if (action === 'getScanResults') return jsonResponse({ status: 'ok', data: loadScanResults() });
    if (action === 'getScanProgress') return jsonResponse({ status: 'ok', data: loadScanProgress() });
    if (action === 'getOHLC') {
      const { market, symbols, date } = body;
      return jsonResponse({ status: 'ok', data: fetchOHLCBatch(market, symbols, date) });
    }
    return jsonResponse({ status: 'error', msg: 'Unknown action' });
  } catch(err) {
    return jsonResponse({ status: 'error', msg: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'saveState') {
      saveStateToSheet(body.data);
      saveTradesFromState(body.data.trades || []);
      saveSnapshotsFromState(body.data.dailySnapshots || []);
      return jsonResponse({ status: 'ok' });
    }

    if (action === 'sendNotification') {
      const sent = sendScanEmail(body.results || [], body.summary || {});
      return jsonResponse({ status: 'ok', sent });
    }

    if (action === 'saveConfig') {
      saveConfigToSheet(body.config || {});
      if (body.config && (body.config.SCAN_TRIGGER_HOUR !== undefined || body.config.SCAN_TRIGGER_MIN !== undefined)) {
        resetDailyTrigger();
      }
      return jsonResponse({ status: 'ok' });
    }

    if (action === 'startScan') {
      loadConfigFromSheet();
      initScanProgress();
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'runScanBatch') ScriptApp.deleteTrigger(t);
      });
      ScriptApp.newTrigger('runScanBatch').timeBased().after(60 * 1000).create();
      Logger.log('startScan：進度已初始化，1 分鐘後開始第一批');
      return jsonResponse({ status: 'ok', msg: '掃描已排程，1 分鐘後開始執行，完成後寄 Email 通知' });
    }

    if (action === 'finalize') {
      // 手動觸發收尾（掃描卡住時使用）
      loadConfigFromSheet();
      finalizeScan();
      return jsonResponse({ status: 'ok', msg: 'finalizeScan 已執行' });
    }

    return jsonResponse({ status: 'error', msg: 'Unknown action' });
  } catch(err) {
    return jsonResponse({ status: 'error', msg: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 處理 CORS preflight（OPTIONS 請求）
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ══════════════════════════════════════════════
// 全市場掃描主流程（接力執行）
// ══════════════════════════════════════════════

/**
 * 每天定時觸發的入口（由 setupAll 設定的 Trigger 呼叫）
 * 重置進度，從第 0 批開始
 */
function dailyScanStart() {
  Logger.log('=== 每日掃描開始 ===');
  loadConfigFromSheet(); // 載入最新設定
  initScanProgress();
  runScanBatch(); // 執行第一批
}

/**
 * 接力觸發的入口（由上一批的 Trigger 呼叫）
 */
function runScanBatch() {
  loadConfigFromSheet(); // 每批都重新讀取最新設定
  const progress = loadScanProgress();
  if (!progress || progress.done) {
    Logger.log('掃描已完成或無進度，跳過');
    return;
  }

  const batchIndex = progress.batchIndex || 0;
  const allCandidates = progress.candidates || [];
  const start = batchIndex * CONFIG.BATCH_SIZE;
  const end = Math.min(start + CONFIG.BATCH_SIZE, allCandidates.length);
  const currentBatch = allCandidates.slice(start, end);

  Logger.log(`執行第 ${batchIndex + 1} 批，處理 ${start}~${end} / ${allCandidates.length} 檔`);

  // 第一批：先做第一階段初篩
  if (batchIndex === 0) {
    const phase1Results = runPhase1();
    if (!phase1Results || phase1Results.length === 0) {
      Logger.log('第一階段初篩無結果，掃描結束');
      markScanDone([], true);
      return;
    }
    // 更新候選清單並執行第一批第二階段
    progress.candidates = phase1Results.map(s => s.id);
    progress.nameMap = {};
    phase1Results.forEach(s => { if (s.name) progress.nameMap[s.id] = s.name; });
    progress.totalCandidates = phase1Results.length;

    // 真正的大盤判斷：加權指數 vs 20日均線
    const market = checkMarketTrend();
    progress.marketPass = market.pass;
    progress.marketIndex = market.index;
    progress.marketMA20 = market.ma20;
    progress.advCount = _phase1AdvDec.adv;
    progress.decCount = _phase1AdvDec.dec;

    // 止穩訊號 Checklist（5項自動判定，只在批次0跑一次）
    try {
      progress.stabilization = checkStabilizationSignals();
      Logger.log(`止穩訊號：${progress.stabilization.verdict}`);
    } catch(e) {
      Logger.log('止穩訊號檢查失敗：' + e.message);
      progress.stabilization = null;
    }

    // 空頭警示：大盤空頭確認 + 庫存清單逐檔檢查（只在批次0跑一次）
    try {
      progress.bearishWarning = checkPortfolioBearishWarning();
      Logger.log(`空頭警示：大盤${progress.bearishWarning.market.verdict}，持股警示 ${progress.bearishWarning.warnedCount}/${progress.bearishWarning.totalHoldings} 檔`);
    } catch(e) {
      Logger.log('空頭警示檢查失敗：' + e.message);
      progress.bearishWarning = null;
    }

    saveScanProgress(progress);
    Logger.log(`第一階段完成：${phase1Results.length} 檔通過初篩，大盤${market.pass ? '多頭' : '空頭(謹慎)'}`);
  }

  // 重新讀取（因為第一批可能更新了 candidates）
  const freshProgress = loadScanProgress();
  const freshCandidates = freshProgress.candidates || [];
  const freshStart = batchIndex === 0 ? 0 : batchIndex * CONFIG.BATCH_SIZE;
  const freshEnd = Math.min(freshStart + CONFIG.BATCH_SIZE, freshCandidates.length);
  const batchIds = freshCandidates.slice(freshStart, freshEnd);

  // 執行第二階段（完整技術分析）
  const batchResults = runPhase2Batch(batchIds, freshProgress.nameMap || {});

  // 合併已有結果
  const existing = loadTempResults(); // 從 ScanTemp 讀取已累積的結果
  const merged = [...existing, ...batchResults];

  const nextBatch = batchIndex + 1;
  const hasMore = freshEnd < freshCandidates.length;

  if (hasMore) {
    freshProgress.batchIndex = nextBatch;
    saveScanProgress(freshProgress); // 只存進度，不存結果
    // 把這批新結果追加到 ScanTemp
    const newResults = batchResults.map(r => ({
      id: r.id, name: r.name || '', tier: r.tier, combo: r.combo, combos: r.combos,
      priority: r.priority, hitCount: r.hitCount,
      entryPrice: r.entryPrice, stopLoss: r.stopLoss,
      target1: r.target1, target2: r.target2, atr: r.atr,
      itrustBuy: r.itrustBuy, foreignBuy: r.foreignBuy, rs: r.rs,
      marginUsage: r.marginUsage, marginWeeklyChange: r.marginWeeklyChange, marginBonus: r.marginBonus
    }));
    saveTempResults(newResults);
    scheduleNextBatch(nextBatch);
    Logger.log(`第 ${batchIndex + 1} 批完成，已找到 ${merged.length} 檔，排程第 ${nextBatch + 1} 批`);
  } else {
    // 全部掃描完成：把結果存入 progress，排程獨立的收尾 Trigger
    Logger.log(`所有批次掃描完成，共 ${merged.length} 檔，排程收尾作業...`);
    freshProgress.batchIndex = nextBatch;
    freshProgress.scanDone = true;
    saveScanProgress(freshProgress);
    // 把最後一批新結果追加到 ScanTemp
    const lastResults = batchResults.map(r => ({
      id: r.id, name: r.name || '', tier: r.tier, combo: r.combo, combos: r.combos,
      priority: r.priority, hitCount: r.hitCount,
      entryPrice: r.entryPrice, stopLoss: r.stopLoss,
      target1: r.target1, target2: r.target2, atr: r.atr,
      itrustBuy: r.itrustBuy, foreignBuy: r.foreignBuy, rs: r.rs,
      marginUsage: r.marginUsage, marginWeeklyChange: r.marginWeeklyChange, marginBonus: r.marginBonus
    }));
    saveTempResults(lastResults);

    // 1分鐘後執行收尾（儲存結果 + Gemini 分析 + Email）
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'finalizeScan') ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('finalizeScan').timeBased().after(60 * 1000).create();
    Logger.log('已排程 finalizeScan，1 分鐘後執行');
  }
}

/**
 * 收尾函式：掃描全部完成後由 Trigger 獨立執行
 * 若 Gemini 分析檔數過多，自動分批接力
 */
function finalizeScan() {
  Logger.log('=== finalizeScan 開始 ===');
  loadConfigFromSheet();

  const progress = loadScanProgress();
  if (!progress || !progress.scanDone) {
    Logger.log('無待收尾的掃描，跳過');
    return;
  }

  let merged = loadTempResults(); // 從 ScanTemp 讀取所有結果
  const totalScanned = progress.totalCandidates || 0;
  Logger.log(`finalizeScan：從 ScanTemp 讀到 ${merged.length} 筆結果`);

  // 排序
  merged.sort((a, b) => (b.priority - a.priority) || ((b.rs || 0) - (a.rs || 0)));

  // Gemini 分析：動態分批，每批最多 20 檔（每檔1次，約20秒，安全範圍內）
  const GEMINI_BATCH = 20;
  if (CONFIG.GEMINI_KEY && merged.length > 0 && progress.geminiQuotaExceeded) {
    Logger.log('Gemini 額度本輪已用完，跳過分析');
  } else if (CONFIG.GEMINI_KEY && merged.length > 0) {
    const geminiOffset = progress.geminiOffset || 0; // 目前跑到第幾檔

    if (geminiOffset < merged.length) {
      const batchEnd = Math.min(geminiOffset + GEMINI_BATCH, merged.length);
      const batchSlice = merged.slice(geminiOffset, batchEnd);

      Logger.log(`Gemini 分析 ${geminiOffset + 1}~${batchEnd} / ${merged.length} 檔`);
      let quotaExceeded = false;
      try {
        quotaExceeded = runGeminiAnalysis(batchSlice);
        // 把分析結果寫回 merged
        batchSlice.forEach((r, i) => { merged[geminiOffset + i] = r; });
      } catch(e) {
        Logger.log('Gemini 分析失敗：' + e.message);
      }

      const nextOffset = batchEnd;
      // 把 Gemini 分析結果存回 ScanTemp（覆蓋）
      const resultSheet = getOrCreateSheet('ScanTemp');
      resultSheet.clearContents();
      saveTempResults(merged);
      progress.geminiOffset = nextOffset;

      if (quotaExceeded) {
        // 額度用完：記錄旗標、通知、本次掃描不再呼叫 Gemini，直接進入收尾
        progress.geminiQuotaExceeded = true;
        saveScanProgress(progress);
        notifyGeminiQuotaExceeded(merged.length - nextOffset);
        Logger.log(`Gemini 額度已用完，剩餘 ${merged.length - nextOffset} 檔不再分析，已寄信通知`);
      } else {
        saveScanProgress(progress);

        if (nextOffset < merged.length) {
          // 還有剩，繼續排程 finalizeScan
          ScriptApp.getProjectTriggers().forEach(t => {
            if (t.getHandlerFunction() === 'finalizeScan') ScriptApp.deleteTrigger(t);
          });
          ScriptApp.newTrigger('finalizeScan').timeBased().after(90 * 1000).create();
          Logger.log(`Gemini 分析未完，排程繼續（剩 ${merged.length - nextOffset} 檔），90 秒後繼續`);
          return; // 先結束，讓下一批繼續
        }

        Logger.log('Gemini 分析全部完成');
      }

      // 重新排序（加入 Gemini 分數後；額度用完的情況下未分析個股 geminiTotal 視為 0）
      merged.sort((a, b) =>
        (b.priority - a.priority) ||
        ((b.geminiTotal || 0) - (a.geminiTotal || 0)) ||
        ((b.rs || 0) - (a.rs || 0))
      );
    }
  }

  // 載入股票名稱
  try {
    const names = loadStockNames();
    merged.forEach(r => { if (!r.name) r.name = names[r.id] || ''; });
  } catch(e) { Logger.log('載入股票名稱失敗：' + e.message); }

  // 儲存最終結果
  try {
    saveScanResultsToSheet(merged);
    Logger.log(`ScanResults 已儲存，共 ${merged.length} 筆`);
  } catch(e) {
    Logger.log('saveScanResultsToSheet 失敗：' + e.message);
  }

  // 更新進度為完成
  try {
    progress.done = true;
    progress.resultCount = merged.length;
    progress.completedAt = new Date().toISOString();
    progress.geminiOffset = 0;
    saveScanProgress(progress);
    // 清空 ScanTemp（結果已存到 ScanResults）
    try { getOrCreateSheet('ScanTemp').clearContents(); } catch(e) {}
  } catch(e) {
    Logger.log('標記完成失敗：' + e.message);
  }

  // 寄 Email
  try {
    sendScanEmail(merged, {
      marketPass: progress.marketPass,
      stabilization: progress.stabilization,
      bearishWarning: progress.bearishWarning,
      totalScanned,
      completedAt: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    });
    Logger.log('Email 已寄出');
  } catch(e) {
    Logger.log('sendScanEmail 失敗：' + e.message);
  }

  Logger.log('=== finalizeScan 全部完成 ===');
}

/**
 * 設定接力 Trigger（在 N 分鐘後執行 runScanBatch）
 */
function scheduleNextBatch(batchIndex) {
  // 刪除舊的接力 Trigger
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runScanBatch' &&
        t.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 設定新 Trigger
  ScriptApp.newTrigger('runScanBatch')
    .timeBased()
    .after(CONFIG.SCAN_INTERVAL_MIN * 60 * 1000)
    .create();

  Logger.log(`已排程第 ${batchIndex + 1} 批，${CONFIG.SCAN_INTERVAL_MIN} 分鐘後執行`);
}

// ══════════════════════════════════════════════
// 第一階段：全市場快速初篩
// ══════════════════════════════════════════════

let _phase1AdvDec = { adv: 0, dec: 0 };

// ══════════════════════════════════════════════
// 大盤判斷：加權指數 vs 20日均線（TWSE FMTQIK，免費）
// ══════════════════════════════════════════════
function checkMarketTrend() {
  try {
    const closes = [];
    const now = new Date();
    // 抓本月 + 上月的大盤日資料（確保滿 20 個交易日）
    for (let m = 1; m >= 0; m--) {
      const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      const dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'yyyyMMdd');
      const url = `https://www.twse.com.tw/rwd/zh/afterTrading/FMTQIK?date=${dateStr}&response=json`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      const json = JSON.parse(res.getContentText());
      if (json.stat !== 'OK' || !json.data) continue;
      // 欄位：[日期, 成交股數, 成交金額, 成交筆數, 發行量加權股價指數, 漲跌點數]
      json.data.forEach(row => {
        const idx = parseFloat(String(row[4]).replace(/,/g, ''));
        if (idx > 0) closes.push(idx);
      });
    }
    if (closes.length < 20) {
      Logger.log(`checkMarketTrend：資料不足（${closes.length} 日），預設謹慎`);
      return { pass: false, index: closes[closes.length-1] || 0, ma20: 0 };
    }
    const last = closes[closes.length - 1];
    const ma20 = closes.slice(-20).reduce((a, b) => a + b, 0) / 20;
    const pass = last >= ma20;
    Logger.log(`checkMarketTrend：指數 ${last.toFixed(0)} vs MA20 ${ma20.toFixed(0)} → ${pass ? '多頭' : '空頭'}`);
    return { pass, index: Math.round(last), ma20: Math.round(ma20) };
  } catch(e) {
    Logger.log('checkMarketTrend 失敗：' + e.message);
    return { pass: false, index: 0, ma20: 0 };
  }
}

// ══════════════════════════════════════════════
// 止穩訊號 Checklist（5項自動判定，7/17崩盤後建立的框架）
// ⚠️ 這個函式抓了4個不同的外部資料源，單次執行可能要10~20秒，
//    建議只在每天第一批（batchIndex===0）呼叫一次，不要每檔股票都呼叫。
// ⚠️ ③的「大盤融資單日大減」門檻是用 FinMind TodayBalance/YesBalance
//    的原始差值去比對，實際單位（元 or 張）未經真實資料驗證，
//    第一次跑完後請對照 Logger.log 印出的原始數字，用真實案例校正
//    STABILIZATION_MARGIN_DROP_THRESHOLD 這個常數。
// ⚠️ ⑤的「爆量長下影」只用成交金額(1.3兆+)當代理指標，沒有驗證
//    K棒下影線的實際形狀（TWSE免費API拿不到大盤指數的日內高低點），
//    這條件成立時建議自己肉眼確認一下K線圖。
// ══════════════════════════════════════════════
const STABILIZATION_MARGIN_DROP_THRESHOLD = 15000000000; // 150億，單位需依實際資料校正

function checkStabilizationSignals() {
  const items = [];

  // ① 台積電（2330）連2日收盤不破前低且站回5日線
  try {
    const start = getDaysAgo(20);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=2330&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const data = (json.data || []).filter(d => d.close > 0);
    if (data.length >= 7) {
      const closes = data.map(d => d.close);
      const last = closes[closes.length - 1], prev = closes[closes.length - 2];
      const priorLow = Math.min(...closes.slice(-7, -2)); // 前低：再往前5天的最低收盤
      const ma5 = calcMA_GAS(closes, 5);
      const pass = last >= priorLow && prev >= priorLow && last >= ma5;
      items.push({ key: '①', name: '台積電止穩', pass, detail: `收盤${last} / 前低${priorLow.toFixed(1)} / 5MA${ma5.toFixed(1)}` });
    } else {
      items.push({ key: '①', name: '台積電止穩', pass: false, detail: '資料不足' });
    }
  } catch(e) { items.push({ key: '①', name: '台積電止穩', pass: false, detail: '抓取失敗：' + e.message }); }

  // ② 外資單日賣超收斂至300億內（整體市場三大法人買賣表，過濾「外資」列）
  try {
    const start = getDaysAgo(5);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockTotalInstitutionalInvestors&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const rows = (json.data || []).filter(d => (d.name || '').indexOf('外資') !== -1);
    if (rows.length > 0) {
      const last = rows[rows.length - 1];
      const net = (last.buy || 0) - (last.sell || 0); // 負值＝賣超
      const pass = net >= -30000000000; // 賣超收斂至300億內，或已轉買超
      items.push({ key: '②', name: '外資賣超收斂', pass, detail: `淨額約 ${(net / 1e8).toFixed(1)} 億` });
    } else {
      items.push({ key: '②', name: '外資賣超收斂', pass: false, detail: '無資料' });
    }
  } catch(e) { items.push({ key: '②', name: '外資賣超收斂', pass: false, detail: '抓取失敗：' + e.message }); }

  // ③ 大盤融資單日大減150億以上（整體市場融資融券表，過濾「融資」列）
  try {
    const start = getDaysAgo(5);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockTotalMarginPurchaseShortSale&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const rows = (json.data || []).filter(d => (d.name || '').indexOf('融資') !== -1);
    if (rows.length > 0) {
      const last = rows[rows.length - 1];
      const diff = (last.TodayBalance || 0) - (last.YesBalance || 0); // 負值＝減少
      Logger.log(`止穩訊號③ 原始數字 TodayBalance=${last.TodayBalance} YesBalance=${last.YesBalance} diff=${diff}`);
      const pass = diff <= -STABILIZATION_MARGIN_DROP_THRESHOLD;
      items.push({ key: '③', name: '融資大減斷頭出清', pass, detail: `單日變化約 ${(diff / 1e8).toFixed(1)} 億（單位待校正）` });
    } else {
      items.push({ key: '③', name: '融資大減斷頭出清', pass: false, detail: '無資料' });
    }
  } catch(e) { items.push({ key: '③', name: '融資大減斷頭出清', pass: false, detail: '抓取失敗：' + e.message }); }

  // ④ 費半（SOX）連2日不創新低或單日彈2%+收高
  try {
    const start = getDaysAgo(15);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=USStockPrice&data_id=^SOX&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const data = (json.data || []).filter(d => d.Close > 0);
    if (data.length >= 5) {
      const closes = data.map(d => d.Close);
      const last = closes[closes.length - 1], prev = closes[closes.length - 2];
      const priorLow = Math.min(...closes.slice(0, -2));
      const bigBounce = prev > 0 && (last - prev) / prev >= 0.02 && last > prev;
      const notNewLow2Days = last >= priorLow && prev >= priorLow;
      const pass = bigBounce || notNewLow2Days;
      items.push({ key: '④', name: '費半止跌', pass, detail: `收盤 ${last.toFixed(0)} / 前波低點 ${priorLow.toFixed(0)}` });
    } else {
      items.push({ key: '④', name: '費半止跌', pass: false, detail: '資料不足' });
    }
  } catch(e) { items.push({ key: '④', name: '費半止跌', pass: false, detail: '抓取失敗：' + e.message }); }

  // ⑤ 大盤爆量（1.3兆+，成交金額代理指標）或收復43500站穩2日
  try {
    const closes = [], amounts = [];
    const now = new Date();
    for (let m = 1; m >= 0; m--) {
      const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      const dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'yyyyMMdd');
      const url = `https://www.twse.com.tw/rwd/zh/afterTrading/FMTQIK?date=${dateStr}&response=json`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      const json = JSON.parse(res.getContentText());
      if (json.stat !== 'OK' || !json.data) continue;
      json.data.forEach(row => {
        const idx = parseFloat(String(row[4]).replace(/,/g, ''));
        const amt = parseFloat(String(row[2]).replace(/,/g, ''));
        if (idx > 0) { closes.push(idx); amounts.push(amt); }
      });
    }
    if (closes.length >= 3) {
      const lastIdx = closes[closes.length - 1], lastAmt = amounts[amounts.length - 1];
      const prevIdx = closes[closes.length - 2];
      const bigVolume = lastAmt >= 1.3e12;
      const holds43500 = lastIdx >= 43500 && prevIdx >= 43500;
      const pass = bigVolume || holds43500;
      items.push({ key: '⑤', name: '大盤反轉確認', pass, detail: `指數 ${lastIdx.toFixed(0)} / 金額 ${(lastAmt / 1e12).toFixed(2)}兆` });
    } else {
      items.push({ key: '⑤', name: '大盤反轉確認', pass: false, detail: '資料不足' });
    }
  } catch(e) { items.push({ key: '⑤', name: '大盤反轉確認', pass: false, detail: '抓取失敗：' + e.message }); }

  const count = items.filter(i => i.pass).length;
  let verdict;
  if (count >= 3) verdict = `${count}/5 成立 → 可投入第2批`;
  else if (count >= 2) verdict = `${count}/5 成立 → 可投入第1批`;
  else verdict = `${count}/5 成立 → 訊號未到，繼續觀望`;

  return { count, total: 5, items, verdict };
}

// ══════════════════════════════════════════════
// 空頭警示系統：大盤空頭確認 + 庫存清單個股警示
// ⚠️ 跟止穩訊號一樣，③的融資單位、⑤的爆量下跌都是同樣的估算限制，
//    需要你跑過真實資料後校正 BEARISH_MARGIN_SURGE_THRESHOLD。
// ══════════════════════════════════════════════
const BEARISH_MARGIN_SURGE_THRESHOLD = 10000000000; // 100億，單位需依實際資料校正

// 個股空頭訊號（checkG的鏡像，但刻意放寬——這是「提醒你注意」不是「已經崩盤才通知」，
// 所以不要求5項全過，只要「均線結構轉空」＋「動能指標至少2項同向轉弱」就示警，
// 這樣才能在真正大跌之前提早提醒，而不是等到損失已經發生才後知後覺）
function checkBearishExit(ohlc, vols) {
  if (ohlc.length < 65) return { warn: false, detail: '資料不足' };
  const cl = ohlc.map(d => d.close);
  const lastClose = cl[cl.length - 1];

  const ma5 = calcMA_GAS(cl, 5), ma10 = calcMA_GAS(cl, 10), ma20 = calcMA_GAS(cl, 20);
  if (!ma5 || !ma10 || !ma20) return { warn: false, detail: '均線資料不足' };
  // 結構性轉空：收盤跌破5MA且5MA<10MA<20MA（不強制60MA，避免早期轉空訊號被壓過久才示警）
  const maBearish = lastClose < ma5 && ma5 < ma10 && ma10 < ma20;

  const rsis = calcRSISeries_GAS(cl, 14);
  let rsiWeak = false, rsiToday = null;
  if (rsis.length >= 2) {
    rsiToday = rsis[rsis.length - 1];
    rsiWeak = rsiToday < 50 && rsiToday < rsis[rsis.length - 2];
  }

  const kd = calcKDSeries_GAS(ohlc, 9);
  let kdDeathCross = false, kToday = null, dToday = null;
  if (kd.k.length >= 1) {
    kToday = kd.k[kd.k.length - 1];
    dToday = kd.d[kd.d.length - 1];
    kdDeathCross = kToday < dToday && kToday <= 50 && dToday <= 50;
  }

  const avgVol5 = vols.slice(-6, -1).reduce((a, b) => a + b, 0) / 5;
  const today = ohlc[ohlc.length - 1];
  const volSpikeDown = vols[vols.length - 1] > avgVol5 && today.close < today.open;

  const macd = calcMACDSeries_GAS(cl);
  let macdBearish = false;
  if (macd.dif.length >= 2) {
    const difToday = macd.dif[macd.dif.length - 1];
    const histToday = macd.hist[macd.hist.length - 1];
    macdBearish = difToday < 0 && histToday < 0 && histToday < macd.hist[macd.hist.length - 2];
  }

  const momentumHits = [rsiWeak, kdDeathCross, volSpikeDown, macdBearish].filter(Boolean).length;
  const warn = maBearish && momentumHits >= 2;

  return {
    warn, maBearish, rsiWeak, kdDeathCross, volSpikeDown, macdBearish, momentumHits,
    detail: `均線${maBearish ? '空頭排列' : '尚未轉空'} / RSI${rsiToday !== null ? rsiToday.toFixed(0) : '-'} / K${kToday !== null ? kToday.toFixed(0) : '-'}D${dToday !== null ? dToday.toFixed(0) : '-'} / 動能轉弱${momentumHits}/4項`
  };
}

// 大盤空頭確認 Checklist（止穩訊號的鏡像，5項訊號判定空軍是否增加）
function checkMarketBearishSignals() {
  const items = [];

  // ① 台積電轉弱：連2日收盤跌破前高且跌破5日線
  try {
    const start = getDaysAgo(20);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=2330&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const data = (json.data || []).filter(d => d.close > 0);
    if (data.length >= 7) {
      const closes = data.map(d => d.close);
      const last = closes[closes.length - 1], prev = closes[closes.length - 2];
      const priorHigh = Math.max(...closes.slice(-7, -2));
      const ma5 = calcMA_GAS(closes, 5);
      const warn = last < priorHigh && prev < priorHigh && last < ma5 && prev < ma5;
      items.push({ key: '①', name: '台積電轉弱跌破5日線', warn, detail: `收盤${last} / 前高${priorHigh.toFixed(1)} / 5MA${ma5.toFixed(1)}` });
    } else {
      items.push({ key: '①', name: '台積電轉弱跌破5日線', warn: false, detail: '資料不足' });
    }
  } catch(e) { items.push({ key: '①', name: '台積電轉弱跌破5日線', warn: false, detail: '抓取失敗：' + e.message }); }

  // ② 外資單日賣超擴大超過300億
  try {
    const start = getDaysAgo(5);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockTotalInstitutionalInvestors&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const rows = (json.data || []).filter(d => (d.name || '').indexOf('外資') !== -1);
    if (rows.length > 0) {
      const last = rows[rows.length - 1];
      const net = (last.buy || 0) - (last.sell || 0);
      const warn = net <= -30000000000;
      items.push({ key: '②', name: '外資賣超擴大', warn, detail: `淨額約 ${(net / 1e8).toFixed(1)} 億` });
    } else {
      items.push({ key: '②', name: '外資賣超擴大', warn: false, detail: '無資料' });
    }
  } catch(e) { items.push({ key: '②', name: '外資賣超擴大', warn: false, detail: '抓取失敗：' + e.message }); }

  // ③ 大盤融資單日大增（散戶還在追高，通常是頭部警訊）
  try {
    const start = getDaysAgo(5);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockTotalMarginPurchaseShortSale&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const rows = (json.data || []).filter(d => (d.name || '').indexOf('融資') !== -1);
    if (rows.length > 0) {
      const last = rows[rows.length - 1];
      const diff = (last.TodayBalance || 0) - (last.YesBalance || 0);
      const warn = diff >= BEARISH_MARGIN_SURGE_THRESHOLD;
      items.push({ key: '③', name: '融資單日大增(散戶追高)', warn, detail: `單日變化約 ${(diff / 1e8).toFixed(1)} 億（單位待校正）` });
    } else {
      items.push({ key: '③', name: '融資單日大增(散戶追高)', warn: false, detail: '無資料' });
    }
  } catch(e) { items.push({ key: '③', name: '融資單日大增(散戶追高)', warn: false, detail: '抓取失敗：' + e.message }); }

  // ④ 費半連2日創新低或單日重挫2%+收低
  try {
    const start = getDaysAgo(15);
    const res = UrlFetchApp.fetch(
      `https://api.finmindtrade.com/api/v4/data?dataset=USStockPrice&data_id=^SOX&start_date=${start}&token=${CONFIG.FINMIND_TOKEN}`,
      { muteHttpExceptions: true }
    );
    const json = JSON.parse(res.getContentText());
    const data = (json.data || []).filter(d => d.Close > 0);
    if (data.length >= 5) {
      const closes = data.map(d => d.Close);
      const last = closes[closes.length - 1], prev = closes[closes.length - 2];
      const priorLow = Math.min(...closes.slice(0, -2));
      const bigDrop = prev > 0 && (last - prev) / prev <= -0.02 && last < prev;
      const newLow2Days = last <= priorLow && prev <= priorLow;
      const warn = bigDrop || newLow2Days;
      items.push({ key: '④', name: '費半重挫或破底', warn, detail: `收盤 ${last.toFixed(0)} / 前低 ${priorLow.toFixed(0)}` });
    } else {
      items.push({ key: '④', name: '費半重挫或破底', warn: false, detail: '資料不足' });
    }
  } catch(e) { items.push({ key: '④', name: '費半重挫或破底', warn: false, detail: '抓取失敗：' + e.message }); }

  // ⑤ 跌破月線(20MA)且爆量下跌（成交金額代理，同樣需要肉眼確認K棒形狀）
  try {
    const closes = [], amounts = [];
    const now = new Date();
    for (let m = 1; m >= 0; m--) {
      const d = new Date(now.getFullYear(), now.getMonth() - m, 1);
      const dateStr = Utilities.formatDate(d, 'Asia/Taipei', 'yyyyMMdd');
      const url = `https://www.twse.com.tw/rwd/zh/afterTrading/FMTQIK?date=${dateStr}&response=json`;
      const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (res.getResponseCode() !== 200) continue;
      const json = JSON.parse(res.getContentText());
      if (json.stat !== 'OK' || !json.data) continue;
      json.data.forEach(row => {
        const idx = parseFloat(String(row[4]).replace(/,/g, ''));
        const amt = parseFloat(String(row[2]).replace(/,/g, ''));
        if (idx > 0) { closes.push(idx); amounts.push(amt); }
      });
    }
    if (closes.length >= 20) {
      const lastIdx = closes[closes.length - 1], lastAmt = amounts[amounts.length - 1];
      const ma20 = calcMA_GAS(closes, 20);
      const belowMA20 = lastIdx < ma20;
      const bigVolume = lastAmt >= 1.3e12;
      const warn = belowMA20 && bigVolume;
      items.push({ key: '⑤', name: '跌破月線且爆量下跌', warn, detail: `指數 ${lastIdx.toFixed(0)} / 20MA ${ma20.toFixed(0)} / 金額 ${(lastAmt / 1e12).toFixed(2)}兆` });
    } else {
      items.push({ key: '⑤', name: '跌破月線且爆量下跌', warn: false, detail: '資料不足' });
    }
  } catch(e) { items.push({ key: '⑤', name: '跌破月線且爆量下跌', warn: false, detail: '抓取失敗：' + e.message }); }

  const count = items.filter(i => i.warn).length;
  let verdict;
  if (count >= 3) verdict = `${count}/5 成立 → 空頭訊號強烈，建議停利出場，暫停加碼`;
  else if (count >= 2) verdict = `${count}/5 成立 → 空軍增加中，提高警覺，準備停利`;
  else verdict = `${count}/5 成立 → 暫無明顯空頭訊號`;

  return { count, total: 5, items, verdict };
}

// 從 Trades 工作表推算目前實際持股（買進累加、賣出扣除，只回傳淨部位>0的）
function getCurrentHoldings_() {
  const trades = loadTrades();
  const positions = {};
  trades.forEach(t => {
    const id = String(t.id || '').trim();
    if (!id) return;
    const shares = Number(t.shares) || 0;
    const action = String(t.action || '').toLowerCase();
    if (!positions[id]) positions[id] = { id, name: t.name || '', netShares: 0 };
    if (action === 'buy') positions[id].netShares += shares;
    else if (action === 'sell') positions[id].netShares -= shares;
    if (t.name) positions[id].name = t.name;
  });
  return Object.values(positions).filter(p => p.netShares > 0);
}

// 整合入口：大盤空頭確認 + 庫存清單逐檔檢查
function checkPortfolioBearishWarning() {
  const holdings = getCurrentHoldings_();
  const results = [];

  holdings.forEach(h => {
    try {
      const start90 = getDaysAgo(90);
      const res = UrlFetchApp.fetch(
        `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=${h.id}&start_date=${start90}&token=${CONFIG.FINMIND_TOKEN}`,
        { muteHttpExceptions: true }
      );
      const json = JSON.parse(res.getContentText());
      if (json.status !== 200 || !json.data || json.data.length < 65) {
        results.push({ id: h.id, name: h.name, netShares: h.netShares, warn: false, detail: '資料不足' });
        return;
      }
      const ohlc = json.data;
      const vols = ohlc.map(d => d.Trading_Volume || 0);
      const signal = checkBearishExit(ohlc, vols);
      results.push(Object.assign({ id: h.id, name: h.name, netShares: h.netShares }, signal));
    } catch(e) {
      results.push({ id: h.id, name: h.name, netShares: h.netShares, warn: false, detail: '抓取失敗：' + e.message });
    }
    Utilities.sleep(500);
  });

  const market = checkMarketBearishSignals();
  const warnedHoldings = results.filter(r => r.warn);

  return { market, holdings: results, warnedHoldings, warnedCount: warnedHoldings.length, totalHoldings: holdings.length };
}

function runPhase1() {
  Logger.log('Phase1 開始，使用 TWSE Open API 抓全市場當日資料');

  // TWSE Open API：不需要 Token，免費，一次回傳全市場
  // 上市股票
  const twseUrl = 'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL';
  // 上櫃股票
  const tpexUrl = 'https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes';

  const allData = [];

  try {
    const res1 = UrlFetchApp.fetch(twseUrl, { muteHttpExceptions: true });
    if (res1.getResponseCode() === 200) {
      const data1 = JSON.parse(res1.getContentText());
      allData.push(...(Array.isArray(data1) ? data1 : []));
      Logger.log(`TWSE 上市：${allData.length} 筆`);
    }
  } catch(e) {
    Logger.log('TWSE API 失敗：' + e.message);
  }

  try {
    const res2 = UrlFetchApp.fetch(tpexUrl, { muteHttpExceptions: true });
    if (res2.getResponseCode() === 200) {
      const data2 = JSON.parse(res2.getContentText());
      const tpexData = Array.isArray(data2) ? data2 : [];
      allData.push(...tpexData);
      Logger.log(`TPEx 上櫃：${tpexData.length} 筆，合計：${allData.length} 筆`);
    }
  } catch(e) {
    Logger.log('TPEx API 失敗：' + e.message);
  }

  if (allData.length === 0) {
    Logger.log('❌ 無資料，可能今天休市');
    return [];
  }

  // 統計全市場漲跌家數（Change 欄位）
  let advCount = 0, decCount = 0;
  allData.forEach(d => {
    const chg = parseFloat(String(d.Change || '0').replace(/,/g, ''));
    if (chg > 0) advCount++;
    else if (chg < 0) decCount++;
  });
  _phase1AdvDec = { adv: advCount, dec: decCount };
  Logger.log(`漲跌家數：漲 ${advCount} / 跌 ${decCount}`);

  const passed = [];

  allData.forEach(d => {
    // TWSE 欄位：Code, Name, TradeVolume, TradeValue, OpeningPrice, HighestPrice, LowestPrice, ClosingPrice, Change, Transaction
    // TPEx 欄位：SecuritiesCompanyCode, CompanyName, Close, Change, Open, High, Low, Volumn, TotalValue
    const id = d.Code || d.SecuritiesCompanyCode || '';
    if (!id) return;
    if (/[A-Za-z]/.test(id)) return;
    if (!/^\d{4,5}$/.test(id)) return;

    const name = d.Name || d.CompanyName || ''; // TWSE: Name, TPEx: CompanyName
    const close = parseFloat(d.ClosingPrice || d.Close || 0);
    const open  = parseFloat(d.OpeningPrice || d.Open || close);
    const high  = parseFloat(d.HighestPrice || d.High || close);
    const low   = parseFloat(d.LowestPrice  || d.Low  || close);
    // 成交量（TWSE 用千股，TPEx 用張）
    const volRaw = d.TradeVolume || d.Volumn || '0';
    const vol = parseInt(String(volRaw).replace(/,/g, '')) || 0;
    const amount = parseFloat(String(d.TradeValue || d.TotalValue || '0').replace(/,/g, '')) || (close * vol);

    if (!close || close < 5) return;
    if (vol === 0) return;
    if (amount < 5000000) return;
    if (close < open) return; // 收紅初篩

    let tier = 'mid';
    if (amount > 5e8) tier = 'large';
    else if (amount < 5e7) tier = 'small';

    passed.push({ id, name, close, open, high, low, vol, amount, tier });
  });

  Logger.log(`第一階段：${allData.length} 檔 → ${passed.length} 檔通過初篩`);
  return passed;
}


// ══════════════════════════════════════════════
// GEMINI AI 分析（消息面 / 券商目標價 / 基本面）
// ══════════════════════════════════════════════

/**
 * 對所有符合條件的個股執行 Gemini 深度分析
 * 結果直接寫入 result 物件
 */
/**
 * @returns {boolean} 是否在本次呼叫中遇到額度用完（QUOTA_EXCEEDED）
 */
function runGeminiAnalysis(results) {
  const names = loadStockNames();

  // 只分析優先級 4 星以上的個股，節省額度
  const highPriority = results.filter(r => r.priority >= 4);
  Logger.log(`Gemini 分析：共 ${results.length} 檔，其中 ${highPriority.length} 檔優先級 ≥ 4 星需要分析`);

  for (let i = 0; i < highPriority.length; i++) {
    const r = highPriority[i];
    try {
      Logger.log(`Gemini 分析 ${i + 1}/${highPriority.length}: ${r.id}`);
      const name = names[r.id] || r.id;

      // 三合一分析（1次呼叫取代原本3次）
      const result = geminiAnalyzeAll(r.id, name, r.entryPrice);

      r.newsAnalysis   = result.news;
      r.brokerAnalysis = result.broker;
      r.fundAnalysis   = result.fund;

      const total = (result.news.score || 0) + (result.broker.score || 0) + (result.fund.score || 0);
      r.geminiTotal = total;
      r.geminiBonus = total >= 4 ? 2 : total >= 2 ? 1 : 0;
      if (r.geminiBonus > 0) r.priority = r.priority + r.geminiBonus; // 無上限

      Utilities.sleep(2000);
    } catch(e) {
      if (e.message === 'QUOTA_EXCEEDED') {
        Logger.log(`Gemini 額度已用完，停止本次與後續分析（已處理 ${i}/${highPriority.length} 檔）`);
        return true; // 真正中止迴圈，不再呼叫
      }
      Logger.log(`${r.id} Gemini 分析失敗：${e.message}`);
      r.geminiTotal = 0;
      r.geminiBonus = 0;
    }
  }
  return false;
}

/**
 * 三合一分析：一次呼叫同時取得消息面、券商目標價、基本面
 * 節省 66% Gemini 額度（3次→1次）
 */
function geminiAnalyzeAll(stockId, stockName, currentPrice) {
  const today = getTodayStr();
  const prompt = `今天是 ${today}。請使用 Google Search 搜尋台灣上市股票「${stockId} ${stockName}」的最新資訊，進行以下三個面向的分析。

目前股價：${currentPrice} 元

【面向一：消息面】搜尋近 14 天重大新聞、法說會、訂單、產業動向
【面向二：券商目標價】搜尋近 6 個月各大券商目標價與評級（富邦、國泰、元大、凱基、摩根士丹利等）
【面向三：基本面】搜尋近四季 EPS、月營收趨勢、本益比、毛利率、財務健康

各面向評分規則（score 必須是 -2、-1、0、1、2 其中一個整數）：
消息面：+2 明確重大利多 / 0 中性或資料不足 / -2 明確重大利空
券商目標：+2 目標價空間>20%且多數買進 / 0 資料不足 / -2 主要券商下調
基本面：+2 EPS成長+營收成長+財務健康 / 0 好壞參半或資料不足 / -2 EPS虧損或大幅衰退

若搜尋不到資訊或不確定，score 給 0，summary 標示「資料不足」。

只回傳以下 JSON，不要任何其他文字：
{
  "news": {"score":0,"summary":"2句繁體中文","keyNews":["消息1","消息2"],"riskAlert":"風險或無明顯利空"},
  "broker": {"score":0,"targetMedian":null,"targetHigh":null,"targetLow":null,"upsidePct":null,"consensusRating":"資料不足","recentChange":"資料不足","brokerList":[],"summary":"2句繁體中文"},
  "fund": {"score":0,"epsGrowth":"成長","revenueGrowth":"成長","peRatio":null,"peAssessment":"合理","marginTrend":"持平","financialHealth":"健康","industryCycle":"擴張","summary":"2句繁體中文"}
}`;

  const empty = {
    news:   { score: 0, summary: '─', keyNews: [], riskAlert: '─' },
    broker: { score: 0, targetMedian: null, consensusRating: '資料不足', recentChange: '資料不足', brokerList: [], summary: '─' },
    fund:   { score: 0, epsGrowth: '─', revenueGrowth: '─', peRatio: null, peAssessment: '─', marginTrend: '─', financialHealth: '─', industryCycle: '─', summary: '─' }
  };

  try {
    const raw = callGeminiGAS(prompt);
    const parsed = JSON.parse(raw);
    // 確保 score 都是整數
    ['news','broker','fund'].forEach(k => {
      if (parsed[k]) parsed[k].score = Math.max(-2, Math.min(2, parseInt(parsed[k].score) || 0));
    });
    return {
      news:   { ...empty.news,   ...parsed.news   },
      broker: { ...empty.broker, ...parsed.broker },
      fund:   { ...empty.fund,   ...parsed.fund   }
    };
  } catch(e) {
    if (e.message === 'QUOTA_EXCEEDED') throw e;
    Logger.log(`geminiAnalyzeAll 解析失敗：${e.message}`);
    return empty;
  }
}

/**
 * 呼叫 Gemini API（帶 Google Search grounding，用到免費額度上限為止）
 */
function callGeminiGAS(prompt) {
  const models = [
    { model: 'gemini-2.5-flash',      api: 'v1beta' },
    { model: 'gemini-2.5-flash-lite', api: 'v1beta' },
  ];

  for (const { model, api } of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/${api}/models/${model}:generateContent?key=${CONFIG.GEMINI_KEY}`;
      const payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 1024 },
        tools: [{ googleSearch: {} }],
      };

      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = res.getResponseCode();
      const body = res.getContentText();

      Logger.log(`Gemini ${model} → ${code}: ${body.slice(0, 200)}`);

      if (code === 429 || body.includes('RESOURCE_EXHAUSTED') || body.includes('quota')) {
        throw new Error('QUOTA_EXCEEDED');
      }

      if (code !== 200) {
        Logger.log(`Gemini ${model} 非200：${body.slice(0, 300)}`);
        continue;
      }

      const json = JSON.parse(body);
      const parts = (json.candidates || [])[0]?.content?.parts || [];
      const text = parts.map(p => p.text || '').join('');
      if (!text) continue;
      return extractJson(text);

    } catch(e) {
      if (e.message === 'QUOTA_EXCEEDED') throw e;
      Logger.log(`模型 ${model} 失敗：${e.message}`);
      continue;
    }
  }
  throw new Error('所有 Gemini 模型均失敗');
}

/**
 * 從 Gemini 回應中提取 JSON
 */
function extractJson(text) {
  // 移除 markdown
  let clean = text.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
  const start = clean.indexOf('{');
  const end = clean.lastIndexOf('}');
  if (start !== -1 && end !== -1) clean = clean.slice(start, end + 1);
  return JSON.parse(clean);
}

// ══════════════════════════════════════════════
// 第二階段：完整技術分析（單批）
// ══════════════════════════════════════════════

function runPhase2Batch(stockIds, nameMap) {
  nameMap = nameMap || {};
  const today = getTodayStr();
  const start90 = getDaysAgo(90);
  const start15 = getDaysAgo(15);
  const results = [];

  stockIds.forEach(id => {
    try {
      // 抓 90 日 OHLCV
      const ohlcRes = UrlFetchApp.fetch(
        `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&data_id=${id}&start_date=${start90}&token=${CONFIG.FINMIND_TOKEN}`,
        { muteHttpExceptions: true }
      );
      const ohlcJson = JSON.parse(ohlcRes.getContentText());
      if (ohlcJson.status !== 200 || !ohlcJson.data?.length) return;

      const ohlc = ohlcJson.data;
      if (ohlc.length < 20) return;

      const closes = ohlc.map(d => d.close);
      const volumes = ohlc.map(d => d.Trading_Volume || 0);
      const lastClose = closes[closes.length - 1];

      // 排除條件
      const todayVol = volumes[volumes.length - 1];
      const todayAmount = lastClose * todayVol;
      if (todayAmount < 5000000) return;

      const recent5 = ohlc.slice(-5);
      const hasLimit = recent5.some(d => (d.close - d.open) / (d.open || 1) < -0.095);
      if (hasLimit) return;

      // 籌碼資料
      let itrustArr = [], foreignArr = [];
      try {
        const instRes = UrlFetchApp.fetch(
          `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInstitutionalInvestorsBuySell&data_id=${id}&start_date=${start15}&token=${CONFIG.FINMIND_TOKEN}`,
          { muteHttpExceptions: true }
        );
        const instJson = JSON.parse(instRes.getContentText());
        if (instJson.status === 200 && instJson.data) {
          itrustArr = instJson.data.filter(d => d.name === '投信').map(d => d.buy - d.sell);
          foreignArr = instJson.data.filter(d => d.name === '外資').map(d => d.buy - d.sell);
        }
      } catch(e) {}

      // 融資融券排除層：使用率過高(>85%)或單週暴增(>30%)代表籌碼過熱，直接排除
      // （對應合晶6182案例：單週融資暴增近2萬張後最終走空）
      let marginSignal = null;
      try {
        const marginRes = UrlFetchApp.fetch(
          `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockMarginPurchaseShortSale&data_id=${id}&start_date=${start15}&token=${CONFIG.FINMIND_TOKEN}`,
          { muteHttpExceptions: true }
        );
        const marginJson = JSON.parse(marginRes.getContentText());
        if (marginJson.status === 200 && marginJson.data && marginJson.data.length >= 6) {
          const md = marginJson.data.slice().sort((a, b) => a.date < b.date ? -1 : 1);
          const mToday = md[md.length - 1];
          const mWeekAgo = md[md.length - 6];
          const usageRatio = mToday.MarginPurchaseLimit > 0
            ? mToday.MarginPurchaseTodayBalance / mToday.MarginPurchaseLimit : null;
          const weeklyChangeRatio = mWeekAgo.MarginPurchaseTodayBalance > 0
            ? (mToday.MarginPurchaseTodayBalance - mWeekAgo.MarginPurchaseTodayBalance) / mWeekAgo.MarginPurchaseTodayBalance
            : null;
          marginSignal = { usageRatio, weeklyChangeRatio };
          const exclude = (usageRatio !== null && usageRatio > 0.85) ||
                           (weeklyChangeRatio !== null && weeklyChangeRatio > 0.30);
          if (exclude) return; // 融資排除層：籌碼過熱，跳過這檔
        }
      } catch(e) { Logger.log(`${id} 融資資料抓取失敗：${e.message}`); }

      // ATR
      const atr = calcATR_GAS(ohlc);
      if (!atr) return;

      // 市值分層（用成交金額估算）
      let tier = 'mid';
      if (todayAmount > 5e8) tier = 'large';
      else if (todayAmount < 5e7) tier = 'small';

      // 組合判斷
      const hitCombos = [];
      if (tier === 'small') {
        if (checkF(ohlc, volumes))               hitCombos.push({ combo: 'F', base: 4 });
        if (checkD(ohlc, volumes, tier))         hitCombos.push({ combo: 'D', base: 5 });
        if (checkA(ohlc, volumes, tier))         hitCombos.push({ combo: 'A', base: 4 });
        if (checkB(ohlc, itrustArr, foreignArr, tier)) hitCombos.push({ combo: 'B', base: 4 });
        if (checkC(ohlc, volumes, tier))         hitCombos.push({ combo: 'C', base: 3 });
        if (checkE(ohlc, volumes))               hitCombos.push({ combo: 'E', base: 2 });
        if (checkG(ohlc, volumes))               hitCombos.push({ combo: 'G', base: 6 });
      } else {
        if (checkB(ohlc, itrustArr, foreignArr, tier)) hitCombos.push({ combo: 'B', base: 5 });
        if (checkD(ohlc, volumes, tier))         hitCombos.push({ combo: 'D', base: 5 });
        if (checkA(ohlc, volumes, tier))         hitCombos.push({ combo: 'A', base: 4 });
        if (checkC(ohlc, volumes, tier))         hitCombos.push({ combo: 'C', base: 3 });
        if (checkE(ohlc, volumes))               hitCombos.push({ combo: 'E', base: 2 });
        if (checkG(ohlc, volumes))               hitCombos.push({ combo: 'G', base: 6 });
      }
      if (hitCombos.length === 0) return;

      // 融資加分層：融資週增率為負(籌碼退潮)且股價創20日新高＝法人推動而非融資堆疊的上漲
      let marginBonus = 0;
      if (marginSignal && marginSignal.weeklyChangeRatio !== null && marginSignal.weeklyChangeRatio < 0) {
        const high20 = Math.max(...ohlc.slice(-21, -1).map(d => d.max || d.high || d.close));
        if (lastClose > high20) marginBonus = 1;
      }

      const basePriority = Math.max(...hitCombos.map(c => c.base));
      const crossBonus = hitCombos.length >= 3 ? 2 : hitCombos.length === 2 ? 1 : 0;
      const priority = basePriority + crossBonus + marginBonus; // 無上限，有幾顆顯示幾顆
      const combos = hitCombos.map(c => c.combo);
      const combo = combos[0];
      const atrMult = { B: 2.0, D: 2.0, A: 1.5, C: 1.5, E: 1.0, F: 1.5, G: 2.0 }[combo] || 1.5;

      // 建議進場價：依組合類型，用技術回檔點而非直接用收盤價
      // A/D 突破型：等回檔 0.3~0.5 ATR 進場，避免追高
      // B 籌碼型：站上均線後，微幅回檔到收盤 -0.2 ATR
      // C 反轉型：反彈確認後，接近低點支撐
      // E 強勢型：不等回檔，直接用收盤（已是強勢慣性，回檔少）
      // F 小型爆量：保守，等 -0.3 ATR
      const entryDiscount = { A: 0.4, B: 0.2, C: 0.5, D: 0.35, E: 0.0, F: 0.3, G: 0.35 }[combo] || 0.3;
      const entryPrice = Math.round((lastClose - atr * entryDiscount) * 100) / 100;

      const stopLoss = Math.round((entryPrice - atr * atrMult) * 100) / 100;
      const target1 = Math.round((entryPrice + atr * atrMult * 2) * 100) / 100;
      const target2 = Math.round((entryPrice + atr * atrMult * 3) * 100) / 100;
      if (stopLoss >= entryPrice) return;

      // RS（相對強弱，暫時用5日漲幅代替）
      const rs = closes.length >= 5
        ? (lastClose - closes[closes.length - 6]) / (closes[closes.length - 6] || 1) : 0;

      results.push({
        id, name: nameMap[id] || '', tier, combo, combos, priority,
        hitCount: hitCombos.length, crossBonus, marginBonus,
        entryPrice, lastClose, stopLoss, target1, target2,
        atr: Math.round(atr * 100) / 100,
        chaseLimit: Math.round(lastClose * 1.02 * 100) / 100,  // 追價上限仍以收盤價為基準
        rs, itrustBuy: itrustArr.slice(-5).filter(v => v > 0).length,
        foreignBuy: foreignArr.slice(-3).filter(v => v > 0).length,
        marginUsage: marginSignal && marginSignal.usageRatio !== null ? Math.round(marginSignal.usageRatio * 1000) / 10 : null,
        marginWeeklyChange: marginSignal && marginSignal.weeklyChangeRatio !== null ? Math.round(marginSignal.weeklyChangeRatio * 1000) / 10 : null,
      });

    } catch(e) {
      Logger.log(`${id} 分析失敗：${e.message}`);
    }

    // 每檔之間稍微暫停，避免 FinMind rate limit
    Utilities.sleep(500);
  });

  return results;
}

// ══════════════════════════════════════════════
// GAS 版技術指標函式
// ══════════════════════════════════════════════

function calcMA_GAS(prices, period) {
  if (prices.length < period) return null;
  return prices.slice(-period).reduce((a, b) => a + b, 0) / period;
}

function calcATR_GAS(ohlc, period) {
  period = period || 14;
  if (ohlc.length < period + 1) return null;
  const trs = [];
  for (let i = 1; i < ohlc.length; i++) {
    const h = ohlc[i].max || ohlc[i].high || ohlc[i].close;
    const l = ohlc[i].min || ohlc[i].low || ohlc[i].close;
    const pc = ohlc[i-1].close;
    trs.push(Math.max(h - l, Math.abs(h - pc), Math.abs(l - pc)));
  }
  return trs.slice(-period).reduce((a, b) => a + b, 0) / period;
}

function getVolMult(tier) {
  return { large: 1.5, mid: 1.8, small: 2.0 }[tier] || 1.8;
}

// ── 日線趨勢篩選法（checkG）用的指標序列 ──────────

// RSI 序列（Wilder's smoothing），回傳整段序列方便比較「今天 vs 昨天」
function calcRSISeries_GAS(closes, period) {
  period = period || 14;
  if (closes.length < period + 1) return [];
  let gains = 0, losses = 0;
  for (let i = 1; i <= period; i++) {
    const diff = closes[i] - closes[i - 1];
    if (diff >= 0) gains += diff; else losses -= diff;
  }
  let avgGain = gains / period, avgLoss = losses / period;
  const rsis = [avgLoss === 0 ? 100 : 100 - 100 / (1 + avgGain / avgLoss)];
  for (let i = period + 1; i < closes.length; i++) {
    const diff = closes[i] - closes[i - 1];
    const gain = diff > 0 ? diff : 0;
    const loss = diff < 0 ? -diff : 0;
    avgGain = (avgGain * (period - 1) + gain) / period;
    avgLoss = (avgLoss * (period - 1) + loss) / period;
    rsis.push(avgLoss === 0 ? 100 : 100 - 100 / (1 + avgGain / avgLoss));
  }
  return rsis;
}

// KD 序列（RSV 9 日，K/D 用 1/3 平滑，台股慣用參數 9,3,3）
function calcKDSeries_GAS(ohlc, period) {
  period = period || 9;
  if (ohlc.length < period) return { k: [], d: [] };
  const kArr = [], dArr = [];
  let prevK = 50, prevD = 50;
  for (let i = period - 1; i < ohlc.length; i++) {
    const slice = ohlc.slice(i - period + 1, i + 1);
    const high = Math.max(...slice.map(d => d.max || d.high || d.close));
    const low = Math.min(...slice.map(d => d.min || d.low || d.close));
    const close = ohlc[i].close;
    const rsv = high === low ? 50 : (close - low) / (high - low) * 100;
    const k = prevK * 2 / 3 + rsv * 1 / 3;
    const d = prevD * 2 / 3 + k * 1 / 3;
    kArr.push(k); dArr.push(d);
    prevK = k; prevD = d;
  }
  return { k: kArr, d: dArr };
}

// EMA 序列，MACD 的基礎
function calcEMASeries_GAS(values, period) {
  const k = 2 / (period + 1);
  const emaArr = [values[0]];
  for (let i = 1; i < values.length; i++) {
    emaArr.push(values[i] * k + emaArr[i - 1] * (1 - k));
  }
  return emaArr;
}

// MACD 序列（12,26,9），dif=快慢EMA差, dea=dif的9日EMA, hist=柱狀體(dif-dea)
function calcMACDSeries_GAS(closes) {
  if (closes.length < 35) return { dif: [], dea: [], hist: [] };
  const emaFast = calcEMASeries_GAS(closes, 12);
  const emaSlow = calcEMASeries_GAS(closes, 26);
  const dif = closes.map((_, i) => emaFast[i] - emaSlow[i]);
  const dea = calcEMASeries_GAS(dif, 9);
  const hist = dif.map((v, i) => v - dea[i]);
  return { dif, dea, hist };
}

function checkA(ohlc, vols, tier) {
  if (ohlc.length < 60) return false;
  const cl = ohlc.map(d => d.close);
  const ma5 = calcMA_GAS(cl, 5), ma20 = calcMA_GAS(cl, 20), ma60 = calcMA_GAS(cl, 60);
  if (!ma5 || !ma20 || !ma60) return false;
  const high20 = Math.max(...ohlc.slice(-21, -1).map(d => d.max || d.high || d.close));
  const avgVol5 = vols.slice(-6, -1).reduce((a, b) => a + b, 0) / 5;
  return cl[cl.length-1] > high20 && vols[vols.length-1] > avgVol5 * getVolMult(tier) && ma5 > ma20 && ma20 > ma60;
}

function checkB(ohlc, itrust, foreign, tier) {
  if (ohlc.length < 20) return false;
  const cl = ohlc.map(d => d.close);
  const ma20 = calcMA_GAS(cl, 20);
  if (!ma20) return false;
  const last = cl[cl.length - 1];
  if (tier === 'small') return last > ma20 && itrust.slice(-5).filter(v => v > 0).length >= 2;
  return last > ma20 && itrust.slice(-5).filter(v => v > 0).length >= 3 && foreign.slice(-3).filter(v => v > 0).length >= 2;
}

function checkC(ohlc, vols, tier) {
  if (ohlc.length < 10) return false;
  const hasBlack = ohlc.slice(-5, -1).some(d => d.close < d.open);
  const today = ohlc[ohlc.length - 1];
  const avgVol10 = vols.slice(-11, -1).reduce((a, b) => a + b, 0) / 10;
  const low10 = Math.min(...ohlc.slice(-11, -1).map(d => d.min || d.low || d.close));
  const mult = tier === 'small' ? 2.5 : 2.0;
  return hasBlack && today.close > today.open && vols[vols.length-1] > avgVol10 * mult && (today.close - low10) / low10 > 0.05;
}

function checkD(ohlc, vols, tier) {
  if (ohlc.length < 20) return false;
  const box = ohlc.slice(-21, -1);
  const boxH = Math.max(...box.map(d => d.max || d.high || d.close));
  const boxL = Math.min(...box.map(d => d.min || d.low || d.close));
  const avgVol = vols.slice(-21, -1).reduce((a, b) => a + b, 0) / 20;
  const today = ohlc[ohlc.length - 1];
  return (boxH - boxL) / boxL < 0.08 && today.close > boxH && vols[vols.length-1] > avgVol * getVolMult(tier);
}

function checkE(ohlc, vols) {
  if (ohlc.length < 5) return false;
  const cl = ohlc.map(d => d.close);
  const gain5 = (cl[cl.length-1] - cl[cl.length-6]) / cl[cl.length-6];
  const today = ohlc[ohlc.length - 1];
  const volUp = vols[vols.length-1] > vols[vols.length-2] && vols[vols.length-2] > vols[vols.length-3];
  return gain5 > 0.10 && today.close > today.open && volUp && (today.close - today.open) / today.open < 0.095;
}

function checkF(ohlc, vols) {
  if (ohlc.length < 60) return false;
  const today = ohlc[ohlc.length - 1];
  const yesterday = ohlc[ohlc.length - 2];
  const avgVol20 = vols.slice(-21, -1).reduce((a, b) => a + b, 0) / 20;
  if (vols[vols.length-1] <= avgVol20 * 2.5) return false;
  if (today.close <= (yesterday.max || yesterday.high || yesterday.close)) return false;
  if (today.close <= today.open) return false;
  const high60 = Math.max(...ohlc.slice(-60).map(d => d.max || d.high || d.close));
  const low60 = Math.min(...ohlc.slice(-60).map(d => d.min || d.low || d.close));
  return (high60 - low60) / low60 > 0.15 && today.close < high60 * 0.92;
}

// G：日線趨勢篩選法。三關（均線排列 / RSI+KD / 量能+MACD）要「全部」過，
// 缺一關就不算，跟 A~F 用「命中即算一種型態」的邏輯不同，這是嚴格 AND。
function checkG(ohlc, vols) {
  if (ohlc.length < 65) return false; // 60MA + MACD/RSI 穩定所需的緩衝天數
  const cl = ohlc.map(d => d.close);
  const lastClose = cl[cl.length - 1];

  // ① 均線排列：K棒站上所有均線，且 5>10>20>60（多頭排列）
  const ma5 = calcMA_GAS(cl, 5), ma10 = calcMA_GAS(cl, 10);
  const ma20 = calcMA_GAS(cl, 20), ma60 = calcMA_GAS(cl, 60);
  if (!ma5 || !ma10 || !ma20 || !ma60) return false;
  if (!(lastClose > ma5 && ma5 > ma10 && ma10 > ma20 && ma20 > ma60)) return false;

  // ② RSI 50以上且比昨天高 ＋ KD 50以上黃金交叉
  const rsis = calcRSISeries_GAS(cl, 14);
  if (rsis.length < 2) return false;
  const rsiToday = rsis[rsis.length - 1], rsiYest = rsis[rsis.length - 2];
  if (!(rsiToday > 50 && rsiToday > rsiYest)) return false;

  // 註：原始定義是「當天剛好黃金交叉」，但這樣太嚴格——趨勢一旦成立，
  // K會持續站在D之上好幾天，只有交叉當天才算會讓每天幾乎掃不到任何股票。
  // 改成「K>D 且兩者都在50以上」，代表仍處於交叉後的轉強區間，較實用。
  // 如果你想要更嚴格、只抓「剛轉強的當天」，把下面這行換成：
  // const kdGoldenCross = kd.k[kd.k.length-2] <= kd.d[kd.d.length-2] && kToday > dToday && kToday >= 50 && dToday >= 50;
  const kd = calcKDSeries_GAS(ohlc, 9);
  if (kd.k.length < 1) return false;
  const kToday = kd.k[kd.k.length - 1], dToday = kd.d[kd.d.length - 1];
  const kdGoldenCross = kToday > dToday && kToday >= 50 && dToday >= 50;
  if (!kdGoldenCross) return false;

  // ③ 成交量站上5日均量 ＋ MACD零軸以上且柱狀體翻紅並放大
  const avgVol5 = vols.slice(-6, -1).reduce((a, b) => a + b, 0) / 5;
  if (!(vols[vols.length - 1] > avgVol5)) return false;

  const macd = calcMACDSeries_GAS(cl);
  if (macd.dif.length < 2) return false;
  const difToday = macd.dif[macd.dif.length - 1];
  const histToday = macd.hist[macd.hist.length - 1];
  const histYest = macd.hist[macd.hist.length - 2];
  if (!(difToday > 0 && histToday > 0 && histToday > histYest)) return false;

  return true;
}

// ══════════════════════════════════════════════
// 掃描進度管理
// ══════════════════════════════════════════════

function initScanProgress() {
  const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS);
  sheet.clearContents();
  const progress = {
    batchIndex: 0,
    candidates: [],
    done: false,
    scanDone: false,
    marketPass: null,
    startedAt: new Date().toISOString(),
    totalCandidates: 0,
    geminiOffset: 0,
    geminiQuotaExceeded: false,
    resultCount: 0
    // partialResults 不存在 progress，改用獨立工作表
  };
  sheet.getRange(1, 1).setValue(JSON.stringify(progress));
  // 清空暫存結果
  const resultSheet = getOrCreateSheet('ScanTemp');
  resultSheet.clearContents();
  Logger.log('掃描進度已初始化');
}

function loadScanProgress() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS);
    const val = sheet.getRange(1, 1).getValue();
    if (!val) return null;
    return JSON.parse(val);
  } catch(e) { return null; }
}

function saveScanProgress(progress) {
  // 不存 partialResults 到 progress，獨立存到 ScanTemp
  const slim = Object.assign({}, progress);
  delete slim.partialResults;
  const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS);
  const json = JSON.stringify(slim);
  if (json.length > 49000) {
    Logger.log('警告：progress JSON 仍過大：' + json.length);
  }
  sheet.getRange(1, 1).setValue(json);
}

// 暫存掃描結果（分批累積）
function saveTempResults(results) {
  const sheet = getOrCreateSheet('ScanTemp');
  if (!results || results.length === 0) return;
  // 每行存一筆 JSON
  const lastRow = sheet.getLastRow();
  const rows = results.map(r => [JSON.stringify(r)]);
  sheet.getRange(lastRow + 1, 1, rows.length, 1).setValues(rows);
}

function loadTempResults() {
  try {
    const sheet = getOrCreateSheet('ScanTemp');
    const lastRow = sheet.getLastRow();
    if (lastRow === 0) return [];
    const data = sheet.getRange(1, 1, lastRow, 1).getValues();
    return data.map(row => {
      try { return JSON.parse(row[0]); } catch(e) { return null; }
    }).filter(Boolean);
  } catch(e) { return []; }
}

function markScanDone(results, isEmpty) {
  const progress = loadScanProgress() || {};
  progress.done = true;
  progress.completedAt = new Date().toISOString();
  progress.resultCount = results.length;
  saveScanProgress(progress);
}

// ══════════════════════════════════════════════
// 掃描結果儲存與讀取
// ══════════════════════════════════════════════

// 讀寫 ScanResults 時只看表尾這麼多列：資料永遠是逐日append在最後面，
// 「今日資料」一定落在這個視窗內，不需要隨著累積的歷史天數讀越多、越讀越慢
const SCAN_READ_WINDOW = 1000;

// 欄位定義（共 28 欄，每個個股獨立一列，最前面是日期）
const SCAN_HEADERS = [
  '掃描日期',   // 0  A
  '代號',       // 1  B
  '名稱',       // 2  C
  '市值層級',   // 3  D
  '主組合',     // 4  E
  '交叉組合',   // 5  F
  '優先級',     // 6  G
  '交叉數',     // 7  H
  '進場價',     // 8  I
  '停損',       // 9  J
  '目標一',     // 10 K
  '目標二',     // 11 L
  'ATR',        // 12 M
  '投信買超天', // 13 N
  '外資買超天', // 14 O
  'RS%',        // 15 P
  '法人目標價', // 16 Q
  '法人空間%',  // 17 R
  '消息面分',   // 18 S
  '消息面摘要', // 19 T
  '券商目標',   // 20 U
  '券商評級',   // 21 V
  '券商評分',   // 22 W
  '基本面分',   // 23 X
  'EPS趨勢',    // 24 Y
  '營收趨勢',   // 25 Z
  '基本面摘要', // 26 AA
  'Gemini綜合分'// 27 AB
];

function saveScanResultsToSheet(results) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SCAN);
  const today = getTodayStr();

  // 確保標題列存在
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SCAN_HEADERS.length).setValues([SCAN_HEADERS]);
    sheet.getRange(1, 1, 1, SCAN_HEADERS.length)
      .setBackground('#2c4a6e').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // 刪除今日舊資料（避免重複）
  // 只往回看最後一段視窗：今日資料一定寫在表尾，不需要隨歷史累積逐天變慢
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const scanStart = Math.max(2, lastRow - SCAN_READ_WINDOW + 1);
    const dates = sheet.getRange(scanStart, 1, lastRow - scanStart + 1, 1).getValues();
    const todayRowNums = [];
    dates.forEach((d, i) => { if (d[0] === today) todayRowNums.push(scanStart + i); });
    // 今日資料是連續寫入的同一段，合併成連續區段後一次刪除，避免逐列呼叫 deleteRow（每次都是一次慢速 API 往返）
    if (todayRowNums.length > 0) {
      const segments = [];
      let segStart = todayRowNums[0], segEnd = todayRowNums[0];
      for (let i = 1; i < todayRowNums.length; i++) {
        if (todayRowNums[i] === segEnd + 1) { segEnd = todayRowNums[i]; }
        else { segments.push([segStart, segEnd]); segStart = segEnd = todayRowNums[i]; }
      }
      segments.push([segStart, segEnd]);
      // 從後往前刪，避免索引位移
      for (let i = segments.length - 1; i >= 0; i--) {
        sheet.deleteRows(segments[i][0], segments[i][1] - segments[i][0] + 1);
      }
    }
  }

  if (results.length === 0) {
    Logger.log('無結果可儲存');
    return;
  }

  const names = loadStockNames();
  const insertRow = sheet.getLastRow() + 1;

  const rows = results.map(r => {
    const n = r.newsAnalysis   || {};
    const b = r.brokerAnalysis || {};
    const f = r.fundAnalysis   || {};
    return [
      today,                                          // 掃描日期
      String(r.id || ''),                             // 代號
      String(r.name || names[r.id] || ''),            // 名稱
      String(r.tier || 'mid'),                        // 市值層級
      String(r.combo || ''),                          // 主組合
      String((r.combos || [r.combo]).join('+')),      // 交叉組合
      Number(r.priority || 0),                        // 優先級
      Number(r.hitCount || 1),                        // 交叉數
      Number(r.entryPrice || 0),                      // 進場價
      Number(r.stopLoss || 0),                        // 停損
      Number(r.target1 || 0),                         // 目標一
      Number(r.target2 || 0),                         // 目標二
      Number(r.atr || 0),                             // ATR
      Number(r.itrustBuy || 0),                       // 投信買超天
      Number(r.foreignBuy || 0),                      // 外資買超天
      Number(Math.round((r.rs || 0) * 1000) / 10),   // RS%
      r.analystTarget != null ? Number(r.analystTarget) : '', // 法人目標價
      r.analystUpside != null ? Number(r.analystUpside) : '', // 法人空間%
      n.score != null ? Number(n.score) : '',         // 消息面分
      String((n.summary || '').slice(0, 150)),        // 消息面摘要（限150字）
      b.targetMedian != null ? Number(b.targetMedian) : '',   // 券商目標
      String(b.consensusRating || ''),                // 券商評級
      b.score != null ? Number(b.score) : '',         // 券商評分
      f.score != null ? Number(f.score) : '',         // 基本面分
      String(f.epsGrowth || ''),                      // EPS趨勢
      String(f.revenueGrowth || ''),                  // 營收趨勢
      String((f.summary || '').slice(0, 150)),        // 基本面摘要（限150字）
      r.geminiTotal != null ? Number(r.geminiTotal) : ''      // Gemini綜合分
    ];
  });

  sheet.getRange(insertRow, 1, rows.length, SCAN_HEADERS.length).setValues(rows);

  // 色彩標記優先級（一次性批次設定，避免逐列呼叫 setBackground）
  const bgColors = rows.map(row => {
    const bg = row[6] >= 5 ? '#e8f4ff' : row[6] >= 4 ? '#f0f8f0' : '#fffef0';
    return new Array(SCAN_HEADERS.length).fill(bg);
  });
  sheet.getRange(insertRow, 1, rows.length, SCAN_HEADERS.length).setBackgrounds(bgColors);

  Logger.log(`已儲存 ${results.length} 筆掃描結果（${today}）`);
}

function loadScanResults() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.SCAN);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // 只讀表尾視窗，不要隨歷史累積把整張表讀出來找「最新日期」（見 SCAN_READ_WINDOW 說明）
    const readCount = Math.min(lastRow - 1, SCAN_READ_WINDOW);
    const startRow = lastRow - readCount + 1;
    const data = sheet.getRange(startRow, 1, readCount, SCAN_HEADERS.length).getValues();
    if (data.length === 0) return [];

    // 日期欄位可能是 Date 物件或字串，統一轉成 YYYY-MM-DD 字串
    const toDateStr = v => {
      if (!v) return '';
      if (v instanceof Date) {
        const y = v.getFullYear();
        const m = String(v.getMonth() + 1).padStart(2, '0');
        const d = String(v.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      return String(v).slice(0, 10);
    };

    // 找最新日期
    let latestDate = '';
    data.forEach(r => {
      const d = toDateStr(r[0]);
      if (d && d > latestDate) latestDate = d;
    });
    if (!latestDate) return [];

    Logger.log(`loadScanResults：最新日期=${latestDate}`);

    return data
      .filter(r => toDateStr(r[0]) === latestDate && r[1])
      .map(r => ({
        date:           toDateStr(r[0]),
        id:             String(r[1]),
        name:           String(r[2]),
        tier:           String(r[3]),
        combo:          String(r[4]),
        combos:         r[5] ? String(r[5]).split('+') : [String(r[4])],
        priority:       Number(r[6]),
        hitCount:       Number(r[7]),
        entryPrice:     Number(r[8]),
        stopLoss:       Number(r[9]),
        target1:        Number(r[10]),
        target2:        Number(r[11]),
        atr:            Number(r[12]),
        itrustBuy:      Number(r[13]),
        foreignBuy:     Number(r[14]),
        rs:             Number(r[15]) / 100,
        analystTarget:  r[16] !== '' ? Number(r[16]) : null,
        analystUpside:  r[17] !== '' ? Number(r[17]) : null,
        newsAnalysis:   r[18] !== '' ? { score: Number(r[18]), summary: String(r[19]) } : null,
        brokerAnalysis: r[20] !== '' ? { targetMedian: Number(r[20]), consensusRating: String(r[21]), score: Number(r[22]) } : null,
        fundAnalysis:   r[23] !== '' ? { score: Number(r[23]), epsGrowth: String(r[24]), revenueGrowth: String(r[25]), summary: String(r[26]) } : null,
        geminiTotal:    r[27] !== '' ? Number(r[27]) : null,
      }));
  } catch(e) {
    Logger.log('loadScanResults 失敗：' + e.message);
    return [];
  }
}

// ══════════════════════════════════════════════
// 股票名稱
// ══════════════════════════════════════════════


// ═══════════════════════════════════════════════════
// OHLC 代理（供前端呼叫，繞過 CORS）
// ═══════════════════════════════════════════════════
function fetchOHLCBatch(market, symbols, dateStr) {
  const results = {};
  if (!symbols || symbols.length === 0) return results;

  if (market === 'TW') {
    // TWSE 單股當日資料
    const d = (dateStr || getTodayStr()).replace(/-/g, '');
    symbols.forEach(id => {
      try {
        const url = `https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?stockNo=${id}&date=${d}&response=json`;
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        const json = JSON.parse(res.getContentText());
        if (json.stat === 'OK' && json.data && json.data.length > 0) {
          const row = json.data[json.data.length - 1];
          const p = s => parseFloat(String(s).replace(/,/g, '')) || 0;
          results[id] = { open: p(row[3]), high: p(row[4]), low: p(row[5]), close: p(row[6]), date: dateStr };
        } else {
          // 收盤前或休市，嘗試 TPEx
          const tpexUrl = `https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes`;
          const res2 = UrlFetchApp.fetch(tpexUrl, { muteHttpExceptions: true });
          const arr = JSON.parse(res2.getContentText());
          const row2 = arr.find(r => (r.SecuritiesCompanyCode || r.Code) === id);
          if (row2) {
            const p = s => parseFloat(String(s).replace(/,/g, '')) || 0;
            results[id] = {
              open:  p(row2.Open  || row2.OpeningPrice),
              high:  p(row2.High  || row2.HighestPrice),
              low:   p(row2.Low   || row2.LowestPrice),
              close: p(row2.Close || row2.ClosingPrice),
              date: dateStr
            };
          }
        }
      } catch(e) {
        Logger.log('fetchOHLC TW 失敗 ' + id + ':' + e.message);
      }
    });
  } else if (market === 'US') {
    // Yahoo Finance
    symbols.forEach(sym => {
      try {
        const url = `https://query1.finance.yahoo.com/v8/finance/chart/${sym}?interval=1d&range=5d`;
        const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
        const json = JSON.parse(res.getContentText());
        const result = json?.chart?.result?.[0];
        if (!result) return;
        const ts = result.timestamp || [];
        const q  = result.indicators?.quote?.[0];
        if (!q || ts.length === 0) return;
        const idx = ts.length - 1;
        results[sym] = {
          open:  Math.round((q.open[idx]  || 0) * 100) / 100,
          high:  Math.round((q.high[idx]  || 0) * 100) / 100,
          low:   Math.round((q.low[idx]   || 0) * 100) / 100,
          close: Math.round((q.close[idx] || 0) * 100) / 100,
          date:  dateStr
        };
      } catch(e) {
        Logger.log('fetchOHLC US 失敗 ' + sym + ':' + e.message);
      }
    });
  }
  return results;
}

function loadStockNames() {
  // 優先從 ScanTemp 工作表取名稱（Phase1 已帶入，不消耗 UrlFetchApp 配額）
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TEMP);
    if (sheet && sheet.getLastRow() > 0) {
      const names = {};
      const rows = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
      rows.forEach(row => {
        try {
          const obj = JSON.parse(row[0]);
          if (obj.id && obj.name) names[obj.id] = obj.name;
        } catch(e) {}
      });
      const count = Object.keys(names).length;
      if (count > 0) {
        Logger.log(`loadStockNames：從 ScanTemp 取得 ${count} 筆名稱`);
        return names;
      }
    }
  } catch(e) {
    Logger.log('loadStockNames ScanTemp 失敗：' + e.message);
  }

  // Fallback：FinMind（但可能因配額失敗）
  try {
    const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockInfo&token=${CONFIG.FINMIND_TOKEN}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const json = JSON.parse(res.getContentText());
    const names = {};
    (json.data || []).forEach(d => {
      const id = d.stock_id || d.StockID;
      const name = d.stock_name || d.StockName || d.company_name || '';
      if (id && name) names[id] = name;
    });
    Logger.log(`loadStockNames：從 FinMind 取得 ${Object.keys(names).length} 筆名稱`);
    return names;
  } catch(e) {
    Logger.log('loadStockNames FinMind 失敗：' + e.message);
    return {};
  }
}

// ══════════════════════════════════════════════
// EMAIL 通知
// ══════════════════════════════════════════════

function sendScanEmail(results, summary) {
  try {
    const to = CONFIG.NOTIFY_EMAIL || Session.getActiveUser().getEmail();
    if (!to) return false;

    const now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    const count = results.length;
    const tierCount = { large: 0, mid: 0, small: 0 };
    results.forEach(r => tierCount[r.tier || 'mid']++);

    const subject = `📊 DANVELOPER 選股完成｜${now.split(' ')[0]}｜共 ${count} 檔符合條件`;

    const tableRows = results.slice(0, 20).map(r => {
      const name = r.name || '';
      const tier = { large: '大型', mid: '中型', small: '小型' }[r.tier || 'mid'];
      const cross = r.hitCount >= 2 ? `✦ ${(r.combos||[r.combo]).join('+')}` : r.combo;
      return `<tr style="border-bottom:1px solid #e8e3d8;">
        <td style="padding:8px 10px;font-weight:700;font-family:monospace;">${r.id}</td>
        <td style="padding:8px 10px;color:#7a7168;">${name}</td>
        <td style="padding:8px 10px;">${tier}</td>
        <td style="padding:8px 10px;">${'⭐'.repeat(r.priority||1)}</td>
        <td style="padding:8px 10px;font-family:monospace;">${cross}</td>
        <td style="padding:8px 10px;font-family:monospace;color:#7c9cb5;font-weight:700;">${r.entryPrice}</td>
        <td style="padding:8px 10px;font-family:monospace;color:#b5726a;">${r.stopLoss}</td>
        <td style="padding:8px 10px;font-family:monospace;color:#7aab8a;">${r.target1}</td>
      </tr>`;
    }).join('');

    const stab = summary.stabilization;
    const bear = summary.bearishWarning;

    const bearHoldingsRows = bear && bear.warnedHoldings.length > 0
      ? bear.warnedHoldings.map(h => `<tr style="border-bottom:1px solid #f0dede;">
          <td style="padding:6px 8px;font-weight:700;font-family:monospace;">${h.id}</td>
          <td style="padding:6px 8px;color:#7a7168;">${h.name || ''}</td>
          <td style="padding:6px 8px;color:#b5726a;font-size:11px;">${h.detail}</td>
        </tr>`).join('')
      : '';

    const bearBlock = bear ? `
  <div style="padding:16px 28px 0;">
    <div style="font-size:13px;font-weight:700;color:#b5726a;margin-bottom:8px;">🐻 空頭警示：${bear.market.verdict}</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:10px;">
      ${bear.market.items.map(i => `<tr style="border-bottom:1px solid #edeae2;">
        <td style="padding:6px 8px;color:${i.warn ? '#b5726a' : '#b0a99f'};">${i.warn ? '⚠' : '─'} ${i.key}</td>
        <td style="padding:6px 8px;">${i.name}</td>
        <td style="padding:6px 8px;color:#7a7168;">${i.detail}</td>
      </tr>`).join('')}
    </table>
    ${bear.totalHoldings > 0 ? `
    <div style="font-size:12px;font-weight:700;color:#2c4a6e;margin-bottom:6px;">庫存清單檢查：${bear.warnedCount}/${bear.totalHoldings} 檔出現空頭警示</div>
    ${bearHoldingsRows ? `<table style="width:100%;border-collapse:collapse;font-size:12px;background:#fbf3f3;border-radius:6px;">${bearHoldingsRows}</table>` : '<div style="font-size:12px;color:#7aab8a;">目前庫存股沒有出現空頭警示</div>'}
    ` : '<div style="font-size:12px;color:#b0a99f;">Trades 工作表目前沒有偵測到持股</div>'}
  </div>` : '';

    const stabBlock = stab ? `
  <div style="padding:16px 28px 0;">
    <div style="font-size:13px;font-weight:700;color:#2c4a6e;margin-bottom:8px;">📍 止穩訊號 Checklist：${stab.verdict}</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;">
      ${stab.items.map(i => `<tr style="border-bottom:1px solid #edeae2;">
        <td style="padding:6px 8px;color:${i.pass ? '#7aab8a' : '#b0a99f'};">${i.pass ? '✓' : '✗'} ${i.key}</td>
        <td style="padding:6px 8px;">${i.name}</td>
        <td style="padding:6px 8px;color:#7a7168;">${i.detail}</td>
      </tr>`).join('')}
    </table>
  </div>` : '';

    const html = `<!DOCTYPE html><html><body style="font-family:sans-serif;background:#f5f2ec;margin:0;padding:20px;">
<div style="max-width:700px;margin:0 auto;background:#fff;border-radius:10px;overflow:hidden;">
  <div style="background:#2c4a6e;padding:24px 28px;">
    <div style="font-size:11px;color:#7c9cb5;letter-spacing:3px;margin-bottom:4px;">DANVELOPER SYSTEM</div>
    <div style="font-size:22px;font-weight:700;color:#fff;">📊 盤後選股完成通知</div>
    <div style="font-size:12px;color:#a8c4d8;margin-top:4px;">${now}　掃描 ${summary.totalScanned || '─'} 檔</div>
  </div>
  <div style="background:#edeae2;padding:16px 28px;display:flex;gap:20px;flex-wrap:wrap;">
    <div style="text-align:center;"><div style="font-size:26px;font-weight:700;color:#2c4a6e;">${count}</div><div style="font-size:11px;color:#7a7168;">符合個股</div></div>
    <div style="text-align:center;"><div style="font-size:20px;font-weight:700;color:#4a7ba8;">${tierCount.large}</div><div style="font-size:11px;color:#7a7168;">大型股</div></div>
    <div style="text-align:center;"><div style="font-size:20px;font-weight:700;color:#7aab8a;">${tierCount.mid}</div><div style="font-size:11px;color:#7a7168;">中型股</div></div>
    <div style="text-align:center;"><div style="font-size:20px;font-weight:700;color:#c9a96e;">${tierCount.small}</div><div style="font-size:11px;color:#7a7168;">小型股</div></div>
    <div style="text-align:center;"><div style="font-size:16px;font-weight:700;color:${summary.marketPass ? '#7aab8a' : '#b5726a'};">${summary.marketPass ? '✓ 多頭' : '✗ 謹慎'}</div><div style="font-size:11px;color:#7a7168;">大盤</div></div>
  </div>
  ${bearBlock}
  ${stabBlock}
  <div style="padding:20px 28px;">
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <thead><tr style="background:#edeae2;">
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">代號</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">名稱</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">市值</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">優先</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">組合</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">買入價</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">停損</th>
        <th style="padding:8px 10px;text-align:left;font-size:11px;color:#7a7168;">目標一</th>
      </tr></thead>
      <tbody>${tableRows}</tbody>
    </table>
    ${count > 20 ? '<p style="color:#b0a99f;font-size:12px;margin-top:10px;">僅顯示前 20 檔，完整清單請至網頁查看</p>' : ''}
  </div>
  <div style="background:#edeae2;padding:12px 28px;font-size:11px;color:#b0a99f;text-align:center;">
    DANVELOPER 選股模擬系統 | 本信件由系統自動發送，不構成投資建議
  </div>
</div></body></html>`;

    const plain = `DANVELOPER 盤後選股完成\n${now}\n共 ${count} 檔符合\n\n` +
      results.slice(0,20).map(r => `${r.id} | ${r.combo} ${'★'.repeat(r.priority||1)} | 買入 ${r.entryPrice} | 停損 ${r.stopLoss} | 目標 ${r.target1}`).join('\n');

    MailApp.sendEmail({ to, subject, body: plain, htmlBody: html });
    logNotification_(now, count, to);
    return true;
  } catch(e) {
    Logger.log('Email 寄送失敗：' + e.message);
    return false;
  }
}

function logNotification_(time, count, email) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.NOTIF);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1,1,1,4).setValues([['時間','符合檔數','Email','狀態']]);
    }
    sheet.appendRow([time, count, email, '已寄出']);
  } catch(e) {}
}

/**
 * Gemini Google Search grounding 免費額度用完時通知，本次掃描不再呼叫
 */
function notifyGeminiQuotaExceeded(skippedCount) {
  try {
    const to = CONFIG.NOTIFY_EMAIL || Session.getActiveUser().getEmail();
    if (!to) return;
    const now = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
    const subject = `⚠️ DANVELOPER Gemini 額度已用完｜${now.split(' ')[0]}`;
    const body = `Gemini（Google Search grounding）免費額度已於今日用完（${now}）。\n\n` +
      `本次掃描剩餘 ${skippedCount} 檔個股將不再呼叫 Gemini 分析，優先級僅以技術面計算，不會產生額外費用。\n` +
      `額度會定期重置，下次掃描將自動恢復嘗試。`;
    MailApp.sendEmail({ to, subject, body });
    logNotification_(now, 0, to + '（Gemini額度用完通知）');
    Logger.log('已寄送 Gemini 額度用完通知');
  } catch(e) {
    Logger.log('寄送 Gemini 額度用完通知失敗：' + e.message);
  }
}

// ══════════════════════════════════════════════
// 日期工具
// ══════════════════════════════════════════════

function getTodayStr() {
  const d = new Date();
  const tz = 'Asia/Taipei';
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function getDaysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
}

// ══════════════════════════════════════════════
// SHEET 工具
// ══════════════════════════════════════════════

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function saveStateToSheet(stateObj) {
  const sheet = getOrCreateSheet(SHEET_NAMES.STATE);
  const now = new Date().toISOString();
  const slim = JSON.stringify({ ...stateObj, recommendations: [] }); // 不存推薦，節省空間
  const data = sheet.getDataRange().getValues();
  let found = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'full_state') { found = i + 1; break; }
  }
  if (found > 0) sheet.getRange(found, 1, 1, 3).setValues([['full_state', slim, now]]);
  else sheet.appendRow(['full_state', slim, now]);
}

function loadStateFromSheet() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.STATE);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'full_state' && data[i][1]) return JSON.parse(data[i][1]);
    }
    return null;
  } catch(e) { return null; }
}

function saveTradesFromState(trades) {
  const sheet = getOrCreateSheet(SHEET_NAMES.TRADES);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,8).setValues([['日期','代號','名稱','動作','價格','股數','金額','損益']]);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 8).clearContent();
  if (trades.length === 0) return;
  const rows = trades.map(t => [t.date, t.id, t.name||'', t.action, t.price, t.shares||t.lots||0, t.amount, t.pnl||'']);
  sheet.getRange(2, 1, rows.length, 8).setValues(rows);
}

function loadTrades() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.TRADES);
    const data = sheet.getDataRange().getValues();
    return data.slice(1).filter(r => r[0]).map(r => ({
      date: r[0], id: r[1], name: r[2], action: r[3], price: r[4], shares: r[5], amount: r[6], pnl: r[7]
    }));
  } catch(e) { return []; }
}

function saveSnapshotsFromState(snapshots) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SNAPSHOTS);
  if (sheet.getLastRow() === 0) sheet.getRange(1,1,1,2).setValues([['日期','總資產']]);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  if (snapshots.length === 0) return;
  sheet.getRange(2, 1, snapshots.length, 2).setValues(snapshots.map(s => [s.date, s.totalAsset]));
}

function loadSnapshots() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.SNAPSHOTS);
    const data = sheet.getDataRange().getValues();
    return data.slice(1).filter(r => r[0]).map(r => ({ date: r[0], totalAsset: r[1] }));
  } catch(e) { return []; }
}


// ══════════════════════════════════════════════
// 動態設定管理（從 HTML 設定頁同步）
// ══════════════════════════════════════════════

/**
 * 從 Sheets 載入設定，覆蓋 CONFIG
 * 每次掃描前呼叫，確保使用最新設定
 */
function loadConfigFromSheet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
    if (!sheet) {
      Logger.log('Config 工作表不存在，使用 Code.gs 內的 CONFIG 預設值');
      return; // 直接用 CONFIG 內的值
    }
    const data = sheet.getDataRange().getValues();
    data.forEach(row => {
      const key = row[0], val = row[1];
      if (key && val !== '' && CONFIG.hasOwnProperty(key)) {
        if (['BATCH_SIZE','SCAN_TRIGGER_HOUR','SCAN_TRIGGER_MIN','SCAN_INTERVAL_MIN'].includes(key)) {
          CONFIG[key] = parseInt(val) || CONFIG[key];
        } else {
          CONFIG[key] = String(val);
        }
      }
    });
    Logger.log('設定已從 Sheets 載入：Token長度=' + (CONFIG.FINMIND_TOKEN || '').length);
  } catch(e) {
    Logger.log('loadConfigFromSheet 失敗：' + e.message + '，使用預設值');
  }
}

/**
 * 從 HTML 設定頁接收設定值並存入 Sheets
 */
function saveConfigToSheet(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config');
  if (!sheet) {
    sheet = ss.insertSheet('Config');
    sheet.getRange(1,1,1,3).setValues([['參數名稱','值','說明']]);
    sheet.getRange(1,1,1,3).setBackground('#2c4a6e').setFontColor('#ffffff').setFontWeight('bold');
  }

  const fieldMap = {
    FINMIND_TOKEN:      'FinMind API Token',
    GEMINI_KEY:         'Gemini API Key',
    NOTIFY_EMAIL:       '通知 Email',
    BATCH_SIZE:         '每批掃描檔數',
    SCAN_TRIGGER_HOUR:  '每天掃描時間（小時）',
    SCAN_TRIGGER_MIN:   '每天掃描時間（分鐘）',
    SCAN_INTERVAL_MIN:  '批次間隔分鐘',
  };

  // 先清除再寫入
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 3).clearContent();

  const rows = Object.entries(fieldMap).map(([key, desc]) => [
    key,
    config[key] !== undefined ? config[key] : CONFIG[key],
    desc
  ]);

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // 同時更新記憶體中的 CONFIG
  rows.forEach(([key, val]) => {
    if (CONFIG.hasOwnProperty(key) && val !== '') {
      if (['BATCH_SIZE','SCAN_TRIGGER_HOUR','SCAN_TRIGGER_MIN','SCAN_INTERVAL_MIN'].includes(key)) {
        CONFIG[key] = parseInt(val) || CONFIG[key];
      } else {
        CONFIG[key] = String(val);
      }
    }
  });

  Logger.log('設定已儲存至 Sheets');
}

/**
 * 重設每日掃描觸發器（設定時間改變時呼叫）
 */
function resetDailyTrigger() {
  // 刪除舊的每日觸發器
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyScanStart') ScriptApp.deleteTrigger(t);
  });
  // 用新時間建立
  ScriptApp.newTrigger('dailyScanStart')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.SCAN_TRIGGER_HOUR)
    .nearMinute(CONFIG.SCAN_TRIGGER_MIN)
    .inTimezone('Asia/Taipei')
    .create();
  Logger.log(`觸發器已重設為 ${CONFIG.SCAN_TRIGGER_HOUR}:${CONFIG.SCAN_TRIGGER_MIN}`);
}

// ══════════════════════════════════════════════
// 一次性設定（只需執行一次）
// ══════════════════════════════════════════════

function setupAll() {
  // 1. 建立所有工作表
  Object.values(SHEET_NAMES).forEach(name => getOrCreateSheet(name));

  // 2. 刪除舊的每日觸發器
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyScanStart') ScriptApp.deleteTrigger(t);
  });

  // 3. 建立每日 14:30 觸發器
  ScriptApp.newTrigger('dailyScanStart')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.SCAN_TRIGGER_HOUR)
    .nearMinute(CONFIG.SCAN_TRIGGER_MIN)
    .inTimezone('Asia/Taipei')
    .create();

  Logger.log('✅ 設定完成！');
  Logger.log('📋 工作表已建立：' + Object.values(SHEET_NAMES).join(', '));
  Logger.log('⏰ 每日 ' + CONFIG.SCAN_TRIGGER_HOUR + ':' + CONFIG.SCAN_TRIGGER_MIN + ' 自動掃描');
  Logger.log('⚠️  請記得在 CONFIG 填入 FINMIND_TOKEN 和 NOTIFY_EMAIL');
}

/**
 * 手動測試用：立刻執行一次完整掃描（不等定時觸發）
 */
function testScanNow() {
  Logger.log('=== 手動測試掃描 ===');
  dailyScanStart();
}

// 手動測試止穩訊號 Checklist（不用跑整個掃描，Apps Script編輯器直接執行這個函式看Logger輸出）
function testStabilizationSignals() {
  Logger.log('=== 手動測試止穩訊號 ===');
  const result = checkStabilizationSignals();
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log(`判定：${result.verdict}`);
  return result;
}

// 手動測試空頭警示（大盤空頭確認 + 庫存清單逐檔檢查）
function testBearishWarning() {
  Logger.log('=== 手動測試空頭警示 ===');
  const result = checkPortfolioBearishWarning();
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log(`大盤：${result.market.verdict}`);
  Logger.log(`持股警示：${result.warnedCount}/${result.totalHoldings} 檔`);
  return result;
}
