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
      // 載入最新設定
      loadConfigFromSheet();
      // 重置進度
      initScanProgress();
      // 設定 1 分鐘後觸發第一批（讓 doPost 先結束，避免逾時）
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'runScanBatch') ScriptApp.deleteTrigger(t);
      });
      ScriptApp.newTrigger('runScanBatch').timeBased().after(60 * 1000).create();
      Logger.log('startScan：進度已初始化，1 分鐘後開始第一批');
      return jsonResponse({ status: 'ok', msg: '掃描已排程，1 分鐘後開始執行，完成後寄 Email 通知' });
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
    progress.totalCandidates = phase1Results.length;
    
    saveScanProgress(progress);
    progress.marketPass = true; // 預設多頭，可在 runPhase1 中更新
    Logger.log(`第一階段完成：${phase1Results.length} 檔通過初篩`);
  }

  // 重新讀取（因為第一批可能更新了 candidates）
  const freshProgress = loadScanProgress();
  const freshCandidates = freshProgress.candidates || [];
  const freshStart = batchIndex === 0 ? 0 : batchIndex * CONFIG.BATCH_SIZE;
  const freshEnd = Math.min(freshStart + CONFIG.BATCH_SIZE, freshCandidates.length);
  const batchIds = freshCandidates.slice(freshStart, freshEnd);

  // 執行第二階段（完整技術分析）
  const batchResults = runPhase2Batch(batchIds);

  // 合併已有結果
  const existing = freshProgress.partialResults || [];
  const merged = [...existing, ...batchResults];

  const nextBatch = batchIndex + 1;
  const hasMore = freshEnd < freshCandidates.length;

  if (hasMore) {
    // 還有下一批：更新進度，設定接力 Trigger
    freshProgress.batchIndex = nextBatch;
    freshProgress.partialResults = merged;
    saveScanProgress(freshProgress);
    scheduleNextBatch(nextBatch);
    Logger.log(`第 ${batchIndex + 1} 批完成，已找到 ${merged.length} 檔，排程第 ${nextBatch + 1} 批`);
  } else {
    // 全部完成
    Logger.log(`所有批次完成，共找到 ${merged.length} 檔符合條件`);

    // 排序
    merged.sort((a, b) => (b.priority - a.priority) || (b.rs - a.rs));

    // Gemini AI 深度分析（只對技術面通過的個股）
    if (CONFIG.GEMINI_KEY && merged.length > 0) {
      Logger.log(`開始 Gemini 分析 ${merged.length} 檔...`);
      runGeminiAnalysis(merged);
      // 重新排序（加入 Gemini 加成後優先級可能改變）
      merged.sort((a, b) =>
        (b.priority - a.priority) ||
        ((b.geminiTotal || 0) - (a.geminiTotal || 0)) ||
        (b.rs - a.rs)
      );
      Logger.log('Gemini 分析完成');
    }

    // 儲存最終結果
    saveScanResultsToSheet(merged);
    markScanDone(merged, false);

    // 寄 Email
    sendScanEmail(merged, {
      marketPass: freshProgress.marketPass,
      totalScanned: freshCandidates.length,
      completedAt: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
    });

    Logger.log('=== 掃描全部完成，Email 已寄出 ===');
  }
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

function runPhase1() {
  if (!CONFIG.FINMIND_TOKEN) {
    Logger.log('❌ 錯誤：FINMIND_TOKEN 未設定，請在 Code.gs 的 CONFIG 填入 Token');
    return [];
  }

  const today = getTodayStr();
  const yesterday = getDaysAgo(3);
  Logger.log(`Phase1 開始，日期範圍：${yesterday} ~ ${today}，Token 長度：${CONFIG.FINMIND_TOKEN.length}`);

  // 一次抓全市場當日資料
  const url = `https://api.finmindtrade.com/api/v4/data?dataset=TaiwanStockPrice&start_date=${yesterday}&token=${CONFIG.FINMIND_TOKEN}`;
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const json = JSON.parse(res.getContentText());

  if (json.status !== 200) {
    Logger.log('FinMind API 錯誤：' + json.msg);
    return [];
  }

  const allData = json.data || [];

  // 只保留最新日期的資料
  const dateMap = {};
  allData.forEach(d => {
    if (!dateMap[d.stock_id] || d.date > dateMap[d.stock_id].date) {
      dateMap[d.stock_id] = d;
    }
  });

  const passed = [];
  Object.values(dateMap).forEach(d => {
    const id = d.stock_id;
    if (!id) return;

    // 排除非個股代號
    if (/[A-Za-z]/.test(id)) return;
    if (!/^\d{4,5}$/.test(id)) return;

    const close = d.close || 0;
    const open = d.open || close;
    const vol = d.Trading_Volume || 0;
    const amount = close * vol;

    if (amount < 5000000) return;  // 成交金額 < 500萬
    if (close < 5) return;          // 股價 < 5元
    if (vol === 0) return;

    // 收紅初篩
    if (close < open) return;

    // 市值分層
    let tier = 'mid';
    if (amount > 5e8) tier = 'large';
    else if (amount < 5e7) tier = 'small';

    passed.push({ id, close, open, vol, amount, tier,
      high: d.max || close, low: d.min || close });
  });

  Logger.log(`第一階段：${Object.keys(dateMap).length} 檔 → ${passed.length} 檔通過`);
  return passed;
}


// ══════════════════════════════════════════════
// GEMINI AI 分析（消息面 / 券商目標價 / 基本面）
// ══════════════════════════════════════════════

/**
 * 對所有符合條件的個股執行 Gemini 深度分析
 * 結果直接寫入 result 物件
 */
function runGeminiAnalysis(results) {
  const names = loadStockNames();

  results.forEach((r, i) => {
    try {
      Logger.log(`Gemini 分析 ${i + 1}/${results.length}: ${r.id}`);
      const name = names[r.id] || r.id;

      // 三項分析
      const news = geminiAnalyzeNews(r.id, name);
      const broker = geminiAnalyzeBroker(r.id, name, r.entryPrice);
      const fund = geminiAnalyzeFund(r.id, name, r.entryPrice);

      r.newsAnalysis = news;
      r.brokerAnalysis = broker;
      r.fundAnalysis = fund;

      // 綜合評分
      const total = (news.score || 0) + (broker.score || 0) + (fund.score || 0);
      r.geminiTotal = total;
      r.geminiBonus = total >= 4 ? 2 : total >= 2 ? 1 : 0;

      // 優先級加成
      if (r.geminiBonus > 0) {
        r.priority = Math.min(5, r.priority + r.geminiBonus);
      }

      // 每檔之間暫停 2 秒，避免 Gemini rate limit
      Utilities.sleep(2000);
    } catch(e) {
      Logger.log(`${r.id} Gemini 分析失敗：${e.message}`);
      r.newsAnalysis = { score: 0, summary: '分析失敗', keyNews: [], riskAlert: '─' };
      r.brokerAnalysis = { score: 0, targetMedian: null, consensusRating: '資料不足', recentChange: '資料不足', brokerList: [], summary: '分析失敗' };
      r.fundAnalysis = { score: 0, epsGrowth: '資料不足', revenueGrowth: '資料不足', peRatio: null, peAssessment: '資料不足', marginTrend: '資料不足', financialHealth: '資料不足', industryCycle: '不確定', summary: '分析失敗' };
      r.geminiTotal = 0;
      r.geminiBonus = 0;
    }
  });
}

/**
 * 呼叫 Gemini API（帶 Google Search grounding）
 */
function callGeminiGAS(prompt) {
  const models = ['gemini-2.0-flash', 'gemini-1.5-flash', 'gemini-1.5-flash-001'];

  for (const model of models) {
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${CONFIG.GEMINI_KEY}`;
      const payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 1024 },
        tools: [{ googleSearch: {} }]
      };

      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const json = JSON.parse(res.getContentText());

      if (res.getResponseCode() !== 200) {
        // grounding 不支援，fallback 不帶 search
        const payload2 = {
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 1024 }
        };
        const res2 = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          payload: JSON.stringify(payload2),
          muteHttpExceptions: true
        });
        if (res2.getResponseCode() !== 200) continue;
        const json2 = JSON.parse(res2.getContentText());
        const text2 = (json2.candidates || [])[0]?.content?.parts?.map(p => p.text || '').join('') || '';
        return extractJson(text2);
      }

      const parts = (json.candidates || [])[0]?.content?.parts || [];
      const text = parts.map(p => p.text || '').join('');
      return extractJson(text);

    } catch(e) {
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

/**
 * 消息面分析
 */
function geminiAnalyzeNews(stockId, stockName) {
  const today = getTodayStr();
  const prompt = `今天是 ${today}。請使用 Google Search 搜尋台灣上市股票「${stockId} ${stockName}」的最新消息面。

搜尋重點：
1. 近 14 天重大新聞（法說會、財報、訂單、策略合作、產品發表）
2. 近期利空風險（虧損、客戶流失、競爭加劇、主管異動）
3. 所屬產業景氣方向

score 必須是 -2、-1、0、1、2 其中一個整數：
+2 = 明確重大利多  +1 = 溫和正面  0 = 中性或資料不足  -1 = 溫和負面  -2 = 明確重大利空

只回傳 JSON，不要其他文字：
{"score":0,"summary":"2句繁體中文摘要","keyNews":["消息1","消息2"],"riskAlert":"風險說明或目前無明顯利空"}`;

  try {
    const result = callGeminiGAS(prompt);
    result.score = Math.max(-2, Math.min(2, parseInt(result.score) || 0));
    return result;
  } catch(e) {
    return { score: 0, summary: `分析失敗：${e.message.slice(0,30)}`, keyNews: [], riskAlert: '─' };
  }
}

/**
 * 券商目標價分析
 */
function geminiAnalyzeBroker(stockId, stockName, currentPrice) {
  const today = getTodayStr();
  const prompt = `今天是 ${today}。請使用 Google Search 搜尋台灣股票「${stockId} ${stockName}」的券商目標價資訊。

搜尋重點：
1. 近 6 個月各大券商目標價（富邦、國泰、元大、凱基、摩根士丹利、高盛、花旗）
2. 多數券商評級共識（強力買進/買進/中立/賣出）
3. 近期是否有券商上調或下調

目前股價：${currentPrice} 元

score 必須是 -2、-1、0、1、2 其中一個整數：
+2 = 目標價空間>20%且多數評級買進  +1 = 目標價空間10~20%  0 = 資料不足或分歧  -1 = 保守評級  -2 = 下調目標價

只回傳 JSON，不要其他文字：
{"score":0,"targetHigh":null,"targetLow":null,"targetMedian":null,"upsidePct":null,"consensusRating":"資料不足","recentChange":"資料不足","brokerList":[],"summary":"2句繁體中文摘要"}`;

  try {
    const result = callGeminiGAS(prompt);
    result.score = Math.max(-2, Math.min(2, parseInt(result.score) || 0));
    return result;
  } catch(e) {
    return { score: 0, targetMedian: null, consensusRating: '資料不足', recentChange: '資料不足', brokerList: [], summary: `分析失敗：${e.message.slice(0,30)}` };
  }
}

/**
 * 基本面分析
 */
function geminiAnalyzeFund(stockId, stockName, currentPrice) {
  const today = getTodayStr();
  const prompt = `今天是 ${today}。請使用 Google Search 搜尋台灣股票「${stockId} ${stockName}」的基本面財務資訊。

搜尋重點：
1. 近四季 EPS 數字與趨勢（成長/持平/衰退）
2. 近三個月月營收與年增率 YoY
3. 本益比 PE（現價 ${currentPrice} 除以近四季 EPS 合計）與同業比較
4. 毛利率趨勢（擴張/持平/收縮）
5. 財務健康（負債比率、現金流）
6. 所屬產業景氣循環位置

score 必須是 -2、-1、0、1、2 其中一個整數：
+2 = EPS成長+營收成長+PE合理+財務健康  +1 = 多數指標正面  0 = 好壞參半  -1 = 多數偏弱  -2 = EPS虧損或大幅衰退

只回傳 JSON，不要其他文字：
{"score":0,"epsGrowth":"成長","revenueGrowth":"成長","peRatio":null,"peAssessment":"合理","marginTrend":"持平","financialHealth":"健康","industryCycle":"擴張","summary":"2至3句繁體中文摘要","recentEps":[],"recentRevenue":[]}`;

  try {
    const result = callGeminiGAS(prompt);
    result.score = Math.max(-2, Math.min(2, parseInt(result.score) || 0));
    return result;
  } catch(e) {
    return { score: 0, epsGrowth: '資料不足', revenueGrowth: '資料不足', peRatio: null, peAssessment: '資料不足', marginTrend: '資料不足', financialHealth: '資料不足', industryCycle: '不確定', summary: `分析失敗：${e.message.slice(0,30)}`, recentEps: [], recentRevenue: [] };
  }
}

// ══════════════════════════════════════════════
// 第二階段：完整技術分析（單批）
// ══════════════════════════════════════════════

function runPhase2Batch(stockIds) {
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
      } else {
        if (checkB(ohlc, itrustArr, foreignArr, tier)) hitCombos.push({ combo: 'B', base: 5 });
        if (checkD(ohlc, volumes, tier))         hitCombos.push({ combo: 'D', base: 5 });
        if (checkA(ohlc, volumes, tier))         hitCombos.push({ combo: 'A', base: 4 });
        if (checkC(ohlc, volumes, tier))         hitCombos.push({ combo: 'C', base: 3 });
        if (checkE(ohlc, volumes))               hitCombos.push({ combo: 'E', base: 2 });
      }
      if (hitCombos.length === 0) return;

      const basePriority = Math.max(...hitCombos.map(c => c.base));
      const crossBonus = hitCombos.length >= 3 ? 2 : hitCombos.length === 2 ? 1 : 0;
      const priority = Math.min(basePriority + crossBonus, 5);
      const combos = hitCombos.map(c => c.combo);
      const combo = combos[0];
      const atrMult = { B: 2.0, D: 2.0, A: 1.5, C: 1.5, E: 1.0, F: 1.5 }[combo] || 1.5;

      const stopLoss = Math.round((lastClose - atr * atrMult) * 100) / 100;
      const target1 = Math.round((lastClose + atr * atrMult * 2) * 100) / 100;
      const target2 = Math.round((lastClose + atr * atrMult * 3) * 100) / 100;
      if (stopLoss >= lastClose) return;

      // RS（相對強弱，暫時用5日漲幅代替）
      const rs = closes.length >= 5
        ? (lastClose - closes[closes.length - 6]) / (closes[closes.length - 6] || 1) : 0;

      results.push({
        id, tier, combo, combos, priority,
        hitCount: hitCombos.length, crossBonus,
        entryPrice: lastClose, stopLoss, target1, target2,
        atr: Math.round(atr * 100) / 100,
        chaseLimit: Math.round(lastClose * 1.02 * 100) / 100,
        rs, itrustBuy: itrustArr.slice(-5).filter(v => v > 0).length,
        foreignBuy: foreignArr.slice(-3).filter(v => v > 0).length,
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

// ══════════════════════════════════════════════
// 掃描進度管理
// ══════════════════════════════════════════════

function initScanProgress() {
  const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS);
  sheet.clearContents();
  const progress = {
    batchIndex: 0,
    candidates: [],
    partialResults: [],
    done: false,
    marketPass: true,
    startedAt: new Date().toISOString(),
    totalCandidates: 0
  };
  sheet.getRange(1, 1).setValue(JSON.stringify(progress));
  Logger.log('掃描進度已初始化');
}

function loadScanProgress() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS);
    const val = sheet.getRange(1, 1).getValue();
    if (!val) return null;
    return JSON.parse(val);
  } catch(e) {
    return null;
  }
}

function saveScanProgress(progress) {
  const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS);
  sheet.getRange(1, 1).setValue(JSON.stringify(progress));
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

function saveScanResultsToSheet(results) {
  const sheet = getOrCreateSheet(SHEET_NAMES.SCAN);
  sheet.clearContents();

  const today = getTodayStr();
  const headers = [['掃描日期', '代號', '名稱', '市值層級', '主組合', '所有組合', '優先級',
    '交叉數', '進場價', '停損', '目標一', '目標二', 'ATR', '投信買超天', '外資買超天', 'RS值', '原始JSON']];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(1, 1, 1, headers[0].length)
    .setBackground('#2c4a6e').setFontColor('#ffffff').setFontWeight('bold');

  if (results.length === 0) return;

  // 讀取股票名稱
  const names = loadStockNames();

  const rows = results.map(r => [
    today,
    r.id,
    names[r.id] || '',
    r.tier || 'mid',
    r.combo,
    (r.combos || [r.combo]).join('+'),
    r.priority,
    r.hitCount || 1,
    r.entryPrice,
    r.stopLoss,
    r.target1,
    r.target2,
    r.atr,
    r.itrustBuy || 0,
    r.foreignBuy || 0,
    Math.round((r.rs || 0) * 1000) / 10,
    JSON.stringify(r)
  ]);

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // 色彩標記優先級
  rows.forEach((row, i) => {
    const rowNum = i + 2;
    const bg = row[6] >= 5 ? '#e8f4ff' : row[6] >= 4 ? '#f0f8f0' : '#fffef0';
    sheet.getRange(rowNum, 1, 1, rows[0].length).setBackground(bg);
  });

  Logger.log(`已儲存 ${results.length} 筆掃描結果`);
}

function loadScanResults() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.SCAN);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    return data
      .filter(r => r[0] && r[1])
      .map(r => {
        try { return JSON.parse(r[16]); } // 從原始 JSON 還原
        catch(e) {
          return { // fallback：從欄位重建
            id: r[1], name: r[2], tier: r[3], combo: r[4],
            combos: r[5].split('+'), priority: r[6],
            entryPrice: r[8], stopLoss: r[9], target1: r[10],
            target2: r[11], atr: r[12]
          };
        }
      });
  } catch(e) {
    return [];
  }
}

// ══════════════════════════════════════════════
// 股票名稱
// ══════════════════════════════════════════════

function loadStockNames() {
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
    return names;
  } catch(e) {
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
