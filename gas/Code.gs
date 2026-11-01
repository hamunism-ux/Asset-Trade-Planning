/** @OnlyCurrentDoc */
/**
 * ============================================================
 *  資產買賣紀錄與掛單規劃本 Asset Trade & Ladder Planner｜網頁 App 版（Google Apps Script Web App）
 *  版本 2.0（2026-10）
 *
 *  Google 試算表只當資料庫，所有操作都在網頁介面上做。
 *  檔案：Code.gs（本檔）、Seed.gs（原 Excel 資料）、App.html（網頁介面）
 *
 *  安裝：
 *   1. 在 Apps Script 編輯器上方選函式「setup」→ 執行（第一次會要求授權）
 *   2. 部署 → 新增部署作業 → 類型「網頁應用程式」→ 執行身分「我」→ 存取權「只有我自己」→ 部署
 *   3. 打開部署後給的網址（可加入書籤，手機也能開）
 * ============================================================
 */

const APP_VER = '3.7';
const SH = { ORD: '掛單簿', LOG: '成交紀錄', REAL: '真實交易紀錄', DONE: '已結束項目', HOLD: '持倉', GRP: '組參數', ASSET: '資產', SET: '設定' };
const TABLES = ['ORD', 'LOG', 'REAL', 'DONE', 'HOLD', 'GRP', 'ASSET', 'SET'];
const OLD_SHEETS = ['儀表板', '階梯檢視', '說明與筆記', '_圖表資料'];   // v1 留下的工作表

const ST = { PLAN: '規劃', PLACED: '已掛單', FILLED: '已成交', PAUSED: '暫停' };
const STATUS_LIST = [ST.PLAN, ST.PLACED, ST.FILLED, ST.PAUSED];
const SIDE = { SELL: '賣', BUY: '買' };

const COLS = {
  ORD: ['ID', '帳戶', '資產', '方向', '層級', '設定價格', '基本數量', '手動調整', '最終數量', '狀態', '成交紀錄ID', '備註', '更新時間'],
  LOG: ['ID', '成交時間', '帳戶', '資產', '方向', '成交價', '成交數量', '成交金額', '手續費', '交易所', '對應掛單ID', '備註'],
  REAL: ['ID', '時間', '帳戶', '資產', '方向', '價格', '數量', '金額', '手續費', '交易所', '來源', '對應掛單成交ID', '已確認', '備註', '幣別', '類型'],
  DONE: ['ID', '帳戶', '項目', '已實現金額', '幣別', '年份', '備註'],
  HOLD: ['帳戶', '資產', '基準持倉量', '基準時間', '備註', '模式'],
  GRP: ['帳戶', '資產', '方向', '起始價', '價差模式', '價差', '層數', '曲線', '起值', '參數', '目標類型', '目標值', '更新時間', '備註'],
  ASSET: ['資產', 'GOOGLEFINANCE代碼', 'CoinGecko ID', '手動價格', '數量小數位', '價格小數位', 'GOOGLEFINANCE現價(公式)', '備註', '圖示網址', '計價幣別', '數量單位'],
  SET: ['項目', '值', '說明']
};
const KEYS = {
  ORD: ['id', 'acct', 'asset', 'side', 'lvl', 'price', 'base', 'adj', 'qty', 'status', 'logId', 'note', 'updated'],
  LOG: ['id', 'time', 'acct', 'asset', 'side', 'price', 'qty', 'amt', 'fee', 'exch', 'oid', 'note'],
  REAL: ['id', 'time', 'acct', 'asset', 'side', 'price', 'qty', 'amt', 'fee', 'exch', 'source', 'logId', 'confirmed', 'note', 'ccy', 'kind'],
  DONE: ['id', 'acct', 'name', 'amt', 'ccy', 'year', 'note'],
  HOLD: ['acct', 'asset', 'qty', 'time', 'note', 'mode'],
  GRP: ['acct', 'asset', 'side', 'start', 'stepMode', 'step', 'levels', 'curve', 'startQty', 'param', 'targetType', 'targetVal', 'updated', 'note'],
  ASSET: ['name', 'gf', 'cg', 'manual', 'qd', 'pd', 'gfPrice', 'note', 'logo', 'ccy', 'unit'],
  SET: ['key', 'value', 'desc']
};
const SETTING_DEFS = [
  ['accounts', '帳戶清單', '我,MOM,DAD', '用逗號分隔'],
  ['alertPct', '提醒門檻', 0.15, '規劃中的單距離現價在此範圍內 → 提示可以掛單'],
  ['feeRate', '預設手續費率', 0.001, '登記成交時：手續費 = 成交金額 × 費率'],
  ['defExch', '預設交易所', '', '登記成交時預設帶入']
];
const AI_SHEET = 'AI診斷';
const AI_COLS = ['ID', '時間', '帳戶', '資產', '模型', '內容', '來源'];
const AI_MODELS = [
  ['claude-sonnet-5-5', 'Claude Sonnet 5.5（速度與品質平衡，建議）'],
  ['claude-opus-5-5', 'Claude Opus 5.5（分析較深入，較慢較貴）'],
  ['claude-haiku-5-5', 'Claude Haiku 5.5（最快最便宜）']
];
const NOTE_PREFIX = '筆記:';   // 每個帳戶一列：「筆記:我」「筆記:MOM」…
const OLD_IMPORT_NOTE = '原表匯入（成交日期／實際成交價待補）';
const IMPORT_NOTE = '原表匯入（2026 以前）';

// ───────────────────────── 入口 ─────────────────────────
function doGet(e) {
  // 開了手機離線版（有同步金鑰）之後，「任何人」部署的網址也會走到這裡，所以網頁介面要帶金鑰才顯示
  const key = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  if (key && !(e && e.parameter && e.parameter.k && safeEq_(String(e.parameter.k), key))) {
    return HtmlService.createHtmlOutput('<div style="font:15px/1.6 sans-serif;padding:40px;max-width:520px;margin:auto">' +
      '<h3>資產買賣紀錄與掛單規劃本</h3><p>這個網址只提供手機離線版同步用。</p><p>要在這裡打開網頁介面，請在網址後面加上 <code>?k=同步金鑰</code>，或直接使用手機離線版（GitHub Pages 的網址）。</p></div>')
      .setTitle('資產買賣紀錄與掛單規劃本');
  }
  return uiHtml_(e)
    .setTitle('資產買賣紀錄與掛單規劃本 Asset Trade & Ladder Planner')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

// ───────────────────────── 從 GitHub 取得介面／更新程式 ─────────────────────────
// 介面（App.html）預設直接從 GitHub 讀最新版，前端改了不用重新貼檔、不用重新部署。
// 伺服器程式（Code.gs）用 updateFromGitHub 一鍵更新並重新部署（需要一次性設定，見 README）。
const GH_RAW = 'https://raw.githubusercontent.com/hamunism-ux/Asset-Trade-Planning/main/gas/';
function ghFetch_(name) {
  const res = UrlFetchApp.fetch(GH_RAW + name + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('GitHub 回應 ' + res.getResponseCode() + '（' + name + '）');
  return res.getContentText();
}
function uiSource_() { return PropertiesService.getScriptProperties().getProperty('UI_SOURCE') === 'local' ? 'local' : 'github'; }
/** 介面 HTML：GitHub 最新版（快取 10 分鐘）→ 失敗就用專案裡的 App.html */
function uiHtml_(e) {
  const forceLocal = e && e.parameter && e.parameter.local;
  if (uiSource_() !== 'local' && !forceLocal) {
    try {
      const cache = CacheService.getScriptCache();
      const meta = cache.get('ui:meta');
      let html = '';
      if (meta && !(e && e.parameter && e.parameter.fresh)) {
        const m = JSON.parse(meta); const parts = [];
        for (let i = 0; i < m.n; i++) { const c = cache.get('ui:' + i); if (c === null) { parts.length = 0; break; } parts.push(c); }
        html = parts.join('');
      }
      if (!html) {
        html = ghFetch_('App.html');
        if (html.indexOf('APP_UI_VER') < 0 || html.indexOf('</html>') < 0) throw new Error('GitHub 上的 App.html 不完整');
        const CH = 90000, n = Math.ceil(html.length / CH), kv = {};
        for (let i = 0; i < n; i++) kv['ui:' + i] = html.slice(i * CH, (i + 1) * CH);
        kv['ui:meta'] = JSON.stringify({ n: n, at: Date.now() });
        cache.putAll(kv, 600);
      }
      return HtmlService.createHtmlOutput(html);
    } catch (err) { console.warn('介面改用專案內的 App.html：' + err.message); }
  }
  return HtmlService.createHtmlOutputFromFile('App');
}
/** 選單：切換介面來源（GitHub 最新版 ↔ 專案裡的 App.html） */
function uiSourceToggle() {
  const p = PropertiesService.getScriptProperties();
  const next = uiSource_() === 'local' ? 'github' : 'local';
  p.setProperty('UI_SOURCE', next);
  CacheService.getScriptCache().remove('ui:meta');
  const msg = next === 'github' ? '介面改為：GitHub 最新版（前端更新不用再貼檔）' : '介面改為：專案裡的 App.html（要自己貼檔更新）';
  try { SpreadsheetApp.getUi().alert(msg); } catch (x) { console.log(msg); }
}
/** 選單：一鍵把 Code.gs／App.html 更新成 GitHub 最新版，並把所有網頁部署換成新版本 */
function updateFromGitHub() {
  const say = m => { console.log(m); try { SpreadsheetApp.getUi().alert('更新程式', m, SpreadsheetApp.getUi().ButtonSet.OK); } catch (x) { /* 編輯器執行 */ } return m; };
  let code, html;
  try { code = ghFetch_('Code.gs'); html = ghFetch_('App.html'); } catch (e) { return say('抓不到 GitHub 上的檔案：' + e.message); }
  if (code.indexOf('function doGet') < 0 || code.indexOf('const APP_VER') < 0 || html.indexOf('APP_UI_VER') < 0 || html.indexOf('</html>') < 0) return say('GitHub 上的檔案不完整，沒有更新。');
  const newVer = (/APP_VER = '([^']+)'/.exec(code) || [0, '?'])[1];
  if (newVer === APP_VER && (/APP_UI_VER = '([^']+)'/.exec(html) || [0, '?'])[1] === newVer) { /* 仍然重新部署，確保部署是最新版 */ }
  const id = ScriptApp.getScriptId();
  const api = (path, method, body) => {
    const res = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + id + path, {
      method: method || 'get', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, payload: body ? JSON.stringify(body) : undefined
    });
    const c = res.getResponseCode(), t = res.getContentText();
    if (c === 403 && /not been used|disabled|PERMISSION_DENIED/i.test(t)) throw new Error('還沒開啟 Google Apps Script API：請到 https://script.google.com/home/usersettings 把「Google Apps Script API」打開，再執行一次。');
    if (c === 401 || (c === 403 && /insufficient/i.test(t))) throw new Error('權限不足：請確認 appsscript.json 的 oauthScopes 有 script.projects 與 script.deployments，然後重新授權。');
    if (c < 200 || c >= 300) throw new Error('Apps Script API ' + c + '：' + t.slice(0, 200));
    return t ? JSON.parse(t) : {};
  };
  try {
    const content = api('/content');
    const files = content.files || [];
    const put = (name, type, source) => { const f = files.find(x => x.name === name && x.type === type); if (f) f.source = source; else files.push({ name: name, type: type, source: source }); };
    put('Code', 'SERVER_JS', code);
    put('App', 'HTML', html);
    // 資訊清單：保留原有設定，只確保權限與網頁設定
    const mf = files.find(x => x.type === 'JSON' && x.name === 'appsscript');
    const m = mf ? JSON.parse(mf.source || '{}') : {};
    m.runtimeVersion = 'V8';
    const need = ['https://www.googleapis.com/auth/spreadsheets.currentonly', 'https://www.googleapis.com/auth/script.external_request',
      'https://www.googleapis.com/auth/script.container.ui', 'https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/script.deployments'];
    m.oauthScopes = (m.oauthScopes || []).filter(x => need.indexOf(x) < 0).concat(need);
    m.webapp = Object.assign({ executeAs: 'USER_DEPLOYING', access: 'ANYONE' }, m.webapp || {});
    put('appsscript', 'JSON', JSON.stringify(m, null, 2));
    api('/content', 'put', { files: files });
    const ver = api('/versions', 'post', { description: 'v' + newVer + ' 自動更新 ' + Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm') });
    const deps = (api('/deployments').deployments || []).filter(d => d.deploymentConfig && d.deploymentConfig.versionNumber && (d.entryPoints || []).some(p => p.entryPointType === 'WEB_APP'));
    deps.forEach(d => api('/deployments/' + d.deploymentId, 'put', { deploymentConfig: { scriptId: id, versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: d.deploymentConfig.description || '' } }));
    CacheService.getScriptCache().remove('ui:meta');
    return say('已更新到 v' + newVer + '（版本 ' + ver.versionNumber + '），' + deps.length + ' 個網頁部署已換成新版本。重新整理網頁即可。');
  } catch (e) { return say('更新失敗：' + e.message); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧭 掛單工具')
    .addItem('開啟網頁介面', 'showWebAppLink')
    .addItem('📱 手機離線版：同步金鑰與設定', 'syncSetup')
    .addSeparator()
    .addItem('⬆ 更新程式到最新版（GitHub）', 'updateFromGitHub')
    .addItem('介面來源：GitHub ↔ 專案內檔案', 'uiSourceToggle')
    .addSeparator()
    .addItem('初始化（建立資料表＋匯入原 Excel）', 'setup')
    .addToUi();
}

function showWebAppLink() {
  let url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { /* 尚未部署 */ }
  const html = url
    ? `<div style="font:14px system-ui;padding:8px"><p>網頁介面網址：</p><p><a href="${url}" target="_blank">${url}</a></p><p style="color:#666">建議加入書籤；手機登入同一個 Google 帳號也能開。</p></div>`
    : `<div style="font:14px system-ui;padding:8px"><p>還沒部署成網頁應用程式。</p><p>到 Apps Script 編輯器：部署 → 新增部署作業 → 類型選「網頁應用程式」→ 執行身分「我」→ 存取權「只有我自己」→ 部署。</p></div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(200), '資產買賣紀錄與掛單規劃本');
}

/** 在編輯器執行一次：建立資料表、匯入原 Excel 資料 */
function setup() {
  const msg = doSetup_(true, false);
  try { SpreadsheetApp.getUi().alert('資產買賣紀錄與掛單規劃本', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* 從編輯器執行時沒有 UI */ }
  console.log(msg);
  return msg;
}

// ───────────────────────── 初始化 ─────────────────────────
/** 新版加入的資料表：舊安裝第一次用到時自動建立 */
function ensureTable_(t) {
  const ss = ss_();
  if (ss.getSheetByName(SH[t])) return;
  const s = ss.insertSheet(SH[t]);
  s.getRange(1, 1, 1, COLS[t].length).setValues([COLS[t]]).setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
  s.setFrozenRows(1);
}
function isReady_() {
  const sh = ss_().getSheetByName(SH.ORD);
  if (!sh || sh.getMaxColumns() < COLS.ORD.length) return false;
  const h = sh.getRange(1, 1, 1, COLS.ORD.length).getValues()[0];
  return h.join('|') === COLS.ORD.join('|') && !!ss_().getSheetByName(SH.ASSET) && !!ss_().getSheetByName(SH.SET);
}

function doSetup_(withSeed, force) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    // 有資料就不動（除非 force）；舊版（v1）的資料會搬到新格式
    const existing = ss.getSheetByName(SH.ORD);
    const hasData = existing && existing.getLastRow() >= 2 && String(existing.getRange(2, 1).getValue()).trim();
    const v1 = force ? null : readV1_();
    if (hasData && !force && !v1) {
      return isReady_() ? '已經初始化過了（掛單簿有資料），沒有做任何變更。' : '「掛單簿」已有資料但格式認不得，為了安全沒有變更。';
    }
    // 清掉 v1 的命名範圍
    ss.getNamedRanges().forEach(n => {
      if (/^(O_|T_|H_|D_|V_|ASSET_|ACCT_|ALERT_PCT$|FEE_RATE$|DEF_EXCH$)/.test(n.getName())) n.remove();
    });
    // 先建一張暫存表，才能刪除舊表
    const tmp = ss.insertSheet('_setup_' + Date.now());
    const names = OLD_SHEETS.concat(TABLES.map(t => SH[t]));
    ss.getSheets().forEach(s => { if (names.indexOf(s.getName()) >= 0) ss.deleteSheet(s); });
    TABLES.forEach(t => {
      const s = ss.insertSheet(SH[t]);
      s.getRange(1, 1, 1, COLS[t].length).setValues([COLS[t]])
        .setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
      s.setFrozenRows(1);
      if (s.getMaxColumns() > COLS[t].length + 1) s.deleteColumns(COLS[t].length + 2, s.getMaxColumns() - COLS[t].length - 1);
    });
    const hasSeed = typeof SEED !== 'undefined';
    const seedAssets = hasSeed ? SEED.assets.map(a => ({ name: a[0], gf: a[1], cg: a[2], manual: '', qd: a[3], pd: a[4], note: a[5] })) : [];
    const notesText = hasSeed && SEED.notes ? SEED.notes.join('\n') : '';
    let msg = '資料表已建立。';
    if (v1) {
      // 舊版資料搬家
      writeSettings_(Object.assign({ notes: notesText }, v1.settings));
      writeAssets_(v1.assets.length ? v1.assets : seedAssets);
      writeOrders_(v1.orders);
      writeT_('LOG', v1.logs);
      writeT_('GRP', v1.groups);
      writeT_('HOLD', v1.holds);
      msg += `已從舊版搬過來：${v1.orders.length} 張掛單、${v1.logs.length} 筆成交紀錄、${v1.holds.length} 列持倉、${v1.groups.length} 組參數（狀態與修改都保留）。`;
    } else {
      writeSettings_({ accounts: hasSeed ? SEED.accounts : ['我'], alertPct: 0.15, feeRate: 0.001, defExch: '', notes: notesText });
      writeAssets_(seedAssets);
      if (withSeed && hasSeed) msg += seed_();
    }
    ss.deleteSheet(tmp);
    ss.getSheets().forEach(s => {
      if (/^(Sheet\d*|工作表\d*)$/.test(s.getName()) && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
    });
    // 試算表上留一個說明
    sheet_(SH.SET).getRange('D1').setValue('⚠ 請用網頁介面操作（試算表選單「🧭 掛單工具 → 開啟網頁介面」可看網址），不要直接改這幾張表的欄位順序。');
    ss.setActiveSheet(sheet_(SH.ORD));
    return msg + '\n\n下一步：部署 → 新增部署作業 → 網頁應用程式 → 執行身分「我」、存取權「只有我自己」→ 部署，打開給你的網址。';
  } finally {
    lock.releaseLock();
  }
}

/** 讀取舊版（v1：試算表內操作版）的資料；不是 v1 或沒有資料就回傳 null */
function readV1_() {
  const ss = ss_();
  const ord = ss.getSheetByName('掛單簿');
  if (!ord || ord.getMaxColumns() < 14 || ord.getLastRow() < 2) return null;
  const h = ord.getRange(1, 1, 1, 14).getValues()[0];
  if (str_(h[0]) !== 'ID' || str_(h[11]) !== '狀態' || str_(h[12]) !== '成交紀錄ID') return null;
  const rows = ord.getRange(2, 1, ord.getLastRow() - 1, 14).getValues();
  const orders = rows.filter(r => str_(r[0]) || num_(r[5])).map(r => ({
    id: str_(r[0]), acct: str_(r[1]), asset: str_(r[2]), side: str_(r[3]), lvl: num_(r[4]) === null ? '' : num_(r[4]),
    price: num_(r[5]), base: num_(r[7]) || 0, adj: num_(r[8]) || 0,
    status: STATUS_LIST.indexOf(str_(r[11])) >= 0 ? str_(r[11]) : ST.PLAN, logId: str_(r[12]), note: str_(r[13]), updated: new Date()
  })).filter(o => o.acct && o.asset && o.side && o.price > 0);
  if (!orders.length) return null;
  // 成交紀錄
  const logs = [];
  const log = ss.getSheetByName('成交紀錄');
  if (log && log.getMaxColumns() >= 13 && log.getLastRow() >= 2 && str_(log.getRange(1, 1).getValue()) === '紀錄ID') {
    log.getRange(2, 1, log.getLastRow() - 1, 13).getValues().forEach(r => {
      if (!str_(r[0])) return;
      const price = num_(r[5]), qty = num_(r[6]);
      logs.push({ id: str_(r[0]), time: r[1] instanceof Date ? r[1] : '', acct: str_(r[2]), asset: str_(r[3]), side: str_(r[4]),
        price: price, qty: qty, amt: price && qty ? round_(price * qty, 2) : '', fee: num_(r[8]) === null ? '' : num_(r[8]),
        exch: str_(r[10]), oid: str_(r[11]), note: str_(r[12]) });
    });
  }
  // 持倉
  const holds = [];
  const hold = ss.getSheetByName('持倉');
  if (hold && hold.getLastRow() >= 2 && str_(hold.getRange(1, 3).getValue()) === '基準持倉量') {
    const nc = Math.min(13, hold.getMaxColumns());
    hold.getRange(2, 1, hold.getLastRow() - 1, nc).getValues().forEach(r => {
      if (!str_(r[0]) || !str_(r[1])) return;
      holds.push({ acct: str_(r[0]), asset: str_(r[1]), qty: num_(r[2]) === null ? '' : num_(r[2]),
        time: r[3] instanceof Date ? r[3] : '', note: nc >= 13 ? str_(r[12]) : '' });
    });
  }
  // 組參數（欄位和新版相同）
  const groups = [];
  const grp = ss.getSheetByName('組參數');
  if (grp && grp.getMaxColumns() >= 14 && grp.getLastRow() >= 2) {
    grp.getRange(2, 1, grp.getLastRow() - 1, 14).getValues().forEach(r => {
      if (!str_(r[0]) || !str_(r[1]) || !str_(r[2])) return;
      const g = {};
      KEYS.GRP.forEach((k, i) => g[k] = r[i]);
      groups.push(g);
    });
  }
  // 設定（資產表 A4:K23、帳戶 M4:M13、參數 P4:P6）
  const assets = [];
  const settings = {};
  const set = ss.getSheetByName('設定');
  if (set && set.getMaxColumns() >= 16 && str_(set.getRange(3, 1).getValue()) === '資產') {
    set.getRange(4, 1, 20, 11).getValues().forEach(r => {
      if (!str_(r[0])) return;
      assets.push({ name: str_(r[0]), gf: str_(r[3]), cg: str_(r[5]), manual: num_(r[4]) === null ? '' : num_(r[4]),
        qd: intOr_(r[8], 4), pd: intOr_(r[9], 2), note: str_(r[10]) });
    });
    const accts = set.getRange(4, 13, 10, 1).getValues().map(r => str_(r[0])).filter(Boolean);
    if (accts.length) settings.accounts = accts;
    const p = set.getRange(4, 16, 3, 1).getValues();
    if (num_(p[0][0]) !== null) settings.alertPct = num_(p[0][0]);
    if (num_(p[1][0]) !== null) settings.feeRate = num_(p[1][0]);
    settings.defExch = str_(p[2][0]);
  }
  if (!settings.accounts) settings.accounts = orders.map(o => o.acct).filter((a, i, arr) => arr.indexOf(a) === i);
  if (settings.alertPct === undefined) settings.alertPct = 0.15;
  if (settings.feeRate === undefined) settings.feeRate = 0.001;
  if (settings.defExch === undefined) settings.defExch = '';
  return { orders: orders, logs: logs, holds: holds, groups: groups, assets: assets, settings: settings };
}

function seed_() {
  const assets = readAssetRows_();
  const am = {};
  assets.forEach(a => am[a.name] = a);
  const now = new Date();
  const orders = SEED.orders.map(o => {
    const base = o[5], adj = o[6] || 0, a = am[o[1]] || { qd: 6 };
    return {
      id: '', acct: o[0], asset: o[1], side: o[2], lvl: o[3], price: o[4], base: base, adj: adj,
      qty: round_(base + adj, intOr_(a.qd, 6)), status: o[7] === ST.FILLED ? SEED_FILLED_STATUS : o[7],
      logId: '', note: o[8] || '', updated: now
    };
  });
  assignIds_(orders, 'O');
  const logs = [];
  if (SEED_FILLED_STATUS === ST.FILLED) {
    orders.filter(o => o.status === ST.FILLED).forEach(o => {
      const id = 'T' + String(logs.length + 1).padStart(4, '0');
      o.logId = id;
      logs.push({ id: id, time: '', acct: o.acct, asset: o.asset, side: o.side, price: o.price, qty: o.qty,
        amt: round_(o.price * o.qty, 2), fee: '', exch: '', oid: o.id, note: IMPORT_NOTE });
    });
  }
  writeOrders_(orders);
  writeT_('LOG', logs);
  writeT_('GRP', SEED.groups.map(g => Object.assign({}, g, { updated: now })));
  const seen = {}, holds = [];
  SEED.groups.forEach(g => { const k = g.acct + '|' + g.asset; if (!seen[k]) { seen[k] = 1; holds.push({ acct: g.acct, asset: g.asset, qty: '', time: '', note: '' }); } });
  writeT_('HOLD', holds);
  return `已匯入 ${orders.length} 張掛單、${logs.length} 筆成交紀錄、${SEED.groups.length} 組參數。`;
}

// ───────────────────────── 讀取全部資料 ─────────────────────────
function api_load() {
  if (!isReady_()) return { needsSetup: true, hasSeed: typeof SEED !== 'undefined' };
  fixImportNotes_();
  ensureTable_('REAL');
  ensureTable_('DONE');
  fixRealKinds_();
  const settings = readSettings_();
  const ai = aiCfg_();
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const trades = readT_('LOG').map(t => ({
    id: str_(t.id), time: ms_(t.time), d: t.time instanceof Date ? Utilities.formatDate(t.time, tz, 'yyyy-MM-dd') : '',
    acct: str_(t.acct), asset: str_(t.asset), side: str_(t.side),
    price: num_(t.price), qty: num_(t.qty), fee: num_(t.fee), exch: str_(t.exch), oid: str_(t.oid), note: str_(t.note)
  })).filter(t => t.id);
  const real = readT_('REAL').map(t => ({
    id: str_(t.id), time: ms_(t.time), d: t.time instanceof Date ? Utilities.formatDate(t.time, tz, 'yyyy-MM-dd') : '',
    acct: str_(t.acct), asset: str_(t.asset), side: str_(t.side), price: num_(t.price), qty: num_(t.qty), fee: num_(t.fee),
    exch: str_(t.exch), source: str_(t.source) || '手動', logId: str_(t.logId), confirmed: str_(t.confirmed) === '✓', note: str_(t.note),
    ccy: str_(t.ccy) ? normCcy_(t.ccy) : '',   // 空白 = 該資產的計價幣別
    kind: realKind_(t.kind, t.note)
  })).filter(t => t.id);
  return {
    fx: fxForLoad_(trades.concat(real)),
    real: real,
    done: readT_('DONE').map(d => ({ id: str_(d.id), acct: str_(d.acct), name: str_(d.name), amt: num_(d.amt) || 0, ccy: normCcy_(d.ccy || 'TWD'), year: str_(d.year), note: str_(d.note) })).filter(d => d.id && d.name),
    ver: APP_VER, now: Date.now(),
    ai: { hasKey: !!ai.key, hint: ai.key ? 'sk-ant-…' + ai.key.slice(-4) : '', model: ai.model, search: ai.search, models: AI_MODELS },
    settings: settings,
    assets: readAssets_(),
    orders: readT_('ORD').map(o => ({
      id: str_(o.id), acct: str_(o.acct), asset: str_(o.asset), side: str_(o.side), lvl: num_(o.lvl),
      price: num_(o.price), base: num_(o.base), adj: num_(o.adj) || 0, qty: num_(o.qty),
      status: str_(o.status) || ST.PLAN, logId: str_(o.logId), note: str_(o.note)
    })).filter(o => o.id || o.price),
    trades: trades,
    holdings: readT_('HOLD').map(h => ({ acct: str_(h.acct), asset: str_(h.asset), qty: num_(h.qty), time: ms_(h.time), note: str_(h.note), mode: str_(h.mode) === '真實交易' ? '真實交易' : '手動' }))
      .filter(h => h.acct && h.asset)
  };
}

function api_setup(force) {
  doSetup_(true, !!force);
  return api_load();
}

// ───────────────────────── 掛單操作 ─────────────────────────
/** 改狀態（不含已成交；已成交請用 api_recordFills）。opts.deleteLogs：從已成交改回時一併刪除成交紀錄 */
function api_setStatus(ids, status, opts) {
  if (STATUS_LIST.indexOf(status) < 0 || status === ST.FILLED) throw new Error('狀態不正確：' + status);
  opts = opts || {};
  return mutate_(() => {
    const orders = readOrders_();
    let logs = null;
    const set = toSet_(ids), now = new Date();
    orders.forEach(o => {
      if (!set[o.id] || o.status === status) return;
      if (o.status === ST.FILLED && o.logId && opts.deleteLogs) {
        logs = logs || readT_('LOG');
        logs = logs.filter(l => str_(l.id) !== o.logId);
        o.logId = '';
      }
      o.status = status; o.updated = now;
    });
    writeOrders_(orders);
    if (logs) writeT_('LOG', logs);
  });
}

/** 登記成交 items=[{id, price, qty}]，common={time(ms), feeRate(%), exch, note} */
function api_recordFills(items, common) {
  common = common || {};
  return mutate_(() => {
    const st = readSettings_();
    const rate = common.feeRate === '' || common.feeRate === null || common.feeRate === undefined ? st.feeRate : (num_(common.feeRate) || 0) / 100;
    const exch = common.exch === undefined ? st.defExch : str_(common.exch);
    const time = num_(common.time) ? new Date(num_(common.time)) : new Date();
    const orders = readOrders_();
    const logs = readT_('LOG');
    const real = readT_('REAL');
    const om = {};
    orders.forEach(o => om[o.id] = o);
    const ids = nextIds_('T', logs.map(l => l.id), items.length);
    const rids = nextIds_('R', real.map(r => r.id), items.length);
    let k = 0, kr = 0;
    items.forEach(it => {
      const o = om[it.id];
      if (!o || o.status === ST.FILLED) return;
      const wasPlaced = o.status === ST.PLACED;
      const price = num_(it.price) || o.price, qty = num_(it.qty) || o.qty || 0;
      const amt = round_(price * qty, 2), fee = round_(amt * rate, 2);
      let id = ids[k++];
      const c = cid_(it.logId);
      if (c && !logs.some(l => str_(l.id) === c)) id = c;
      logs.push({ id: id, time: time, acct: o.acct, asset: o.asset, side: o.side, price: price, qty: qty, amt: amt,
        fee: fee, exch: exch, oid: o.id, note: str_(common.note) });
      o.status = ST.FILLED; o.logId = id; o.updated = new Date();
      // 原本已經掛在交易所的單成交了 → 自動寫一筆真實交易，標示待確認
      if (wasPlaced && !real.some(r => str_(r.logId) === id)) {
        let rid = rids[kr++];
        const rc = cid_(it.realId);
        if (rc && !real.some(r => str_(r.id) === rc)) rid = rc;
        real.push({ id: rid, time: time, acct: o.acct, asset: o.asset, side: o.side, price: price, qty: qty, amt: amt, fee: fee, exch: exch,
          source: '掛單成交', logId: id, confirmed: '', note: '由掛單 ' + o.id + ' 成交自動寫入，請確認價格與數量', ccy: '', kind: '法幣' });
      }
    });
    if (!k) throw new Error('沒有可以登記的單（可能都已經是已成交）');
    writeT_('LOG', logs);
    if (kr) { ensureTable_('REAL'); writeT_('REAL', real); }
    writeOrders_(orders);
  });
}

function api_updateOrder(id, patch) {
  return mutate_(() => {
    const orders = readOrders_();
    const o = orders.find(x => x.id === id);
    if (!o) throw new Error('找不到掛單 ' + id);
    if (patch.qty !== undefined && num_(patch.qty) !== null) { o.base = num_(patch.qty); o.adj = 0; }
    ['price', 'base', 'adj'].forEach(k => { if (patch[k] !== undefined) o[k] = patch[k] === '' || patch[k] === null ? (k === 'adj' ? 0 : '') : num_(patch[k]); });
    ['note', 'acct', 'asset', 'side'].forEach(k => { if (patch[k] !== undefined) o[k] = str_(patch[k]); });
    if (!(o.price > 0)) throw new Error('價格要大於 0');
    o.updated = new Date();
    writeOrders_(orders);
  });
}

function api_addOrder(p) {
  return mutate_(() => {
    if (!p.acct || !p.asset || !p.side || !(num_(p.price) > 0)) throw new Error('請填帳戶、資產、方向與價格');
    const orders = readOrders_();
    const id = cid_(p.id);
    if (id && orders.some(o => o.id === id)) return;   // 已經新增過（重送）
    orders.push({ id: id, acct: str_(p.acct), asset: str_(p.asset), side: str_(p.side), lvl: num_(p.lvl) || '', price: num_(p.price),
      base: num_(p.base) || 0, adj: num_(p.adj) || 0, status: STATUS_LIST.indexOf(p.status) >= 0 && p.status !== ST.FILLED ? p.status : ST.PLAN,
      logId: '', note: str_(p.note), updated: new Date() });
    writeOrders_(orders);
  });
}

function api_deleteOrders(ids) {
  return mutate_(() => {
    const set = toSet_(ids);
    writeOrders_(readOrders_().filter(o => !set[o.id]));
  });
}

function api_duplicate(ids, newIds) {
  return mutate_(() => {
    const orders = readOrders_();
    const set = toSet_(ids), have = toSet_(orders.map(o => o.id));
    const nid = {}; (ids || []).forEach((id, i) => { const c = newIds && cid_(newIds[i]); if (c && !have[c]) nid[id] = c; });
    const copies = orders.filter(o => set[o.id]).map(o => Object.assign({}, o, {
      id: nid[o.id] || '', status: ST.PLAN, logId: '', note: ('複製自 ' + o.id + ' ' + (o.note || '')).trim(), updated: new Date()
    }));
    writeOrders_(orders.concat(copies));
  });
}

/** 模擬器：把一組新價位「新增」到掛單簿（不動既有掛單）
 *  p={acct, asset, side, status, note, skipDup, levels:[{lvl, price, qty}]} */
function api_addOrders(p) {
  return mutate_(() => {
    if (!p || !p.acct || !p.asset || !p.side) throw new Error('缺少帳戶／資產／方向');
    let orders = readOrders_();
    // 取代：先刪除這組（帳戶·資產·方向）所有未成交的單，已成交保留
    if (p.mode === 'replace') orders = orders.filter(o => !(o.acct === p.acct && o.asset === p.asset && o.side === p.side && o.status !== ST.FILLED));
    const have = {};
    orders.filter(o => o.acct === p.acct && o.asset === p.asset && o.side === p.side).forEach(o => have[String(o.price)] = 1);
    const status = p.status === ST.PLACED ? ST.PLACED : ST.PLAN;
    const now = new Date();
    let n = 0;
    (p.levels || []).forEach(l => {
      const price = num_(l.price), qty = num_(l.qty);
      if (!(price > 0) || !(qty > 0)) return;
      if (p.mode !== 'replace' && p.skipDup && have[String(price)]) return;
      const lid = cid_(l.id);
      if (lid && orders.some(o => o.id === lid)) return;
      orders.push({ id: lid, acct: p.acct, asset: p.asset, side: p.side, lvl: num_(l.lvl) || '', price: price, base: qty, adj: 0,
        status: status, logId: '', note: str_(p.note), updated: now });
      n++;
    });
    if (!n) throw new Error('沒有新增任何掛單（價位可能都重複了）');
    writeOrders_(orders);
  });
}

// ───────────────────────── 成交紀錄 ─────────────────────────
/** 新增或修改成交紀錄 t={id?, time(ms|null), acct, asset, side, price, qty, fee, exch, oid, note} */
function api_saveTrade(t) {
  return mutate_(() => {
    const price = num_(t.price), qty = num_(t.qty);
    if (!t.acct || !t.asset || !t.side || !(price > 0) || !(qty > 0)) throw new Error('請填帳戶、資產、方向、成交價、數量');
    const logs = readT_('LOG');
    let row = t.id ? logs.find(l => str_(l.id) === t.id) : null;
    if (t.id && !row && !(t._new && cid_(t.id))) throw new Error('找不到成交紀錄 ' + t.id);
    if (!row) { row = { id: t._new && cid_(t.id) ? cid_(t.id) : nextIds_('T', logs.map(l => l.id), 1)[0], oid: '' }; logs.push(row); }
    Object.assign(row, {
      time: num_(t.time) ? new Date(num_(t.time)) : '', acct: str_(t.acct), asset: str_(t.asset), side: str_(t.side),
      price: price, qty: qty, amt: round_(price * qty, 2), fee: num_(t.fee) === null ? '' : num_(t.fee), exch: str_(t.exch), note: str_(t.note)
    });
    writeT_('LOG', logs);
  });
}

/** 刪除掛單成交紀錄；若有對應掛單，掛單改回「已掛單」；自動寫入且尚未確認的真實交易一併移除（已確認的保留） */
function api_deleteTrade(id) {
  return mutate_(() => {
    const logs = readT_('LOG');
    writeT_('LOG', logs.filter(l => str_(l.id) !== id));
    const real = readT_('REAL');
    let rt = false;
    const keep = real.filter(r => { if (str_(r.logId) !== id) return true; if (str_(r.confirmed) === '✓') { r.logId = ''; r.note = (str_(r.note) + '（對應的掛單成交紀錄已刪除）').trim(); rt = true; return true; } rt = true; return false; });
    if (rt) writeT_('REAL', keep);
    const orders = readOrders_();
    let touched = false;
    orders.forEach(o => { if (o.logId === id) { o.logId = ''; if (o.status === ST.FILLED) o.status = ST.PLACED; o.updated = new Date(); touched = true; } });
    if (touched) writeOrders_(orders);
  });
}

// ───────────────────────── 真實交易紀錄 ─────────────────────────
/** 新增或修改 t={id?, _new?, time(ms|null), acct, asset, side, price, qty, fee, exch, note} */
function api_saveReal(t) {
  return mutate_(() => {
    ensureTable_('REAL');
    const price = num_(t.price), qty = num_(t.qty);
    if (!t.acct || !t.asset || !t.side || !(price >= 0) || !(qty > 0)) throw new Error('請填帳戶、資產、方向、價格、數量');
    const rows = readT_('REAL');
    let row = t.id ? rows.find(r => str_(r.id) === t.id) : null;
    if (t.id && !row && !(t._new && cid_(t.id))) throw new Error('找不到交易紀錄 ' + t.id);
    if (!row) { row = { id: t._new && cid_(t.id) ? cid_(t.id) : nextIds_('R', rows.map(r => r.id), 1)[0], source: '手動', logId: '', confirmed: '✓' }; rows.push(row); }
    else if (str_(row.source) === '掛單成交') row.confirmed = '✓';   // 使用者動手修改過，視同確認
    Object.assign(row, {
      time: num_(t.time) ? new Date(num_(t.time)) : '', acct: str_(t.acct), asset: str_(t.asset), side: str_(t.side),
      price: price, qty: qty, amt: round_(price * qty, 2), fee: num_(t.fee) === null ? '' : num_(t.fee), exch: str_(t.exch), note: str_(t.note),
      ccy: str_(t.ccy) ? normCcy_(t.ccy) : '', kind: REAL_KINDS.indexOf(str_(t.kind)) >= 0 ? str_(t.kind) : guessKind_(t.note)
    });
    writeT_('REAL', rows);
  });
}
function api_deleteReal(ids) {
  return mutate_(() => { const set = toSet_(ids); writeT_('REAL', readT_('REAL').filter(r => !set[str_(r.id)])); });
}
function api_confirmReal(ids) {
  return mutate_(() => {
    const set = toSet_(ids), rows = readT_('REAL');
    let n = 0;
    rows.forEach(r => { if (set[str_(r.id)] && str_(r.confirmed) !== '✓') { r.confirmed = '✓'; n++; } });
    if (n) writeT_('REAL', rows);
  });
}
/** 批次匯入 rows=[{id?, time(ms|null), asset, side, price, qty, fee, exch, note, ccy?}]，全部記在 acct 下；ccy = 這批價格的幣別（空白 = 各資產原幣） */
function api_importReal(acct, items, ccy) {
  return mutate_(() => {
    ensureTable_('REAL');
    const st = readSettings_();
    if (st.accounts.indexOf(acct) < 0) throw new Error('找不到帳戶 ' + acct);
    const assets = {}; readAssetRows_().forEach(a => assets[a.name.toUpperCase()] = a.name);
    const rows = readT_('REAL'), have = toSet_(rows.map(r => str_(r.id)));
    const ids = nextIds_('R', rows.map(r => r.id), (items || []).length);
    let k = 0, n = 0;
    (items || []).forEach(it => {
      const price = num_(it.price), qty = num_(it.qty), asset = assets[str_(it.asset).toUpperCase()];
      if (!asset || (it.side !== '買' && it.side !== '賣') || !(price >= 0) || !(qty > 0)) return;
      let id = ids[k++];
      const c = cid_(it.id);
      if (c) { if (have[c]) return; id = c; }
      have[id] = 1;
      rows.push({ id: id, time: num_(it.time) ? new Date(num_(it.time)) : '', acct: acct, asset: asset, side: it.side, price: price, qty: qty,
        amt: round_(price * qty, 2), fee: num_(it.fee) === null ? '' : num_(it.fee), exch: str_(it.exch), source: '匯入', logId: '', confirmed: '✓', note: str_(it.note),
        ccy: str_(it.ccy || ccy) ? normCcy_(it.ccy || ccy) : '', kind: REAL_KINDS.indexOf(str_(it.kind)) >= 0 ? str_(it.kind) : guessKind_(it.note) });
      n++;
    });
    if (!n) throw new Error('沒有可以匯入的資料（請檢查資產名稱、方向、價格與數量）');
    writeT_('REAL', rows);
  });
}

// ───────────────────────── 已結束項目（只記一個已實現金額） ─────────────────────────
function api_saveDone(d) {
  return mutate_(() => {
    ensureTable_('DONE');
    const name = str_(d.name), amt = num_(d.amt);
    if (!d.acct || !name || amt === null) throw new Error('請填帳戶、項目名稱與金額');
    const rows = readT_('DONE');
    let row = d.id ? rows.find(r => str_(r.id) === d.id) : null;
    if (d.id && !row && !(d._new && cid_(d.id))) throw new Error('找不到項目 ' + d.id);
    if (!row) { row = { id: d._new && cid_(d.id) ? cid_(d.id) : nextIds_('D', rows.map(r => r.id), 1)[0] }; rows.push(row); }
    Object.assign(row, { acct: str_(d.acct), name: name, amt: amt, ccy: normCcy_(d.ccy || 'TWD'), year: str_(d.year).slice(0, 12), note: str_(d.note) });
    writeT_('DONE', rows);
  });
}
function api_deleteDone(ids) {
  return mutate_(() => { const set = toSet_(ids); writeT_('DONE', readT_('DONE').filter(r => !set[str_(r.id)])); });
}

// ───────────────────────── 持倉／設定 ─────────────────────────
function api_saveHolding(h) {
  return mutate_(() => {
    if (!h.acct || !h.asset) throw new Error('請選帳戶與資產');
    const rows = readT_('HOLD');
    let r = rows.find(x => str_(x.acct) === h.acct && str_(x.asset) === h.asset);
    if (!r) { r = { acct: h.acct, asset: h.asset }; rows.push(r); }
    if (h.qty !== undefined) r.qty = num_(h.qty) === null ? '' : num_(h.qty);
    if (h.time !== undefined) r.time = num_(h.time) ? new Date(num_(h.time)) : '';
    if (h.note !== undefined) r.note = str_(h.note);
    if (h.mode !== undefined) r.mode = h.mode === '真實交易' ? '真實交易' : '手動';
    writeT_('HOLD', rows);
  });
}

function api_deleteHolding(acct, asset) {
  return mutate_(() => writeT_('HOLD', readT_('HOLD').filter(x => !(str_(x.acct) === acct && str_(x.asset) === asset))));
}

function api_saveSettings(s) {
  return mutate_(() => {
    const next = readSettings_();
    ['alertPct', 'feeRate'].forEach(k => { if (s[k] !== undefined && num_(s[k]) !== null) next[k] = num_(s[k]); });
    if (s.defExch !== undefined) next.defExch = String(s.defExch);
    writeSettings_(next);
  });
}

function api_saveNotes(acct, text) {
  return mutate_(() => {
    const st = readSettings_();
    if (st.accounts.indexOf(acct) < 0) throw new Error('找不到帳戶 ' + acct);
    st.notes[acct] = String(text || '');
    writeSettings_(st);
  });
}

// ───────────────────────── AI 診斷 ─────────────────────────
/** API Key 存在「使用者屬性」（只有你自己的帳號讀得到），不會回傳到網頁 */
function aiCfg_() {
  const up = PropertiesService.getUserProperties();
  const model = up.getProperty('AI_MODEL');
  return { key: up.getProperty('ANTHROPIC_API_KEY') || '', model: AI_MODELS.some(m => m[0] === model) ? model : AI_MODELS[0][0],
    search: up.getProperty('AI_SEARCH') !== '0' };
}

function api_saveAi(c) {
  const up = PropertiesService.getUserProperties();
  if (c.clearKey) up.deleteProperty('ANTHROPIC_API_KEY');
  if (c.key) {
    const k = String(c.key).trim();
    if (!/^sk-ant-[A-Za-z0-9_\-]{20,}$/.test(k)) throw new Error('API Key 格式不對（應該是 sk-ant- 開頭）');
    up.setProperty('ANTHROPIC_API_KEY', k);
  }
  if (c.model && AI_MODELS.some(m => m[0] === c.model)) up.setProperty('AI_MODEL', c.model);
  if (c.search !== undefined) up.setProperty('AI_SEARCH', c.search ? '1' : '0');
  return api_load();
}

function aiSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(AI_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AI_SHEET);
    sh.getRange(1, 1, 1, AI_COLS.length).setValues([AI_COLS]).setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 這個帳戶＋資產的歷次診斷（新到舊，最多 20 筆） */
function api_aiList(acct, asset) {
  const sh = ss_().getSheetByName(AI_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, AI_COLS.length).getValues()
    .filter(r => str_(r[2]) === acct && str_(r[3]) === asset)
    .map(r => ({ id: str_(r[0]), time: ms_(r[1]), model: str_(r[4]), content: String(r[5] || ''), sources: (() => { try { return JSON.parse(r[6] || '[]'); } catch (e) { return []; } })() }))
    .sort((a, b) => (b.time || 0) - (a.time || 0)).slice(0, 20);
}

function api_aiDelete(id) {
  const sh = ss_().getSheetByName(AI_SHEET);
  if (!sh || sh.getLastRow() < 2) return true;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (str_(ids[i][0]) === id) sh.deleteRow(i + 2);
  return true;
}

// Apps Script 的 UrlFetchApp 每個請求約 60 秒就會逾時，
// 所以診斷拆成三步（各自一次呼叫，網頁依序叫）：1 搜尋消息 → 2 掛單結構 → 3 市場、情境與建議
const AI_FAST_MODEL = 'claude-haiku-5-5';
const AI_BASE_RULES = [
  '你是一位專業、務實的加密貨幣與資產配置分析師，協助使用者檢視「階梯式限價掛單」計畫（賣單越漲越賣、買單越跌越買）。',
  '用繁體中文（台灣用語）、Markdown 回答，條理清楚、數字具體，不寒暄、不重複題目。',
  '只根據提供的資料推論，不要捏造數字；資料不足就直接說缺什麼。'
].join('\n');

function aiContext_(acct, asset) {
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const fmtD = d => d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '';
  const a = readAssets_().find(x => x.name === asset);
  if (!a) throw new Error('找不到資產 ' + asset);
  const orders = readOrders_().filter(o => o.acct === acct && o.asset === asset);
  const logs = readT_('LOG').filter(l => str_(l.acct) === acct && str_(l.asset) === asset);
  const hold = readT_('HOLD').find(h => str_(h.acct) === acct && str_(h.asset) === asset);
  const st = readSettings_();
  const cur = a.price;
  const L = [];
  L.push(`今天：${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')}（${tz}）`);
  L.push(`帳戶：${acct}　資產：${asset}　現價：${cur || '未知'} ${a.ccy}（來源：${a.source || '無'}）。以下所有價格與金額都以 ${a.ccy} 計價。`);
  if (hold && str_(hold.mode) === '真實交易') {
    const rl = readT_('REAL').filter(r => str_(r.acct) === acct && str_(r.asset) === asset);
    const pos = rl.reduce((x, r) => x + (str_(r.side) === '買' ? 1 : -1) * (num_(r.qty) || 0), 0);
    L.push(`估計目前持倉：${round_(pos, a.qd)} ${asset}（依真實交易紀錄的淨部位）${cur ? '，市值約 ' + Math.round(pos * cur) + ' ' + a.ccy : ''}`);
  } else if (hold && num_(hold.qty) !== null) {
    let b = 0, s2 = 0;
    const t0 = hold.time instanceof Date ? hold.time.getTime() : null;
    if (t0) logs.forEach(l => { if (l.time instanceof Date && l.time.getTime() > t0) { if (str_(l.side) === '買') b += num_(l.qty) || 0; else s2 += num_(l.qty) || 0; } });
    const est = num_(hold.qty) + b - s2;
    L.push(`估計目前持倉：${round_(est, a.qd)} ${asset}（基準 ${num_(hold.qty)} @ ${fmtD(hold.time) || '未填日期'}，之後買入 ${round_(b, a.qd)}、賣出 ${round_(s2, a.qd)}）${cur ? '，市值約 ' + Math.round(est * cur) + ' ' + a.ccy : ''}`);
  } else L.push('持倉：使用者尚未設定');
  [['賣', '賣單（價格由低到高）'], ['買', '買單（價格由高到低）']].forEach(sd => {
    const os = orders.filter(o => o.side === sd[0]).sort((x, y) => sd[0] === '買' ? y.price - x.price : x.price - y.price);
    if (!os.length) { L.push(`\n${sd[1]}：無`); return; }
    const by = s => os.filter(o => o.status === s);
    const q = l => round_(l.reduce((x, o) => x + (o.qty || 0), 0), a.qd), v = l => Math.round(l.reduce((x, o) => x + (o.qty || 0) * o.price, 0));
    L.push(`\n${sd[1]}：共 ${os.length} 張；已成交 ${by(ST.FILLED).length} 張 ${q(by(ST.FILLED))}；已掛單 ${by(ST.PLACED).length} 張 ${q(by(ST.PLACED))}；規劃 ${by(ST.PLAN).length} 張 ${q(by(ST.PLAN))}；暫停 ${by(ST.PAUSED).length} 張。未成交（已掛＋規劃）價值約 ${v(by(ST.PLAN).concat(by(ST.PLACED)))} ${a.ccy}`);
    L.push('價格 | 數量 | 狀態 | 距現價');
    os.slice(0, 150).forEach(o => L.push(`${o.price} | ${o.qty} | ${o.status} | ${cur ? ((o.price / cur - 1) * 100).toFixed(1) + '%' : ''}`));
    if (os.length > 150) L.push(`…（其餘 ${os.length - 150} 張省略）`);
  });
  const sum2 = side => { const l = logs.filter(x => str_(x.side) === side); const q = l.reduce((x, t) => x + (num_(t.qty) || 0), 0), v = l.reduce((x, t) => x + (num_(t.qty) || 0) * (num_(t.price) || 0), 0); return `${l.length} 筆，數量 ${round_(q, a.qd)}，金額 ${Math.round(v)} ${a.ccy}${q ? '，均價 ' + round_(v / q, a.pd) : ''}`; };
  L.push(`\n成交紀錄：賣出 ${sum2('賣')}；買入 ${sum2('買')}。（沒有日期的是 2026 年以前從原表匯入的）`);
  const recent = logs.filter(l => l.time instanceof Date).sort((x, y) => y.time - x.time).slice(0, 10);
  if (recent.length) { L.push('最近成交：'); recent.forEach(l => L.push(`${fmtD(l.time)} ${str_(l.side)} ${num_(l.qty)} @ ${num_(l.price)}`)); }
  const reals = readT_('REAL').filter(r => str_(r.acct) === acct && str_(r.asset) === asset);
  if (reals.length) {
    // 紀錄若用別的幣別（例如台幣買的），依交易當天匯率換成資產的計價幣別
    const need = reals.filter(r => str_(r.ccy) && normCcy_(r.ccy) !== a.ccy && r.time instanceof Date).map(r => fmtD(r.time));
    let hist = {}; try { hist = need.length ? fxHist_(need) : {}; } catch (e) { hist = {}; }
    const latest = need.length ? fxLatest_() : null;
    const rate = (c, d) => c === 'USD' ? 1 : (hist[d] && hist[d][c]) || (latest && latest[c]) || null;
    const px = r => { const c = str_(r.ccy) ? normCcy_(r.ccy) : a.ccy; if (c === a.ccy) return num_(r.price) || 0; const d = r.time instanceof Date ? fmtD(r.time) : ''; const f = rate(c, d), t = rate(a.ccy, d); return f && t ? (num_(r.price) || 0) / f * t : 0; };
    const rb = reals.filter(r => str_(r.side) === '買'), rs = reals.filter(r => str_(r.side) === '賣');
    const q = l => l.reduce((x, t) => x + (num_(t.qty) || 0), 0), v = l => l.reduce((x, t) => x + (num_(t.qty) || 0) * px(t), 0);
    const fiat = reals.filter(r => realKind_(r.kind, r.note) === '法幣');
    const net = v(fiat.filter(r => str_(r.side) === '買')) - v(fiat.filter(r => str_(r.side) === '賣'));
    L.push(`\n真實交易紀錄（使用者實際成交的買賣，含掛單以外的交易）：買入 ${rb.length} 筆 ${round_(q(rb), a.qd)}（均價 ${q(rb) ? round_(v(rb) / q(rb), a.pd) : '-'}）；賣出 ${rs.length} 筆 ${round_(q(rs), a.qd)}（均價 ${q(rs) ? round_(v(rs) / q(rs), a.pd) : '-'}）；淨部位 ${round_(q(rb) - q(rs), a.qd)} ${asset}；法幣淨投入（只算用錢買賣，不含幣幣交換與轉帳）${net > 0 ? round_(net, 0) + ' ' + a.ccy : '已全部回本（多收回 ' + round_(-net, 0) + ' ' + a.ccy + '）'}`);
    const rr = reals.filter(r => r.time instanceof Date).sort((x, y) => y.time - x.time).slice(0, 8);
    if (rr.length) { L.push('最近真實交易：'); rr.forEach(r => L.push(`${fmtD(r.time)} ${str_(r.side)} ${num_(r.qty)} @ ${round_(px(r), a.pd)}${str_(r.note) ? '（' + str_(r.note).slice(0, 30) + '）' : ''}`)); }
  }
  const note = (st.notes[acct] || '').trim();
  L.push(`\n使用者的策略筆記：${note ? '\n' + note.slice(0, 2000) : '（無）'}`);
  return L.join('\n');
}

/**
 * 5.5 世代的模型預設會先「思考」，思考的 token 也算在 max_tokens 裡，
 * 而且會拉長回應時間（Apps Script 每次連線約 60 秒上限）。這裡依模型把思考關掉或壓低：
 * Haiku → 關閉；Sonnet → between_tools（不先思考）；Opus 無法關閉 → effort 設 low 並多留額度。
 */
function thinkingFor_(model) {
  if (/haiku/.test(model)) return { thinking: { type: 'disabled' } };
  if (/sonnet/.test(model)) return { thinking: { type: 'between_tools' } };
  if (/opus/.test(model)) return { output_config: { effort: 'low' }, extraTokens: 6000 };
  return {};
}

/** 呼叫一次 Claude（含 pause_turn 續跑）；回傳 { text, sources, usage, stop } */
function claude_(cfg, o) {
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const model = o.model || cfg.model, th = thinkingFor_(model);
  const base = { model: model, max_tokens: (o.maxTokens || 2000) + (th.extraTokens || 0), system: o.system };
  if (th.thinking) base.thinking = th.thinking;
  if (th.output_config) base.output_config = th.output_config;
  let fellBack = false;
  if (o.search) base.tools = [{ type: 'web_search_20250305', name: 'web_search', max_uses: o.search, user_location: { type: 'approximate', timezone: tz } }];
  const user = { role: 'user', content: o.prompt };
  let all = [], data = null;
  const usage = { input: 0, output: 0, searches: 0 };
  for (let round = 0; round < 4; round++) {
    const msgs = all.length ? [user, { role: 'assistant', content: all }] : [user];
    let res;
    try {
      res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': cfg.key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify(Object.assign({}, base, { messages: msgs }))
      });
    } catch (e) {
      const to = /timeout|timed out|逾時/i.test(e.message);
      throw new Error((o.label || 'AI') + (to ? ' 超過 Google 的 60 秒連線上限。' : ' 連線失敗：' + e.message + '。') +
        '再按一次通常就會過；常發生的話，到「設定 → AI 診斷」改用 Haiku 或關掉網路搜尋。');
    }
    const code = res.getResponseCode(), text = res.getContentText();
    try { data = JSON.parse(text); } catch (e) { data = null; }
    if (code === 400 && !fellBack && (base.thinking || base.output_config) && data && data.error && /thinking|effort|output_config/i.test(data.error.message || '')) {
      // 萬一這個模型不接受上面的思考設定：拿掉設定、加大額度再試一次
      console.warn('thinking config rejected, retrying without it:', data.error.message);
      delete base.thinking; delete base.output_config; base.max_tokens += 8000; fellBack = true; round--; continue;
    }
    if (code !== 200) {
      const m = data && data.error ? data.error.message : text.slice(0, 300);
      if (code === 401) throw new Error('API Key 無效或已撤銷（401）。請到設定重新填入。');
      if (code === 429 || code === 529) throw new Error('Anthropic 目前忙碌或達到速率上限（' + code + '），稍後再試。');
      if (code === 400 && /credit balance/i.test(m)) throw new Error('Anthropic 帳戶餘額不足，請到 Claude Console 的 Billing 加值。（' + m + '）');
      throw new Error('Anthropic API 錯誤 ' + code + '：' + m);
    }
    all = all.concat(data.content || []);
    if (data.usage) { usage.input += data.usage.input_tokens || 0; usage.output += data.usage.output_tokens || 0; if (data.usage.server_tool_use) usage.searches += data.usage.server_tool_use.web_search_requests || 0; }
    if (data.stop_reason !== 'pause_turn') break;
  }
  const sources = (o.sources || []).slice(), idx = {};
  sources.forEach((s, i) => idx[s.url] = i + 1);
  let out = '';
  all.forEach(b => {
    if (b.type === 'web_search_tool_result' && b.content && b.content.error_code) console.warn('web_search error', b.content.error_code);
    if (b.type !== 'text') return;
    out += b.text;
    const refs = [];
    (b.citations || []).forEach(c => {
      if (!c.url) return;
      if (!idx[c.url]) { sources.push({ url: c.url, title: c.title || c.url }); idx[c.url] = sources.length; }
      if (refs.indexOf(idx[c.url]) < 0) refs.push(idx[c.url]);
    });
    if (refs.length) out += refs.map(n => `[${n}]`).join('');
  });
  return { text: out.trim(), sources: sources, usage: usage, stop: data && data.stop_reason };
}

function aiNeedKey_() {
  const cfg = aiCfg_();
  if (!cfg.key) throw new Error('請先到「設定 → AI 診斷」填入 Anthropic API Key');
  return cfg;
}

/**
 * 診斷的其中一步（網頁依序呼叫 news → p1 → p2）。
 * news：搜尋近期消息，整理成重點（用 Haiku，比較快）
 * p1：現況摘要＋掛單結構評估
 * p2：一句話結論、市場面、情境、建議；組合全文並存檔
 */
function api_aiStep(acct, asset, step, prev) {
  const cfg = aiNeedKey_();
  prev = prev || {};
  if (step === 'news') {
    if (!cfg.search) return { text: '', sources: [], usage: null };
    const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
    const r = claude_(cfg, {
      label: '第 1 步（搜尋消息）', model: AI_FAST_MODEL, maxTokens: 2500, search: 3,
      system: '你是金融研究助理。用網路搜尋找出資料，整理成精簡的繁體中文重點，每點都附日期。不要給投資建議。',
      prompt: `今天是 ${Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd')}。請搜尋「${asset}」最近 1–4 週的重要資訊，整理成 8–12 條重點（每條一行，開頭寫日期），分三類：\n` +
        `### 價格、資金流與市場情緒\n### 總體經濟（利率、通膨、美元、流動性、監管）\n### 科技與產業發展\n最後一行寫「整體氛圍：偏多／中性／偏空」並用一句話說明。總長 500 字以內。`
    });
    return { text: r.text, sources: r.sources, usage: r.usage };
  }
  const ctx = aiContext_(acct, asset);
  if (step === 'p1') {
    const r = claude_(cfg, {
      label: '第 2 步（掛單結構）', maxTokens: 3500, system: AI_BASE_RULES,
      prompt: ctx + '\n\n請只輸出以下兩段（合計約 400–700 字）：\n' +
        '## 現況摘要\n（價格位置、持倉、已實現的賣出／買入、目前在階梯的哪一段）\n' +
        '## 掛單結構評估\n（價格分布與密度、距現價、每張數量的曲線形狀、未成交占持倉比例、均價；有沒有賣超過持倉、買單資金是否足夠、有沒有太集中或斷層）'
    });
    if (!r.text) throw new Error('第 2 步沒有回傳內容（' + r.stop + (r.stop === 'max_tokens' ? '：模型思考用完了額度' : '') + '），請再試一次或改用其他模型');
    return { text: r.text + (r.stop === 'max_tokens' ? '\n\n（這段內容過長被截斷）' : ''), usage: r.usage };
  }
  if (step === 'p2') {
    const news = prev.news || { text: '', sources: [] };
    const r = claude_(cfg, {
      label: '第 3 步（市場、情境與建議）', maxTokens: 4000, system: AI_BASE_RULES, sources: news.sources,
      prompt: ctx + '\n\n=== 前一步的分析 ===\n' + (prev.p1 || '（無）') +
        '\n\n=== 近期消息摘要 ===\n' + (news.text || '（使用者關閉了網路搜尋，沒有最新消息；請在市場段落說明這點，只做一般性的提醒）') +
        '\n\n請接著輸出以下段落（合計約 600–900 字）：\n' +
        '## 一句話結論\n（綜合以上，一到兩句）\n' +
        '## 市場、總經與科技面\n（根據消息摘要說明對這個資產的影響；引用消息時保留原本的 [數字] 標記與日期）\n' +
        '## 情境推演\n（上漲、盤整、下跌各一小段：會成交哪些掛單、兌現多少、剩多少持倉或要多少資金）\n' +
        '## 具體建議\n（3–6 點，可直接執行，附價格區間或數量，標註優先順序 高／中／低）\n' +
        '最後單獨一行：「以上為 AI 依資料產生的參考意見，不構成投資建議。」'
    });
    if (!r.text) throw new Error('第 3 步沒有回傳內容（' + r.stop + (r.stop === 'max_tokens' ? '：模型思考用完了額度' : '') + '），請再試一次或改用其他模型');
    // 把「一句話結論」移到最前面
    let t2 = r.text, concl = '';
    const m = /(^|\n)##\s*一句話結論[^\n]*\n([\s\S]*?)(?=\n##\s|$)/.exec(t2);
    if (m) { concl = '## 一句話結論\n' + m[2].trim(); t2 = (t2.slice(0, m.index) + t2.slice(m.index + m[0].length)).trim(); }
    const content = [concl, prev.p1 || '', t2].filter(Boolean).join('\n\n') + (r.stop === 'max_tokens' ? '\n\n（內容過長被截斷）' : '');
    const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
    const sh = aiSheet_();
    const id = 'A' + Date.now();
    const safe = t => /^[=+\-@]/.test(t) ? "'" + t : t;   // 避免被試算表當成公式
    sh.appendRow([id, new Date(), acct, asset, cfg.model + (cfg.search ? '+search' : ''), safe(content.slice(0, 45000)), JSON.stringify(r.sources)]);
    return { id: id, time: Date.now(), model: cfg.model, content: content, sources: r.sources, usage: r.usage, tz: tz };
  }
  throw new Error('未知的步驟 ' + step);
}

// ───────────────────────── 帳戶 ─────────────────────────

function cleanAcct_(n) {
  n = str_(n).replace(/[,，|]/g, '');
  if (!n) throw new Error('帳戶名稱不能空白');
  if (n.length > 20) throw new Error('帳戶名稱最多 20 個字');
  return n;
}

function api_addAccount(name) {
  return mutate_(() => {
    const st = readSettings_(), n = cleanAcct_(name);
    if (st.accounts.indexOf(n) >= 0) throw new Error('已經有帳戶「' + n + '」');
    st.accounts.push(n);
    writeSettings_(st);
  });
}

/** 改名：掛單、成交紀錄、持倉、組參數、筆記全部跟著改 */
function api_renameAccount(oldName, newName) {
  return mutate_(() => {
    const st = readSettings_(), n = cleanAcct_(newName), i = st.accounts.indexOf(oldName);
    if (i < 0) throw new Error('找不到帳戶 ' + oldName);
    if (n === oldName) return;
    if (st.accounts.indexOf(n) >= 0) throw new Error('已經有帳戶「' + n + '」');
    st.accounts[i] = n;
    if (st.notes[oldName] !== undefined) { st.notes[n] = st.notes[oldName]; delete st.notes[oldName]; }
    if (st.pw[oldName]) { delete st.pw[oldName]; }   // 雜湊含帳戶名稱，改名後密碼失效，需重設
    const aiSh = ss_().getSheetByName(AI_SHEET);
    if (aiSh && aiSh.getLastRow() >= 2) {
      const rg = aiSh.getRange(2, 3, aiSh.getLastRow() - 1, 1), v = rg.getValues();
      if (v.some(r => str_(r[0]) === oldName)) rg.setValues(v.map(r => [str_(r[0]) === oldName ? n : r[0]]));
    }
    ['ORD', 'LOG', 'REAL', 'DONE', 'HOLD', 'GRP'].forEach(t => {
      const rows = readT_(t);
      let touched = false;
      rows.forEach(r => { if (str_(r.acct) === oldName) { r.acct = n; touched = true; } });
      if (touched) writeT_(t, rows);
    });
    writeSettings_(st);
  });
}

/** 刪除帳戶：連同它的掛單、成交紀錄、持倉、組參數與筆記 */
function api_deleteAccount(name) {
  return mutate_(() => {
    const st = readSettings_();
    if (st.accounts.indexOf(name) < 0) throw new Error('找不到帳戶 ' + name);
    if (st.accounts.length <= 1) throw new Error('至少要保留一個帳戶');
    st.accounts = st.accounts.filter(a => a !== name);
    delete st.notes[name]; delete st.pw[name];
    const aiSh = ss_().getSheetByName(AI_SHEET);
    if (aiSh && aiSh.getLastRow() >= 2) {
      const v = aiSh.getRange(2, 3, aiSh.getLastRow() - 1, 1).getValues();
      for (let i = v.length - 1; i >= 0; i--) if (str_(v[i][0]) === name) aiSh.deleteRow(i + 2);
    }
    ['ORD', 'LOG', 'REAL', 'DONE', 'HOLD', 'GRP'].forEach(t => {
      const rows = readT_(t), keep = rows.filter(r => str_(r.acct) !== name);
      if (keep.length !== rows.length) writeT_(t, keep);
    });
    writeSettings_(st);
  });
}

function api_saveAssets(list) {
  return mutate_(() => {
    const seen = {};
    const rows = (list || []).map(a => ({ name: str_(a.name).toUpperCase(), gf: str_(a.gf).replace(/^=+/, '').trim(), cg: str_(a.cg).toLowerCase(),
      manual: num_(a.manual) === null ? '' : num_(a.manual), qd: intOr_(a.qd, 4), pd: intOr_(a.pd, 2), note: str_(a.note), logo: str_(a.logo), ccy: normCcy_(a.ccy), unit: str_(a.unit).slice(0, 6) }))
      .filter(a => a.name && !seen[a.name] && (seen[a.name] = 1));
    if (!rows.length) throw new Error('至少要有一個資產');
    // 已有掛單或成交紀錄的資產不能改計價幣別（舊的價格會被誤當成新幣別）
    const old = {}; readAssetRows_().forEach(a => old[a.name] = a.ccy);
    const used = {}; readT_('ORD').concat(readT_('LOG'), readT_('REAL')).forEach(r => used[str_(r.asset)] = 1);
    rows.forEach(a => { if (old[a.name] && old[a.name] !== a.ccy && used[a.name]) throw new Error(`${a.name} 已經有掛單或成交紀錄，不能把計價幣別從 ${old[a.name]} 改成 ${a.ccy}`); });
    writeAssets_(rows);
    CacheService.getScriptCache().remove(cgCacheKey_(rows.map(a => a.cg).filter(Boolean)));
  });
}

function api_refreshPrices() {
  const ids = readAssetRows_().map(a => a.cg).filter(Boolean);
  CacheService.getScriptCache().remove(cgCacheKey_(ids));
  CacheService.getScriptCache().remove('fx:latest');
  return api_load();
}

// 真實交易的類型：法幣（用錢買賣）、幣幣（用別的幣換）、轉帳（送出／收到，沒有買賣）
const REAL_KINDS = ['法幣', '幣幣', '轉帳'];
function realKind_(kind, note) {
  kind = str_(kind);
  if (REAL_KINDS.indexOf(kind) >= 0) return kind;
  return guessKind_(note);
}
function guessKind_(note) {
  const n = str_(note);
  if (/轉出|轉入|give to|receive from|received from|transfer|送出|收到/i.test(n)) return '轉帳';
  if (/換成|換入|幣幣|exchange|swap|convert|hard fork|分叉|投入 |buy from (eth|btc|ada|usdt|usdc)/i.test(n)) return '幣幣';
  return '法幣';
}
/** 一次性：舊資料沒有「類型」欄的，依備註推斷並寫回 */
function fixRealKinds_() {
  const sh = ss_().getSheetByName(SH.REAL);
  if (!sh || sh.getLastRow() < 2) return;
  if (sh.getMaxColumns() < COLS.REAL.length) sh.insertColumnsAfter(sh.getMaxColumns(), COLS.REAL.length - sh.getMaxColumns());
  if (str_(sh.getRange(1, COLS.REAL.length).getValue()) !== '類型') sh.getRange(1, COLS.REAL.length).setValue('類型').setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
  const rows = readT_('REAL');
  if (!rows.some(r => str_(r.id) && !str_(r.kind))) return;
  rows.forEach(r => { if (!str_(r.kind)) r.kind = guessKind_(r.note); });
  writeT_('REAL', rows);
}

/** 一次性：把舊的「日期待補」說明改成「2026 以前」 */
function fixImportNotes_() {
  const logs = readT_('LOG');
  if (!logs.some(l => str_(l.note) === OLD_IMPORT_NOTE)) return;
  logs.forEach(l => { if (str_(l.note) === OLD_IMPORT_NOTE) l.note = IMPORT_NOTE; });
  writeT_('LOG', logs);
}

// ───────────────────────── 資料層 ─────────────────────────
function mutate_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try { fn(); SpreadsheetApp.flush(); }
  finally { lock.releaseLock(); }
  return BATCH_ ? null : api_load();
}
let BATCH_ = false;   // 同步批次中：每個操作不必各自回傳整份資料
/** 前端（手機離線版）自己產生的 ID；格式不對就忽略，改用伺服器編號 */
function cid_(v) { v = str_(v); return /^[A-Za-z0-9_-]{4,40}$/.test(v) ? v : ''; }

function readT_(t) {
  const sh = ss_().getSheetByName(SH[t]);
  if (!sh) return [];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const nc = Math.min(COLS[t].length, sh.getMaxColumns());
  const v = sh.getRange(2, 1, n, nc).getValues();
  return v.filter(r => r.some(c => c !== '' && c !== null)).map(r => {
    const o = {};
    KEYS[t].forEach((k, i) => o[k] = r[i]);
    return o;
  });
}

function writeT_(t, objs) {
  const sh = sheet_(SH[t]);
  const nc = COLS[t].length;
  if (sh.getMaxColumns() < nc) sh.insertColumnsAfter(sh.getMaxColumns(), nc - sh.getMaxColumns());
  const old = Math.max(0, sh.getLastRow() - 1);
  const rows = objs.map(o => KEYS[t].map(k => cell_(o[k])));
  if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
  if (rows.length) sh.getRange(2, 1, rows.length, nc).setValues(rows);
  if (old > rows.length) sh.getRange(rows.length + 2, 1, old - rows.length, nc).clearContent();
}

function readOrders_() {
  return readT_('ORD').map(o => ({
    id: str_(o.id), acct: str_(o.acct), asset: str_(o.asset), side: str_(o.side), lvl: o.lvl === '' ? '' : num_(o.lvl),
    price: num_(o.price), base: num_(o.base), adj: num_(o.adj) || 0, qty: num_(o.qty),
    status: str_(o.status) || ST.PLAN, logId: str_(o.logId), note: str_(o.note), updated: o.updated
  }));
}

/** 寫回全部掛單：排序、補 ID、重算最終數量 */
function writeOrders_(orders) {
  const st = readSettings_();
  const assets = readAssetRows_();
  const am = {};
  assets.forEach((a, i) => { am[a.name] = a; a.i = i; });
  const ai = n => { const i = st.accounts.indexOf(n); return i < 0 ? 999 : i; };
  const si = n => am[n] ? am[n].i : 999;
  orders.forEach(o => {
    const a = am[o.asset];
    const q = (num_(o.base) || 0) + (num_(o.adj) || 0);
    o.qty = a ? round_(q, intOr_(a.qd, 6)) : q;
  });
  orders.sort((a, b) => (ai(a.acct) - ai(b.acct)) || String(a.acct).localeCompare(String(b.acct)) ||
    (si(a.asset) - si(b.asset)) || String(a.asset).localeCompare(String(b.asset)) ||
    ((a.side === SIDE.SELL ? 0 : 1) - (b.side === SIDE.SELL ? 0 : 1)) ||
    (a.side === SIDE.BUY ? (b.price || 0) - (a.price || 0) : (a.price || 0) - (b.price || 0)));
  // 層級一律依價格順序：賣單由低到高、買單由高到低（每組從 1 開始）
  let prevKey = '', lv = 0;
  orders.forEach(o => { const k = o.acct + '|' + o.asset + '|' + o.side; lv = k === prevKey ? lv + 1 : 1; prevKey = k; o.lvl = lv; });
  assignIds_(orders, 'O');
  writeT_('ORD', orders);
}

function readSettings_() {
  const rows = readT_('SET');
  const m = {};
  rows.forEach(r => m[str_(r.key)] = r.value);
  const get = (label, def) => (m[label] !== undefined && m[label] !== '') ? m[label] : def;
  const d = {};
  SETTING_DEFS.forEach(x => d[x[0]] = get(x[1], x[2]));
  const accounts = String(d.accounts).split(',').map(x => x.trim()).filter(Boolean);
  const notes = {}, pw = {};
  Object.keys(m).forEach(k => { if (k.indexOf(NOTE_PREFIX) === 0) notes[k.slice(NOTE_PREFIX.length)] = String(m[k]); if (k.indexOf(PW_PREFIX) === 0 && str_(m[k])) pw[k.slice(PW_PREFIX.length)] = str_(m[k]); });
  // 舊版只有一份「筆記」→ 放到「我」（沒有「我」就放第一個帳戶）
  if (m['筆記'] !== undefined && String(m['筆記']).trim()) {
    const owner = accounts.indexOf('我') >= 0 ? '我' : accounts[0];
    if (owner && !notes[owner]) notes[owner] = String(m['筆記']);
  }
  return {
    accounts: accounts,
    alertPct: num_(d.alertPct) === null ? 0.15 : num_(d.alertPct),
    feeRate: num_(d.feeRate) === null ? 0.001 : num_(d.feeRate),
    defExch: str_(d.defExch), notes: notes, pw: pw
  };
}
const PW_PREFIX = '密碼:';
/** 密碼雜湊（和網頁端用同一個算法：SHA-256( 帳戶 + ':' + 密碼 ) 的十六進位） */
function pwHash_(acct, pw) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, acct + ':' + pw, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + ((x + 256) % 256).toString(16)).slice(-2)).join('');
}
/** 設定／更改／移除帳戶密碼。有舊密碼時要先驗證。newPw 空字串 = 移除 */
function api_setPassword(acct, oldPw, newPw) {
  return mutate_(() => {
    const st = readSettings_();
    if (st.accounts.indexOf(acct) < 0) throw new Error('找不到帳戶 ' + acct);
    if (st.pw[acct] && !safeEq_(st.pw[acct], pwHash_(acct, String(oldPw || '')))) { Utilities.sleep(800); throw new Error('目前的密碼不正確'); }
    newPw = String(newPw || '');
    if (newPw && newPw.length < 4) throw new Error('密碼至少 4 個字');
    if (newPw) st.pw[acct] = pwHash_(acct, newPw); else delete st.pw[acct];
    writeSettings_(st);
  });
}

function writeSettings_(s) {
  const accounts = Array.isArray(s.accounts) ? s.accounts : String(s.accounts || '').split(',').map(x => x.trim()).filter(Boolean);
  let notes = s.notes || {};
  if (typeof notes === 'string') { const owner = accounts.indexOf('我') >= 0 ? '我' : accounts[0]; notes = {}; if (owner) notes[owner] = s.notes; }
  const rows = SETTING_DEFS.map(d => ({
    key: d[1], value: d[0] === 'accounts' ? accounts.join(',') : (s[d[0]] === undefined ? d[2] : s[d[0]]), desc: d[3]
  }));
  accounts.forEach(a => rows.push({ key: NOTE_PREFIX + a, value: notes[a] || '', desc: a + ' 的策略筆記' }));
  const pw = s.pw || {};
  accounts.forEach(a => { if (pw[a]) rows.push({ key: PW_PREFIX + a, value: pw[a], desc: a + ' 的密碼（雜湊）' }); });
  writeT_('SET', rows);
}

function isGfExpr_(g) { return /[()*\/+]/.test(String(g || '')); }

function readAssetRows_() {
  return readT_('ASSET').map(a => ({
    name: str_(a.name), gf: str_(a.gf), cg: str_(a.cg), manual: num_(a.manual), qd: intOr_(a.qd, 4), pd: intOr_(a.pd, 2),
    gfPrice: num_(a.gfPrice), note: str_(a.note), logo: str_(a.logo), ccy: normCcy_(a.ccy), unit: str_(a.unit)
  })).filter(a => a.name);
}

function writeAssets_(rows) {
  const sh = sheet_(SH.ASSET);
  if (sh.getMaxColumns() < COLS.ASSET.length) sh.insertColumnsAfter(sh.getMaxColumns(), COLS.ASSET.length - sh.getMaxColumns());
  sh.getRange(1, 1, 1, COLS.ASSET.length).setValues([COLS.ASSET]).setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
  if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setNumberFormat('@');   // 名稱、代碼一律當文字（避免運算式被解讀）
  writeT_('ASSET', rows.map(a => Object.assign({}, a, { gfPrice: '' })));
  if (rows.length) {
    // 代碼欄可以是單純代碼（NASDAQ:AAPL），也可以是運算式（GOOGLEFINANCE("TPE:2330")*1000）
    const f = rows.map((a, i) => [isGfExpr_(a.gf) ? `=IFERROR(${String(a.gf).replace(/^=+/, '')},"")` : `=IFERROR(GOOGLEFINANCE(B${i + 2}),"")`]);
    sh.getRange(2, 7, rows.length, 1).setFormulas(f);
  }
}

/** 資產＋現價（手動 ＞ CoinGecko ＞ GOOGLEFINANCE） */
function readAssets_() {
  const rows = readAssetRows_();
  const cg = cgPrices_(rows.map(a => a.cg).filter(Boolean));
  return rows.map(a => {
    let price = null, source = '';
    if (a.manual > 0) { price = a.manual; source = '手動'; }
    else if (a.cg && cg[a.cg] && cg[a.cg][a.ccy.toLowerCase()] > 0) { price = cg[a.cg][a.ccy.toLowerCase()]; source = 'CoinGecko'; }
    else if (a.gfPrice > 0) { price = a.gfPrice; source = 'GOOGLEFINANCE'; }
    return { name: a.name, gf: a.gf, cg: a.cg, manual: a.manual, qd: a.qd, pd: a.pd, note: a.note, logo: a.logo, ccy: a.ccy, unit: a.unit, price: price, source: source };
  });
}

function cgCacheKey_(ids) { return 'cg2:' + ids.slice().sort().join(','); }

function cgPrices_(ids) {
  if (!ids.length) return {};
  const cache = CacheService.getScriptCache();
  const key = cgCacheKey_(ids);
  const hit = cache.get(key);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* 重抓 */ } }
  const props = PropertiesService.getScriptProperties();
  let out = {};
  try {
    const url = 'https://api.coingecko.com/api/v3/simple/price?vs_currencies=usd,cny,twd,hkd&ids=' + ids.map(encodeURIComponent).join(',');
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const data = JSON.parse(res.getContentText());
      ids.forEach(id => { if (data[id] && data[id].usd) out[id] = { usd: data[id].usd, cny: data[id].cny, twd: data[id].twd, hkd: data[id].hkd }; });
      cache.put(key, JSON.stringify(out), 180);
      props.setProperty('cgLast2', JSON.stringify(out));
      return out;
    }
  } catch (e) { console.warn('CoinGecko', e); }
  try { out = JSON.parse(props.getProperty('cgLast2') || 'null'); } catch (e) { out = null; }
  if (!out) {   // 舊版只存美元
    try { const o = JSON.parse(props.getProperty('cgLast') || '{}'); out = {}; Object.keys(o).forEach(k => out[k] = { usd: o[k] }); } catch (e) { out = {}; }
  }
  return out;
}

// ───────────────────────── 匯率（美元／人民幣／台幣） ─────────────────────────
// 來源：fawazahmed0 currency-api（免金鑰，每日匯率，2024-03 起有歷史資料）；
// 備援網址 currency-api.pages.dev。歷史匯率存在「匯率」工作表，抓過就不再抓。
const CCYS = ['USD', 'CNY', 'TWD', 'HKD'];
const FX_SHEET = '匯率';
const FX_FIRST = '2024-03-06';
// 2024-03 以前沒有每日匯率，用年平均（近似值；1 USD = ?）
const FX_OLD = {
  2013: { CNY: 6.15, TWD: 29.8, HKD: 7.76 }, 2014: { CNY: 6.16, TWD: 30.4, HKD: 7.75 }, 2015: { CNY: 6.23, TWD: 31.8, HKD: 7.75 },
  2016: { CNY: 6.64, TWD: 32.3, HKD: 7.76 }, 2017: { CNY: 6.76, TWD: 30.4, HKD: 7.79 }, 2018: { CNY: 6.62, TWD: 30.2, HKD: 7.84 },
  2019: { CNY: 6.91, TWD: 30.9, HKD: 7.84 }, 2020: { CNY: 6.90, TWD: 29.5, HKD: 7.76 }, 2021: { CNY: 6.45, TWD: 28.0, HKD: 7.77 },
  2022: { CNY: 6.73, TWD: 29.8, HKD: 7.83 }, 2023: { CNY: 7.08, TWD: 31.2, HKD: 7.83 }, 2024: { CNY: 7.19, TWD: 31.4, HKD: 7.82 }
};
function normCcy_(c) { c = str_(c).toUpperCase(); if (c === 'RMB') c = 'CNY'; if (c === 'NTD') c = 'TWD'; return CCYS.indexOf(c) >= 0 ? c : 'USD'; }
function fxUrls_(d) {
  return ['https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@' + d + '/v1/currencies/usd.min.json',
    'https://' + d + '.currency-api.pages.dev/v1/currencies/usd.min.json'];
}
function fxParse_(res) {
  try {
    if (res.getResponseCode() !== 200) return null;
    const j = JSON.parse(res.getContentText());
    const u = j && j.usd;
    if (!u || !(u.cny > 0) || !(u.twd > 0) || !(u.hkd > 0)) return null;
    return { CNY: u.cny, TWD: u.twd, HKD: u.hkd, date: str_(j.date) };
  } catch (e) { return null; }
}
/** 一次抓多個日期（含 'latest'），先主網址、失敗的再用備援 */
function fxFetchMany_(dates) {
  const out = {};
  if (!dates.length) return out;
  [0, 1].forEach(k => {
    const need = dates.filter(d => !out[d]);
    if (!need.length) return;
    let res = [];
    try { res = UrlFetchApp.fetchAll(need.map(d => ({ url: fxUrls_(d)[k], muteHttpExceptions: true }))); } catch (e) { console.warn('fx fetch', e); return; }
    need.forEach((d, i) => { const r = res[i] && fxParse_(res[i]); if (r) out[d] = r; });
  });
  return out;
}
function fxLatest_() {
  const cache = CacheService.getScriptCache(), props = PropertiesService.getScriptProperties();
  const hit = cache.get('fx:latest');
  if (hit) { try { const h = JSON.parse(hit); if (h && h.HKD) return h; } catch (e) { /* 重抓 */ } }
  const r = fxFetchMany_(['latest']).latest;
  if (r) { r.source = 'currency-api'; cache.put('fx:latest', JSON.stringify(r), 3600); props.setProperty('fxLast', JSON.stringify(r)); return r; }
  try { const last = JSON.parse(props.getProperty('fxLast') || 'null'); if (last) { last.stale = true; return last; } } catch (e) { /* 無 */ }
  return null;
}
/** 取得這些日期（yyyy-MM-dd）的匯率；沒抓過的會去抓（一次最多 25 個）並存到「匯率」工作表 */
function fxHist_(dates) {
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const ss = ss_();
  let sh = ss.getSheetByName(FX_SHEET);
  const have = {};
  if (sh && sh.getLastRow() > 1) {
    const nc = Math.min(5, sh.getMaxColumns());
    sh.getRange(2, 1, sh.getLastRow() - 1, nc).getValues().forEach(r => {
      const d = r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : str_(r[0]);
      if (d && num_(r[1]) > 0 && num_(r[2]) > 0 && num_(r[4]) > 0) have[d] = { CNY: num_(r[1]), TWD: num_(r[2]), HKD: num_(r[4]) };
    });
  }
  const key = d => d < FX_FIRST ? FX_FIRST : d;
  const want = [];
  dates.forEach(d => { if (d < FX_FIRST) return; const k = key(d); if (k < today && !have[k] && want.indexOf(k) < 0) want.push(k); });
  const got = fxFetchMany_(want.slice(0, 25));
  const rows = Object.keys(got).map(d => [d, got[d].CNY, got[d].TWD, got[d].date || d, got[d].HKD]);
  if (rows.length) {
    if (!sh) {
      sh = ss.insertSheet(FX_SHEET);
      sh.getRange(1, 1, 1, 5).setValues([['日期', 'USD→CNY', 'USD→TWD', '資料日期', 'USD→HKD']]).setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    if (sh.getMaxColumns() < 5) sh.insertColumnsAfter(sh.getMaxColumns(), 5 - sh.getMaxColumns());
    if (!str_(sh.getRange(1, 5).getValue())) sh.getRange(1, 5).setValue('USD→HKD').setFontWeight('bold').setBackground('#22313f').setFontColor('#ffffff');
    const r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(r0, 4, rows.length, 1).setNumberFormat('@');
    sh.getRange(r0, 1, rows.length, 5).setValues(rows);
    rows.forEach(r => have[r[0]] = { CNY: r[1], TWD: r[2], HKD: r[4] });
  }
  const out = {};
  dates.forEach(d => {
    if (d < FX_FIRST) { const y = +d.slice(0, 4), r = FX_OLD[y] || FX_OLD[y < 2013 ? 2013 : 2024]; out[d] = Object.assign({ approx: true }, r); return; }
    const k = key(d); if (have[k]) out[d] = have[k];
  });
  return out;
}
function fxForLoad_(trades) {
  try {
    const latest = fxLatest_();
    const ds = [];
    trades.forEach(t => { if (t.d && ds.indexOf(t.d) < 0) ds.push(t.d); });
    let hist = {};
    try { hist = fxHist_(ds); } catch (e) { console.warn('fx hist', e); }
    return { latest: latest, hist: hist };
  } catch (e) { console.warn('fx', e); return { latest: null, hist: {} }; }
}


// ───────────────────────── 手機離線版（PWA）同步 API ─────────────────────────
// 前端放在 GitHub Pages，用 POST 呼叫這裡；每個請求都要帶同步金鑰（存在指令碼屬性 SYNC_KEY）。
// 寫入是一批「操作」：依序套用、每個操作有唯一 ID，重送不會重複執行。
function syncMutations_() {
  return { api_setStatus, api_recordFills, api_updateOrder, api_addOrder, api_deleteOrders, api_duplicate, api_addOrders,
    api_saveTrade, api_deleteTrade, api_saveReal, api_deleteReal, api_confirmReal, api_importReal, api_saveDone, api_deleteDone,
    api_saveHolding, api_deleteHolding, api_saveSettings, api_saveNotes, api_addAccount };
}
function syncCalls_() {
  return { api_load, api_refreshPrices, api_saveAi, api_aiList, api_aiDelete, api_aiStep, api_renameAccount, api_deleteAccount,
    api_saveAssets, api_addAccount, api_setPassword, api_priceHistory };
}
function safeEq_(a, b) {
  if (a.length !== b.length) return false;
  let d = 0;
  for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return d === 0;
}
function doPost(e) {
  const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
  let req;
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (x) { return out({ ok: false, error: '請求格式錯誤' }); }
  const key = PropertiesService.getScriptProperties().getProperty('SYNC_KEY');
  if (!key) return out({ ok: false, auth: true, error: '還沒啟用手機離線版：請在 Apps Script 執行 syncSetup' });
  if (!req.key || !safeEq_(String(req.key), key)) { Utilities.sleep(1000); return out({ ok: false, auth: true, error: '同步金鑰不正確' }); }
  try {
    if (Array.isArray(req.ops)) return out(syncOps_(req.ops));
    if (req.call) {
      const f = syncCalls_()[req.call];
      if (!f) throw new Error('不支援的功能 ' + req.call);
      return out({ ok: true, result: f.apply(null, Array.isArray(req.args) ? req.args : []) });
    }
    return out({ ok: true, result: api_load() });
  } catch (err) {
    return out({ ok: false, error: String(err && err.message || err) });
  }
}
const SYNC_SHEET = '_同步紀錄';
/** 暫時性的錯誤（鎖、服務忙碌、逾時）不算處理過，前端會再送 */
function syncTransient_(msg) { return /lock|timed? ?out|timeout|service.*(unavailable|failed|error)|internal error|try again|too many|exceeded maximum/i.test(msg); }
function syncOps_(ops) {
  const ss = ss_();
  let sh = ss.getSheetByName(SYNC_SHEET);
  if (!sh) {
    sh = ss.insertSheet(SYNC_SHEET);
    sh.getRange(1, 1, 1, 3).setValues([['操作 ID', '時間', '結果']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    try { sh.hideSheet(); } catch (x) { /* 隱藏失敗沒關係 */ }
  }
  const seen = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => seen[str_(r[0])] = 1);
  const fns = syncMutations_(), results = [], log = [];
  BATCH_ = true;
  try {
    for (let i = 0; i < Math.min(ops.length, 100); i++) {
      const op = ops[i] || {};
      const id = str_(op.id);
      if (!id) continue;
      if (seen[id]) { results.push({ id: id, ok: true, dup: true }); continue; }
      const f = fns[op.fn];
      if (!f) { results.push({ id: id, ok: false, error: '不支援的操作 ' + op.fn }); log.push([id, new Date(), 'err']); continue; }
      try {
        f.apply(null, Array.isArray(op.args) ? op.args : []);
        results.push({ id: id, ok: true }); log.push([id, new Date(), 'ok']); seen[id] = 1;
      } catch (err) {
        const msg = String(err && err.message || err);
        if (syncTransient_(msg)) { results.push({ id: id, ok: false, retry: true, error: msg }); break; }   // 保持順序：後面的下次再送
        results.push({ id: id, ok: false, error: msg }); log.push([id, new Date(), 'err: ' + msg.slice(0, 200)]); seen[id] = 1;
      }
    }
  } finally { BATCH_ = false; }
  if (log.length) {
    sh.getRange(sh.getLastRow() + 1, 1, log.length, 3).setValues(log);
    const extra = sh.getLastRow() - 1 - 3000;   // 只留最近 3000 筆
    if (extra > 0) sh.deleteRows(2, extra);
  }
  return { ok: true, results: results, state: api_load() };
}
/** 在編輯器執行（或從選單）：產生同步金鑰並顯示設定步驟。已經有金鑰就顯示現有的。 */
function syncSetup() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty('SYNC_KEY');
  const isNew = !key;
  if (!key) { key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); props.setProperty('SYNC_KEY', key); }
  const msg = (isNew ? '已產生同步金鑰。' : '目前的同步金鑰：') + '\n\n' + key + '\n\n' +
    '接下來：\n1. 部署 → 新增部署作業 → 網頁應用程式：執行身分「我」、存取權「任何人」→ 部署，複製網址（/exec 結尾）。\n' +
    '2. 打開手機離線版（GitHub Pages 網址），貼上這個網址和金鑰。\n\n' +
    '金鑰等於這份資料的鑰匙，不要分享。要換金鑰請執行 syncResetKey。\n' +
    '提醒：啟用後，原本的網頁介面網址要在後面加上 ?k=金鑰 才能開。';
  console.log(msg);
  try { SpreadsheetApp.getUi().alert('手機離線版設定', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* 從編輯器執行時沒有對話框 */ }
  return msg;
}
/** 換一把新的同步金鑰（舊金鑰立刻失效，各裝置要重新輸入） */
function syncResetKey() {
  PropertiesService.getScriptProperties().deleteProperty('SYNC_KEY');
  return syncSetup();
}

// ───────────────────────── 歷史價格（真實交易紀錄的圖表用） ─────────────────────────
const PH_SHEET = '歷史價格';
/** 回傳 { asset, ccy, source, points: [[yyyy-MM-dd, price], ...]（每週）, updated } */
function api_priceHistory(asset, force) {
  const a = readAssetRows_().find(x => x.name === asset);
  if (!a) throw new Error('找不到資產 ' + asset);
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const ss = ss_();
  let sh = ss.getSheetByName(PH_SHEET);
  const cached = phRead_(sh, asset, tz);
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (cached && !force && cached.points.length && daysBetween_(cached.updated, today) < 7) return cached;
  let res = null;
  const tries = [];
  try { res = phYahoo_(a); tries.push('yahoo:' + (res ? res.points.length : 'x')); } catch (e) { tries.push('yahoo:' + e.message); }
  if ((!res || res.points.length < 10) && (a.cg || !a.gf || /^CURRENCY:/.test(a.gf))) { try { const r4 = phKraken_(a); tries.push('kraken:' + (r4 ? r4.points.length : 'x')); if (r4 && (!res || r4.points.length > res.points.length)) res = r4; } catch (e) { tries.push('kraken:' + e.message); } }
  if ((!res || res.points.length < 10) && a.cg) { try { const r2 = phCoinGecko_(a); tries.push('cg:' + (r2 ? r2.points.length : 'x')); if (r2 && (!res || r2.points.length > res.points.length)) res = r2; } catch (e) { tries.push('cg:' + e.message); } }
  if ((!res || res.points.length < 10) && (a.gf || /^[A-Z]{2,6}$/.test(a.name))) { try { const r3 = phGoogleFinance_(a, sh); tries.push('gf:' + (r3 ? r3.points.length : 'x')); if (r3 && (!res || r3.points.length > res.points.length)) res = r3; } catch (e) { tries.push('gf:' + e.message); } }
  console.log('priceHistory ' + asset + ' ' + tries.join(' | '));
  if (!res || !res.points.length) {
    if (cached && cached.points.length) return cached;
    return { asset: asset, ccy: a.ccy, source: '', points: [], updated: today, note: '抓不到歷史價格（' + tries.join('；') + '）' };
  }
  // 存進工作表（整個資產重寫）
  if (!sh) { sh = ss.insertSheet(PH_SHEET); sh.getRange(1, 1, 1, 5).setValues([['資產', '日期', '價格', '來源', '更新日']]).setFontWeight('bold'); sh.setFrozenRows(1); try { sh.hideSheet(); } catch (e) { /* 無 */ } }
  const keep = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().filter(r => str_(r[0]) !== asset) : [];
  const rows = keep.concat(res.points.map(p => [asset, p[0], p[1], res.source, today]));
  sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 5).clearContent();
  if (rows.length) { if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows()); sh.getRange(2, 2, rows.length, 1).setNumberFormat('@'); sh.getRange(2, 1, rows.length, 5).setValues(rows); }
  return { asset: asset, ccy: a.ccy, source: res.source, points: res.points, updated: today };
}
function daysBetween_(d1, d2) { return Math.round((new Date(d2) - new Date(d1)) / 864e5); }
function phRead_(sh, asset, tz) {
  if (!sh || sh.getLastRow() < 2) return null;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues().filter(r => str_(r[0]) === asset);
  if (!rows.length) return null;
  const f = v => v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : str_(v);
  const points = rows.map(r => [f(r[1]), num_(r[2])]).filter(p => p[0] && p[1] > 0).sort((x, y) => x[0] < y[0] ? -1 : 1);
  return { asset: asset, ccy: '', source: str_(rows[0][3]), points: points, updated: f(rows[0][4]) };
}
/** 把每日或任意間隔的資料壓成每週一點（取該週最後一筆），並去掉無效值 */
function phWeekly_(pairs) {
  const out = [], seen = {};
  pairs.forEach(p => {
    if (!(p[1] > 0)) return;
    const d = new Date(p[0]); const wk = Math.floor((d.getTime() + 3 * 864e5) / (7 * 864e5));
    seen[wk] = p;
  });
  Object.keys(seen).sort((x, y) => x - y).forEach(k => out.push([seen[k][0], round_(seen[k][1], 6)]));
  return out;
}
function phYahooSymbol_(a) {
  const gf = String(a.gf || '');
  const m = /([A-Z]{2,8}):([A-Za-z0-9.\-]+)/.exec(gf);
  if (m) {
    const ex = m[1], t = m[2];
    if (ex === 'TPE') return t + '.TW';
    if (ex === 'TWO') return t + '.TWO';
    if (ex === 'HKG') return ('0000' + t).slice(-4) + '.HK';
    if (ex === 'SHA') return t + '.SS';
    if (ex === 'SHE') return t + '.SZ';
    if (ex === 'TYO') return t + '.T';
    if (ex === 'CURRENCY') return t.length === 6 ? t.slice(0, 3) + '-' + t.slice(3) : null;
    return t;
  }
  if (/^\d{4,5}$/.test(gf)) return null;
  if (a.cg || !gf) return a.name.toUpperCase() + '-' + a.ccy;   // 加密貨幣
  return null;
}
function phYahoo_(a) {
  const sym = phYahooSymbol_(a);
  if (!sym) return null;
  const mult = (/\)\s*\*\s*(\d+(?:\.\d+)?)/.exec(String(a.gf || '')) || [0, 1])[1] * 1 || 1;
  const hosts = ['query1.finance.yahoo.com', 'query2.finance.yahoo.com'];
  let res = null, code = 0;
  for (let i = 0; i < hosts.length; i++) {
    const url = 'https://' + hosts[i] + '/v8/finance/chart/' + encodeURIComponent(sym) + '?range=max&interval=1wk';
    res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' } });
    code = res.getResponseCode();
    if (code === 200) break;
  }
  if (code !== 200) throw new Error('HTTP ' + code + '（' + sym + '）');
  const j = JSON.parse(res.getContentText());
  const r = j && j.chart && j.chart.result && j.chart.result[0];
  if (!r || !r.timestamp) throw new Error('無資料');
  const close = r.indicators.quote[0].close;
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const pairs = r.timestamp.map((t, i) => [Utilities.formatDate(new Date(t * 1000), tz, 'yyyy-MM-dd'), close[i] * mult]);
  const quoteCcy = r.meta && r.meta.currency ? String(r.meta.currency).toUpperCase() : a.ccy;
  if (quoteCcy !== a.ccy && !(quoteCcy === 'TWD' && a.ccy === 'TWD')) {
    // 報價幣別和資產設定不同（例如 Yahoo 的 ADA-USD 但資產設 TWD）→ 用最新匯率換
    const fx = fxLatest_();
    const rate = c => c === 'USD' ? 1 : fx && fx[c];
    if (fx && rate(quoteCcy) && rate(a.ccy)) pairs.forEach(p => { p[1] = p[1] / rate(quoteCcy) * rate(a.ccy); });
  }
  return { source: 'Yahoo Finance（' + sym + '，每週）', points: phWeekly_(pairs) };
}
/** Kraken 公開 OHLC：每週一點，最多 720 點（約 14 年），不用金鑰 */
function phKraken_(a) {
  const base = a.name.toUpperCase() === 'BTC' ? 'XBT' : a.name.toUpperCase();
  const quote = a.ccy === 'USD' ? 'USD' : 'USD';   // Kraken 以美元報價，其他幣別用匯率換
  const url = 'https://api.kraken.com/0/public/OHLC?pair=' + base + quote + '&interval=10080';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  const j = JSON.parse(res.getContentText());
  if (!j || (j.error && j.error.length)) throw new Error((j && j.error && j.error[0]) || '無資料');
  const key = Object.keys(j.result || {}).filter(k => k !== 'last')[0];
  if (!key) throw new Error('無資料');
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  const pairs = j.result[key].map(r => [Utilities.formatDate(new Date(r[0] * 1000), tz, 'yyyy-MM-dd'), +r[4]]);
  if (a.ccy !== 'USD') { const fx = fxLatest_(); if (fx && fx[a.ccy]) pairs.forEach(p => { p[1] = p[1] * fx[a.ccy]; }); }
  return { source: 'Kraken（' + base + '/USD，每週）', points: phWeekly_(pairs) };
}
function phCoinGecko_(a) {
  const url = 'https://api.coingecko.com/api/v3/coins/' + encodeURIComponent(a.cg) + '/market_chart?vs_currency=' + a.ccy.toLowerCase() + '&days=365&interval=daily';
  let res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() === 429) { Utilities.sleep(2500); res = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); }
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode());
  const j = JSON.parse(res.getContentText());
  if (!j || !j.prices) throw new Error('無資料');
  const tz = ss_().getSpreadsheetTimeZone() || 'Asia/Taipei';
  return { source: 'CoinGecko（最近一年，每週）', points: phWeekly_(j.prices.map(p => [Utilities.formatDate(new Date(p[0]), tz, 'yyyy-MM-dd'), p[1]])) };
}
/** 用 GOOGLEFINANCE 的歷史報價（在隱藏工作表暫時放一個公式） */
function phGoogleFinance_(a, sh) {
  const gf = String(a.gf || '');
  const mm = /GOOGLEFINANCE\("([^"]+)"\)\s*(?:\*\s*(\d+(?:\.\d+)?))?/.exec(gf);
  const ticker = mm ? mm[1] : (gf && !isGfExpr_(gf) ? gf : 'CURRENCY:' + a.name.toUpperCase() + a.ccy);
  const mult = mm && mm[2] ? +mm[2] : 1;
  const ss = ss_();
  const tmp = ss.insertSheet('_ph_' + Date.now());
  try {
    tmp.getRange('A1').setFormula('=IFERROR(GOOGLEFINANCE("' + ticker + '","price",DATE(2013,1,1),TODAY(),"WEEKLY"),"")');
    SpreadsheetApp.flush();
    let n = 0;
    for (let i = 0; i < 8 && n < 3; i++) { Utilities.sleep(1500); SpreadsheetApp.flush(); n = tmp.getLastRow(); }
    if (n < 3) throw new Error('GOOGLEFINANCE 沒有歷史資料');
    const v = tmp.getRange(2, 1, n - 1, 2).getValues();
    const tz = ss.getSpreadsheetTimeZone() || 'Asia/Taipei';
    const pairs = v.filter(r => r[0] instanceof Date && num_(r[1]) > 0).map(r => [Utilities.formatDate(r[0], tz, 'yyyy-MM-dd'), num_(r[1]) * mult]);
    return { source: 'GOOGLEFINANCE（' + ticker + '，每週）', points: phWeekly_(pairs) };
  } finally { ss.deleteSheet(tmp); }
}

// ───────────────────────── 小工具 ─────────────────────────
function ss_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('找不到試算表：這份 Apps Script 必須從試算表的「擴充功能 → Apps Script」建立。');
  return ss;
}
function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('找不到工作表「' + name + '」，請先執行 setup');
  return s;
}
function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : null;
  if (v === null || v === undefined || v === '' || v instanceof Date) return null;
  const n = parseFloat(String(v).replace(/[,$\s]/g, ''));
  return isFinite(n) ? n : null;
}
function ms_(v) { return v instanceof Date ? v.getTime() : (typeof v === 'number' && v > 1e11 ? v : null); }
function cell_(v) { return v === null || v === undefined || (typeof v === 'number' && !isFinite(v)) ? '' : v; }
function intOr_(v, d) { const n = parseInt(v, 10); return isNaN(n) ? d : n; }
function round_(v, d) { const m = Math.pow(10, d); return Math.round(v * m) / m; }
function toSet_(arr) { const s = {}; (arr || []).forEach(x => s[x] = 1); return s; }
function nextIds_(prefix, existing, n) {
  let max = 0;
  const re = new RegExp('^' + prefix + '(\\d+)$');
  (existing || []).forEach(x => { const m = re.exec(String(x)); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  const out = [];
  for (let i = 0; i < n; i++) out.push(prefix + String(++max).padStart(4, '0'));
  return out;
}
function assignIds_(orders, prefix) {
  const seen = {}, need = [];
  orders.forEach(o => { if (o.id && !seen[o.id]) seen[o.id] = 1; else need.push(o); });
  const ids = nextIds_(prefix, Object.keys(seen), need.length);
  need.forEach((o, i) => o.id = ids[i]);
}
