/** @OnlyCurrentDoc */
/**
 * ============================================================
 *  資產買賣紀錄與掛單規劃本 Asset Trade & Ladder Planner｜載入器
 *
 *  這個檔案只做一件事：從 GitHub 取得最新的伺服器程式（gas/core.js）與介面（gas/App.html）並執行。
 *  之後程式更新都不用再貼檔、不用重新部署：重新整理網頁就是新版（最多延遲 10 分鐘），
 *  或在試算表選單按「🧭 掛單工具 → ⬆ 更新程式到最新版」立刻更新。
 *
 *  安全：程式只從 https://github.com/hamunism-ux/Asset-Trade-Planning 讀取；抓不到時用上次成功的版本（存在指令碼屬性）。
 *  資料、同步金鑰、API Key 都不在 GitHub 上。
 * ============================================================
 */
const LOADER_VER = '1.0';
const CORE_RAW = 'https://raw.githubusercontent.com/hamunism-ux/Asset-Trade-Planning/main/gas/core.js';

/** 取得程式原始碼：快取（10 分鐘）→ GitHub → 上次成功的版本 */
function coreSource_() {
  const cache = CacheService.getScriptCache();
  const meta = cache.get('core:meta');
  if (meta) {
    try { const m = JSON.parse(meta); const parts = []; for (let i = 0; i < m.n; i++) { const c = cache.get('core:' + i); if (c === null) { parts.length = 0; break; } parts.push(c); } if (parts.length) return parts.join(''); } catch (e) { /* 重抓 */ }
  }
  let src = '';
  try {
    const res = UrlFetchApp.fetch(CORE_RAW + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) src = res.getContentText();
  } catch (e) { console.warn('GitHub 連線失敗：' + e.message); }
  if (src && /APP_VER = '/.test(src) && src.indexOf('function doGet') >= 0 && src.indexOf('__CORE_EXPORTS__') >= 0) {
    coreCache_(src); coreStore_(src);
    return src;
  }
  const last = coreLast_();
  if (last) { coreCache_(last); return last; }
  throw new Error('抓不到程式：GitHub 連不上，而且這份試算表還沒有存過任何版本。請檢查網路或稍後再試。');
}
function coreCache_(src) {
  const CH = 90000, n = Math.ceil(src.length / CH), kv = {};
  for (let i = 0; i < n; i++) kv['core:' + i] = src.slice(i * CH, (i + 1) * CH);
  kv['core:meta'] = JSON.stringify({ n: n, at: Date.now() });
  try { CacheService.getScriptCache().putAll(kv, 600); } catch (e) { /* 快取失敗沒關係 */ }
}
/** 上次成功的版本（指令碼屬性，每格最多 9KB → 分段存） */
function coreStore_(src) {
  const p = PropertiesService.getScriptProperties();
  const h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, src, Utilities.Charset.UTF_8));
  if (p.getProperty('coreLG:hash') === h) return;
  const CH = 8000, n = Math.ceil(src.length / CH), kv = { 'coreLG:n': String(n), 'coreLG:hash': h };
  for (let i = 0; i < n; i++) kv['coreLG:' + i] = src.slice(i * CH, (i + 1) * CH);
  p.setProperties(kv);
}
function coreLast_() {
  const p = PropertiesService.getScriptProperties();
  const n = +(p.getProperty('coreLG:n') || 0);
  if (!n) return '';
  const parts = [];
  for (let i = 0; i < n; i++) { const c = p.getProperty('coreLG:' + i); if (c === null) return ''; parts.push(c); }
  return parts.join('');
}
/** 執行程式：所有函式掛到全域（Seed.gs、Import*.gs 這類一次性腳本也能呼叫） */
let CORE_ = null, CORE_ERR_ = '';
function core_() {
  if (CORE_) return CORE_;
  const src = coreSource_();
  const C = (new Function(src))();
  Object.keys(C).forEach(k => { if (!(k in globalThis) || /^api_/.test(k)) globalThis[k] = C[k]; });
  CORE_ = C;
  return C;
}
try { core_(); } catch (e) { CORE_ERR_ = e.message; console.error('載入程式失敗：' + e.message); }

// ── 固定入口（Apps Script 需要靜態宣告的函式）──
function doGet(e) {
  if (CORE_ERR_ && !CORE_) return HtmlService.createHtmlOutput('<div style="font:15px/1.6 sans-serif;padding:40px;max-width:560px;margin:auto"><h3>載入程式失敗</h3><p>' + CORE_ERR_.replace(/</g, '&lt;') + '</p><p>請稍後重新整理。</p></div>').setTitle('資產買賣紀錄與掛單規劃本');
  return core_().doGet(e);
}
function doPost(e) { return core_().doPost(e); }
function rpc(name, args) { return core_().rpc(name, args); }
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🧭 掛單工具')
    .addItem('開啟網頁介面', 'showWebAppLink')
    .addItem('📱 手機離線版：同步金鑰與設定', 'syncSetup')
    .addSeparator()
    .addItem('⬆ 更新程式到最新版（GitHub）', 'updateFromGitHub')
    .addItem('介面來源：GitHub ↔ 專案內檔案', 'uiSourceToggle')
    .addItem('目前版本', 'showVersion')
    .addSeparator()
    .addItem('初始化（建立資料表＋匯入原 Excel）', 'setup')
    .addToUi();
}
function showWebAppLink() { return core_().showWebAppLink(); }
function syncSetup() { return core_().syncSetup(); }
function syncResetKey() { return core_().syncResetKey(); }
function setup() { return core_().setup(); }
function uiSourceToggle() { return core_().uiSourceToggle(); }
function updateFromGitHub() {
  CacheService.getScriptCache().remove('core:meta');
  CORE_ = null;
  return core_().updateFromGitHub();
}
function showVersion() {
  const v = core_().APP_VER;
  const msg = '伺服器程式 v' + v + '（載入器 v' + LOADER_VER + '）\n介面來源：' + core_().uiSource_() + '\n程式來源：GitHub hamunism-ux/Asset-Trade-Planning';
  try { SpreadsheetApp.getUi().alert('目前版本', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { console.log(msg); }
  return msg;
}
