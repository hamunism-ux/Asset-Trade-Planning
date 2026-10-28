// 資產掛單規劃本：離線快取（只快取介面檔案，不快取任何資料請求）
const SHELL = 'ladder-shell-8c48e6bd';
const RUNTIME = 'ladder-runtime';
const FILES = ['./', 'index.html', 'manifest.webmanifest', 'icon-180.png', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(SHELL).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== SHELL && k !== RUNTIME).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // 同步（POST）一律直接走網路
  const url = new URL(req.url);
  if (url.origin === location.origin) {
    // 先用快取（打開就有畫面），同時在背景抓新版，下次打開就是新版
    e.respondWith(caches.open(SHELL).then(c => c.match(req, { ignoreSearch: true }).then(hit => {
      const net = fetch(req).then(res => { if (res.ok) c.put(req, res.clone()); return res; }).catch(() => hit);
      return hit || net;
    })));
    return;
  }
  if (url.hostname === 'cdnjs.cloudflare.com') {             // PDF 元件：第一次用到時快取
    e.respondWith(caches.open(RUNTIME).then(c => c.match(req).then(hit => hit || fetch(req).then(res => { if (res.ok) c.put(req, res.clone()); return res; }))));
  }
  // 其他（Apps Script、CoinGecko…）不處理，直接走網路
});
