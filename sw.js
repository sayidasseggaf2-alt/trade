const CACHE = 'tradelog-v1';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'apple-touch-icon.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const h = new URL(r.url).hostname;
  if (h.endsWith('script.google.com') || h.endsWith('googleusercontent.com')) return;   // data tidak di-cache
  e.respondWith(fetch(r).then(res => {
    const copy = res.clone(); caches.open(CACHE).then(c => c.put(r, copy)).catch(() => {}); return res;
  }).catch(() => caches.match(r).then(m => m || caches.match('index.html'))));
});
