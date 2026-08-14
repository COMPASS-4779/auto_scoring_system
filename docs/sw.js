/* COMPASS 採点システム PWA – 最小サービスワーカー（インストール可能化＋シェルキャッシュ） */
const CACHE = 'scoring-pwa-v1';
const SHELL = ['./','./index.html','./manifest.webmanifest'];
self.addEventListener('install', (e) => { e.waitUntil(caches.open(CACHE).then((c)=>c.addAll(SHELL)).then(()=>self.skipWaiting())); });
self.addEventListener('activate', (e) => { e.waitUntil(caches.keys().then((ks)=>Promise.all(ks.filter((k)=>k!==CACHE).map((k)=>caches.delete(k)))).then(()=>self.clients.claim())); });
self.addEventListener('fetch', (e) => {
  const url = new URL(e.request.url);
  if (url.origin === self.location.origin) { e.respondWith(caches.match(e.request).then((r)=>r||fetch(e.request))); }
});
