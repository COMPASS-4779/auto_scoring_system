/* 理解度・評価 統合教育システム PWA – 最小サービスワーカー */
const CACHE='rikai-pwa-v1';
const SHELL=['./','./index.html','./manifest.webmanifest'];
self.addEventListener('install',(e)=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));});
self.addEventListener('activate',(e)=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',(e)=>{const u=new URL(e.request.url); if(u.origin===self.location.origin){e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));}});
