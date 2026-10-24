// Cache-first offline copy of the app. Bump CACHE on every release so installed phones fetch the new files.
const CACHE = 'inventory-v15';
const FILES = ['./', 'index.html', 'style.css', 'app.js', 'logic.js', 'excel.js', 'exceljs.min.js', 'manifest.json', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  self.skipWaiting();
  // cache: 'reload' skips the browser's HTTP cache, so a quick follow-up release can't store the previous files.
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES.map(f => new Request(f, { cache: 'reload' })))));
});
self.addEventListener('activate', e => e.waitUntil(
  caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())));
self.addEventListener('fetch', e => e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request))));
