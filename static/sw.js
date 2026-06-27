// Service Worker for 工程單助手 PWA
const CACHE = 'eg-don-' + self.registration.scope.replace(/[^a-z0-9]/g,'');
const ASSETS = [
  '/',
  '/static/dist/output.css',
  '/static/style.css',
  '/manifest.json',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((cache) => cache.addAll(ASSETS).catch(() => {}))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  // Network-first for API calls
  if (e.request.url.includes('/api/') || e.request.url.includes('/generate') || e.request.url.includes('/download/')) {
    e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
    return;
  }
  // Cache-first for static assets
  e.respondWith(
    caches.match(e.request).then((cached) => cached || fetch(e.request).then((resp) => {
      if (resp.ok) {
        const clone = resp.clone();
        caches.open(CACHE).then((cache) => cache.put(e.request, clone));
      }
      return resp;
    }))
  );
});
