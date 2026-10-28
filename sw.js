const CACHE_NAME = 'correctgloss-home-v1';
const APP_FILES = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];
self.addEventListener('install', event => {
  event.waitUntil(caches.open(CACHE_NAME).then(cache => cache.addAll(APP_FILES)));
  self.skipWaiting();
});
self.addEventListener('activate', event => {
  event.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)))));
  self.clients.claim();
});
self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.method !== 'GET') return;
  event.respondWith(caches.match(req).then(hit => hit || fetch(req).then(resp => {
    const copy = resp.clone();
    if (new URL(req.url).origin === self.location.origin) caches.open(CACHE_NAME).then(cache => cache.put(req, copy));
    return resp;
  }).catch(() => hit)));
});
