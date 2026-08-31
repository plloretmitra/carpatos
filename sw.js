const CACHE = 'carpatos-v1';

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll([
    '/',
    'carpatos_pwa.html',
    'manifest.json',
    'portada.jpg'
  ])));
  self.skipWaiting();
});

self.addEventListener('fetch', e => {
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
