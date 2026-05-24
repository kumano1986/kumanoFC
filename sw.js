// ServiceWorker完全無効化（PWAのキャッシュ問題対策）
self.addEventListener('install', function() { self.skipWaiting(); });
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.map(function(k){ return caches.delete(k); }));
    }).then(function() {
      return self.registration.unregister();
    }).then(function() {
      return self.clients.matchAll();
    }).then(function(clients) {
      clients.forEach(function(c) { c.navigate(c.url); });
    })
  );
});
self.addEventListener('fetch', function(e) {
  // 何もキャッシュせず常にネットワークから取得
  e.respondWith(fetch(e.request, { cache: 'no-store' }));
});
