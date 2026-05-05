// Service Worker - キャッシュを使わず常にネットワークから取得
self.addEventListener('install', function(e) {
  self.skipWaiting();
});
self.addEventListener('activate', function(e) {
  // 古いキャッシュを全削除
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.map(function(key) { return caches.delete(key); }));
    })
  );
  self.clients.claim();
});
self.addEventListener('fetch', function(e) {
  // 常にネットワークから取得（キャッシュしない）
  e.respondWith(fetch(e.request));
});
