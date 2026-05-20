// ★ HTMLを更新するたびにこのバージョン番号を上げてください
const CACHE_VERSION = 'v2';
const CACHE_NAME = 'fielder-tool-' + CACHE_VERSION;

// インストール時：キャッシュを作成
self.addEventListener('install', event => {
  self.skipWaiting();
});

// アクティベート時：古いキャッシュを全削除
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

// フェッチ時：常にネットワークを優先し、失敗時のみキャッシュを使用
self.addEventListener('fetch', event => {
  event.respondWith(
    fetch(event.request)
      .then(response => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
