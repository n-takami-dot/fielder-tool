// Service Worker：常にネットワークから取得（キャッシュなし）
// オフライン対応より常に最新版を優先する設計

const CACHE_NAME = 'fielder-tool-v46';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  // 古いキャッシュを全て削除
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.map(key => caches.delete(key)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  // GASへのリクエストはService Worker経由させない
  if(event.request.url.includes('script.google.com')){
    return;
  }
  // HTMLファイルは常にネットワークから取得（キャッシュしない）
  if(event.request.url.includes('.html') || event.request.url.endsWith('/')){
    event.respondWith(
      fetch(event.request, {cache: 'no-store'})
        .catch(() => caches.match(event.request))
    );
    return;
  }
  // CSS・JS・画像はキャッシュ優先（CDNリソース等）
  event.respondWith(
    caches.match(event.request).then(cached => {
      const fetchPromise = fetch(event.request).then(response => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        return response;
      });
      return cached || fetchPromise;
    })
  );
});
