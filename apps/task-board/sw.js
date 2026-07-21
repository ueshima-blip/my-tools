/* タスクボード Service Worker
   - GitHub Pages など https で「アプリとしてインストール」したとき、
     オフラインでも起動できるよう、アプリ一式をキャッシュします。
   - file:// では登録されません（index.html 側でガード）。 */
const CACHE = 'taskboard-v2';
const ASSETS = ['./', './index.html', './manifest.webmanifest', './icon.svg'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;
  e.respondWith(
    caches.match(req).then((hit) => {
      if (hit) return hit;
      return fetch(req)
        .then((res) => {
          // 同一オリジンの GET は取得ついでにキャッシュ
          try {
            const url = new URL(req.url);
            if (url.origin === self.location.origin) {
              const copy = res.clone();
              caches.open(CACHE).then((c) => c.put(req, copy));
            }
          } catch (_) {}
          return res;
        })
        .catch(() => caches.match('./index.html')); // オフライン時はアプリ本体を返す
    })
  );
});
