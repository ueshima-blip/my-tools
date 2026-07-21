/* タスクボード Service Worker
   - GitHub Pages など https で「アプリとしてインストール」したとき、
     オフラインでも起動できるよう、アプリ一式をキャッシュします。
   - アプリ本体(HTML)は「ネットワーク優先」＝オンライン時は常に最新を取得し、
     更新がすぐ反映されます。オフライン時だけキャッシュを使います。
   - file:// では登録されません（index.html 側でガード）。 */
const CACHE = 'taskboard-v3';
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

  const isPage = req.mode === 'navigate' || req.destination === 'document';
  if (isPage) {
    // アプリ本体：ネットワーク優先（最新を取りに行き、取れたらキャッシュ更新。ダメならキャッシュ）
    e.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(CACHE).then((c) => c.put('./index.html', copy)).catch(() => {});
          return res;
        })
        .catch(() => caches.match('./index.html').then((r) => r || caches.match('./')))
    );
    return;
  }

  // それ以外（アイコン・manifest 等）：キャッシュ優先
  e.respondWith(
    caches.match(req).then((hit) => {
      if (hit) return hit;
      return fetch(req)
        .then((res) => {
          try {
            const url = new URL(req.url);
            if (url.origin === self.location.origin) {
              const copy = res.clone();
              caches.open(CACHE).then((c) => c.put(req, copy));
            }
          } catch (_) {}
          return res;
        })
        .catch(() => caches.match('./index.html'));
    })
  );
});
