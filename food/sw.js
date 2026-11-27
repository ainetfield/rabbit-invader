// 食材きげん帳 Service Worker: アプリ本体をキャッシュしてオフラインでも開けるようにする
const CACHE = "kigencho-v1";
const ASSETS = ["./", "./index.html", "./manifest.webmanifest", "./icon.svg", "./icon-192.png", "./icon-180.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener("fetch", (e) => {
  const req = e.request;
  // 同一オリジンの GET のみ対象（共有サーバーへの通信はそのまま通す）
  if (req.method !== "GET" || new URL(req.url).origin !== location.origin) return;
  // ネットワーク優先・失敗したらキャッシュ
  e.respondWith(
    fetch(req)
      .then((res) => {
        const copy = res.clone();
        caches.open(CACHE).then((c) => c.put(req, copy));
        return res;
      })
      .catch(() => caches.match(req).then((r) => r || caches.match("./")))
  );
});
