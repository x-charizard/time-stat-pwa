const CACHE = "time-stat-v41";

self.addEventListener("install", (e) => {
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

/** 導航：線上優先；靜態資源：線上優先再寫入 cache（避免永遠食舊 app.js） */
self.addEventListener("fetch", (e) => {
  const req = e.request;
  const url = new URL(req.url);

  if (req.mode === "navigate") {
    e.respondWith(fetch(req).catch(() => caches.match("./index.html")));
    return;
  }

  const path = url.pathname;
  const isAsset =
    path.endsWith(".js") ||
    path.endsWith(".css") ||
    path.endsWith(".json") ||
    path.endsWith(".svg");

  if (isAsset) {
    e.respondWith(
      fetch(req)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(req, copy));
          }
          return res;
        })
        .catch(() => caches.match(req, { ignoreSearch: true }))
    );
    return;
  }

  e.respondWith(caches.match(req).then((r) => r || fetch(req)));
});
