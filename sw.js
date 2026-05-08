const CACHE = "time-stat-v42";

self.addEventListener("install", (e) => {
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
  );
  self.clients.claim();
});

function indexHref() {
  return new URL("index.html", self.registration.scope).href;
}

self.addEventListener("fetch", (e) => {
  const req = e.request;
  const url = new URL(req.url);

  if (req.mode === "navigate") {
    e.respondWith(
      (async () => {
        try {
          return await fetch(req);
        } catch {
          const c = await caches.match(indexHref());
          if (c) return c;
          return await fetch(indexHref());
        }
      })().catch(
        () =>
          new Response(
            "<!DOCTYPE html><html lang=zh-Hant><meta charset=utf-8><title>無法載入</title><p>請檢查本機 server 是否仍開住。</p></html>",
            { headers: { "Content-Type": "text/html; charset=utf-8" }, status: 503 }
          )
      )
    );
    return;
  }

  const isAsset =
    url.pathname.endsWith(".js") ||
    url.pathname.endsWith(".css") ||
    url.pathname.endsWith(".json") ||
    url.pathname.endsWith(".svg");

  if (isAsset) {
    e.respondWith(
      (async () => {
        try {
          const res = await fetch(req);
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE).then((cache) => cache.put(req, copy));
          }
          return res;
        } catch {
          const cached = await caches.match(req, { ignoreSearch: true });
          if (cached) return cached;
          return await fetch(req);
        }
      })().catch(
        () => new Response("/* network error */", { status: 503, headers: { "Content-Type": "text/plain" } })
      )
    );
    return;
  }

  e.respondWith(
    caches
      .match(req)
      .then((r) => r || fetch(req))
      .catch(() => fetch(req))
      .catch(() => new Response("", { status: 503 }))
  );
});
