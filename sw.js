const CACHE = "time-stat-v43";

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

function isLocalDevScope() {
  try {
    const u = new URL(self.registration.scope);
    if (u.hostname === "localhost" || u.hostname === "127.0.0.1" || u.hostname === "[::1]") return true;
    if (u.port === "8765") return true;
  } catch (_) {}
  return false;
}

self.addEventListener("fetch", (e) => {
  if (isLocalDevScope()) {
    e.respondWith(fetch(e.request));
    return;
  }

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
            "<!DOCTYPE html><html lang=zh-Hant><meta charset=utf-8><title>無法載入</title><p>請檢查網絡或稍後再試。</p></html>",
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
