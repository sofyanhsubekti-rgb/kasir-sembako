/* Jaringan dulu, cache sebagai cadangan: aplikasi tetap terbuka saat internet
   putus, dan pembaruan tetap terambil begitu jaringan hidup lagi. */
const CACHE = "kasir-v2";

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil(self.clients.claim()));

self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  const u = new URL(e.request.url);
  if (u.origin !== self.location.origin) return;
  if (u.pathname.endsWith("/print") || u.pathname.endsWith("/ping")) return;

  e.respondWith((async () => {
    try {
      const r = await fetch(e.request);
      if (r && r.ok) (await caches.open(CACHE)).put(e.request, r.clone());
      return r;
    } catch (err) {
      const c = await caches.open(CACHE);
      const hit = (await c.match(e.request)) || (await c.match("./")) || (await c.match("index.html"));
      if (hit) return hit;
      throw err;
    }
  })());
});
