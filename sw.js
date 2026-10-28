// Klubbsalg: gjør at appen kan startes uten nett etter første besøk.
// Selve salgsdataene ligger i telefonen og i Google Sheet, ikke her.
const CACHE = "klubbsalg-v8b";
const FILES = ["./", "./index.html"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const req = e.request, url = new URL(req.url);
  if (req.method !== "GET" || url.origin !== self.location.origin) return;   // Google Sheet-kall går alltid rett til nettet
  if (/\.(mp3|m4a|wav|ogg)$/i.test(url.pathname) || req.headers.has("range")) return;   // lyd håndteres av nettleseren (Safari krever delvise svar)
  e.respondWith(
    caches.open(CACHE).then(cache =>
      cache.match(req, { ignoreSearch: true }).then(hit => {
        const net = fetch(req).then(res => { if (res && res.status === 200) cache.put(req, res.clone()); return res; }).catch(() => null);
        if (hit) { e.waitUntil(net); return hit; }          // vis det som er lagret med en gang, og oppdater i bakgrunnen
        return net.then(res => res || (req.mode === "navigate" ? cache.match("./index.html") : Response.error()));
      })
    )
  );
});
