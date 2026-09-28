const CACHE = "pagos-tienda-v3";
const ARCHIVOS = ["./", "index.html", "estilos.css", "manifest.webmanifest", "js/app.js", "js/excel.js",
  "js/gemini.js", "js/carpeta.js", "vendor/exceljs.min.js", "iconos/icono-192.png", "iconos/icono-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ARCHIVOS)));
  self.skipWaiting();
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

// primero la red, así las mejoras llegan solas; si no hay internet, lo guardado
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request)
      .then(r => {
        const copia = r.clone();
        caches.open(CACHE).then(c => c.put(e.request, copia));
        return r;
      })
      .catch(() => caches.match(e.request))
  );
});
