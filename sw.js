// Nombre de cache versionado para forzar actualizacion cuando cambia la PWA.
const CACHE_NAME = "crm-jdl-pwa-v18";
// Archivos minimos que permiten abrir la app instalada con su apariencia base.
const APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./app.js",
  "./manifest.json?v=20260321c",
  "./favicon.ico?v=20260321c",
  "./icons/apple-touch-icon.png?v=20260321c",
  "./icons/icon-192.png?v=20260321c",
  "./icons/icon-512.png?v=20260321c",
  "./icons/icon-512-maskable.png?v=20260321c",
  "./Oficial_JDL_acua.png",
  "./Encabezadojdl.png",
  "./piedepaginajdl.png"
];

// Durante la instalacion se precargan los recursos principales de la app.
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

// Al activar, se eliminan caches viejos para evitar mezclar iconos o archivos obsoletos.
self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

// Intercepta lecturas GET del mismo origen para responder desde cache cuando conviene.
// Las rutas API se excluyen para no servir datos viejos del backend.
self.addEventListener("fetch", (event) => {
  const request = event.request;
  if (request.method !== "GET") return;

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;
  if (url.pathname.includes("/api/")) return;

  if (request.mode === "navigate") {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put("./index.html", copy));
          return response;
        })
        .catch(() => caches.match(request).then((cached) => cached || caches.match("./index.html") || caches.match("./")))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return fetch(request).then((response) => {
        if (!response || response.status !== 200 || response.type !== "basic") {
          return response;
        }
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      });
    })
  );
});
