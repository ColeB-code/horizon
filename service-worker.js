const CACHE_NAME = "family-horizon-v29";
const APP_SHELL = [
  "index.html",
  "css/styles.css",
  "js/app.js",
  "js/auth.js",
  "js/graph.js",
  "manifest.json",
  "assets/icon-192.png",
  "assets/icon-512.png"
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Never cache auth or live Graph API calls.
  if (url.origin.includes("microsoftonline.com") || url.origin.includes("graph.microsoft.com")) {
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
