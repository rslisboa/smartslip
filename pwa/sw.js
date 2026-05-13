const CACHE_NAME = "smartslip-shell-v2";

const CORE_ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icons/smartslip-192.png",
  "./icons/smartslip-512.png"
];

self.addEventListener("install", function(event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(CORE_ASSETS);
    })
  );

  self.skipWaiting();
});

self.addEventListener("activate", function(event) {
  event.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys
          .filter(function(key) {
            return key !== CACHE_NAME;
          })
          .map(function(key) {
            return caches.delete(key);
          })
      );
    })
  );

  self.clients.claim();
});

self.addEventListener("fetch", function(event) {
  const request = event.request;

  if (request.method !== "GET") {
    return;
  }

  const url = new URL(request.url);

  if (url.origin !== self.location.origin) {
    return;
  }

  const isHtmlNavigation =
    request.mode === "navigate" ||
    (request.headers.get("accept") || "").includes("text/html");

  if (isHtmlNavigation) {
    event.respondWith(
      fetch(request)
        .then(function(response) {
          const copy = response.clone();

          caches.open(CACHE_NAME).then(function(cache) {
            cache.put(request, copy);
          });

          return response;
        })
        .catch(function() {
          return caches.match("./index.html");
        })
    );

    return;
  }

  event.respondWith(
    caches.match(request).then(function(cached) {
      return cached || fetch(request).then(function(response) {
        const copy = response.clone();

        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(request, copy);
        });

        return response;
      });
    })
  );
});
