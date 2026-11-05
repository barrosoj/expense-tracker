// Service worker: caches the app shell so the app works offline once loaded.
// Bump CACHE_VERSION whenever index.html, styles.css or app.js change.
var CACHE_VERSION = "v6";
var CACHE_NAME = "quadro-tarefas-" + CACHE_VERSION;
var APP_SHELL = [
  "./",
  "./index.html",
  "./styles.css",
  "./config.js",
  "./vendor/supabase.js",
  "./app.js",
  "./auth.js",
  "./icon.png",
  "./manifest.webmanifest"
];

self.addEventListener("install", function(event){
  event.waitUntil(
    caches.open(CACHE_NAME).then(function(cache){ return cache.addAll(APP_SHELL); })
      .then(function(){ return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function(event){
  event.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){
        return k.indexOf("quadro-tarefas-") === 0 && k !== CACHE_NAME;
      }).map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

// Network-first for same-origin GETs (so updates show up when online),
// falling back to the cache when offline.
self.addEventListener("fetch", function(event){
  var req = event.request;
  // Only same-origin GETs; Supabase API/auth calls (other origin) always go straight to the network.
  if (req.method !== "GET" || new URL(req.url).origin !== self.location.origin) return;
  event.respondWith(
    // cache:"no-cache" revalidates with the server (cheap 304) so updates aren't hidden by the HTTP cache.
    fetch(req, { cache:"no-cache" }).then(function(res){
      if (res && res.ok){
        var copy = res.clone();
        caches.open(CACHE_NAME).then(function(cache){ cache.put(req, copy); });
      }
      return res;
    }).catch(function(){
      return caches.match(req, { ignoreSearch:true }).then(function(hit){
        return hit || caches.match("./index.html");
      });
    })
  );
});
