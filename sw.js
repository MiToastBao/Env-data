const CACHE_NAME = 'ttrpg-tracker-v3';
const CORE_ASSETS = ['./', './index.html', './manifest.json', './version.json'];
const ALLOWED_CROSS_ORIGIN = ['cdnjs.cloudflare.com'];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS)).catch(()=>{})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) => Promise.all(
      names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n))
    ))
  );
  self.clients.claim();
});

// Network-first for the app itself: always try to fetch the freshest copy so a
// new GitHub Pages deploy is picked up immediately. Falls back to cache when offline.
// For the one allow-listed CDN (used for optional Excel import), use cache-first
// once it's been downloaded, since that library version never changes.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;
  const url = new URL(event.request.url);

  if (url.origin !== self.location.origin) {
    if (!ALLOWED_CROSS_ORIGIN.includes(url.hostname)) return;
    event.respondWith(
      caches.match(event.request).then((cached) => {
        if (cached) return cached;
        return fetch(event.request).then((resp) => {
          if (resp && resp.ok) {
            const copy = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return resp;
        });
      })
    );
    return;
  }

  event.respondWith(
    fetch(event.request, { cache: 'no-store' }).then((resp) => {
      if (resp && resp.ok) {
        const copy = resp.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
      }
      return resp;
    }).catch(() => caches.match(event.request))
  );
});
