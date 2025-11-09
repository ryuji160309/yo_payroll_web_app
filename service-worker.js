const CACHE_VERSION = 'v1.6.0';
const CACHE_NAME = `yo-payroll-${CACHE_VERSION}`;
const STATIC_ASSETS = [
  '/',
  '/index.html',
  '/payroll.html',
  '/settings.html',
  '/sheets.html',
  '/style.css',
  '/manifest.json',
  '/app.js',
  '/index.js',
  '/payroll.js',
  '/settings.js',
  '/sheets.js',
  '/calc.js',
  '/help.js',
  '/help/top.txt',
  '/help/payroll.txt',
  '/help/settings.txt',
  '/help/sheets.txt',
  '/announcements.txt'
];

const OPTIONAL_ASSETS = ['/icons/icon-192.png', '/icons/icon-512.png'];
const CACHE_FIRST_PATHS = new Set([...STATIC_ASSETS, ...OPTIONAL_ASSETS]);

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      await cache.addAll(STATIC_ASSETS);
      await Promise.all(
        OPTIONAL_ASSETS.map(async asset => {
          try {
            const response = await fetch(asset);
            if (response.ok) {
              await cache.put(asset, response.clone());
            }
          } catch (error) {
            // Ignore optional assets that are not available during install.
          }
        })
      );
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(key => key.startsWith('yo-payroll-') && key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const { request } = event;
  if (request.method !== 'GET') {
    return;
  }

  const url = new URL(request.url);
  if (url.origin !== self.location.origin) {
    return;
  }

  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then(response => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
          return response;
        })
        .catch(() => caches.match(request).then(res => res || caches.match('/index.html')))
    );
    return;
  }

  const cacheFirst = CACHE_FIRST_PATHS.has(url.pathname);

  if (cacheFirst) {
    event.respondWith(
      caches.match(request).then(cached => {
        if (cached) {
          return cached;
        }
        return fetch(request).then(response => {
          if (response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
          }
          return response;
        });
      })
    );
    return;
  }

  event.respondWith(
    fetch(request)
      .then(response => {
        if (response.ok) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
        }
        return response;
      })
      .catch(() => caches.match(request))
  );
});
