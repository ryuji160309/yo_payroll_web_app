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
const NETWORK_ONLY_PATHS = new Set(['/version.json']);
const SPREADSHEET_HOST_SUFFIXES = ['googleusercontent.com'];
const SPREADSHEET_HOSTS = new Set(['docs.google.com']);
const CACHE_FIRST_PATHS = new Set([...STATIC_ASSETS, ...OPTIONAL_ASSETS]);

function isSpreadsheetRequest(url) {
  if (SPREADSHEET_HOSTS.has(url.hostname) && url.pathname.includes('/spreadsheets/')) {
    return true;
  }
  return SPREADSHEET_HOST_SUFFIXES.some(suffix => url.hostname.endsWith(suffix));
}

function shouldCacheResponse(response) {
  return response && (response.ok || response.type === 'opaque');
}

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      await cache.addAll(STATIC_ASSETS);
      await Promise.all(
        OPTIONAL_ASSETS.map(async asset => {
          try {
            const response = await fetch(asset);
            if (shouldCacheResponse(response)) {
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
  const sameOrigin = url.origin === self.location.origin;

  if (sameOrigin && request.mode === 'navigate') {
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

  if (sameOrigin && NETWORK_ONLY_PATHS.has(url.pathname)) {
    event.respondWith(fetch(request));
    return;
  }

  if (isSpreadsheetRequest(url)) {
    event.respondWith(fetch(request));
    return;
  }

  const cacheFirst = sameOrigin && CACHE_FIRST_PATHS.has(url.pathname);

  if (cacheFirst) {
    event.respondWith(
      caches.match(request).then(cached => {
        if (cached) {
          return cached;
        }
        return fetch(request).then(response => {
          if (shouldCacheResponse(response)) {
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
        if (shouldCacheResponse(response)) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(request, copy));
        }
        return response;
      })
      .catch(() => caches.match(request))
  );
});
