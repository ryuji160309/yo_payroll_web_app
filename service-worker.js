const CACHE_VERSION = 'v1.9.5';
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
  '/help/sheets.txt'
];

const OPTIONAL_ASSETS = ['/icons/icon-192.png', '/icons/icon-512.png'];
const NETWORK_ONLY_PATHS = new Set(['/version.json', '/announcements.txt']);
const SPREADSHEET_HOST_SUFFIXES = ['googleusercontent.com'];
const SPREADSHEET_HOSTS = new Set(['docs.google.com']);
const CACHE_FIRST_PATHS = new Set([...STATIC_ASSETS, ...OPTIONAL_ASSETS]);

function uniquePaths(paths) {
  return Array.from(new Set(paths));
}

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

self.addEventListener('message', event => {
  const { data } = event;
  if (!data || data.type !== 'WARMUP_CACHE') {
    return;
  }

  const rawPaths = Array.isArray(data.paths) ? data.paths : [];
  const pathsToWarm = uniquePaths(
    rawPaths.filter(
      path =>
        typeof path === 'string' &&
        path.startsWith('/') &&
        !NETWORK_ONLY_PATHS.has(path)
    )
  );

  if (!pathsToWarm.length) {
    return;
  }

  event.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      await Promise.all(
        pathsToWarm.map(async path => {
          try {
            const request = new Request(path, { cache: 'no-store' });
            const response = await fetch(request);
            if (shouldCacheResponse(response)) {
              await cache.put(request, response.clone());
            }
          } catch (error) {
            // Ignore warm-up failures to avoid breaking message handling.
          }
        })
      );
    })
  );
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
          const copyForRequest = response.clone();
          const copyForPath = response.clone();
          caches.open(CACHE_NAME).then(cache => {
            cache.put(request, copyForRequest);
            cache.put(url.pathname, copyForPath).catch(() => {
              // Ignore errors writing the path-only cache entry.
            });
          });
          return response;
        })
        .catch(async () => {
          const cached = await caches.match(request, { ignoreSearch: true });
          if (cached) {
            return cached;
          }
          const fallback = await caches.match('/index.html');
          if (fallback) {
            return fallback;
          }
          return new Response('Offline', {
            status: 503,
            headers: { 'Content-Type': 'text/plain; charset=utf-8' }
          });
        })
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
