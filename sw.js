// Matteo Orders — minimal service worker
//
// Purpose: satisfy PWA "installability" requirements on Android/Chrome and give
// a tiny bit of offline resilience for the app shell. It deliberately does NOT
// cache API calls (Supabase) — orders/products must always be fetched fresh.
//
// Strategy: network-first for navigation/app files, falling back to cache only
// when the network is unavailable. Cache is versioned so pushing a new deploy
// automatically invalidates old assets.

const CACHE_NAME = 'matteo-orders-shell-v1';
const APP_SHELL = [
  '/',
  '/index.html',
  '/manifest.json',
  '/icons/icon-192.png',
  '/icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;

  // Never intercept API calls (Supabase, Netlify functions) — always go to network.
  if (request.url.includes('/.netlify/functions/') || request.url.includes('supabase.co')) {
    return;
  }
  // Only handle GET requests.
  if (request.method !== 'GET') return;

  event.respondWith(
    fetch(request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy)).catch(() => {});
        return response;
      })
      .catch(() => caches.match(request).then((cached) => cached || caches.match('/index.html')))
  );
});
