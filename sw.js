// =============================================================================
// Service Worker — vela-v3
// Offline-first PWA strategy with per-route caching policies
// =============================================================================

const CACHE = 'vela-v3';

// Files to pre-cache during install
const FILES = [
  '/',
  './index.html',
  './feltai_2.html',
  './landing.html',
  './manifest.json',
];

// API route prefixes that must never be cached
const API_PREFIXES = [
  '/bw-api',
  '/fiskdir-api',
  '/claude',
  '/transcribe',
  '/login',
  '/register',
];

// HTML files that use cache-first with background revalidation
const CACHE_FIRST_HTML = [
  'feltai_2.html',
  'landing.html',
];

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Returns true if the request URL path starts with any API prefix */
function isApiRequest(url) {
  const { pathname } = new URL(url);
  return API_PREFIXES.some(prefix => pathname.startsWith(prefix));
}

/** Returns true if the request URL is for one of the cache-first HTML files */
function isCacheFirstHtml(url) {
  const { pathname } = new URL(url);
  return CACHE_FIRST_HTML.some(file => pathname.endsWith(file));
}

/** Returns true if the request is same-origin */
function isSameOrigin(url) {
  return new URL(url).origin === self.location.origin;
}

/**
 * Notify all connected clients that a new service worker version has activated.
 * Clients can listen for this to show a "New version available" toast.
 */
async function notifyClients(type, payload = {}) {
  const allClients = await self.clients.matchAll({ includeUncontrolled: true });
  for (const client of allClients) {
    client.postMessage({ type, ...payload });
  }
}

// ---------------------------------------------------------------------------
// Install — pre-cache static shell
// ---------------------------------------------------------------------------

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE).then(cache => cache.addAll(FILES))
  );
  // Take control immediately; don't wait for old SW to become idle
  self.skipWaiting();
});

// ---------------------------------------------------------------------------
// Activate — prune stale caches and claim all clients
// ---------------------------------------------------------------------------

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys =>
        Promise.all(
          keys
            .filter(key => key !== CACHE)
            .map(key => caches.delete(key))
        )
      )
      .then(() => {
        // Tell every open tab it is now controlled by this SW
        self.clients.claim();
        // Inform clients so they can show a "ready to work offline" or
        // "new version installed" toast
        notifyClients('SW_ACTIVATED', { version: CACHE });
      })
  );
});

// ---------------------------------------------------------------------------
// Fetch — per-route strategy
// ---------------------------------------------------------------------------

self.addEventListener('fetch', (event) => {
  const { request } = event;
  const url = request.url;

  // Only handle GET requests; let others (POST, etc.) pass through
  if (request.method !== 'GET') return;

  // ── 1. API routes — network-only, never cache ──────────────────────────
  if (isSameOrigin(url) && isApiRequest(url)) {
    // Fall through; do not call respondWith so the browser handles it normally
    return;
  }

  // ── 2. Same-origin HTML — cache-first with network fallback + cache update
  if (isSameOrigin(url) && isCacheFirstHtml(url)) {
    event.respondWith(cacheFirstWithUpdate(request));
    return;
  }

  // ── 3. Everything else — network-first with cache fallback ────────────
  event.respondWith(networkFirstWithCacheFallback(request));
});

// ---------------------------------------------------------------------------
// Fetch strategies
// ---------------------------------------------------------------------------

/**
 * Cache-first: serve from cache immediately if available, then fetch from
 * network in the background and refresh the cache entry for next time.
 */
async function cacheFirstWithUpdate(request) {
  const cache = await caches.open(CACHE);
  const cached = await cache.match(request);

  // Kick off a background network fetch to keep the cache fresh
  const networkFetch = fetch(request)
    .then(response => {
      if (response && response.ok) {
        cache.put(request, response.clone());
      }
      return response;
    })
    .catch(() => null); // Silently ignore network errors during background update

  // Return the cached version instantly, or wait for network if nothing cached
  return cached || networkFetch;
}

/**
 * Network-first: try the network, fall back to cache when offline.
 * Successful network responses are stored in the cache for future fallback.
 */
async function networkFirstWithCacheFallback(request) {
  const cache = await caches.open(CACHE);

  try {
    const response = await fetch(request);
    // Only cache valid responses (not opaque/error responses for cross-origin)
    if (response && response.ok) {
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    // Network unavailable — serve from cache
    const cached = await cache.match(request);
    return cached || new Response('Offline', {
      status: 503,
      statusText: 'Service Unavailable',
      headers: { 'Content-Type': 'text/plain' },
    });
  }
}

// ---------------------------------------------------------------------------
// Message handler — allow controlled SW updates from the page
// ---------------------------------------------------------------------------

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    // Page explicitly requested the waiting SW to take over.
    // Typically triggered after the user acknowledges an "Update available" toast.
    self.skipWaiting();
  }
});
