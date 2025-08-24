//@ts-nocheck
const CACHE_NAME = 'my-pwa-cache-v1'; // Increment this version number when you update your app's assets
const OFFLINE_URL = '/offline.html'; // Path to your custom offline page (recommended for better UX)

// List of essential files that make up your app shell, to be precached
const urlsToCache = [
  '/', // Your root URL
  '/index.html',
  '/styles.css',
  '/app.js',
  '/images/logo.png', // Example: include any crucial images
  OFFLINE_URL
  // Add other critical static assets here (fonts, icons, etc.)
];

// ------------------- INSTALL EVENT -------------------
self.addEventListener('install', (event) => {
  console.log('Service Worker: Installing...');
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('Service Worker: Caching app shell');
        return cache.addAll(urlsToCache);
      })
      .then(() => {
        // Force the waiting service worker to become the active service worker
        // This is often desired during development or for simpler apps
        (self as ServiceWorkerGlobalScope).skipWaiting();
      })
      .catch((error) => {
        console.error('Service Worker: Caching failed during install', error);
      })
  );
});

// ------------------- ACTIVATE EVENT -------------------
self.addEventListener('activate', (event) => {
  console.log('Service Worker: Activating...');
  event.waitUntil(
    // Delete old caches that are no longer needed
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('Service Worker: Deleting old cache', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(() => {
      // Ensure the service worker takes control of clients immediately
      // This is crucial for navigating clients to use the newly activated SW
      (self as ServiceWorkerGlobalScope).clients.claim();
      console.log('Service Worker: Activated and claimed clients');
    })
  );
});

// ------------------- FETCH EVENT -------------------
self.addEventListener('fetch', (event) => {
  // We only want to handle GET requests for caching purposes
  if (event.request.method !== 'GET') {
    return;
  }

  // Strategy: Network-First with Cache Fallback
  // Try to fetch from the network first. If that fails, serve from cache.
  event.respondWith(
    fetch(event.request)
      .then((networkResponse) => {
        // Check if we received a valid response from the network
        if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic') {
          // If not valid, don't cache it, just return the response (or try cache)
          return networkResponse;
        }

        // IMPORTANT: Clone the response. A response is a stream and can only be
        // consumed once. We consume it once to cache it and once the browser
        // consumes it.
        const responseToCache = networkResponse.clone();

        caches.open(CACHE_NAME)
          .then((cache) => {
            cache.put(event.request, responseToCache);
            console.log('Service Worker: Caching new resource', event.request.url);
          })
          .catch((error) => {
            console.warn('Service Worker: Failed to cache network response', event.request.url, error);
          });

        return networkResponse; // Return the live network response
      })
      .catch(() => {
        // Network fetch failed (e.g., offline or network error)
        console.log('Service Worker: Network fetch failed, trying cache for', event.request.url);
        return caches.match(event.request)
          .then((cachedResponse) => {
            if (cachedResponse) {
              console.log('Service Worker: Serving from cache', event.request.url);
              return cachedResponse;
            }

            // If the resource is not in the cache either, provide a fallback
            // Especially for navigation requests, serve the offline page
            if (event.request.mode === 'navigate') {
              console.log('Service Worker: Serving offline page for navigation');
              return caches.match(OFFLINE_URL);
            }

            // For other resource types (images, scripts, etc.), you might
            // return a generic error or a placeholder
            console.warn('Service Worker: Resource not in cache and network failed', event.request.url);
            return new Response('Network error and resource not cached.', {
              status: 503,
              headers: { 'Content-Type': 'text/plain' }
            });
          });
      })
  );
});