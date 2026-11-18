
'use strict';

/* Service Worker de CapturaDatos
   - Guarda en caché los archivos de la app para que funcione sin internet.
   - Al publicar cambios, sube el número de CACHE_VERSION para forzar la actualización. */

const CACHE_VERSION = 'v1';
const CACHE_NAME = 'capturadatos-' + CACHE_VERSION;

const APP_SHELL = [
  './',
  './index.html',
  './css/styles.css',
  './js/app.js',
  './js/xlsx.js',
  './manifest.json',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/apple-touch-icon.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) =>
        Promise.all(
          keys
            .filter((k) => k.startsWith('capturadatos-') && k !== CACHE_NAME)
            .map((k) => caches.delete(k))
        )
      )
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // Google y otros dominios: conexión directa, sin caché.
  if (url.origin !== self.location.origin) return;

  // Páginas: red primero, caché como alternativa.
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          if (res && res.ok) {
            const copy = res.clone();
            caches.open(CACHE_NAME)
              .then((cache) => cache.put('./index.html', copy));
          }

          return res;
        })
        .catch(async () => {
          const cached = await caches.match('./index.html');

          if (cached) return cached;

          return new Response(
            'CapturaDatos no está disponible. Conéctate a internet para cargar la aplicación.',
            {
              status: 503,
              statusText: 'Service Unavailable',
              headers: { 'Content-Type': 'text/plain; charset=utf-8' }
            }
          );
        })
    );

    return;
  }

  // Archivos: caché primero y red como alternativa.
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;

      return fetch(req).catch(() =>
        new Response('Recurso no disponible sin conexión.', {
          status: 503,
          statusText: 'Service Unavailable'
        })
      );
    })
  );
});
