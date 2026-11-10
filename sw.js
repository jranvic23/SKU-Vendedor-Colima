const CACHE_NAME = 'sku-vendedor-v3';

self.addEventListener('install', event => {
  self.skipWaiting();

  event.waitUntil(
    caches.open(CACHE_NAME).then(cache =>
      cache.addAll([
        './',
        './index.html',
        './manifest.json'
      ])
    )
  );
});

self.addEventListener('activate', event => {

  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    )
  );

  self.clients.claim();

});

self.addEventListener('fetch', event => {

  if (event.request.url.includes('clientes.json')) {

    event.respondWith(

      fetch(event.request)
        .then(response => {

          const copia = response.clone();

          caches.open(CACHE_NAME)
            .then(cache => cache.put(event.request, copia));

          return response;

        })
        .catch(() => caches.match(event.request))

    );

    return;
  }

  event.respondWith(
    caches.match(event.request)
      .then(response => response || fetch(event.request))
  );

});
