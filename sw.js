/**
 * Service worker — Asistencia Grupo Robles
 *
 * Hace dos cosas:
 *  1. Permite que el navegador ofrezca "Instalar aplicación"
 *  2. Guarda la app para que abra aunque no haya internet
 *
 * Lo que NUNCA se guarda en caché son las llamadas al servidor: los marcajes y
 * el padrón siempre se piden en vivo. Si no hay red, la app lo detecta y usa su
 * propia cola, que ya sabe reintentar.
 */

var CACHE = 'asistencia-v1';
var ARCHIVOS = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE)
      .then(function (c) { return c.addAll(ARCHIVOS); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (nombres) {
        return Promise.all(nombres.map(function (n) {
          return n === CACHE ? null : caches.delete(n);
        }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var url = e.request.url;

  // El servidor nunca se cachea: la asistencia tiene que ser siempre la de verdad
  if (url.indexOf('script.google.com') >= 0 || url.indexOf('googleusercontent.com') >= 0) return;
  if (e.request.method !== 'GET') return;

  // Para la app: primero la red, y si falla, lo guardado.
  // Así una versión nueva llega sola en cuanto haya internet.
  e.respondWith(
    fetch(e.request)
      .then(function (r) {
        if (r && r.status === 200 && r.type === 'basic') {
          var copia = r.clone();
          caches.open(CACHE).then(function (c) { c.put(e.request, copia); });
        }
        return r;
      })
      .catch(function () {
        return caches.match(e.request).then(function (c) {
          return c || caches.match('./index.html');
        });
      })
  );
});
