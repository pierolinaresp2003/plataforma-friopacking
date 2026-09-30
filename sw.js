// App instalable (30.09.2026, Piero: "¿cómo se podría volver esto un aplicativo tal cual?").
// Primero la red, para que siempre se vea lo último publicado (index.html y data.json.gz); si no hay
// señal, se abre la última copia guardada en el celular. Solo archivos de este mismo sitio: lo que va a
// GitHub (Guardar y compartir) y a Google no pasa por aquí.
const CACHE = 'friopacking-v3';
self.addEventListener('install', function(e){
  self.skipWaiting();
  // La primera vez se guarda la página misma, para que abra aunque luego no haya señal.
  e.waitUntil(caches.open(CACHE).then(function(c){ return c.addAll(['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png']); }).catch(function(){}));
});
self.addEventListener('activate', function(e){
  e.waitUntil(caches.keys().then(function(ks){ return Promise.all(ks.filter(function(k){ return k !== CACHE; }).map(function(k){ return caches.delete(k); })); }).then(function(){ return self.clients.claim(); }));
});
self.addEventListener('fetch', function(e){
  const req = e.request, url = new URL(req.url);
  if(req.method !== 'GET' || url.origin !== self.location.origin) return;
  // Se guarda sin el "?_=hora" con que se pide data.json.gz: una sola copia por archivo, no una por visita.
  const clave = url.origin + url.pathname;
  e.respondWith(fetch(req).then(function(res){
    if(res && res.ok){ const copia = res.clone(); caches.open(CACHE).then(function(c){ return c.put(clave, copia); }); }
    return res;
  }).catch(function(){
    return caches.match(clave).then(function(r){
      if(r) return r;
      return req.mode === 'navigate' ? caches.match(new URL('./index.html', self.registration.scope).href).then(function(x){ return x || caches.match(self.registration.scope); }) : Response.error();
    });
  }));
});
