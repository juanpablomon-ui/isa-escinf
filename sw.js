/* Guarda la app en el equipo para que funcione sin internet. Subir VERSION en cada cambio. */
const VERSION = 'isa-1.0.0';
const ARCHIVOS = ['./', 'index.html', 'estilos.css', 'leer_ppt.js', 'plazos.js', 'excel.js', 'app.js', 'lib/xlsx.full.min.js', 'manifest.webmanifest',
  'iconos/icono.svg', 'iconos/icono-180.png', 'iconos/icono-192.png', 'iconos/icono-512.png'];
self.addEventListener('install', e=>{ e.waitUntil(caches.open(VERSION).then(c=>c.addAll(ARCHIVOS))); self.skipWaiting(); });
self.addEventListener('activate', e=>{ e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==VERSION).map(k=>caches.delete(k))))); self.clients.claim(); });
// Primero la red (para recibir cambios); sin conexión, lo guardado.
self.addEventListener('fetch', e=>{
  if(e.request.method!=='GET' || new URL(e.request.url).origin!==location.origin) return;
  const rq = e.request.mode!=='navigate' ? new Request(e.request, {cache:'no-cache'}) : e.request;
  e.respondWith(fetch(rq).then(r=>{ const c = r.clone(); caches.open(VERSION).then(k=>k.put(e.request, c)); return r; })
    .catch(()=>caches.match(e.request, {ignoreSearch:true}).then(r=>r || caches.match('index.html'))));
});
