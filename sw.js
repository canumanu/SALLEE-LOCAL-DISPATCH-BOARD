// Minimal service worker: lets the app open instantly / install, and falls back to the last copy if offline.
// Always asks the network first so a new upload is picked up right away. Never touches Microsoft or SharePoint calls.
const CACHE = 'local-dispatch-shell-v1';
self.addEventListener('install', e=>{ self.skipWaiting(); });
self.addEventListener('activate', e=>{ e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', e=>{
  const req = e.request, url = new URL(req.url);
  if(req.method !== 'GET' || url.origin !== location.origin) return;   // only our own files
  e.respondWith(
    fetch(req).then(res=>{
      const copy = res.clone();
      caches.open(CACHE).then(c=>c.put(req, copy)).catch(()=>{});
      return res;
    }).catch(()=>caches.match(req).then(r=>r || caches.match('./')))
  );
});
