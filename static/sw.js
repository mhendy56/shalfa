const CACHE_NAME='shalfa-ops-v14-6';
const STATIC_CACHE=[
  '/static/style.css',
  '/static/icon-192.png',
  '/static/icon-512.png',
  '/static/apple-touch-icon.png',
  '/static/shalfa-logo.png',
  '/static/tbc-logo.png',
  '/static/offline.html',
  '/static/pwa.js'
];

self.addEventListener('install',event=>{
  event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(STATIC_CACHE)));
  self.skipWaiting();
});

self.addEventListener('activate',event=>{
  event.waitUntil(
    caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE_NAME).map(k=>caches.delete(k))))
  );
  self.clients.claim();
});

self.addEventListener('fetch',event=>{
  const req=event.request;
  if(req.method!=='GET') return;

  const url=new URL(req.url);
  if(url.origin!==location.origin) return;

  // Never cache APIs, login state, exports or PDFs.
  if(
    url.pathname.startsWith('/api/') ||
    url.pathname.startsWith('/login') ||
    url.pathname.startsWith('/logout') ||
    url.pathname.endsWith('.xlsx') ||
    url.pathname.endsWith('.pdf')
  ){
    event.respondWith(fetch(req));
    return;
  }

  if(req.mode==='navigate'){
    event.respondWith(
      fetch(req).catch(()=>caches.match('/static/offline.html'))
    );
    return;
  }

  event.respondWith(
    caches.match(req).then(cached=>{
      const network=fetch(req).then(response=>{
        if(response && response.ok){
          const copy=response.clone();
          caches.open(CACHE_NAME).then(cache=>cache.put(req,copy));
        }
        return response;
      }).catch(()=>cached);
      return cached || network;
    })
  );
});

self.addEventListener('message',event=>{
  if(event.data && event.data.type==='SKIP_WAITING') self.skipWaiting();
});
