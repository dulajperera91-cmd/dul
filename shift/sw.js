// Opens even with no internet; always tries the newest version first.
const CACHE='dul-shift-v7';
const FILES=['./','./index.html','./config.js','./manifest.json','./icon-192.png','./icon-512.png'];
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(FILES.map(u=>new Request(u,{cache:'reload'})))).then(()=>self.skipWaiting()));});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k.startsWith('dul-shift')&&k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener('fetch',e=>{const u=new URL(e.request.url);
  if(e.request.method!=='GET'||u.origin!==location.origin)return;
  e.respondWith(fetch(new Request(e.request.url,{cache:'no-cache',credentials:'same-origin'})).then(r=>{const c=r.clone();caches.open(CACHE).then(x=>x.put(e.request,c));return r;})
    .catch(()=>caches.match(e.request,{ignoreSearch:true}).then(r=>r||caches.match('./index.html'))));});
