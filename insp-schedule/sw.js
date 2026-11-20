// 스마트 물길점검 — 앱 설치용(주소창 없이 열기). 항상 최신 화면을 받고, 인터넷이 끊기면 마지막 화면을 보여줌.
const C='mulgil-v1';
self.addEventListener('install',e=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET'||new URL(r.url).origin!==location.origin)return;
 e.respondWith(fetch(r).then(res=>{const k=res.clone();caches.open(C).then(c=>c.put(r,k));return res}).catch(()=>caches.match(r)))});
