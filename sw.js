const CACHE = 'ent-or-v5';
const ASSETS = [
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  'https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap'
];

// ติดตั้ง — cache ไฟล์หลัก
self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(c){ return c.addAll(ASSETS); })
  );
  self.skipWaiting();
});

// Activate — ลบ cache เก่า
self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(
        keys.filter(function(k){ return k !== CACHE; })
            .map(function(k){ return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Fetch — Network first สำหรับ API และหน้าเว็บ, Cache first สำหรับ assets
self.addEventListener('fetch', function(e){
  var url = e.request.url;

  // Apps Script API — ไม่ cache เพราะต้องดึงข้อมูลสดเสมอ
  if(url.includes('script.google.com')){
    e.respondWith(fetch(e.request).catch(function(){
      return new Response(JSON.stringify({status:'offline'}),{
        headers:{'Content-Type':'application/json'}
      });
    }));
    return;
  }

  // หน้าเว็บ (index.html) — Network first เพื่อให้ผู้ใช้ได้เวอร์ชันใหม่เสมอ
  // (ถ้าออฟไลน์ค่อยใช้ตัวที่ cache ไว้)
  if(e.request.mode === 'navigate' || url.includes('index.html')){
    e.respondWith(
      fetch(e.request).then(function(res){
        if(res.ok){
          var clone = res.clone();
          caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        }
        return res;
      }).catch(function(){
        return caches.match(e.request).then(function(cached){
          return cached || caches.match('./index.html');
        });
      })
    );
    return;
  }

  // Assets — Cache first, fallback to network
  e.respondWith(
    caches.match(e.request).then(function(cached){
      return cached || fetch(e.request).then(function(res){
        // cache ไฟล์ใหม่ที่ดึงมา
        if(res.ok){
          var clone = res.clone();
          caches.open(CACHE).then(function(c){ c.put(e.request, clone); });
        }
        return res;
      });
    })
  );
});
