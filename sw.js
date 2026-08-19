// Service worker สำหรับ "ระบบติดตามวันหยุดเสาร์"
// แคชแค่ app shell (ไฟล์หน้าตา) — ข้อมูลจริง (getData/updateStatus) ยิงตรงไป
// Apps Script เสมอ ไม่แคช เพื่อไม่ให้เห็นข้อมูลเก่าค้าง
const CACHE_NAME = 'sat-off-tracker-v2';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './favicon-32.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // ไม่แคช request ที่ยิงไปหา Apps Script (ข้อมูลต้องสดเสมอ)
  if (url.hostname.indexOf('script.google') !== -1 || url.hostname.indexOf('googleusercontent') !== -1) {
    return;
  }

  // เฉพาะไฟล์ในโดเมนตัวเอง: cache-first แล้ว fallback ไป network
  if (url.origin === self.location.origin) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        return cached || fetch(event.request).then((res) => {
          const resClone = res.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone));
          return res;
        });
      }).catch(() => caches.match('./index.html'))
    );
  }
});
