// Service Worker ขั้นต่ำ — มีไว้เพื่อให้ Chrome รู้จำว่าเว็บนี้เป็น PWA ติดตั้งได้จริง
// (ทำให้เปิดจากหน้าจอโฮมแล้วไม่มีแถบ URL) พร้อม cache หน้าเว็บไว้เผื่อเน็ตหลุด
const CACHE_NAME = "sat-off-v1";
const APP_SHELL = ["./", "./index.html", "./manifest.json"];

self.addEventListener("install", (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL).catch(() => {}))
  );
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  // เรียก API ของ Apps Script ปล่อยผ่านตรงๆ เสมอ ไม่ cache (ต้องได้ข้อมูลสดใหม่)
  if (event.request.url.indexOf("script.google.com") !== -1) return;

  event.respondWith(
    fetch(event.request)
      .then((res) => {
        var resClone = res.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, resClone));
        return res;
      })
      .catch(() => caches.match(event.request))
  );
});
