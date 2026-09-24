// 轻量 Service Worker：纯网络模式，永远拉最新资源，不做任何缓存。
// 目的：避免 PWA 缓存旧 app.js 导致「改了不生效」。
// 版本号随 app.js 一起递增；新 SW 安装即激活并接管页面。
const CACHE = 'workbench-v10-networkonly';

self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('activate', (e) => {
  // 清掉历史所有缓存，确保没有旧文件残留
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', (e) => {
  if (e.request.method !== 'GET') return;
  // 页面壳（HTML）必须拿最新的，否则会一直引用旧的 ?v= 版本号
  const isHTML = e.request.mode === 'navigate'
    || (e.request.headers.get('accept') || '').includes('text/html');
  // 全部请求强制绕过 HTTP 缓存；离线才退回首页，保证离线也能打开壳子
  e.respondWith(
    fetch(e.request, { cache: 'no-store' }).catch(() =>
      caches.match('./index.html').then((hit) => hit || new Response('离线不可用', { status: 503 })))
  );
});
