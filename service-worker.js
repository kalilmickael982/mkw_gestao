/* =========================================================
   MKW Gestão — Service Worker
   Faz cache do "app shell" (as páginas, o manifest e a logo) para
   abrir rápido e funcionar offline. NÃO intercepta as chamadas ao
   Apps Script (Google) nem a CDN — essas sempre vão à rede, para
   que os dados das planilhas estejam sempre atualizados.

   Ao mudar qualquer arquivo do app, suba a versão do cache
   (CACHE_NOME) para forçar a atualização nos dispositivos.
   ========================================================= */
const CACHE_NOME = 'mkw-gestao-v1';
const ASSETS = [
  './',
  './index.html',
  './parceiros.html',
  './incluir_parceiro.html',
  './produtos_servicos.html',
  './incluir_produtos_servicos.html',
  './manifest.json',
  './logomarca.png'
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE_NOME)
      .then((c) => c.addAll(ASSETS))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((chaves) => Promise.all(chaves.filter((k) => k !== CACHE_NOME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  // Só gerencia o próprio site (app shell). Apps Script, CDN e Sheets passam direto.
  if (url.origin !== self.location.origin) return;

  const aceitaHtml = (req.headers.get('accept') || '').includes('text/html');

  if (req.mode === 'navigate' || aceitaHtml) {
    // HTML: rede primeiro (pega atualizações), cai para o cache se offline.
    e.respondWith(
      fetch(req)
        .then((resp) => {
          const copia = resp.clone();
          caches.open(CACHE_NOME).then((c) => c.put(req, copia));
          return resp;
        })
        .catch(() => caches.match(req).then((r) => r || caches.match('./index.html')))
    );
  } else {
    // Demais arquivos locais: cache primeiro, busca na rede se não houver.
    e.respondWith(
      caches.match(req).then((r) => r || fetch(req).then((resp) => {
        const copia = resp.clone();
        caches.open(CACHE_NOME).then((c) => c.put(req, copia));
        return resp;
      }))
    );
  }
});
