import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { createHash } from 'node:crypto';

const root = path.dirname(fileURLToPath(import.meta.url));
const sources = [
  { id: 'rf', name: 'Официальный интернет-портал правовой информации РФ', url: 'https://publication.pravo.gov.ru/' },
  { id: 'lnr', name: 'Нормативно-правовая база Народного Совета ЛНР', url: 'https://nslnr.su/zakonodatelstvo/normativno-pravovaya-baza/' }
];
const files = { '/': ['index.html', 'text/html; charset=utf-8'], '/app.js': ['app.js', 'text/javascript; charset=utf-8'], '/docx.js': ['docx.js', 'text/javascript; charset=utf-8'], '/style.css': ['style.css', 'text/css; charset=utf-8'], '/form.css': ['form.css', 'text/css; charset=utf-8'] };

export async function checkSources(fetcher = fetch) {
  return Promise.all(sources.map(async (source) => {
    try {
      const response = await fetcher(source.url, { signal: AbortSignal.timeout(12000), headers: { 'User-Agent': 'InstructionStudio/1.0 (legal source change monitor)' } });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const html = (await response.text()).slice(0, 2_000_000);
      // A fingerprint is a change notification, not a legal analysis or an assertion of entry into force.
      const text = html.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, ' ').replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, ' ').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      return { ...source, fingerprint: createHash('sha256').update(text).digest('hex'), checkedAt: new Date().toISOString() };
    } catch (error) {
      return { ...source, error: String(error.message || error), checkedAt: new Date().toISOString() };
    }
  }));
}

export function listenAddress(env = process.env) {
  return { port: Number(env.PORT) || 3210, host: env.HOST || (env.RENDER ? '0.0.0.0' : '127.0.0.1') };
}

export function createServer({ sourceChecker = checkSources, cacheMs = 300_000 } = {}) {
  let cachedCheck;
  let cacheUntil = 0;
  let pendingCheck;
  return http.createServer(async (req, res) => {
    const pathname = new URL(req.url, 'http://localhost').pathname;
    if (req.method !== 'GET') { res.writeHead(405); res.end(); return; }
    if (pathname === '/health') {
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
      res.end(JSON.stringify({ status: 'ok' }));
      return;
    }
    if (pathname === '/api/check') {
      try {
        if (!cachedCheck || Date.now() >= cacheUntil) {
          pendingCheck ||= Promise.resolve().then(sourceChecker).then(result => {
            cachedCheck = result;
            cacheUntil = Date.now() + cacheMs;
            return result;
          }).finally(() => { pendingCheck = null; });
        }
        const result = pendingCheck ? await pendingCheck : cachedCheck;
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' });
        res.end(JSON.stringify(result));
      } catch {
        res.writeHead(503, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(JSON.stringify({ error: 'Источник временно недоступен' }));
      }
      return;
    }
    const entry = files[pathname];
    if (!entry) { res.writeHead(404); res.end('Not found'); return; }
    try {
      const body = await readFile(path.join(root, entry[0]));
      res.writeHead(200, { 'Content-Type': entry[1], 'X-Content-Type-Options': 'nosniff', 'Cache-Control': 'no-cache' });
      res.end(body);
    } catch { res.writeHead(500); res.end('Server error'); }
  });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { host, port } = listenAddress();
  createServer().listen(port, host, () => console.log(`Конструктор инструкций запущен: http://${host}:${port}`));
}
