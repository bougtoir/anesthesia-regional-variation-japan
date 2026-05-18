/**
 * ProjectHub Server
 *
 * Zero-dependency Node.js server that:
 *  1. Serves the built frontend (dist/)
 *  2. Provides a REST API for project CRUD
 *  3. Broadcasts real-time updates via SSE
 *  4. Loads plugins from plugins/ directory
 */

import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { join, extname, resolve, sep } from 'node:path';
import { Store } from './store.mjs';
import { SSEBroadcaster } from './sse.mjs';
import { PluginManager } from './plugin-manager.mjs';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
};

/**
 * Read and parse JSON body from an IncomingMessage.
 * @param {import('node:http').IncomingMessage} req
 */
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString()));
      } catch (e) {
        reject(e);
      }
    });
    req.on('error', reject);
  });
}

function json(res, status, data) {
  res.writeHead(status, { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' });
  res.end(JSON.stringify(data));
}

/**
 * @param {{ rootDir: string, distDir: string, port: number, config: object }} opts
 */
export async function startServer({ rootDir, distDir, port, config }) {
  const store = new Store(rootDir);
  const sse = new SSEBroadcaster();

  // Wire store events → SSE
  for (const ev of ['add', 'update', 'delete', 'replace']) {
    store.on(ev, (data) => sse.broadcast(`project-${ev}`, data));
  }

  // Plugin manager
  const pluginManager = new PluginManager({
    store,
    broadcast: (event, data) => sse.broadcast(event, data),
    rootDir,
    pluginConfigs: config.plugins ?? {},
  });
  await pluginManager.loadAll(join(rootDir, 'plugins'));

  // HTTP server
  const server = createServer(async (req, res) => {
    const method = req.method;
    const url = new URL(req.url, `http://localhost:${port}`);
    const path = url.pathname;

    // CORS preflight
    if (method === 'OPTIONS') {
      res.writeHead(204, {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
      });
      res.end();
      return;
    }

    // ── API routes ──────────────────────────────────

    // SSE stream
    if (path === '/api/events' && method === 'GET') {
      sse.addClient(res);
      return;
    }

    // List plugins
    if (path === '/api/plugins' && method === 'GET') {
      return json(res, 200, pluginManager.list());
    }

    // GET /api/projects
    if (path === '/api/projects' && method === 'GET') {
      return json(res, 200, store.getAll());
    }

    // POST /api/projects
    if (path === '/api/projects' && method === 'POST') {
      try {
        const body = await readBody(req);
        store.add(body);
        return json(res, 201, body);
      } catch {
        return json(res, 400, { error: 'Invalid JSON' });
      }
    }

    // PUT /api/projects/replace (bulk import)
    if (path === '/api/projects/replace' && method === 'PUT') {
      try {
        const body = await readBody(req);
        if (Array.isArray(body)) {
          store.replaceAll(body);
          return json(res, 200, { ok: true });
        }
        return json(res, 400, { error: 'Expected array' });
      } catch {
        return json(res, 400, { error: 'Invalid JSON' });
      }
    }

    // Single project routes: /api/projects/:id
    const projectMatch = path.match(/^\/api\/projects\/([^/]+)$/);
    if (projectMatch) {
      const id = decodeURIComponent(projectMatch[1]);

      if (method === 'GET') {
        const p = store.getById(id);
        return p ? json(res, 200, p) : json(res, 404, { error: 'Not found' });
      }
      if (method === 'PUT') {
        try {
          const body = await readBody(req);
          const ok = store.update({ ...body, id });
          return ok ? json(res, 200, body) : json(res, 404, { error: 'Not found' });
        } catch {
          return json(res, 400, { error: 'Invalid JSON' });
        }
      }
      if (method === 'DELETE') {
        const ok = store.delete(id);
        return ok ? json(res, 200, { ok: true }) : json(res, 404, { error: 'Not found' });
      }
    }

    // ── Static files ────────────────────────────────

    let filePath = path === '/' ? '/index.html' : path;
    // Normalize URL forward slashes to OS path separators
    filePath = filePath.split('/').join(sep);
    const fullPath = resolve(join(distDir, filePath));

    // Prevent directory traversal (use resolve to normalize both paths)
    if (!fullPath.startsWith(resolve(distDir))) {
      res.writeHead(403);
      res.end();
      return;
    }

    try {
      const data = await readFile(fullPath);
      const ext = extname(fullPath);
      res.writeHead(200, { 'Content-Type': MIME[ext] || 'application/octet-stream' });
      res.end(data);
    } catch {
      // SPA fallback
      try {
        const index = await readFile(join(distDir, 'index.html'));
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(index);
      } catch {
        res.writeHead(404);
        res.end('Not Found');
      }
    }
  });

  // Graceful shutdown
  for (const sig of ['SIGINT', 'SIGTERM']) {
    process.on(sig, async () => {
      console.log(`\n  Shutting down...`);
      await pluginManager.destroyAll();
      server.close();
      process.exit(0);
    });
  }

  server.listen(port, () => {
    console.log(`\n  ProjectHub server running at http://localhost:${port}\n`);
    console.log(`  API:     http://localhost:${port}/api/projects`);
    console.log(`  Events:  http://localhost:${port}/api/events (SSE)`);
    console.log(`  Plugins: ${pluginManager.list().map((p) => p.name).join(', ') || '(none active)'}\n`);
    console.log('  Press Ctrl+C to stop.\n');
  });

  return { server, store, sse, pluginManager };
}
