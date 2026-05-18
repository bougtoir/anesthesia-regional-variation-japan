#!/usr/bin/env node
/**
 * PaperHub — Launch the server and open the browser.
 *
 * Usage:
 *   node start.mjs                    # default config
 *   node start.mjs --port 8080        # custom port
 *   node start.mjs --no-open          # don't open browser
 */

import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { readFileSync, existsSync } from 'node:fs';
import { exec } from 'node:child_process';
import { startServer } from './server/index.mjs';

const __dirname = fileURLToPath(new URL('.', import.meta.url));

// Load config
let config = {};
const configPath = join(__dirname, 'config.json');
if (existsSync(configPath)) {
  try {
    config = JSON.parse(readFileSync(configPath, 'utf-8'));
  } catch {
    console.warn('  Warning: could not parse config.json, using defaults.');
  }
}

// CLI overrides
const args = process.argv.slice(2);
const portIdx = args.indexOf('--port');
if (portIdx !== -1 && args[portIdx + 1]) config.port = Number(args[portIdx + 1]);
const noOpen = args.includes('--no-open');
if (noOpen) config.openBrowser = false;

const port = config.port ?? 3456;
const distDir = join(__dirname, 'dist');

const { server } = await startServer({ rootDir: __dirname, distDir, port, config });

// Open browser
if (config.openBrowser !== false) {
  const url = `http://localhost:${port}`;
  const cmd =
    process.platform === 'darwin' ? `open "${url}"` :
    process.platform === 'win32' ? `start "" "${url}"` :
    `xdg-open "${url}"`;
  exec(cmd, (err) => { if (err) console.log('  Could not open browser automatically. Open manually:', url); });
}
