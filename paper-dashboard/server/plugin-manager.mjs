/**
 * Plugin Manager — discovers and manages lifecycle of plugins.
 *
 * Plugin interface:
 * ─────────────────
 *   export default {
 *     name: 'my-plugin',           // unique identifier
 *     version: '1.0.0',            // semver
 *     description: '...',          // short description
 *
 *     async init(ctx) { ... },     // called on startup
 *     async destroy() { ... },     // called on shutdown (optional)
 *   }
 *
 * Context object (ctx):
 * ─────────────────────
 *   ctx.store     — Store instance (getAll, getById, add, update, delete)
 *   ctx.broadcast — fn(event, data) to push SSE to all browsers
 *   ctx.config    — plugin-specific config from config.json
 *   ctx.log       — fn(msg) scoped logger
 *   ctx.rootDir   — absolute path to app root
 */

import { readdirSync, existsSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

export class PluginManager {
  #plugins = new Map();
  #ctx;

  /**
   * @param {{ store: Store, broadcast: Function, rootDir: string, pluginConfigs: Object }} opts
   */
  constructor({ store, broadcast, rootDir, pluginConfigs = {} }) {
    this.#ctx = { store, broadcast, rootDir, pluginConfigs };
  }

  /** Discover and load all plugins from `plugins/` directory. */
  async loadAll(pluginsDir) {
    if (!existsSync(pluginsDir)) return;

    const entries = readdirSync(pluginsDir);
    for (const entry of entries) {
      const full = join(pluginsDir, entry);
      let modulePath;

      if (entry.endsWith('.mjs') || entry.endsWith('.js')) {
        modulePath = full;
      } else if (statSync(full).isDirectory()) {
        const idx = join(full, 'index.mjs');
        if (existsSync(idx)) modulePath = idx;
        else {
          const idx2 = join(full, 'index.js');
          if (existsSync(idx2)) modulePath = idx2;
        }
      }
      if (entry === 'README.md') continue;
      if (!modulePath) continue;

      try {
        const mod = await import(pathToFileURL(modulePath).href);
        const plugin = mod.default ?? mod;
        if (!plugin.name) {
          console.warn(`  [plugin] skipping ${entry}: no "name" export`);
          continue;
        }
        await this.#initPlugin(plugin);
      } catch (err) {
        console.error(`  [plugin] failed to load ${entry}:`, err.message);
      }
    }
  }

  async #initPlugin(plugin) {
    const config = this.#ctx.pluginConfigs[plugin.name] ?? {};
    const ctx = {
      store: this.#ctx.store,
      broadcast: this.#ctx.broadcast,
      config,
      rootDir: this.#ctx.rootDir,
      log: (msg) => console.log(`  [${plugin.name}] ${msg}`),
    };

    await plugin.init(ctx);
    this.#plugins.set(plugin.name, plugin);
    console.log(`  [plugin] loaded: ${plugin.name} v${plugin.version ?? '?'}`);
  }

  /** Gracefully shut down all plugins. */
  async destroyAll() {
    for (const [name, plugin] of this.#plugins) {
      try {
        if (plugin.destroy) await plugin.destroy();
      } catch (err) {
        console.error(`  [plugin] error destroying ${name}:`, err.message);
      }
    }
    this.#plugins.clear();
  }

  /** List loaded plugin names. */
  list() {
    return [...this.#plugins.entries()].map(([name, p]) => ({
      name,
      version: p.version ?? '?',
      description: p.description ?? '',
    }));
  }
}
