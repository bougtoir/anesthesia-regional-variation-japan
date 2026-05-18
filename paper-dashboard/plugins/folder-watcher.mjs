/**
 * Folder Watcher Plugin
 *
 * Monitors configured directories and automatically registers/updates
 * project progress based on file system changes.
 *
 * Configuration (in config.json → plugins.folder-watcher):
 * ─────────────────────────────────────────────────────────
 *   {
 *     "watchPaths": ["/path/to/projects"],
 *     "debounceMs": 1000,
 *     "patterns": {
 *       "deliverable": ["*.md", "*.pdf", "docs/**"],
 *       "source":      ["src/**", "*.ts", "*.tsx", "*.js"]
 *     }
 *   }
 *
 * Events emitted to browser via SSE:
 *   - "file-change"   { projectId, path, type: "add"|"change"|"remove" }
 *   - "auto-progress" { projectId, field, oldValue, newValue }
 */

import { watch } from 'node:fs';
import { basename, relative, resolve, sep } from 'node:path';
import { existsSync, statSync, readdirSync } from 'node:fs';

let ctx = null;
let watchers = [];

/** Debounce helper */
function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

/** Try to match a file change to an existing project by folder name. */
function findProjectByPath(watchRoot, filePath) {
  const rel = relative(watchRoot, filePath);
  const topFolder = rel.split(/[\\/]/)[0];
  const projects = ctx.store.getAll();
  return projects.find(
    (p) =>
      p.shortTitle?.toLowerCase() === topFolder.toLowerCase() ||
      p.id === topFolder,
  );
}

/** Count files matching simple glob patterns (shallow check). */
function countFiles(dir) {
  if (!existsSync(dir) || !statSync(dir).isDirectory()) return 0;
  try {
    return readdirSync(dir, { recursive: true }).length;
  } catch {
    return 0;
  }
}

/** Handle a file change event within a watched folder. */
function handleChange(watchRoot, eventType, filename) {
  if (!filename) return;
  const fullPath = resolve(watchRoot, filename);

  const project = findProjectByPath(watchRoot, fullPath);
  if (!project) {
    ctx.log(`unmatched change: ${filename}`);
    return;
  }

  ctx.log(`change detected: ${filename} → project "${project.shortTitle}"`);

  // Broadcast raw file change event
  ctx.broadcast('file-change', {
    projectId: project.id,
    path: filename,
    type: eventType,
  });

  // ── Auto-progress logic (extend as needed) ──────────────
  // Example: count files to auto-update statistics
  const rel = relative(watchRoot, fullPath);
  const topFolder = rel.split(/[\\/]/)[0];
  const projectDir = resolve(watchRoot, topFolder);
  const fileCount = countFiles(projectDir);

  // Update the project's statistics with file count
  const updated = { ...project };
  const oldCount = updated.statistics?.wordCount ?? 0;
  updated.statistics = { ...updated.statistics, wordCount: fileCount };

  if (fileCount !== oldCount) {
    ctx.store.update(updated);
    ctx.broadcast('auto-progress', {
      projectId: project.id,
      field: 'statistics.wordCount',
      oldValue: oldCount,
      newValue: fileCount,
    });
    ctx.log(`auto-updated file count: ${oldCount} → ${fileCount}`);
  }
}

export default {
  name: 'folder-watcher',
  version: '1.0.0',
  description: 'Monitors directories and auto-registers project progress from file changes.',

  async init(_ctx) {
    ctx = _ctx;
    const config = ctx.config ?? {};
    const watchPaths = config.watchPaths ?? [];
    const debounceMs = config.debounceMs ?? 1000;

    if (watchPaths.length === 0) {
      ctx.log('no watchPaths configured — plugin idle. Add paths in config.json → plugins.folder-watcher.watchPaths');
      return;
    }

    for (const watchPath of watchPaths) {
      const resolved = resolve(ctx.rootDir, watchPath);
      if (!existsSync(resolved)) {
        ctx.log(`watch path does not exist: ${resolved}`);
        continue;
      }

      const debouncedHandle = debounce(
        (eventType, filename) => handleChange(resolved, eventType, filename),
        debounceMs,
      );

      try {
        const watcher = watch(resolved, { recursive: true }, debouncedHandle);
        watchers.push(watcher);
        ctx.log(`watching: ${resolved}`);
      } catch (err) {
        ctx.log(`failed to watch ${resolved}: ${err.message}`);
      }
    }
  },

  async destroy() {
    for (const w of watchers) w.close();
    watchers = [];
    ctx?.log('stopped all watchers');
  },
};
