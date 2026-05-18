/**
 * File-based JSON data store with event hooks.
 *
 * Data is persisted to `data/projects.json` beside the project root.
 * Every mutation fires an event so plugins and SSE can react.
 */

import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { EventEmitter } from 'node:events';

export class Store extends EventEmitter {
  #filePath;
  #data;

  /**
   * @param {string} rootDir  Absolute path to the app root
   * @param {string} [file]   Filename inside `data/` dir
   */
  constructor(rootDir, file = 'projects.json') {
    super();
    const dataDir = join(rootDir, 'data');
    if (!existsSync(dataDir)) mkdirSync(dataDir, { recursive: true });
    this.#filePath = join(dataDir, file);
    this.#data = this.#load();
  }

  /* ── helpers ────────────────────────────────────── */

  #load() {
    try {
      return JSON.parse(readFileSync(this.#filePath, 'utf-8'));
    } catch {
      return [];
    }
  }

  #save() {
    writeFileSync(this.#filePath, JSON.stringify(this.#data, null, 2));
  }

  /* ── public API ─────────────────────────────────── */

  /** Return all projects. */
  getAll() {
    return structuredClone(this.#data);
  }

  /** Return one project by id. */
  getById(id) {
    return structuredClone(this.#data.find((p) => p.id === id));
  }

  /** Replace the full data set (e.g. on import). */
  replaceAll(projects) {
    this.#data = projects;
    this.#save();
    this.emit('replace', { projects });
  }

  /** Add a new project. */
  add(project) {
    this.#data.push(project);
    this.#save();
    this.emit('add', { project });
  }

  /** Update an existing project (matched by id). */
  update(project) {
    const idx = this.#data.findIndex((p) => p.id === project.id);
    if (idx === -1) return false;
    this.#data[idx] = project;
    this.#save();
    this.emit('update', { project });
    return true;
  }

  /** Delete a project by id. */
  delete(id) {
    const before = this.#data.length;
    this.#data = this.#data.filter((p) => p.id !== id);
    if (this.#data.length < before) {
      this.#save();
      this.emit('delete', { id });
      return true;
    }
    return false;
  }
}
