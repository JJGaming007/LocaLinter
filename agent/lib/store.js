'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const RUNS_DIR = path.join(__dirname, '..', 'runs');

/** One scan run: status, live event log, screenshots on disk, collected issues. */
class Run {
  constructor(options) {
    this.id = new Date().toISOString().replace(/[:.]/g, '-') + '-' + crypto.randomBytes(3).toString('hex');
    this.dir = path.join(RUNS_DIR, this.id);
    fs.mkdirSync(path.join(this.dir, 'screens'), { recursive: true });

    this.options = options || {};
    this.status = 'starting';       // starting | running | done | failed | stopped
    this.error = null;
    this.startedAt = Date.now();
    this.finishedAt = null;
    this.stopRequested = false;

    this.screens = [];              // [{ id, file, scene, sig, summary, issueCount, depth, path }]
    this.issues = [];
    this.skipped = [];              // elements we refused to tap
    this.warnings = [];
    this.summary = '';
    this.usage = null;

    this.events = [];
    this.listeners = new Set();
  }

  emit(type, data = {}) {
    const ev = { t: Date.now(), type, ...data };
    this.events.push(ev);
    if (this.events.length > 4000) this.events.splice(0, 1000);
    for (const l of this.listeners) {
      try { l(ev); } catch { /* a dead SSE client must not kill the crawl */ }
    }
  }

  log(message, level = 'info') {
    this.emit('log', { level, message });
  }

  subscribe(fn) {
    this.listeners.add(fn);
    return () => this.listeners.delete(fn);
  }

  saveScreenshot(name, buffer) {
    const file = `${name}.png`;
    fs.writeFileSync(path.join(this.dir, 'screens', file), buffer);
    return file;
  }

  addScreen(screen) {
    this.screens.push(screen);
    this.emit('screen', { screen });
  }

  addIssues(list) {
    if (!list || !list.length) return;
    for (const i of list) {
      i.id = crypto.randomBytes(6).toString('hex');
      this.issues.push(i);
    }
    this.emit('issues', { count: list.length, issues: list });
  }

  finish(status, error) {
    this.status = status;
    this.error = error ? String(error.message || error) : null;
    this.finishedAt = Date.now();
    this.emit('done', { status: this.status, error: this.error, issues: this.issues.length });
    try {
      fs.writeFileSync(path.join(this.dir, 'report.json'), JSON.stringify(this.toJSON(), null, 2), 'utf8');
    } catch (e) {
      console.warn('[run] could not write report.json:', e.message);
    }
  }

  toJSON() {
    return {
      id: this.id,
      status: this.status,
      error: this.error,
      startedAt: this.startedAt,
      finishedAt: this.finishedAt,
      options: { ...this.options, sheet: undefined },
      screens: this.screens,
      issues: this.issues,
      skipped: this.skipped,
      warnings: this.warnings,
      summary: this.summary,
      usage: this.usage,
      counts: this.counts(),
    };
  }

  counts() {
    const c = { high: 0, medium: 0, low: 0, byType: {} };
    for (const i of this.issues) {
      c[i.severity] = (c[i.severity] || 0) + 1;
      c.byType[i.type] = (c.byType[i.type] || 0) + 1;
    }
    return c;
  }
}

const runs = new Map();

function createRun(options) {
  const run = new Run(options);
  runs.set(run.id, run);
  // keep memory bounded across a long-lived agent process
  if (runs.size > 25) {
    const oldest = [...runs.values()].sort((a, b) => a.startedAt - b.startedAt)[0];
    if (oldest && oldest.status !== 'running') runs.delete(oldest.id);
  }
  return run;
}

function getRun(id) {
  return runs.get(id) || null;
}

function listRuns() {
  return [...runs.values()]
    .sort((a, b) => b.startedAt - a.startedAt)
    .map((r) => ({
      id: r.id,
      status: r.status,
      startedAt: r.startedAt,
      finishedAt: r.finishedAt,
      screens: r.screens.length,
      issues: r.issues.length,
      language: r.options.targetHeader,
    }));
}

module.exports = { Run, createRun, getRun, listRuns, RUNS_DIR };
