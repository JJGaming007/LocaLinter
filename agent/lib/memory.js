'use strict';

/**
 * What the agent knows about an app, kept between runs.
 *
 * Every scan used to start from nothing: it rediscovered the launch ad, walked
 * into the same dead ends, and re-reported the same false positives a human had
 * already waved off. None of that is knowledge about *this* run — it is
 * knowledge about the app, and it should outlive the run that learned it.
 *
 * One file per package under memory/. Three kinds of thing live in it:
 *
 *   obstacles  what gets in the way on launch, and what cleared it last time
 *   dismissed  findings a human said were not real, so we stop repeating them
 *   notes      a short description of the app, written by Claude after a run,
 *              handed back to the vision pass on the next one as context
 *
 * Deliberately conservative: memory only ever suppresses a finding a person
 * explicitly dismissed, and only for the same string and issue type. A store
 * that quietly learned to hide real bugs would be worse than no store at all.
 */

const fs = require('fs');
const path = require('path');

const { DATA_DIR } = require('./paths');
const { norm } = require('./sheet');

const MEM_DIR = path.join(DATA_DIR, 'memory');

const EMPTY = {
  version: 1,
  package: '',
  runs: 0,
  updatedAt: null,
  notes: '',            // Claude's description of the app
  obstacles: [],        // [{ kind, hint, dismissal, seen, lastSeen }]
  dismissed: [],        // [{ type, text, key, reason, at }]
  screens: {},          // name -> { summary, seen }
};

function fileFor(pkg) {
  const safe = String(pkg || 'unknown').replace(/[^a-zA-Z0-9._-]/g, '_');
  return path.join(MEM_DIR, `${safe}.json`);
}

function load(pkg) {
  if (!pkg) return { ...EMPTY };
  try {
    const raw = JSON.parse(fs.readFileSync(fileFor(pkg), 'utf8'));
    return { ...EMPTY, ...raw, package: pkg };
  } catch {
    return { ...EMPTY, package: pkg };
  }
}

function save(mem) {
  if (!mem || !mem.package) return mem;
  try {
    fs.mkdirSync(MEM_DIR, { recursive: true });
    mem.updatedAt = new Date().toISOString();
    fs.writeFileSync(fileFor(mem.package), JSON.stringify(mem, null, 2), 'utf8');
  } catch (e) {
    console.warn('[memory] could not save:', e.message);
  }
  return mem;
}

/* ── obstacles ───────────────────────────────────────────────────────────── */

/**
 * Record that something blocked the way and how it was cleared, so the next
 * run reaches for that first instead of restarting the app three times.
 */
function rememberObstacle(mem, { kind, hint = '', dismissal }) {
  if (!mem || !kind || !dismissal) return mem;
  const found = mem.obstacles.find((o) => o.kind === kind && o.dismissal === dismissal);
  if (found) {
    found.seen += 1;
    found.lastSeen = new Date().toISOString();
    if (hint) found.hint = hint;
  } else {
    mem.obstacles.push({ kind, hint, dismissal, seen: 1, lastSeen: new Date().toISOString() });
  }
  return mem;
}

/** The dismissal that has worked most often for this kind of obstacle. */
function bestDismissal(mem, kind) {
  const candidates = (mem.obstacles || []).filter((o) => o.kind === kind);
  if (!candidates.length) return null;
  return candidates.sort((a, b) => b.seen - a.seen)[0].dismissal;
}

/* ── dismissed findings ──────────────────────────────────────────────────── */

const fingerprint = (issue) => `${issue.type}::${norm(issue.text || '')}`;

function dismiss(mem, issue, reason = '') {
  if (!mem || !issue || !issue.type) return mem;
  const fp = fingerprint(issue);
  if (mem.dismissed.some((d) => `${d.type}::${norm(d.text || '')}` === fp)) return mem;
  mem.dismissed.push({
    type: issue.type,
    text: issue.text || '',
    key: issue.key || '',
    reason,
    at: new Date().toISOString(),
  });
  return mem;
}

function undismiss(mem, issue) {
  if (!mem || !issue) return mem;
  const fp = fingerprint(issue);
  mem.dismissed = mem.dismissed.filter((d) => `${d.type}::${norm(d.text || '')}` !== fp);
  return mem;
}

/**
 * Drop findings a person has already judged not to be real.
 *
 * Matched on issue type *and* exact string: a dismissal of "PrimeRush is not
 * Portuguese" must never silence a genuine truncation of the same word.
 */
function filterDismissed(mem, issues) {
  if (!mem || !mem.dismissed.length) return { kept: issues, removed: 0 };
  const seen = new Set(mem.dismissed.map((d) => `${d.type}::${norm(d.text || '')}`));
  const kept = issues.filter((i) => !seen.has(fingerprint(i)));
  return { kept, removed: issues.length - kept.length };
}

/* ── prompt context ──────────────────────────────────────────────────────── */

/**
 * What to tell Claude about this app before it looks at a screen. Kept short —
 * it rides in the cached system block on every screen of every run, so length
 * here is paid for many times over.
 */
function promptContext(mem) {
  if (!mem) return '';
  const parts = [];
  if (mem.notes) parts.push(`What previous scans learned about this app:\n${mem.notes}`);

  const dismissed = (mem.dismissed || []).slice(-25);
  if (dismissed.length) {
    parts.push(
      'A human reviewed earlier scans and said these are NOT defects. Do not report them again ' +
      'unless something else about them is genuinely wrong:\n' +
      dismissed.map((d) => `- ${d.type}: "${d.text}"${d.reason ? ` (${d.reason})` : ''}`).join('\n')
    );
  }

  const obstacles = (mem.obstacles || []).filter((o) => o.seen > 1);
  if (obstacles.length) {
    parts.push(
      'Known interruptions in this app:\n' +
      obstacles.map((o) => `- ${o.kind}${o.hint ? ` (${o.hint})` : ''} — cleared by: ${o.dismissal}`).join('\n')
    );
  }
  return parts.join('\n\n');
}

module.exports = {
  load, save, MEM_DIR,
  rememberObstacle, bestDismissal,
  dismiss, undismiss, filterDismissed, fingerprint,
  promptContext,
};
