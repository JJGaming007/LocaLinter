'use strict';

/**
 * Two small languages the tester writes in the Device Scan panel.
 *
 *   Custom checks — deterministic rules applied to every string on every
 *   screen, alongside the built-in ones in checks.js. They cost nothing, run
 *   before the vision pass, and encode the house style that no generic linter
 *   could know ("buttons stay under 24 characters", "never say ROBINET").
 *
 *   Steps — a short script the crawler replays on the device before it starts
 *   exploring: dismiss the daily popup, sign in, walk to the screen that
 *   actually needs checking. Without it every run starts at the title screen
 *   and spends its budget getting to the interesting part.
 *
 * Both are parsed here rather than in the crawler so a typo is reported to the
 * tester as a line number instead of silently doing nothing.
 */

const SEVERITIES = new Set(['high', 'medium', 'low']);

// ── custom checks ──────────────────────────────────────────────────────────
//
//   forbid:  <regex> [| severity] [| message]
//   maxlen:  <n> [on <regex>] [| severity] [| message]
//   casing:  upper|lower|title [on <regex>] [| severity] [| message]
//   require: <regex> on <regex> [| severity] [| message]
//
// `on <regex>` narrows a rule to elements whose path, id or sheet key matches.
// Blank lines and lines starting with # are ignored.

function splitTail(rest) {
  const parts = String(rest).split('|').map((s) => s.trim());
  const head = parts.shift() || '';
  let severity = null;
  let message = '';
  for (const p of parts) {
    if (!severity && SEVERITIES.has(p.toLowerCase())) severity = p.toLowerCase();
    else if (p) message = message ? `${message} ${p}` : p;
  }
  return { head, severity, message };
}

/** "<something> on <regex>" -> { body, scope } */
function splitScope(head) {
  const m = /^(.*?)\s+on\s+(.+)$/i.exec(head);
  if (!m) return { body: head.trim(), scope: null };
  return { body: m[1].trim(), scope: m[2].trim() };
}

function compileRegex(source, flags = 'i') {
  // Accept both `/foo/i` and a bare `foo`, because testers write both.
  const slashed = /^\/(.*)\/([gimsuy]*)$/.exec(source);
  // g and y carry lastIndex between strings, which would make a rule match
  // every other label. They mean nothing useful here, so drop them.
  if (slashed) return new RegExp(slashed[1], (slashed[2] || flags).replace(/[gy]/g, '') || flags);
  return new RegExp(source, flags);
}

/**
 * @param {string|string[]} input  the rules as typed
 * @returns {{ rules: Array, errors: string[] }}
 */
function compileRules(input) {
  const lines = Array.isArray(input) ? input : String(input || '').split('\n');
  const rules = [];
  const errors = [];

  lines.forEach((raw, idx) => {
    const line = String(raw || '').trim();
    if (!line || line.startsWith('#')) return;
    const n = idx + 1;

    const m = /^(forbid|maxlen|casing|require)\s*:\s*(.+)$/i.exec(line);
    if (!m) {
      errors.push(`Line ${n}: "${line}" does not start with forbid:, maxlen:, casing: or require:.`);
      return;
    }
    const verb = m[1].toLowerCase();
    const { head, severity, message } = splitTail(m[2]);
    const { body, scope } = splitScope(head);

    let scopeRe = null;
    if (scope) {
      try {
        scopeRe = compileRegex(scope);
      } catch (e) {
        errors.push(`Line ${n}: "on ${scope}" is not a valid pattern (${e.message}).`);
        return;
      }
    }

    try {
      if (verb === 'forbid') {
        if (!body) throw new Error('nothing to forbid');
        rules.push({
          verb, line: n, scopeRe, severity: severity || 'high',
          re: compileRegex(body),
          message: message || `Text matches a forbidden pattern (${body}).`,
        });
      } else if (verb === 'maxlen') {
        const max = Number(body);
        if (!Number.isFinite(max) || max <= 0) throw new Error(`"${body}" is not a character count`);
        rules.push({
          verb, line: n, scopeRe, severity: severity || 'medium', max,
          message: message || `Longer than the ${max}-character limit for this element.`,
        });
      } else if (verb === 'casing') {
        const want = body.toLowerCase();
        if (!['upper', 'lower', 'title'].includes(want)) throw new Error(`"${body}" is not upper, lower or title`);
        rules.push({
          verb, line: n, scopeRe, severity: severity || 'low', want,
          message: message || `Should be ${want}case.`,
        });
      } else {
        // require: <regex> on <regex>  — the scope is what makes it meaningful
        if (!scopeRe) throw new Error('require needs "on <pattern>" to say which elements must match');
        rules.push({
          verb, line: n, scopeRe, severity: severity || 'medium',
          re: compileRegex(body),
          message: message || `Does not match the required pattern (${body}).`,
        });
      }
    } catch (e) {
      errors.push(`Line ${n}: ${e.message}.`);
    }
  });

  return { rules, errors };
}

function titleCased(s) {
  return s.replace(/\S+/g, (w) => (/^[\p{Lu}\p{N}\p{P}]+$/u.test(w) ? w : w.charAt(0).toUpperCase() + w.slice(1)));
}

/**
 * Applies compiled rules to one rendered string.
 *
 * @param {Array}  rules   from compileRules()
 * @param {string} text    what the screen shows
 * @param {object} where   { element, key } used by "on <regex>"
 * @returns {Array} plain issues — the caller stamps on element/rect/screen
 */
function applyRules(rules, text, where = {}) {
  const out = [];
  const raw = String(text == null ? '' : text);
  if (!raw.trim() || !rules || !rules.length) return out;
  const subject = `${where.element || ''} ${where.key || ''}`.trim();

  for (const r of rules) {
    if (r.scopeRe && !r.scopeRe.test(subject)) continue;

    if (r.verb === 'forbid') {
      const hit = r.re.exec(raw);
      if (hit) out.push({ type: 'custom_forbidden', severity: r.severity, message: `${r.message} (matched "${hit[0]}")` });
    } else if (r.verb === 'maxlen') {
      if (raw.length > r.max) {
        out.push({
          type: 'custom_too_long',
          severity: r.severity,
          message: `${r.message} — ${raw.length} characters against a ${r.max}-character limit.`,
        });
      }
    } else if (r.verb === 'casing') {
      const bad =
        r.want === 'upper' ? raw !== raw.toLocaleUpperCase()
          : r.want === 'lower' ? raw !== raw.toLocaleLowerCase()
            : raw !== titleCased(raw);
      if (bad) out.push({ type: 'custom_casing', severity: r.severity, message: r.message });
    } else if (r.verb === 'require') {
      if (!r.re.test(raw)) out.push({ type: 'custom_required', severity: r.severity, message: r.message });
    }
  }
  return out;
}

// ── steps ──────────────────────────────────────────────────────────────────
//
//   tap 0.5 0.86         coordinates 0–1 are a fraction of the screen,
//   longpress 0.5 0.5 900  anything larger is device pixels
//   swipe 0.5 0.8 0.5 0.2 400
//   text Hello there
//   key KEYCODE_ENTER
//   wait 1500
//   back / home
//   launch / restart      (needs the Android package)
//   shell am start -n …

const STEP_VERBS = new Set(['tap', 'longpress', 'swipe', 'text', 'key', 'wait', 'back', 'home', 'launch', 'restart', 'shell']);

function num(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

/**
 * @param {string} script
 * @returns {{ steps: Array, errors: string[] }}
 */
function parseSteps(script) {
  const steps = [];
  const errors = [];
  const lines = String(script || '').split('\n');

  lines.forEach((raw, idx) => {
    const line = String(raw || '').trim();
    if (!line || line.startsWith('#')) return;
    const n = idx + 1;
    const [verbRaw, ...rest] = line.split(/\s+/);
    const verb = verbRaw.toLowerCase();
    const arg = line.slice(verbRaw.length).trim();

    if (!STEP_VERBS.has(verb)) {
      errors.push(`Line ${n}: "${verbRaw}" is not a step. Use ${[...STEP_VERBS].join(', ')}.`);
      return;
    }

    if (verb === 'tap' || verb === 'longpress') {
      const x = num(rest[0]);
      const y = num(rest[1]);
      if (x == null || y == null) {
        errors.push(`Line ${n}: ${verb} needs an x and a y (0–1 as a fraction of the screen, or pixels).`);
        return;
      }
      steps.push({ verb, x, y, ms: num(rest[2]) || null, line: n, source: line });
    } else if (verb === 'swipe') {
      const c = rest.slice(0, 4).map(num);
      if (c.some((v) => v == null)) {
        errors.push(`Line ${n}: swipe needs four coordinates (x1 y1 x2 y2) and an optional duration.`);
        return;
      }
      steps.push({ verb, x1: c[0], y1: c[1], x2: c[2], y2: c[3], ms: num(rest[4]) || 400, line: n, source: line });
    } else if (verb === 'wait') {
      const ms = num(rest[0]);
      if (ms == null || ms < 0) {
        errors.push(`Line ${n}: wait needs a number of milliseconds.`);
        return;
      }
      steps.push({ verb, ms: Math.min(ms, 120000), line: n, source: line });
    } else if (verb === 'text' || verb === 'shell') {
      if (!arg) {
        errors.push(`Line ${n}: ${verb} needs something to send.`);
        return;
      }
      steps.push({ verb, arg, line: n, source: line });
    } else if (verb === 'key') {
      if (!arg) {
        errors.push(`Line ${n}: key needs a keycode, e.g. KEYCODE_ENTER or 66.`);
        return;
      }
      steps.push({ verb, arg, line: n, source: line });
    } else {
      steps.push({ verb, line: n, source: line });
    }
  });

  return { steps, errors };
}

module.exports = { compileRules, applyRules, parseSteps };
