'use strict';

/**
 * LocaLinter Device Scan — local agent.
 *
 * Binds to 127.0.0.1 only. The LocaLinter web app (localhost or the deployed
 * site) calls it from the browser; nothing here is reachable from the network.
 *
 *   npm install && npm start
 */

const http = require('http');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

const config = require('./lib/config');
const { Adb } = require('./lib/adb');
const { Bridge } = require('./lib/bridge');
const { SheetIndex, RTL_CODES, norm } = require('./lib/sheet');
const { ClaudeAnalyzer } = require('./lib/claude');
const { Crawler } = require('./lib/crawler');
const store = require('./lib/store');
const routeMaps = require('./lib/routes');
const { compileRules, parseSteps } = require('./lib/rules');
const paths = require('./lib/paths');
const tools = require('./lib/tools');
const autostart = require('./lib/autostart');
const google = require('./lib/google');
const memory = require('./lib/memory');

const PORT = Number(process.env.PORT || 8790);
const HOST = '127.0.0.1';
const VERSION = '1.0.0';

// ── helpers ───────────────────────────────────────────────────────────────

/**
 * Deliberately no Access-Control-Allow-Origin.
 *
 * The web build served the UI from a different origin, so this had to answer
 * `*` — and, once Chrome's private-network rules landed, Allow-Private-Network
 * too. The app now serves its own UI, so every legitimate call is same-origin
 * and needs no CORS header at all.
 *
 * Restoring those headers would be a real hole rather than a convenience:
 * /api/auth/token hands back a live Google access token, so `*` would let any
 * page you happened to visit read it straight off loopback.
 */
function cors(res) {
  res.setHeader('Vary', 'Origin');
}

function json(res, status, body) {
  const data = JSON.stringify(body);
  cors(res);
  res.writeHead(status, { 'content-type': 'application/json; charset=utf-8', 'content-length': Buffer.byteLength(data) });
  res.end(data);
}

function readBody(req, limit = 64 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let size = 0;
    req.on('data', (c) => {
      size += c.length;
      if (size > limit) {
        reject(new Error('request body too large'));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8');
      if (!raw.trim()) return resolve({});
      try { resolve(JSON.parse(raw)); } catch (e) { reject(new Error('invalid JSON body')); }
    });
    req.on('error', reject);
  });
}

// ── route handlers ────────────────────────────────────────────────────────

// adb may be the tester's own install, a copy the agent downloaded, or missing
// entirely. Resolving it in one place means every route below gets a binary
// that actually runs, and falls back to the configured value so adb.js can
// report "not found" in its own words.
async function adbFor(cfg) {
  const found = await tools.resolveAdb(cfg);
  return found ? found.path : (cfg.adbPath || 'adb');
}

const routes = {
  'GET /api/health': async () => ({
    ok: true,
    version: VERSION,
    node: process.version,
    platform: process.platform,
  }),

  'GET /api/config': async () => ({ config: config.redact(config.load()) }),

  // Route maps: what a previous pass learned about a game — its screens, the
  // coordinates of the info badges that open flyouts, how to switch language,
  // how to recover when it strands itself. Served so the browser can offer
  // them and a scan can start knowing where things are.
  'GET /api/routes': async () => ({ routes: routeMaps.list() }),

  'POST /api/config': async (req) => {
    const body = await readBody(req);
    const allowed = Object.keys(config.DEFAULTS);
    const patch = {};
    for (const k of allowed) {
      if (Object.prototype.hasOwnProperty.call(body, k)) patch[k] = body[k];
    }
    // an empty apiKey means "leave it alone", not "erase it"
    if (patch.apiKey === '') delete patch.apiKey;
    if (body.clearApiKey === true) patch.apiKey = '';
    const next = config.save(patch);
    return { config: config.redact(next), savedTo: config.CONFIG_PATH };
  },

  // Lints the custom checks and setup steps as they are typed, so a mistake
  // surfaces in the panel rather than halfway through a twenty-minute scan.
  'POST /api/validate': async (req) => {
    const body = await readBody(req, 256 * 1024);
    const rules = compileRules(body.customRules || '');
    const steps = parseSteps(body.preSteps || '');
    return {
      rules: { count: rules.rules.length, errors: rules.errors },
      steps: { count: steps.steps.length, errors: steps.errors },
    };
  },

  'GET /api/devices': async () => {
    const cfg = config.load();
    const bin = await adbFor(cfg);
    try {
      const devices = await Adb.devices(bin);
      return { devices, adbPath: bin };
    } catch (e) {
      return { devices: [], error: e.message, adbPath: bin };
    }
  },

  // Whether adb is usable, and where it came from. The browser asks so it can
  // offer the download instead of leaving a tester staring at "No devices".
  'GET /api/tools/adb': async () => {
    const found = await tools.resolveAdb(config.load());
    return found ? { found: true, ...found } : { found: false, downloadFrom: tools.ZIP_URL };
  },

  // ── Google sign-in (installed-app flow) ──
  // The UI never sees the refresh token; it asks for an access token and gets
  // one that is already valid, renewed here when needed.
  'GET /api/auth/session': async () => ({
    configured: google.configured(),
    session: google.session(),
    flow: google.flow(),
  }),

  'POST /api/auth/start': async (req) => {
    const body = await readBody(req);
    google.begin({ port: PORT, hint: body.hint || '' });
    return { ok: true };
  },

  'GET /api/auth/token': async () => {
    const auth = await google.token();
    if (!auth) {
      const err = new Error('Not signed in.');
      err.status = 401;
      throw err;
    }
    return { access_token: auth.access_token, expires_at: auth.expires_at, user: auth.user };
  },

  'POST /api/auth/signout': async () => {
    await google.signOut();
    return { ok: true };
  },

  // Start-at-login. The browser cannot launch this process, so the next best
  // thing is letting the tester ask the OS to do it for them, once.
  'GET /api/autostart': async () => autostart.status(),

  'POST /api/autostart': async (req) => {
    const body = await readBody(req);
    return body.enable === false ? autostart.disable() : autostart.enable();
  },

  // Explicitly triggered from the browser: fetch Google's platform-tools and
  // remember where it landed, so later runs need no network.
  'POST /api/tools/adb': async () => {
    const log = [];
    const adbPath = await tools.downloadAdb((m) => log.push(m));
    const next = config.save({ adbPath });
    return { adbPath, log, config: config.redact(next) };
  },

  'POST /api/probe': async (req) => {
    const cfg = config.load();
    const body = await readBody(req);
    const mode = body.mode === 'editor' ? 'editor' : 'device';
    const port = Number(body.bridgePort || cfg.bridgePort);
    const result = { mode, bridgePort: port, bridge: null, device: null, forwarded: false, errors: [] };

    if (mode === 'device') {
      const serial = body.serial || null;
      const adb = new Adb(await adbFor(cfg), serial);
      try {
        const size = await adb.screenSize();
        const activity = await adb.currentActivity();
        const locale = await adb.deviceLocale();
        result.device = { serial, size, ...activity, locale };
      } catch (e) {
        result.errors.push(`Device: ${e.message}`);
      }
      try {
        await adb.forward(port);
        result.forwarded = true;
      } catch (e) {
        result.errors.push(`Port forward: ${e.message}`);
      }
    }

    const bridge = new Bridge(port);
    const info = await bridge.connect();
    result.bridge = info;
    // Device mode renders its own "no bridge" guidance, so only the editor —
    // where the bridge is mandatory — needs an error here.
    if (!info && mode === 'editor') {
      result.errors.push(`No bridge on 127.0.0.1:${port}. Add LocaLinterBridge.cs to the project and enter Play Mode.`);
    }
    return result;
  },

  'POST /api/run/start': async (req) => {
    const cfg = config.load();
    const body = await readBody(req);

    if (!cfg.apiKey) {
      const err = new Error('No Anthropic API key. Set it in the Device Scan settings first.');
      err.status = 400;
      throw err;
    }
    // Fail here rather than after a minute of crawling that ends in a 401.
    const keyProblem = config.keyLooksWrong(cfg.apiKey, cfg.baseUrl);
    if (keyProblem) {
      const err = new Error(`${keyProblem} Get one from console.anthropic.com and save it again.`);
      err.status = 400;
      throw err;
    }
    if (!body.sheet || !Array.isArray(body.sheet.headers) || !Array.isArray(body.sheet.rows)) {
      const err = new Error('No localization sheet was sent. Load a sheet in LocaLinter first.');
      err.status = 400;
      throw err;
    }

    const sheet = new SheetIndex(body.sheet.headers, body.sheet.rows);
    if (!sheet.entries.length) {
      const err = new Error('The localization sheet has no usable rows.');
      err.status = 400;
      throw err;
    }

    // A language was asked for explicitly, so an unresolvable one is an error,
    // not an invitation to guess. Falling back silently meant a scan could
    // check a completely different column from the one selected — and then
    // report every string on screen as belonging to the wrong language.
    let target;
    if (body.targetLanguage) {
      target = sheet.languageByCodeOrHeader(body.targetLanguage);
      if (!target) {
        const err = new Error(
          `The sheet has no column matching "${body.targetLanguage}". ` +
          `It has: ${sheet.languages.map((l) => l.header).join(', ') || 'no language columns'}.`
        );
        err.status = 400;
        throw err;
      }
    } else {
      target = sheet.languages.find((l) => l.code !== 'en') || sheet.languages[0];
    }
    if (!target) {
      const err = new Error('The sheet has no language columns.');
      err.status = 400;
      throw err;
    }

    const mode = body.mode === 'editor' ? 'editor' : 'device';
    const runCfg = { ...cfg, ...pickOverrides(body.options) };
    const route = body.route ? routeMaps.get(body.route) : null;
    if (body.route && !route) {
      const err = new Error(`No route map named "${body.route}".`);
      err.status = 400;
      throw err;
    }

    const run = store.createRun({
      mode,
      serial: body.serial || null,
      targetHeader: target.header,
      targetCode: target.code,
      sheetSummary: sheet.summary(),
      limits: {
        maxScreens: runCfg.maxScreens,
        maxActions: runCfg.maxActions,
        maxDepth: runCfg.maxDepth,
      },
    });

    // Fire and forget: the browser follows along over SSE.
    startScan({ run, cfg: runCfg, sheet, target, mode, serial: body.serial || null, route }).catch((e) => {
      run.log(`Scan failed: ${e.message}`, 'error');
      run.finish('failed', e);
    });

    return { runId: run.id, language: target.header, mode };
  },

  'GET /api/runs': async () => ({ runs: store.listRuns() }),

  // Past runs and their screenshots pile up with no way to be rid of them.
  'POST /api/runs/clear': async () => store.clearRuns(),

  /**
   * Explain what the scan would decide about one string, and why.
   *
   * The judgement is several rules deep and most of them end in silence, so
   * "why did it flag this / why did it not" was unanswerable without reading
   * the source. This answers it against the sheet actually loaded.
   */
  'POST /api/sheet/explain': async (req) => {
    const body = await readBody(req);
    if (!body.sheet || !Array.isArray(body.sheet.headers)) {
      throw Object.assign(new Error('Load a sheet first.'), { status: 400 });
    }
    const sheet = new SheetIndex(body.sheet.headers, body.sheet.rows || []);
    const target = body.targetLanguage
      ? sheet.languageByCodeOrHeader(body.targetLanguage)
      : sheet.languages.find((l) => l.code !== 'en');
    if (!target) throw Object.assign(new Error('No such language column.'), { status: 400 });

    const text = String(body.text || '').trim();
    const sourceHeader = sheet.englishCol ? sheet.englishCol.header : null;
    const hits = sheet.lookupExact(text).filter((h) => h.header !== '__key__');
    const rows = hits.slice(0, 5).map((h) => ({
      key: h.entry.key,
      row: h.entry.rowNumber,
      matchedColumn: h.header,
      source: sourceHeader ? h.entry.values[sourceHeader] : '',
      target: h.entry.values[target.header] || '',
    }));

    let verdict = 'reported';
    let reason;
    if (!text) {
      verdict = 'ignored'; reason = 'Empty string.';
    } else if (hits.some((h) => h.header === target.header)) {
      verdict = 'ignored';
      reason = `This is exactly what the "${target.header}" column contains, so it is correct.`;
    } else if (sourceHeader && hits.some((h) => h.header === sourceHeader)) {
      const hit = hits.find((h) => h.header === sourceHeader);
      const expected = hit.entry.values[target.header] || '';
      if (!expected.trim()) {
        verdict = 'noted';
        reason = `Found in "${sourceHeader}" (key ${hit.entry.key}), but "${target.header}" is empty — the build has nothing else to show. A sheet gap, not a build defect.`;
      } else if (norm(expected) === norm(hit.entry.values[sourceHeader] || '')) {
        verdict = 'ignored';
        reason = `Found in "${sourceHeader}" (key ${hit.entry.key}), and "${target.header}" holds the same text — the build matches the sheet.`;
      } else {
        reason = `Found in "${sourceHeader}" (key ${hit.entry.key}) while "${target.header}" says "${expected}" — the build is not using the translation.`;
      }
    } else {
      const fuzzy = (sheet.lookupFuzzy(text, { limit: 3 }) || []).map((f) => ({
        key: f.entry.key, row: f.entry.rowNumber, matchedColumn: f.header,
        score: Math.round(f.score * 100), value: f.entry.values[f.header],
      }));
      reason = fuzzy.length
        ? 'No exact match anywhere in the sheet. Closest entries are listed — if none is right, the string was probably hardcoded.'
        : 'Not in the sheet in any language — probably hardcoded and never sent for translation.';
      return { text, targetHeader: target.header, verdict, reason, rows, near: fuzzy };
    }
    return { text, targetHeader: target.header, verdict, reason, rows };
  },

  // ── what the agent remembers about an app ──
  'GET /api/memory': async (req, res, url) => {
    const pkg = url.searchParams.get('package') || config.load().androidPackage;
    return { memory: memory.load(pkg) };
  },

  // A human saying "that is not a defect" is the most valuable signal there is,
  // so it is stored against the app and respected on every later run.
  'POST /api/memory/dismiss': async (req) => {
    const body = await readBody(req);
    const pkg = body.package || config.load().androidPackage;
    if (!pkg) throw Object.assign(new Error('Set an Android package first — memory is stored per app.'), { status: 400 });
    const mem = memory.load(pkg);
    if (body.undo) memory.undismiss(mem, body.issue || {});
    else memory.dismiss(mem, body.issue || {}, body.reason || '');
    memory.save(mem);
    return { memory: mem };
  },

  'POST /api/memory/clear': async (req) => {
    const body = await readBody(req);
    const pkg = body.package || config.load().androidPackage;
    const mem = { ...memory.load(pkg), runs: 0, notes: '', obstacles: [], dismissed: [], screens: {} };
    memory.save(mem);
    return { memory: mem };
  },
};

/**
 * Does a knownIssues entry actually describe the language being scanned?
 *
 * This used to guess, by comparing the first two letters of the entry's key
 * against the language code. Two letters is not an identifier: `thaiGlyphCoverage`
 * begins "th" and so was reported as a known-broken language every time Thai was
 * scanned — a record that Thai *passed* was announced to the tester as a defect.
 * `arabicTofu` ("ar"), `launchPromo` ("la") and `settingsLanguageLeak` ("se") are
 * all one language code away from the same collision.
 *
 * So the entry has to say what it applies to: either an explicit `language`, or a
 * key that is exactly the sheet column or its code. Anything scoped to the whole
 * app stays out of the per-language warning, which is where it belonged anyway.
 */
function routeIssueAppliesTo(key, note, target) {
  if (note && typeof note === 'object' && (note.resolved || note.status === 'PASS')) return false;
  const wanted = [String(target.header || ''), String(target.code || '')]
    .map((s) => s.trim().toLowerCase())
    .filter(Boolean);
  const declared = note && typeof note === 'object' && note.language ? [String(note.language)] : [key];
  return declared.some((d) => wanted.includes(String(d).trim().toLowerCase()));
}

/** A route-map note is a string, or an object carrying one. Never print [object Object]. */
function describeRouteIssue(note) {
  if (typeof note === 'string') return note;
  if (note && typeof note === 'object') {
    const text = note.note || note.$comment || note.message;
    if (typeof text === 'string') return text;
    if (text && typeof text === 'object' && typeof text.note === 'string') return text.note;
  }
  try { return JSON.stringify(note); } catch { return String(note); }
}

function pickOverrides(options) {
  const out = {};
  if (!options || typeof options !== 'object') return out;
  const numeric = [
    'maxScreens', 'maxActions', 'maxDepth', 'settleMs', 'scrollSteps', 'longPressMs', 'bridgePort',
    'maxMinutes', 'stopAfterHighIssues',
  ];
  for (const k of numeric) {
    if (options[k] != null && Number.isFinite(Number(options[k]))) out[k] = Number(options[k]);
  }
  for (const k of ['visionEnabled', 'scrollProbe', 'longPressProbe', 'routeSetLanguage']) {
    if (typeof options[k] === 'boolean') out[k] = options[k];
  }
  for (const k of ['androidPackage', 'model', 'effort', 'baseUrl', 'extraChecks']) {
    if (typeof options[k] === 'string' && options[k].trim()) out[k] = options[k].trim();
  }
  // These are meaningful when empty — a run that clears the focus list has to
  // be able to say so, which a truthiness test would swallow.
  for (const k of ['customRules', 'preSteps']) {
    if (typeof options[k] === 'string') out[k] = options[k];
  }
  for (const k of ['blockedLabels', 'focusLabels', 'onlyLabels']) {
    if (Array.isArray(options[k])) out[k] = options[k].map(String);
  }
  return out;
}

/**
 * Turns the tester's rule and step text into something the crawler can run,
 * and surfaces every syntax error instead of quietly ignoring the line.
 */
function compileAutomation(cfg, run) {
  const { rules, errors: ruleErrors } = compileRules(cfg.customRules || '');
  const { steps, errors: stepErrors } = parseSteps(cfg.preSteps || '');

  for (const e of ruleErrors) {
    run.log(`Custom check ignored — ${e}`, 'warn');
    run.warnings.push(`Custom check ignored — ${e}`);
  }
  for (const e of stepErrors) {
    run.log(`Setup step ignored — ${e}`, 'warn');
    run.warnings.push(`Setup step ignored — ${e}`);
  }
  if (rules.length) run.log(`${rules.length} custom check${rules.length === 1 ? '' : 's'} active.`);
  if (steps.length) run.log(`${steps.length} setup step${steps.length === 1 ? '' : 's'} will run before the crawl.`);

  return { compiledRules: rules, compiledSteps: steps };
}

// ── the scan itself ───────────────────────────────────────────────────────

async function startScan({ run, cfg, sheet, target, mode, serial, route }) {
  run.log(`Scanning "${target.header}" against ${sheet.entries.length} sheet rows.`);

  if (route) {
    const app = route.app || {};
    run.log(`Route map "${app.name || 'unnamed'}" loaded: ${Object.keys(route.screens || {}).length} known screens.`);

    // A route recorded against another build is worse than none, so say so
    // rather than driving taps at coordinates from a different layout.
    const env = routeMaps.packageFor(route, cfg.androidPackage);
    if (cfg.androidPackage && !env) {
      run.warnings.push(
        `The route map does not list package "${cfg.androidPackage}". Its coordinates were recorded against ${Object.values(app.packages || {}).join(', ') || 'another build'}.`
      );
    } else if (env) {
      run.log(`Package matches the route's "${env}" build.`);
    }

    // Languages the route marks as unimplemented would otherwise fill the
    // report with noise that is already known and accepted.
    const known = Object.entries(route.knownIssues || {}).filter(([k]) => k !== '$comment');
    for (const [key, note] of known) {
      if (!routeIssueAppliesTo(key, note, target)) continue;
      run.warnings.push(`Route map flags this language as known-broken on the recorded build: ${describeRouteIssue(note)}`);
    }
    if (route.capabilities && route.capabilities.accessibilityText === false) {
      run.log('Route map: this game exposes no accessibility text, so strings come from pixels.');
    }
  }

  const adb = mode === 'device' ? new Adb(await adbFor(cfg), serial) : null;
  if (adb) {
    try {
      await adb.forward(cfg.bridgePort);
    } catch (e) {
      run.log(`Could not forward port ${cfg.bridgePort}: ${e.message}`, 'warn');
    }
  }

  const bridge = new Bridge(cfg.bridgePort);
  const info = await bridge.connect();
  if (info) {
    run.log(`In-game bridge connected (${info.mode}, ${info.product || 'unknown product'}, ${info.screen ? `${info.screen.width}x${info.screen.height}` : '?'}).`);
  } else if (mode === 'editor') {
    throw new Error(
      `No bridge on 127.0.0.1:${cfg.bridgePort}. Add LocaLinterBridge.cs to the Unity project and enter Play Mode.`
    );
  } else {
    // When the route map records that this game ships without the bridge, that
    // is a settled decision, not a defect — say so once and move on. Nagging
    // about it on every run buries the findings that matter.
    const bridgeExpected = !(route && route.capabilities && route.capabilities.bridge === false);
    if (bridgeExpected) {
      run.log('No in-game bridge — falling back to screenshots plus vision. Add LocaLinterBridge.cs to the build for exact strings and far better coverage.', 'warn');
      run.warnings.push('Ran without the in-game bridge: strings were read from screenshots, so truncation and overflow are judged visually rather than measured.');
    } else {
      run.log('Reading strings from screenshots, as recorded for this game — truncation and overflow are judged visually rather than measured.');
    }
  }

  // If the game exposes its locale, tell the operator when it disagrees with
  // the column being tested — the single most common cause of a noisy report.
  if (info) {
    try {
      const loc = await bridge.locale();
      if (loc && loc.code) {
        run.log(`Game locale reports "${loc.code}"${loc.name ? ` (${loc.name})` : ''}.`);
        const matched = sheet.languageByCodeOrHeader(loc.code);
        if (matched && matched.header !== target.header) {
          run.warnings.push(
            `The game is running in "${loc.code}" but the scan is checking the "${target.header}" column. Switch the game's language or pick the matching column.`
          );
        }
      }
    } catch { /* Unity Localization not installed — fine */ }
  }

  // What earlier runs learned about this app, handed to the model as context
  // and to the crawler as hints about what gets in the way.
  const mem = memory.load(cfg.androidPackage);
  // Two kinds of learned knowledge, and the model needs both: the route map is
  // what a person worked out by playing the build with their hands, and memory
  // is what previous automated runs accumulated. Only the second used to be
  // sent, so every run rediscovered the app's rendering quirks from scratch.
  const learned = [routeMaps.promptContext(route), memory.promptContext(mem)]
    .filter(Boolean).join('\n\n');
  const analyzer = new ClaudeAnalyzer({
    apiKey: cfg.apiKey, model: cfg.model, effort: cfg.effort, baseUrl: cfg.baseUrl,
    extraChecks: cfg.extraChecks,
    memory: learned,
  });
  if (route) run.log(`Loaded the ${route.app && route.app.name ? route.app.name : 'saved'} route map — its screens, probing techniques and known issues are in play.`);
  if (mem.runs) run.log(`Recalling ${mem.runs} previous scan${mem.runs === 1 ? '' : 's'} of ${cfg.androidPackage}.`);
  if (cfg.extraChecks) run.log(`Applying extra checks from the run settings.`);
  if (cfg.baseUrl) run.log(`Model endpoint: ${cfg.baseUrl}`);
  run.log(`Using ${cfg.model} at ${cfg.effort} effort.`);

  const expectsNonLatin = /^(ar|he|fa|ur|hi|bn|ta|te|mr|th|ja|ko|zh|ru|uk|el)/.test(String(target.code || ''));
  const automation = compileAutomation(cfg, run);
  if ((cfg.onlyLabels || []).length) {
    run.log(`Only tapping controls matching: ${cfg.onlyLabels.join(', ')}`);
  }
  if ((cfg.focusLabels || []).length) {
    run.log(`Exploring first: ${cfg.focusLabels.join(', ')}`);
  }
  const crawler = new Crawler({
    cfg: { ...cfg, ...automation },
    adb,
    bridge,
    sheet,
    analyzer,
    run,
    route,
    target: {
      header: target.header,
      code: target.code,
      rtl: RTL_CODES.has(target.code),
      sourceHeader: sheet.englishCol ? sheet.englishCol.header : null,
      expectsNonLatin,
    },
  });

  // Same object the analyzer mutates, so a report fetched mid-run is current.
  run.usage = analyzer.usage;

  try {
    await crawler.start();
    run.usage = analyzer.usage;
    if (run.issues.length) {
      try {
        run.summary = await analyzer.summarize(run.issues, {
          screens: run.screens.length,
          targetHeader: target.header,
        });
        run.usage = analyzer.usage;
      } catch (e) {
        run.log(`Could not write the summary: ${e.message}`, 'warn');
      }
    }
    run.finish(run.stopRequested ? 'stopped' : 'done');
    run.log(`Finished: ${run.screens.length} screens, ${run.issues.length} issues.`);
  } catch (e) {
    run.usage = analyzer.usage;
    throw e;
  }
}

// ── static UI ─────────────────────────────────────────────────────────────

// Set by the desktop shell to the folder holding index.html. Unset when the
// agent runs on its own, in which case it serves the API and nothing else.
const UI_DIR = process.env.LOCALINTER_UI_DIR ? path.resolve(process.env.LOCALINTER_UI_DIR) : '';

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
  '.woff2': 'font/woff2',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

// The UI folder is the whole install, which also contains the agent's own
// source and its config.json — and that file holds the Anthropic API key.
// Only the front-end may be served; everything else is off limits even though
// it sits inside UI_DIR.
const UI_DENY = new Set(['agent', 'node_modules', 'desktop', 'api', '.git']);

function escapeHtmlText(s) {
  return String(s).replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

function serveUi(pathname, res) {
  const rel = pathname === '/' ? 'index.html' : decodeURIComponent(pathname).replace(/^\/+/, '');
  const file = path.resolve(UI_DIR, rel);
  // Never serve outside the UI folder, however the path was spelled.
  if (file !== UI_DIR && !file.startsWith(UI_DIR + path.sep)) return false;

  const segments = path.relative(UI_DIR, file).split(/[\\/]/);
  if (segments.some((s) => UI_DENY.has(s) || s.startsWith('.'))) return false;
  // An extension we do not recognise is not part of the front-end.
  if (!MIME[path.extname(file).toLowerCase()]) return false;

  if (!fs.existsSync(file) || !fs.statSync(file).isFile()) return false;

  res.writeHead(200, {
    'content-type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
    'cache-control': 'no-cache',
  });
  fs.createReadStream(file).pipe(res);
  return true;
}

// ── server ────────────────────────────────────────────────────────────────

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, `http://${HOST}:${PORT}`);
  const pathname = url.pathname;

  if (req.method === 'OPTIONS') {
    cors(res);
    res.writeHead(204);
    res.end();
    return;
  }

  // Static routes win over the dynamic /api/run/:id pattern (so /api/run/start
  // is the start endpoint, not a run whose id happens to be "start").
  const handler = routes[`${req.method} ${pathname}`];

  // ── per-run endpoints (dynamic path segments) ──
  const runMatch = !handler && /^\/api\/run\/([^/]+)(?:\/(events|stop|screenshot))?(?:\/(.+))?$/.exec(pathname);
  if (runMatch) {
    const run = store.getRun(runMatch[1]);
    if (!run) return json(res, 404, { error: 'no such run' });
    const sub = runMatch[2];

    if (!sub && req.method === 'GET') return json(res, 200, run.toJSON());

    if (sub === 'stop' && req.method === 'POST') {
      run.stopRequested = true;
      run.log('Stop requested.');
      return json(res, 200, { ok: true });
    }

    if (sub === 'screenshot' && req.method === 'GET') {
      const name = path.basename(runMatch[3] || '');
      const file = path.join(run.dir, 'screens', name);
      if (!name || !file.startsWith(path.join(run.dir, 'screens')) || !fs.existsSync(file)) {
        return json(res, 404, { error: 'no such screenshot' });
      }
      cors(res);
      res.writeHead(200, { 'content-type': 'image/png', 'cache-control': 'public, max-age=3600' });
      fs.createReadStream(file).pipe(res);
      return;
    }

    if (sub === 'events' && req.method === 'GET') {
      cors(res);
      res.writeHead(200, {
        'content-type': 'text/event-stream',
        'cache-control': 'no-cache',
        connection: 'keep-alive',
      });
      const send = (ev) => {
        try { res.write(`data: ${JSON.stringify(ev)}\n\n`); } catch { /* client vanished */ }
      };
      // replay what already happened so a late subscriber sees the whole run
      for (const ev of run.events) send(ev);
      if (run.status === 'done' || run.status === 'failed' || run.status === 'stopped') {
        send({ type: 'done', status: run.status, error: run.error, issues: run.issues.length });
        res.end();
        return;
      }
      const unsubscribe = run.subscribe(send);
      const ping = setInterval(() => {
        try { res.write(': ping\n\n'); } catch { /* ignore */ }
      }, 15000);
      req.on('close', () => {
        clearInterval(ping);
        unsubscribe();
      });
      return;
    }

    return json(res, 405, { error: 'method not allowed' });
  }

  // ── OAuth landing ──
  // Google sends the *system browser* here, not the app window, so this has to
  // answer with a human-readable page rather than JSON.
  if (pathname === '/oauth/callback' && req.method === 'GET') {
    const done = (ok, heading, detail) => {
      res.writeHead(ok ? 200 : 400, { 'content-type': 'text/html; charset=utf-8' });
      res.end(`<!doctype html><meta charset="utf-8"><title>LocaLinter</title>
<style>body{background:#0a0a0b;color:#ecebe8;font:15px/1.6 'Segoe UI',system-ui,sans-serif;
display:grid;place-items:center;height:100vh;margin:0;text-align:center}
h1{font-size:1.35rem;font-weight:500;margin:0 0 .4rem;color:${ok ? '#74a06a' : '#d4594f'}}
p{color:#8b877f;margin:0;max-width:44ch}</style>
<div><h1>${heading}</h1><p>${detail}</p></div>`);
    };

    const error = url.searchParams.get('error');
    if (error) {
      const state = google.fail(error);
      done(false, state.status === 'cancelled' ? 'Sign-in cancelled' : 'Sign-in failed',
        'Nothing was changed. You can close this tab and go back to LocaLinter.');
      return;
    }
    try {
      const auth = await google.exchange(
        url.searchParams.get('code'),
        url.searchParams.get('state'),
        PORT
      );
      console.log(`Signed in as ${auth.user && auth.user.email}`);
      done(true, 'Signed in', 'You can close this tab and go back to LocaLinter.');
    } catch (e) {
      done(false, 'Sign-in failed', escapeHtmlText(e.message));
    }
    return;
  }

  // ── the app itself ──
  // Serving the UI from this same origin is what makes the desktop build
  // simple: the window loads http://127.0.0.1:8790 rather than a file:// URL,
  // so every /api call is same-origin (no CORS, no private-network preflight)
  // and Google sign-in still sees a real http origin it will accept.
  if (!handler && UI_DIR && req.method === 'GET') {
    if (serveUi(pathname, res)) return;
  }

  if (!handler) return json(res, 404, { error: `no route for ${req.method} ${pathname}` });

  try {
    const body = await handler(req, res, url);
    if (body !== undefined) json(res, 200, body);
  } catch (e) {
    const status = e.status || 500;
    // 4xx is the caller's problem and already reported in the response body;
    // only unexpected failures deserve a stack trace in the agent console.
    if (status >= 500) console.error(`[${req.method} ${pathname}]`, e);
    else console.warn(`[${req.method} ${pathname}] ${e.message}`);
    json(res, status, { error: e.message || String(e) });
  }
});

// The data directory has to exist before anything reads config or writes a
// run, and a packaged build carries its route maps inside the executable.
paths.ensure();
paths.seedRoutes();

/**
 * Registering start-at-login is also useful without a browser — a tester who
 * has just downloaded the exe can run it once with a flag and be done, and it
 * gives IT something scriptable to push out.
 */
async function runCli(flag) {
  try {
    if (flag === '--install-autostart') {
      const s = await autostart.enable();
      console.log(`Start at login: ENABLED\n  ${s.command}`);
    } else if (flag === '--uninstall-autostart') {
      await autostart.disable();
      console.log('Start at login: DISABLED');
    } else {
      const s = await autostart.status();
      console.log(`Start at login: ${s.enabled ? 'ENABLED' : 'disabled'}`);
      if (s.enabled) console.log(`  ${s.command}`);
      if (s.stale) console.log(`  STALE — points somewhere else. Re-run --install-autostart.`);
    }
    process.exit(0);
  } catch (e) {
    console.error(`Failed: ${e.message}`);
    process.exit(1);
  }
}

const cliFlag = process.argv.find((a) => /^--(install|uninstall)-autostart$|^--autostart-status$/.test(a));
if (cliFlag) {
  runCli(cliFlag);
} else {
  server.listen(PORT, HOST, async () => {
    const cfg = config.load();
    // Inside the desktop app this is a service, not something you visit; run
    // from a checkout it is still the standalone agent it always was.
    const embedded = !!process.versions.electron;

    console.log(`${embedded ? 'LocaLinter' : 'LocaLinter agent'} ${VERSION} — service on http://${HOST}:${PORT}`);
    console.log(`  data:     ${paths.DATA_DIR}`);
    const adb = await tools.resolveAdb(cfg);
    console.log(`  adb:      ${adb ? `${adb.path} (${adb.source})` : 'NOT FOUND — Device scan can fetch it'}`);
    console.log(`  model:    ${cfg.model}`);
    console.log(`  API key:  ${cfg.apiKey ? 'configured' : 'NOT SET — add it under Device scan'}`);
    console.log(`  bridge:   127.0.0.1:${cfg.bridgePort}`);
    if (embedded) {
      console.log('Window opening…');
    } else {
      const auto = await autostart.status();
      console.log(`  autostart: ${auto.enabled ? (auto.stale ? 'enabled (STALE — re-enable it)' : 'enabled') : 'off'}`);
      console.log('Open LocaLinter and switch to the Device Scan tab.');
    }
  });
}
