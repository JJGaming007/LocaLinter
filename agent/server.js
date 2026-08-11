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
const { SheetIndex, RTL_CODES } = require('./lib/sheet');
const { ClaudeAnalyzer } = require('./lib/claude');
const { Crawler } = require('./lib/crawler');
const store = require('./lib/store');
const routeMaps = require('./lib/routes');

const PORT = Number(process.env.PORT || 8790);
const HOST = '127.0.0.1';
const VERSION = '1.0.0';

// ── helpers ───────────────────────────────────────────────────────────────

function cors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'content-type');
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

  'GET /api/devices': async () => {
    const cfg = config.load();
    try {
      const devices = await Adb.devices(cfg.adbPath);
      return { devices, adbPath: cfg.adbPath || 'adb' };
    } catch (e) {
      return { devices: [], error: e.message, adbPath: cfg.adbPath || 'adb' };
    }
  },

  'POST /api/probe': async (req) => {
    const cfg = config.load();
    const body = await readBody(req);
    const mode = body.mode === 'editor' ? 'editor' : 'device';
    const port = Number(body.bridgePort || cfg.bridgePort);
    const result = { mode, bridgePort: port, bridge: null, device: null, forwarded: false, errors: [] };

    if (mode === 'device') {
      const serial = body.serial || null;
      const adb = new Adb(cfg.adbPath, serial);
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

    const target =
      sheet.languageByCodeOrHeader(body.targetLanguage) ||
      sheet.languages.find((l) => l.code !== 'en') ||
      sheet.languages[0];
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
};

function pickOverrides(options) {
  const out = {};
  if (!options || typeof options !== 'object') return out;
  const numeric = ['maxScreens', 'maxActions', 'maxDepth', 'settleMs', 'scrollSteps', 'longPressMs', 'bridgePort'];
  for (const k of numeric) {
    if (options[k] != null && Number.isFinite(Number(options[k]))) out[k] = Number(options[k]);
  }
  for (const k of ['visionEnabled', 'scrollProbe', 'longPressProbe', 'routeSetLanguage']) {
    if (typeof options[k] === 'boolean') out[k] = options[k];
  }
  for (const k of ['androidPackage', 'model', 'effort']) {
    if (typeof options[k] === 'string' && options[k].trim()) out[k] = options[k].trim();
  }
  if (Array.isArray(options.blockedLabels)) out.blockedLabels = options.blockedLabels.map(String);
  return out;
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
      if (String(target.code || '').toLowerCase().startsWith(key.slice(0, 2).toLowerCase())
          || String(target.header || '').toLowerCase() === key.toLowerCase()) {
        run.warnings.push(`Route map flags this language as known-broken on the recorded build: ${note}`);
      }
    }
    if (route.capabilities && route.capabilities.accessibilityText === false) {
      run.log('Route map: this game exposes no accessibility text, so strings come from pixels.');
    }
  }

  const adb = mode === 'device' ? new Adb(cfg.adbPath, serial) : null;
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
    run.log('No in-game bridge — falling back to screenshots plus vision. Add LocaLinterBridge.cs to the build for exact strings and far better coverage.', 'warn');
    run.warnings.push('Ran without the in-game bridge: strings were read from screenshots, so truncation and overflow are judged visually rather than measured.');
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

  const analyzer = new ClaudeAnalyzer({ apiKey: cfg.apiKey, model: cfg.model, effort: cfg.effort });
  run.log(`Using ${cfg.model} at ${cfg.effort} effort.`);

  const expectsNonLatin = /^(ar|he|fa|ur|hi|bn|ta|te|mr|th|ja|ko|zh|ru|uk|el)/.test(String(target.code || ''));
  const crawler = new Crawler({
    cfg,
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

server.listen(PORT, HOST, () => {
  const cfg = config.load();
  console.log(`LocaLinter agent ${VERSION} listening on http://${HOST}:${PORT}`);
  console.log(`  adb:      ${cfg.adbPath || 'adb (from PATH)'}`);
  console.log(`  model:    ${cfg.model}`);
  console.log(`  API key:  ${cfg.apiKey ? 'configured' : 'NOT SET — add it in the Device Scan tab'}`);
  console.log(`  bridge:   127.0.0.1:${cfg.bridgePort}`);
  console.log('Open LocaLinter in your browser and switch to the Device Scan tab.');
});
