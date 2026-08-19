'use strict';

const fs = require('fs');

const { CONFIG_PATH } = require('./paths');

const DEFAULTS = {
  // Claude
  apiKey: '',                       // falls back to process.env.ANTHROPIC_API_KEY
  baseUrl: '',                      // '' => api.anthropic.com; set for a company gateway
  extraChecks: '',                  // project-specific rules added to the vision pass
  model: 'claude-opus-5',
  effort: 'high',                   // low | medium | high | xhigh | max
  visionEnabled: true,              // run the Claude vision pass per screen

  // Tooling
  adbPath: '',                      // '' => resolve `adb` from PATH
  bridgePort: 8791,                 // LocaLinterBridge listener port (device is adb-forwarded)

  // Crawl limits
  maxScreens: 120,
  maxActions: 400,
  maxDepth: 12,
  settleMs: 900,                    // wait after a tap before capturing
  settleTimeoutMs: 6000,
  maxMinutes: 0,                    // wall-clock budget for a run; 0 => no limit
  stopAfterHighIssues: 0,           // bail out once this many high findings land; 0 => no limit

  // Steering: which controls the crawl spends its budget on.
  // focusLabels are explored first; onlyLabels, when set, is the whole world.
  focusLabels: [],
  onlyLabels: [],

  // Deterministic house rules, one per line, in the mini-language in rules.js:
  //   forbid: \bROBINET\b | high | Never use ROBINET for a tap control
  //   maxlen: 24 on btn_ | medium | Buttons must fit one line
  customRules: '',

  // Replayed on the device before the crawl starts — dismiss the daily popup,
  // sign in, walk to the screen that actually needs checking.
  preSteps: '',

  // Safety: never tap an element whose label matches one of these (case-insensitive).
  // Protects against real purchases / destructive actions on a live device.
  // What a scan must never tap. Deliberately NOT a list of things that cost
  // money: a store build runs on a test account with a test card, nothing is
  // really charged, and the purchase and confirmation dialogs carry some of the
  // most commercially visible text in the app — exactly what a localization
  // scan is for. Blocking them was hiding the screens most worth reading.
  //
  // What is left is the set that ends the run or destroys the account it runs
  // on, which no amount of test credit makes recoverable.
  // Re-examine rendering findings on a magnified crop before reporting them.
  // Truncation, overflow and missing diacritics are all decided on a few pixels
  // in a 2400x1080 frame, which is where a scan invents defects; a second look
  // at 4x retracts the ones that were never there. Costs one call per finding
  // checked, capped per screen.
  // Reach a goal by looking at the screen and deciding the next move, instead
  // of replaying recorded coordinates. Recorded steps assume the app is where
  // it was when they were written, and fail silently when it is not.
  // Ask the model what the sheet says about each string on screen, rather than
  // matching them mechanically. A shipped string is rarely byte-identical to its
  // row once a placeholder is filled and the UI has upper-cased it, and on a
  // build with no accessibility bridge nothing was comparing them at all.
  modelSheetCompare: true,

  modelDrivenNavigation: true,

  zoomVerify: true,
  zoomVerifyMax: 6,

  blockedLabels: [
    'delete account', 'delete', 'remove account', 'reset progress', 'wipe',
    'log ?out', 'sign ?out', 'logout', 'quit', 'exit game', 'uninstall'
  ],
  // Elements matching these are probed even when they look non-interactive
  // (info "i" buttons, "?" help chips, tooltips, dropdown carets…)
  probeLabels: ['^i$', '^\\?$', 'info', 'help', 'details', 'more', 'tooltip', 'about'],

  // Long-press probing (reveals tooltips)
  longPressProbe: true,
  longPressMs: 800,

  // Scroll every ScrollRect through its full range and capture each step
  scrollProbe: true,
  scrollSteps: 4,

  // Drive the route map's language procedure before crawling, so the game is
  // in the language the sheet column is being checked against.
  routeSetLanguage: true,

  // Keep a copy of every recognised screen when scanning in the source
  // language, and compare against it when scanning in any other. The pair is
  // what turns "this label looks too long" into a defect with evidence — and,
  // just as usefully, what clears a suspect label whose container is equally
  // broken in the source.
  englishBaseline: true,

  // Android app under test (used to restart when backtracking fails)
  androidPackage: '',
};

function load() {
  let stored = {};
  try {
    if (fs.existsSync(CONFIG_PATH)) {
      stored = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
    }
  } catch (e) {
    console.warn('[config] could not read config.json:', e.message);
  }
  const cfg = { ...DEFAULTS, ...stored };
  if (!cfg.apiKey && process.env.ANTHROPIC_API_KEY) cfg.apiKey = process.env.ANTHROPIC_API_KEY;
  return cfg;
}

function save(patch) {
  let stored = {};
  try {
    if (fs.existsSync(CONFIG_PATH)) stored = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
  } catch { /* start fresh */ }
  const next = { ...stored, ...patch };
  fs.mkdirSync(require('path').dirname(CONFIG_PATH), { recursive: true });
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(next, null, 2), 'utf8');
  return load();
}

/**
 * Anthropic keys are `sk-ant-…` and around a hundred characters. Checking the
 * shape catches the common mistake — something else pasted into the field —
 * before it costs a minute of scanning and comes back as a bare 401.
 *
 * Only when talking to Anthropic directly, though. A gateway (LiteLLM, a
 * company proxy) mints its own keys in whatever format it likes, and they are
 * usually far shorter, so with a base URL set there is nothing to check: the
 * gateway is the only thing that can judge its own credentials.
 *
 * Deliberately a shape check, not validation — only the far end can say
 * whether a well-formed key is live.
 */
function keyLooksWrong(apiKey, baseUrl = '') {
  if (!apiKey) return '';
  if (baseUrl) return '';
  if (!apiKey.startsWith('sk-ant-')) {
    return 'That does not look like an Anthropic key — they begin with "sk-ant-". If it is for a gateway, set the base URL under Advanced.';
  }
  if (apiKey.length < 40) return `That key is only ${apiKey.length} characters; Anthropic keys are much longer. Was it truncated?`;
  return '';
}

/** Config minus secrets, safe to hand to the browser UI. */
function redact(cfg) {
  const { apiKey, ...rest } = cfg;
  return {
    ...rest,
    apiKeySet: !!apiKey,
    apiKeyHint: apiKey ? `…${apiKey.slice(-4)}` : '',
    apiKeyFromEnv: !!process.env.ANTHROPIC_API_KEY && !readStoredKey(),
    apiKeyWarning: keyLooksWrong(apiKey, cfg.baseUrl),
  };
}

function readStoredKey() {
  try {
    if (!fs.existsSync(CONFIG_PATH)) return '';
    return JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')).apiKey || '';
  } catch {
    return '';
  }
}

module.exports = { load, save, redact, keyLooksWrong, DEFAULTS, CONFIG_PATH };
