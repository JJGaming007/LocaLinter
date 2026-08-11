'use strict';

const fs = require('fs');
const path = require('path');

const CONFIG_PATH = path.join(__dirname, '..', 'config.json');

const DEFAULTS = {
  // Claude
  apiKey: '',                       // falls back to process.env.ANTHROPIC_API_KEY
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

  // Safety: never tap an element whose label matches one of these (case-insensitive).
  // Protects against real purchases / destructive actions on a live device.
  blockedLabels: [
    'buy', 'purchase', 'checkout', 'pay', 'subscribe', 'confirm purchase',
    'delete account', 'delete', 'remove account', 'reset progress', 'wipe',
    'log ?out', 'sign ?out', 'logout', 'quit', 'exit game',
    'uninstall', 'restore purchase', 'redeem', 'top ?up', 'recharge'
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
  fs.writeFileSync(CONFIG_PATH, JSON.stringify(next, null, 2), 'utf8');
  return load();
}

/** Config minus secrets, safe to hand to the browser UI. */
function redact(cfg) {
  const { apiKey, ...rest } = cfg;
  return {
    ...rest,
    apiKeySet: !!apiKey,
    apiKeyHint: apiKey ? `…${apiKey.slice(-4)}` : '',
    apiKeyFromEnv: !!process.env.ANTHROPIC_API_KEY && !readStoredKey(),
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

module.exports = { load, save, redact, DEFAULTS, CONFIG_PATH };
