'use strict';

/**
 * Where the agent keeps its config, route maps and run output.
 *
 * Run from a source checkout, that is `agent/` itself — exactly where these
 * things have always lived, so an existing install keeps its config.json and
 * runs/ without being migrated.
 *
 * Packaged as a single executable there is nothing writable to rely on beside
 * the binary: a tester may well drop it on the Desktop or in Downloads, and
 * writing next to it breaks the moment it lands somewhere like Program Files.
 * So it moves to the per-user location Windows expects, and the route maps
 * shipped inside the executable are seeded there on first run.
 */

const fs = require('fs');
const path = require('path');

function isPackaged() {
  // node:sea exists from Node 20.12; absent when running under an older
  // runtime, in which case this is a source checkout by definition.
  try { return require('node:sea').isSea(); } catch { return false; }
}

function defaultDataDir() {
  if (!isPackaged()) return path.join(__dirname, '..');
  const base = process.env.LOCALAPPDATA || process.env.APPDATA || process.env.HOME || process.cwd();
  return path.join(base, 'LocaLinter', 'agent');
}

const DATA_DIR = path.resolve(process.env.LOCALINTER_DATA_DIR || defaultDataDir());
const CONFIG_PATH = path.join(DATA_DIR, 'config.json');
const ROUTES_DIR = path.join(DATA_DIR, 'routes');
const RUNS_DIR = path.join(DATA_DIR, 'runs');
// platform-tools is downloaded here when the tester has no adb of their own.
const TOOLS_DIR = path.join(DATA_DIR, 'tools');

function ensure() {
  for (const dir of [DATA_DIR, ROUTES_DIR, RUNS_DIR]) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

/**
 * Copies the route maps built into the executable into the data directory,
 * once. They are seeded as real files rather than read straight out of the
 * binary so a tester can open the folder, drop in a route map for their own
 * game, and have it show up in the picker. An existing file is never
 * overwritten — a local edit outlives an upgrade.
 */
function seedRoutes() {
  if (!isPackaged()) return;
  let seed;
  try {
    seed = JSON.parse(require('node:sea').getAsset('routes-seed.json', 'utf8'));
  } catch {
    return; // nothing embedded; the picker just starts empty
  }
  for (const [file, data] of Object.entries(seed)) {
    const target = path.join(ROUTES_DIR, file);
    if (fs.existsSync(target)) continue;
    try {
      fs.writeFileSync(target, JSON.stringify(data, null, 2), 'utf8');
    } catch (e) {
      console.warn(`[paths] could not seed ${file}: ${e.message}`);
    }
  }
}

module.exports = { DATA_DIR, CONFIG_PATH, ROUTES_DIR, RUNS_DIR, TOOLS_DIR, ensure, seedRoutes, isPackaged };
