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
  if (!isPackaged()) return seedRoutesFromDisk();
  let seed;
  try {
    seed = JSON.parse(require('node:sea').getAsset('routes-seed.json', 'utf8'));
  } catch {
    return seedRoutesFromDisk(); // nothing embedded; fall back to the shipped folder
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

/**
 * The same seeding, for every build that is not a single executable.
 *
 * The desktop app is the case this exists for. It is not a SEA, so isPackaged()
 * is false, but it still points LOCALINTER_DATA_DIR at the per-user folder just
 * like the packaged binary does — which left the route maps sitting in the
 * install next to server.js while the agent looked for them somewhere else, and
 * the picker came up empty. A scan then ran with no route map at all: no app to
 * open, no modal auto-dismissed, and the back button live on a lobby where it
 * quits the game.
 *
 * Copying is skipped when the two directories are the same, which is the source
 * checkout — there the files are already where they belong.
 */
function seedRoutesFromDisk() {
  const src = path.join(__dirname, '..', 'routes');
  if (path.resolve(src) === path.resolve(ROUTES_DIR)) return;
  let files;
  try {
    files = fs.readdirSync(src).filter((f) => f.endsWith('.json'));
  } catch {
    return; // no shipped route maps; the picker just starts empty
  }
  for (const file of files) {
    try {
      seedOneRoute(path.join(src, file), path.join(ROUTES_DIR, file));
    } catch (e) {
      console.warn(`[paths] could not seed ${file}: ${e.message}`);
    }
  }
}

/**
 * Copy a shipped route map into the data directory, and keep it up to date.
 *
 * "Never overwrite an existing file" protects a tester's local edits, which is
 * right — but on its own it also means a route map improved upstream never
 * reaches anyone who has run the app once. Everything learned on the device
 * this week was invisible to the desktop build for exactly that reason: the
 * agent kept loading a first-run copy while the real map sat updated beside it,
 * and the scans that followed looked like crawler bugs.
 *
 * So the two cases are separated. A copy the tester has not touched is just a
 * stale cache and gets replaced. A copy they have edited is theirs and is left
 * alone, with a line in the log saying the update was held back — silence there
 * would recreate the same confusion in the other direction.
 */
function seedOneRoute(srcFile, target) {
  const shipped = fs.readFileSync(srcFile);
  const stamp = `${target}.seeded`;

  if (!fs.existsSync(target)) {
    fs.writeFileSync(target, shipped);
    writeStamp(stamp, shipped);
    return;
  }

  const current = fs.readFileSync(target);
  if (current.equals(shipped)) return;               // already up to date

  const seededHash = readStamp(stamp);
  const currentHash = sha256(current);
  if (seededHash && seededHash !== currentHash) {
    console.warn(`[paths] ${path.basename(target)} has local edits — keeping them, not applying the shipped update.`);
    return;
  }

  fs.writeFileSync(target, shipped);
  writeStamp(stamp, shipped);
}

function sha256(buf) {
  return require('crypto').createHash('sha256').update(buf).digest('hex');
}
function readStamp(file) {
  try { return fs.readFileSync(file, 'utf8').trim(); } catch { return null; }
}
function writeStamp(file, buf) {
  try { fs.writeFileSync(file, sha256(buf), 'utf8'); } catch { /* best effort */ }
}

module.exports = { DATA_DIR, CONFIG_PATH, ROUTES_DIR, RUNS_DIR, TOOLS_DIR, ensure, seedRoutes, isPackaged };
