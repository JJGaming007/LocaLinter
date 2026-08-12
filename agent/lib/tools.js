'use strict';

/**
 * Finding adb, and fetching it when the tester has none.
 *
 * A tester who has never installed the Android SDK has no adb, and "No devices
 * found" is a miserable way to discover that. Google publishes platform-tools
 * as a standalone zip, so the agent can fetch it into its own data directory
 * and point itself at it — no SDK install, and nothing of Google's is
 * redistributed in the LocaLinter download.
 *
 * The download is never automatic. It is a few hundred megabits of someone
 * else's bandwidth and a surprise network call, so the browser asks for it
 * explicitly.
 */

const fs = require('fs');
const path = require('path');
const { spawn, execFileSync } = require('child_process');

const { TOOLS_DIR } = require('./paths');

const ZIP_URL = 'https://dl.google.com/android/repository/platform-tools-latest-windows.zip';
const BUNDLED_ADB = path.join(TOOLS_DIR, 'platform-tools', 'adb.exe');

/** True when `bin version` runs, which is the only proof that matters. */
function works(bin) {
  return new Promise((resolve) => {
    const child = spawn(bin, ['version'], { windowsHide: true });
    const done = (ok) => resolve(ok);
    child.on('error', () => done(false));
    child.on('close', (code) => done(code === 0));
    setTimeout(() => { try { child.kill(); } catch { /* gone */ } done(false); }, 8000);
  });
}

/**
 * Where adb is, in order of preference: what the user configured, whatever the
 * agent downloaded earlier, then the PATH.
 */
async function resolveAdb(cfg) {
  const configured = (cfg && cfg.adbPath || '').trim();
  if (configured && await works(configured)) return { path: configured, source: 'configured' };
  if (fs.existsSync(BUNDLED_ADB) && await works(BUNDLED_ADB)) return { path: BUNDLED_ADB, source: 'downloaded' };
  if (await works('adb')) return { path: 'adb', source: 'path' };
  return null;
}

/** Unzip without taking on a dependency: bsdtar ships with Windows 10+. */
function unzip(zip, dest) {
  fs.mkdirSync(dest, { recursive: true });
  try {
    execFileSync('tar', ['-xf', zip, '-C', dest], { stdio: 'pipe', windowsHide: true });
    return;
  } catch (e) {
    // Older builds have no bsdtar; PowerShell is always there.
    execFileSync('powershell', [
      '-NoProfile', '-NonInteractive', '-Command',
      `Expand-Archive -LiteralPath '${zip}' -DestinationPath '${dest}' -Force`,
    ], { stdio: 'pipe', windowsHide: true });
  }
}

/**
 * Fetches platform-tools and returns the adb path. Resolves to the existing
 * copy if one is already there, so calling twice is harmless.
 */
async function downloadAdb(log = () => {}) {
  if (fs.existsSync(BUNDLED_ADB) && await works(BUNDLED_ADB)) {
    log('platform-tools is already installed.');
    return BUNDLED_ADB;
  }

  fs.mkdirSync(TOOLS_DIR, { recursive: true });
  const zip = path.join(TOOLS_DIR, 'platform-tools.zip');

  log(`Downloading platform-tools from ${new URL(ZIP_URL).host}…`);
  const res = await fetch(ZIP_URL);
  if (!res.ok) throw new Error(`platform-tools download failed (HTTP ${res.status})`);
  const bytes = Buffer.from(await res.arrayBuffer());
  fs.writeFileSync(zip, bytes);
  log(`Downloaded ${(bytes.length / 1024 / 1024).toFixed(1)} MB. Extracting…`);

  // A half-extracted previous attempt would shadow the new one.
  fs.rmSync(path.join(TOOLS_DIR, 'platform-tools'), { recursive: true, force: true });
  unzip(zip, TOOLS_DIR);
  fs.rmSync(zip, { force: true });

  if (!fs.existsSync(BUNDLED_ADB)) {
    throw new Error('platform-tools extracted but adb.exe was not where it was expected.');
  }
  if (!await works(BUNDLED_ADB)) {
    throw new Error('adb.exe was downloaded but will not run. A security tool may have quarantined it.');
  }
  log('platform-tools ready.');
  return BUNDLED_ADB;
}

module.exports = { resolveAdb, downloadAdb, BUNDLED_ADB, ZIP_URL };
