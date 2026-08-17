'use strict';

/**
 * Launch the agent when the user logs in.
 *
 * A tester should not have to remember to start anything: the browser cannot
 * spawn a local process (no page may, by design), so the only way the Device
 * Scan tab is reliably Connected is for the agent to already be running. That
 * means the OS has to start it.
 *
 * On Windows this is one string under HKCU\...\CurrentVersion\Run. HKCU rather
 * than HKLM so it never needs an administrator, and `cmd /c start /min` so the
 * console window the agent lives in does not land in the middle of the screen
 * at every login.
 *
 * Other platforms would need launchd or a .desktop file; the agent is Windows
 * only today, so this reports unsupported rather than pretending.
 */

const path = require('path');
const { execFile } = require('child_process');

const RUN_KEY = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run';
const VALUE_NAME = 'LocaLinterAgent';

const supported = process.platform === 'win32';

function reg(args) {
  return new Promise((resolve, reject) => {
    execFile('reg', args, { windowsHide: true }, (err, stdout, stderr) => {
      if (err) {
        // reg.exe exits 1 for "value does not exist", which is a normal answer
        // to a query, not a failure. The caller decides.
        err.stdout = stdout;
        err.stderr = stderr;
        return reject(err);
      }
      resolve(stdout);
    });
  });
}

/**
 * What the Run entry should contain. Packaged, that is the executable itself;
 * from a source checkout it is this machine's node plus server.js, so a
 * developer gets the same behaviour without building an exe first.
 */
function launchCommand() {
  const exe = process.execPath;
  const packaged = path.basename(exe).toLowerCase() !== 'node.exe';
  const target = packaged
    ? `"${exe}"`
    : `"${exe}" "${path.join(__dirname, '..', 'server.js')}"`;
  // The empty "" is start's title argument — without it, start treats the
  // quoted path as the window title and never launches anything.
  return `cmd /c start "" /min ${target}`;
}

async function status() {
  if (!supported) {
    return { supported: false, enabled: false, platform: process.platform, command: '' };
  }
  const expected = launchCommand();
  try {
    const out = await reg(['query', RUN_KEY, '/v', VALUE_NAME]);
    const line = out.split(/\r?\n/).find((l) => l.includes(VALUE_NAME)) || '';
    // "    LocaLinterAgent    REG_SZ    cmd /c start "" /min "C:\…""
    const current = line.split(/REG_SZ\s+/)[1]?.trim() || '';
    return {
      supported: true,
      enabled: true,
      command: current,
      // A moved or rebuilt exe leaves a stale entry pointing at nothing.
      stale: current !== expected,
      expected,
    };
  } catch {
    return { supported: true, enabled: false, command: '', stale: false, expected };
  }
}

async function enable() {
  if (!supported) throw new Error(`Start at login is only implemented for Windows (this is ${process.platform}).`);
  await reg(['add', RUN_KEY, '/v', VALUE_NAME, '/t', 'REG_SZ', '/d', launchCommand(), '/f']);
  return status();
}

async function disable() {
  if (!supported) throw new Error(`Start at login is only implemented for Windows (this is ${process.platform}).`);
  try {
    await reg(['delete', RUN_KEY, '/v', VALUE_NAME, '/f']);
  } catch (e) {
    // Already gone is the state we wanted anyway.
    if (!/cannot find|unable to find/i.test(`${e.stderr || ''}${e.message}`)) throw e;
  }
  return status();
}

module.exports = { status, enable, disable, launchCommand, supported };
