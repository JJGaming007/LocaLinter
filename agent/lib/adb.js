'use strict';

const { spawn } = require('child_process');

/**
 * Thin, promise-based wrapper around the `adb` CLI.
 * Every call is explicit about the target serial so multi-device setups work.
 */
class Adb {
  constructor(adbPath, serial) {
    this.bin = adbPath && adbPath.trim() ? adbPath.trim() : 'adb';
    this.serial = serial || null;
  }

  _args(args) {
    return this.serial ? ['-s', this.serial, ...args] : args;
  }

  /** Runs adb and resolves with a Buffer of stdout (binary-safe). */
  raw(args, { timeoutMs = 30000, input = null } = {}) {
    return new Promise((resolve, reject) => {
      const child = spawn(this.bin, this._args(args), { windowsHide: true });
      const out = [];
      const err = [];
      let settled = false;

      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        try { child.kill(); } catch { /* already gone */ }
        reject(new Error(`adb ${args.join(' ')} timed out after ${timeoutMs}ms`));
      }, timeoutMs);

      child.stdout.on('data', (d) => out.push(d));
      child.stderr.on('data', (d) => err.push(d));
      child.on('error', (e) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(new Error(
          e.code === 'ENOENT'
            ? `adb not found at "${this.bin}". Install platform-tools or set adbPath in the agent settings.`
            : e.message
        ));
      });
      child.on('close', (code) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        const stdout = Buffer.concat(out);
        const stderr = Buffer.concat(err).toString('utf8');
        if (code !== 0) {
          reject(new Error(`adb ${args.join(' ')} failed (${code}): ${stderr.trim() || stdout.toString('utf8').trim()}`));
          return;
        }
        resolve(stdout);
      });

      if (input != null) child.stdin.end(input);
      else child.stdin.end();
    });
  }

  async text(args, opts) {
    return (await this.raw(args, opts)).toString('utf8');
  }

  async shell(cmd, opts) {
    return this.text(['shell', cmd], opts);
  }

  /** Lists attached devices: [{ serial, state, model }] */
  static async devices(adbPath) {
    const adb = new Adb(adbPath, null);
    const out = await adb.text(['devices', '-l'], { timeoutMs: 15000 });
    return out
      .split('\n')
      .slice(1)
      .map((l) => l.trim())
      .filter(Boolean)
      .map((line) => {
        const [serial, state, ...rest] = line.split(/\s+/);
        const meta = rest.join(' ');
        const model = /model:(\S+)/.exec(meta);
        const device = /device:(\S+)/.exec(meta);
        return {
          serial,
          state,
          model: (model && model[1]) || (device && device[1]) || '',
        };
      })
      .filter((d) => d.serial && d.serial !== 'List');
  }

  /** PNG screenshot as a Buffer. */
  async screenshot() {
    // exec-out avoids the CRLF mangling that plagues `adb shell screencap`.
    const buf = await this.raw(['exec-out', 'screencap', '-p'], { timeoutMs: 45000 });
    if (buf.length > 8 && buf[0] === 0x89 && buf[1] === 0x50) return buf;
    // Some older/rooted ROMs still translate newlines — repair the stream.
    const repaired = Buffer.from(buf.toString('binary').replace(/\r\n/g, '\n'), 'binary');
    if (repaired.length > 8 && repaired[0] === 0x89) return repaired;
    throw new Error('screencap did not return a PNG. Is the screen on and unlocked?');
  }

  async screenSize() {
    const out = await this.shell('wm size');
    const m = /Override size:\s*(\d+)x(\d+)/.exec(out) || /Physical size:\s*(\d+)x(\d+)/.exec(out);
    if (!m) throw new Error(`could not parse screen size from: ${out.trim()}`);
    return { width: Number(m[1]), height: Number(m[2]) };
  }

  tap(x, y) {
    return this.shell(`input tap ${Math.round(x)} ${Math.round(y)}`);
  }

  longPress(x, y, ms = 800) {
    const p = `${Math.round(x)} ${Math.round(y)}`;
    return this.shell(`input swipe ${p} ${p} ${Math.round(ms)}`);
  }

  swipe(x1, y1, x2, y2, ms = 300) {
    return this.shell(
      `input swipe ${Math.round(x1)} ${Math.round(y1)} ${Math.round(x2)} ${Math.round(y2)} ${Math.round(ms)}`
    );
  }

  back() {
    return this.shell('input keyevent KEYCODE_BACK');
  }

  /** Forwards a local TCP port to the device so we can reach the in-app bridge. */
  forward(localPort, devicePort = localPort) {
    return this.text(['forward', `tcp:${localPort}`, `tcp:${devicePort}`], { timeoutMs: 15000 });
  }

  removeForward(localPort) {
    return this.text(['forward', '--remove', `tcp:${localPort}`], { timeoutMs: 15000 })
      .catch(() => '');
  }

  /** Foreground package/activity, best-effort across Android versions. */
  async currentActivity() {
    const out = await this.shell(
      "dumpsys activity activities | grep -E 'mResumedActivity|ResumedActivity|topResumedActivity' | head -n 3"
    ).catch(() => '');
    const m = /([A-Za-z0-9_.]+)\/([A-Za-z0-9_.$]+)/.exec(out);
    return m ? { package: m[1], activity: m[2] } : { package: '', activity: '' };
  }

  /** Accessibility hierarchy XML. Empty for Unity/Flutter canvases — that is expected. */
  async uiHierarchy() {
    try {
      const out = await this.raw(['exec-out', 'uiautomator', 'dump', '/dev/tty'], { timeoutMs: 20000 });
      const xml = out.toString('utf8');
      const start = xml.indexOf('<?xml');
      return start >= 0 ? xml.slice(start) : '';
    } catch {
      return '';
    }
  }

  async deviceLocale() {
    const props = ['persist.sys.locale', 'ro.product.locale', 'persist.sys.language'];
    for (const p of props) {
      const v = (await this.shell(`getprop ${p}`).catch(() => '')).trim();
      if (v) return v;
    }
    return '';
  }

  forceStop(pkg) {
    return this.shell(`am force-stop ${pkg}`);
  }

  launch(pkg) {
    return this.shell(`monkey -p ${pkg} -c android.intent.category.LAUNCHER 1`);
  }
}

module.exports = { Adb };
