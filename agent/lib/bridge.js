'use strict';

const http = require('http');

/**
 * Client for LocaLinterBridge.cs — the tiny HTTP listener that runs inside the
 * Unity game (Editor Play Mode or a development build on the device).
 *
 * On a device the port is reached through `adb forward`, so the host is always
 * 127.0.0.1 from the agent's point of view.
 *
 * Everything degrades gracefully: if the bridge is absent, the crawler falls
 * back to screenshot-only mode and the caller sees `available: false`.
 */
class Bridge {
  constructor(port = 8791, host = '127.0.0.1') {
    this.port = port;
    this.host = host;
    this.available = false;
    this.info = null;
  }

  _request(method, path, body, { timeoutMs = 15000, binary = false } = {}) {
    return new Promise((resolve, reject) => {
      const payload = body == null ? null : Buffer.from(JSON.stringify(body), 'utf8');
      const req = http.request(
        {
          host: this.host,
          port: this.port,
          path,
          method,
          headers: payload
            ? { 'content-type': 'application/json', 'content-length': payload.length }
            : {},
          timeout: timeoutMs,
        },
        (res) => {
          const chunks = [];
          res.on('data', (c) => chunks.push(c));
          res.on('end', () => {
            const buf = Buffer.concat(chunks);
            if (res.statusCode >= 400) {
              reject(new Error(`bridge ${method} ${path} → ${res.statusCode}: ${buf.toString('utf8').slice(0, 300)}`));
              return;
            }
            if (binary) return resolve(buf);
            const txt = buf.toString('utf8');
            if (!txt.trim()) return resolve(null);
            try {
              resolve(JSON.parse(txt));
            } catch (e) {
              reject(new Error(`bridge ${path} returned non-JSON: ${txt.slice(0, 200)}`));
            }
          });
        }
      );
      req.on('timeout', () => { req.destroy(new Error(`bridge ${path} timed out`)); });
      req.on('error', reject);
      if (payload) req.write(payload);
      req.end();
    });
  }

  /** Probes the bridge. Never throws — sets `available` and returns info or null. */
  async connect() {
    try {
      const info = await this._request('GET', '/ping', null, { timeoutMs: 4000 });
      this.available = !!(info && info.ok);
      this.info = info;
      return this.available ? info : null;
    } catch {
      this.available = false;
      this.info = null;
      return null;
    }
  }

  /**
   * Full UI snapshot:
   * { scene, locale, screen:{width,height}, texts:[…], interactables:[…], scrolls:[…] }
   */
  state() {
    return this._request('GET', '/state', null, { timeoutMs: 20000 });
  }

  /** PNG bytes rendered by Unity itself (matches the text rects exactly). */
  screenshot() {
    return this._request('GET', '/screenshot', null, { timeoutMs: 30000, binary: true });
  }

  tap(x, y) {
    return this._request('POST', '/tap', { x, y });
  }

  /** Clicks a specific element by the id returned in /state — more reliable than coordinates. */
  click(id) {
    return this._request('POST', '/click', { id });
  }

  longPress(x, y, ms = 800) {
    return this._request('POST', '/longpress', { x, y, ms });
  }

  back() {
    return this._request('POST', '/back', {});
  }

  /** normalized 0..1 vertical position for a ScrollRect id */
  scroll(id, position) {
    return this._request('POST', '/scroll', { id, position });
  }

  locale() {
    return this._request('GET', '/locale', null, { timeoutMs: 8000 });
  }

  setLocale(code) {
    return this._request('POST', '/locale', { code }, { timeoutMs: 20000 });
  }
}

module.exports = { Bridge };
