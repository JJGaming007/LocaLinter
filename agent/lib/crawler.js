'use strict';

const { runChecks } = require('./checks');
const { perceptualHash, hammingDistance, size: pngSize } = require('./png');
const { norm } = require('./sheet');
const memory = require('./memory');

// Findings that are a claim about *which language* a string is in. These are
// the ones the sheet can settle; a truncation or an overlap it cannot.
const LANGUAGE_CLAIMS = new Set([
  'wrong_language', 'untranslated', 'missing_translation', 'probably_untranslated',
]);

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function safeRegexes(patterns) {
  const out = [];
  for (const p of patterns || []) {
    try { out.push(new RegExp(p, 'i')); } catch { /* ignore a bad user pattern */ }
  }
  return out;
}

/**
 * Drives the app and captures every reachable screen.
 *
 * Two capture paths, chosen automatically:
 *   bridge  — LocaLinterBridge.cs is running: exact strings, rects, element ids,
 *             engine-reported truncation, reliable clicking by id.
 *   vision  — no bridge: ADB screenshots plus Claude-proposed tap targets.
 *
 * Coverage is state-graph based, not screen-order based: a flyout, a dropdown
 * with its options open, and a tooltip are each their own state, so opening an
 * info button is explored exactly like navigating to a new menu.
 */
class Crawler {
  constructor({ cfg, adb, bridge, sheet, analyzer, run, target, route, memory: mem }) {
    this.cfg = cfg;
    this.route = route || null;
    this.routeScreens = (route && route.screens) || {};
    this.observedText = new Map();     // screenId -> strings read off that screen
    this.routeScreenFor = new Map();   // screenId -> name in the route map
    this.recoveries = 0;
    this.mem = mem || null;
    this.restarts = 0;
    this.wrongAppSkips = 0;
    this.maxRestarts = 3;
    this.langMismatchChecked = false;
    this.adb = adb;
    this.bridge = bridge;
    this.sheet = sheet;
    this.analyzer = analyzer;
    this.run = run;
    this.target = target;              // { header, code, rtl, sourceHeader, expectsNonLatin }

    this.useBridge = !!(bridge && bridge.available);
    this.mode = this.useBridge ? (bridge.info && bridge.info.mode === 'editor' ? 'unity-editor' : 'unity-bridge') : 'adb-vision';

    this.blocked = safeRegexes(cfg.blockedLabels);
    this.probe = safeRegexes(cfg.probeLabels);

    this.visited = new Map();          // sig -> screenId
    this.hashes = [];                  // [{ hash, sig }] for near-duplicate detection
    this.triedLabels = new Map();      // screenId -> labels already acted on (vision mode)
    this.stack = [];                   // pending { action, parent }
    this.actionCount = 0;
    this.screenCount = 0;
    this.deviceSize = null;
    this.rootSig = null;
  }

  // ── capture ────────────────────────────────────────────────────────────

  async screenshot() {
    if (this.useBridge) {
      try {
        return await this.bridge.screenshot();
      } catch (e) {
        this.run.log(`Bridge screenshot failed (${e.message}); using ADB instead.`, 'warn');
      }
    }
    if (!this.adb) throw new Error('No way to capture a screenshot: no bridge and no device.');
    return this.adb.screenshot();
  }

  async capture() {
    let state = null;
    if (this.useBridge) {
      try {
        state = await this.bridge.state();
      } catch (e) {
        this.run.log(`Bridge state failed: ${e.message}`, 'warn');
      }
    }
    const png = await this.screenshot();
    // `adb shell wm size` reports the panel in its natural orientation, which
    // for a landscape game is the wrong way round — taps computed against it
    // land nowhere near the control. The screenshot is what is actually on
    // screen, so its dimensions are the coordinate space to tap in.
    try {
      const s = pngSize(png);
      if (s.width && s.height) {
        if (!this.deviceSize || this.deviceSize.width !== s.width || this.deviceSize.height !== s.height) {
          if (this.deviceSize) {
            this.run.log(`Screen is ${s.width}x${s.height} (was ${this.deviceSize.width}x${this.deviceSize.height}) — using the screenshot's orientation.`);
          }
          this.deviceSize = s;
        }
      }
    } catch { /* keep whatever we had */ }
    let hash = null;
    try { hash = perceptualHash(png); } catch { /* unusual PNG flavour; fall back to text signature */ }
    return { state, png, hash };
  }

  /** Stable identity for a screen. Text content first, pixels only as a fallback. */
  signature(cap) {
    if (cap.state && Array.isArray(cap.state.texts)) {
      const texts = cap.state.texts
        .filter((t) => t.active !== false && String(t.text || '').trim())
        .map((t) => norm(t.text))
        .sort();
      const acts = (cap.state.interactables || []).map((i) => `${i.path}`).sort();
      return `t:${cap.state.scene || ''}|${texts.join('')}|${acts.join('')}`;
    }
    return `h:${cap.hash || Math.random()}`;
  }

  /**
   * Near-duplicate check for vision mode, where a few animated pixels differ.
   * Kept tight on purpose: a loose threshold silently drops a real screen, and
   * re-analysing a near-identical one only costs a little.
   */
  matchExistingHash(hash) {
    if (!hash) return null;
    for (const h of this.hashes) {
      if (hammingDistance(h.hash, hash) <= 3) return h.sig;
    }
    return null;
  }

  /** Waits until the UI stops changing, so we never analyse a mid-transition frame. */
  async settle() {
    const deadline = Date.now() + this.cfg.settleTimeoutMs;
    await sleep(this.cfg.settleMs);
    let prev = await this.capture();
    while (Date.now() < deadline) {
      await sleep(400);
      const next = await this.capture();
      const sameText = this.signature(prev) === this.signature(next);
      const samePixels = prev.hash && next.hash ? hammingDistance(prev.hash, next.hash) <= 4 : true;
      if (sameText && samePixels) return next;
      prev = next;
    }
    return prev;
  }

  // ── actions ────────────────────────────────────────────────────────────

  isBlocked(label) {
    const l = String(label || '').trim();
    if (!l) return false;
    return this.blocked.some((re) => re.test(l));
  }

  shouldProbe(label) {
    const l = String(label || '').trim();
    return this.probe.some((re) => re.test(l));
  }

  /** Enumerates the actions available on the current screen. */
  async actionsFor(cap, screenId) {
    const actions = [];

    if (cap.state && Array.isArray(cap.state.interactables)) {
      for (const el of cap.state.interactables) {
        const label = el.label || el.name || el.path || '';
        if (this.isBlocked(label)) {
          this.run.skipped.push({ screenId, label, path: el.path, reason: 'matches a blocked-label pattern' });
          continue;
        }
        actions.push({ key: `click:${el.path}`, kind: 'click', id: el.id, label, path: el.path });
        if (this.cfg.longPressProbe && this.shouldProbe(label)) {
          actions.push({
            key: `long:${el.path}`, kind: 'long_press', label,
            x: el.rect ? el.rect.x + el.rect.w / 2 : null,
            y: el.rect ? el.rect.y + el.rect.h / 2 : null,
            id: el.id, path: el.path,
          });
        }
      }
    } else {
      // vision mode — let the model find the controls
      let targets = [];
      try {
        targets = await this.analyzer.proposeTargets(cap.png, this.triedLabels.get(screenId) || []);
      } catch (e) {
        this.run.log(`Could not propose tap targets: ${e.message}`, 'warn');
      }
      const size = this.deviceSize || { width: 1080, height: 1920 };
      for (const t of targets) {
        if (this.isBlocked(t.label)) {
          this.run.skipped.push({ screenId, label: t.label, reason: 'matches a blocked-label pattern' });
          continue;
        }
        const x = Math.max(0, Math.min(1, Number(t.x))) * size.width;
        const y = Math.max(0, Math.min(1, Number(t.y))) * size.height;
        if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
        const kind =
          t.kind === 'long_press' ? 'long_press'
            : t.kind === 'scroll_down' ? 'swipe_up'
              : t.kind === 'scroll_right' ? 'swipe_left'
                : t.kind === 'back' ? 'back'
                  : 'tap';
        actions.push({ key: `${kind}:${t.label}:${Math.round(x)},${Math.round(y)}`, kind, x, y, label: t.label, priority: t.priority });
      }
      // deepest-looking controls first
      actions.sort((a, b) => (a.priority === 'high' ? -1 : 1) - (b.priority === 'high' ? -1 : 1));
    }

    // Anything the route map already knows about this screen goes first, and
    // wins over a guessed tap at roughly the same spot.
    const fromRoute = this.routeActionsFor(screenId);
    if (fromRoute.length) {
      const near = (a, b) => Math.hypot(a.x - b.x, a.y - b.y) < 40 && a.kind === b.kind;
      const kept = actions.filter((a) => !fromRoute.some((r) => near(a, r)));
      this.run.log(`Route map contributed ${fromRoute.length} known controls on ${screenId}.`);
      return [...fromRoute, ...kept];
    }

    return actions;
  }

  // ── route map ──────────────────────────────────────────────────────────
  // A route map is what an earlier pass worked out about this game. Using it
  // means the crawl does not have to rediscover where the settings tabs and
  // info badges are on every run, and can recognise a stuck state instead of
  // tapping at a modal until it runs out of budget.

  /** Normalised route coordinates -> device pixels. */
  routePoint(pair) {
    const size = this.deviceSize || { width: 1080, height: 1920 };
    if (!Array.isArray(pair) || pair.length !== 2) return null;
    const x = Number(pair[0]) * size.width;
    const y = Number(pair[1]) * size.height;
    return Number.isFinite(x) && Number.isFinite(y) ? { x, y } : null;
  }

  /** Which known screen are we looking at, judged by the strings on it. */
  identifyRouteScreen(texts) {
    if (!texts || !texts.length) return null;
    const hay = texts.map((t) => String(t).toLowerCase());
    // Score rather than take the first hit: one screen's marker can be a
    // substring of another's ("PARAMÈTRES" sits inside "PARAMÈTRES DU
    // COMPTE"), and the longest, most-matched signature is the right screen.
    let best = null;
    for (const [name, def] of Object.entries(this.routeScreens)) {
      const wanted = (def.signature && def.signature.anyText) || [];
      let score = 0;
      for (const w of wanted) {
        const needle = String(w).toLowerCase();
        if (!needle) continue;
        if (hay.some((h) => h === needle)) score += needle.length * 2;   // exact line
        else if (hay.some((h) => h.includes(needle))) score += needle.length;
      }
      if (score > 0 && (!best || score > best.score)) best = { name, def, score };
    }
    return best ? { name: best.name, def: best.def } : null;
  }

  /** Taps the route already knows about, so they are not left to guesswork. */
  routeActionsFor(screenId) {
    const match = this.routeScreenFor.get(screenId);
    if (!match) return [];
    const { name, def } = match;
    const out = [];

    for (const [label, pair] of Object.entries(def.controls || {})) {
      const p = this.routePoint(pair);
      if (!p) continue;
      if (this.isBlocked(label)) {
        this.run.skipped.push({ screenId, label, reason: 'matches a blocked-label pattern' });
        continue;
      }
      out.push({
        key: `tap:${name}.${label}:${Math.round(p.x)},${Math.round(p.y)}`,
        kind: 'tap', x: p.x, y: p.y, label: `${name}.${label}`, priority: 'high', fromRoute: true,
      });
    }

    // Info badges are the whole point of recording a route: the tooltips they
    // open carry text that appears on no other screen, and a crawler working
    // from a screenshot rarely guesses that a small "?" chip is worth a tap.
    for (const [label, pair] of Object.entries(def.infoBadges || {})) {
      const p = this.routePoint(pair);
      if (!p) continue;
      out.push({
        key: `tap:${name}.badge.${label}:${Math.round(p.x)},${Math.round(p.y)}`,
        kind: 'tap', x: p.x, y: p.y, label: `${name} info: ${label}`, priority: 'high', fromRoute: true,
      });
      if (this.cfg.longPressProbe) {
        out.push({
          key: `long_press:${name}.badge.${label}:${Math.round(p.x)},${Math.round(p.y)}`,
          kind: 'long_press', x: p.x, y: p.y, label: `${name} info held: ${label}`, priority: 'high', fromRoute: true,
        });
      }
    }

    out.push(...this.routeGridActions(name, def));
    out.push(...this.routeRowDetailActions(name, def));
    return out;
  }

  /**
   * Item grids (Inventory, Abilities) put one item's name and description on
   * screen at a time and swap them when another cell is selected. Screenshotting
   * such a screen once captures a single item out of dozens, so every cell is
   * walked deliberately. Recorded by hand: two adjacent avatar cells produced
   * entirely different names and lore.
   */
  routeGridActions(name, def) {
    const out = [];
    for (const [gridName, grid] of Object.entries(def.grids || {})) {
      const cols = Array.isArray(grid.cols) ? grid.cols : [];
      const rows = Array.isArray(grid.rows) ? grid.rows : [];
      for (let r = 0; r < rows.length; r++) {
        for (let c = 0; c < cols.length; c++) {
          const p = this.routePoint([cols[c], rows[r]]);
          if (!p) continue;
          out.push({
            key: `tap:${name}.${gridName}[${r},${c}]:${Math.round(p.x)},${Math.round(p.y)}`,
            kind: 'tap', x: p.x, y: p.y,
            label: `${name} ${gridName} cell r${r + 1}c${c + 1}`,
            priority: 'high', fromRoute: true, revealsDetail: true,
          });
        }
      }
    }
    return out;
  }

  /**
   * Settings > Gameplay shows a description for the selected row only, in a
   * panel elsewhere on the screen. Every row has to be selected for its text to
   * exist anywhere on screen even once.
   */
  routeRowDetailActions(name, def) {
    const rd = def.rowDetails;
    if (!rd || !Array.isArray(rd.rows)) return [];
    const out = [];
    rd.rows.forEach((pair, i) => {
      const p = this.routePoint(pair);
      if (!p) return;
      out.push({
        key: `tap:${name}.row${i + 1}:${Math.round(p.x)},${Math.round(p.y)}`,
        kind: 'tap', x: p.x, y: p.y,
        label: `${name} option row ${i + 1}`,
        priority: 'high', fromRoute: true, revealsDetail: true,
      });
    });
    return out;
  }

  /**
   * Tag findings the team has already accepted instead of hiding them.
   *
   * Indus has one long-standing bug — switching the account language leaves a
   * fixed set of Settings keys rendering in the previous language — that would
   * otherwise dominate every run's findings. Suppressing the whole Settings
   * subtree would also bury any genuinely new defect there, so the route map
   * names the exact strings and they are marked rather than dropped: the list
   * stays complete, and the UI can collapse them out of the way.
   */
  markKnownIssues(issues) {
    const known = (this.route && this.route.knownIssues) || {};
    const rules = Object.entries(known)
      .filter(([k, v]) => !k.startsWith('$') && v && Array.isArray(v.matchAny) && v.mark === 'known')
      .map(([key, v]) => ({
        key,
        note: v.note || '',
        severity: v.severity || 'low',
        needles: v.matchAny.map((m) => String(m).trim().toLowerCase()).filter(Boolean),
      }));
    if (!rules.length) return issues;

    let tagged = 0;
    const out = issues.map((issue) => {
      const hay = String(issue.text || '').trim().toLowerCase();
      if (!hay) return issue;
      const hit = rules.find((r) => r.needles.some((n) => hay === n || hay.includes(n)));
      if (!hit) return issue;
      tagged += 1;
      return {
        ...issue,
        known: true,
        knownIssue: hit.key,
        knownNote: hit.note,
        severity: hit.severity,
      };
    });
    if (tagged) {
      this.run.log(`Marked ${tagged} finding${tagged === 1 ? '' : 's'} as a known accepted issue — still listed, but collapsed and dropped to low severity.`);
    }
    return out;
  }

  /**
   * A scan checking the wrong column reports almost every string as a defect.
   *
   * This used to warn after the fact, once the findings were already in — which
   * meant a run against the wrong language produced a page of confident
   * nonsense with an explanation buried in the log underneath it. Nothing
   * downstream can be trusted once the language is wrong, so the run stops
   * instead, before the vision pass has been paid for and before a single
   * false finding is recorded.
   */
  detectLanguageMismatch(staticIssues) {
    const wrong = staticIssues.filter((i) => i.type === 'wrong_language' && i.matchedLanguage);
    if (wrong.length < 3) return null;

    const tally = new Map();
    for (const i of wrong) tally.set(i.matchedLanguage, (tally.get(i.matchedLanguage) || 0) + 1);
    const [lang, n] = [...tally.entries()].sort((a, b) => b[1] - a[1])[0];
    if (n < 3) return null;
    return { lang, n };
  }

  /** Turn a detected mismatch into an actionable stop. */
  abortForLanguage({ lang, n }) {
    // Naming the column they should have picked is the difference between
    // "something is wrong" and "click this instead".
    const column = this.sheet.languages.find(
      (l) => l.header === lang || l.code === lang || String(l.header).toLowerCase() === String(lang).toLowerCase()
    );
    const fix = column
      ? `Set Language to "${column.header}" and scan again, or switch the game to ${this.target.header}.`
      : `Switch the game to ${this.target.header}, or add a "${lang}" column to the sheet.`;

    const msg =
      `Stopped: the game is running in "${lang}" but the scan was set to check "${this.target.header}". ` +
      `${n} strings on the first screen matched ${lang}, so every finding would be a false positive. ${fix}`;

    this.run.log(msg, 'error');
    this.run.warnings.push(msg);
    const err = new Error(msg);
    err.languageMismatch = { detected: lang, expected: this.target.header, suggestColumn: column ? column.header : '' };
    throw err;
  }

  /**
   * Does the sheet say this exact string belongs in the column under test?
   * If so, no amount of "it looks Portuguese" makes it a language defect —
   * Portuguese is what the Portuguese column is supposed to contain.
   */
  sheetAgreesWithTarget(text) {
    const raw = String(text || '').trim();
    if (!raw || !this.target || !this.target.header) return false;
    const header = this.target.header;
    const source = this.target.sourceHeader;

    const hits = this.sheet.lookupExact(raw).filter((h) => h.header !== '__key__');

    // 2. The string is what the column under test is supposed to contain.
    if (hits.some((h) => h.header === header)) return true;

    // 3. It is the source string. Whether that is a defect depends entirely on
    //    what the sheet holds for the same key — the same judgement a tester
    //    makes by looking across the row.
    const srcHit = source && hits.find((h) => h.header === source);
    if (srcHit) {
      const expected = srcHit.entry.values[header];
      // Nothing to translate to, or the translation *is* the source: the build
      // is showing exactly what the sheet says. Not a build defect.
      if (!expected || !expected.trim()) return true;
      if (norm(expected) === norm(srcHit.entry.values[source] || '')) return true;
      return false;   // a real translation exists and is not being used
    }

    // Vision transcribes from pixels, so allow for a stray character.
    const fuzzy = this.sheet.lookupFuzzy(raw, { limit: 3 }) || [];
    return fuzzy.some((f) => f.header === header && f.score >= 0.9);
  }

  /** Is the app under test the thing actually on screen right now? */
  async appIsInForeground() {
    const pkg = (this.cfg.androidPackage || '').trim();
    if (!this.adb || !pkg) return true;      // nothing to compare against
    try {
      const cur = await this.adb.currentActivity();
      return cur.package === pkg;
    } catch {
      return true;                            // cannot tell; do not block the run
    }
  }

  /**
   * Refuse to analyse anything that is not the app under test.
   *
   * Without this the crawler happily captured the launcher, the Play Store, or
   * whatever an ad had opened, and reported its strings against the sheet —
   * confident, detailed findings about the wrong application.
   */
  async ensureAppInForeground() {
    const pkg = (this.cfg.androidPackage || '').trim();
    if (!this.adb || !pkg) return true;
    // One read, then reason about it — asking twice can report a package that
    // has already changed underneath us.
    let cur = await this.adb.currentActivity().catch(() => ({ package: '' }));
    if (cur.package === pkg) return true;
    this.run.log(`${cur.package || 'Something else'} is in front instead of ${pkg} — getting back to the game.`, 'warn');

    // Back out of whatever it is, then resume the app if that was not enough.
    await this.dismissOverlays(2);
    if (await this.appIsInForeground()) return true;

    await this.adb.launch(pkg);
    await sleep(3000);
    if (await this.appIsInForeground()) return true;

    cur = await this.adb.currentActivity().catch(() => ({ package: '' }));
    this.run.log(`Could not get back to ${pkg} — ${cur.package || 'nothing'} is in front.`, 'error');
    return false;
  }

  /** The app strands itself sometimes; the route says how to get it back. */
  async maybeRecover(texts) {
    const rec = this.route && this.route.recovery;
    if (!rec || !this.adb || this.recoveries >= 2) return false;
    const wanted = (rec.detect && rec.detect.anyText) || [];
    const hay = (texts || []).map((t) => String(t).toLowerCase());
    const stuck = wanted.some((w) => hay.some((h) => h.includes(String(w).toLowerCase())));
    if (!stuck) return false;

    const pkg = this.cfg.androidPackage;
    if (!pkg) {
      this.run.warnings.push('Hit a state the route map calls stuck, but no Android package is set, so it could not restart the app.');
      return false;
    }
    this.recoveries++;
    this.run.log(`Route map recognised a stuck state — restarting ${pkg}.`, 'warn');
    this.run.warnings.push(`Recovered from a stuck state (${wanted.join(', ')}) by restarting the app.`);
    for (const step of rec.steps || []) {
      try {
        if (step.action === 'shell') await this.adb.shell(String(step.cmd).replace('{package}', pkg));
        else if (step.action === 'launch') await this.adb.shell(`monkey -p ${pkg} -c android.intent.category.LAUNCHER 1`);
        else if (step.action === 'wait') await sleep(Number(step.ms) || 1000);
      } catch (e) {
        this.run.log(`Recovery step failed: ${e.message}`, 'warn');
      }
    }
    return true;
  }

  /** Put the game into the language the sheet column is being checked against. */
  async applyRouteLanguage() {
    const proc = this.route && this.route.procedures && this.route.procedures.setLanguage;
    if (!proc || !this.adb) return false;

    const picker = this.routeScreens['settings.languagePicker'];
    const wanted = String(this.target.header || '').trim().toLowerCase();
    const control = picker && Object.keys(picker.controls || {})
      .find((k) => k.toLowerCase() === wanted || wanted.startsWith(k.toLowerCase()));
    if (!control) {
      this.run.log(`Route map has no language option matching "${this.target.header}" — leaving the game's language alone.`, 'warn');
      return false;
    }

    this.run.log(`Setting the game's language to ${this.target.header} using the route map.`);
    for (const step of proc.steps || []) {
      if (this.run.stopRequested) return false;
      const ref = String(step.tap || '').replace('{language}', control);
      const point = this.resolveRef(ref);
      if (!point) {
        this.run.log(`Route step "${ref}" does not resolve — stopping the language switch.`, 'warn');
        return false;
      }
      try {
        await this.tapAt(point.x, point.y);
      } catch (e) {
        this.run.log(`Route step "${ref}" failed: ${e.message}`, 'warn');
        return false;
      }
      await sleep(Number(step.wait) || this.cfg.settleMs);
    }
    this.run.log('Language switch done.');
    return true;
  }

  /** "settings.account.languageChange" -> a point on screen. */
  resolveRef(ref) {
    const parts = String(ref).split('.');
    if (parts.length < 2) return null;
    const key = parts.pop();
    const screen = parts.join('.');
    const def = this.routeScreens[screen];
    if (!def) return null;
    const pair = (def.controls && def.controls[key]) || (def.infoBadges && def.infoBadges[key]);
    return pair ? this.routePoint(pair) : null;
  }

  async perform(action) {
    this.actionCount++;
    this.run.emit('action', { action: { kind: action.kind, label: action.label }, count: this.actionCount });

    switch (action.kind) {
      case 'click':
        if (this.useBridge && action.id != null) return this.bridge.click(action.id);
        return this.tapAt(action.x, action.y);
      case 'tap':
        return this.tapAt(action.x, action.y);
      case 'long_press':
        if (this.useBridge && action.x != null) return this.bridge.longPress(action.x, action.y, this.cfg.longPressMs);
        if (this.adb) return this.adb.longPress(action.x, action.y, this.cfg.longPressMs);
        return this.bridge.click(action.id);
      case 'swipe_up': {
        const s = this.deviceSize || { width: 1080, height: 1920 };
        return this.adb
          ? this.adb.swipe(action.x, action.y + s.height * 0.2, action.x, action.y - s.height * 0.2, 320)
          : null;
      }
      case 'swipe_left': {
        const s = this.deviceSize || { width: 1080, height: 1920 };
        return this.adb
          ? this.adb.swipe(action.x + s.width * 0.25, action.y, action.x - s.width * 0.25, action.y, 320)
          : null;
      }
      case 'scroll':
        return this.bridge.scroll(action.id, action.position);
      case 'back':
        return this.goBack();
      default:
        return null;
    }
  }

  async tapAt(x, y) {
    if (this.useBridge) return this.bridge.tap(x, y);
    if (this.adb) return this.adb.tap(x, y);
    throw new Error('no input transport available');
  }

  async goBack() {
    if (this.useBridge) {
      try { return await this.bridge.back(); } catch { /* fall through to the OS back key */ }
    }
    if (this.adb) return this.adb.back();
    return null;
  }

  // ── navigation ─────────────────────────────────────────────────────────

  async currentSig() {
    const cap = await this.capture();
    return { sig: this.signature(cap), cap };
  }

  /**
   * Games open with things in the way: an interstitial ad, a daily reward, a
   * rate-us prompt. Back usually clears them, and doing so is much cheaper than
   * a restart — and unlike a restart it does not put a *fresh* ad on screen.
   */
  async dismissOverlays(max = 3) {
    if (!this.adb) return false;
    // Past runs may already know what clears this app's launch interstitial.
    if (this.mem && memory.bestDismissal(this.mem, 'launch_interruption') === 'back') max = Math.max(max, 2);
    for (let i = 0; i < max; i++) {
      const pkg = (this.cfg.androidPackage || '').trim();
      if (pkg) {
        // An ad can hand control to a browser or another activity entirely.
        const cur = await this.adb.currentActivity().catch(() => ({ package: '' }));
        if (cur.package && cur.package !== pkg) {
          this.run.log(`${cur.package} is in front of the game — backing out of it.`, 'warn');
        }
      }
      await this.goBack();
      await sleep(this.cfg.settleMs);
      const { sig } = await this.currentSig();
      if (sig === this.rootSig) {
        if (this.mem) {
          memory.rememberObstacle(this.mem, {
            kind: 'launch_interruption',
            hint: `cleared after ${i + 1} back press${i ? 'es' : ''}`,
            dismissal: 'back',
          });
        }
        return true;
      }
    }
    return false;
  }

  async resetToRoot() {
    // 1. cheap: back out until we recognise the root
    for (let i = 0; i < 6; i++) {
      const { sig } = await this.currentSig();
      if (sig === this.rootSig) return true;
      await this.goBack();
      await sleep(this.cfg.settleMs);
    }

    // 2. still lost — try clearing whatever is on top before reaching for the
    //    hammer. A launch-time ad is the usual reason Back alone did not work.
    if (await this.dismissOverlays()) return true;

    // 3. bring the app forward without killing it. A LAUNCHER intent resumes
    //    the existing task, so state, session and position survive — and it
    //    does not trigger a fresh launch ad the way a cold start does.
    if (this.adb && this.cfg.androidPackage) {
      await this.adb.launch(this.cfg.androidPackage);
      await sleep(1500);
      await this.dismissOverlays(1);
      const { sig: resumed } = await this.currentSig();
      if (resumed === this.rootSig) return true;
    }

    // 4. last resort: a cold restart. Kept behind the gentler attempt above
    //    because force-stopping is what "it closed my game" looks like, and it
    //    is capped because a fresh interstitial on every launch would
    //    otherwise loop forever.
    if (this.adb && this.cfg.androidPackage && this.restarts < this.maxRestarts) {
      this.restarts += 1;
      this.run.log(`Restarting ${this.cfg.androidPackage} to get back to the first screen (${this.restarts}/${this.maxRestarts}).`);
      await this.adb.forceStop(this.cfg.androidPackage);
      await sleep(1200);
      await this.adb.launch(this.cfg.androidPackage);
      await sleep(4000);
      // A cold start can fail outright — the app crashes on boot, or the
      // launcher never hands over. Continuing here is what led to scanning
      // the Play Store, so confirm the app is actually in front.
      if (!(await this.appIsInForeground())) {
        this.run.log(`${this.cfg.androidPackage} did not come back after the restart.`, 'error');
        return false;
      }
      await this.settle();

      let { sig } = await this.currentSig();
      if (sig === this.rootSig) return true;

      // A launch-time ad lands here every time. Clear it before concluding
      // that the app simply starts somewhere new.
      if (await this.dismissOverlays(2)) return true;
      ({ sig } = await this.currentSig());
      if (sig === this.rootSig) return true;

      // Genuinely a different first screen (a daily popup, an A/B variant).
      this.run.log('The app started on a different screen; treating that as the new starting point.', 'warn');
      this.rootSig = sig;
      return true;
    }

    if (this.restarts >= this.maxRestarts) {
      this.run.log(
        `Restarted ${this.restarts} times without getting back to a known screen — something keeps appearing on launch ` +
        '(an ad or a daily popup). Carrying on from the current screen instead of restarting again.',
        'warn'
      );
    }
    return false;
  }

  /** Gets the app into `state`, replaying from the root if we have drifted. */
  async navigateTo(node) {
    const { sig } = await this.currentSig();
    if (sig === node.sig) return true;

    // one step up is the common case when finishing a leaf
    for (let i = 0; i < 2; i++) {
      await this.goBack();
      await sleep(this.cfg.settleMs);
      const now = await this.currentSig();
      if (now.sig === node.sig) return true;
    }

    if (!(await this.resetToRoot())) {
      this.run.warnings.push('Could not return to the first screen; some paths may be skipped. Set the Android package name to allow app restarts.');
      return false;
    }
    for (const action of node.path) {
      await this.perform(action);
      await this.settle();
    }
    const final = await this.currentSig();
    if (final.sig !== node.sig) {
      this.run.log(`Replaying the path to ${node.screenId} landed somewhere else; continuing from here.`, 'warn');
      return false;
    }
    return true;
  }

  // ── analysis ───────────────────────────────────────────────────────────

  matchesFor(text) {
    const exact = this.sheet.lookupExact(text).filter((h) => h.header !== '__key__');
    if (exact.length) {
      return exact.slice(0, 3).map((h) => ({
        key: h.entry.key, row: h.entry.rowNumber, header: h.header, value: h.entry.values[h.header], score: null,
      }));
    }
    return this.sheet.lookupFuzzy(text, { limit: 3 }).map((h) => ({
      key: h.entry.key, row: h.entry.rowNumber, header: h.header, value: h.entry.values[h.header], score: h.score,
    }));
  }

  async analyzeScreen(cap, screenId, depth, pathLabels) {
    // Never spend a vision call, or file findings, against something that is
    // not the app under test.
    if (!(await this.ensureAppInForeground())) {
      this.wrongAppSkips = (this.wrongAppSkips || 0) + 1;
      this.run.log(`Skipped ${screenId}: the app under test was not on screen.`, 'warn');
      if (this.wrongAppSkips >= 3) {
        throw new Error(
          `Stopped: ${this.cfg.androidPackage} kept leaving the foreground and could not be brought back. ` +
          'Check the build launches and stays open on the device, then scan again.'
        );
      }
      return;
    }
    this.wrongAppSkips = 0;

    const file = this.run.saveScreenshot(screenId, cap.png);

    const ctx = {
      screenId,
      scene: cap.state ? cap.state.scene : '',
      targetHeader: this.target.header,
      targetCode: this.target.code,
      rtl: this.target.rtl,
      sourceHeader: this.target.sourceHeader,
      mode: this.mode,
    };

    // 1. deterministic pass
    let staticIssues = [];
    if (cap.state && Array.isArray(cap.state.texts)) {
      staticIssues = runChecks(cap.state, this.sheet, {
        targetHeader: this.target.header,
        englishHeader: this.target.sourceHeader,
        rtl: this.target.rtl,
        expectsNonLatin: this.target.expectsNonLatin,
      });
    }

    // The deterministic pass has already matched strings against every column,
    // so it knows what language is actually on screen. Ask it before spending
    // a vision call, and before any of this screen's findings are kept.
    if (!this.langMismatchChecked) {
      this.langMismatchChecked = true;
      const mismatch = this.detectLanguageMismatch(staticIssues);
      if (mismatch) this.abortForLanguage(mismatch);
    }

    // 2. sheet context for the vision pass
    const extracted = (cap.state && cap.state.texts ? cap.state.texts : [])
      .filter((t) => t.active !== false && String(t.text || '').trim())
      .slice(0, 120)
      .map((t) => ({ path: t.path, text: t.text, rect: t.rect, matches: this.matchesFor(t.text) }));

    // 3. vision pass
    let vision = { issues: [], unlisted_text: [], screen_summary: '' };
    if (this.cfg.visionEnabled) {
      try {
        vision = await this.analyzer.analyzeScreen(cap.png, ctx, extracted, staticIssues);
      } catch (e) {
        this.run.log(`Vision pass failed on ${screenId}: ${e.message}`, 'warn');
        this.run.warnings.push(`Vision pass failed on ${screenId}: ${e.message}`);
      }
    }

    // Remember what this screen said. The route map is matched against it, and
    // a stuck state is recognised the same way.
    const seen = [
      ...extracted.map((t) => t.text),
      ...((vision.unlisted_text || []).map((t) => (typeof t === 'string' ? t : t && t.text))),
    ].filter(Boolean);
    this.observedText.set(screenId, seen);
    const known = this.identifyRouteScreen(seen);
    if (known) {
      this.routeScreenFor.set(screenId, known);
      this.run.log(`Route map recognises ${screenId} as "${known.name}".`);
    }
    await this.maybeRecover(seen);

    // 4. strings only the model could see get the same sheet treatment
    const aiStatic = [];
    if (vision.unlisted_text && vision.unlisted_text.length) {
      const pseudo = {
        screen: cap.state ? cap.state.screen : null,
        texts: vision.unlisted_text.map((u, i) => ({
          id: `ai-${i}`, path: `(screenshot) ${u.where}`, text: u.text, rect: null, active: true,
        })),
      };
      for (const issue of runChecks(pseudo, this.sheet, {
        targetHeader: this.target.header,
        englishHeader: this.target.sourceHeader,
        rtl: this.target.rtl,
        expectsNonLatin: this.target.expectsNonLatin,
      })) {
        issue.source = 'static-ocr';
        aiStatic.push(issue);
      }
    }

    let issues = [
      ...staticIssues,
      ...aiStatic,
      ...(vision.issues || []).map((i) => ({
        source: 'vision',
        type: i.type,
        severity: i.severity,
        confidence: i.confidence,
        element: i.element || i.where || '',
        text: i.text,
        where: i.where,
        key: i.key || '',
        expected: i.expected || '',
        message: i.message,
      })),
    ].map((i) => ({ ...i, screenId, screenFile: file, scene: ctx.scene, depth, path: pathLabels }));

    // The vision pass judges by eye. Without the bridge it is handed no sheet
    // data at all, so its language verdicts are guesses about a book it has not
    // read — and it will call a legitimate target-language string "the wrong
    // language" because it looks foreign. The sheet is the authority, so every
    // such verdict is checked against it before being kept.
    const verified = [];
    let overruled = 0;
    for (const i of issues) {
      if (i.source === 'vision' && LANGUAGE_CLAIMS.has(i.type) && this.sheetAgreesWithTarget(i.text)) {
        overruled += 1;
        continue;
      }
      verified.push(i);
    }
    if (overruled) {
      this.run.log(`Dropped ${overruled} vision finding${overruled === 1 ? '' : 's'} that the sheet contradicts — the text is in the "${this.target.header}" column.`);
    }
    issues = verified;

    let deduped = dedupe(issues);
    deduped = this.markKnownIssues(deduped);
    if (this.mem) {
      const { kept, removed } = memory.filterDismissed(this.mem, deduped);
      if (removed) this.run.log(`Ignored ${removed} finding${removed === 1 ? '' : 's'} you dismissed on an earlier scan.`);
      deduped = kept;
    }
    this.run.addIssues(deduped);
    this.run.addScreen({
      id: screenId,
      file,
      scene: ctx.scene,
      depth,
      path: pathLabels,
      summary: vision.screen_summary || '',
      textCount: extracted.length,
      issueCount: deduped.length,
      unlisted: (vision.unlisted_text || []).map((u) => u.text),
    });

    return deduped;
  }

  /**
   * Text below the fold is invisible to a single screenshot, so every screen
   * gets scrolled through before we move on. With the bridge we drive each
   * ScrollRect exactly; without it we swipe and watch for the screen to stop
   * changing, which is what "reached the bottom" looks like from outside.
   */
  async probeScrolls(cap, baseId, depth, pathLabels) {
    if (!this.cfg.scrollProbe) return;
    if (this.useBridge && cap.state && Array.isArray(cap.state.scrolls)) {
      return this.probeBridgeScrolls(cap, baseId, depth, pathLabels);
    }
    return this.probeSwipeScrolls(cap, baseId, depth, pathLabels);
  }

  /** Vision mode: swipe down the screen until nothing new appears. */
  async probeSwipeScrolls(cap, baseId, depth, pathLabels) {
    if (!this.adb) return;
    const size = this.deviceSize || { width: 1080, height: 1920 };
    const x = Math.round(size.width / 2);
    const from = Math.round(size.height * 0.72);
    const to = Math.round(size.height * 0.28);

    let prevSig = this.signature(cap);
    let steps = 0;

    for (let step = 1; step <= this.cfg.scrollSteps; step++) {
      if (this.run.stopRequested) break;
      try {
        await this.adb.swipe(x, from, x, to, 400);
      } catch (e) {
        this.run.log(`Could not swipe on ${baseId}: ${e.message}`, 'warn');
        break;
      }
      const next = await this.settle();
      const sig = this.signature(next);

      // Nothing moved — this screen does not scroll, or we hit the bottom.
      if (sig === prevSig) break;
      const near = this.matchExistingHash(next.hash);
      if (near && !next.state) { prevSig = sig; steps++; continue; }
      prevSig = sig;
      steps++;

      if (this.visited.has(sig)) continue;
      const id = `${baseId}-scroll${step}`;
      this.visited.set(sig, id);
      if (next.hash) this.hashes.push({ hash: next.hash, sig });
      this.screenCount++;
      this.run.log(`Scrolled ${baseId} down (${step}) — capturing ${id}`);
      await this.analyzeScreen(next, id, depth, [...pathLabels, `scroll down ${step}`]);
    }

    // Put the screen back where we found it so the parent state still matches.
    for (let i = 0; i < steps; i++) {
      try { await this.adb.swipe(x, to, x, from, 400); } catch { break; }
    }
    if (steps) await sleep(this.cfg.settleMs);
  }

  /** Bridge mode: step each ScrollRect through its full range. */
  async probeBridgeScrolls(cap, baseId, depth, pathLabels) {
    const scrolls = (cap.state && cap.state.scrolls) || [];
    for (const s of scrolls) {
      if (!s.canScroll) continue;
      for (let step = 1; step <= this.cfg.scrollSteps; step++) {
        if (this.run.stopRequested) return;
        const pos = step / this.cfg.scrollSteps;
        try {
          await this.bridge.scroll(s.id, s.vertical ? 1 - pos : pos);
        } catch (e) {
          this.run.log(`Scroll failed on ${s.path}: ${e.message}`, 'warn');
          break;
        }
        const next = await this.settle();
        const sig = this.signature(next);
        if (this.visited.has(sig)) continue;
        const id = `${baseId}-scroll${step}`;
        this.visited.set(sig, id);
        if (next.hash) this.hashes.push({ hash: next.hash, sig });
        this.screenCount++;
        this.run.log(`Scrolled ${s.path} to ${Math.round(pos * 100)}% — capturing ${id}`);
        await this.analyzeScreen(next, id, depth, [...pathLabels, `scroll ${s.path} ${Math.round(pos * 100)}%`]);
      }
      try { await this.bridge.scroll(s.id, s.vertical ? 1 : 0); } catch { /* best effort reset */ }
      await sleep(300);
    }
  }

  // ── main loop ──────────────────────────────────────────────────────────

  /**
   * Bring the app under test to the front before anything is captured.
   *
   * Deliberately does not force-stop when it is already running: a tester may
   * have signed in, picked a save slot or dismissed a first-run flow, and
   * throwing that away would make the scan harder to set up, not easier. It
   * only launches when something else is in front.
   *
   * With no package configured there is nothing to open, and the run will
   * scan whatever is on screen — which is legitimate (the Unity editor, a
   * build already running) but is worth saying out loud, because silently
   * scanning the launcher looks like a broken scan.
   */
  async openAppUnderTest() {
    const pkg = (this.cfg.androidPackage || '').trim();
    let current = { package: '', activity: '' };
    try { current = await this.adb.currentActivity(); } catch { /* asked below anyway */ }

    if (!pkg) {
      this.run.log(
        `No Android package set — scanning whatever is on screen right now (${current.package || 'unknown app'}). ` +
        'Set "Android package" to have the scan open your build itself.',
        'warn'
      );
      return;
    }

    if (current.package === pkg) {
      this.run.log(`${pkg} is already in the foreground.`);
      return;
    }

    this.run.log(`Foreground app is ${current.package || 'unknown'} — launching ${pkg}.`);
    try {
      await this.adb.launch(pkg);
      await sleep(4000);
      const now = await this.adb.currentActivity();
      if (now.package === pkg) {
        this.run.log(`${pkg} is up (${now.activity}).`);
      } else {
        // Better a loud warning than a page of findings about the wrong app.
        this.run.log(
          `Asked Android to open ${pkg} but ${now.package || 'nothing'} is in front. ` +
          'Is the package name right, and is that build installed?',
          'error'
        );
      }
    } catch (e) {
      this.run.log(`Could not launch ${pkg}: ${e.message}`, 'error');
    }
  }

  async start() {
    this.run.status = 'running';
    this.run.emit('status', { status: 'running', mode: this.mode });
    this.run.log(`Capture mode: ${this.mode}`);

    if (this.adb) {
      try {
        const reported = await this.adb.screenSize();
        if (!this.deviceSize) this.deviceSize = reported;
        this.run.log(`Device reports ${reported.width}x${reported.height}; taps use the screenshot's own size.`);
      } catch (e) {
        this.run.log(`Could not read the screen size: ${e.message}`, 'warn');
      }
    }

    // Open the app the run is about.
    //
    // This used to be missing entirely: androidPackage was consulted only when
    // backtracking got lost, so a scan simply captured whatever happened to be
    // in the foreground — the launcher, or the Play Store — and dutifully
    // reported its strings against the sheet. If a package is named, it is the
    // subject of the scan and has to be on screen before the first capture.
    if (this.adb) {
      await this.openAppUnderTest();
    }

    // The app may open straight into an ad. Analysing that as screen-001 makes
    // the ad the root, and then every later reset tries to return to it.
    if (this.adb) {
      const { sig } = await this.currentSig();
      this.rootSig = sig;                 // so dismissOverlays has something to compare
      await this.dismissOverlays(2);
      this.rootSig = null;
    }

    if (this.route && this.cfg.routeSetLanguage !== false) {
      try {
        await this.applyRouteLanguage();
      } catch (e) {
        this.run.log(`Could not apply the route's language switch: ${e.message}`, 'warn');
      }
    }

    const first = await this.settle();
    const rootSig = this.signature(first);
    this.rootSig = rootSig;
    const rootId = 'screen-001';
    this.visited.set(rootSig, rootId);
    if (first.hash) this.hashes.push({ hash: first.hash, sig: rootSig });
    this.screenCount = 1;

    this.run.log('Analysing the first screen.');
    await this.analyzeScreen(first, rootId, 0, []);
    await this.probeScrolls(first, rootId, 0, []);

    const rootNode = { sig: rootSig, screenId: rootId, depth: 0, path: [] };
    const rootActions = await this.actionsFor(first, rootId);
    this.run.log(`${rootActions.length} controls to explore on the first screen.`);
    for (const a of rootActions.slice().reverse()) this.stack.push({ action: a, parent: rootNode });

    while (this.stack.length) {
      if (this.run.stopRequested) {
        this.run.log('Stop requested — wrapping up.');
        break;
      }
      if (this.screenCount >= this.cfg.maxScreens) {
        this.run.warnings.push(`Stopped at the ${this.cfg.maxScreens}-screen limit; raise "Max screens" to go deeper.`);
        break;
      }
      if (this.actionCount >= this.cfg.maxActions) {
        this.run.warnings.push(`Stopped at the ${this.cfg.maxActions}-action limit.`);
        break;
      }

      const task = this.stack.pop();
      this.run.emit('progress', {
        screens: this.screenCount,
        actions: this.actionCount,
        queued: this.stack.length,
        issues: this.run.issues.length,
        usage: this.analyzer ? this.analyzer.usage : null,
      });

      const onCourse = await this.navigateTo(task.parent);
      if (!onCourse) {
        // We are somewhere unexpected. Re-anchor rather than blindly tapping.
        const here = await this.currentSig();
        if (!this.visited.has(here.sig)) {
          const id = `screen-${String(++this.screenCount).padStart(3, '0')}`;
          this.visited.set(here.sig, id);
          if (here.cap.hash) this.hashes.push({ hash: here.cap.hash, sig: here.sig });
          await this.analyzeScreen(here.cap, id, task.parent.depth, task.parent.path.map((a) => a.label));
        }
        continue;
      }

      let acted = true;
      const tried = this.triedLabels.get(task.parent.screenId) || [];
      if (task.action.label) tried.push(task.action.label);
      this.triedLabels.set(task.parent.screenId, tried);
      try {
        await this.perform(task.action);
      } catch (e) {
        acted = false;
        this.run.log(`Action "${task.action.label}" failed: ${e.message}`, 'warn');
      }
      if (!acted) continue;

      const cap = await this.settle();
      const sig = this.signature(cap);

      if (sig === task.parent.sig) continue;               // nothing happened
      if (this.visited.has(sig)) continue;                 // already seen this state
      const near = this.matchExistingHash(cap.hash);
      if (near && !cap.state) continue;                    // vision mode: visually identical

      const depth = task.parent.depth + 1;
      const id = `screen-${String(++this.screenCount).padStart(3, '0')}`;
      this.visited.set(sig, id);
      if (cap.hash) this.hashes.push({ hash: cap.hash, sig });

      const path = [...task.parent.path, task.action];
      const pathLabels = path.map((a) => a.label);
      this.run.log(`New screen ${id} via "${task.action.label}" (depth ${depth}).`);

      await this.analyzeScreen(cap, id, depth, pathLabels);
      await this.probeScrolls(cap, id, depth, pathLabels);

      if (depth < this.cfg.maxDepth) {
        const node = { sig, screenId: id, depth, path };
        const next = await this.actionsFor(cap, id);
        for (const a of next.slice().reverse()) this.stack.push({ action: a, parent: node });
      }
    }

    this.run.emit('progress', {
      screens: this.screenCount,
      actions: this.actionCount,
      queued: this.stack.length,
      issues: this.run.issues.length,
      usage: this.analyzer ? this.analyzer.usage : null,
    });
    return { screens: this.screenCount, actions: this.actionCount };
  }
}

/** Collapses repeats of the same defect on the same element. */
function dedupe(issues) {
  const seen = new Map();
  for (const i of issues) {
    const k = `${i.type}|${norm(i.text)}|${i.element || i.where || ''}`;
    const prev = seen.get(k);
    if (!prev) {
      seen.set(k, i);
      continue;
    }
    // keep the richer record: static findings carry keys and rects
    if (prev.source === 'vision' && i.source !== 'vision') seen.set(k, { ...i, alsoSeenBy: 'vision' });
    else if (prev.source !== 'vision' && i.source === 'vision') seen.set(k, { ...prev, alsoSeenBy: 'vision' });
  }
  return [...seen.values()];
}

module.exports = { Crawler };
