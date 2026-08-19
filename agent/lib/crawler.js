'use strict';

const { runChecks } = require('./checks');
const { perceptualHash, hammingDistance, size: pngSize } = require('./png');
const { norm } = require('./sheet');
const memory = require('./memory');
const { cropAndZoom } = require('./zoom');
const baseline = require('./baseline');

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
 * The route map's own blocked list, as anchored patterns.
 *
 * These are control refs — `lobby.play`, `exitConfirm.yes`, `store.*.purchase`
 * — not the free-text labels a tester types in the panel, so they are matched
 * whole rather than as substrings and `*` is the only wildcard.
 *
 * They were being recorded and then ignored: the crawler built its blocked list
 * from the config alone, so every entry a route map author wrote to keep a scan
 * out of a live match, an exit dialog, a settings reset or a purchase confirm
 * did nothing at all. The self-test missed it because the test wired the two
 * together itself, which is exactly the part production was missing.
 */
function routeBlockedPatterns(route) {
  const labels = (route && route.blocked && route.blocked.labels) || [];
  return labels
    .filter((l) => typeof l === 'string' && l.trim())
    .map((l) => `^${l.trim().replace(/[.+?^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*')}$`);
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

    this.blocked = safeRegexes([...(cfg.blockedLabels || []), ...routeBlockedPatterns(route)]);
    this.probe = safeRegexes(cfg.probeLabels);
    // Steering. `focus` reorders the queue, `only` narrows it to the part of
    // the app the tester actually wants this run spent on.
    this.focus = safeRegexes(cfg.focusLabels);
    this.only = safeRegexes(cfg.onlyLabels);
    this.rules = cfg.compiledRules || [];
    this.steps = cfg.compiledSteps || [];
    this.deadline = null;              // set in start() from cfg.maxMinutes
    this.highIssues = 0;

    // An overflow is only reportable as a pair: the same screen in the source
    // language and in the target. A run covers one language, so the source
    // pass writes its screens to disk and every later pass reads them back.
    this.routeName = (route && route.app && route.app.name) || (target && target.routeName) || '';
    this.sourceLanguage = (target && target.sourceHeader) || '';
    this.capturingBaseline = !!(this.sourceLanguage && target
      && String(target.header).trim().toLowerCase() === String(this.sourceLanguage).trim().toLowerCase());
    this.baselineEnabled = !!(this.routeName && this.sourceLanguage && cfg.englishBaseline !== false);

    this.visited = new Map();          // sig -> screenId
    this.autoDismissScreens = new Set(); // screens the route says to close, not explore
    this.clearingModals = false;       // guards the dismissal's own settles
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
    const cap = await this.settleOnly();
    // A modal the route map says to dismiss on sight is not a screen to
    // explore — it is something standing between us and the screen we asked
    // for, and every coordinate underneath it is wrong while it is up.
    const cleared = await this.clearAutoDismissModals(cap);
    return cleared || cap;
  }

  /** The raw settle, with no modal handling — used by the dismissal itself. */
  async settleOnly() {
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

  /**
   * Tap a recorded control, then check it actually did something.
   *
   * A recorded coordinate is a guess about a layout, and a blind sequence of
   * them fails in the worst possible way: every tap lands somewhere, nothing
   * errors, and the procedure reports success having navigated nowhere. That is
   * how the language switch spent two minutes on the lobby and then announced
   * it could not find GERMAN in a picker it had never opened.
   *
   * So a step may name what should be on screen afterwards (`expect`). If it is
   * not, the model is asked to find that control on the current screen and the
   * tap is retried where it actually is. Recorded coordinates stay the fast
   * path; looking is the fallback that keeps a shifted layout, a stale
   * recording or an unexpected overlay from silently voiding the whole run.
   */
  async tapRouteStep(step) {
    const ref = String(step.tap);
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
    if (!step.expect) return true;

    await sleep(Number(step.wait) || this.cfg.settleMs);
    const cap = await this.settleOnly();
    const size = this.deviceSize || { width: 1080, height: 1920 };
    let found = null;
    try {
      found = await this.analyzer.locateText(cap.png, String(step.expect));
    } catch (e) {
      this.run.log(`Could not confirm "${ref}": ${e.message}`, 'warn');
      return true;                       // an unreadable screen is not proof of failure
    }
    if (found) return true;

    // Before deciding the coordinate is wrong, rule out the two things that
    // make a perfectly good coordinate miss: something covering it, and an app
    // that has not finished launching. Both are ordinary here — the promo
    // interstitial is on screen more often than not, and a cold start takes the
    // better part of a minute.
    this.run.log(`"${ref}" did not open ${step.expect} — clearing anything on top and trying once more.`);
    await this.clearAutoDismissModals(cap);
    await sleep(4000);
    try {
      await this.tapAt(point.x, point.y);
      await sleep(Number(step.wait) || this.cfg.settleMs);
      const retry = await this.settleOnly();
      if (await this.analyzer.locateText(retry.png, String(step.expect))) {
        this.run.log(`"${ref}" worked on the second try.`);
        return true;
      }
    } catch { /* fall through to looking for it */ }

    // Still not there. Look for the control by name and use wherever it is —
    // only useful when the control carries text; an icon has nothing to match.
    if (!step.find) {
      this.run.log(`"${ref}" still did not open ${step.expect} — stopping the language switch.`, 'warn');
      return false;
    }
    this.run.log(`Looking for "${step.find}" on screen instead.`);
    let target = null;
    try {
      const now = await this.settleOnly();
      target = await this.analyzer.locateText(now.png, String(step.find));
    } catch { /* fall through to the failure below */ }
    if (!target) {
      this.run.log(`Could not find "${step.find}" on screen — stopping the language switch.`, 'warn');
      return false;
    }
    await this.tapAt(
      Math.max(0, Math.min(1, target.x)) * size.width,
      Math.max(0, Math.min(1, target.y)) * size.height
    );
    this.run.log(`Recovered: tapped "${step.find}" where it actually is.`);
    return true;
  }

  /**
   * Name the current screen against the whole route map, by looking at it.
   *
   * The fallback when the recorded signatures do not match, which on a build
   * being scanned in a new language is most of the time — the samples were
   * written in the languages someone happened to record. Getting a name back
   * is what turns the map from a list of coordinates into the veto, the shared
   * vocabulary and the "this screen does not scroll" it is actually for.
   */
  async identifyScreenByVision(cap) {
    if (!this.analyzer || !cap || !cap.png) return null;
    const candidates = Object.entries(this.routeScreens)
      .map(([name, def]) => ({ name, anyText: (def.signature && def.signature.anyText) || [] }))
      .filter((c) => c.anyText.length);
    if (!candidates.length) return null;

    let name = null;
    try {
      name = await this.analyzer.identifyScreen(cap.png, candidates);
    } catch (e) {
      this.run.log(`Could not identify the screen: ${e.message}`, 'warn');
      return null;
    }
    if (!name || !this.routeScreens[name]) return null;
    this.run.log(`Recognised this screen as "${name}" by looking — its recorded text did not match.`);
    return { name, def: this.routeScreens[name] };
  }

  /**
   * Which dismiss-on-sight screen is in front of us, decided by looking.
   *
   * Restricted to the screens the route map marks autoDismiss, so the question
   * put to the model is small and closed: "is this one of these three modals,
   * or not?" — not "what am I looking at?". The answer is cached per run
   * signature so a modal that reappears between screens is not re-identified
   * from scratch every time.
   */
  async identifyOverlayByVision(png) {
    const candidates = Object.entries(this.routeScreens)
      .filter(([, def]) => def && def.autoDismiss && def.autoDismiss.tap)
      .map(([name, def]) => ({ name, anyText: (def.signature && def.signature.anyText) || [] }))
      .filter((c) => c.anyText.length);
    if (!candidates.length) return null;

    let name = null;
    try {
      name = await this.analyzer.identifyScreen(png, candidates);
    } catch (e) {
      this.run.log(`Could not check for an overlay: ${e.message}`, 'warn');
      return null;
    }
    if (!name || !this.routeScreens[name]) return null;
    return { name, def: this.routeScreens[name] };
  }

  /**
   * Close the modals the route map marks `autoDismiss`, using their own close
   * control rather than the back key.
   *
   * Two things made this necessary, both hit by hand while recording Indus.
   * A daily-reward modal appears over the lobby after launch *and after every
   * language change*, and it swallowed three separate navigation sequences:
   * the taps landed on the modal, the crawl believed it had reached Settings,
   * and the screenshots that followed were of the wrong screen — captured,
   * analysed and reported with complete confidence. And the obvious way to
   * clear such a thing, pressing back, is the one thing that must not happen
   * here: back on the lobby opens the exit-game confirmation, whose yes button
   * quits the app and ends the run.
   *
   * So the route map names the control that closes each modal, and this taps
   * exactly that.
   */
  async clearAutoDismissModals(cap, max = 3) {
    if (!this.route || this.clearingModals) return null;
    this.clearingModals = true;
    try {
      let current = cap;
      let cleared = null;
      for (let i = 0; i < max; i++) {
        const texts = this.textsOf(current);
        let known = this.identifyRouteScreen(texts);

        // Without the bridge there are no strings yet, so the match above is
        // made against an empty array and always fails. Ask the model instead —
        // one small call, against only the screens the map says to close on
        // sight, rather than letting an overlay swallow the next four taps.
        if (!known && !texts.length && this.analyzer && current && current.png) {
          known = await this.identifyOverlayByVision(current.png);
        }

        const auto = known && known.def && known.def.autoDismiss;
        if (!auto || !auto.tap) break;

        const point = this.resolveRef(auto.tap);
        if (!point) {
          this.run.log(`${known.name} should be dismissed but "${auto.tap}" does not resolve — leaving it up.`, 'warn');
          break;
        }
        this.run.log(`${known.name} is covering the screen — closing it with ${auto.tap}.`);
        try {
          await this.tapAt(point.x, point.y);
        } catch (e) {
          this.run.log(`Could not close ${known.name}: ${e.message}`, 'warn');
          break;
        }
        current = await this.settleOnly();
        cleared = current;
      }
      return cleared;
    } finally {
      this.clearingModals = false;
    }
  }

  /**
   * Does this route map say the back key is unsafe here?
   *
   * Indus answers yes: back on the lobby opens a confirmation whose yes button
   * quits the game. Rather than hard-code that, any hazard whose text says so
   * is enough to stop the generic back-until-it-goes-away routine.
   */
  hazardWarnsAgainstBack() {
    const hazards = (this.route && this.route.hazards) || {};
    return Object.entries(hazards).some(([name, text]) =>
      /back/i.test(`${name} ${text}`) && /unsafe|exit|quit|kills the app|ends the run/i.test(String(text)));
  }

  /** Every string we can see on a capture, from the bridge or from memory. */
  textsOf(cap) {
    if (cap && cap.state && Array.isArray(cap.state.texts)) {
      return cap.state.texts
        .filter((t) => t.active !== false && String(t.text || '').trim())
        .map((t) => t.text);
    }
    // Vision mode has no strings until the analyser has looked at the screen.
    // The signature is all we have, so fall back to whatever this exact screen
    // said last time we were here.
    const sig = cap ? this.signature(cap) : null;
    const id = sig ? this.visited.get(sig) : null;
    return (id && this.observedText.get(id)) || [];
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

  isFocused(label) {
    const l = String(label || '').trim();
    return !!l && this.focus.some((re) => re.test(l));
  }

  /** With an "only tap" list set, everything outside it is off the table. */
  isAllowed(label) {
    if (!this.only.length) return true;
    const l = String(label || '').trim();
    return !!l && this.only.some((re) => re.test(l));
  }

  /**
   * Puts the tester's focus patterns at the front of the queue and drops what
   * an "only tap" list excludes. Route-map controls keep their own precedence;
   * this only sorts within what is left.
   */
  steer(actions, screenId) {
    let out = actions;
    if (this.only.length) {
      out = [];
      for (const a of actions) {
        if (a.fromRoute || this.isAllowed(a.label)) out.push(a);
        else this.run.skipped.push({ screenId, label: a.label, reason: 'outside the "only tap" list' });
      }
    }
    if (!this.focus.length) return out;
    const rank = (a) => (a.fromRoute ? 0 : this.isFocused(a.label) ? 1 : 2);
    return out.slice().sort((a, b) => rank(a) - rank(b));
  }

  /** Enumerates the actions available on the current screen. */
  async actionsFor(cap, screenId) {
    // A dismiss-on-sight modal has no controls worth queueing: exploring it
    // means tapping CLAIM, or the exit-game dialog's yes button. Close it and
    // let the crawl carry on with whatever it was covering.
    if (this.autoDismissScreens.has(screenId)) {
      await this.clearAutoDismissModals(cap);
      return [];
    }

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

    return this.steer(this.applyRouteToProposals(actions, screenId), screenId);
  }

  /**
   * The route map's say over what the model decided to do.
   *
   * It used to be the other way round: recorded controls went first and any
   * proposal near one was thrown away, so the crawl was really following a
   * hand-drawn map with the model filling gaps. That inverts here. The model
   * looks at the screen and says what is worth tapping — it is better at that
   * than a coordinate recorded against one build on one device, and it is the
   * part that keeps working when a layout moves.
   *
   * What the map keeps is the three things looking at a screenshot cannot tell
   * you:
   *
   *   veto       — PLAY commits the account to a live match with an abandon
   *                penalty; the red button on the promo is a real charge. None
   *                of that is visible. A proposal landing on a control the map
   *                forbids is dropped, whatever the model called it.
   *   name       — a proposal near a recorded control inherits its ref, so
   *                blocked patterns, known findings and the log all line up on
   *                one vocabulary instead of whatever text the model read.
   *   supplement — controls the model missed are added at the back rather than
   *                the front. Mostly info badges: the tooltips behind a small
   *                "?" carry text found on no other screen, and they are the
   *                easiest thing on a screen to overlook.
   */
  applyRouteToProposals(proposals, screenId) {
    const known = this.routeControlIndex(screenId);
    if (!known.length) return proposals;

    const near = (a, b) => Number.isFinite(a.x) && Number.isFinite(a.y)
      && Math.hypot(a.x - b.x, a.y - b.y) < 48;

    // On a root screen there is nothing behind, so going back opens the
    // exit-game dialog. The model cannot tell that from a screenshot — a back
    // arrow looks like a back arrow — and it proposed one every single pass: the
    // dialog opened, exitConfirm answered NO, the crawl landed on the lobby and
    // the arrow was proposed again. Three screens in ten minutes, nothing broken
    // and nothing gained.
    const match0 = this.routeScreenFor.get(screenId);
    const backExits = !!(match0 && match0.def && match0.def.backExits);
    const goesBack = (a) => a.kind === 'back'
      || /\b(back|return|previous|close)\b/i.test(String(a.label || ''))
      || /\.back$/i.test(String(a.label || ''));

    const kept = [];
    let vetoed = 0;
    for (const a of proposals) {
      const match = known.find((k) => near(a, k));
      const label = match ? match.ref : a.label;
      if (backExits && goesBack(a)) {
        this.run.skipped.push({ screenId, label, reason: 'going back from here opens the exit-game dialog' });
        vetoed += 1;
        continue;
      }
      if (this.isBlocked(label) || this.isBlocked(a.label)) {
        this.run.skipped.push({ screenId, label, reason: 'matches a blocked-label pattern' });
        vetoed += 1;
        continue;
      }
      kept.push(match ? { ...a, label, knownAs: match.ref } : a);
    }

    // Whatever the model did not see. Blocked ones never make it into the
    // index, so nothing dangerous can be added back here.
    const missed = this.routeActionsFor(screenId)
      .filter((r) => !proposals.some((p) => near(p, r)))
      .map((r) => ({ ...r, priority: 'medium' }));

    const bits = [];
    if (vetoed) bits.push(`vetoed ${vetoed}`);
    if (missed.length) bits.push(`added ${missed.length} it missed`);
    this.run.log(
      `${screenId}: the model proposed ${proposals.length} control${proposals.length === 1 ? '' : 's'}`
      + (bits.length ? `; the route map ${bits.join(' and ')}.` : '; the route map had nothing to add.')
    );
    return [...kept, ...missed];
  }

  /**
   * Every control the route map records for this screen, blocked ones included.
   *
   * routeActionsFor drops the blocked ones because it produces things to do.
   * This produces things to recognise, and a control has to stay recognisable
   * precisely so a proposal that lands on it can be refused.
   */
  routeControlIndex(screenId) {
    const match = this.routeScreenFor.get(screenId);
    if (!match) return [];
    const { name, def } = match;
    const out = [];
    const add = (ref, pair) => {
      const p = this.routePoint(pair);
      if (p) out.push({ ref, x: p.x, y: p.y });
    };
    for (const [label, pair] of Object.entries(def.controls || {})) add(`${name}.${label}`, pair);
    for (const [label, pair] of Object.entries(def.infoBadges || {})) add(`${name}.badge.${label}`, pair);
    return out;
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
      // The route's blocked list names controls in full — `lobby.play`, not
      // `play` — so the qualified ref is what has to be tested. Checking only
      // the bare name let PLAY through on a map that explicitly forbids it.
      const ref = `${name}.${label}`;
      if (this.isBlocked(ref) || this.isBlocked(label)) {
        this.run.skipped.push({ screenId, label: ref, reason: 'matches a blocked-label pattern' });
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

  /**
   * Put the game into the language the sheet column is being checked against.
   *
   * The picker used to be a handful of fixed coordinates. It is not: it is a
   * scrolling list of nineteen entries, so a language has to be found by
   * reading the screen. Three things about it were learned the hard way and
   * are all handled below — the list keeps gliding after a swipe, it re-anchors
   * itself to the currently selected entry a couple of seconds later, and the
   * app goes on rendering parts of its UI in the language it launched with
   * until it is restarted.
   */
  async applyRouteLanguage() {
    const proc = this.route && this.route.procedures && this.route.procedures.setLanguage;
    if (!proc || !this.adb) return false;

    const wanted = this.pickerEntryFor(this.target.header);
    if (!wanted) return false;

    // Drive it by looking, not by replaying coordinates. The recorded step list
    // below is kept only as a fallback: it assumes the app is sitting where it
    // was when someone wrote it, and when it is not — a promo over the lobby, a
    // shifted layout — every tap after the first miss lands somewhere arbitrary
    // while the procedure still reports success.
    if (this.cfg.modelDrivenNavigation !== false) {
      const hints = Object.entries((this.route && this.route.hazards) || {})
        .filter(([k]) => k !== '$comment')
        .map(([, v]) => String(v).split('.')[0] + '.')
        .slice(0, 6);

      const ok = await this.pursue(
        `Set the game's account language to "${wanted}". It is under Settings (the gear icon) on the Account tab, `
        + 'behind the CHANGE button on the Language row, which opens a scrolling list of languages. '
        + `Select "${wanted}" in that list so its checkbox is ticked, then press the confirm button. `
        + 'You are done only once the list is closed and the Language row shows the new language.',
        { maxSteps: 16, hints }
      );
      if (ok) {
        this.run.log('Language switch done.');
        return this.finishLanguageSwitch(proc, wanted);
      }
      this.run.log('Falling back to the recorded steps for the language switch.', 'warn');
    }

    this.run.log(`Setting the game's language to ${wanted} using the route map.`);
    for (const step of proc.steps || []) {
      if (this.run.stopRequested) return false;

      if (step.findAndTap) {
        const ok = await this.findAndTapListEntry(
          String(step.findAndTap).replace('{language}', wanted),
          step.in || 'settings.languagePicker'
        );
        if (!ok) return false;
      } else if (step.verifyChecked) {
        const ok = await this.verifyLanguageChecked(String(step.verifyChecked).replace('{language}', wanted));
        if (!ok) return false;
      } else if (step.action === 'restart' || step.restart) {
        if (!(await this.restartApp('so the UI stops rendering the language it launched with'))) return false;
      } else if (step.verifyLanguageApplied) {
        const ok = await this.verifyLanguageApplied(
          String(step.verifyLanguageApplied).replace('{language}', wanted), step
        );
        if (!ok && step.onMismatch === 'abort') return false;
      } else if (step.dismiss) {
        await this.clearAutoDismissModals(await this.settleOnly());
      } else if (step.tap) {
        if (!(await this.tapRouteStep(step))) return false;
      }
      if (step.wait) await sleep(Number(step.wait));
      else if (!step.action && !step.dismiss) await sleep(this.cfg.settleMs);
    }
    this.run.log('Language switch done.');
    return true;
  }

  /**
   * The picker row that corresponds to a sheet column.
   *
   * The two rarely read the same. A column headed "Portuguese (Brazil)" has to
   * find a row reading "PORTUGUESE (BRAZIL)", and a column headed "pt-BR" has
   * to find it without sharing a single word. Matching against the recorded
   * list of entries also means an unknown language is caught here — before the
   * picker is opened and something gets tapped — rather than after a fruitless
   * scroll through nineteen rows.
   */
  pickerEntryFor(header) {
    const wanted = String(header || '').trim();
    if (!wanted) return null;
    const picker = this.routeScreens['settings.languagePicker'];
    const entries = (picker && picker.entries && picker.entries.order) || [];
    if (!entries.length) return wanted;      // nothing recorded; try the header as written

    const squash = (s) => String(s).toLowerCase().replace(/[^a-z]/g, '');
    const a = squash(wanted);
    const exact = entries.find((e) => squash(e) === a);
    if (exact) return exact;
    // "Portuguese (Brazil)" against "PORTUGUESE (BRAZIL)" is already handled;
    // this catches "Portuguese" against the two Portuguese rows by preferring
    // the one the header is a prefix of.
    const partial = entries.find((e) => squash(e).startsWith(a) || a.startsWith(squash(e)));
    if (partial) {
      this.run.log(`Sheet column "${wanted}" matched the picker entry "${partial}".`);
      return partial;
    }
    this.run.log(`The language picker has no entry matching the sheet column "${wanted}" — leaving the game's language alone. Recorded entries: ${entries.join(', ')}.`, 'warn');
    return null;
  }

  /**
   * Find a row by its text in a scrolling list and tap it.
   *
   * Scrolling toward the row does not work on this picker, and the failure is
   * silent: the list re-anchors the CHECKED entry back to the top a second or
   * two after each swipe settles, so every capture shows the same few rows no
   * matter how far you swiped. Driven by hand on 2026-08-19 that produced
   * "scrolled to the end without finding THAI" against a list that plainly
   * contains THAI, and before that three wrong selections in a row.
   *
   * What does work is to stop fighting the re-anchor and use it. Tapping a row
   * only moves the checkbox — nothing is applied until CONFIRM — so a row on
   * the way to the target is a free stepping stone. Tap the furthest reachable
   * row toward the target, let the list re-anchor onto it, and the window has
   * moved down the list. Repeat until the target is in view, then tap it.
   *
   * The recorded entries.order says which direction the target lies in, so the
   * hops always go the right way. A list with no recorded order falls back to
   * swiping, which is right for every app that does not do this.
   */
  async findAndTapListEntry(needle, screenName) {
    const def = this.routeScreens[screenName];
    const scroll = def && def.scrollable;
    const order = ((def && def.entries && def.entries.order) || []).map((s) => String(s).toUpperCase());
    const size = this.deviceSize || { width: 1080, height: 1920 };
    const wantIdx = order.indexOf(String(needle).toUpperCase());
    const maxHops = 10;

    const tapRow = async (row) => {
      await this.tapAt(
        Math.max(0, Math.min(1, row.x)) * size.width,
        Math.max(0, Math.min(1, row.y)) * size.height
      );
    };

    // Preferred path: read the whole visible window and step toward the target.
    if (wantIdx >= 0) {
      let lastAnchor = null;
      for (let hop = 0; hop <= maxHops; hop++) {
        if (this.run.stopRequested) return false;

        // Settle fully. Racing the re-anchor is what the old code did; letting
        // it finish is what makes the next capture describe a list that will
        // still be there when the tap lands.
        const cap = await this.settleOnly();
        let rows = [];
        try {
          rows = await this.analyzer.readListRows(cap.png);
        } catch (e) {
          this.run.log(`Could not read ${screenName}: ${e.message}`, 'warn');
          break;
        }
        if (!rows.length) {
          this.run.log(`Could not read any list rows on ${screenName} — the picker may not be open.`, 'warn');
          break;
        }
        this.run.log(`${screenName} hop ${hop}: ${rows.map((r) => `${r.label}${r.checked ? '*' : ''}${r.visible ? '' : '(clipped)'}`).join(' | ')}`);

        const hit = rows.find((r) => r.visible && r.label.toUpperCase() === String(needle).toUpperCase());
        if (hit) {
          // Confirmed by the next pass rather than assumed here. Reading the
          // list costs twenty seconds, which is far longer than the re-anchor
          // takes to fire — so on a freshly opened picker, which always shows
          // the top of the list before snapping back to the checked entry, the
          // row really was where the model said and is simply not there any
          // more by the time the tap lands. Tapping and looking again turns
          // that from a failed run into one more cheap iteration.
          if (hit.checked) return true;
          await tapRow(hit);
          await sleep(2500);
          continue;
        }

        // Which recorded entries can actually be tapped right now?
        const reachable = rows
          .filter((r) => r.visible)
          .map((r) => ({ row: r, idx: order.indexOf(r.label.toUpperCase()) }))
          .filter((r) => r.idx >= 0);
        if (!reachable.length) {
          this.run.log(`Nothing on screen matches the recorded ${screenName} entries (saw: ${rows.map((r) => r.label).filter(Boolean).join(', ') || 'nothing'}).`, 'warn');
          break;
        }

        const nearest = reachable.reduce((best, r) =>
          Math.abs(r.idx - wantIdx) < Math.abs(best.idx - wantIdx) ? r : best);

        // Going down the list is the easy direction: the re-anchor parks the
        // checked row at the TOP of the viewport, so the rows below it are all
        // reachable and tapping the lowest one drags the window down.
        if (nearest.idx < wantIdx && nearest.idx !== lastAnchor) {
          lastAnchor = nearest.idx;
          this.run.log(`"${needle}" is below the visible rows; stepping to "${nearest.row.label}".`);
          await tapRow(nearest.row);
          await sleep(2500);                        // let the re-anchor finish
          continue;
        }

        // Scrolling up, which works — but only if the drag starts inside the
        // list rather than at its edge.
        //
        // This was misdiagnosed once, expensively. Drags beginning near the top
        // of the recorded region moved nothing, twice, at two speeds, and the
        // conclusion drawn was that the picker refuses to scroll above the
        // selected entry. It does not. A drag starting mid-list walks straight
        // up to ENGLISH at the top. The gesture was landing on the modal's edge
        // instead of the list, so the list never saw it.
        //
        // The lesson is worth more than the fix: "the control did not respond"
        // and "the app forbids this" look identical from outside, and only one
        // of them is a bug worth reporting.
        if (!scroll || !Array.isArray(scroll.region)) {
          this.run.log(`${screenName} records no scrollable region — cannot reach "${needle}".`, 'warn');
          break;
        }
        const [sx0, sy0, sx1, sy1] = scroll.region;
        const midX = (sx0 + sx1) / 2 * size.width;
        const height = sy1 - sy0;
        const minBefore = Math.min(...reachable.map((r) => r.idx));

        try {
          // Well inside the list at both ends: 40% down to 85% down.
          await this.adb.swipe(
            Math.round(midX), Math.round((sy0 + height * 0.40) * size.height),
            Math.round(midX), Math.round((sy0 + height * 0.85) * size.height), 600
          );
        } catch (e) {
          this.run.log(`Could not scroll ${screenName}: ${e.message}`, 'warn');
          break;
        }
        await sleep(900);

        if (minBefore === lastAnchor) {
          this.run.log(`Scrolling up past "${order[minBefore]}" is not getting any further.`, 'warn');
          break;
        }
        lastAnchor = minBefore;
        this.run.log(`Scrolled ${screenName} up toward "${needle}".`);
      }
    }

    // Fallback for lists with no recorded order: the original swipe sweep.
    for (let sweep = 0; sweep <= 8; sweep++) {
      if (this.run.stopRequested) return false;

      const cap = await this.settleOnly();
      let point = null;
      try {
        point = await this.analyzer.locateText(cap.png, needle);
      } catch (e) {
        this.run.log(`Could not look for "${needle}": ${e.message}`, 'warn');
        return false;
      }
      if (point) {
        await tapRow({ x: point.x, y: point.y });
        await sleep(2500);
        // Same race as above: locating costs long enough for the list to
        // re-anchor underneath the tap. Saying "found it" without checking is
        // what turned a missed tap into a whole run scanned in the wrong
        // language, so the sweep only stops once the row reads back as checked.
        try {
          const after = await this.analyzer.readListRows(await this.screenshot());
          const now = after.find((r) => r.label.toUpperCase() === String(needle).toUpperCase());
          if (now && now.checked) return true;
          this.run.log(`Tapped "${needle}" but it did not take — trying again.`);
        } catch {
          return true;                     // cannot check; assume it landed
        }
      }

      if (!scroll || !Array.isArray(scroll.region)) {
        this.run.log(`"${needle}" is not on screen and ${screenName} has no scrollable region recorded.`, 'warn');
        return false;
      }
      const [x0, y0, x1, y1] = scroll.region;
      const midX = (x0 + x1) / 2 * size.width;
      const top = (y0 + (y1 - y0) * 0.25) * size.height;
      const bottom = (y0 + (y1 - y0) * 0.75) * size.height;
      try {
        await this.adb.swipe(Math.round(midX), Math.round(bottom), Math.round(midX), Math.round(top), 500);
      } catch (e) {
        this.run.log(`Could not scroll ${screenName}: ${e.message}`, 'warn');
        return false;
      }
    }
    this.run.log(`Could not reach "${needle}" in ${screenName}.`, 'warn');
    return false;
  }

  /**
   * Read the checkbox back before confirming.
   *
   * Cheap, and it catches the case this list is prone to: the tap landed a row
   * off, or on nothing at all because the list moved underneath it. Confirming
   * without checking means the whole scan then runs against the wrong language
   * and every finding in it is worthless.
   */
  async verifyLanguageChecked(needle) {
    const cap = await this.settleOnly();
    let checked = null;
    try {
      checked = await this.analyzer.locateText(cap.png, needle, { requireChecked: true });
    } catch (e) {
      this.run.log(`Could not verify the language selection: ${e.message}`, 'warn');
      return true;      // do not abandon the switch over a failed read
    }
    if (!checked) {
      this.run.log(`"${needle}" does not look selected — not confirming, and leaving the language alone.`, 'warn');
      this.run.warnings.push(`Could not select ${needle} in the language picker; the scan ran in whatever language the game was already in.`);
      return false;
    }
    this.run.log(`${needle} is selected.`);
    return true;
  }

  /**
   * The part of the language switch that stays scripted, on purpose.
   *
   * The restart and the read-back are not navigation and are not a matter of
   * judgement: the account language reverts intermittently, so it is confirmed
   * from Settings whichever way the switch was driven, and a mismatch stops the
   * run rather than scanning a language nobody asked for.
   */
  async finishLanguageSwitch(proc, wanted) {
    const steps = (proc && proc.steps) || [];
    if (steps.some((s) => s.action === 'restart' || s.restart)) {
      if (!(await this.restartApp('so the UI stops rendering the language it launched with'))) return false;
      await sleep(3000);
      await this.clearAutoDismissModals(await this.settleOnly());
    }
    const verify = steps.find((s) => s.verifyLanguageApplied);
    if (!verify) return true;
    const ok = await this.verifyLanguageApplied(wanted, verify);
    return ok || verify.onMismatch !== 'abort';
  }

  /**
   * Chase a goal by looking, deciding and acting, one step at a time.
   *
   * The general replacement for writing procedures down. A recorded step list
   * assumes the app is where it was when someone wrote it; on this title it
   * usually is not, and the failure is silent — the taps land somewhere, the
   * script reports success, and the run continues in the wrong place. The
   * language switch failed that way repeatedly.
   *
   * Here nothing is assumed between steps. Screenshot, decide, act, screenshot
   * again. It costs a call per step, which is the price of not being wrong
   * about where you are.
   *
   * The route map is still consulted, for the one thing looking cannot supply:
   * an action the map forbids is refused however sensible it looked, and
   * hazards are passed in as things worth knowing.
   */
  async pursue(goal, { maxSteps = 12, hints = [] } = {}) {
    if (!this.adb || !this.analyzer) return false;
    const size = this.deviceSize || { width: 1080, height: 1920 };
    const history = [];
    let lastSig = null;
    let stalled = 0;

    this.run.log(`Working towards: ${goal}`);

    for (let step = 0; step < maxSteps; step++) {
      if (this.run.stopRequested) return false;

      const cap = await this.settleOnly();
      const sig = this.signature(cap);
      if (sig && sig === lastSig) {
        stalled += 1;
        if (stalled >= 3) {
          this.run.log('The screen has stopped responding to anything tried here.', 'warn');
          return false;
        }
      } else {
        stalled = 0;
      }
      lastSig = sig;

      let move;
      try {
        move = await this.analyzer.decideNextAction(cap.png, goal, history, hints);
      } catch (e) {
        this.run.log(`Could not decide what to do next: ${e.message}`, 'warn');
        return false;
      }
      if (!move) return false;

      if (move.action === 'done') {
        this.run.log(`Reached it: ${move.why}`);
        return true;
      }
      if (move.action === 'stuck') {
        this.run.log(`Cannot get there from this screen: ${move.why}`, 'warn');
        return false;
      }
      if (move.action === 'wait') {
        history.push('waited for the screen to settle');
        await sleep(2000);
        continue;
      }

      // The one veto looking cannot supply.
      const known = this.routeControlIndex(this.currentScreenId || '');
      const hit = known.find((k) => move.x != null
        && Math.hypot(k.x - move.x * size.width, k.y - move.y * size.height) < 48);
      if (hit && this.isBlocked(hit.ref)) {
        this.run.log(`Refused "${move.target || hit.ref}" — the route map forbids it.`, 'warn');
        history.push(`tried ${move.target || hit.ref}, refused as unsafe`);
        continue;
      }

      try {
        if (move.action === 'back') {
          await this.adb.shell('input keyevent KEYCODE_BACK');
          history.push('pressed back');
        } else if (move.action === 'swipe' && move.x2 != null) {
          await this.adb.swipe(
            Math.round(move.x * size.width), Math.round(move.y * size.height),
            Math.round(move.x2 * size.width), Math.round(move.y2 * size.height), 500
          );
          history.push(`swiped ${move.target || ''}`.trim());
        } else if (move.x != null) {
          await this.tapAt(move.x * size.width, move.y * size.height);
          history.push(`tapped ${move.target || `${move.x.toFixed(2)},${move.y.toFixed(2)}`}`);
        } else {
          continue;
        }
      } catch (e) {
        this.run.log(`Could not ${move.action}: ${e.message}`, 'warn');
        return false;
      }
      this.run.log(`${move.action} ${move.target ? `"${move.target}"` : ''} — ${move.why}`);
      await sleep(this.cfg.settleMs);
    }

    this.run.log(`Gave up on "${goal}" after ${maxSteps} steps.`, 'warn');
    return false;
  }

  /**
   * Look again at the findings that were decided on a few pixels.
   *
   * Nine of the issue types are judgements about rendering rather than meaning
   * — truncation, overflow, clipping, overlap, contrast, mojibake, wrong
   * glyphs. All of them are settled at full-screen size on marks a few pixels
   * tall, and that is exactly where a scan starts inventing defects. A Thai
   * tone mark or a Vietnamese hook is at the limit of what is visible in a
   * 2400x1080 frame.
   *
   * So each one gets cropped, magnified and looked at again. Driven by hand
   * this retracted two findings and confirmed a third, and the two it retracted
   * were correctly spelled strings — precisely the kind of mistake that makes a
   * translator stop trusting the whole report.
   *
   * Semantic findings are left alone: whether "Store" was translated as a verb
   * is not a question magnification can answer, and paying for a call to ask it
   * would be waste.
   */
  async confirmUnderMagnification(issues, cap) {
    if (!this.cfg.zoomVerify || !this.analyzer || !cap || !cap.png) return issues;

    const PIXEL_JUDGEMENTS = new Set([
      'truncated', 'overflow_horizontal', 'overflow_vertical', 'offscreen',
      'overlap', 'clipped_by_art', 'unreadable_contrast', 'mojibake', 'wrong_font_glyphs',
    ]);
    const budget = Number(this.cfg.zoomVerifyMax) || 6;

    const out = [];
    let checked = 0, dropped = 0;
    for (const issue of issues) {
      const wants = PIXEL_JUDGEMENTS.has(issue.type) && issue.rect && checked < budget;
      if (!wants) { out.push(issue); continue; }

      const crop = cropAndZoom(cap.png, issue.rect);
      if (!crop) { out.push(issue); continue; }
      checked += 1;

      let verdict;
      try {
        verdict = await this.analyzer.verifyFinding(crop.buffer, issue, { language: this.target.header });
      } catch (e) {
        this.run.log(`Could not re-check "${issue.text}": ${e.message}`, 'warn');
        out.push(issue);
        continue;
      }

      if (!verdict.holds) {
        dropped += 1;
        this.run.log(`Dropped "${issue.text}" (${issue.type}) — at ${crop.factor}x it does not hold: ${verdict.why}`);
        continue;
      }
      out.push({
        ...issue,
        severity: verdict.severity || issue.severity,
        text: verdict.text || issue.text,
        confidence: 'certain',
        zoomVerified: crop.factor,
      });
    }

    if (checked) {
      this.run.log(
        `Re-checked ${checked} rendering finding${checked === 1 ? '' : 's'} magnified` +
        (dropped ? `; ${dropped} did not survive a closer look.` : '; all held up.')
      );
    }
    return out;
  }

  /**
   * Read the account language back after the restart, and refuse to scan if it
   * is not the language that was asked for.
   *
   * The restart is not free. Driving it by hand on 2026-08-19 reverted the
   * account to English three times out of four — Thai twice and Japanese once,
   * on a healthy network — and held only on the fourth. Intermittent is the
   * dangerous kind: the run continues, captures a build rendering English,
   * compares it against the Thai column and reports every string as
   * untranslated. Nothing about that run looks wrong from the outside.
   *
   * So the switch is not believed, it is checked. Navigating to Settings >
   * Account costs two taps and turns a silent worthless run into a loud one.
   */
  async verifyLanguageApplied(wanted, step) {
    const via = Array.isArray(step.via) ? step.via : [];
    for (const ref of via) {
      const point = this.resolveRef(String(ref));
      if (!point) {
        this.run.log(`Cannot reach the language row — "${ref}" does not resolve, so the switch is unverified.`, 'warn');
        return true;                       // a broken map must not fail the run
      }
      await this.tapAt(point.x, point.y);
      await sleep(this.cfg.settleMs * 2);
    }

    const cap = await this.settleOnly();
    let found = null;
    try {
      found = await this.analyzer.locateText(cap.png, wanted, {});
    } catch (e) {
      this.run.log(`Could not read the language back: ${e.message}`, 'warn');
      return true;                         // an unreadable screen is not proof of failure
    }
    if (found) {
      this.run.log(`Language reads back as ${wanted} after the restart.`);
      return true;
    }

    const msg = `The restart did not keep ${wanted} — Settings still reports a different language. `
      + 'Scanning now would report the whole build as untranslated, so the run is stopping instead.';
    this.run.log(msg, 'warn');
    this.run.warnings.push(msg);
    return false;
  }

  /** Force-stop and relaunch, then wait for the app to come back. */
  async restartApp(why) {
    const pkg = (this.cfg.androidPackage || '').trim();
    if (!pkg) {
      this.run.warnings.push('The route map asks for a restart after the language change, but no Android package is set, so it was skipped — some strings may still be in the previous language.');
      return true;
    }
    this.run.log(`Restarting ${pkg} ${why}.`);
    try {
      await this.adb.shell(`am force-stop ${pkg}`);
      await sleep(2000);
      await this.adb.shell(`monkey -p ${pkg} -c android.intent.category.LAUNCHER 1`);
    } catch (e) {
      this.run.log(`Restart failed: ${e.message}`, 'warn');
      return false;
    }
    return true;
  }

  /**
   * Replays the tester's setup script before the crawl: dismiss the daily
   * popup, sign in, walk to the part of the app that needs checking. Steps run
   * in order and a failure stops the script rather than carrying on from a
   * screen the rest of the steps were never written for.
   */
  async runSteps(steps, label = 'Setup steps') {
    if (!steps || !steps.length) return true;
    this.run.log(`${label}: replaying ${steps.length} step${steps.length === 1 ? '' : 's'}.`);
    const size = () => this.deviceSize || { width: 1080, height: 1920 };
    // A coordinate of 0–1 is a fraction of the screen, so a script survives a
    // different device; anything larger is taken as device pixels.
    const px = (v, span) => (Math.abs(v) <= 1 ? v * span : v);

    for (const step of steps) {
      if (this.run.stopRequested) return false;
      const s = size();
      try {
        switch (step.verb) {
          case 'tap':
            await this.tapAt(px(step.x, s.width), px(step.y, s.height));
            break;
          case 'longpress': {
            const x = px(step.x, s.width);
            const y = px(step.y, s.height);
            const ms = step.ms || this.cfg.longPressMs;
            if (this.useBridge) await this.bridge.longPress(x, y, ms);
            else if (this.adb) await this.adb.longPress(x, y, ms);
            else throw new Error('no input transport available');
            break;
          }
          case 'swipe':
            if (!this.adb) throw new Error('swipe needs a device connection');
            await this.adb.swipe(px(step.x1, s.width), px(step.y1, s.height), px(step.x2, s.width), px(step.y2, s.height), step.ms);
            break;
          case 'text':
            if (!this.adb) throw new Error('text needs a device connection');
            await this.adb.typeText(step.arg);
            break;
          case 'key':
            if (!this.adb) throw new Error('key needs a device connection');
            await this.adb.keyevent(step.arg);
            break;
          case 'wait':
            await sleep(step.ms);
            break;
          case 'back':
            await this.goBack();
            break;
          case 'home':
            if (!this.adb) throw new Error('home needs a device connection');
            await this.adb.home();
            break;
          case 'launch':
          case 'restart': {
            if (!this.adb) throw new Error(`${step.verb} needs a device connection`);
            const pkg = this.cfg.androidPackage;
            if (!pkg) throw new Error(`${step.verb} needs the Android package name`);
            if (step.verb === 'restart') {
              await this.adb.forceStop(pkg);
              await sleep(1200);
            }
            await this.adb.launch(pkg);
            await sleep(3000);
            break;
          }
          case 'shell':
            if (!this.adb) throw new Error('shell needs a device connection');
            await this.adb.shell(step.arg);
            break;
          default:
            break;
        }
      } catch (e) {
        const msg = `${label} stopped at line ${step.line} ("${step.source}"): ${e.message}`;
        this.run.log(msg, 'warn');
        this.run.warnings.push(msg);
        return false;
      }
      if (step.verb !== 'wait') await sleep(this.cfg.settleMs);
    }
    this.run.log(`${label}: done.`);
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

    // Try the route map's own close controls first. Back is a blunt instrument
    // and in some games a dangerous one: on the Indus lobby it opens the
    // exit-game confirmation, so a generic "press back until the overlay goes
    // away" loop eventually quits the app and ends the run. If the route map
    // names the button that closes this modal, press that instead.
    if (this.route) {
      const here = await this.currentSig();
      if (await this.clearAutoDismissModals(here.cap)) {
        const after = await this.currentSig();
        if (!this.rootSig || after.sig === this.rootSig) return true;
      }
      if (this.hazardWarnsAgainstBack()) {
        this.run.log('The route map flags back as unsafe in this app — not using it to clear overlays.', 'warn');
        return false;
      }
    }

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
    // Where we are, for anything that needs the route map's opinion about this
    // screen while it is being worked on — the veto in pursue(), in particular.
    this.currentScreenId = screenId;

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
        customRules: this.rules,
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
    // Signature matching is a substring test against a list of sample strings,
    // which fails the moment a screen is in a language nobody recorded samples
    // in — and then the map contributes nothing at all: no veto over what the
    // model wants to tap, no names in the log, and a generic swipe down a
    // screen the map already says does not scroll. Asking is cheap and it is
    // the same question, put to something that can actually read the screen.
    let known = this.identifyRouteScreen(seen);
    if (!known) known = await this.identifyScreenByVision(cap);
    if (known) {
      this.routeScreenFor.set(screenId, known);
      this.run.log(`Route map recognises ${screenId} as "${known.name}".`);
      // Without the bridge there are no strings until this point, so a modal
      // marked "dismiss on sight" can only be recognised once it has been
      // read. Note it here; actionsFor closes it instead of exploring it.
      if (known.def && known.def.autoDismiss) {
        this.autoDismissScreens.add(screenId);
        this.run.log(`${screenId} is "${known.name}", which the route map says to dismiss rather than explore.`);
      }
    }

    // Which route screen this is can only be known once its strings have been
    // read, so the source-language pair is dealt with here rather than before
    // the vision pass.
    const baselineIssues = await this.compareWithBaseline(known, cap, seen, screenId, file, ctx, depth, pathLabels);
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
        customRules: this.rules,
      })) {
        issue.source = 'static-ocr';
        aiStatic.push(issue);
      }
    }

    let issues = [
      ...staticIssues,
      ...aiStatic,
      ...baselineIssues,
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
    deduped = await this.confirmUnderMagnification(deduped, cap);
    if (this.mem) {
      const { kept, removed } = memory.filterDismissed(this.mem, deduped);
      if (removed) this.run.log(`Ignored ${removed} finding${removed === 1 ? '' : 's'} you dismissed on an earlier scan.`);
      deduped = kept;
    }
    this.highIssues += deduped.filter((i) => i.severity === 'high').length;
    // The language-mismatch call that used to sit here came from an older
    // branch and was carried in by a merge. It has no definition here, and it
    // would be the wrong place for one: detectLanguageMismatch already runs
    // above, off the deterministic pass, before a vision call is spent rather
    // than after the findings are built.
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
   * Compare this screen with the same screen in the source language.
   *
   * On a source-language run this only records: the screenshot and its strings
   * are filed under the route screen's name so later runs have something to
   * compare against. On a target-language run, if a recording exists, the two
   * captures are put side by side and the differences between them are the
   * findings — which strings did not change, and which ones no longer fit.
   *
   * The comparison is worth its extra call for a reason that is easy to miss:
   * it is the only check here that can *clear* a finding. A label running off
   * its plate looks like a translation defect right up until the same label
   * runs off the same plate in English, at which point it stops being the
   * translator's problem and becomes the layout's. Reporting it against the
   * translation would send the wrong team after it.
   */
  async compareWithBaseline(known, cap, seen, screenId, file, ctx, depth, pathLabels) {
    if (!this.baselineEnabled || !known) return [];
    const screenName = known.name;

    if (this.capturingBaseline) {
      baseline.record(this.routeName, this.sourceLanguage, screenName, { texts: seen, png: cap.png });
      return [];
    }

    const pair = baseline.get(this.routeName, this.sourceLanguage, screenName);
    if (!pair) {
      this.run.log(`No ${this.sourceLanguage} capture of "${screenName}" yet — scan once in ${this.sourceLanguage} to get overflow comparisons here.`);
      return [];
    }

    // The cheap half needs no model at all: a string that is still exactly the
    // source string was not translated.
    const identical = baseline.untranslatedCandidates(pair.texts, seen);

    if (!pair.png || !this.cfg.visionEnabled) {
      return identical.map((text) => ({
        source: 'baseline',
        type: 'untranslated',
        severity: 'medium',
        confidence: 0.8,
        text,
        where: screenName,
        message: `"${text}" is character-for-character the ${this.sourceLanguage} string on this screen.`,
      }));
    }

    let result = { issues: [] };
    try {
      result = await this.analyzer.compareToBaseline(pair.png, cap.png, {
        ...ctx,
        screenName,
        baselineLanguage: this.sourceLanguage,
        identical,
      });
    } catch (e) {
      this.run.log(`Could not compare "${screenName}" with its ${this.sourceLanguage} capture: ${e.message}`, 'warn');
      return [];
    }

    const issues = (result.issues || []).map((i) => ({
      source: 'baseline',
      type: i.type,
      severity: i.severity,
      confidence: i.confidence,
      text: i.text,
      where: i.where || screenName,
      element: i.element || '',
      expected: i.expected || '',
      message: i.message,
    }));
    if (issues.length) {
      this.run.log(`Comparing "${screenName}" with its ${this.sourceLanguage} capture found ${issues.length} difference${issues.length === 1 ? '' : 's'}.`);
    }
    return issues;
  }

  /**
   * Text below the fold is invisible to a single screenshot, so every screen
   * gets scrolled through before we move on. With the bridge we drive each
   * ScrollRect exactly; without it we swipe and watch for the screen to stop
   * changing, which is what "reached the bottom" looks like from outside.
   */
  async probeScrolls(cap, baseId, depth, pathLabels) {
    if (!this.cfg.scrollProbe) return;

    // A modal that is about to be closed has nothing below the fold worth a
    // vision call. Without this the daily-login popup was analysed, then
    // swiped and analysed three more times, and only then dismissed — four
    // calls and four screens of the run's budget on one modal.
    if (this.autoDismissScreens.has(baseId)) return;

    if (this.useBridge && cap.state && Array.isArray(cap.state.scrolls)) {
      return this.probeBridgeScrolls(cap, baseId, depth, pathLabels);
    }
    const regions = this.routeScrollRegions(baseId);
    if (regions.length) return this.probeRouteScrolls(regions, baseId, depth, pathLabels);

    // The route map recognised this screen and recorded no scrollable region,
    // which is a statement that it does not scroll — not an invitation to
    // guess. Blind-swiping a recognised screen is how a static modal over an
    // animated lobby looked like three new screens: the background kept
    // moving, so "nothing changed" never became true.
    if (this.routeScreenFor.has(baseId)) return;
    return this.probeSwipeScrolls(cap, baseId, depth, pathLabels);
  }

  /**
   * The scrollable areas an earlier pass recorded for this screen.
   *
   * The generic probe swipes vertically up the middle of the screen, which is
   * right often enough to be worth keeping and wrong in exactly the cases that
   * hide the most text: a weapon strip that scrolls sideways along the bottom,
   * a rewards rail a third of the way across that is narrower than the swipe,
   * an item grid beside a detail panel that must not be dragged. Swiping the
   * centre of those screens either does nothing or drags the wrong thing.
   */
  routeScrollRegions(screenId) {
    const match = this.routeScreenFor.get(screenId);
    if (!match) return [];
    const { name, def } = match;
    const out = [];
    const add = (label, s) => {
      if (!s || !Array.isArray(s.region) || s.region.length !== 4) return;
      out.push({ label: `${name}.${label}`, region: s.region, axis: s.axis === 'horizontal' ? 'horizontal' : 'vertical' });
    };
    add('scroll', def.scrollable);
    for (const [gridName, grid] of Object.entries(def.grids || {})) add(gridName, grid && grid.scrollable);
    return out;
  }

  /** Swipe inside each recorded region, along the axis it actually scrolls. */
  async probeRouteScrolls(regions, baseId, depth, pathLabels) {
    if (!this.adb) return;
    const size = this.deviceSize || { width: 1080, height: 1920 };

    for (const region of regions) {
      const [x0, y0, x1, y1] = region.region;
      const midX = (x0 + x1) / 2 * size.width;
      const midY = (y0 + y1) / 2 * size.height;
      // Swipe across 60% of the region so the gesture stays inside it — a
      // longer drag starting outside picks up whatever is next door.
      const spanX = (x1 - x0) * size.width * 0.3;
      const spanY = (y1 - y0) * size.height * 0.3;
      const horizontal = region.axis === 'horizontal';

      const fwd = horizontal
        ? [midX + spanX, midY, midX - spanX, midY]
        : [midX, midY + spanY, midX, midY - spanY];

      let prevSig = null;
      let steps = 0;
      for (let step = 1; step <= this.cfg.scrollSteps; step++) {
        if (this.run.stopRequested) break;
        try {
          await this.adb.swipe(Math.round(fwd[0]), Math.round(fwd[1]), Math.round(fwd[2]), Math.round(fwd[3]), 400);
        } catch (e) {
          this.run.log(`Could not swipe ${region.label}: ${e.message}`, 'warn');
          break;
        }
        const next = await this.settle();
        const sig = this.signature(next);
        if (sig === prevSig) break;            // reached the end of the list
        prevSig = sig;
        steps++;
        if (this.visited.has(sig)) continue;

        const id = `${baseId}-${region.axis === 'horizontal' ? 'pan' : 'scroll'}${step}`;
        this.visited.set(sig, id);
        if (next.hash) this.hashes.push({ hash: next.hash, sig });
        this.screenCount++;
        this.run.log(`Scrolled ${region.label} (${step}) — capturing ${id}`);
        await this.analyzeScreen(next, id, depth, [...pathLabels, `${region.label} ${step}`]);
      }

      // Put it back, so the parent screen still matches when we return to it.
      for (let i = 0; i < steps; i++) {
        try {
          await this.adb.swipe(Math.round(fwd[2]), Math.round(fwd[3]), Math.round(fwd[0]), Math.round(fwd[1]), 400);
        } catch { break; }
      }
      if (steps) await sleep(this.cfg.settleMs);
    }
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
    // Everything this screen has already shown, so a scroll that adds nothing
    // to it can be recognised as wasted.
    const seenText = new Set(
      (this.observedText.get(baseId) || []).map((t) => String(t).trim().toLowerCase())
    );

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

      // A scrolled view that reveals no string the screen has not already
      // shown is that screen photographed twice. The pixel checks above miss
      // this whenever something behind the content animates — a character
      // idling, a carousel rotating — because then the frame really did change
      // even though nothing readable did. Text is the thing being scanned, so
      // text is what decides whether the scroll was worth anything.
      const before = seenText.size;
      for (const t of (this.observedText.get(id) || [])) seenText.add(String(t).trim().toLowerCase());
      if (seenText.size === before) {
        this.run.log(`${id} showed nothing ${baseId} had not already shown — moving on.`);
        break;
      }
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

    if (Number(this.cfg.maxMinutes) > 0) {
      this.deadline = Date.now() + Number(this.cfg.maxMinutes) * 60000;
      this.run.log(`Time budget: ${this.cfg.maxMinutes} minutes.`);
    }

    // The setup script runs before anything else — including the language
    // switch, which usually needs the game past its title screen anyway.
    if (this.steps.length) {
      if (!this.deviceSize) {
        try { this.deviceSize = pngSize(await this.screenshot()); } catch { /* runSteps falls back to a default */ }
      }
      await this.runSteps(this.steps);
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
      if (this.deadline && Date.now() > this.deadline) {
        this.run.warnings.push(`Stopped at the ${this.cfg.maxMinutes}-minute time budget with ${this.stack.length} controls still queued.`);
        break;
      }
      const highLimit = Number(this.cfg.stopAfterHighIssues) || 0;
      if (highLimit > 0 && this.highIssues >= highLimit) {
        this.run.warnings.push(`Stopped after ${this.highIssues} high-severity findings, as configured.`);
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
