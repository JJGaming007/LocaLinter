'use strict';

const { runChecks } = require('./checks');
const { perceptualHash, hammingDistance } = require('./png');
const { norm } = require('./sheet');

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
  constructor({ cfg, adb, bridge, sheet, analyzer, run, target }) {
    this.cfg = cfg;
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

    return actions;
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

  async resetToRoot() {
    // 1. cheap: back out until we recognise the root
    for (let i = 0; i < 6; i++) {
      const { sig } = await this.currentSig();
      if (sig === this.rootSig) return true;
      await this.goBack();
      await sleep(this.cfg.settleMs);
    }
    // 2. expensive: restart the app
    if (this.adb && this.cfg.androidPackage) {
      this.run.log(`Restarting ${this.cfg.androidPackage} to get back to the first screen.`);
      await this.adb.forceStop(this.cfg.androidPackage);
      await sleep(1200);
      await this.adb.launch(this.cfg.androidPackage);
      await sleep(4000);
      await this.settle();
      const { sig } = await this.currentSig();
      if (sig === this.rootSig) return true;
      // a fresh launch may legitimately land somewhere new (a daily popup, say)
      this.rootSig = sig;
      return true;
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

    const issues = [
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

    const deduped = dedupe(issues);
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

  async start() {
    this.run.status = 'running';
    this.run.emit('status', { status: 'running', mode: this.mode });
    this.run.log(`Capture mode: ${this.mode}`);

    if (this.adb) {
      try {
        this.deviceSize = await this.adb.screenSize();
        this.run.log(`Device screen: ${this.deviceSize.width}x${this.deviceSize.height}`);
      } catch (e) {
        this.run.log(`Could not read the screen size: ${e.message}`, 'warn');
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
