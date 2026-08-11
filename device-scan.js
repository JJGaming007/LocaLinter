/**
 * Device Scan tab.
 *
 * The browser cannot speak ADB or hold an API key safely, so everything real
 * happens in the local agent (agent/server.js). This file is the control panel:
 * it hands the already-loaded sheet to the agent, streams the run back over
 * SSE, and renders the findings.
 */
(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const AGENT_KEY = 'localinter_agent_url';

  const el = {};
  let agentOk = false;
  let agentConfig = null;
  let runId = null;
  let source = null;           // EventSource
  let issues = [];
  let screens = new Map();
  let usageSent = null;          // cumulative usage already forwarded for this run
  let running = false;

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    [
      'ds-agent-url', 'ds-reconnect', 'ds-agent-status', 'ds-api-key', 'ds-save-key', 'ds-key-state',
      'ds-mode', 'ds-device', 'ds-sheet', 'ds-language', 'ds-package', 'ds-probe', 'ds-start', 'ds-stop',
      'ds-advanced-toggle', 'ds-advanced', 'ds-probe-out', 'ds-agent-out', 'ds-target-status',
      'ds-max-screens', 'ds-max-actions', 'ds-max-depth', 'ds-model', 'ds-effort', 'ds-adb-path',
      'ds-vision', 'ds-scroll', 'ds-longpress', 'ds-blocked',
      'ds-run-panel', 'ds-run-status', 'ds-stat-screens', 'ds-stat-actions', 'ds-stat-issues',
      'ds-stat-queue', 'ds-summary', 'ds-log', 'ds-results-panel', 'ds-findings',
      'ds-filter-severity', 'ds-filter-type', 'ds-filter-text', 'ds-export',
      'ds-shot-overlay', 'ds-shot-img', 'ds-shot-caption', 'ds-shot-close', 'badge-device',
    ].forEach((id) => { el[id.replace(/-([a-z])/g, (_, c) => c.toUpperCase())] = $(id); });

    if (!el.dsAgentUrl) return; // tab not present

    const saved = localStorage.getItem(AGENT_KEY);
    if (saved) el.dsAgentUrl.value = saved;

    el.dsReconnect.addEventListener('click', connect);
    el.dsSaveKey.addEventListener('click', saveKey);
    el.dsAdvancedToggle.addEventListener('click', () => el.dsAdvanced.classList.toggle('hidden'));
    el.dsMode.addEventListener('change', onModeChange);
    el.dsProbe.addEventListener('click', probe);
    el.dsStart.addEventListener('click', start);
    el.dsStop.addEventListener('click', stop);
    el.dsExport.addEventListener('click', exportJson);
    el.dsShotClose.addEventListener('click', () => el.dsShotOverlay.classList.add('hidden'));
    el.dsShotOverlay.addEventListener('click', (e) => {
      if (e.target === el.dsShotOverlay) el.dsShotOverlay.classList.add('hidden');
    });
    [el.dsFilterSeverity, el.dsFilterType].forEach((s) => s.addEventListener('change', renderFindings));
    el.dsFilterText.addEventListener('input', renderFindings);

    el.dsSheet.addEventListener('change', onSheetChange);

    document.addEventListener('localinter:tab', (e) => {
      if (e.detail.tabId === 'device-scan') refreshSheets();
    });
    // Loading a sheet anywhere in the app re-points the scan at it.
    document.addEventListener('localinter:sheet', refreshSheets);

    onModeChange();
    refreshSheets();
    connect();
  }

  // ── agent ───────────────────────────────────────────────────────────────

  function agentUrl() {
    return (el.dsAgentUrl.value || 'http://127.0.0.1:8790').replace(/\/+$/, '');
  }

  async function api(path, options) {
    let res;
    try {
      res = await fetch(agentUrl() + path, {
        headers: { 'content-type': 'application/json' },
        ...options,
      });
    } catch (e) {
      // fetch only rejects when the request never reached the agent. "Failed to
      // fetch" tells the user nothing, so say what is actually wrong.
      agentOk = false;
      setStatus(el.dsAgentStatus, 'Not running', 'bad');
      showAgentError(`Cannot reach the agent at ${agentUrl()}.`);
      updateStartState();
      throw new Error(`The agent is not running at ${agentUrl()} — start it with "npm start" in agent/, then hit Reconnect.`);
    }
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `agent returned ${res.status}`);
    return body;
  }

  function showAgentError(headline) {
    el.dsAgentOut.innerHTML =
      `<div class="ds-msg bad">${escapeHtml(headline)}<br>` +
      `Open a terminal in <code>agent/</code> and run <code>npm install</code> then <code>npm start</code>.</div>`;
  }

  async function connect() {
    setStatus(el.dsAgentStatus, 'Connecting…', 'pending');
    localStorage.setItem(AGENT_KEY, agentUrl());
    // Whatever the last attempt said no longer applies.
    el.dsAgentOut.innerHTML = '';
    try {
      const health = await api('/api/health');
      agentOk = true;
      setStatus(el.dsAgentStatus, `Connected · v${health.version} · node ${health.node}`, 'ok');
      const { config } = await api('/api/config');
      agentConfig = config;
      applyConfig(config);
      await refreshDevices();
      refreshLanguages();
    } catch (e) {
      agentOk = false;
      setStatus(el.dsAgentStatus, 'Not running', 'bad');
      // This is about the agent connection, so it belongs in the agent panel.
      // api() already fills that in for an unreachable agent; anything else
      // (a bad response, a wrong URL that resolves) needs its own line.
      if (!el.dsAgentOut.innerHTML) showAgentError(`Could not reach the agent at ${agentUrl()}: ${e.message}`);
      updateStartState();
    }
  }

  function applyConfig(c) {
    el.dsKeyState.textContent = c.apiKeySet
      ? `Key configured ${c.apiKeyHint}${c.apiKeyFromEnv ? ' (from ANTHROPIC_API_KEY)' : ''}.`
      : 'No API key set — the scan cannot run without one.';
    el.dsKeyState.className = 'ds-hint ' + (c.apiKeySet ? 'ok' : 'bad');
    if (c.model) el.dsModel.value = c.model;
    if (c.effort) el.dsEffort.value = c.effort;
    el.dsAdbPath.value = c.adbPath || '';
    el.dsPackage.value = c.androidPackage || '';
    el.dsMaxScreens.value = c.maxScreens;
    el.dsMaxActions.value = c.maxActions;
    el.dsMaxDepth.value = c.maxDepth;
    el.dsVision.checked = c.visionEnabled !== false;
    el.dsScroll.checked = c.scrollProbe !== false;
    el.dsLongpress.checked = c.longPressProbe !== false;
    el.dsBlocked.value = (c.blockedLabels || []).join('\n');
    updateStartState();
  }

  async function saveKey() {
    const key = el.dsApiKey.value.trim();
    if (!key) {
      toast('Paste your Anthropic API key first.', 'error');
      return;
    }
    // An unreachable agent takes a couple of seconds to refuse the connection,
    // so say something immediately rather than looking like a dead button.
    const label = el.dsSaveKey.textContent;
    el.dsSaveKey.disabled = true;
    el.dsSaveKey.textContent = 'Saving…';
    try {
      const { config } = await api('/api/config', { method: 'POST', body: JSON.stringify({ apiKey: key }) });
      el.dsApiKey.value = '';
      agentConfig = config;
      applyConfig(config);
      toast('API key saved to the agent.');
    } catch (e) {
      toast(e.message, 'error');
    } finally {
      el.dsSaveKey.disabled = false;
      el.dsSaveKey.textContent = label;
    }
  }

  async function refreshDevices() {
    if (!agentOk) return;
    try {
      const { devices, error } = await api('/api/devices');
      el.dsDevice.innerHTML = '';
      if (error || !devices.length) {
        el.dsDevice.innerHTML = `<option value="">${escapeHtml(error || 'No devices found')}</option>`;
      } else {
        for (const d of devices) {
          const opt = document.createElement('option');
          opt.value = d.serial;
          opt.textContent = `${d.model || d.serial} — ${d.state}`;
          if (d.state !== 'device') opt.disabled = true;
          el.dsDevice.appendChild(opt);
        }
      }
    } catch (e) {
      el.dsDevice.innerHTML = `<option value="">${escapeHtml(e.message)}</option>`;
    }
    updateStartState();
  }

  function onModeChange() {
    const isDevice = el.dsMode.value === 'device';
    el.dsDevice.disabled = !isDevice;
    el.dsPackage.disabled = !isDevice;
    if (isDevice) refreshDevices();
    updateStartState();
  }

  // ── sheet ───────────────────────────────────────────────────────────────

  function sheet() {
    return window.LocaLinter && window.LocaLinter.getSheet ? window.LocaLinter.getSheet() : null;
  }

  // ── which sheet to compare against ──────────────────────────────────────

  function api_() { return window.LocaLinter || {}; }

  function refreshSheets() {
    const L = api_();
    if (!el.dsSheet || !L.getPresets) return;
    const presets = L.getPresets();
    const currentKey = L.getCurrentPresetKey ? L.getCurrentPresetKey() : null;
    const s = sheet();

    el.dsSheet.innerHTML = '';
    const loaded = document.createElement('option');
    loaded.value = '';
    loaded.textContent = s ? `Loaded sheet — ${s.name}` : 'No sheet loaded';
    el.dsSheet.appendChild(loaded);

    presets.forEach((p) => {
      const opt = document.createElement('option');
      opt.value = p.key;
      opt.textContent = p.label;
      el.dsSheet.appendChild(opt);
    });

    el.dsSheet.value = currentKey && presets.some((p) => p.key === currentKey) ? currentKey : '';
    refreshLanguages();
  }

  async function onSheetChange() {
    const key = el.dsSheet.value;
    if (!key) return;
    const L = api_();
    if (!L.loadPreset) return;
    const label = el.dsSheet.options[el.dsSheet.selectedIndex].textContent;
    el.dsSheet.disabled = true;
    setStatus(el.dsTargetStatus, `Loading ${label}…`, 'pending');
    try {
      await L.loadPreset(key);
    } catch (e) {
      toast(e.message || `Could not load ${label}.`, 'error');
      setStatus(el.dsTargetStatus, `Could not load ${label}`, 'bad');
    } finally {
      el.dsSheet.disabled = false;
      refreshSheets();
    }
  }

  function refreshLanguages() {
    const s = sheet();
    const prev = el.dsLanguage.value;
    el.dsLanguage.innerHTML = '';
    if (!s) {
      el.dsLanguage.innerHTML = '<option value="">Load a sheet first</option>';
      setStatus(el.dsTargetStatus, 'No sheet loaded', 'bad');
      updateStartState();
      return;
    }
    s.headers.slice(1).forEach((h) => {
      if (!h || !h.trim()) return;
      const opt = document.createElement('option');
      opt.value = h;
      opt.textContent = h;
      el.dsLanguage.appendChild(opt);
    });
    if (prev && [...el.dsLanguage.options].some((o) => o.value === prev)) el.dsLanguage.value = prev;
    setStatus(el.dsTargetStatus, `${s.rows.length} rows · ${el.dsLanguage.options.length} languages`, 'ok');
    updateStartState();
  }

  function updateStartState() {
    const ready =
      agentOk &&
      !!sheet() &&
      !!el.dsLanguage.value &&
      (agentConfig ? agentConfig.apiKeySet : false) &&
      !running &&
      (el.dsMode.value !== 'device' || !!el.dsDevice.value);
    el.dsStart.disabled = !ready;
  }

  function options() {
    return {
      maxScreens: Number(el.dsMaxScreens.value),
      maxActions: Number(el.dsMaxActions.value),
      maxDepth: Number(el.dsMaxDepth.value),
      model: el.dsModel.value,
      effort: el.dsEffort.value,
      visionEnabled: el.dsVision.checked,
      scrollProbe: el.dsScroll.checked,
      longPressProbe: el.dsLongpress.checked,
      androidPackage: el.dsPackage.value.trim(),
      blockedLabels: el.dsBlocked.value.split('\n').map((s) => s.trim()).filter(Boolean),
    };
  }

  // ── run control ─────────────────────────────────────────────────────────

  async function probe() {
    el.dsProbeOut.innerHTML = '<div class="ds-msg">Testing…</div>';
    try {
      // persist the connection-shaped settings so the run uses them
      await api('/api/config', {
        method: 'POST',
        body: JSON.stringify({ adbPath: el.dsAdbPath.value.trim(), androidPackage: el.dsPackage.value.trim() }),
      });
      const r = await api('/api/probe', {
        method: 'POST',
        body: JSON.stringify({ mode: el.dsMode.value, serial: el.dsDevice.value || null }),
      });
      const parts = [];
      if (r.device) {
        parts.push(`<div class="ds-msg ok">Device ${escapeHtml(r.device.serial || '')} · ${r.device.size.width}×${r.device.size.height}` +
          `${r.device.package ? ` · foreground: <code>${escapeHtml(r.device.package)}</code>` : ''}` +
          `${r.device.locale ? ` · locale ${escapeHtml(r.device.locale)}` : ''}</div>`);
      }
      if (r.bridge) {
        parts.push(`<div class="ds-msg ok">In-game bridge connected — ${escapeHtml(r.bridge.product || '')} ` +
          `(${escapeHtml(r.bridge.mode)}, Unity ${escapeHtml(r.bridge.unity || '?')}). Exact strings and rects are available.</div>`);
      } else {
        parts.push('<div class="ds-msg warn">No in-game bridge. The scan will read text from screenshots instead — ' +
          'add <code>agent/unity/LocaLinterBridge.cs</code> to the Unity project for exact strings, measured overflow and reliable clicking.</div>');
      }
      for (const e of r.errors || []) parts.push(`<div class="ds-msg warn">${escapeHtml(e)}</div>`);
      el.dsProbeOut.innerHTML = parts.join('');
      if (r.device && r.device.package && !el.dsPackage.value) el.dsPackage.value = r.device.package;
    } catch (e) {
      el.dsProbeOut.innerHTML = `<div class="ds-msg bad">${escapeHtml(e.message)}</div>`;
    }
  }

  async function start() {
    const s = sheet();
    if (!s) return toast('Load a localization sheet first.', 'error');

    issues = [];
    screens = new Map();
    usageSent = null;
    el.dsFindings.innerHTML = '';
    el.dsLog.innerHTML = '';
    el.dsSummary.classList.add('hidden');
    el.dsRunPanel.classList.remove('hidden');
    el.dsResultsPanel.classList.remove('hidden');
    el.dsFilterType.innerHTML = '<option value="all">All types</option>';
    setStatus(el.dsRunStatus, 'Starting…', 'pending');

    try {
      await api('/api/config', {
        method: 'POST',
        body: JSON.stringify({ adbPath: el.dsAdbPath.value.trim(), ...options() }),
      });
      const r = await api('/api/run/start', {
        method: 'POST',
        body: JSON.stringify({
          mode: el.dsMode.value,
          serial: el.dsDevice.value || null,
          targetLanguage: el.dsLanguage.value,
          sheet: s,
          options: options(),
        }),
      });
      runId = r.runId;
      running = true;
      el.dsStop.classList.remove('hidden');
      updateStartState();
      log(`Scanning ${r.language} in ${r.mode} mode.`, 'info');
      listen();
    } catch (e) {
      setStatus(el.dsRunStatus, 'Failed to start', 'bad');
      log(e.message, 'error');
      toast(e.message, 'error');
    }
  }

  async function stop() {
    if (!runId) return;
    try {
      await api(`/api/run/${runId}/stop`, { method: 'POST' });
      log('Stop requested — finishing the current screen.', 'warn');
    } catch (e) {
      toast(e.message, 'error');
    }
  }

  function listen() {
    if (source) source.close();
    source = new EventSource(`${agentUrl()}/api/run/${runId}/events`);
    source.onmessage = (e) => {
      let ev;
      try { ev = JSON.parse(e.data); } catch { return; }
      handleEvent(ev);
    };
    source.onerror = () => {
      if (!running) return;
      setStatus(el.dsRunStatus, 'Connection to the agent dropped', 'bad');
    };
  }

  /**
   * The agent reports usage as a running total, so forward the difference since
   * the last report. The budget meter can then add up deltas without ever
   * counting the same tokens twice.
   */
  function reportUsage(usage, { final = false } = {}) {
    if (!usage) return;
    const zero = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, calls: 0, costUSD: 0 };
    const prev = usageSent || zero;
    const delta = {};
    let changed = false;
    for (const k of Object.keys(zero)) {
      delta[k] = (usage[k] || 0) - (prev[k] || 0);
      if (delta[k] > 0) changed = true;
    }
    usageSent = { ...zero, ...usage };
    if (!changed && !final) return;
    delta.priced = usage.priced !== false;
    delta.model = el.dsModel ? el.dsModel.value : '';
    delta.runFinished = final;
    document.dispatchEvent(new CustomEvent('localinter:usage', { detail: delta }));
  }

  function handleEvent(ev) {
    switch (ev.type) {
      case 'status':
        setStatus(el.dsRunStatus, `Running (${ev.mode})`, 'pending');
        break;
      case 'log':
        log(ev.message, ev.level);
        break;
      case 'progress':
        el.dsStatScreens.textContent = `Screens: ${ev.screens}`;
        el.dsStatActions.textContent = `Actions: ${ev.actions}`;
        el.dsStatIssues.textContent = `Issues: ${ev.issues}`;
        el.dsStatQueue.textContent = `Queued: ${ev.queued}`;
        reportUsage(ev.usage);
        break;
      case 'action':
        setStatus(el.dsRunStatus, `Tapping “${truncate(ev.action.label, 40)}”`, 'pending');
        break;
      case 'screen':
        screens.set(ev.screen.id, ev.screen);
        renderFindings();
        break;
      case 'issues':
        issues.push(...ev.issues);
        el.badgeDevice.textContent = String(issues.length);
        el.dsStatIssues.textContent = `Issues: ${issues.length}`;
        renderFindings();
        break;
      case 'done':
        running = false;
        el.dsStop.classList.add('hidden');
        updateStartState();
        if (source) source.close();
        setStatus(
          el.dsRunStatus,
          ev.status === 'done' ? `Finished — ${ev.issues} issues` : `${ev.status}${ev.error ? `: ${ev.error}` : ''}`,
          ev.status === 'done' ? 'ok' : ev.status === 'stopped' ? 'warn' : 'bad'
        );
        finalize();
        break;
    }
  }

  async function finalize() {
    if (!runId) return;
    try {
      const report = await api(`/api/run/${runId}`);
      issues = report.issues;
      screens = new Map(report.screens.map((s) => [s.id, s]));
      if (report.summary) {
        el.dsSummary.textContent = report.summary;
        el.dsSummary.classList.remove('hidden');
      }
      if (report.warnings && report.warnings.length) {
        for (const w of report.warnings) log(w, 'warn');
      }
      if (report.skipped && report.skipped.length) {
        log(`${report.skipped.length} controls were not tapped because they matched a blocked-label pattern.`, 'warn');
      }
      if (report.usage) {
        const u = report.usage;
        const cost = u.priced === false ? '' : ` ≈ $${(u.costUSD || 0).toFixed(4)}`;
        log(`Claude usage: ${u.calls} calls, ${u.input} in / ${u.output} out tokens (${u.cacheRead} cached)${cost}.`, 'info');
        reportUsage(u, { final: true });
      }
      el.badgeDevice.textContent = String(issues.length);
      renderFindings();
    } catch (e) {
      log(`Could not load the final report: ${e.message}`, 'warn');
    }
  }

  // ── rendering ───────────────────────────────────────────────────────────

  const SEVERITY_RANK = { high: 3, medium: 2, low: 1 };

  function renderFindings() {
    // keep the type filter in sync with what we have actually seen
    const types = [...new Set(issues.map((i) => i.type))].sort();
    const current = el.dsFilterType.value;
    if (el.dsFilterType.options.length !== types.length + 1) {
      el.dsFilterType.innerHTML = '<option value="all">All types</option>';
      for (const t of types) {
        const o = document.createElement('option');
        o.value = t;
        o.textContent = t.replace(/_/g, ' ');
        el.dsFilterType.appendChild(o);
      }
      if (types.includes(current)) el.dsFilterType.value = current;
    }

    const minSev = el.dsFilterSeverity.value === 'high' ? 3 : el.dsFilterSeverity.value === 'medium' ? 2 : 1;
    const typeFilter = el.dsFilterType.value;
    const q = el.dsFilterText.value.trim().toLowerCase();

    const shown = issues.filter((i) => {
      if ((SEVERITY_RANK[i.severity] || 1) < minSev) return false;
      if (typeFilter !== 'all' && i.type !== typeFilter) return false;
      if (q) {
        const hay = `${i.text} ${i.message} ${i.element || ''} ${i.key || ''} ${i.screenId}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });

    if (!shown.length) {
      el.dsFindings.innerHTML = issues.length
        ? '<div class="ds-empty">No findings match this filter.</div>'
        : '<div class="ds-empty">No issues yet.</div>';
      return;
    }

    // group by screen, screens in capture order
    const byScreen = new Map();
    for (const i of shown) {
      if (!byScreen.has(i.screenId)) byScreen.set(i.screenId, []);
      byScreen.get(i.screenId).push(i);
    }

    const html = [...byScreen.entries()]
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([screenId, list]) => {
        const screen = screens.get(screenId);
        const shot = screen ? `${agentUrl()}/api/run/${runId}/screenshot/${screen.file}` : '';
        list.sort((a, b) => (SEVERITY_RANK[b.severity] || 0) - (SEVERITY_RANK[a.severity] || 0));
        return `
          <div class="ds-screen">
            <div class="ds-screen-shot">
              ${shot ? `<img src="${shot}" alt="${escapeHtml(screenId)}" data-shot="${shot}" data-caption="${escapeHtml((screen && screen.summary) || screenId)}" loading="lazy" />` : ''}
            </div>
            <div class="ds-screen-body">
              <div class="ds-screen-head">
                <strong>${escapeHtml(screenId)}</strong>
                <span class="ds-screen-meta">${escapeHtml((screen && screen.summary) || '')}</span>
              </div>
              ${screen && screen.path && screen.path.length
            ? `<div class="ds-path">Reached via: ${screen.path.map((p) => `<span>${escapeHtml(truncate(p, 28))}</span>`).join(' › ')}</div>`
            : ''}
              <div class="ds-issue-list">
                ${list.map(issueRow).join('')}
              </div>
            </div>
          </div>`;
      })
      .join('');

    el.dsFindings.innerHTML = html;
    el.dsFindings.querySelectorAll('img[data-shot]').forEach((img) => {
      img.addEventListener('click', () => {
        el.dsShotImg.src = img.dataset.shot;
        el.dsShotCaption.textContent = img.dataset.caption;
        el.dsShotOverlay.classList.remove('hidden');
      });
    });
  }

  function issueRow(i) {
    const bits = [];
    if (i.key) bits.push(`key <code>${escapeHtml(i.key)}</code>`);
    if (i.row) bits.push(`row ${i.row}`);
    if (i.element) bits.push(`<span class="ds-elem">${escapeHtml(truncate(i.element, 60))}</span>`);
    if (i.where && !i.element) bits.push(escapeHtml(i.where));
    return `
      <div class="ds-issue sev-${escapeHtml(i.severity)}">
        <div class="ds-issue-top">
          <span class="ds-type">${escapeHtml(i.type.replace(/_/g, ' '))}</span>
          <span class="ds-sev">${escapeHtml(i.severity)}</span>
          <span class="ds-src" title="${i.source === 'vision' ? 'Found by Claude vision' : 'Found by a deterministic check'}">${i.source === 'vision' ? 'vision' : 'check'}${i.alsoSeenBy ? ' + vision' : ''}</span>
          ${i.confidence ? `<span class="ds-conf">${escapeHtml(i.confidence)}</span>` : ''}
        </div>
        <div class="ds-issue-text">“${escapeHtml(truncate(i.text, 200))}”</div>
        ${i.expected ? `<div class="ds-issue-expected">Sheet says: “${escapeHtml(truncate(i.expected, 200))}”</div>` : ''}
        <div class="ds-issue-msg">${escapeHtml(i.message)}</div>
        ${bits.length ? `<div class="ds-issue-meta">${bits.join(' · ')}</div>` : ''}
      </div>`;
  }

  function exportJson() {
    const payload = {
      runId,
      exportedAt: new Date().toISOString(),
      language: el.dsLanguage.value,
      screens: [...screens.values()],
      issues,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `device-scan-${runId || 'report'}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  // ── small helpers ───────────────────────────────────────────────────────

  function log(message, level = 'info') {
    const div = document.createElement('div');
    div.className = `ds-log-line ${level}`;
    div.textContent = `${new Date().toLocaleTimeString()}  ${message}`;
    el.dsLog.appendChild(div);
    el.dsLog.scrollTop = el.dsLog.scrollHeight;
  }

  function setStatus(node, text, kind) {
    if (!node) return;
    node.textContent = text;
    node.className = `ds-status ${kind || ''}`;
  }

  function toast(message, type) {
    if (window.LocaLinter && window.LocaLinter.showToast) window.LocaLinter.showToast(message, type);
    else console.log(message);
  }

  function truncate(s, n) {
    s = String(s == null ? '' : s);
    return s.length > n ? `${s.slice(0, n - 1)}…` : s;
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
    ));
  }
})();
