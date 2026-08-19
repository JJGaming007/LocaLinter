/**
 * Device Scan tab.
 *
 * A window cannot speak ADB or hold an API key safely, so the real work happens
 * in the app's own service (agent/server.js) running in the main process. This
 * file is the control panel: it hands the already-loaded sheet over, streams the
 * run back via SSE, and renders the findings.
 */
(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);

  const el = {};
  let agentOk = false;
  let agentConfig = null;
  let runId = null;
  let source = null;           // EventSource
  let issues = [];
  let screens = new Map();
  let running = false;
  let currentLanguage = '';

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    [
      'ds-agent-status', 'ds-api-key', 'ds-save-key', 'ds-key-state',
      'ds-adb-row', 'ds-adb-state', 'ds-get-adb',
      'ds-mode', 'ds-device', 'ds-sheet', 'ds-language', 'ds-package', 'ds-build', 'ds-route', 'ds-route-card', 'ds-probe', 'ds-start', 'ds-stop',
      'ds-advanced-toggle', 'ds-advanced', 'ds-probe-out', 'ds-agent-out', 'ds-target-status',
      'ds-max-screens', 'ds-max-actions', 'ds-max-depth', 'ds-model', 'ds-effort', 'ds-adb-path',
      'ds-vision', 'ds-scroll', 'ds-longpress', 'ds-blocked', 'ds-base-url', 'ds-extra-checks',
      'ds-settle', 'ds-scroll-steps', 'ds-longpress-ms', 'ds-max-minutes', 'ds-stop-high',
      'ds-route-lang', 'ds-only', 'ds-focus', 'ds-rules', 'ds-steps',
      'ds-validate', 'ds-validate-out', 'ds-reset-defaults', 'ds-lang-queue', 'ds-queue-wrap',
      'ds-profile', 'ds-profile-save', 'ds-profile-update', 'ds-profile-delete', 'ds-profile-hint',
      'ds-run-panel', 'ds-run-status', 'ds-stat-screens', 'ds-stat-actions', 'ds-stat-issues',
      'ds-stat-queue', 'ds-summary', 'ds-log', 'ds-results-panel', 'ds-findings',
      'ds-filter-severity', 'ds-filter-type', 'ds-filter-text', 'ds-export', 'ds-clear',
      'ds-explain', 'ds-explain-btn', 'ds-explain-out',
      'ds-filter-severity', 'ds-filter-type', 'ds-filter-text', 'ds-export', 'ds-export-csv',
      'ds-shot-overlay', 'ds-shot-img', 'ds-shot-caption', 'ds-shot-close', 'badge-device',
      'ds-blockers', 'ds-blockers-list', 'ds-device-refresh',
    ].forEach((id) => { el[id.replace(/-([a-z])/g, (_, c) => c.toUpperCase())] = $(id); });

    if (!el.dsAgentStatus) return; // tab not present

    el.dsSaveKey.addEventListener('click', saveKey);
    el.dsGetAdb.addEventListener('click', downloadAdb);
    el.dsAdvancedToggle.addEventListener('click', () => {
      const open = !el.dsAdvanced.classList.toggle('hidden');
      // A toggle that looks identical open and shut tells you nothing.
      el.dsAdvancedToggle.classList.toggle('is-open', open);
      el.dsAdvancedToggle.setAttribute('aria-expanded', String(open));
    });
    el.dsMode.addEventListener('change', onModeChange);
    el.dsProbe.addEventListener('click', probe);
    el.dsStart.addEventListener('click', start);
    el.dsStop.addEventListener('click', stop);
    el.dsExport.addEventListener('click', exportJson);
    el.dsClear.addEventListener('click', clearHistory);
    el.dsExplainBtn.addEventListener('click', explainString);
    el.dsExplain.addEventListener('keydown', (e) => { if (e.key === 'Enter') explainString(); });
    el.dsExportCsv.addEventListener('click', exportCsv);
    el.dsValidate.addEventListener('click', validateAutomation);
    el.dsResetDefaults.addEventListener('click', resetToDefaults);
    el.dsProfile.addEventListener('change', onProfilePick);
    el.dsProfileSave.addEventListener('click', saveProfileAs);
    el.dsProfileUpdate.addEventListener('click', updateProfile);
    el.dsProfileDelete.addEventListener('click', deleteProfile);
    // Whatever the panel was last set to is what the next scan almost always
    // wants, so it survives a reload without anyone pressing Save.
    SETTING_IDS.forEach((id) => {
      const node = el[camel(id)];
      if (node) node.addEventListener('change', rememberSettings);
    });
    el.dsShotClose.addEventListener('click', () => el.dsShotOverlay.classList.add('hidden'));
    el.dsShotOverlay.addEventListener('click', (e) => {
      if (e.target === el.dsShotOverlay) el.dsShotOverlay.classList.add('hidden');
    });
    [el.dsFilterSeverity, el.dsFilterType].forEach((s) => s.addEventListener('change', renderFindings));
    el.dsFilterText.addEventListener('input', renderFindings);

    fillBuilds();
    el.dsBuild.addEventListener('change', () => {
      if (el.dsBuild.value) el.dsPackage.value = el.dsBuild.value;
      updateStartState();
    });
    el.dsPackage.addEventListener('input', syncBuildSelect);

    // Picking a device or a language column is exactly what clears the last
    // blocker, so both have to re-run the readiness check.
    el.dsDevice.addEventListener('change', updateStartState);
    el.dsLanguage.addEventListener('change', updateStartState);

    el.dsSheet.addEventListener('change', onSheetChange);
    el.dsRoute.addEventListener('change', renderRouteCard);
    el.dsLanguage.addEventListener('change', () => { renderLangQueue(); updateStartState(); });

    if (el.dsDeviceRefresh) el.dsDeviceRefresh.addEventListener('click', refreshDevices);

    // In the web build the agent was a separate process the user had to start,
    // so contacting it at load only painted a failure nobody had asked about.
    // In the desktop app it runs inside this very process and is always up, so
    // waiting meant the status bar claimed the service was stopped and the
    // device list stayed empty until you happened to open this tab. Connect
    // once at startup instead; the per-tab retry below still covers a service
    // that has genuinely fallen over.
    connectAndLoadRoutes();

    document.addEventListener('localinter:tab', (e) => {
      const here = e.detail.tabId === 'device-scan';
      setDevicePolling(here);
      if (!here) return;
      refreshSheets();
      if (!agentOk && !running) connectAndLoadRoutes();
      // Already connected from an earlier visit: re-check rather than trusting
      // a device list that may be minutes old.
      else if (!running) refreshDevices();
    });
    // Loading a sheet anywhere in the app re-points the scan at it.
    document.addEventListener('localinter:sheet', () => {
      refreshSheets();
      // The sheet and the build are two halves of one environment.
      const key = window.LocaLinter && window.LocaLinter.getCurrentPresetKey && window.LocaLinter.getCurrentPresetKey();
      const preset = key && (window.LocaLinter.getPresets() || []).find((p) => p.key === key);
      if (preset && preset.pkg && el.dsPackage.value.trim() !== preset.pkg) {
        el.dsPackage.value = preset.pkg;
        syncBuildSelect();
        log(`Sheet is ${preset.label}, so the scan will target ${preset.pkg}.`, 'info');
      }
    });

    renderProfiles();
    restoreSettings();
    onModeChange();
    refreshSheets();
    if (document.getElementById('tab-device-scan').classList.contains('active')) {
      setDevicePolling(true);
    }
  }

  /**
   * The same game ships as five builds, one per environment, and each has its
   * own package — and its own memory. Choosing from the list is the difference
   * between scanning Latam and teaching the Global Prod memory by mistake.
   */
  function fillBuilds() {
    const presets = (window.LocaLinter && window.LocaLinter.getPresets) ? window.LocaLinter.getPresets() : [];
    for (const p of presets.filter((x) => x.pkg)) {
      const o = document.createElement('option');
      o.value = p.pkg;
      o.textContent = `${p.label} — ${p.pkg}`;
      el.dsBuild.appendChild(o);
    }
    syncBuildSelect();
  }

  /** Keep the dropdown honest when the package is typed or loaded from config. */
  function syncBuildSelect() {
    if (!el.dsBuild) return;
    const pkg = (el.dsPackage.value || '').trim();
    const match = [...el.dsBuild.options].some((o) => o.value === pkg);
    el.dsBuild.value = match ? pkg : '';
  }

  function camel(id) {
    return id.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
  }

  // ── settings, profiles ──────────────────────────────────────────────────
  // Everything a tester tunes lives in one list, so remembering it, exporting
  // it as a profile and resetting it are the same three lines rather than
  // three places to forget a new field.

  const SETTING_IDS = [
    'ds-mode', 'ds-package', 'ds-route', 'ds-extra-checks',
    'ds-max-screens', 'ds-max-actions', 'ds-max-depth', 'ds-model', 'ds-effort',
    'ds-base-url', 'ds-adb-path', 'ds-settle', 'ds-scroll-steps', 'ds-longpress-ms',
    'ds-max-minutes', 'ds-stop-high',
    'ds-vision', 'ds-scroll', 'ds-longpress', 'ds-route-lang',
    'ds-blocked', 'ds-only', 'ds-focus', 'ds-rules', 'ds-steps',
  ];
  const PROFILE_KEY = 'localinter_ds_profiles';
  const LAST_KEY = 'localinter_ds_last';
  const PROFILE_PICK_KEY = 'localinter_ds_profile';

  // Three starting points that cover most of what people actually ask for.
  // Only the fields a preset cares about are listed; the rest stay as they are.
  const BUILT_IN = {
    '': { label: 'Custom (whatever is set below)' },
    quick: {
      label: 'Quick smoke — first screens, cheap',
      values: {
        'ds-max-screens': 15, 'ds-max-actions': 50, 'ds-max-depth': 3,
        'ds-model': 'claude-sonnet-5', 'ds-effort': 'medium',
        'ds-scroll-steps': 2, 'ds-max-minutes': 10, 'ds-stop-high': 0,
        'ds-vision': true, 'ds-scroll': true, 'ds-longpress': false,
      },
    },
    standard: {
      label: 'Standard pass — balanced',
      values: {
        'ds-max-screens': 120, 'ds-max-actions': 400, 'ds-max-depth': 12,
        'ds-model': 'claude-opus-5', 'ds-effort': 'high',
        'ds-scroll-steps': 4, 'ds-max-minutes': 0, 'ds-stop-high': 0,
        'ds-vision': true, 'ds-scroll': true, 'ds-longpress': true,
      },
    },
    deep: {
      label: 'Deep audit — everything, slowly',
      values: {
        'ds-max-screens': 400, 'ds-max-actions': 2000, 'ds-max-depth': 25,
        'ds-model': 'claude-opus-5', 'ds-effort': 'xhigh',
        'ds-scroll-steps': 8, 'ds-settle': 1400, 'ds-max-minutes': 0, 'ds-stop-high': 0,
        'ds-vision': true, 'ds-scroll': true, 'ds-longpress': true,
      },
    },
    triage: {
      label: 'Triage — stop at the first ten serious findings',
      values: {
        'ds-max-screens': 200, 'ds-max-actions': 600, 'ds-max-depth': 12,
        'ds-model': 'claude-sonnet-5', 'ds-effort': 'high',
        'ds-stop-high': 10, 'ds-max-minutes': 30,
        'ds-vision': true, 'ds-scroll': true, 'ds-longpress': false,
      },
    },
  };

  function readSettings() {
    const out = {};
    SETTING_IDS.forEach((id) => {
      const node = el[camel(id)];
      if (!node) return;
      out[id] = node.type === 'checkbox' ? node.checked : node.value;
    });
    return out;
  }

  function writeSettings(values) {
    if (!values) return;
    Object.entries(values).forEach(([id, v]) => {
      const node = el[camel(id)];
      if (!node) return;
      if (node.type === 'checkbox') node.checked = !!v;
      else node.value = v;
    });
    onModeChange();
  }

  function rememberSettings() {
    try { localStorage.setItem(LAST_KEY, JSON.stringify(readSettings())); } catch { /* private mode */ }
  }

  function rememberedSetting(id) {
    try {
      const raw = localStorage.getItem(LAST_KEY);
      return raw ? (JSON.parse(raw) || {})[id] : null;
    } catch { return null; }
  }

  function restoreSettings() {
    try {
      const raw = localStorage.getItem(LAST_KEY);
      if (raw) writeSettings(JSON.parse(raw));
    } catch { /* nothing worth reporting */ }
    const picked = localStorage.getItem(PROFILE_PICK_KEY);
    if (picked && el.dsProfile.querySelector(`option[value="${CSS.escape(picked)}"]`)) el.dsProfile.value = picked;
  }

  function customProfiles() {
    try { return JSON.parse(localStorage.getItem(PROFILE_KEY)) || {}; } catch { return {}; }
  }

  function saveCustomProfiles(map) {
    try { localStorage.setItem(PROFILE_KEY, JSON.stringify(map)); } catch {
      toast('This browser will not let the page store profiles.', 'error');
    }
  }

  function renderProfiles(select) {
    const custom = customProfiles();
    const prev = select || el.dsProfile.value;
    el.dsProfile.innerHTML = '';
    const add = (value, label, group) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = label;
      (group || el.dsProfile).appendChild(opt);
    };
    Object.entries(BUILT_IN).forEach(([k, p]) => add(k, p.label));
    const names = Object.keys(custom).sort();
    if (names.length) {
      const group = document.createElement('optgroup');
      group.label = 'Saved';
      names.forEach((n) => add(`custom:${n}`, n, group));
      el.dsProfile.appendChild(group);
    }
    if (prev && [...el.dsProfile.querySelectorAll('option')].some((o) => o.value === prev)) el.dsProfile.value = prev;
    updateProfileButtons();
  }

  function updateProfileButtons() {
    const isCustom = el.dsProfile.value.startsWith('custom:');
    el.dsProfileUpdate.disabled = !isCustom;
    el.dsProfileDelete.disabled = !isCustom;
  }

  function onProfilePick() {
    const value = el.dsProfile.value;
    localStorage.setItem(PROFILE_PICK_KEY, value);
    updateProfileButtons();
    if (!value) return;
    if (value.startsWith('custom:')) {
      const saved = customProfiles()[value.slice(7)];
      if (!saved) return;
      writeSettings(saved.values);
      el.dsProfileHint.textContent = `Loaded "${value.slice(7)}".`;
    } else {
      const preset = BUILT_IN[value];
      if (!preset || !preset.values) return;
      writeSettings(preset.values);
      el.dsProfileHint.textContent = `${preset.label} — limits and model set; your patterns and steps are untouched.`;
    }
    rememberSettings();
  }

  function saveProfileAs() {
    const name = (prompt('Name this profile', 'My scan') || '').trim();
    if (!name) return;
    const map = customProfiles();
    map[name] = { values: readSettings(), savedAt: Date.now() };
    saveCustomProfiles(map);
    renderProfiles(`custom:${name}`);
    el.dsProfile.value = `custom:${name}`;
    localStorage.setItem(PROFILE_PICK_KEY, el.dsProfile.value);
    updateProfileButtons();
    el.dsProfileHint.textContent = `Saved "${name}". It covers every field in Advanced, including your checks and steps.`;
    toast(`Profile "${name}" saved.`);
  }

  function updateProfile() {
    const value = el.dsProfile.value;
    if (!value.startsWith('custom:')) return;
    const name = value.slice(7);
    const map = customProfiles();
    map[name] = { values: readSettings(), savedAt: Date.now() };
    saveCustomProfiles(map);
    el.dsProfileHint.textContent = `Updated "${name}".`;
    toast(`Profile "${name}" updated.`);
  }

  function deleteProfile() {
    const value = el.dsProfile.value;
    if (!value.startsWith('custom:')) return;
    const name = value.slice(7);
    if (!confirm(`Delete the profile "${name}"?`)) return;
    const map = customProfiles();
    delete map[name];
    saveCustomProfiles(map);
    renderProfiles('');
    el.dsProfile.value = '';
    el.dsProfileHint.textContent = `Deleted "${name}".`;
  }

  // Back to the agent's own defaults rather than to whatever this browser last
  // held, which is what "reset" has to mean when a run has gone strange.
  function resetToDefaults() {
    if (!agentConfig) {
      writeSettings(BUILT_IN.standard.values);
    } else {
      applyConfig(agentConfig);
      writeSettings(BUILT_IN.standard.values);
    }
    el.dsOnly.value = '';
    el.dsFocus.value = '';
    el.dsRules.value = '';
    el.dsSteps.value = '';
    rememberSettings();
    el.dsProfileHint.textContent = 'Back to the defaults the agent shipped with.';
  }

  async function validateAutomation() {
    el.dsValidateOut.innerHTML = '<div class="ds-msg">Checking…</div>';
    try {
      const r = await api('/api/validate', {
        method: 'POST',
        body: JSON.stringify({ customRules: el.dsRules.value, preSteps: el.dsSteps.value }),
      });
      const parts = [];
      const errors = [...r.rules.errors, ...r.steps.errors];
      parts.push(`<div class="ds-msg ${errors.length ? 'warn' : 'ok'}">${r.rules.count} custom check${
        r.rules.count === 1 ? '' : 's'} and ${r.steps.count} setup step${r.steps.count === 1 ? '' : 's'} understood.</div>`);
      for (const e of errors) parts.push(`<div class="ds-msg bad">${escapeHtml(e)}</div>`);
      el.dsValidateOut.innerHTML = parts.join('');
    } catch (e) {
      el.dsValidateOut.innerHTML = `<div class="ds-msg bad">${escapeHtml(e.message)}</div>`;
    }
  }

  async function connectAndLoadRoutes() {
    await connect();
    if (agentOk) loadRoutes();
  }

  // ── agent ───────────────────────────────────────────────────────────────

  // Same-origin: the app serves both the UI and this API from one process.
  function agentUrl() {
    return '';
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
      setStatus(el.dsAgentStatus, 'Stopped', 'bad');
      showAgentError('The scanning service stopped responding.');
      updateStartState();
      throw new Error('The scanning service stopped responding. Restart LocaLinter.');
    }
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `agent returned ${res.status}`);
    return body;
  }

  function showAgentError(headline) {
    el.dsAgentOut.innerHTML = `<div class="ds-msg bad">${escapeHtml(headline)}</div>`;
  }


  async function connect() {
    setStatus(el.dsAgentStatus, 'Connecting…', 'pending');
    // Whatever the last attempt said no longer applies.
    el.dsAgentOut.innerHTML = '';
    try {
      const health = await api('/api/health');
      agentOk = true;
      setStatus(el.dsAgentStatus, `Connected · v${health.version} · node ${health.node}`, 'ok');
      const { config } = await api('/api/config');
      agentConfig = config;
      applyConfig(config);
      // The agent's stored config is the baseline; anything the tester has
      // since typed into this panel is what they are actually looking at, so
      // it wins over what connecting just painted in.
      restoreSettings();
      await refreshAdb();
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
    // A stored key that cannot possibly work should say so here, not surface
    // as a 401 a minute into a scan.
    if (c.apiKeySet && c.apiKeyWarning) {
      el.dsKeyState.textContent = `${c.apiKeyWarning} Get one from console.anthropic.com.`;
      el.dsKeyState.className = 'ds-hint bad';
    } else {
      el.dsKeyState.textContent = c.apiKeySet
        ? `Key configured ${c.apiKeyHint}${c.apiKeyFromEnv ? ' (from ANTHROPIC_API_KEY)' : ''}.`
        : 'No API key set — the scan cannot run without one.';
      el.dsKeyState.className = 'ds-hint ' + (c.apiKeySet ? 'ok' : 'bad');
    }
    if (c.model) el.dsModel.value = c.model;
    if (c.effort) el.dsEffort.value = c.effort;
    el.dsAdbPath.value = c.adbPath || '';
    el.dsBaseUrl.value = c.baseUrl || '';
    el.dsExtraChecks.value = c.extraChecks || '';
    el.dsPackage.value = c.androidPackage || '';
    syncBuildSelect();
    el.dsMaxScreens.value = c.maxScreens;
    el.dsMaxActions.value = c.maxActions;
    el.dsMaxDepth.value = c.maxDepth;
    el.dsVision.checked = c.visionEnabled !== false;
    el.dsScroll.checked = c.scrollProbe !== false;
    el.dsLongpress.checked = c.longPressProbe !== false;
    el.dsRouteLang.checked = c.routeSetLanguage !== false;
    el.dsBlocked.value = (c.blockedLabels || []).join('\n');
    el.dsOnly.value = (c.onlyLabels || []).join('\n');
    el.dsFocus.value = (c.focusLabels || []).join('\n');
    el.dsRules.value = c.customRules || '';
    el.dsSteps.value = c.preSteps || '';
    if (c.settleMs) el.dsSettle.value = c.settleMs;
    if (c.scrollSteps) el.dsScrollSteps.value = c.scrollSteps;
    if (c.longPressMs) el.dsLongpressMs.value = c.longPressMs;
    el.dsMaxMinutes.value = c.maxMinutes || 0;
    el.dsStopHigh.value = c.stopAfterHighIssues || 0;
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

  // A tester who has never installed the Android SDK has no adb, and the only
  // symptom would otherwise be an empty device list. Say so plainly, and offer
  // to fetch it rather than sending them to Google's download page.
  async function refreshAdb() {
    if (!el.dsAdbRow) return;
    let state;
    try {
      state = await api('/api/tools/adb');
    } catch {
      return; // the agent itself is the problem; that message is already up
    }
    el.dsAdbRow.classList.remove('hidden');
    if (state.found) {
      const where = state.source === 'path' ? 'found on your PATH'
        : state.source === 'downloaded' ? 'downloaded by the agent'
        : 'set in Advanced';
      el.dsAdbState.textContent = `Ready — ${where}.`;
      el.dsAdbState.className = 'ds-hint ok';
      el.dsGetAdb.classList.add('hidden');
    } else {
      el.dsAdbState.textContent = 'Not installed — a scan cannot reach your device without it.';
      el.dsAdbState.className = 'ds-hint bad';
      el.dsGetAdb.classList.remove('hidden');
    }
  }

  async function downloadAdb() {
    const label = el.dsGetAdb.textContent;
    el.dsGetAdb.disabled = true;
    el.dsGetAdb.textContent = 'Downloading…';
    el.dsAdbState.textContent = 'Fetching platform-tools from Google — this takes a moment.';
    el.dsAdbState.className = 'ds-hint';
    try {
      await api('/api/tools/adb', { method: 'POST' });
      toast('ADB installed.');
      await refreshAdb();
      await refreshDevices();
    } catch (e) {
      el.dsAdbState.textContent = e.message;
      el.dsAdbState.className = 'ds-hint bad';
      toast(e.message, 'error');
    } finally {
      el.dsGetAdb.disabled = false;
      el.dsGetAdb.textContent = label;
    }
  }

  /**
   * Devices come and go while the app is open — a cable gets plugged in, the
   * "allow USB debugging" prompt is finally accepted, adb's daemon takes a
   * moment to start on the first call. This used to run once, on connect, so
   * anything that arrived afterwards was never noticed and the dropdown sat on
   * "No devices found" with no way to retry.
   */
  async function refreshDevices() {
    if (!agentOk) return;
    const prev = el.dsDevice.value;
    if (el.dsDeviceRefresh) el.dsDeviceRefresh.classList.add('is-busy');
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
        // Re-selecting what was chosen keeps a poll from stealing the choice.
        if (prev && [...el.dsDevice.options].some((o) => o.value === prev && !o.disabled)) {
          el.dsDevice.value = prev;
        } else {
          const first = [...el.dsDevice.options].find((o) => o.value && !o.disabled);
          if (first) el.dsDevice.value = first.value;
        }
      }
    } catch (e) {
      el.dsDevice.innerHTML = `<option value="">${escapeHtml(e.message)}</option>`;
    } finally {
      if (el.dsDeviceRefresh) el.dsDeviceRefresh.classList.remove('is-busy');
    }
    updateStartState();
  }

  /**
   * Poll only while the Device scan tab is actually on screen and no scan is
   * running — often enough that plugging a cable in feels instant, cheap
   * enough that it costs nothing when the tab is closed.
   */
  let devicePollTimer = null;
  function setDevicePolling(on) {
    if (on && !devicePollTimer) {
      devicePollTimer = setInterval(() => {
        if (agentOk && !running && el.dsMode.value === 'device' && !document.hidden) refreshDevices();
      }, 3000);
    } else if (!on && devicePollTimer) {
      clearInterval(devicePollTimer);
      devicePollTimer = null;
    }
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

  // ── route maps ──────────────────────────────────────────────────────────
  // What earlier passes worked out about a game: its screens, where the info
  // badges that open flyouts sit, how to switch language, how to recover when
  // it strands itself. Shown here so a run starts from that knowledge instead
  // of rediscovering it, and so what the agent knows is visible rather than
  // buried in a file on someone's machine.

  let routeList = [];
  let reconnectAttempts = 0;
  let reconnectTimer = null;

  async function loadRoutes() {
    if (!el.dsRoute) return;
    try {
      const r = await api('/api/routes');
      routeList = r.routes || [];
    } catch {
      routeList = []; // agent offline; the picker just stays empty
    }
    const prev = el.dsRoute.value;
    el.dsRoute.innerHTML = '<option value="">Explore from scratch</option>';
    routeList.forEach((rt) => {
      const opt = document.createElement('option');
      opt.value = rt.name;
      opt.textContent = `${rt.label} — ${rt.screens} screens`;
      el.dsRoute.appendChild(opt);
    });
    // The picker is empty until the agent answers, which is after the panel has
    // restored its settings — so the remembered route is applied here instead.
    const remembered = prev || rememberedSetting('ds-route');
    if (remembered && routeList.some((rt) => rt.name === remembered)) el.dsRoute.value = remembered;
    else autoSelectRoute();
    renderRouteCard();
  }

  // Pick the route matching the package being scanned, so the common case
  // needs no thought.
  function autoSelectRoute() {
    const pkg = (el.dsPackage.value || '').trim();
    if (!pkg) return;
    const hit = routeList.find((rt) => Object.values(rt.packages || {}).includes(pkg));
    if (hit) el.dsRoute.value = hit.name;
  }

  function renderRouteCard() {
    const rt = routeList.find((r) => r.name === el.dsRoute.value);
    if (!rt) {
      el.dsRouteCard.classList.add('hidden');
      el.dsRouteCard.innerHTML = '';
      return;
    }
    const pkg = (el.dsPackage.value || '').trim();
    const envs = Object.entries(rt.packages || {});
    const matched = envs.find(([, v]) => v === pkg);
    const mismatch = pkg && !matched;

    const bits = [];
    bits.push(`<div class="ds-route-line"><strong>${escapeHtml(rt.label)}</strong>
      <span class="ds-route-meta">${rt.screens} screens · ${rt.infoBadges} info badges · ${rt.procedures.length} procedures</span></div>`);
    if (rt.screenNames.length) {
      bits.push(`<div class="ds-route-line ds-route-dim">Knows: ${rt.screenNames.map(escapeHtml).join(', ')}</div>`);
    }
    if (envs.length) {
      bits.push(`<div class="ds-route-line ds-route-dim">Builds: ${envs.map(([k, v]) =>
        `${escapeHtml(k)} <code>${escapeHtml(v)}</code>`).join(' · ')}</div>`);
    }
    if (rt.recordedOn) {
      const d = rt.recordedOn;
      bits.push(`<div class="ds-route-line ds-route-dim">Recorded on ${escapeHtml(d.device || 'unknown device')}${
        d.resolution ? ` at ${d.resolution.join('×')}` : ''}${d.build ? `, build ${escapeHtml(d.build)}` : ''}</div>`);
    }
    rt.knownIssues.forEach((ki) => {
      bits.push(`<div class="ds-route-line ds-route-warn">Known issue — ${escapeHtml(ki.key)}: ${escapeHtml(ki.note)}</div>`);
    });
    if (mismatch) {
      bits.push(`<div class="ds-route-line ds-route-warn">This route was recorded against a different package than
        <code>${escapeHtml(pkg)}</code>; its coordinates may not line up.</div>`);
    }
    el.dsRouteCard.innerHTML = bits.join('');
    el.dsRouteCard.classList.remove('hidden');
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
    renderLangQueue();
    updateStartState();
  }

  /**
   * Start has five independent preconditions. Disabling the button on its own
   * left you guessing which one you had missed, so the same check now produces
   * the list of what is still outstanding and shows it above the button.
   */
  function missingForStart() {
    const missing = [];
    if (!agentOk) missing.push('The scanning service to be running — it starts with the app; try “Test connection”.');
    if (!(agentConfig && agentConfig.apiKeySet)) missing.push('An Anthropic API key — add one in step 1 above.');
    if (!sheet()) missing.push('A sheet to compare against — open one from the sidebar, or pick a different sheet above.');
    else if (!el.dsLanguage.value) missing.push('A language column to check — pick one under “Language column”.');
    if (el.dsMode.value === 'device' && !el.dsDevice.value) {
      missing.push('An Android device connected over USB with USB debugging turned on.');
    }
    return missing;
  }

  function updateStartState() {
    const missing = missingForStart();
    el.dsStart.disabled = running || missing.length > 0;

    if (!el.dsBlockers) return;
    // While a scan is running the button is disabled for an obvious reason,
    // so the checklist would only be noise.
    if (running || !missing.length) {
      el.dsBlockers.classList.add('hidden');
      el.dsStart.removeAttribute('title');
      return;
    }
    el.dsBlockersList.innerHTML = '';
    missing.forEach((text) => {
      const li = document.createElement('li');
      li.textContent = text;
      el.dsBlockersList.appendChild(li);
    });
    el.dsBlockers.classList.remove('hidden');
    el.dsStart.title = `Not ready yet — ${missing.length} thing${missing.length > 1 ? 's' : ''} still needed.`;
  }

  function lines(node) {
    return node.value.split('\n').map((s) => s.trim()).filter(Boolean);
  }

  function options() {
    return {
      maxScreens: Number(el.dsMaxScreens.value),
      maxActions: Number(el.dsMaxActions.value),
      maxDepth: Number(el.dsMaxDepth.value),
      settleMs: Number(el.dsSettle.value),
      scrollSteps: Number(el.dsScrollSteps.value),
      longPressMs: Number(el.dsLongpressMs.value),
      maxMinutes: Number(el.dsMaxMinutes.value),
      stopAfterHighIssues: Number(el.dsStopHigh.value),
      model: el.dsModel.value,
      effort: el.dsEffort.value,
      visionEnabled: el.dsVision.checked,
      scrollProbe: el.dsScroll.checked,
      longPressProbe: el.dsLongpress.checked,
      routeSetLanguage: el.dsRouteLang.checked,
      androidPackage: el.dsPackage.value.trim(),
      blockedLabels: lines(el.dsBlocked),
      onlyLabels: lines(el.dsOnly),
      focusLabels: lines(el.dsFocus),
      customRules: el.dsRules.value,
      preSteps: el.dsSteps.value,
    };
  }

  // ── language queue ──────────────────────────────────────────────────────
  // One scan covers one column. Testers have five languages to get through, so
  // the extra ones queue up and run themselves rather than needing someone at
  // the keyboard between them.

  let queued = [];               // language headers still to scan
  let batchPrimary = '';         // the language the tester actually picked
  let batchRan = [];             // the queue as it stood when Start was pressed
  const results = new Map();     // language -> { issues, screens, runId }

  function renderLangQueue() {
    const s = sheet();
    if (!el.dsLangQueue) return;
    const chosen = new Set(queuedFromUi());
    el.dsLangQueue.innerHTML = '';
    if (!s) {
      el.dsQueueWrap.classList.add('hidden');
      return;
    }
    const current = el.dsLanguage.value;
    const others = s.headers.slice(1).filter((h) => h && h.trim() && h !== current);
    if (!others.length) {
      el.dsQueueWrap.classList.add('hidden');
      return;
    }
    el.dsQueueWrap.classList.remove('hidden');
    others.forEach((h) => {
      const id = `ds-q-${h.replace(/\W+/g, '_')}`;
      const label = document.createElement('label');
      label.className = 'ds-chip';
      label.innerHTML = `<input type="checkbox" id="${id}" value="${escapeHtml(h)}"${chosen.has(h) ? ' checked' : ''} /> ${escapeHtml(h)}`;
      el.dsLangQueue.appendChild(label);
    });
  }

  function queuedFromUi() {
    return [...el.dsLangQueue.querySelectorAll('input:checked')].map((i) => i.value);
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
        const rt = routeList.find((x) => x.name === el.dsRoute.value);
        if (rt && rt.capabilities && rt.capabilities.bridge === false) {
          parts.push('<div class="ds-msg">Reading text from screenshots, as recorded for ' +
            `${escapeHtml(rt.label)}. Truncation and overflow are judged visually.</div>`);
        } else {
          parts.push('<div class="ds-msg warn">No in-game bridge. The scan will read text from screenshots instead — ' +
            'add <code>agent/unity/LocaLinterBridge.cs</code> to the Unity project for exact strings, measured overflow and reliable clicking.</div>');
        }
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
    queued = queuedFromUi();
    batchRan = queued.slice();
    batchPrimary = el.dsLanguage.value;
    results.clear();
    rememberSettings();
    if (queued.length) {
      log(`${queued.length} more language${queued.length === 1 ? '' : 's'} queued after this one: ${queued.join(', ')}.`, 'info');
    }
    return beginRun(el.dsLanguage.value);
  }

  async function beginRun(language) {
    const s = sheet();
    if (!s) return toast('Load a localization sheet first.', 'error');

    issues = [];
    screens = new Map();
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
        body: JSON.stringify({ adbPath: el.dsAdbPath.value.trim(), baseUrl: el.dsBaseUrl.value.trim(), extraChecks: el.dsExtraChecks.value.trim(), ...options() }),
      });
      const r = await api('/api/run/start', {
        method: 'POST',
        body: JSON.stringify({
          mode: el.dsMode.value,
          serial: el.dsDevice.value || null,
          targetLanguage: language,
          route: el.dsRoute.value || null,
          sheet: s,
          options: options(),
        }),
      });
      runId = r.runId;
      currentLanguage = r.language || language;
      running = true;
      reconnectAttempts = 0;
      el.dsStop.disabled = false;
      el.dsStop.textContent = 'Stop scan';
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
    // Stopping means stopping, not "stop this one and start French".
    if (queued.length) {
      log(`Cancelled the ${queued.length} queued language scan${queued.length === 1 ? '' : 's'}.`, 'warn');
      queued = [];
    }
    // The scan finishes the screen it is on before it unwinds, which with the
    // vision pass can take the better part of a minute. Say that, rather than
    // leaving a dead-looking button.
    el.dsStop.disabled = true;
    el.dsStop.textContent = 'Stopping…';
    setStatus(el.dsRunStatus, 'Stopping — finishing the current screen', 'pending');
    try {
      await api(`/api/run/${runId}/stop`, { method: 'POST' });
      log('Stop requested — finishing the current screen, then wrapping up.', 'warn');
    } catch (e) {
      toast(e.message, 'error');
      el.dsStop.disabled = false;
      el.dsStop.textContent = 'Stop scan';
    }
  }

  function listen() {
    if (source) source.close();
    source = new EventSource(`${agentUrl()}/api/run/${runId}/events`);
    source.onmessage = (e) => {
      reconnectAttempts = 0;
      let ev;
      try { ev = JSON.parse(e.data); } catch { return; }
      handleEvent(ev);
    };
    // A dropped stream is not a dead scan. The run lives in the agent, so
    // reconnect and keep following it — an agent restart, a laptop sleeping,
    // or a flaky loopback should not throw away a scan that is still going.
    source.onerror = () => {
      if (!running) return;
      if (reconnectTimer) return;
      reconnectAttempts++;
      if (reconnectAttempts > 12) {
        setStatus(el.dsRunStatus, 'Lost the agent — findings so far are kept below', 'bad');
        log('Gave up reconnecting to the agent. The findings already received are still listed.', 'error');
        running = false;
        el.dsStop.disabled = true;
        updateStartState();
        return;
      }
      setStatus(el.dsRunStatus, `Reconnecting to the agent (${reconnectAttempts})…`, 'pending');
      reconnectTimer = setTimeout(async () => {
        reconnectTimer = null;
        try {
          // If the agent came back without the run, say so rather than
          // reconnecting forever to something that no longer exists.
          const { runs } = await api('/api/runs');
          const mine = (runs || []).find((r) => r.id === runId);
          if (mine && mine.status !== 'running') {
            running = false;
            setStatus(el.dsRunStatus, `Run ${mine.status} while disconnected — ${mine.issues} issues`, mine.status === 'done' ? 'ok' : 'bad');
            el.dsStop.disabled = true;
            updateStartState();
            return;
          }
          if (!mine) {
            running = false;
            setStatus(el.dsRunStatus, 'The agent restarted and lost this run', 'bad');
            log('The agent restarted while the scan was running, so the run was lost. Findings received before that are listed below.', 'error');
            el.dsStop.disabled = true;
            updateStartState();
            return;
          }
        } catch { /* agent still down; the retry below keeps trying */ }
        listen();
      }, Math.min(2000 * reconnectAttempts, 10000));
    };
  }

  /**
   * The issues tile was marked danger in the markup, so it sat red through an
   * entire clean scan. Colour it from the count instead: neutral until the
   * scan has run, green at zero, red once something is actually wrong.
   */
  function setIssueStat(count) {
    el.dsStatIssues.textContent = `Issues: ${count}`;
    const tile = el.dsStatIssues.closest('.stat-pill');
    if (tile) tile.className = `stat-pill ${count > 0 ? 'danger' : 'success'}`;
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
        setIssueStat(ev.issues);
        el.dsStatQueue.textContent = `Queued: ${ev.queued}`;
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
        (window.locaLinterSetBadge || (() => {}))(el.badgeDevice, issues.length);
        setIssueStat(issues.length);
        renderFindings();
        break;
      case 'done':
        running = false;
        el.dsStop.disabled = true;
        el.dsStop.textContent = 'Stop scan';
        updateStartState();
        if (source) source.close();
        setStatus(
          el.dsRunStatus,
          ev.status === 'done' ? `Finished — ${ev.issues} issues` : `${ev.status}${ev.error ? `: ${ev.error}` : ''}`,
          ev.status === 'done' ? 'ok' : ev.status === 'stopped' ? 'warn' : 'bad'
        );
        finalize(ev.status);
        break;
    }
  }

  async function finalize(status) {
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
        log(`Claude usage: ${u.calls} calls, ${u.input} in / ${u.output} out tokens (${u.cacheRead} cached).`, 'info');
      }
      (window.locaLinterSetBadge || (() => {}))(el.badgeDevice, issues.length);
      renderFindings();
    } catch (e) {
      log(`Could not load the final report: ${e.message}`, 'warn');
    }
    // Keep each language's findings so the export covers the whole batch, not
    // just whichever scan happened to finish last.
    results.set(currentLanguage, { runId, issues: issues.slice(), screens: [...screens.values()] });
    runNextQueued(status);
  }

  /**
   * A batch walks the language picker through every queued column, which would
   * otherwise leave it parked on whichever one happened to be last. Put it back
   * where the tester left it, so pressing Start again repeats the same batch.
   */
  function endBatch() {
    if (!batchPrimary || el.dsLanguage.value === batchPrimary) return;
    if (![...el.dsLanguage.options].some((o) => o.value === batchPrimary)) return;
    el.dsLanguage.value = batchPrimary;
    renderLangQueue();
    // The language that was primary a moment ago had no chip to stay ticked
    // in, so the batch is re-ticked from what it actually ran.
    el.dsLangQueue.querySelectorAll('input').forEach((i) => {
      if (batchRan.includes(i.value)) i.checked = true;
    });
    updateStartState();
  }

  function runNextQueued(status) {
    if (!queued.length) return endBatch();
    if (status !== 'done') {
      log(`Run ${status} — not starting the ${queued.length} queued language scan${queued.length === 1 ? '' : 's'}.`, 'warn');
      queued = [];
      endBatch();
      return;
    }
    const next = queued.shift();
    const done = [...results.keys()].join(', ');
    log(`Finished ${done}. Starting ${next}${queued.length ? ` (${queued.length} still queued)` : ''}.`, 'info');
    el.dsLanguage.value = next;
    setTimeout(() => beginRun(next), 500);
  }

  // ── rendering ───────────────────────────────────────────────────────────

  const SEVERITY_RANK = { high: 3, medium: 2, low: 1 };

  /**
   * Findings are shown as a screen list beside one screen's detail, not as one
   * long scroll.
   *
   * Stacking every screen with its shot and its issues meant a 120-screen run
   * produced a page metres long, where finding the screen you cared about was
   * the hard part. A rail of thumbnails answers "which screens are bad" at a
   * glance, and the pane beside it answers "what is wrong here" — both scroll
   * independently inside a panel that never grows.
   */
  let selectedScreenId = null;

  function visibleIssues() {
    const minSev = el.dsFilterSeverity.value === 'high' ? 3 : el.dsFilterSeverity.value === 'medium' ? 2 : 1;
    const typeFilter = el.dsFilterType.value;
    const q = el.dsFilterText.value.trim().toLowerCase();

    return issues.filter((i) => {
      if ((SEVERITY_RANK[i.severity] || 1) < minSev) return false;
      if (typeFilter !== 'all' && i.type !== typeFilter) return false;
      if (q) {
        const hay = `${i.text} ${i.message} ${i.element || ''} ${i.key || ''} ${i.screenId}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    });
  }

  function syncTypeFilter() {
    const types = [...new Set(issues.map((i) => i.type))].sort();
    const current = el.dsFilterType.value;
    if (el.dsFilterType.options.length === types.length + 1) return;
    el.dsFilterType.innerHTML = '<option value="all">All types</option>';
    for (const t of types) {
      const o = document.createElement('option');
      o.value = t;
      o.textContent = t.replace(/_/g, ' ');
      el.dsFilterType.appendChild(o);
    }
    if (types.includes(current)) el.dsFilterType.value = current;
  }

  function shotUrl(screen) {
    return screen ? `${agentUrl()}/api/run/${runId}/screenshot/${screen.file}` : '';
  }

  function renderFindings() {
    syncTypeFilter();
    const shown = visibleIssues();

    if (!shown.length) {
      el.dsFindings.innerHTML = issues.length
        ? '<div class="ds-empty">No findings match this filter.</div>'
        : '<div class="ds-empty">No issues yet.</div>';
      return;
    }

    // Every captured screen appears, so a clean screen is visible as clean
    // rather than simply absent — but only screens with a match when filtering.
    const counts = new Map();
    for (const i of shown) counts.set(i.screenId, (counts.get(i.screenId) || 0) + 1);
    const worst = new Map();
    for (const i of shown) {
      const rank = SEVERITY_RANK[i.severity] || 1;
      if (rank > (worst.get(i.screenId) || 0)) worst.set(i.screenId, rank);
    }

    const filtering = el.dsFilterSeverity.value !== 'all' || el.dsFilterType.value !== 'all' || !!el.dsFilterText.value.trim();
    const ids = [...screens.keys()]
      .sort((a, b) => a.localeCompare(b))
      .filter((id) => !filtering || counts.has(id));

    if (!ids.includes(selectedScreenId)) {
      // Land on the worst screen rather than the first — that is the one you
      // opened the panel to look at.
      selectedScreenId = ids.slice().sort((a, b) =>
        (worst.get(b) || 0) - (worst.get(a) || 0) || (counts.get(b) || 0) - (counts.get(a) || 0)
      )[0] || ids[0] || null;
    }

    const sevWord = (rank) => (rank >= 3 ? 'high' : rank === 2 ? 'medium' : 'low');

    const rail = ids.map((id) => {
      const screen = screens.get(id);
      const n = counts.get(id) || 0;
      const shot = shotUrl(screen);
      return `
        <button class="ds-rail-item${id === selectedScreenId ? ' active' : ''}" data-screen="${escapeHtml(id)}" type="button">
          ${shot ? `<img src="${shot}" alt="" loading="lazy" />` : '<span class="ds-rail-noshot"></span>'}
          <span class="ds-rail-meta">
            <span class="ds-rail-id">${escapeHtml(id.replace(/^screen-/, ''))}</span>
            <span class="ds-rail-count ${n ? 'sev-' + sevWord(worst.get(id) || 1) : 'clean'}">${n || '✓'}</span>
          </span>
        </button>`;
    }).join('');

    const screen = screens.get(selectedScreenId);
    const list = shown
      .filter((i) => i.screenId === selectedScreenId)
      .sort((a, b) => (SEVERITY_RANK[b.severity] || 0) - (SEVERITY_RANK[a.severity] || 0));
    const shot = shotUrl(screen);

    const detail = `
      <div class="ds-detail-head">
        <strong>${escapeHtml(selectedScreenId || '')}</strong>
        <span class="ds-screen-meta">${escapeHtml((screen && screen.summary) || '')}</span>
        <span class="ds-detail-count">${list.length} finding${list.length === 1 ? '' : 's'}</span>
      </div>
      ${screen && screen.path && screen.path.length
        ? `<div class="ds-path">Reached via: ${screen.path.map((p) => `<span>${escapeHtml(truncate(p, 28))}</span>`).join(' › ')}</div>`
        : ''}
      <div class="ds-detail-body">
        <div class="ds-detail-shot">
          ${shot ? `<img src="${shot}" alt="${escapeHtml(selectedScreenId || '')}" data-shot="${shot}" data-caption="${escapeHtml((screen && screen.summary) || selectedScreenId || '')}" />` : ''}
        </div>
        <div class="ds-issue-list">
          ${list.length ? list.map(issueRow).join('') : '<div class="ds-empty">Nothing wrong on this screen.</div>'}
        </div>
      </div>`;

    el.dsFindings.innerHTML =
      `<div class="ds-results">
         <div class="ds-rail" role="listbox" aria-label="Captured screens">${rail}</div>
         <div class="ds-detail">${detail}</div>
       </div>`;

    el.dsFindings.querySelectorAll('.ds-rail-item').forEach((btn) => {
      btn.addEventListener('click', () => {
        selectedScreenId = btn.dataset.screen;
        renderFindings();
      });
    });
    const active = el.dsFindings.querySelector('.ds-rail-item.active');
    if (active) active.scrollIntoView({ block: 'nearest' });

    el.dsFindings.querySelectorAll('.ds-dismiss').forEach((btn) => {
      btn.addEventListener('click', () => dismissFinding(btn));
    });

    el.dsFindings.querySelectorAll('img[data-shot]').forEach((img) => {
      img.addEventListener('click', () => {
        el.dsShotImg.src = img.dataset.shot;
        el.dsShotCaption.textContent = img.dataset.caption;
        el.dsShotOverlay.classList.remove('hidden');
      });
    });
  }

  /** Wipes stored runs, and the current view with them. */
  async function clearHistory() {
    if (running) return toast('Stop the scan before clearing.', 'error');
    try {
      const r = await api('/api/runs/clear', { method: 'POST' });
      issues = [];
      screens = new Map();
      selectedScreenId = null;
      runId = null;
      el.dsFindings.innerHTML = '';
      el.dsLog.innerHTML = '';
      el.dsSummary.classList.add('hidden');
      el.dsResultsPanel.classList.add('hidden');
      el.dsRunPanel.classList.add('hidden');
      (window.locaLinterSetBadge || (() => {}))(el.badgeDevice, 0);
      const mb = (r.bytes || 0) / (1024 * 1024);
      toast(r.removed ? `Cleared ${r.removed} run${r.removed === 1 ? '' : 's'} (${mb.toFixed(1)} MB).` : 'Nothing to clear.');
    } catch (e) {
      toast(e.message, 'error');
    }
  }

  /**
   * "Not an issue" is the only signal a human gives the agent, so it is stored
   * against the app rather than the run: it suppresses the finding on every
   * later scan, and is quoted back to the model as context.
   */
  async function dismissFinding(btn) {
    const issue = { type: btn.dataset.type, text: btn.dataset.text, key: btn.dataset.key };
    btn.disabled = true;
    btn.textContent = 'Remembering…';
    try {
      await api('/api/memory/dismiss', { method: 'POST', body: JSON.stringify({ issue }) });
      issues = issues.filter((i) => !(i.type === issue.type && (i.text || '') === issue.text));
      (window.locaLinterSetBadge || (() => {}))(el.badgeDevice, issues.length);
      renderFindings();
      toast('Noted — I will not report that again.');
    } catch (e) {
      btn.disabled = false;
      btn.textContent = 'Not an issue';
      toast(e.message, 'error');
    }
  }

  /**
   * Answers "why did it flag this" and, more usefully, "why did it not" —
   * against the sheet that is actually loaded, not a description of the rules.
   */
  async function explainString() {
    const text = el.dsExplain.value.trim();
    const s = sheet();
    if (!text) return;
    if (!s) return toast('Load a sheet first.', 'error');

    el.dsExplainOut.classList.remove('hidden');
    el.dsExplainOut.innerHTML = '<div class="ds-hint">Checking…</div>';
    try {
      const r = await api('/api/sheet/explain', {
        method: 'POST',
        body: JSON.stringify({ sheet: s, text, targetLanguage: el.dsLanguage.value }),
      });
      const rows = (r.rows || []).map((m) => `
        <tr><td>${escapeHtml(m.key)}</td><td>${m.row}</td><td>${escapeHtml(m.matchedColumn)}</td>
        <td>${escapeHtml(m.source || '')}</td><td>${escapeHtml(m.target || '')}</td></tr>`).join('');
      const near = (r.near || []).map((m) => `
        <tr><td>${escapeHtml(m.key)}</td><td>${m.row}</td><td>${escapeHtml(m.matchedColumn)} · ${m.score}%</td>
        <td colspan="2">${escapeHtml(m.value || '')}</td></tr>`).join('');
      el.dsExplainOut.innerHTML = `
        <div class="ds-verdict ${escapeHtml(r.verdict)}">${escapeHtml(r.verdict)}</div>
        <p class="ds-hint">${escapeHtml(r.reason)}</p>
        ${rows || near ? `<table class="ds-explain-table">
          <thead><tr><th>Key</th><th>Row</th><th>Matched</th><th>${escapeHtml(r.targetHeader ? 'Source' : '')}</th><th>${escapeHtml(r.targetHeader || '')}</th></tr></thead>
          <tbody>${rows}${near}</tbody></table>` : ''}`;
    } catch (e) {
      el.dsExplainOut.innerHTML = `<div class="ds-msg bad">${escapeHtml(e.message)}</div>`;
    }
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
        <div class="ds-issue-foot">
          ${bits.length ? `<span class="ds-issue-meta">${bits.join(' · ')}</span>` : '<span></span>'}
          <button class="ds-dismiss" type="button" title="Remember that this is not a defect, and stop reporting it"
            data-type="${escapeHtml(i.type)}" data-text="${escapeHtml(i.text || '')}" data-key="${escapeHtml(i.key || '')}">Not an issue</button>
        </div>
      </div>`;
  }

  function exportJson() {
    const payload = {
      runId,
      exportedAt: new Date().toISOString(),
      language: currentLanguage || el.dsLanguage.value,
      settings: options(),
      screens: [...screens.values()],
      issues,
      // A batch of languages is one job to whoever reads the report.
      batch: [...results.entries()].map(([language, r]) => ({
        language, runId: r.runId, issues: r.issues, screens: r.screens,
      })),
    };
    download(`device-scan-${runId || 'report'}.json`, JSON.stringify(payload, null, 2), 'application/json');
  }

  // Spreadsheets are where triage actually happens: one row per finding, with
  // the sheet key and row number so a fix goes straight back into the source.
  function exportCsv() {
    const cols = ['language', 'severity', 'type', 'source', 'screen', 'text', 'expected', 'key', 'sheetRow', 'element', 'message', 'reachedVia'];
    const rows = [];
    const batch = results.size ? [...results.entries()] : [[currentLanguage || el.dsLanguage.value, { issues, screens: [...screens.values()] }]];
    for (const [language, r] of batch) {
      const byId = new Map((r.screens || []).map((s) => [s.id, s]));
      for (const i of r.issues || []) {
        const screen = byId.get(i.screenId);
        rows.push([
          language, i.severity, i.type, i.alsoSeenBy ? `${i.source}+vision` : i.source, i.screenId,
          i.text, i.expected || '', i.key || '', i.row || '', i.element || i.where || '', i.message,
          screen && screen.path ? screen.path.join(' > ') : '',
        ]);
      }
    }
    const esc = (v) => {
      const s = String(v == null ? '' : v);
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const csv = [cols.join(','), ...rows.map((r) => r.map(esc).join(','))].join('\r\n');
    // A leading BOM is what makes Excel open UTF-8 without mangling the very
    // translations this report is about.
    download(`device-scan-${runId || 'report'}.csv`, `﻿${csv}`, 'text/csv');
  }

  function download(name, content, type) {
    const blob = new Blob([content], { type: `${type};charset=utf-8` });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
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
