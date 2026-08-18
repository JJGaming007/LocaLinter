/**
 * App shell.
 *
 * main.js owns the sheet, the validation and the tables. This file owns the
 * frame around them: the sidebar, the breadcrumb and the page heading. It
 * talks to main.js only through the DOM and the `localinter:tab` event that
 * switchTab() already dispatches, so nothing here needs main.js to know it
 * exists.
 */
(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);

  const VIEWS = {
    'format-issues': {
      crumb: 'Format issues',
      title: 'Format issues',
      sub: 'Bracket, spacing and variable problems found in the translated strings.',
    },
    'missing-locales': {
      crumb: 'Missing locales',
      title: 'Missing locales',
      sub: 'Keys that have English copy but no translation in one or more languages.',
    },
    search: {
      crumb: 'Search',
      title: 'Search',
      sub: 'Query every column of the loaded sheet — contains, exact, word or regex.',
    },
    'device-scan': {
      crumb: 'Device scan',
      title: 'Device scan',
      sub: 'Claude drives a build on a device or in the editor and reports what a player would see.',
    },
  };

  const THEME_KEY = 'locaLinterTheme';

  // wireHelp installs the real one.
  let openHelp = () => {};

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    wireSidebar();
    wirePresetSelect();
    wireSheetMirror();
    wireAgentStatus();
    wireTheme();
    wireHelp();
    wireShortcuts();

    document.addEventListener('localinter:tab', (e) => setView(e.detail.tabId));

    const storedTab = localStorage.getItem('locaLinterActiveTab');
    setView(storedTab && VIEWS[storedTab] ? storedTab : 'format-issues');
  }

  /* ── views ───────────────────────────────────────────────── */

  // main.js persists the active tab; this only paints the chrome around it.
  function setView(tabId) {
    paint(VIEWS[tabId] || VIEWS['format-issues']);
  }

  function paint(view) {
    $('crumb-tab').textContent = view.crumb;
    $('page-title').innerHTML = `${escapeHtml(view.title)}<span class="accent-text">.</span>`;
    $('page-sub').textContent = view.sub;
  }

  /* ── sidebar ─────────────────────────────────────────────── */

  /**
   * The sidebar is one fixed width and never collapses — an icon-only rail
   * hid every label and left the brand mark shifting around in the gap. On a
   * narrow window it becomes a drawer instead, which is the only mode where
   * hiding it is worth anything.
   */
  function wireSidebar() {
    $('sidebar-open').addEventListener('click', () => {
      document.body.classList.toggle('sb-open');
    });

    // Tapping a nav item on a phone should close the drawer behind it.
    document.querySelectorAll('.sb-item').forEach((btn) => {
      btn.addEventListener('click', () => document.body.classList.remove('sb-open'));
    });

    document.addEventListener('click', (e) => {
      if (!document.body.classList.contains('sb-open')) return;
      if ($('sidebar').contains(e.target) || $('sidebar-open').contains(e.target)) return;
      document.body.classList.remove('sb-open');
    });
  }

  /* ── sheet source ────────────────────────────────────────── */

  /**
   * The dropdown is the single place a sheet is chosen, so it has to be honest
   * about what is open — including sheets it did not open itself (a restore
   * from the local cache, a pasted URL, a drag-and-drop). It therefore never
   * holds a stale choice: every load re-syncs it from what is actually loaded.
   */
  let sheetLoading = false;

  function wirePresetSelect() {
    const select = $('sb-preset');

    select.addEventListener('change', () => {
      const choice = select.value;

      if (choice === '__local__') {
        // The dialog is cancellable and gives no callback when dismissed, so
        // the select must not be left showing an action that may not happen.
        syncPresetSelect();
        $('file-input').click();
        return;
      }
      if (!choice) return;                       // the placeholder is not a choice
      loadPreset(choice);
    });

    // Every load — preset, URL, file, restore — ends here.
    document.addEventListener('localinter:sheet', syncPresetSelect);
    syncPresetSelect();
  }

  /**
   * One load at a time. Two quick choices used to race, and whichever fetch
   * happened to finish last won — which is not necessarily the one clicked
   * last, so you could end up on a sheet you did not pick.
   */
  function loadPreset(key) {
    if (sheetLoading) return;
    const api = window.LocaLinter;
    if (!api || !api.loadPreset) return;

    const label = labelForKey(key);
    setSheetLoading(true, label);

    // loadFromSheetUrl reports its own failures and resolves either way, so
    // rather than trust the outcome, re-read what is actually loaded.
    Promise.resolve(api.loadPreset(key))
      .catch(() => {})
      .finally(() => {
        setSheetLoading(false);
        syncPresetSelect();
      });
  }

  function labelForKey(key) {
    const opt = $('sb-preset').querySelector(`option[value="${key}"]`);
    return opt ? opt.textContent : key;
  }

  function setSheetLoading(on, label) {
    sheetLoading = on;
    $('sb-preset').disabled = on;
    $('sb-sheet').classList.toggle('is-loading', on);
    if (on) $('sb-sheet-name').textContent = `Loading ${label}…`;
  }

  /** Point the dropdown at whatever is genuinely open. */
  function syncPresetSelect() {
    if (sheetLoading) return;
    const select = $('sb-preset');
    const api = window.LocaLinter;
    const key = api && api.getCurrentPresetKey ? api.getCurrentPresetKey() : null;

    if (key) {
      select.value = key;
      return;
    }
    // Not one of the environments: name it in the hidden placeholder so the
    // control still reads as "this is what is open" rather than going blank.
    const loaded = $('loaded-content');
    const hasSheet = loaded && !loaded.classList.contains('hidden');
    const name = $('loaded-file-name').textContent.trim().replace(/^Loaded:\s*/i, '');
    $('sb-preset-placeholder').textContent = hasSheet && name ? name : 'No sheet loaded';
    select.value = '';
  }


  /* ── sheet identity in the chrome ────────────────────────── */

  function wireSheetMirror() {
    const nameEl = $('loaded-file-name');
    const loaded = $('loaded-content');
    const scanned = $('stat-scanned');

    const sync = () => {
      const hasFile = !loaded.classList.contains('hidden') && nameEl.textContent.trim() !== '';
      document.body.classList.toggle('has-sheet', hasFile);
      // main.js labels the file "Loaded: <name>"; the chrome only wants the name.
      const name = hasFile ? nameEl.textContent.trim().replace(/^Loaded:\s*/i, '') : 'No sheet';
      $('crumb-sheet').textContent = name;
      $('sb-sheet-name').textContent = hasFile ? name : 'No sheet loaded';
      $('sb-sheet-dot').classList.toggle('on', hasFile);

      const rows = scanned ? scanned.textContent.replace(/[^\d,]/g, '') : '';
      $('topbar-meta').textContent = hasFile && rows ? `${rows} rows` : '';

      // The name is already in the breadcrumb and the sidebar, so the status
      // bar carries only what neither of them shows: how big the sheet is.
      $('sbar-sheet-dot').classList.toggle('on', hasFile);
      $('sbar-counts').textContent = hasFile
        ? (rows ? `${rows} rows` : 'Sheet loaded')
        : 'No sheet loaded';
    };

    new MutationObserver(sync).observe(loaded, {
      subtree: true,
      childList: true,
      characterData: true,
      attributes: true,
      attributeFilter: ['class'],
    });
    if (scanned) {
      new MutationObserver(sync).observe(scanned, { childList: true, characterData: true, subtree: true });
    }
    sync();
  }

  /* ── agent presence ──────────────────────────────────────── */

  /**
   * The device-scan tab owns the real connection status; the sidebar just
   * mirrors it so you can tell the agent is up without leaving the tab
   * you are on.
   */
  function wireAgentStatus() {
    const source = $('ds-agent-status');
    const line = $('sb-agent-state');
    const dot = $('tb-agent');
    if (!source) return;

    const sync = () => {
      const on = source.classList.contains('ok');
      const busy = source.classList.contains('pending');
      const label = busy ? 'Service starting…' : on ? 'Scanning service ready' : 'Scanning service stopped';
      for (const el of [line, dot]) {
        if (!el) continue;
        el.classList.toggle('on', on);
        el.classList.toggle('busy', busy);
        el.classList.toggle('bad', !on && !busy);
      }
      if (line) line.textContent = label;
      if (dot) dot.title = label;
    };
    new MutationObserver(sync).observe(source, {
      attributes: true,
      attributeFilter: ['class'],
      childList: true,
      characterData: true,
      subtree: true,
    });
    sync();
  }


  /* ── theme ───────────────────────────────────────────────── */

  /**
   * Light is the default: this is a document tool used in daylight, and the
   * paper ground is what the palette was drawn for. The choice is remembered,
   * and the Windows title-bar strip is repainted to match — it belongs to the
   * main process, so preload.js carries the colours across.
   */
  const TITLEBAR = {
    light: { color: '#ffffff', symbolColor: '#6b6a64' },
    dark: { color: '#161412', symbolColor: '#ada89f' },
  };

  function wireTheme() {
    const btn = $('theme-toggle');
    const label = $('theme-toggle-text');

    const apply = (theme) => {
      document.documentElement.setAttribute('data-theme', theme);
      if (label) label.textContent = theme === 'dark' ? 'Light mode' : 'Dark mode';
      if (btn) btn.title = theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode';
      if (window.localinterShell && window.localinterShell.setTitleBarTheme) {
        window.localinterShell.setTitleBarTheme(TITLEBAR[theme]);
      }
    };

    const stored = localStorage.getItem(THEME_KEY);
    apply(stored === 'dark' || stored === 'light' ? stored : 'light');

    if (!btn) return;
    btn.addEventListener('click', () => {
      const next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
      localStorage.setItem(THEME_KEY, next);
      apply(next);
      closeAccountMenu();
    });
  }

  /** main.js owns the flyout; this only dismisses it after an action. */
  function closeAccountMenu() {
    const menu = $('account-menu');
    const btn = $('account-btn');
    if (menu) menu.classList.add('hidden');
    if (btn) btn.setAttribute('aria-expanded', 'false');
  }

  /* ── help ────────────────────────────────────────────────── */

  /**
   * Shares the settings modal's overlay so only one dimmer ever exists, but
   * owns its own open/close so main.js's Escape handler cannot half-close it.
   */
  function wireHelp() {
    const modal = $('help-modal');
    const overlay = $('modal-overlay');
    if (!modal || !overlay) return;

    openHelp = () => {
      $('settings-modal').classList.add('hidden');
      modal.classList.remove('hidden');
      overlay.classList.remove('hidden');
      $('help-close').focus();
    };
    const close = () => {
      modal.classList.add('hidden');
      // Settings may not have been what opened the overlay; hide it either way.
      if ($('settings-modal').classList.contains('hidden')) overlay.classList.add('hidden');
    };

    $('help-btn').addEventListener('click', openHelp);
    // Settings now lives in the account flyout; it opens the same modal as the
    // toolbar gear, and closes the flyout behind it.
    $('sb-settings').addEventListener('click', () => {
      closeAccountMenu();
      $('global-settings-btn').click();
    });
    $('help-close').addEventListener('click', close);
    overlay.addEventListener('click', close);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
  }

  /* ── shortcuts ───────────────────────────────────────────── */

  function wireShortcuts() {
    document.addEventListener('keydown', (e) => {
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      const t = e.target;
      if (t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable)) return;
      if (e.key === 't' || e.key === 'T') {
        e.preventDefault();
        $('qt-toggle').click();
      }
      if (e.key === '?') {
        e.preventDefault();
        openHelp();
      }
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
    ));
  }
})();
