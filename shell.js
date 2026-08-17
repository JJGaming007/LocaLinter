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

  const COLLAPSE_KEY = 'locaLinterSidebarCollapsed';

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    wireSidebar();
    wirePresetSelect();
    wireSheetMirror();
    wireSheetPicker();
    wireAgentStatus();
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

  function wireSidebar() {
    if (localStorage.getItem(COLLAPSE_KEY) === '1') {
      document.body.classList.add('sb-collapsed');
    }

    $('sidebar-toggle').addEventListener('click', () => {
      if (window.matchMedia('(max-width: 860px)').matches) {
        document.body.classList.remove('sb-open');
        return;
      }
      const collapsed = document.body.classList.toggle('sb-collapsed');
      localStorage.setItem(COLLAPSE_KEY, collapsed ? '1' : '0');
    });

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

  /**
   * With a sheet loaded the source panel collapses to a one-line bar, which
   * used to leave the full picker reachable only from the Home tab. This
   * re-opens it in place, so opening a different sheet no longer means
   * navigating away from your results.
   */
  function wireSheetPicker() {
    const buttons = [$('change-sheet-btn'), $('toolbar-change-btn')].filter(Boolean);
    const picker = $('drop-content-idle');
    if (!buttons.length || !picker) return;

    const setOpen = (open) => {
      document.body.classList.toggle('picker-open', open);
      buttons.forEach((b) => {
        b.classList.toggle('is-open', open);
        b.setAttribute('aria-expanded', String(open));
      });
    };

    buttons.forEach((b) => b.addEventListener('click', () => {
      setOpen(!document.body.classList.contains('picker-open'));
    }));

    // Loading a sheet is the point of the picker, so it closes itself.
    document.addEventListener('localinter:sheet', () => setOpen(false));
    document.querySelectorAll('.quick-open-btn').forEach((b) => {
      b.addEventListener('click', () => setOpen(false));
    });
  }

  function wirePresetSelect() {
    const select = $('sb-preset');
    select.addEventListener('change', () => {
      const preset = select.value;
      if (!preset) return;
      const btn = document.querySelector(`.quick-open-btn[data-preset="${preset}"]`);
      if (btn) btn.click();
    });

    document.querySelectorAll('.quick-open-btn').forEach((btn) => {
      btn.addEventListener('click', () => { select.value = btn.dataset.preset || ''; });
    });
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
    });
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
    ));
  }
})();
