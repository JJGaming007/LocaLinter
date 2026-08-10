/**
 * App shell.
 *
 * main.js owns the sheet, the validation and the tables. This file owns the
 * frame around them: the sidebar, the breadcrumb, the page heading, and the
 * AI budget meter. It talks to main.js only through the DOM and the
 * `localinter:tab` event that switchTab() already dispatches, so nothing here
 * needs main.js to know it exists.
 */
(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);

  const VIEWS = {
    home: {
      crumb: 'Home',
      title: 'Home',
      sub: 'Drop an .xlsx or .csv sheet to check formatting, brackets and variables.',
    },
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

  const VIEW_KEY = 'locaLinterView';
  const COLLAPSE_KEY = 'locaLinterSidebarCollapsed';
  const BUDGET_KEY = 'localinter_ai_budget';

  let restoringHome = false;

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    wireSidebar();
    wirePresetSelect();
    wireSheetMirror();
    wireBudget();
    wireShortcuts();

    document.addEventListener('localinter:tab', (e) => {
      if (restoringHome) {
        restoringHome = false;
        showHome();
        return;
      }
      setView(e.detail.tabId);
    });

    // main.js restores the last tab asynchronously (after IndexedDB answers),
    // so a stored "home" has to win once that lands.
    const storedView = localStorage.getItem(VIEW_KEY);
    const storedTab = localStorage.getItem('locaLinterActiveTab');
    if (!storedView || storedView === 'home') {
      restoringHome = true;
      showHome();
      // If no sheet is stored, no tab event ever fires — home is already right.
      setTimeout(() => { restoringHome = false; }, 3000);
    } else {
      setView(storedTab && VIEWS[storedTab] ? storedTab : 'format-issues');
    }
  }

  /* ── views ───────────────────────────────────────────────── */

  function setView(tabId) {
    const view = VIEWS[tabId] || VIEWS.home;
    document.body.classList.remove('view-home');
    $('nav-home').classList.remove('active');
    paint(view);
    localStorage.setItem(VIEW_KEY, tabId);
  }

  function showHome() {
    document.body.classList.add('view-home');
    document.querySelectorAll('.tab-btn').forEach((b) => b.classList.remove('active'));
    $('nav-home').classList.add('active');
    paint(VIEWS.home);
    localStorage.setItem(VIEW_KEY, 'home');
  }

  function paint(view) {
    $('crumb-tab').textContent = view.crumb;
    $('page-title').innerHTML = `${escapeHtml(view.title)}<span class="accent-text">.</span>`;
    $('page-sub').textContent = view.sub;
  }

  /* ── sidebar ─────────────────────────────────────────────── */

  function wireSidebar() {
    $('nav-home').addEventListener('click', showHome);

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

  /* ── AI budget ───────────────────────────────────────────── */

  const EMPTY_SPEND = {
    input: 0, output: 0, cacheRead: 0, cacheWrite: 0, calls: 0, runs: 0, costUSD: 0, unpriced: false,
  };

  function readBudget() {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(BUDGET_KEY)) || {}; } catch { /* corrupt, start fresh */ }
    const month = new Date().toISOString().slice(0, 7);
    // v1 counted tokens against a made-up cap; those numbers mean nothing here.
    if (saved.v !== 2) saved = { v: 2, month };
    if (saved.month !== month) saved = { v: 2, month, cap: saved.cap ?? null };
    return { v: 2, month, cap: null, ...EMPTY_SPEND, ...saved };
  }

  function writeBudget(b) {
    localStorage.setItem(BUDGET_KEY, JSON.stringify(b));
    renderBudget(b);
  }

  /**
   * Everything shown here is measured: the tokens come from the usage counts the
   * Claude API returns for each scan, and the dollar figure is those counts at
   * the published per-model rates. With nothing scanned there is nothing to
   * report, so we say that instead of drawing a meter against an invented limit.
   */
  function renderBudget(b) {
    const tokens = (b.input || 0) + (b.output || 0) + (b.cacheRead || 0) + (b.cacheWrite || 0);
    const hasData = (b.calls || 0) > 0 || tokens > 0;
    const meterWrap = $('ai-meter-wrap');
    const capLabel = b.cap > 0 ? `of $${trimNum(b.cap)}` : 'set cap';

    if (!hasData) {
      $('ai-spend').textContent = '—';
      $('ai-detail').textContent = 'No scans this month';
      meterWrap.classList.add('hidden');
      $('ai-edit').textContent = b.cap > 0 ? capLabel : 'set cap';
      return;
    }

    // A model with no published rate gives tokens only — never a made-up price.
    const spend = b.unpriced ? null : b.costUSD || 0;
    $('ai-spend').textContent = spend == null
      ? `${compact(tokens)} tok`
      : `$${spend.toFixed(spend < 1 ? 3 : 2)}`;

    const runs = `${b.runs || 0} scan${b.runs === 1 ? '' : 's'}`;
    const detail = $('ai-detail');
    detail.textContent = spend == null
      ? `${runs} · ${b.calls} calls · no published rate`
      : `${compact(tokens)} tokens · ${runs}`;
    detail.title = `${b.calls} API calls this month · ${resetLabel()}`;

    if (b.cap > 0 && spend != null) {
      const pct = (spend / b.cap) * 100;
      meterWrap.classList.remove('hidden');
      const fill = $('ai-meter');
      fill.style.width = `${Math.min(100, pct)}%`;
      fill.classList.toggle('warn', pct >= 75 && pct < 100);
      fill.classList.toggle('over', pct >= 100);
    } else {
      meterWrap.classList.add('hidden');
    }
    $('ai-edit').textContent = capLabel;
  }

  function trimNum(n) {
    return String(parseFloat(Number(n).toFixed(2)));
  }

  function resetLabel() {
    const now = new Date();
    const next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    return `resets ${next.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })}`;
  }

  function compact(n) {
    const round = (v, d) => String(parseFloat(v.toFixed(d)));
    if (n >= 1e9) return `${round(n / 1e9, 2)}B`;
    if (n >= 1e6) return `${round(n / 1e6, 2)}M`;
    if (n >= 1e3) return `${round(n / 1e3, 1)}K`;
    return String(n);
  }

  function wireBudget() {
    renderBudget(readBudget());

    // device-scan.js forwards each scan's usage as deltas, live during the run.
    document.addEventListener('localinter:usage', (e) => {
      const u = e.detail || {};
      const b = readBudget();
      b.input += u.input || 0;
      b.output += u.output || 0;
      b.cacheRead += u.cacheRead || 0;
      b.cacheWrite += u.cacheWrite || 0;
      b.calls += u.calls || 0;
      b.costUSD += u.costUSD || 0;
      if (u.priced === false) b.unpriced = true;
      if (u.runFinished) b.runs += 1;
      writeBudget(b);
    });

    // Mirror the agent's connection state into the sidebar.
    const agentStatus = $('ds-agent-status');
    if (agentStatus) {
      const sync = () => {
        const on = agentStatus.classList.contains('ok');
        const busy = agentStatus.classList.contains('pending');
        const dot = $('ai-agent');
        dot.classList.toggle('on', on);
        dot.classList.toggle('busy', busy);
        dot.textContent = busy ? 'agent connecting' : on ? 'agent online' : 'agent offline';
      };
      new MutationObserver(sync).observe(agentStatus, {
        attributes: true,
        attributeFilter: ['class'],
        childList: true,
        characterData: true,
        subtree: true,
      });
      sync();
    }

    // Settings: the optional cap lives next to the ignore list.
    const capInput = $('ai-cap-input');
    const showCap = (cap) => { capInput.value = cap > 0 ? trimNum(cap) : ''; };
    showCap(readBudget().cap);

    $('ai-edit').addEventListener('click', () => {
      $('global-settings-btn').click();
      capInput.focus();
      capInput.select();
    });

    $('save-ignore-btn').addEventListener('click', () => {
      const parsed = parseFloat(String(capInput.value).replace(/[^\d.]/g, ''));
      const b = readBudget();
      b.cap = parsed > 0 ? parsed : null;   // empty or nonsense means "no cap"
      showCap(b.cap);
      writeBudget(b);
    });

    $('ai-reset-btn').addEventListener('click', () => {
      const b = { ...readBudget(), ...EMPTY_SPEND };
      writeBudget(b);
      showCap(b.cap);
    });
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
