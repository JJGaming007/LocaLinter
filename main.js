window.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const browseBtn = document.getElementById('browse-btn');
  const resultsContainer = document.getElementById('results-container');
  const resultsHeader = document.querySelector('.results-header');
  const resultsBody = document.getElementById('results-body');
  const missingBody = document.getElementById('missing-body');
  const statScanned = document.getElementById('stat-scanned');
  const statIssues = document.getElementById('stat-issues');
  const statScannedContainer = document.getElementById('stat-scanned-container');
  const statIssuesContainer = document.getElementById('stat-issues-container');
  const badgeFormat = document.getElementById('badge-format');
  const badgeMissing = document.getElementById('badge-missing');
  const tabBtns = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');
  const dropText = document.querySelector('#drop-content-idle p');
  const dropContentIdle = document.getElementById('drop-content-idle');
  const loadedContent = document.getElementById('loaded-content');
  const loadedFileName = document.getElementById('loaded-file-name');
  const removeFileBtn = document.getElementById('remove-file-btn');
  const ignoreListInput = document.getElementById('ignore-list');
  const saveIgnoreBtn = document.getElementById('save-ignore-btn');
  const langFilter = document.getElementById('lang-filter');

  // ── Search tab DOM refs ──
  const searchQueryInput = document.getElementById('search-query');
  const searchClearX = document.getElementById('search-clear-x');
  const searchOptionsBtn = document.getElementById('search-options-btn');
  const searchOptionsPanel = document.getElementById('search-options-panel');
  const searchModesEl = document.getElementById('searchModes');
  const sCaseChk = document.getElementById('s-case');
  const sWrapChk = document.getElementById('s-wrap');
  const searchColChips = document.getElementById('searchColChips');
  const searchSummary = document.getElementById('searchSummary');
  const searchTableWrap = document.getElementById('searchTableWrap');
  const searchThead = document.getElementById('search-thead');
  const searchTbody = document.getElementById('search-tbody');
  const searchPagination = document.getElementById('searchPagination');
  
  const globalWrapChk = document.getElementById('global-wrap');
  const formatTableWrap = document.getElementById('format-table-wrap');
  const missingTableWrap = document.getElementById('missing-table-wrap');
  const formatSummary = document.getElementById('format-summary');
  const missingSummary = document.getElementById('missing-summary');
  const validationInfoGroup = document.getElementById('validation-info-group');
  const formatPagination = document.getElementById('format-pagination');
  const missingPagination = document.getElementById('missing-pagination');
  
  const globalSettingsBtn = document.getElementById('global-settings-btn');
  const settingsModal = document.getElementById('settings-modal');
  const modalOverlay = document.getElementById('modal-overlay');
  const modalClose = document.getElementById('modal-close');
  
  const container = document.querySelector('.container');
  const formatSearchInput = document.getElementById('format-search');
  const formatSearchClear = document.getElementById('format-search-clear');
  const missingSearchInput = document.getElementById('missing-search');
  const missingSearchClear = document.getElementById('missing-search-clear');

  // ── Search state ──
  const srch = {
    query: '',
    mode: 'contains',
    caseSensitive: false,
    cols: [],         // active columns to search
    allCols: [],      // all columns from sheet
    rows: [],         // flat object rows derived from currentRows
    page: 1,
    pageSize: 50,
  };

  let allFormatIssues = [];
  let allMissingIssues = [];
  let currentScannedCount = 0;
  
  const scanResults = {
    format: { page: 1, pageSize: 25, query: '' },
    missing: { page: 1, pageSize: 25, query: '' }
  };

  langFilter.addEventListener('change', () => {
    filterResults();
  });

  // Handle inline translations in the missing table
  missingBody.addEventListener('click', async (e) => {
    const copyBtn = e.target.closest('.inline-copy-btn');
    if (copyBtn) {
      const input = copyBtn.previousElementSibling;
      const textToCopy = input ? input.value : copyBtn.dataset.translation;

      navigator.clipboard.writeText(textToCopy).then(() => {
        copyBtn.classList.add('copied');
        setTimeout(() => copyBtn.classList.remove('copied'), 2000);
        showToast('Translation copied!');
      });
      return;
    }

    const btn = e.target.closest('.inline-translate-btn');
    if (!btn) return;

    const englishText = btn.dataset.text;
    const langCode = btn.dataset.lang;

    // Use default 'auto' if language wasn't matched properly
    const safeLangCode = (langCode === 'auto') ? btn.closest('tr').children[1].textContent : langCode;

    btn.textContent = 'Translating...';
    btn.disabled = true;

    const responseObj = await fetchTranslation(englishText, 'en', safeLangCode === 'auto' ? '' : langCode);
    const translation = responseObj.text;

    if (translation && !translation.startsWith('Error')) {
      const td = btn.closest('td');
      const showSave = !!currentSheetRef;
      td.innerHTML = `
        <div style="display: flex; gap: 0.5rem; align-items: center;">
          <input type="text" class="text-snippet inline-translation-input" value="${escapeHtml(translation).replace(/"/g, '&quot;')}" style="background: rgba(63,185,80,0.1); border: 1px solid rgba(63,185,80,0.3); color: var(--success); width: 100%; padding: 0.2rem 0.5rem; outline: none; border-radius: 4px; font-family: var(--font);" />
          <button class="qt-copy-btn inline-copy-btn" title="Copy" style="position: static; flex-shrink: 0;">
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
          </button>
          ${showSave ? `<button class="btn sm-btn inline-save-btn" title="Save to Google Sheet" style="flex-shrink: 0;">
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>
            Save
          </button>` : ''}
        </div>
      `;
    } else {
      btn.textContent = 'Failed. Retry?';
      btn.disabled = false;
    }
  });

  // Inline save-to-sheet handler
  missingBody.addEventListener('click', async (e) => {
    const saveBtn = e.target.closest('.inline-save-btn');
    if (!saveBtn) return;
    if (!currentSheetRef) {
      showToast('No Google Sheet is loaded.', 'error');
      return;
    }
    const tr = saveBtn.closest('tr');
    const input = tr.querySelector('.inline-translation-input');
    if (!input) return;
    const value = input.value;
    const rowNum = Number(tr.dataset.rowNum);
    const colIndex = Number(tr.dataset.colIndex);
    if (!rowNum || Number.isNaN(colIndex)) {
      showToast('Cannot determine target cell.', 'error');
      return;
    }
    const originalLabel = saveBtn.innerHTML;
    saveBtn.disabled = true;
    saveBtn.textContent = 'Saving…';
    try {
      await writeSheetCell({
        spreadsheetId: currentSheetRef.spreadsheetId,
        sheetTitle: currentSheetRef.sheetTitle,
        rowNum,
        colIndex,
        value,
      });
      // Update local cache so re-validation reflects the change
      if (currentRows && currentRows[rowNum - 1]) {
        currentRows[rowNum - 1][colIndex] = value;
        saveToDB(currentFileName, currentRows, currentSheetRef);
      }
      input.style.borderColor = 'var(--success)';
      saveBtn.classList.add('saved');
      saveBtn.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
        Saved
      `;
      showToast('Saved to Google Sheet.');
    } catch (err) {
      saveBtn.disabled = false;
      saveBtn.innerHTML = originalLabel;
      showToast(err.message || 'Failed to save.', 'error');
    }
  });

  let currentRows = null;
  let currentFileName = "";
  const DB_NAME = 'LocaLinterDB';
  const STORE_NAME = 'cache';

  function saveToDB(fileName, rows, sheetRef = null) {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = (e) => e.target.result.createObjectStore(STORE_NAME);
    request.onsuccess = (e) => e.target.result.transaction(STORE_NAME, 'readwrite').objectStore(STORE_NAME).put({ fileName, rows, sheetRef }, 'currentFile');
  }

  function loadFromDB(callback) {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = (e) => e.target.result.createObjectStore(STORE_NAME);
    request.onsuccess = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) return;
      const getReq = db.transaction(STORE_NAME, 'readonly').objectStore(STORE_NAME).get('currentFile');
      getReq.onsuccess = () => { if (getReq.result) callback(getReq.result); };
    };
  }

  function clearDB() {
    const request = indexedDB.open(DB_NAME, 1);
    request.onsuccess = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) return;
      db.transaction(STORE_NAME, 'readwrite').objectStore(STORE_NAME).delete('currentFile');
    };
  }

  // Load Ignore List from LocalStorage
  const savedIgnoreList = localStorage.getItem('localinter_ignore_list') || '';
  if (savedIgnoreList) {
    ignoreListInput.value = savedIgnoreList;
  }

  // ── Toast helper ──
  function showToast(message, type = 'success') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.innerHTML = `<span class="toast-dot"></span>${message}`;
    container.appendChild(toast);
    setTimeout(() => {
      toast.classList.add('hide');
      toast.addEventListener('animationend', () => toast.remove());
    }, 2500);
  }

  saveIgnoreBtn.addEventListener('click', () => {
    localStorage.setItem('localinter_ignore_list', ignoreListInput.value);
    showToast('Ignore list saved!');
    closeSettings();
    if (currentRows) {
      validateData(currentRows);
    }
  });

  function openSettings() {
    settingsModal.classList.remove('hidden');
    modalOverlay.classList.remove('hidden');
    document.body.style.overflow = 'hidden';
  }

  function closeSettings() {
    settingsModal.classList.add('hidden');
    modalOverlay.classList.add('hidden');
    document.body.style.overflow = '';
  }

  globalSettingsBtn.addEventListener('click', openSettings);
  modalClose.addEventListener('click', closeSettings);
  modalOverlay.addEventListener('click', closeSettings);

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeSettings();
  });

  removeFileBtn.addEventListener('click', () => {
    clearDB();
    currentRows = null;
    currentFileName = "";
    currentSheetRef = null;
    const refreshBtn = document.getElementById('refresh-sheet-btn');
    if (refreshBtn) refreshBtn.classList.add('hidden');
    const linkWrap = document.getElementById('link-sync-wrap');
    if (linkWrap) linkWrap.classList.add('hidden');
    container.classList.remove('has-results');
    dropzone.classList.remove('file-loaded');
    dropContentIdle.classList.remove('hidden');
    loadedContent.classList.add('hidden');
    resultsContainer.classList.add('hidden');
    fileInput.value = '';
    
    // Clear tab searches
    formatSearchInput.value = '';
    scanResults.format.query = '';
    missingSearchInput.value = '';
    scanResults.missing.query = '';
  });

  function switchTab(tabId) {
    const targetBtn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
    if (!targetBtn) return;
    tabBtns.forEach(b => b.classList.remove('active'));
    tabContents.forEach(c => c.classList.remove('active'));
    targetBtn.classList.add('active');
    const content = document.getElementById('tab-' + tabId);
    if (content) content.classList.add('active');

    // UI Isolation for Search Tab
    const isSearch = tabId === 'search';
    if (globalSettingsBtn) {
      if (isSearch) globalSettingsBtn.classList.add('v-hidden');
      else globalSettingsBtn.classList.remove('v-hidden');
    }
    if (resultsHeader) {
      if (isSearch) resultsHeader.classList.add('hidden');
      else resultsHeader.classList.remove('hidden');
    }

    localStorage.setItem('locaLinterActiveTab', tabId);
  }

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      switchTab(btn.dataset.tab);
    });
  });

  // ── Tab Search Listeners ──
  formatSearchInput.addEventListener('input', () => {
    scanResults.format.query = formatSearchInput.value;
    formatSearchClear.style.display = scanResults.format.query ? 'flex' : 'none';
    filterResults();
  });

  formatSearchClear.addEventListener('click', () => {
    formatSearchInput.value = '';
    scanResults.format.query = '';
    formatSearchClear.style.display = 'none';
    filterResults();
    formatSearchInput.focus();
  });

  missingSearchInput.addEventListener('input', () => {
    scanResults.missing.query = missingSearchInput.value;
    missingSearchClear.style.display = scanResults.missing.query ? 'flex' : 'none';
    filterResults();
  });

  missingSearchClear.addEventListener('click', () => {
    missingSearchInput.value = '';
    scanResults.missing.query = '';
    missingSearchClear.style.display = 'none';
    filterResults();
    missingSearchInput.focus();
  });

  browseBtn.addEventListener('click', (e) => {
    e.preventDefault();
    fileInput.click();
  });

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  ['dragleave', 'dragend'].forEach(type => {
    dropzone.addEventListener(type, () => {
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      handleFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
      handleFile(e.target.files[0]);
    }
  });

  function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = window.XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const json = window.XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      currentFileName = file.name;
      currentRows = json;
      currentSheetRef = null;
      const refreshBtn = document.getElementById('refresh-sheet-btn');
      if (refreshBtn) refreshBtn.classList.add('hidden');
      saveToDB(currentFileName, currentRows);

      showLoadedState(currentFileName);
      validateData(currentRows);
      initSearchTab(currentRows);
    };
    reader.readAsArrayBuffer(file);
  }

  function showLoadedState(name) {
    dropzone.classList.add('file-loaded');
    dropContentIdle.classList.add('hidden');
    loadedContent.classList.remove('hidden');
    loadedFileName.textContent = `Loaded: ${name}`;
    if (typeof updateLinkSyncUI === 'function') updateLinkSyncUI();
  }

  // ── Google Sheets loader (OAuth via Google Identity Services) ──
  const gsheetUrlInput = document.getElementById('gsheet-url');
  const gsheetLoadBtn = document.getElementById('gsheet-load-btn');
  const gsheetSignInBtn = document.getElementById('gsheet-signin-btn');
  const gsheetAuthStatus = document.getElementById('gsheet-auth-status');
  const headerSignInBtn = document.getElementById('header-signin-btn');
  const accountWidget = document.getElementById('account-widget');
  const accountBtn = document.getElementById('account-btn');
  const accountAvatar = document.getElementById('account-avatar');
  const accountInitials = document.getElementById('account-initials');
  const accountMenu = document.getElementById('account-menu');
  const accountMenuAvatar = document.getElementById('account-menu-avatar');
  const accountMenuInitials = document.getElementById('account-menu-initials');
  const accountMenuName = document.getElementById('account-menu-name');
  const accountMenuEmail = document.getElementById('account-menu-email');
  const accountSignOutBtn = document.getElementById('account-signout-btn');
  const SHEETS_SCOPE = 'https://www.googleapis.com/auth/userinfo.email https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/spreadsheets';
  const TOKEN_STORAGE_KEY = 'localinter_gsheet_token_v3';
  const USER_STORAGE_KEY = 'localinter_gsheet_user_v3';

  function getOAuthClientId() {
    const meta = document.querySelector('meta[name="google-oauth-client-id"]');
    return meta ? (meta.getAttribute('content') || '').trim() : '';
  }

  function loadStoredToken() {
    try {
      const raw = sessionStorage.getItem(TOKEN_STORAGE_KEY);
      if (!raw) return null;
      const t = JSON.parse(raw);
      if (!t.access_token || !t.expires_at || Date.now() >= t.expires_at) {
        sessionStorage.removeItem(TOKEN_STORAGE_KEY);
        return null;
      }
      return t;
    } catch { return null; }
  }

  function storeToken(tok) {
    sessionStorage.setItem(TOKEN_STORAGE_KEY, JSON.stringify(tok));
  }

  function clearToken() {
    sessionStorage.removeItem(TOKEN_STORAGE_KEY);
    sessionStorage.removeItem(USER_STORAGE_KEY);
  }

  function loadStoredUser() {
    try {
      const raw = sessionStorage.getItem(USER_STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch { return null; }
  }

  function storeUser(user) {
    sessionStorage.setItem(USER_STORAGE_KEY, JSON.stringify(user));
  }

  function getInitials(name, email) {
    const src = (name || email || '?').trim();
    const parts = src.split(/\s+/).filter(Boolean);
    if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
    return src.slice(0, 2).toUpperCase();
  }

  function setAvatar(imgEl, initialsEl, user) {
    const initials = getInitials(user.name, user.email);
    initialsEl.textContent = initials;
    if (user.picture) {
      imgEl.src = user.picture;
      imgEl.classList.remove('hidden');
      initialsEl.classList.add('hidden');
      imgEl.onerror = () => {
        imgEl.classList.add('hidden');
        initialsEl.classList.remove('hidden');
      };
    } else {
      imgEl.classList.add('hidden');
      initialsEl.classList.remove('hidden');
    }
  }

  function updateAuthUI() {
    const tok = loadStoredToken();
    const user = loadStoredUser();
    const authRow = document.querySelector('.gsheet-auth-row');
    if (tok && user) {
      if (authRow) authRow.classList.add('signed-in');
      gsheetSignInBtn.classList.add('hidden');
      headerSignInBtn.classList.add('hidden');
      accountWidget.classList.remove('hidden');
      setAvatar(accountAvatar, accountInitials, user);
      setAvatar(accountMenuAvatar, accountMenuInitials, user);
      accountMenuName.textContent = user.name || 'Google account';
      accountMenuEmail.textContent = user.email || '';
    } else {
      if (authRow) authRow.classList.remove('signed-in');
      gsheetAuthStatus.textContent = 'Sign in to access private sheets';
      gsheetSignInBtn.classList.remove('hidden');
      headerSignInBtn.classList.remove('hidden');
      accountWidget.classList.add('hidden');
      accountMenu.classList.add('hidden');
      accountBtn.setAttribute('aria-expanded', 'false');
    }
  }

  async function fetchUserInfo(accessToken) {
    const res = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${accessToken}` },
    });
    if (!res.ok) throw new Error('Failed to fetch user profile.');
    const data = await res.json();
    return { name: data.name, email: data.email, picture: data.picture };
  }

  function requestAccessToken({ silent = false } = {}) {
    return new Promise((resolve, reject) => {
      const clientId = getOAuthClientId();
      if (!clientId) {
        reject(new Error('Google OAuth Client ID is not configured. Set the meta tag in index.html.'));
        return;
      }
      if (!window.google || !google.accounts || !google.accounts.oauth2) {
        reject(new Error('Google Identity Services not loaded yet. Try again in a moment.'));
        return;
      }
      const tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SHEETS_SCOPE,
        prompt: silent ? '' : 'consent',
        callback: async (response) => {
          if (response.error) {
            reject(new Error(response.error_description || response.error));
            return;
          }
          const tok = {
            access_token: response.access_token,
            expires_at: Date.now() + (Number(response.expires_in || 3600) - 60) * 1000,
          };
          storeToken(tok);
          try {
            const user = await fetchUserInfo(tok.access_token);
            storeUser(user);
          } catch { /* avatar/profile is optional */ }
          updateAuthUI();
          resolve(tok);
        },
        error_callback: (err) => reject(new Error(err && err.message ? err.message : 'Sign-in cancelled')),
      });
      tokenClient.requestAccessToken({ prompt: silent ? '' : 'consent' });
    });
  }

  async function getValidToken() {
    const cached = loadStoredToken();
    if (cached) return cached;
    try {
      return await requestAccessToken({ silent: true });
    } catch {
      return requestAccessToken({ silent: false });
    }
  }

  function parseSheetUrl(url) {
    if (!url) return null;
    const idMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!idMatch) return null;
    const gidMatch = url.match(/[#?&]gid=([0-9]+)/);
    return { spreadsheetId: idMatch[1], gid: gidMatch ? Number(gidMatch[1]) : null };
  }

  function colIndexToA1(idx) {
    let n = idx + 1, s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  async function writeSheetCell({ spreadsheetId, sheetTitle, rowNum, colIndex, value }) {
    const token = await getValidToken();
    const a1 = `${sheetTitle}!${colIndexToA1(colIndex)}${rowNum}`;
    const res = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(a1)}?valueInputOption=RAW`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token.access_token}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ values: [[value]] }),
      },
    );
    if (res.status === 401) {
      clearToken();
      updateAuthUI();
      throw new Error('Session expired. Please sign in again.');
    }
    if (!res.ok) {
      let msg = `Sheets write failed (${res.status})`;
      try {
        const err = await res.json();
        if (err.error && err.error.message) msg = err.error.message;
      } catch {}
      throw new Error(msg);
    }
    return res.json();
  }

  async function sheetsApi(path, token) {
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets${path}`, {
      headers: { Authorization: `Bearer ${token.access_token}` },
    });
    if (res.status === 401) {
      clearToken();
      updateAuthUI();
      throw new Error('Session expired. Please sign in again.');
    }
    if (!res.ok) {
      let msg = `Sheets API error (${res.status})`;
      try {
        const err = await res.json();
        if (err.error && err.error.message) msg = err.error.message;
      } catch {}
      throw new Error(msg);
    }
    return res.json();
  }

  const refreshSheetBtn = document.getElementById('refresh-sheet-btn');
  let currentSheetRef = null; // { spreadsheetId, gid, sheetTitle }

  async function fetchSheetRows({ spreadsheetId, gid, preferredTitle }) {
    const token = await getValidToken();
    let sheetTitle = preferredTitle;
    let workbookTitle = '';
    const meta = await sheetsApi(
      `/${spreadsheetId}?fields=properties.title,sheets.properties(title,sheetId)`,
      token,
    );
    workbookTitle = meta.properties.title;
    let chosen = null;
    if (preferredTitle) {
      chosen = (meta.sheets || []).find(s => s.properties.title === preferredTitle);
    }
    if (!chosen && gid != null) {
      chosen = (meta.sheets || []).find(s => s.properties.sheetId === gid);
    }
    if (!chosen) chosen = (meta.sheets || [])[0];
    if (!chosen) throw new Error('Spreadsheet has no sheets.');
    sheetTitle = chosen.properties.title;
    const data = await sheetsApi(
      `/${spreadsheetId}/values/${encodeURIComponent(sheetTitle)}`,
      token,
    );
    return { rows: data.values || [], sheetTitle, workbookTitle, sheetId: chosen.properties.sheetId };
  }

  function setLoadBtnLoading(isLoading) {
    const def = gsheetLoadBtn.querySelector('.gsheet-icon-default');
    const ld = gsheetLoadBtn.querySelector('.gsheet-icon-loading');
    const label = gsheetLoadBtn.querySelector('.gsheet-load-label');
    if (isLoading) {
      gsheetLoadBtn.classList.add('is-loading');
      gsheetLoadBtn.disabled = true;
      def && def.classList.add('hidden');
      ld && ld.classList.remove('hidden');
      if (label) label.textContent = 'Loading…';
    } else {
      gsheetLoadBtn.classList.remove('is-loading');
      gsheetLoadBtn.disabled = false;
      def && def.classList.remove('hidden');
      ld && ld.classList.add('hidden');
      if (label) label.textContent = 'Load Sheet';
    }
  }

  async function loadFromSheetUrl(rawUrl) {
    const parsed = parseSheetUrl(rawUrl);
    if (!parsed) {
      showToast('Invalid Google Sheets URL.', 'error');
      return;
    }
    try {
      setLoadBtnLoading(true);
      showToast('Fetching sheet from Google…');
      const { rows, sheetTitle, workbookTitle, sheetId } = await fetchSheetRows({
        spreadsheetId: parsed.spreadsheetId,
        gid: parsed.gid,
      });
      if (!rows.length) {
        showToast('Sheet is empty.', 'error');
        return;
      }
      currentFileName = `${workbookTitle} — ${sheetTitle}`;
      currentRows = rows;
      currentSheetRef = { spreadsheetId: parsed.spreadsheetId, gid: sheetId, sheetTitle };
      saveToDB(currentFileName, currentRows, currentSheetRef);
      showLoadedState(currentFileName);
      validateData(currentRows);
      initSearchTab(currentRows);
      refreshSheetBtn.classList.remove('hidden');
      updateLinkSyncUI();
      showToast('Sheet loaded.');
    } catch (e) {
      showToast(e.message || 'Failed to load sheet.', 'error');
    } finally {
      setLoadBtnLoading(false);
    }
  }

  async function refreshCurrentSheet() {
    if (!currentSheetRef) return;
    try {
      refreshSheetBtn.disabled = true;
      refreshSheetBtn.classList.add('is-loading');
      const { rows, sheetTitle, workbookTitle, sheetId } = await fetchSheetRows({
        spreadsheetId: currentSheetRef.spreadsheetId,
        gid: currentSheetRef.gid,
        preferredTitle: currentSheetRef.sheetTitle,
      });
      if (!rows.length) {
        showToast('Sheet is empty.', 'error');
        return;
      }
      currentFileName = `${workbookTitle} — ${sheetTitle}`;
      currentRows = rows;
      currentSheetRef = { spreadsheetId: currentSheetRef.spreadsheetId, gid: sheetId, sheetTitle };
      saveToDB(currentFileName, currentRows, currentSheetRef);
      loadedFileName.textContent = `Loaded: ${currentFileName}`;
      validateData(currentRows);
      initSearchTab(currentRows);
      linkedSheetCache.delete(currentSheetRef.spreadsheetId);
      updateLinkSyncUI();
      showToast('Sheet refreshed.');
    } catch (e) {
      showToast(e.message || 'Failed to refresh.', 'error');
    } finally {
      refreshSheetBtn.disabled = false;
      refreshSheetBtn.classList.remove('is-loading');
    }
  }

  refreshSheetBtn.addEventListener('click', refreshCurrentSheet);

  gsheetLoadBtn.addEventListener('click', () => loadFromSheetUrl(gsheetUrlInput.value.trim()));
  gsheetUrlInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      loadFromSheetUrl(gsheetUrlInput.value.trim());
    }
  });

  // ── Sheet presets (Dev / Staging / Global Prod) ──
  const SHEET_PRESETS = {
    dev:     { id: '1yyDiN-QlqXXNzFKVSA1NOINMBQyEex-0VYcnRwzZKzY', label: 'Dev' },
    staging: { id: '17lZ5oMqMuCxX6dig--YvsmAM-MuP3w_nBtDgu3_z45A', label: 'Staging' },
    prod:    { id: '1ucDqgNFb74WZCYbTVQSoUW2cTfBDZodR5ypSfC-fBBQ', label: 'Global Prod' },
  };

  function presetKeyForSpreadsheetId(id) {
    if (!id) return null;
    const found = Object.entries(SHEET_PRESETS).find(([, p]) => p.id === id);
    return found ? found[0] : null;
  }

  document.querySelectorAll('.quick-open-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const preset = SHEET_PRESETS[btn.dataset.preset];
      if (!preset) return;
      loadFromSheetUrl(`https://docs.google.com/spreadsheets/d/${preset.id}/edit?gid=0#gid=0`);
    });
  });

  // ── Linked-sheet edit propagation ──
  const linkedSheetCache = new Map(); // spreadsheetId -> { sheetTitle, headerRow, keyToRowNum }
  const linkedSheetDenied = new Set(); // spreadsheetIds we've confirmed are inaccessible (403/404)
  const linkSyncToggle = document.getElementById('link-sync-toggle');
  const linkSyncWrap = document.getElementById('link-sync-wrap');
  const linkSyncCurrent = document.getElementById('link-sync-current');
  const linkSyncRefreshBtn = document.getElementById('link-sync-refresh-btn');

  if (linkSyncToggle) {
    linkSyncToggle.checked = localStorage.getItem('locaLinterLinkSync') === '1';
    linkSyncToggle.addEventListener('change', () => {
      localStorage.setItem('locaLinterLinkSync', linkSyncToggle.checked ? '1' : '0');
      if (linkSyncToggle.checked) linkedSheetCache.clear();
    });
  }
  if (linkSyncRefreshBtn) {
    linkSyncRefreshBtn.addEventListener('click', () => {
      linkedSheetCache.clear();
      linkedSheetDenied.clear();
      showToast('Linked sheets cache cleared.');
    });
  }

  function linkSyncEnabled() {
    return !!(linkSyncToggle && linkSyncToggle.checked);
  }

  function updateLinkSyncUI() {
    if (!linkSyncWrap) return;
    const presetKey = currentSheetRef && presetKeyForSpreadsheetId(currentSheetRef.spreadsheetId);
    if (presetKey) {
      linkSyncWrap.classList.remove('hidden');
      const others = Object.entries(SHEET_PRESETS)
        .filter(([k]) => k !== presetKey)
        .map(([, p]) => p.label)
        .join(' & ');
      if (linkSyncCurrent) {
        linkSyncCurrent.textContent = `Current: ${SHEET_PRESETS[presetKey].label} → mirrors to ${others}`;
      }
    } else {
      linkSyncWrap.classList.add('hidden');
    }
  }

  function normalizeKey(s) {
    return ((s || '') + '').trim().toLowerCase();
  }
  function normalizeHeader(s) {
    return ((s || '') + '').trim().replace(/^\$+/, '').toLowerCase();
  }

  async function getLinkedSheetData(spreadsheetId) {
    if (linkedSheetCache.has(spreadsheetId)) return linkedSheetCache.get(spreadsheetId);
    const { rows, sheetTitle } = await fetchSheetRows({ spreadsheetId, gid: 0 });
    const headerRow = rows[0] || [];
    const keyToRowNums = new Map(); // key -> [rowNum, ...]
    for (let i = 1; i < rows.length; i++) {
      const k = normalizeKey(rows[i] && rows[i][0]);
      if (!k) continue;
      const list = keyToRowNums.get(k);
      if (list) list.push(i + 1);
      else keyToRowNums.set(k, [i + 1]);
    }
    const data = { sheetTitle, headerRow, keyToRowNums };
    linkedSheetCache.set(spreadsheetId, data);
    return data;
  }

  function findHeaderColIndex(headerRow, target) {
    const t = normalizeHeader(target);
    if (!t) return -1;
    for (let i = 0; i < headerRow.length; i++) {
      if (normalizeHeader(headerRow[i]) === t) return i;
    }
    return -1;
  }

  async function propagateEditToLinkedSheets({ rowNum, colIndex, value, sourceKey }) {
    if (!linkSyncEnabled() || !currentSheetRef) return;
    const presetKey = presetKeyForSpreadsheetId(currentSheetRef.spreadsheetId);
    if (!presetKey) return;
    if (!currentRows || !currentRows[rowNum - 1] || !currentRows[0]) return;
    if (colIndex === 0) {
      showToast('Linked sync skipped — editing the key column is not propagated.', 'error');
      return;
    }

    const key = sourceKey != null
      ? ((sourceKey + '').trim())
      : ((currentRows[rowNum - 1][0] || '') + '').trim();
    const langHeader = ((currentRows[0][colIndex] || '') + '').trim();
    if (!key) { showToast('Linked sync skipped — row has no key.', 'error'); return; }
    if (!langHeader) { showToast('Linked sync skipped — column has no header.', 'error'); return; }

    const targets = Object.entries(SHEET_PRESETS).filter(([k]) => k !== presetKey);
    const results = [];
    for (const [, target] of targets) {
      if (linkedSheetDenied.has(target.id)) {
        results.push(`${target.label}: no access (skipped)`);
        continue;
      }
      try {
        const data = await getLinkedSheetData(target.id);
        const targetRows = data.keyToRowNums.get(normalizeKey(key));
        const targetCol = findHeaderColIndex(data.headerRow, langHeader);
        if (!targetRows || !targetRows.length) { results.push(`${target.label}: key not found`); continue; }
        if (targetCol < 0) { results.push(`${target.label}: language not found`); continue; }
        for (const rn of targetRows) {
          await writeSheetCell({
            spreadsheetId: target.id,
            sheetTitle: data.sheetTitle,
            rowNum: rn,
            colIndex: targetCol,
            value,
          });
        }
        const suffix = targetRows.length > 1 ? ` (${targetRows.length} rows)` : '';
        results.push(`${target.label} ✓${suffix}`);
      } catch (e) {
        const msg = (e && e.message) || 'error';
        if (/permission|403|not found|404/i.test(msg)) {
          linkedSheetDenied.add(target.id);
          results.push(`${target.label}: no access — won't retry this session`);
        } else {
          results.push(`${target.label}: ${msg}`);
        }
      }
    }
    if (results.length) {
      const ok = results.every(r => r.includes('✓'));
      showToast('Linked sync — ' + results.join(', '), ok ? undefined : 'error');
    }
  }

  async function doSignIn() {
    try { await requestAccessToken({ silent: false }); showToast('Signed in to Google.'); }
    catch (e) { showToast(e.message || 'Sign-in failed.', 'error'); }
  }
  function doSignOut() {
    const tok = loadStoredToken();
    if (tok && window.google && google.accounts && google.accounts.oauth2) {
      try { google.accounts.oauth2.revoke(tok.access_token, () => {}); } catch {}
    }
    clearToken();
    updateAuthUI();
    showToast('Signed out.');
  }

  gsheetSignInBtn.addEventListener('click', doSignIn);
  headerSignInBtn.addEventListener('click', doSignIn);
  accountSignOutBtn.addEventListener('click', () => {
    accountMenu.classList.add('hidden');
    accountBtn.setAttribute('aria-expanded', 'false');
    doSignOut();
  });
  accountBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    const isHidden = accountMenu.classList.toggle('hidden');
    accountBtn.setAttribute('aria-expanded', String(!isHidden));
  });
  document.addEventListener('click', (e) => {
    if (!accountMenu.classList.contains('hidden') && !accountWidget.contains(e.target)) {
      accountMenu.classList.add('hidden');
      accountBtn.setAttribute('aria-expanded', 'false');
    }
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !accountMenu.classList.contains('hidden')) {
      accountMenu.classList.add('hidden');
      accountBtn.setAttribute('aria-expanded', 'false');
    }
  });
  updateAuthUI();

  // Initial load check
  loadFromDB((data) => {
    currentFileName = data.fileName;
    currentRows = data.rows;
    currentSheetRef = data.sheetRef || null;
    showLoadedState(currentFileName);

    const willAutoRefresh = currentSheetRef && loadStoredToken();
    if (!willAutoRefresh) {
      validateData(currentRows);
      initSearchTab(currentRows);
    }
    if (currentSheetRef) {
      refreshSheetBtn.classList.remove('hidden');
      if (willAutoRefresh) {
        // Yield two frames so the spinner animation kicks off on the
        // compositor before validateData blocks the main thread.
        requestAnimationFrame(() => requestAnimationFrame(() => refreshCurrentSheet()));
      }
    }

    // Restore active tab after data is ready
    const savedTab = localStorage.getItem('locaLinterActiveTab');
    if (savedTab) {
      switchTab(savedTab);
    }
  });

  function checkBrackets(text) {
    const stack = [];
    const pairs = { ')': '(', ']': '[', '}': '{' };
    const opening = ['(', '[', '{'];
    for (let i = 0; i < text.length; i++) {
      const char = text[i];
      if (opening.includes(char)) {
        stack.push({ char, index: i });
      } else if (pairs[char]) {
        if (stack.length === 0) {
          return `Unmatched closing bracket '${char}'`;
        }
        const last = stack.pop();
        if (last.char !== pairs[char]) {
          return `Mismatched bracket '${char}' (expected closing for '${last.char}')`;
        }
      }
    }
    if (stack.length > 0) {
      return `Unclosed bracket '${stack[stack.length - 1].char}'`;
    }
    return null;
  }

  function extractVars(text) {
    const varRegex = /{[a-zA-Z0-9_.]+}/g;
    return text.match(varRegex) || [];
  }

  function extractTags(text) {
    const tagRegex = /<[^>]+>/g;
    return text.match(tagRegex) || [];
  }

  function countLineBreaks(text) {
    return (text.match(/\n/g) || []).length;
  }

  function validateData(rows) {
    if (!rows || rows.length < 2) return;
    const headers = rows[0];

    // Get current ignore list
    const ignoreVal = (ignoreListInput.value || '').toLowerCase();
    const ignoredTermsList = ignoreVal.split(',').map(s => s.trim()).filter(Boolean);

    // Find base language (English)
    let englishColIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].toLowerCase() === 'english') {
        englishColIndex = i;
        break;
      }
    }

    if (englishColIndex === -1) {
      // Fallback: assume column 1 is English
      englishColIndex = 1;
    }

    const formatIssues = [];
    const missingIssues = [];
    let rowsScanned = rows.length - 1;

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      if (!row || row.length === 0) continue;

      const keyInfo = row[0] || `Row ${rowIndex + 1}`;
      const englishText = row[englishColIndex] ? String(row[englishColIndex]) : "";
      const englishVars = extractVars(englishText);

      // Check if row has ANY target translations
      let hasAnyTranslation = false;
      for (let col = 1; col < headers.length; col++) {
        if (col === englishColIndex) continue;
        if (row[col] && String(row[col]).trim() !== "") {
          hasAnyTranslation = true;
          break;
        }
      }

      // Validate base language first (English)
      if (englishText) {
        let baseBracketErr = checkBrackets(englishText);
        if (baseBracketErr) {
          formatIssues.push({ key: keyInfo, rowNum: rowIndex + 1, lang: headers[englishColIndex] || "English", err: `Base text err: ${baseBracketErr}`, snippet: englishText });
        }
      }

      // If there are zero translations for this string across all languages, it's likely a non-localized string (character name, weapon skin, etc.). Skip flagging missing languages.
      if (!hasAnyTranslation) {
        continue;
      }

      // Ignore specifically requested terms (matching english text or keyInfo substring)
      const isIgnored = ignoredTermsList.some(term =>
        (keyInfo && keyInfo.toLowerCase().includes(term)) ||
        (englishText && englishText.toLowerCase() === term)
      );

      // Hardcode exclusion for voice comms (e.g. Config.AudioEmitter.VoiceComms)
      const isConfigVoiceComm = keyInfo && keyInfo.toLowerCase().startsWith('config.');

      if (isIgnored || isConfigVoiceComm) {
        continue;
      }

      for (let col = 1; col < headers.length; col++) {
        if (col === englishColIndex) continue;

        let targetText = row[col];
        // Strictly evaluate if the cell is completely empty or just whitespace
        if (targetText === undefined || targetText === null || String(targetText).trim() === "") {
          if (englishText && String(englishText).trim() !== "") {
            missingIssues.push({
              key: keyInfo,
              rowNum: rowIndex + 1,
              colIndex: col,
              lang: headers[col] || `Col ${col}`,
              englishText: englishText
            });
          }
          continue;
        }

        // It has text, ensure we operate on a string
        targetText = String(targetText);

        // 1. Check Brackets (General nesting/mismatch)
        let bracketErr = checkBrackets(targetText);
        let formatErrs = [];
        if (bracketErr) formatErrs.push(bracketErr);

        // 2. Check Variables match
        if (englishText) {
          const englishVars = extractVars(englishText);
          const targetVars = extractVars(targetText);
          const missingVars = englishVars.filter(v => !targetVars.includes(v));
          const extraVars = targetVars.filter(v => !englishVars.includes(v));

          if (missingVars.length > 0) formatErrs.push(`Missing vars: ${missingVars.join(', ')}`);
          if (extraVars.length > 0) formatErrs.push(`Extra vars: ${extraVars.join(', ')}`);

          // 3. Check Tags match (e.g. <color=red>)
          const englishTags = extractTags(englishText);
          const targetTags = extractTags(targetText);
          const missingTags = englishTags.filter(t => !targetTags.includes(t));
          const extraTags = targetTags.filter(t => !englishTags.includes(t));

          if (missingTags.length > 0) formatErrs.push(`Missing tags: ${missingTags.join(' ')}`);
          if (extraTags.length > 0) formatErrs.push(`Extra tags: ${extraTags.join(' ')}`);

          // 4. Check Newlines (\n)
          const englishNL = countLineBreaks(englishText);
          const targetNL = countLineBreaks(targetText);
          if (englishNL !== targetNL) {
            formatErrs.push(`Newline mismatch: expected ${englishNL}, found ${targetNL}`);
          }

          // 5. Check Leading/Trailing Whitespace
          if (englishText.startsWith(' ') && !targetText.startsWith(' ')) formatErrs.push('Missing leading space');
          if (!englishText.startsWith(' ') && targetText.startsWith(' ')) formatErrs.push('Extra leading space');
          if (englishText.endsWith(' ') && !targetText.endsWith(' ')) formatErrs.push('Missing trailing space');
          if (!englishText.endsWith(' ') && targetText.endsWith(' ')) formatErrs.push('Extra trailing space');
        }

        // 6. Check for Double Spaces (Internal consistency)
        if (targetText.includes('  ')) {
          formatErrs.push('Contains double spaces');
        }

        if (formatErrs.length > 0) {
          formatIssues.push({
            key: keyInfo,
            rowNum: rowIndex + 1,
            lang: headers[col] || `Col ${col}`,
            err: formatErrs.join('; '),
            snippet: targetText
          });
        }
      }
    }

    allFormatIssues = formatIssues;
    allMissingIssues = missingIssues;
    currentScannedCount = rowsScanned;

    populateFilter(formatIssues, missingIssues);
    renderResults(rowsScanned, formatIssues, missingIssues);
  }

  function populateFilter(formatIssues, missingIssues) {
    const allIssues = [...formatIssues, ...missingIssues];
    const languages = [...new Set(allIssues.map(i => i.lang))].sort();
    const currentValue = langFilter.value;

    // Clear except "All"
    langFilter.innerHTML = '<option value="all">All Languages</option>';

    languages.forEach(lang => {
      const option = document.createElement('option');
      option.value = lang;
      option.textContent = lang;
      langFilter.appendChild(option);
    });

    // Restore value if it still exists
    if ([...langFilter.options].some(o => o.value === currentValue)) {
      langFilter.value = currentValue;
    }
  }

  function filterResults() {
    const selectedLang = langFilter.value;
    const fQuery = scanResults.format.query.toLowerCase();
    const mQuery = scanResults.missing.query.toLowerCase();

    const filteredFormat = allFormatIssues.filter(i => {
      const matchLang = selectedLang === 'all' || i.lang === selectedLang;
      const matchQuery = !fQuery || i.key.toLowerCase().includes(fQuery) || i.err.toLowerCase().includes(fQuery) || i.snippet.toLowerCase().includes(fQuery);
      return matchLang && matchQuery;
    });

    const filteredMissing = allMissingIssues.filter(i => {
      const matchLang = selectedLang === 'all' || i.lang === selectedLang;
      const matchQuery = !mQuery || i.key.toLowerCase().includes(mQuery) || i.englishText.toLowerCase().includes(mQuery);
      return matchLang && matchQuery;
    });

    // Reset pagination on filter
    scanResults.format.page = 1;
    scanResults.missing.page = 1;

    renderResults(currentScannedCount, filteredFormat, filteredMissing, true);
  }

  function renderPagination(container, currentPage, totalPages, totalItems, pageSize, onPageChange) {
    container.innerHTML = '';
    if (totalPages <= 1) {
      container.style.display = 'none';
      return;
    }
    container.style.display = 'flex';

    const mk = (label, pg, disabled, active) => {
      const b = document.createElement('button');
      b.className = 'spg-btn' + (active ? ' active' : '');
      b.textContent = label;
      b.disabled = !!disabled;
      b.addEventListener('click', () => onPageChange(pg));
      return b;
    };

    container.appendChild(mk('‹', currentPage - 1, currentPage === 1));
    const range = new Set([1, totalPages]);
    for (let i = Math.max(2, currentPage - 2); i <= Math.min(totalPages - 1, currentPage + 2); i++) range.add(i);
    let last = 0;
    [...range].sort((a, b) => a - b).forEach(p => {
      if (p - last > 1) {
        const el = document.createElement('span');
        el.className = 'spg-ellipsis';
        el.textContent = '…';
        container.appendChild(el);
      }
      container.appendChild(mk(p, p, false, p === currentPage));
      last = p;
    });
    container.appendChild(mk('›', currentPage + 1, currentPage === totalPages));

    const info = document.createElement('span');
    info.className = 'spg-info';
    info.textContent = `${(currentPage - 1) * pageSize + 1}–${Math.min(currentPage * pageSize, totalItems)} of ${totalItems.toLocaleString()}`;
    container.appendChild(info);
  }

  function renderFormatPage(issues, isFiltered) {
    const total = issues.length;
    const pageSize = scanResults.format.pageSize;
    const totalPg = Math.max(1, Math.ceil(total / pageSize));
    scanResults.format.page = Math.min(scanResults.format.page, totalPg);
    const start = (scanResults.format.page - 1) * pageSize;
    const pageIssues = issues.slice(start, start + pageSize);

    const filters = [];
    if (langFilter.value !== 'all') filters.push(`<strong>${escapeHtml(langFilter.value)}</strong>`);
    if (scanResults.format.query) filters.push(`"<strong>${escapeHtml(scanResults.format.query)}</strong>"`);
    
    formatSummary.innerHTML = `Found <span class="srch-count">${total.toLocaleString()}</span> issue${total !== 1 ? 's' : ''}${filters.length ? ` for ${filters.join(' and ')}` : ''}`;

    resultsBody.innerHTML = '';
    if (total === 0) {
      resultsBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>No formatting issues!</h3></td></tr>`;
    } else {
      pageIssues.forEach(issue => {
        const tr = document.createElement('tr');
        const rowSpan = issue.rowNum ? ` <span class="row-num" title="Excel Row Number">(Row ${issue.rowNum})</span>` : '';
        tr.innerHTML = `
          <td><strong>${escapeHtml(issue.key)}</strong>${rowSpan}</td>
          <td>${escapeHtml(issue.lang)}</td>
          <td class="issue-cell">${escapeHtml(issue.err)}</td>
          <td><span class="text-snippet">${escapeHtml(issue.snippet)}</span></td>
        `;
        resultsBody.appendChild(tr);
      });
    }

    renderPagination(formatPagination, scanResults.format.page, totalPg, total, pageSize, (pg) => {
      scanResults.format.page = pg;
      renderFormatPage(issues, isFiltered);
      document.getElementById('format-table-wrap').scrollIntoView({ behavior: 'smooth' });
    });
  }

  function renderMissingPage(issues, isFiltered) {
    const total = issues.length;
    const pageSize = scanResults.missing.pageSize;
    const totalPg = Math.max(1, Math.ceil(total / pageSize));
    scanResults.missing.page = Math.min(scanResults.missing.page, totalPg);
    const start = (scanResults.missing.page - 1) * pageSize;
    const pageIssues = issues.slice(start, start + pageSize);

    const filters = [];
    if (langFilter.value !== 'all') filters.push(`<strong>${escapeHtml(langFilter.value)}</strong>`);
    if (scanResults.missing.query) filters.push(`"<strong>${escapeHtml(scanResults.missing.query)}</strong>"`);
    
    missingSummary.innerHTML = `Found <span class="srch-count">${total.toLocaleString()}</span> missing item${total !== 1 ? 's' : ''}${filters.length ? ` for ${filters.join(' and ')}` : ''}`;

    missingBody.innerHTML = '';
    if (total === 0) {
      missingBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>All localizations present!</h3></td></tr>`;
    } else {
      pageIssues.forEach(issue => {
        const tr = document.createElement('tr');
        const langCode = getLangCodeForName(issue.lang);
        const rowSpan = issue.rowNum ? ` <span class="row-num" title="Excel Row Number">(Row ${issue.rowNum})</span>` : '';
        tr.dataset.rowNum = issue.rowNum;
        tr.dataset.colIndex = issue.colIndex != null ? issue.colIndex : '';
        tr.innerHTML = `
          <td><strong>${escapeHtml(issue.key)}</strong>${rowSpan}</td>
          <td>${escapeHtml(issue.lang)}</td>
          <td><span class="text-snippet">${escapeHtml(issue.englishText)}</span></td>
          <td class="inline-trans-cell">
            <button class="btn btn-primary sm-btn inline-translate-btn" data-text="${escapeHtml(issue.englishText)}" data-lang="${langCode}">
              <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
              Translate
            </button>
          </td>
        `;
        missingBody.appendChild(tr);
      });
    }

    renderPagination(missingPagination, scanResults.missing.page, totalPg, total, pageSize, (pg) => {
      scanResults.missing.page = pg;
      renderMissingPage(issues, isFiltered);
      document.getElementById('missing-table-wrap').scrollIntoView({ behavior: 'smooth' });
    });
  }

  function renderResults(scannedCount, formatIssues, missingIssues, isFiltered = false) {
    resultsContainer.classList.remove('hidden');
    container.classList.add('has-results');
    dropzone.classList.add('file-loaded');
    statScanned.textContent = `Rows Scanned: ${scannedCount}`;

    const totalFormat = allFormatIssues.length;
    const totalMissing = allMissingIssues.length;
    const totalIssues = totalFormat + totalMissing;

    badgeFormat.textContent = allFormatIssues.length;
    badgeMissing.textContent = allMissingIssues.length;

    // If it's a fresh scan (not filtered), reset pages
    if (!isFiltered) {
      scanResults.format.page = 1;
      scanResults.missing.page = 1;
    }

    renderFormatPage(formatIssues, isFiltered);
    renderMissingPage(missingIssues, isFiltered);

    if (totalIssues === 0) {
      statIssuesContainer.className = 'stat-pill success';
      statIssues.textContent = 'All Clear! No issues.';
      return;
    }

    statIssuesContainer.className = 'stat-pill danger';
    let issuesText = `Issues Found: ${totalIssues} (${totalFormat} Format, ${totalMissing} Missing)`;
    if (isFiltered) {
      issuesText += ` — Showing ${formatIssues.length + missingIssues.length} for ${langFilter.value}`;
    }
    statIssues.textContent = issuesText;
  }

  function getLangCodeForName(langName) {
    const langMap = {
      'french': 'fr', 'spanish': 'es', 'german': 'de', 'italian': 'it',
      'portuguese': 'pt', 'russian': 'ru', 'japanese': 'ja', 'korean': 'ko',
      'chinese': 'zh-CN', 'arabic': 'ar', 'hindi': 'hi', 'turkish': 'tr',
      'dutch': 'nl', 'polish': 'pl', 'vietnamese': 'vi', 'thai': 'th',
      'indonesian': 'id', 'swedish': 'sv', 'danish': 'da', 'finnish': 'fi'
    };

    let tl = 'auto';
    const langLower = (langName || "").toLowerCase();
    for (const [key, val] of Object.entries(langMap)) {
      if (langLower.includes(key)) {
        tl = val;
        break;
      }
    }
    return tl;
  }

  async function fetchTranslation(text, sourceLang, targetLang) {
    if (!text.trim()) return { text: '', detected: '' };
    const leadMatch = text.match(/^\s*/);
    const trailMatch = text.match(/\s*$/);
    const lead = leadMatch ? leadMatch[0] : '';
    const trail = trailMatch ? trailMatch[0] : '';
    const trimmed = text.slice(lead.length, text.length - trail.length);
    try {
      const isHttp = location.protocol === 'http:' || location.protocol === 'https:';
      const url = isHttp
        ? `/api/translate?sl=${sourceLang}&tl=${targetLang}&q=${encodeURIComponent(trimmed)}`
        : `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sourceLang}&tl=${targetLang}&dt=t&q=${encodeURIComponent(trimmed)}`;
      const response = await fetch(url);
      const data = await response.json();
      const transText = data[0].map(x => x[0]).join('');
      let detected = '';
      if (sourceLang === 'auto' && data[2]) {
        detected = data[2]; // e.g. 'es'
      }
      return { text: lead + transText + trail, detected };
    } catch (e) {
      console.error('Translation failed:', e);
      return { text: 'Error fetching translation. Please try again.', detected: '' };
    }
  }

  function escapeHtml(unsafe) {
    return (unsafe || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }

  /* ============================================================
     SEARCH TAB
     ============================================================ */

  function initSearchTab(rows) {
    if (!rows || rows.length < 1) return;
    let maxLen = 0;
    for (const r of rows) if (Array.isArray(r) && r.length > maxLen) maxLen = r.length;
    const headerRow = rows[0] || [];
    const headers = Array.from({ length: maxLen }, (_, i) => {
      const h = headerRow[i];
      return (h != null && String(h).trim()) ? String(h).trim() : `Col ${i + 1}`;
    });
    srch.allCols = headers;

    const savedColsStr = localStorage.getItem('locaLinterSearchCols');
    if (savedColsStr) {
      try {
        const savedCols = JSON.parse(savedColsStr);
        srch.cols = headers.filter(h => savedCols.includes(h));
        if (srch.cols.length === 0) srch.cols = [...headers];
      } catch (e) {
        srch.cols = [...headers];
      }
    } else {
      srch.cols = [...headers];
    }

    const savedCase = localStorage.getItem('locaLinterSearchCase');
    if (savedCase) {
      sCaseChk.checked = savedCase === 'true';
      srch.caseSensitive = sCaseChk.checked;
    }

    const savedWrap = localStorage.getItem('locaLinterSearchWrap');
    if (savedWrap !== 'false') { // Default to true if not explicitly saved as false
      sWrapChk.checked = true;
      searchTableWrap.classList.add('wrap-text');
    } else {
      sWrapChk.checked = false;
      searchTableWrap.classList.remove('wrap-text');
    }

    srch.rows = buildFlatRows(rows, headers);

    // Restore search query
    const savedQuery = localStorage.getItem('locaLinterSearchQuery');
    if (savedQuery !== null && savedQuery !== undefined) {
      srch.query = savedQuery;
      searchQueryInput.value = savedQuery;
      searchClearX.style.display = srch.query ? 'flex' : 'none';
    } else {
      srch.query = '';
      searchQueryInput.value = '';
      searchClearX.style.display = 'none';
    }

    // Restore match mode
    const savedMode = localStorage.getItem('locaLinterSearchMode');
    if (savedMode) {
      srch.mode = savedMode;
      searchModesEl.querySelectorAll('.smode').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.mode === savedMode);
      });
    }

    srch.page = 1;
    buildColChips(headers);
    renderSearch();
  }

  function buildFlatRows(rawRows, headers) {
    const flat = [];
    for (let i = 1; i < rawRows.length; i++) {
      const r = rawRows[i];
      if (!r || r.every(c => c == null || String(c).trim() === '')) continue;
      const obj = {};
      headers.forEach((h, idx) => { obj[h] = r[idx] != null ? String(r[idx]) : ''; });
      obj.__rowNum = i + 1;
      flat.push(obj);
    }
    return flat;
  }

  function buildColChips(headers) {
    searchColChips.innerHTML = '';
    if (headers.length <= 1) {
      searchColChips.parentElement.style.display = 'none';
      return;
    }
    searchColChips.parentElement.style.display = 'flex';

    const actions = document.createElement('div');
    actions.style.display = 'flex';
    actions.style.gap = '0.5rem';
    actions.style.marginBottom = '0.6rem';

    const selectAll = document.createElement('button');
    selectAll.className = 'smode';
    selectAll.style.padding = '0.15rem 0.5rem';
    selectAll.textContent = 'Select All';

    const clearAll = document.createElement('button');
    clearAll.className = 'smode';
    clearAll.style.padding = '0.15rem 0.5rem';
    clearAll.textContent = 'Clear All';

    actions.appendChild(selectAll);
    actions.appendChild(clearAll);
    searchColChips.appendChild(actions);

    const grid = document.createElement('div');
    grid.className = 'cols-grid';

    const checkboxes = [];

    function updateAll() {
      checkboxes.forEach(cb => {
        cb.checked = srch.cols.includes(cb.dataset.col);
      });
      localStorage.setItem('locaLinterSearchCols', JSON.stringify(srch.cols));
      srch.page = 1;
      renderSearch();
    }

    selectAll.onclick = () => { srch.cols = [...headers]; updateAll(); };
    clearAll.onclick = () => { srch.cols = []; updateAll(); };

    headers.forEach(col => {
      const label = document.createElement('label');
      label.className = 'col-checkbox';

      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.dataset.col = col;
      checkbox.checked = srch.cols.includes(col);

      const text = document.createElement('span');
      text.textContent = col;

      label.appendChild(checkbox);
      label.appendChild(text);

      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          srch.cols.push(col);
        } else {
          if (srch.cols.length === 1) {
            checkbox.checked = true;
            return;
          }
          srch.cols = srch.cols.filter(c => c !== col);
        }
        updateAll();
      });

      checkboxes.push(checkbox);
      grid.appendChild(label);
    });

    searchColChips.appendChild(grid);
  }

  // ── Search matcher ──
  function buildMatcher(query, mode, cs) {
    const q = cs ? query : query.toLowerCase();
    const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    switch (mode) {
      case 'exact': return v => (cs ? v : v.toLowerCase()) === q;
      case 'startsWith': return v => (cs ? v : v.toLowerCase()).startsWith(q);
      case 'endsWith': return v => (cs ? v : v.toLowerCase()).endsWith(q);
      case 'word': try { const wr = new RegExp(`\\b${esc(query)}\\b`, cs ? 'g' : 'gi'); return v => wr.test(v); } catch (_) { return () => false; }
      case 'regex': try { const rr = new RegExp(query, cs ? 'g' : 'gi'); return v => rr.test(v); } catch (_) { return () => false; }
      default: return v => (cs ? v : v.toLowerCase()).includes(q);
    }
  }

  function getSearchResults() {
    if (!srch.query.trim()) return srch.rows;
    const matcher = buildMatcher(srch.query, srch.mode, srch.caseSensitive);
    return srch.rows.filter(row => srch.cols.some(col => matcher(row[col] ?? '')));
  }

  // ── Highlight matches ──
  function hlMatch(text, query, mode, cs) {
    if (!query) return escapeHtml(text);
    const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    try {
      let pat;
      if (mode === 'regex') pat = query;
      else if (mode === 'word') pat = `\\b${esc(query)}\\b`;
      else if (mode === 'startsWith') pat = `^${esc(query)}`;
      else if (mode === 'endsWith') pat = `${esc(query)}$`;
      else if (mode === 'exact') pat = `^${esc(query)}$`;
      else pat = esc(query);
      const re = new RegExp(`(${pat})`, cs ? 'g' : 'gi');
      return text.split(re).map((p, i) =>
        i % 2 === 1 ? `<mark class="srch-hl">${escapeHtml(p)}</mark>` : escapeHtml(p)
      ).join('');
    } catch (_) { return escapeHtml(text); }
  }

  // ── Render search results ──
  function renderSearch() {
    const query = srch.query.trim();
    const matched = getSearchResults();
    const total = matched.length;
    const totalPg = Math.max(1, Math.ceil(total / srch.pageSize));
    srch.page = Math.min(srch.page, totalPg);
    const start = (srch.page - 1) * srch.pageSize;
    const pageRows = matched.slice(start, start + srch.pageSize);

    // Summary
    if (!srch.rows.length) {
      searchSummary.textContent = 'Load a file and type a query to search.';
      searchTableWrap.style.display = 'none';
      searchPagination.innerHTML = '';
      return;
    }
    if (!query) {
      searchSummary.innerHTML = `<span class="srch-count">${srch.rows.length.toLocaleString()}</span> rows loaded — type a query above.`;
      searchTableWrap.style.display = 'none';
      searchPagination.innerHTML = '';
      return;
    }

    searchSummary.innerHTML =
      `<span class="srch-count">${total.toLocaleString()}</span> result${total !== 1 ? 's' : ''} for
       <strong>&ldquo;${escapeHtml(query)}&rdquo;</strong>
       <span class="srch-mode-tag">${srch.mode}</span>
       <span style="color:var(--text-muted)">of ${srch.rows.length.toLocaleString()} rows</span>`;

    // Determine which columns to display (Always show Key [idx 0], plus any selected columns)
    const displayCols = srch.allCols.filter((col, idx) => idx === 0 || srch.cols.includes(col));

    // Build table head
    searchThead.innerHTML = '';
    const hr = document.createElement('tr');
    displayCols.forEach(col => {
      const th = document.createElement('th');
      th.textContent = col;
      hr.appendChild(th);
    });
    searchThead.appendChild(hr);

    // Build table body
    searchTbody.innerHTML = '';
    if (!pageRows.length) {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td colspan="${displayCols.length}" class="success-state"><h3>No matches found</h3><p>Try a different mode or query.</p></td>`;
      searchTbody.appendChild(tr);
    } else {
      const editable = !!currentSheetRef;
      pageRows.forEach(row => {
        const tr = document.createElement('tr');
        if (row.__rowNum) tr.dataset.rowNum = row.__rowNum;
        displayCols.forEach(col => {
          const td = document.createElement('td');
          const val = row[col] ?? '';
          td.innerHTML = srch.cols.includes(col) && query
            ? hlMatch(val, query, srch.mode, srch.caseSensitive)
            : escapeHtml(val);
          td.title = val;
          if (editable) {
            const colIdx = srch.allCols.indexOf(col);
            if (colIdx >= 0) {
              td.contentEditable = 'true';
              td.spellcheck = false;
              td.classList.add('editable-cell');
              td.dataset.colIndex = colIdx;
              td.dataset.original = val;
            }
          }
          tr.appendChild(td);
        });
        searchTbody.appendChild(tr);
      });
    }
    searchTableWrap.style.display = '';

    // Pagination
    renderPagination(searchPagination, srch.page, totalPg, total, srch.pageSize, (pg) => {
      srch.page = pg;
      renderSearch();
      document.getElementById('tab-search').scrollIntoView({ behavior: 'smooth' });
    });
  }


  // ── Search event listeners ──
  let srchDebounce;
  searchQueryInput.addEventListener('input', () => {
    srch.query = searchQueryInput.value;
    localStorage.setItem('locaLinterSearchQuery', srch.query);
    searchClearX.style.display = srch.query ? 'flex' : 'none';
    srch.page = 1;
    clearTimeout(srchDebounce);
    srchDebounce = setTimeout(renderSearch, 150);
  });
  searchQueryInput.addEventListener('keydown', e => {
    if (e.key === 'Escape') { searchQueryInput.value = ''; srch.query = ''; searchClearX.style.display = 'none'; srch.page = 1; renderSearch(); }
  });
  searchClearX.addEventListener('click', () => {
    searchQueryInput.value = ''; srch.query = '';
    localStorage.setItem('locaLinterSearchQuery', '');
    searchClearX.style.display = 'none'; srch.page = 1; renderSearch(); searchQueryInput.focus();
  });

  searchOptionsBtn.addEventListener('click', () => {
    searchOptionsPanel.classList.toggle('hidden');
    if (!searchOptionsPanel.classList.contains('hidden')) {
      searchOptionsBtn.style.color = 'var(--text)';
    } else {
      searchOptionsBtn.style.color = 'var(--muted)';
    }
  });

  let searchDebounce;
  searchModesEl.addEventListener('click', e => {
    const btn = e.target.closest('.smode');
    if (!btn) return;
    searchModesEl.querySelectorAll('.smode').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    srch.mode = btn.dataset.mode;
    localStorage.setItem('locaLinterSearchMode', srch.mode);
    srch.page = 1;
    renderSearch();
  });
  sCaseChk.addEventListener('change', () => {
    srch.caseSensitive = sCaseChk.checked;
    localStorage.setItem('locaLinterSearchCase', sCaseChk.checked);
    srch.page = 1;
    renderSearch();
  });

  // ── Inline edit on Search tab cells ──
  const editUndoStack = []; // { rowNum, colIndex, prev, next }

  async function pushSheetEdit({ rowNum, colIndex, value, prev }) {
    const sourceKey = currentRows && currentRows[rowNum - 1]
      ? ((currentRows[rowNum - 1][0] || '') + '')
      : '';
    await writeSheetCell({
      spreadsheetId: currentSheetRef.spreadsheetId,
      sheetTitle: currentSheetRef.sheetTitle,
      rowNum,
      colIndex,
      value,
    });
    if (currentRows && currentRows[rowNum - 1]) {
      currentRows[rowNum - 1][colIndex] = value;
      saveToDB(currentFileName, currentRows, currentSheetRef);
    }
    editUndoStack.push({ rowNum, colIndex, prev, next: value, sourceKey });
    if (editUndoStack.length > 50) editUndoStack.shift();
    propagateEditToLinkedSheets({ rowNum, colIndex, value, sourceKey }).catch(() => {});
  }

  async function commitSearchCellEdit(td) {
    if (!currentSheetRef) return false;
    const tr = td.closest('tr');
    const rowNum = Number(tr && tr.dataset.rowNum);
    const colIndex = Number(td.dataset.colIndex);
    const original = td.dataset.original ?? '';
    const newValue = td.innerText.replace(/ /g, ' ');
    if (!rowNum || Number.isNaN(colIndex)) return false;
    if (newValue === original) return false;
    td.classList.add('saving');
    try {
      await pushSheetEdit({ rowNum, colIndex, value: newValue, prev: original });
      td.dataset.original = newValue;
      td.title = newValue;
      td.classList.remove('saving');
      td.classList.add('saved');
      setTimeout(() => td.classList.remove('saved'), 1200);
      showToast('Saved. Press Ctrl+Z to undo.');
      return true;
    } catch (err) {
      td.classList.remove('saving');
      td.classList.add('save-error');
      setTimeout(() => td.classList.remove('save-error'), 2000);
      td.innerText = original;
      showToast(err.message || 'Failed to save.', 'error');
      return false;
    }
  }

  searchTbody.addEventListener('keydown', async (e) => {
    const td = e.target.closest('td.editable-cell');
    if (!td) return;
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      const saved = await commitSearchCellEdit(td);
      if (saved) {
        td.dataset.skipRevert = '1';
        td.blur();
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      td.innerText = td.dataset.original ?? '';
      td.dataset.skipRevert = '1';
      td.blur();
    }
  });
  // On blur (without Enter), silently revert — never save by mistake.
  searchTbody.addEventListener('focusout', (e) => {
    const td = e.target.closest && e.target.closest('td.editable-cell');
    if (!td) return;
    if (td.dataset.skipRevert === '1') { delete td.dataset.skipRevert; return; }
    const original = td.dataset.original ?? '';
    if (td.innerText !== original) td.innerText = original;
  });

  // Global undo for sheet edits (Ctrl/Cmd + Z when not focused inside an editable element)
  document.addEventListener('keydown', async (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.shiftKey) return;
    if (e.key !== 'z' && e.key !== 'Z') return;
    const t = e.target;
    if (t && t.matches && (t.matches('input, textarea') || t.isContentEditable)) return;
    if (!editUndoStack.length || !currentSheetRef) return;
    e.preventDefault();
    const last = editUndoStack.pop();
    try {
      await writeSheetCell({
        spreadsheetId: currentSheetRef.spreadsheetId,
        sheetTitle: currentSheetRef.sheetTitle,
        rowNum: last.rowNum,
        colIndex: last.colIndex,
        value: last.prev,
      });
      if (currentRows && currentRows[last.rowNum - 1]) {
        currentRows[last.rowNum - 1][last.colIndex] = last.prev;
        saveToDB(currentFileName, currentRows, currentSheetRef);
      }
      renderSearch();
      showToast('Edit undone.');
      propagateEditToLinkedSheets({
        rowNum: last.rowNum,
        colIndex: last.colIndex,
        value: last.prev,
        sourceKey: last.sourceKey,
      }).catch(() => {});
    } catch (err) {
      editUndoStack.push(last); // restore on failure
      showToast(err.message || 'Undo failed.', 'error');
    }
  });

  function setWrap(enabled) {
    if (globalWrapChk) globalWrapChk.checked = enabled;
    if (sWrapChk) sWrapChk.checked = enabled;
    [formatTableWrap, missingTableWrap, searchTableWrap].forEach(el => {
      if (el) el.classList.toggle('wrap-text', enabled);
    });
    localStorage.setItem('locaLinterGlobalWrap', enabled);
  }

  if (globalWrapChk) globalWrapChk.addEventListener('change', () => setWrap(globalWrapChk.checked));
  if (sWrapChk) sWrapChk.addEventListener('change', () => setWrap(sWrapChk.checked));

  // Initialize wrap state from localStorage (default true if not set)
  const savedWrap = localStorage.getItem('locaLinterGlobalWrap') !== 'false';
  setWrap(savedWrap);

  // Double-click to copy for all result tables
  [resultsBody, missingBody, searchTbody].forEach(tbody => {
    tbody.addEventListener('dblclick', e => {
      const td = e.target.closest('td');
      if (!td) return;
      if (td.classList.contains('editable-cell')) return; // dblclick selects word for editing
      navigator.clipboard.writeText(td.textContent).then(() => {
        showToast('Copied to clipboard!');
      }).catch(() => { });
    });
  });

  /* ============================================================
     QUICK TRANSLATOR WIDGET
     ============================================================ */
  const qtToggle = document.getElementById('qt-toggle');
  const qtPanel = document.getElementById('qt-panel');
  const qtClose = document.getElementById('qt-close');
  const qtInput = document.getElementById('qt-input');
  const qtSourceLang = document.getElementById('qt-source-lang');
  const qtLang = document.getElementById('qt-lang');
  const qtBtn = document.getElementById('qt-translate-btn');
  const qtOutput = document.getElementById('qt-output');
  const qtCopy = document.getElementById('qt-copy');

  // Convert selects to searchable custom dropdowns
  function makeSearchableSelect(selectEl) {
    selectEl.style.display = 'none';

    const wrapper = document.createElement('div');
    wrapper.className = 'cs-wrapper';

    const selectedBox = document.createElement('div');
    selectedBox.className = 'cs-selected';

    const dropdown = document.createElement('div');
    dropdown.className = 'cs-dropdown hidden';

    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'cs-search';
    searchInput.placeholder = 'Search language...';

    const list = document.createElement('ul');
    list.className = 'cs-list';

    const options = Array.from(selectEl.options);

    function renderList(filter = '') {
      list.innerHTML = '';
      const f = filter.toLowerCase();
      options.forEach(opt => {
        if (opt.text.toLowerCase().includes(f) || opt.value.toLowerCase().includes(f)) {
          const li = document.createElement('li');
          li.textContent = opt.text;
          li.dataset.value = opt.value;
          if (opt.value === selectEl.value) {
            li.classList.add('active');
            selectedBox.textContent = opt.text;
          }
          li.addEventListener('click', () => {
            selectEl.value = opt.value;
            selectEl.dispatchEvent(new Event('change'));
            dropdown.classList.add('hidden');
            selectedBox.textContent = opt.text;
            renderList(); // reset filter
            searchInput.value = '';
          });
          list.appendChild(li);
        }
      });
    }

    renderList();

    selectedBox.addEventListener('click', (e) => {
      e.stopPropagation();
      const isHidden = dropdown.classList.contains('hidden');
      document.querySelectorAll('.cs-dropdown').forEach(d => d.classList.add('hidden'));
      if (isHidden) {
        dropdown.classList.remove('hidden');
        searchInput.focus();
        
        // Scroll active item into view
        const activeItem = list.querySelector('.active');
        if (activeItem) {
          activeItem.scrollIntoView({ block: 'nearest' });
        }
      }
    });

    searchInput.addEventListener('input', (e) => renderList(e.target.value));
    searchInput.addEventListener('click', e => e.stopPropagation());

    document.addEventListener('click', (e) => {
      if (!wrapper.contains(e.target)) dropdown.classList.add('hidden');
    });

    // Update custom UI when original select changes via JS
    selectEl.addEventListener('change', () => {
      options.forEach(o => o.selected = (o.value === selectEl.value));
      renderList();
    });

    dropdown.appendChild(searchInput);
    dropdown.appendChild(list);
    wrapper.appendChild(selectedBox);
    wrapper.appendChild(dropdown);
    selectEl.parentNode.insertBefore(wrapper, selectEl.nextSibling);
  }

  // Restore QT languages from localStorage
  const savedSource = localStorage.getItem('locaLinterQTSource');
  const savedTarget = localStorage.getItem('locaLinterQTTarget');
  if (savedSource) qtSourceLang.value = savedSource;
  if (savedTarget) qtLang.value = savedTarget;

  makeSearchableSelect(qtSourceLang);
  makeSearchableSelect(qtLang);

  qtSourceLang.addEventListener('change', () => {
    localStorage.setItem('locaLinterQTSource', qtSourceLang.value);
  });
  qtLang.addEventListener('change', () => {
    localStorage.setItem('locaLinterQTTarget', qtLang.value);
  });

  qtToggle.addEventListener('click', () => {
    qtPanel.classList.toggle('hidden');
    if (!qtPanel.classList.contains('hidden')) {
      qtInput.focus();
    }
  });

  qtClose.addEventListener('click', () => {
    qtPanel.classList.add('hidden');
  });

  async function triggerTranslation() {
    const text = qtInput.value;
    const sourceLang = qtSourceLang.value;
    const targetLang = qtLang.value;
    if (!text.trim()) {
      qtOutput.value = '';
      return;
    }

    qtBtn.textContent = '...';
    qtBtn.disabled = true;
    qtOutput.value = 'Translating...';

    const translation = await fetchTranslation(text, sourceLang, targetLang);

    qtOutput.value = translation.text;
    if (translation.detected && sourceLang === 'auto') {
      const detectedName = qtSourceLang.querySelector(`option[value="${translation.detected}"]`)?.textContent || translation.detected.toUpperCase();
      qtBtn.textContent = `Detected: ${detectedName}`;
    } else {
      qtBtn.textContent = 'Translate';
    }
    qtBtn.disabled = false;
  }

  qtBtn.addEventListener('click', triggerTranslation);

  let translateDebounce;
  qtInput.addEventListener('input', () => {
    qtBtn.textContent = 'Translate';
    clearTimeout(translateDebounce);
    translateDebounce = setTimeout(triggerTranslation, 600);
  });

  // Also translate automatically when language changes
  qtSourceLang.addEventListener('change', () => {
    if (qtInput.value.trim()) triggerTranslation();
  });
  qtLang.addEventListener('change', () => {
    if (qtInput.value.trim()) triggerTranslation();
  });

  qtInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
      triggerTranslation();
    }
  });

  qtCopy.addEventListener('click', () => {
    if (!qtOutput.value.trim()) return;
    navigator.clipboard.writeText(qtOutput.value).then(() => {
      qtCopy.classList.add('copied');
      setTimeout(() => qtCopy.classList.remove('copied'), 2000);
      showToast('Translation copied!');
    });
  });

});
