window.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const browseBtn = document.getElementById('browse-btn');
  const resultsContainer = document.getElementById('results-container');
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
  const searchQueryInput  = document.getElementById('search-query');
  const searchClearX      = document.getElementById('search-clear-x');
  const searchOptionsBtn  = document.getElementById('search-options-btn');
  const searchOptionsPanel= document.getElementById('search-options-panel');
  const searchModesEl     = document.getElementById('searchModes');
  const sCaseChk          = document.getElementById('s-case');
  const sWrapChk          = document.getElementById('s-wrap');
  const searchColChips    = document.getElementById('searchColChips');
  const searchSummary     = document.getElementById('searchSummary');
  const searchTableWrap   = document.getElementById('searchTableWrap');
  const searchThead       = document.getElementById('search-thead');
  const searchTbody       = document.getElementById('search-tbody');
  const searchPagination  = document.getElementById('searchPagination');

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
      td.innerHTML = `
        <div style="display: flex; gap: 0.5rem; align-items: center;">
          <input type="text" class="text-snippet" value="${escapeHtml(translation).replace(/"/g, '&quot;')}" style="background: rgba(63,185,80,0.1); border: 1px solid rgba(63,185,80,0.3); color: var(--success); width: 100%; padding: 0.2rem 0.5rem; outline: none; border-radius: 4px; font-family: var(--font);" />
          <button class="qt-copy-btn inline-copy-btn" title="Copy" style="position: static; flex-shrink: 0;">
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
          </button>
        </div>
      `;
    } else {
      btn.textContent = 'Failed. Retry?';
      btn.disabled = false;
    }
  });

  let currentRows = null;
  let currentFileName = "";
  const DB_NAME = 'LocaLinterDB';
  const STORE_NAME = 'cache';

  function saveToDB(fileName, rows) {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = (e) => e.target.result.createObjectStore(STORE_NAME);
    request.onsuccess = (e) => e.target.result.transaction(STORE_NAME, 'readwrite').objectStore(STORE_NAME).put({ fileName, rows }, 'currentFile');
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
    if (currentRows) {
      validateData(currentRows);
    }
  });

  removeFileBtn.addEventListener('click', () => {
    clearDB();
    currentRows = null;
    currentFileName = "";
    dropzone.classList.remove('file-loaded');
    dropContentIdle.classList.remove('hidden');
    loadedContent.classList.add('hidden');
    resultsContainer.classList.add('hidden');
    fileInput.value = '';
  });

  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      tabBtns.forEach(b => b.classList.remove('active'));
      tabContents.forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
    });
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
  }

  // Initial load check
  loadFromDB((data) => {
    currentFileName = data.fileName;
    currentRows = data.rows;
    showLoadedState(currentFileName);
    validateData(currentRows);
    initSearchTab(currentRows);
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
    const filteredFormat = selectedLang === 'all'
      ? allFormatIssues
      : allFormatIssues.filter(i => i.lang === selectedLang);

    const filteredMissing = selectedLang === 'all'
      ? allMissingIssues
      : allMissingIssues.filter(i => i.lang === selectedLang);

    renderResults(currentScannedCount, filteredFormat, filteredMissing, true);
  }

  function renderResults(scannedCount, formatIssues, missingIssues, isFiltered = false) {
    resultsContainer.classList.remove('hidden');
    statScanned.textContent = `Rows Scanned: ${scannedCount}`;

    const totalFormat = allFormatIssues.length;
    const totalMissing = allMissingIssues.length;
    const totalIssues = totalFormat + totalMissing;

    badgeFormat.textContent = formatIssues.length;
    badgeMissing.textContent = missingIssues.length;

    if (totalIssues === 0) {
      statIssuesContainer.className = 'stat-pill success';
      statIssues.textContent = 'All Clear! No issues.';
      resultsBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>Everything looks perfect!</h3><p>No formatting errors were found in this sheet.</p></td></tr>`;
      missingBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>All localizations present!</h3><p>No missing translations found.</p></td></tr>`;
      return;
    }

    statIssuesContainer.className = 'stat-pill danger';
    let issuesText = `Issues Found: ${totalIssues} (${totalFormat} Format, ${totalMissing} Missing)`;
    if (isFiltered) {
      issuesText += ` — Showing ${formatIssues.length + missingIssues.length} for ${langFilter.value}`;
    }
    statIssues.textContent = issuesText;

    resultsBody.innerHTML = '';
    if (formatIssues.length === 0) {
      resultsBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>No formatting issues!</h3>${isFiltered ? `<p>For selected language: ${langFilter.value}</p>` : ''}</td></tr>`;
    } else {
      formatIssues.forEach(issue => {
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

    missingBody.innerHTML = '';
    if (missingIssues.length === 0) {
      missingBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>All localizations present!</h3>${isFiltered ? `<p>For selected language: ${langFilter.value}</p>` : ''}</td></tr>`;
    } else {
      missingIssues.forEach(issue => {
        const tr = document.createElement('tr');
        const langCode = getLangCodeForName(issue.lang);
        const rowSpan = issue.rowNum ? ` <span class="row-num" title="Excel Row Number">(Row ${issue.rowNum})</span>` : '';
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
    try {
      const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${sourceLang}&tl=${targetLang}&dt=t&q=${encodeURIComponent(text)}`;
      const response = await fetch(url);
      const data = await response.json();
      const transText = data[0].map(x => x[0]).join('');
      let detected = '';
      if (sourceLang === 'auto' && data[2]) {
        detected = data[2]; // e.g. 'es'
      }
      return { text: transText, detected };
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
    const headers = rows[0].map((h, i) => (h != null && String(h).trim()) ? String(h).trim() : `Col ${i + 1}`);
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

    const savedWrap = localStorage.getItem('locaLinterSearchWrap');
    if (savedWrap !== 'false') { // Default to true if not explicitly saved as false
      sWrapChk.checked = true;
      searchTableWrap.classList.add('wrap-text');
    } else {
      sWrapChk.checked = false;
      searchTableWrap.classList.remove('wrap-text');
    }

    srch.rows    = buildFlatRows(rows, headers);
    srch.query   = '';
    srch.page    = 1;
    searchQueryInput.value = '';
    searchClearX.style.display = 'none';
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
      case 'exact':      return v => (cs ? v : v.toLowerCase()) === q;
      case 'startsWith': return v => (cs ? v : v.toLowerCase()).startsWith(q);
      case 'endsWith':   return v => (cs ? v : v.toLowerCase()).endsWith(q);
      case 'word': try { const wr = new RegExp(`\\b${esc(query)}\\b`, cs ? 'g' : 'gi'); return v => wr.test(v); } catch(_){ return ()=>false; }
      case 'regex': try { const rr = new RegExp(query, cs ? 'g' : 'gi'); return v => rr.test(v); } catch(_){ return ()=>false; }
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
      if (mode === 'regex')      pat = query;
      else if (mode === 'word')  pat = `\\b${esc(query)}\\b`;
      else if (mode === 'startsWith') pat = `^${esc(query)}`;
      else if (mode === 'endsWith')   pat = `${esc(query)}$`;
      else if (mode === 'exact')      pat = `^${esc(query)}$`;
      else                            pat = esc(query);
      const re = new RegExp(`(${pat})`, cs ? 'g' : 'gi');
      return text.split(re).map((p, i) =>
        i % 2 === 1 ? `<mark class="srch-hl">${escapeHtml(p)}</mark>` : escapeHtml(p)
      ).join('');
    } catch(_) { return escapeHtml(text); }
  }

  // ── Render search results ──
  function renderSearch() {
    const query   = srch.query.trim();
    const matched = getSearchResults();
    const total   = matched.length;
    const totalPg = Math.max(1, Math.ceil(total / srch.pageSize));
    srch.page     = Math.min(srch.page, totalPg);
    const start   = (srch.page - 1) * srch.pageSize;
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
      pageRows.forEach(row => {
        const tr = document.createElement('tr');
        displayCols.forEach(col => {
          const td = document.createElement('td');
          const val = row[col] ?? '';
          td.innerHTML = srch.cols.includes(col) && query
            ? hlMatch(val, query, srch.mode, srch.caseSensitive)
            : escapeHtml(val);
          td.title = val;
          tr.appendChild(td);
        });
        searchTbody.appendChild(tr);
      });
    }
    searchTableWrap.style.display = '';

    // Pagination
    renderSearchPagination(totalPg, total);
  }

  function renderSearchPagination(totalPg, total) {
    searchPagination.innerHTML = '';
    if (totalPg <= 1) return;
    const cur = srch.page;
    const mk = (label, pg, disabled, active) => {
      const b = document.createElement('button');
      b.className = 'spg-btn' + (active ? ' active' : '');
      b.textContent = label;
      b.disabled = !!disabled;
      b.addEventListener('click', () => { srch.page = pg; renderSearch(); document.getElementById('tab-search').scrollIntoView({ behavior: 'smooth' }); });
      return b;
    };
    searchPagination.appendChild(mk('‹', cur - 1, cur === 1));
    const range = new Set([1, totalPg]);
    for (let i = Math.max(2, cur - 2); i <= Math.min(totalPg - 1, cur + 2); i++) range.add(i);
    let last = 0;
    [...range].sort((a,b) => a-b).forEach(p => {
      if (p - last > 1) { const el = document.createElement('span'); el.className = 'spg-ellipsis'; el.textContent = '…'; searchPagination.appendChild(el); }
      searchPagination.appendChild(mk(p, p, false, p === cur));
      last = p;
    });
    searchPagination.appendChild(mk('›', cur + 1, cur === totalPg));
    const info = document.createElement('span');
    info.className = 'spg-info';
    info.textContent = `${(cur-1)*srch.pageSize+1}–${Math.min(cur*srch.pageSize, total)} of ${total.toLocaleString()}`;
    searchPagination.appendChild(info);
  }

  // ── Search event listeners ──
  let srchDebounce;
  searchQueryInput.addEventListener('input', () => {
    srch.query = searchQueryInput.value;
    searchClearX.style.display = srch.query ? 'flex' : 'none';
    srch.page = 1;
    clearTimeout(srchDebounce);
    srchDebounce = setTimeout(renderSearch, 150);
  });
  searchQueryInput.addEventListener('keydown', e => {
    if (e.key === 'Escape') { searchQueryInput.value = ''; srch.query = ''; searchClearX.style.display = 'none'; srch.page = 1; renderSearch(); }
  });
  searchClearX.addEventListener('click', () => {
    searchQueryInput.value = ''; srch.query = ''; searchClearX.style.display = 'none'; srch.page = 1; renderSearch(); searchQueryInput.focus();
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
    srch.page = 1;
    renderSearch();
  });
  sCaseChk.addEventListener('change', () => { srch.caseSensitive = sCaseChk.checked; srch.page = 1; renderSearch(); });
  
  sWrapChk.addEventListener('change', () => {
    localStorage.setItem('locaLinterSearchWrap', sWrapChk.checked);
    if (sWrapChk.checked) {
      searchTableWrap.classList.add('wrap-text');
    } else {
      searchTableWrap.classList.remove('wrap-text');
    }
  });

  // Double-click a search result cell to copy
  searchTbody.addEventListener('dblclick', e => {
    const td = e.target.closest('td');
    if (!td) return;
    navigator.clipboard.writeText(td.textContent).catch(() => {});
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
          if (opt.selected) {
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

  makeSearchableSelect(qtSourceLang);
  makeSearchableSelect(qtLang);

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
