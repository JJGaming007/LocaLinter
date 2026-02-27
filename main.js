window.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const browseBtn = document.getElementById('browse-btn');
  const resultsContainer = document.getElementById('results-container');
  const resultsBody = document.getElementById('results-body');
  const missingBody = document.getElementById('missing-body');
  const statScanned = document.getElementById('stat-scanned');
  const statIssues = document.getElementById('stat-issues');
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

  saveIgnoreBtn.addEventListener('click', () => {
    localStorage.setItem('localinter_ignore_list', ignoreListInput.value);
    const successMsg = document.createElement('span');
    successMsg.textContent = ' Saved!';
    successMsg.style.color = 'var(--success)';
    successMsg.style.marginLeft = '10px';
    saveIgnoreBtn.parentNode.appendChild(successMsg);
    setTimeout(() => successMsg.remove(), 2000);

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
    const varRegex = /{[0-9]+}/g;
    return text.match(varRegex) || [];
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
          formatIssues.push({ key: keyInfo, lang: headers[englishColIndex] || "English", err: `Base text err: ${baseBracketErr}`, snippet: englishText });
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
        const targetText = row[col] ? String(row[col]) : "";

        if (!targetText) {
          if (englishText) {
            missingIssues.push({
              key: keyInfo,
              lang: headers[col] || `Col ${col}`,
              englishText: englishText
            });
          }
          continue;
        }

        // 1. Check Brackets
        let bracketErr = checkBrackets(targetText);
        if (bracketErr) {
          formatIssues.push({
            key: keyInfo,
            lang: headers[col] || `Col ${col}`,
            err: bracketErr,
            snippet: targetText
          });
        }

        // 2. Check Variables match
        if (englishText) {
          const targetVars = extractVars(targetText);
          const missing = englishVars.filter(v => !targetVars.includes(v));
          const extra = targetVars.filter(v => !englishVars.includes(v));

          if (missing.length > 0 || extra.length > 0) {
            let varErrs = [];
            if (missing.length > 0) varErrs.push(`Missing vars: ${missing.join(', ')}`);
            if (extra.length > 0) varErrs.push(`Extra vars: ${extra.join(', ')}`);

            formatIssues.push({
              key: keyInfo,
              lang: headers[col] || `Col ${col}`,
              err: varErrs.join('; '),
              snippet: targetText
            });
          }
        }
      }
    }

    renderResults(rowsScanned, formatIssues, missingIssues);
  }

  function renderResults(scannedCount, formatIssues, missingIssues) {
    resultsContainer.classList.remove('hidden');
    statScanned.textContent = `Rows Scanned: ${scannedCount}`;

    const totalIssues = formatIssues.length + missingIssues.length;
    badgeFormat.textContent = formatIssues.length;
    badgeMissing.textContent = missingIssues.length;

    if (totalIssues === 0) {
      statIssues.className = 'stat-pill success';
      statIssues.textContent = 'All Clear! No issues.';
      resultsBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>Everything looks perfect!</h3><p>No formatting errors were found in this sheet.</p></td></tr>`;
      missingBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>All localizations present!</h3><p>No missing translations found.</p></td></tr>`;
      return;
    }

    statIssues.className = totalIssues > 0 ? 'stat-pill danger' : 'stat-pill success';
    statIssues.textContent = `Issues Found: ${totalIssues} (${formatIssues.length} Format, ${missingIssues.length} Missing)`;

    resultsBody.innerHTML = '';
    if (formatIssues.length === 0) {
      resultsBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>No formatting issues!</h3></td></tr>`;
    } else {
      formatIssues.forEach(issue => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong>${escapeHtml(issue.key)}</strong></td>
          <td>${escapeHtml(issue.lang)}</td>
          <td class="issue-cell">${escapeHtml(issue.err)}</td>
          <td><span class="text-snippet">${escapeHtml(issue.snippet)}</span></td>
        `;
        resultsBody.appendChild(tr);
      });
    }

    missingBody.innerHTML = '';
    if (missingIssues.length === 0) {
      missingBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>All localizations present!</h3></td></tr>`;
    } else {
      missingIssues.forEach(issue => {
        const tr = document.createElement('tr');
        const translateUrl = getGoogleTranslateUrl(issue.lang, issue.englishText);
        tr.innerHTML = `
          <td><strong>${escapeHtml(issue.key)}</strong></td>
          <td>${escapeHtml(issue.lang)}</td>
          <td><span class="text-snippet">${escapeHtml(issue.englishText)}</span></td>
          <td>
            <a href="${translateUrl}" target="_blank" class="translate-btn">
              <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"></path><polyline points="15 3 21 3 21 9"></polyline><line x1="10" y1="14" x2="21" y2="3"></line></svg>
              Translate
            </a>
          </td>
        `;
        missingBody.appendChild(tr);
      });
    }
  }

  function getGoogleTranslateUrl(langName, text) {
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

    return `https://translate.google.com/?sl=en&tl=${tl}&text=${encodeURIComponent(text)}&op=translate`;
  }

  function escapeHtml(unsafe) {
    return (unsafe || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }
});
