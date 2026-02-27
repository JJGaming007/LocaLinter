window.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('file-input');
  const browseBtn = document.getElementById('browse-btn');
  const resultsContainer = document.getElementById('results-container');
  const resultsBody = document.getElementById('results-body');
  const statScanned = document.getElementById('stat-scanned');
  const statIssues = document.getElementById('stat-issues');

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
      validateData(json);
    };
    reader.readAsArrayBuffer(file);
  }

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

    const issues = [];
    let rowsScanned = rows.length - 1;

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      const row = rows[rowIndex];
      if (!row || row.length === 0) continue;

      const keyInfo = row[0] || `Row ${rowIndex + 1}`;
      const englishText = row[englishColIndex] ? String(row[englishColIndex]) : "";
      const englishVars = extractVars(englishText);

      // Validate base language first (English)
      if (englishText) {
        let baseBracketErr = checkBrackets(englishText);
        if (baseBracketErr) {
          issues.push({ key: keyInfo, lang: headers[englishColIndex] || "English", err: `Base text err: ${baseBracketErr}`, snippet: englishText });
        }
      }

      for (let col = 1; col < headers.length; col++) {
        if (col === englishColIndex) continue;
        const targetText = row[col] ? String(row[col]) : "";
        if (!targetText) continue;

        // 1. Check Brackets
        let bracketErr = checkBrackets(targetText);
        if (bracketErr) {
          issues.push({
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

            issues.push({
              key: keyInfo,
              lang: headers[col] || `Col ${col}`,
              err: varErrs.join('; '),
              snippet: targetText
            });
          }
        }
      }
    }

    renderResults(rowsScanned, issues);
  }

  function renderResults(scannedCount, issues) {
    resultsContainer.classList.remove('hidden');
    statScanned.textContent = `Rows Scanned: ${scannedCount}`;

    if (issues.length === 0) {
      statIssues.className = 'stat-pill success';
      statIssues.textContent = 'All Clear! No issues.';
      resultsBody.innerHTML = `<tr><td colspan="4" class="success-state"><h3>Everything looks perfect!</h3><p>No formatting errors were found in this sheet.</p></td></tr>`;
      return;
    }

    statIssues.className = 'stat-pill danger';
    statIssues.textContent = `Issues Found: ${issues.length}`;

    resultsBody.innerHTML = '';
    issues.forEach(issue => {
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

  function escapeHtml(unsafe) {
    return (unsafe || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }
});
