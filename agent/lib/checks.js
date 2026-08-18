'use strict';

const { norm, skeleton, extractPlaceholders, quotedUiReferences } = require('./sheet');

/**
 * Deterministic, zero-cost checks that run on every captured screen before the
 * Claude vision pass. These catch the mechanical failures exactly (truncation,
 * overflow, wrong-language strings, broken placeholders, reversed RTL) so the
 * model can spend its attention on what only vision can see.
 *
 * They require the Unity bridge (exact strings + rects). In screenshot-only
 * mode this module returns [] and the vision pass carries the whole load.
 */

const TOFU_RE = /[�□-]/;
const ELLIPSIS_RE = /(…|\.\.\.)\s*$/;
const VISIBLE_MARKUP_RE = /(\{\d+\}|\{[a-z_][a-z0-9_]*\}|%[sdif@]|%\d+\$[sdif@]|<(color|size|b|i|sprite|font|align)\b[^>]*>|\$\{[^}]+\})/i;

const SEV = { high: 3, medium: 2, low: 1 };

function reverseString(s) {
  return Array.from(String(s)).reverse().join('');
}

function rectsOverlap(a, b) {
  const x = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
  const y = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
  const inter = x * y;
  if (inter <= 0) return 0;
  const minArea = Math.max(1, Math.min(a.w * a.h, b.w * b.h));
  return inter / minArea;
}

function hasStrongScript(text) {
  return /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Thai}\p{Script=Devanagari}\p{Script=Cyrillic}]/u.test(text);
}

/**
 * @param {object} state    bridge snapshot { screen, texts, scene, locale }
 * @param {SheetIndex} sheet
 * @param {object} opts     { targetHeader, englishHeader, rtl }
 * @returns {Array} issues
 */
function runChecks(state, sheet, opts = {}) {
  const issues = [];
  if (!state || !Array.isArray(state.texts)) return issues;

  const targetHeader = opts.targetHeader || null;
  const englishHeader = opts.englishHeader || (sheet.englishCol && sheet.englishCol.header) || null;
  const rtl = !!opts.rtl;
  const screen = state.screen || { width: 0, height: 0 };

  const add = (t, severity, text, message, extra = {}) => {
    issues.push({
      source: 'static',
      type: t,
      severity,
      element: text.path || text.id || '',
      elementId: text.id || '',
      text: text.text || '',
      rect: text.rect || null,
      message,
      ...extra,
    });
  };

  // Strings sourced from the screenshot (vision mode) have no rect — geometric
  // checks are skipped for those, sheet checks still apply.
  const visible = state.texts.filter((t) => t && t.active !== false);

  /**
   * A description that names another control must name it as that control is
   * actually labelled, in this same language.
   *
   * Confirmed by hand on Indus: an avatar's lore reads `Click "Customize"`. If
   * the lore is translated but the button is not — or the button is renamed and
   * the lore is not — the instruction sends the player hunting for something
   * that does not exist under that name. It only shows up when the quoted
   * fragment is compared against the labels on the very same screen, which is
   * why it runs here rather than per-string.
   */
  const onScreenLabels = new Set(
    visible.map((t) => norm(t.text)).filter((v) => v && v.split(' ').length <= 4)
  );
  for (const t of visible) {
    for (const quoted of quotedUiReferences(t.text)) {
      const wanted = norm(quoted);
      if (!wanted || onScreenLabels.has(wanted)) continue;
      // Only complain when the capture actually read the screen. A near-empty
      // text list means we saw nothing, not that the reference is broken —
      // but judge that on everything captured, not just the short labels,
      // since a screen can legitimately hold one button and a paragraph.
      if (visible.length < 2 || !onScreenLabels.size) continue;
      add('ui_reference_mismatch', 'medium', t,
        `This text tells the player to use "${quoted}", but no control on this screen is labelled that. ` +
        'Either the instruction or the control was translated without the other.',
        { quotedReference: quoted });
    }
  }

  for (const t of visible) {
    const raw = String(t.text == null ? '' : t.text);
    const rect = t.rect || null;
    const area = rect ? (rect.w || 0) * (rect.h || 0) : 0;

    // ── blank labels that still occupy space ──────────────────────────────
    if (!raw.trim()) {
      if (area > 900) {
        add('empty_label', 'medium', t, 'Text element is visible and sized but renders nothing — a missing localization key would look like this.');
      }
      continue;
    }

    // ── encoding / font fallback ──────────────────────────────────────────
    if (TOFU_RE.test(raw)) {
      add('mojibake', 'high', t, 'String contains replacement or tofu glyphs — the font is missing characters for this language.');
    }

    // ── raw markup or placeholders leaking to the player ─────────────────
    const leak = VISIBLE_MARKUP_RE.exec(raw);
    if (leak) {
      add('unresolved_placeholder', 'high', t, `Unsubstituted token "${leak[0]}" is rendered as literal text.`);
    }

    // ── truncation / overflow (engine-reported, exact) ────────────────────
    if (t.isTruncated) {
      add('truncated', 'high', t, 'The engine reports this text as truncated — the string does not fit its box.');
    } else if (ELLIPSIS_RE.test(raw.trim()) && raw.trim().length > 4) {
      add('truncated', 'medium', t, 'String ends in an ellipsis, which usually means it was cut off.');
    }
    if (rect && typeof t.preferredWidth === 'number' && rect.w > 0 && t.preferredWidth > rect.w + 2) {
      const over = Math.round(((t.preferredWidth - rect.w) / rect.w) * 100);
      add('overflow_horizontal', over > 8 ? 'high' : 'medium', t,
        `Text needs ${Math.round(t.preferredWidth)}px but its box is ${Math.round(rect.w)}px wide (${over}% over).`);
    }
    if (rect && typeof t.preferredHeight === 'number' && rect.h > 0 && t.preferredHeight > rect.h + 2) {
      add('overflow_vertical', 'medium', t,
        `Text needs ${Math.round(t.preferredHeight)}px of height but its box is ${Math.round(rect.h)}px tall.`);
    }

    // ── pushed off the screen ─────────────────────────────────────────────
    if (rect && screen.width && screen.height) {
      const offLeft = rect.x < -2;
      const offRight = rect.x + rect.w > screen.width + 2;
      const offTop = rect.y < -2;
      const offBottom = rect.y + rect.h > screen.height + 2;
      if (offLeft || offRight || offTop || offBottom) {
        add('offscreen', 'high', t,
          `Text extends past the screen edge (${offLeft ? 'left ' : ''}${offRight ? 'right ' : ''}${offTop ? 'top ' : ''}${offBottom ? 'bottom' : ''}).`.trim());
      }
    }

    // ── sheet comparison ──────────────────────────────────────────────────
    //
    // This mirrors how the check is done by hand:
    //
    //   1. Find the on-screen string in the sheet.
    //   2. Found in the column under test  → correct, say nothing.
    //   3. Found only in the source column → look at what the target column
    //      holds for that same key:
    //        · a real translation      → REPORT: the build is ignoring it
    //        · the English text again  → leave it; the build matches the sheet
    //        · nothing at all          → leave it; there is nothing to show
    //   4. Not found anywhere          → REPORT: probably hardcoded, never
    //                                    sent for translation
    //
    // The distinction that matters: a missing translation is a *sheet* problem
    // and belongs in Missing locales, while a translation the build fails to
    // use is a *build* problem and belongs here.
    const exact = sheet.lookupExact(raw).filter((h) => h.header !== '__key__');
    const exactHeaders = new Set(exact.map((h) => h.header));

    if (exact.length) {
      const inTarget = targetHeader ? exactHeaders.has(targetHeader) : true;
      if (!inTarget) {
        const isEnglish = englishHeader && exactHeaders.has(englishHeader);
        // lookupExact matches by value across the whole sheet, so these hits
        // can be different rows. Report against the row that actually matched
        // the source column — taking hits[0] blindly attributed the finding to
        // an unrelated key and quoted that key's translation as "expected".
        const hit = (isEnglish && exact.find((h) => h.header === englishHeader)) || exact[0];
        const entry = hit.entry;
        const expected = targetHeader ? entry.values[targetHeader] : '';
        // A translation identical to the source is not an untranslated string:
        // "OK", "Menu" and most proper nouns are legitimately the same in both.
        const sameAsSource = englishHeader && expected &&
          norm(expected) === norm(entry.values[englishHeader] || '');
        if (isEnglish && expected && expected.trim() && !sameAsSource) {
          // The sheet has a real translation and the build is not using it.
          // This is the finding worth a tester's time.
          add('untranslated', 'high', t,
            `Source-language string is showing while "${targetHeader}" has a translation.`,
            { key: entry.key, row: entry.rowNumber, expected, matchedLanguage: englishHeader });
        } else if (isEnglish && sameAsSource) {
          // The target column holds the English text too. The build is showing
          // exactly what the sheet says to show, so there is no build defect
          // here — at most a gap in the sheet, which Missing locales covers.
        } else if (isEnglish) {
          // No target value at all: again, the build has nothing else it could
          // display. Kept at low severity as a pointer, not raised as a defect.
          add('missing_translation', 'low', t,
            `The sheet has no "${targetHeader}" value for this key, so the build can only show the source string.`,
            { key: entry.key, row: entry.rowNumber, matchedLanguage: englishHeader });
        } else {
          add('wrong_language', 'high', t,
            `String belongs to "${hit.header}", not the language under test ("${targetHeader}").`,
            { key: entry.key, row: entry.rowNumber, expected, matchedLanguage: hit.header });
        }
      } else if (targetHeader) {
        // matched the right column — verify the placeholder set survived
        const entry = exact.find((h) => h.header === targetHeader).entry;
        const src = englishHeader ? entry.values[englishHeader] : '';
        if (src) {
          const a = extractPlaceholders(src).sort().join('|');
          const b = extractPlaceholders(entry.values[targetHeader]).sort().join('|');
          if (a !== b) {
            add('placeholder_mismatch', 'high', t,
              `Placeholders differ from the source string (source: ${a || 'none'} / translation: ${b || 'none'}).`,
              { key: entry.key, row: entry.rowNumber, expected: entry.values[targetHeader] });
          }
        }
      }
    } else {
      // ── reversed RTL: the classic "Arabic renders backwards" bug ────────
      if (rtl) {
        const rev = sheet.lookupExact(reverseString(raw)).filter((h) => h.header !== '__key__');
        if (rev.length) {
          const hit = rev[0];
          add('rtl_reversed', 'high', t,
            'String matches the sheet only when reversed — it is being rendered in visual instead of logical order.',
            { key: hit.entry.key, row: hit.entry.rowNumber, expected: hit.entry.values[targetHeader] || '', matchedLanguage: hit.header });
          continue;
        }
      }

      const fuzzy = sheet.lookupFuzzy(raw, { limit: 3 });
      const best = fuzzy[0];
      if (best && best.score >= 0.82 && targetHeader && best.header === targetHeader) {
        add('text_mismatch', 'medium', t,
          `Rendered text differs from the sheet value for this key (${Math.round(best.score * 100)}% similar) — the build may be stale or the string edited in code.`,
          { key: best.entry.key, row: best.entry.rowNumber, expected: best.entry.values[targetHeader] });
      } else if (best && best.score >= 0.82 && targetHeader && best.header !== targetHeader) {
        add('wrong_language', 'medium', t,
          `Closest sheet match is in "${best.header}", not "${targetHeader}".`,
          { key: best.entry.key, row: best.entry.rowNumber, expected: best.entry.values[targetHeader] || '', matchedLanguage: best.header });
      } else if (!fuzzy.length && raw.trim().length >= 3 && /\p{L}/u.test(raw) && !/^[\p{N}\p{P}\p{S}\s]+$/u.test(raw)) {
        add('not_in_sheet', 'medium', t,
          'No key or source string anywhere in the sheet matches this text — it was probably hardcoded and never sent for translation.');
      }

      // Target language is non-Latin but the rendered string is plain ASCII words
      if (targetHeader && !hasStrongScript(raw) && opts.expectsNonLatin && /[A-Za-z]{4,}/.test(raw)) {
        add('probably_untranslated', 'medium', t,
          `Latin text on a screen under test in "${targetHeader}", which uses a non-Latin script.`);
      }
    }
  }

  // ── overlapping labels ──────────────────────────────────────────────────
  for (let i = 0; i < visible.length; i++) {
    for (let j = i + 1; j < visible.length; j++) {
      const a = visible[i];
      const b = visible[j];
      if (!a.rect || !b.rect) continue;
      if (!a.text.trim() || !b.text.trim()) continue;
      const ratio = rectsOverlap(a.rect, b.rect);
      if (ratio > 0.35) {
        issues.push({
          source: 'static',
          type: 'overlap',
          severity: ratio > 0.7 ? 'high' : 'medium',
          element: `${a.path} ⨯ ${b.path}`,
          elementId: a.id,
          text: `${a.text} / ${b.text}`,
          rect: a.rect,
          message: `Two text elements overlap by ${Math.round(ratio * 100)}% — the longer translation is colliding with a neighbour.`,
        });
      }
    }
  }

  return issues;
}

module.exports = { runChecks, SEV };
