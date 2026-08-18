'use strict';

/**
 * Screens remembered in the source language, so a later run in another
 * language can be compared against them.
 *
 * An overflow is very hard to report honestly from one screenshot. "This label
 * is too long" is an opinion; the reviewer has no way to know whether the box
 * was always that tight or whether the translation broke it. The pair settles
 * it — English "Rewards" sitting comfortably on one line beside Portuguese
 * "Recompen/sas" split across two with no hyphen is not a matter of taste.
 *
 * It matters just as much for what it disproves. Three findings that looked
 * solid on the Indus scan died against the baseline: a counter running under
 * some artwork, a blank row in a dropdown and a screen called by three
 * different names all turned out to do exactly the same thing in English, so
 * none of them belonged to the translator. A finding filed against a
 * translation for a defect the source already has sends the wrong team after
 * it, and costs more than the finding was worth.
 *
 * A run only ever covers one language, so the pair cannot be gathered in a
 * single pass: the source-language captures are written to disk keyed by route
 * and screen, and read back when a later run recognises the same screen.
 */

const fs = require('fs');
const path = require('path');

const { DATA_DIR } = require('./paths');

const BASELINE_DIR = path.join(DATA_DIR, 'baselines');

function slug(s) {
  return String(s || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'unknown';
}

function dirFor(routeName, language) {
  return path.join(BASELINE_DIR, slug(routeName), slug(language));
}

/**
 * Everything recorded for one route in one language.
 * Missing or unreadable is not an error — it just means no pair is available.
 */
function load(routeName, language) {
  const file = path.join(dirFor(routeName, language), 'index.json');
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return { language, screens: {} };
  }
}

/**
 * Remember one screen. Keyed by its name in the route map rather than by the
 * run's own screen id, because those ids depend on the order a crawl happened
 * to take and will not line up between two runs.
 */
function record(routeName, language, screenName, { texts, png }) {
  if (!routeName || !language || !screenName) return;
  const dir = dirFor(routeName, language);
  fs.mkdirSync(dir, { recursive: true });

  const index = load(routeName, language);
  const file = `${slug(screenName)}.png`;
  if (png) {
    try { fs.writeFileSync(path.join(dir, file), png); } catch { /* the texts alone are still useful */ }
  }
  index.language = language;
  index.screens[screenName] = {
    texts: (texts || []).filter(Boolean).slice(0, 200),
    file,
    at: Date.now(),
  };
  try {
    fs.writeFileSync(path.join(dir, 'index.json'), JSON.stringify(index, null, 2), 'utf8');
  } catch { /* a baseline that cannot be written is not worth failing a run over */ }
}

/** The stored baseline for one screen, with its screenshot if it survived. */
function get(routeName, language, screenName) {
  const index = load(routeName, language);
  const entry = index.screens && index.screens[screenName];
  if (!entry) return null;
  let png = null;
  try { png = fs.readFileSync(path.join(dirFor(routeName, language), entry.file)); } catch { /* texts only */ }
  return { texts: entry.texts || [], png, at: entry.at };
}

/**
 * Source strings that appear unchanged in the target capture.
 *
 * A cheap first pass that needs no vision call: if a string was English in the
 * baseline and is still character-for-character English here, it was not
 * translated. It says nothing about layout — that is what the paired image
 * comparison is for — and it deliberately ignores the short and the numeric,
 * where an identical string usually means a brand name, a number or a unit
 * rather than a missed translation.
 */
function untranslatedCandidates(baselineTexts, targetTexts) {
  const target = new Set((targetTexts || []).map((t) => String(t).trim()));
  const out = [];
  for (const raw of baselineTexts || []) {
    const s = String(raw).trim();
    if (s.length < 4) continue;                 // "OK", "x2", "1/1"
    // Needs at least one run of two or more letters to be a word at all. A
    // single stray letter is a unit or a multiplier, not a translatable
    // string: quantity labels like "x 60" and "x 12000" are identical in every
    // language by design, and a whole currency shelf of them would otherwise
    // be reported as untranslated on the first comparison.
    if (!/[a-z]{2}/i.test(s)) continue;
    if (target.has(s)) out.push(s);
  }
  return out;
}

module.exports = { record, get, load, untranslatedCandidates, BASELINE_DIR };
