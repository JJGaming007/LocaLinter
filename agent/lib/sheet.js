'use strict';

/**
 * Localization sheet index.
 *
 * The browser sends the already-parsed sheet (header row + data rows), so this
 * module never touches xlsx/Google APIs — it only builds the lookup structures
 * the crawler and the Claude prompt need.
 */

const LANG_CODES = [
  [/^key$/i, null],
  [/english|^en\b|\ben\)/i, 'en'],
  [/french|français|^fr\b/i, 'fr'],
  [/german|deutsch|^de\b/i, 'de'],
  [/italian|italiano|^it\b/i, 'it'],
  [/spanish.*(mexico|latam|mx|419)/i, 'es-MX'],
  [/spanish|español|^es\b/i, 'es'],
  [/portuguese.*(brazil|br)/i, 'pt-BR'],
  [/portuguese|português|^pt\b/i, 'pt'],
  [/russian|русский|^ru\b/i, 'ru'],
  [/turkish|türkçe|^tr\b/i, 'tr'],
  [/arabic|العربية|^ar\b/i, 'ar'],
  [/hebrew|עברית|^he\b|^iw\b/i, 'he'],
  [/persian|farsi|^fa\b/i, 'fa'],
  [/urdu|^ur\b/i, 'ur'],
  [/hindi|हिन्दी|^hi\b/i, 'hi'],
  [/bengali|^bn\b/i, 'bn'],
  [/tamil|^ta\b/i, 'ta'],
  [/telugu|^te\b/i, 'te'],
  [/marathi|^mr\b/i, 'mr'],
  [/indonesian|bahasa|^id\b/i, 'id'],
  [/malay|^ms\b/i, 'ms'],
  [/thai|ไทย|^th\b/i, 'th'],
  [/vietnamese|^vi\b/i, 'vi'],
  [/japanese|日本語|^ja\b/i, 'ja'],
  [/korean|한국어|^ko\b/i, 'ko'],
  [/chinese.*(trad|hant|tw|hk)/i, 'zh-Hant'],
  [/chinese|简体|^zh\b/i, 'zh-Hans'],
  [/polish|^pl\b/i, 'pl'],
  [/dutch|^nl\b/i, 'nl'],
  [/swedish|^sv\b/i, 'sv'],
  [/norwegian|^nb\b|^no\b/i, 'nb'],
  [/danish|^da\b/i, 'da'],
  [/finnish|^fi\b/i, 'fi'],
  [/czech|^cs\b/i, 'cs'],
  [/hungarian|^hu\b/i, 'hu'],
  [/romanian|^ro\b/i, 'ro'],
  [/greek|^el\b/i, 'el'],
  [/ukrainian|^uk\b/i, 'uk'],
  [/filipino|tagalog|^tl\b|^fil\b/i, 'fil'],
];

const RTL_CODES = new Set(['ar', 'he', 'fa', 'ur']);

function codeForHeader(header) {
  const h = String(header || '').trim();
  if (!h) return null;
  for (const [re, code] of LANG_CODES) {
    if (re.test(h)) return code;
  }
  return null;
}

/** Placeholder tokens that must survive translation unchanged. */
const PLACEHOLDER_RE = /(\{[^{}]{0,40}\}|<[^<>]{1,60}>|\[[A-Za-z0-9_.:-]{1,30}\]|%[sdif@]|%\d+\$[sdif@]|\$\{[^}]{1,40}\}|\bN\/A\b)/g;

function extractPlaceholders(text) {
  const out = [];
  const s = String(text || '');
  let m;
  PLACEHOLDER_RE.lastIndex = 0;
  while ((m = PLACEHOLDER_RE.exec(s)) !== null) out.push(m[0]);
  return out;
}

/**
 * Rich-text tags a game renders as formatting rather than as characters.
 *
 * Confirmed on Indus 2.14.0: the Gameplay description draws "Tap" in amber and
 * the rest in white, i.e. the sheet value carries colour tags that never reach
 * the screen. Comparing a screenshot's plain text against the tagged sheet
 * value made every such string look like a mismatch, so the tags come out of
 * both sides before anything is compared. Only formatting tags are stripped —
 * a real placeholder like {0} still has to survive, because a missing one is a
 * genuine defect.
 */
const MARKUP_RE = /<\/?(?:color|size|b|i|u|s|sprite|font|align|material|quad|mark|nobr|indent|line-height|cspace|mspace|voffset|width|style|gradient|rotate|link|lowercase|uppercase|smallcaps|sup|sub|alpha|noparse|pos|space)\b[^>]*>/gi;

/** Drop formatting tags, keep everything the player actually reads. */
function stripMarkup(text) {
  return String(text == null ? '' : text).replace(MARKUP_RE, '');
}

/** Aggressive normalization used for exact lookups. */
function norm(text) {
  return stripMarkup(text)
    .replace(/‏|‎|‪|‫|‬|­/g, '') // bidi marks, soft hyphen
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/**
 * UI labels a description quotes by name.
 *
 * Confirmed on Indus: an avatar's lore reads `Click "Customize"`. If the
 * description is translated but the button it names is not (or the reverse) the
 * instruction sends the player looking for a control that does not exist under
 * that name. Pulling the quoted fragments out lets a caller check them against
 * the labels actually on screen, in the same locale.
 */
const QUOTED_RE = /[\"“”'‘’«»„][^\"“”'‘’«»„]{2,40}[\"“”'‘’«»„]/g;
function quotedUiReferences(text) {
  const out = [];
  for (const m of stripMarkup(text).matchAll(QUOTED_RE)) {
    const inner = m[0].slice(1, -1).trim();
    // A quoted sentence is prose, not a control name; labels are short.
    if (inner && inner.split(/\s+/).length <= 4 && !/[.!?]$/.test(inner)) out.push(inner);
  }
  return out;
}

/** Structure-only form: placeholders and punctuation removed. Used for fuzzy matching. */
function skeleton(text) {
  return norm(text)
    .replace(PLACEHOLDER_RE, ' ')
    .replace(/[^\p{L}\p{N}\s]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function trigrams(s) {
  const t = ` ${s} `;
  const out = new Set();
  for (let i = 0; i < t.length - 2; i++) out.add(t.slice(i, i + 3));
  return out;
}

function diceSimilarity(a, b) {
  if (!a.size || !b.size) return 0;
  let shared = 0;
  for (const g of a) if (b.has(g)) shared++;
  return (2 * shared) / (a.size + b.size);
}

class SheetIndex {
  /**
   * @param {string[]} headers  header row
   * @param {any[][]} rows      data rows (row 0 = first data row)
   */
  constructor(headers, rows) {
    this.headers = (headers || []).map((h) => String(h == null ? '' : h));
    this.rows = rows || [];

    this.languages = [];
    this.headers.forEach((h, i) => {
      if (i === 0) return; // column 0 is the key
      const code = codeForHeader(h);
      this.languages.push({ header: h, code: code || h, index: i, rtl: RTL_CODES.has(code) });
    });

    this.englishCol = this.languages.find((l) => l.code === 'en') || this.languages[0] || null;

    this.entries = [];
    this.rows.forEach((row, r) => {
      if (!row || !row.length) return;
      const key = String(row[0] == null ? '' : row[0]).trim();
      const values = {};
      let any = false;
      for (const lang of this.languages) {
        const v = row[lang.index];
        const s = v == null ? '' : String(v);
        values[lang.header] = s;
        if (s.trim()) any = true;
      }
      if (!key && !any) return;
      this.entries.push({ key, rowNumber: r + 2, values }); // +2: 1-based + header row
    });

    // exact index:  normalized text -> [{ entryIdx, header }]
    this.byNorm = new Map();
    // fuzzy index:  entryIdx/header -> trigram set of the skeleton
    this.fuzzy = [];
    this.entries.forEach((e, entryIdx) => {
      for (const lang of this.languages) {
        const raw = e.values[lang.header];
        if (!raw || !raw.trim()) continue;
        const n = norm(raw);
        if (!this.byNorm.has(n)) this.byNorm.set(n, []);
        this.byNorm.get(n).push({ entryIdx, header: lang.header, code: lang.code });
        const sk = skeleton(raw);
        if (sk.length >= 3) {
          this.fuzzy.push({ entryIdx, header: lang.header, code: lang.code, sk, grams: trigrams(sk) });
        }
      }
      if (e.key) {
        const n = norm(e.key);
        if (!this.byNorm.has(n)) this.byNorm.set(n, []);
        this.byNorm.get(n).push({ entryIdx, header: '__key__', code: 'key' });
      }
    });
  }

  entry(idx) {
    return this.entries[idx];
  }

  languageByCodeOrHeader(needle) {
    if (!needle) return null;
    const n = String(needle).toLowerCase();
    return (
      this.languages.find((l) => String(l.code).toLowerCase() === n) ||
      this.languages.find((l) => l.header.toLowerCase() === n) ||
      this.languages.find((l) => String(l.code).toLowerCase().startsWith(n.split(/[-_]/)[0])) ||
      null
    );
  }

  /** Every exact match of a rendered string, across all languages. */
  lookupExact(text) {
    const hits = this.byNorm.get(norm(text)) || [];
    return hits.map((h) => ({ ...h, entry: this.entries[h.entryIdx] }));
  }

  /** Best fuzzy matches, sorted by similarity desc. */
  lookupFuzzy(text, { limit = 4, minScore = 0.45, header = null } = {}) {
    const sk = skeleton(text);
    if (sk.length < 3) return [];
    const grams = trigrams(sk);
    const scored = [];
    for (const cand of this.fuzzy) {
      if (header && cand.header !== header) continue;
      // cheap length gate before the set intersection
      if (Math.abs(cand.sk.length - sk.length) > Math.max(24, sk.length)) continue;
      const score = diceSimilarity(grams, cand.grams);
      if (score >= minScore) scored.push({ ...cand, score, entry: this.entries[cand.entryIdx] });
    }
    scored.sort((a, b) => b.score - a.score);
    // one row may match through several columns — keep the best per row
    const seen = new Set();
    const out = [];
    for (const s of scored) {
      if (seen.has(s.entryIdx)) continue;
      seen.add(s.entryIdx);
      out.push(s);
      if (out.length >= limit) break;
    }
    return out;
  }

  summary() {
    return {
      rows: this.entries.length,
      languages: this.languages.map((l) => ({ header: l.header, code: l.code, rtl: l.rtl })),
      english: this.englishCol ? this.englishCol.header : null,
    };
  }
}

module.exports = {
  stripMarkup,
  quotedUiReferences,
  SheetIndex,
  norm,
  skeleton,
  extractPlaceholders,
  codeForHeader,
  RTL_CODES,
};
