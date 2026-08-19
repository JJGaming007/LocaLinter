'use strict';

const Anthropic = require('@anthropic-ai/sdk');

/**
 * Claude vision analysis for a single captured screen.
 *
 * The deterministic checks in checks.js already found everything that is
 * mechanically detectable from the exact strings and rects. This pass exists to
 * catch what only a pair of eyes can: clipping the engine does not report,
 * collisions with art, unreadable contrast, mirrored layouts, text baked into
 * textures, and translations that are grammatically or contextually wrong.
 */

const ISSUE_TYPES = [
  'truncated',
  'overflow_horizontal',
  'overflow_vertical',
  'offscreen',
  'overlap',
  'clipped_by_art',
  'unreadable_contrast',
  'mojibake',
  'wrong_font_glyphs',
  'untranslated',
  'missing_translation',
  'wrong_language',
  'mixed_language',
  'text_mismatch',
  'placeholder_mismatch',
  'unresolved_placeholder',
  'rtl_reversed',
  'rtl_layout_not_mirrored',
  'number_format',
  'date_format',
  'not_in_sheet',
  'baked_into_texture',
  'grammar_or_wording',
  'inconsistent_terminology',
  'empty_label',
  'other',
];

const RESULT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['issues', 'unlisted_text', 'screen_summary'],
  properties: {
    screen_summary: {
      type: 'string',
      description: 'One short sentence naming what this screen is (e.g. "Main lobby with shop and settings buttons").',
    },
    unlisted_text: {
      type: 'array',
      description:
        'Every piece of readable text visible in the screenshot that does NOT appear in the supplied extracted-strings list. These are strings the engine could not report — baked into textures, rendered by a custom system, or inside a native dialog. Empty array if there are none.',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['text', 'where'],
        properties: {
          text: { type: 'string' },
          where: { type: 'string', description: 'Where on screen, in plain words.' },
        },
      },
    },
    issues: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['type', 'severity', 'confidence', 'text', 'where', 'message'],
        properties: {
          type: { type: 'string', enum: ISSUE_TYPES },
          severity: { type: 'string', enum: ['high', 'medium', 'low'] },
          confidence: { type: 'string', enum: ['certain', 'likely', 'possible'] },
          text: { type: 'string', description: 'The offending on-screen string, verbatim.' },
          where: { type: 'string', description: 'Where on screen it is, in plain words.' },
          element: { type: 'string', description: 'Hierarchy path from the extracted list, or "" if unknown.' },
          key: { type: 'string', description: 'Localization key from the supplied candidates, or "".' },
          expected: { type: 'string', description: 'What the sheet says it should be, or "".' },
          message: { type: 'string', description: 'One or two sentences: what is wrong and what a fixer should do.' },
          // Without this there is nothing to crop around, and the whole
          // magnified second look is impossible. Optional, because a finding
          // about a string the model read but cannot place is still worth
          // having — it just cannot be double-checked.
          rect: {
            type: 'object',
            description: 'Where the offending string sits, as fractions of the image. Omit only if you cannot place it.',
            additionalProperties: false,
            required: ['x', 'y', 'w', 'h'],
            properties: {
              x: { type: 'number', description: 'Left edge, 0 = left of image, 1 = right.' },
              y: { type: 'number', description: 'Top edge, 0 = top of image, 1 = bottom.' },
              w: { type: 'number', description: 'Width as a fraction of the image width.' },
              h: { type: 'number', description: 'Height as a fraction of the image height.' },
            },
          },
        },
      },
    },
  },
};

const SYSTEM_PROMPT = `You are a localization QA engineer reviewing screenshots of a game build against its localization spreadsheet.

Your job on each screen is coverage: find every localization defect a player in this language would notice. Report everything you find, including findings you are only somewhat sure about — mark those with a lower confidence rather than dropping them. A finding that gets filtered out later costs nothing; a missed bug ships.

What to look for, beyond what has already been checked mechanically:
- Text visually cut off, clipped by a panel edge, or running under artwork, even when the engine reports it fits.
- Labels colliding with icons, borders, buttons, or each other because the translation is longer than the source.
- Text that is unreadable against its background, or too small after a font fallback.
- Glyphs the font cannot render: boxes, question marks, disconnected Arabic letterforms, missing diacritics.
- Right-to-left screens: text rendered in visual instead of logical order, punctuation on the wrong side, layouts that should mirror but do not, numbers or Latin fragments in the wrong position inside an RTL run.
- A screen that mixes the target language with the source language.
- Translations that are grammatically wrong, mistranslated for the UI context, or use a term inconsistently with the rest of the screen.
- Numbers, currency, dates, and time formatted with the wrong separators or order for the locale.
- Text baked into a texture or sprite that was never localized.
- Overlapping or duplicated strings, empty labels where a value was expected.

Ground every finding in what is actually visible in the screenshot. Do not invent defects to fill the list, and do not repeat a finding that is already listed under "already detected" unless you are correcting it. If the screen looks correct, return an empty issues array.

Report locations in plain words a person can act on ("the yellow CLAIM button, upper right"), not coordinates.`;

/**
 * Published list prices, US dollars per million tokens. Cache reads bill at
 * 0.1x input and cache writes at 1.25x, so a run's cost can be computed exactly
 * from the token counts the API returns — no estimating.
 */
const PRICING = {
  'claude-opus-5': { input: 5, output: 25 },
  'claude-opus-4-8': { input: 5, output: 25 },
  'claude-sonnet-5': { input: 3, output: 15 },
  'claude-haiku-4-5': { input: 1, output: 5 },
};

function priceFor(model, baseUrl) {
  // Behind a gateway the published rates do not apply — report tokens and let
  // the run say `priced: false` rather than invent a number.
  if (baseUrl) return null;
  return PRICING[model] || null;
}

class ClaudeAnalyzer {
  constructor({ apiKey, model = 'claude-opus-5', effort = 'high', baseUrl = '', extraChecks = '', memory = '' }) {
    if (!apiKey) throw new Error('No Anthropic API key configured.');
    // A base URL points the SDK at a company gateway (LiteLLM and friends)
    // that speaks the same /v1/messages format. Empty means Anthropic direct.
    this.client = new Anthropic({ apiKey, ...(baseUrl ? { baseURL: baseUrl } : {}) });
    this.baseUrl = baseUrl || '';
    this.model = model;
    this.effort = effort;
    // Whatever the team wants looked for on top of the built-in checks —
    // house style, terms that must never appear, a glossary rule.
    this.extraChecks = String(extraChecks || '').trim();
    // What earlier runs learned about this app. Cached like the main prompt,
    // because it is identical on every screen of every run.
    this.memory = String(memory || '').trim();
    this.usage = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, calls: 0, costUSD: 0, priced: true };
    // Server-side refusal fallback is a beta; if this key cannot use it we stop
    // asking rather than failing every screen.
    this.fallbackEnabled = true;
  }

  /** Request options shared by every call, including the refusal fallback. */
  _common() {
    return this.fallbackEnabled
      ? { betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' }
      : {};
  }

  get _messages() {
    return this.fallbackEnabled ? this.client.beta.messages : this.client.messages;
  }

  /**
   * A key without the fallback beta rejects the request outright. That must not
   * cost us the whole scan, so the first such failure disables the feature and
   * the call is retried plain.
   */
  async _send(build) {
    try {
      return await build();
    } catch (e) {
      const msg = String(e && e.message || e);
      if (this.fallbackEnabled && /fallback|beta/i.test(msg) && /400|invalid_request/i.test(msg)) {
        this.fallbackEnabled = false;
        return build();
      }
      throw e;
    }
  }

  _track(usage) {
    if (!usage) return;
    this.usage.calls++;
    const input = usage.input_tokens || 0;
    const output = usage.output_tokens || 0;
    const cacheRead = usage.cache_read_input_tokens || 0;
    const cacheWrite = usage.cache_creation_input_tokens || 0;
    this.usage.input += input;
    this.usage.output += output;
    this.usage.cacheRead += cacheRead;
    this.usage.cacheWrite += cacheWrite;

    const price = priceFor(this.model, this.baseUrl);
    if (!price) {
      // Unknown model: report tokens, but never invent a dollar figure.
      this.usage.priced = false;
      return;
    }
    this.usage.costUSD +=
      (input * price.input + cacheWrite * price.input * 1.25 + cacheRead * price.input * 0.1 + output * price.output)
      / 1e6;
  }

  /**
   * @param {Buffer} png            screenshot
   * @param {object} ctx            { screenId, scene, targetHeader, targetCode, rtl, sourceHeader, mode }
   * @param {Array}  extracted      [{ path, text, rect, matches:[{key,row,header,value,score}] }]
   * @param {Array}  staticIssues   findings from checks.js
   * @returns {Promise<object>} parsed result matching RESULT_SCHEMA
   */
  async analyzeScreen(png, ctx, extracted, staticIssues) {
    const lines = [];
    lines.push(`Screen: ${ctx.screenId}${ctx.scene ? ` (scene: ${ctx.scene})` : ''}`);
    lines.push(`Language under test: ${ctx.targetHeader}${ctx.targetCode ? ` [${ctx.targetCode}]` : ''}${ctx.rtl ? ' — right-to-left' : ''}`);
    lines.push(`Source language column: ${ctx.sourceHeader || 'unknown'}`);
    lines.push(`Capture source: ${ctx.mode}`);
    lines.push('');

    if (extracted && extracted.length) {
      lines.push(`Strings the engine reports on this screen (${extracted.length}), with their sheet matches:`);
      for (const t of extracted) {
        const r = t.rect
          ? ` @${Math.round(t.rect.x)},${Math.round(t.rect.y)} ${Math.round(t.rect.w)}x${Math.round(t.rect.h)}`
          : '';
        lines.push(`- "${t.text}"${r}${t.path ? `  [${t.path}]` : ''}`);
        if (t.matches && t.matches.length) {
          for (const m of t.matches.slice(0, 3)) {
            lines.push(
              `    ↳ sheet row ${m.row} key "${m.key}" · ${m.header}: "${m.value}"${
                m.score != null ? ` (${Math.round(m.score * 100)}% match)` : ' (exact)'
              }`
            );
          }
        } else {
          lines.push('    ↳ no match anywhere in the sheet');
        }
      }
    } else {
      lines.push(
        'The engine could not report any strings for this screen (no in-game bridge, or the UI is drawn outside the supported system). Read every string directly from the screenshot and list them all under unlisted_text.'
      );
    }
    lines.push('');

    if (staticIssues && staticIssues.length) {
      lines.push(`Already detected mechanically — do not repeat these unless you are correcting one (${staticIssues.length}):`);
      for (const i of staticIssues.slice(0, 60)) {
        lines.push(`- [${i.type}] "${i.text}" — ${i.message}`);
      }
    } else {
      lines.push('Nothing was detected mechanically on this screen.');
    }
    lines.push('');
    lines.push('Review the screenshot now and report what those checks cannot see.');

    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 16000,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: this.effort,
        format: { type: 'json_schema', schema: RESULT_SCHEMA },
      },
      system: [
        { type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } },
        ...(this.memory ? [{
          type: 'text',
          text: this.memory,
          cache_control: { type: 'ephemeral' },
        }] : []),
        ...(this.extraChecks ? [{
          type: 'text',
          text: `Additional checks requested for this project. Apply them alongside everything above, and report anything that breaks them as an issue:
${this.extraChecks}`,
          cache_control: { type: 'ephemeral' },
        }] : []),
      ],
      messages: [
        {
          role: 'user',
          content: [
            {
              type: 'image',
              source: { type: 'base64', media_type: 'image/png', data: png.toString('base64') },
            },
            { type: 'text', text: lines.join('\n') },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);

    if (message.stop_reason === 'refusal') {
      throw new Error(
        `Claude declined to analyse this screen${
          message.stop_details && message.stop_details.category ? ` (${message.stop_details.category})` : ''
        }.`
      );
    }

    const text = (message.content || [])
      .filter((b) => b.type === 'text')
      .map((b) => b.text)
      .join('');
    if (!text.trim()) return { issues: [], unlisted_text: [], screen_summary: '' };
    try {
      return JSON.parse(text);
    } catch {
      throw new Error('Claude returned malformed JSON for this screen.');
    }
  }

  /**
   * Vision-driven navigation for builds without the in-game bridge.
   * Returns tap targets in normalized 0..1 screen coordinates.
   *
   * @param {Buffer} png
   * @param {string[]} alreadyTried  labels already tapped on this screen
   */
  async proposeTargets(png, alreadyTried = []) {
    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 4000,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: 'medium',
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['targets'],
            properties: {
              targets: {
                type: 'array',
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['label', 'x', 'y', 'kind', 'priority'],
                  properties: {
                    label: { type: 'string', description: 'What the control says or depicts.' },
                    x: { type: 'number', description: 'Horizontal centre, 0 = left edge, 1 = right edge.' },
                    y: { type: 'number', description: 'Vertical centre, 0 = top edge, 1 = bottom edge.' },
                    kind: { type: 'string', enum: ['tap', 'long_press', 'scroll_down', 'scroll_right', 'back'] },
                    priority: { type: 'string', enum: ['high', 'medium', 'low'] },
                  },
                },
              },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            {
              type: 'image',
              source: { type: 'base64', media_type: 'image/png', data: png.toString('base64') },
            },
            {
              type: 'text',
              text: `This is a screenshot of a game. List every control a player could interact with, so an automated crawler can open every screen that contains text.

Be exhaustive. Include the small ones that are easy to overlook: "i" info badges, "?" help chips, gear and settings icons, dropdown carets, tabs, arrows, close buttons, currency and resource chips at the top of the screen, avatar and profile buttons, and anything that looks like it opens a tooltip, flyout, or popup. Use long_press for anything that looks like it has a tooltip rather than a screen behind it. Add a scroll_down or scroll_right target when a list or row clearly continues past the visible area.

Give coordinates as fractions of the image: x 0 is the left edge and 1 the right edge, y 0 is the top and 1 the bottom. Aim at the centre of each control.

Prioritise controls that lead to text-heavy screens.

${alreadyTried.length ? `Already tried on this screen, skip them: ${alreadyTried.map((l) => `"${l}"`).join(', ')}` : 'Nothing has been tried on this screen yet.'}`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const parsed = JSON.parse(text);
      return Array.isArray(parsed.targets) ? parsed.targets : [];
    } catch {
      return [];
    }
  }

  /**
   * Put the same screen in two languages side by side and report the
   * differences that matter.
   *
   * Everything here rests on one idea: a defect is only the translation's
   * fault if the source does not have it too. The pair is what makes that
   * judgeable, so the prompt spends most of its words insisting the comparison
   * is actually made rather than assumed — including, explicitly, permission
   * to conclude that nothing is wrong with the translation because the source
   * is broken in the same way.
   */
  async compareToBaseline(baselinePng, targetPng, ctx) {
    const identical = (ctx.identical || []).slice(0, 40);
    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 8000,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: this.effort,
        format: { type: 'json_schema', schema: RESULT_SCHEMA },
      },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: baselinePng.toString('base64') } },
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: targetPng.toString('base64') } },
            {
              type: 'text',
              text: `These are the same screen — "${ctx.screenName}" — captured twice.

Image 1 is the source language, ${ctx.baselineLanguage}.
Image 2 is the language under test, ${ctx.targetHeader}.

Report only what the COMPARISON reveals. Anything visible in image 2 alone has already been checked; repeating it here just duplicates findings.

Look for:

1. Text that no longer fits. Find each label in both images and compare how it sits in its container. Report it when the target text is clipped at an edge, breaks mid-word, wraps where the source did not, spills outside its own button or background plate, or runs under artwork that the source clears. Say what the source did and what the target does — "'Rewards' fits on one line, 'Recompensas' breaks mid-word into 'Recompen' / 'sas'" — because that contrast is the evidence.

2. Text that did not change. A string identical to the source is usually untranslated. Ignore brand names, product names, player names and proper nouns, which are meant to stay put.${identical.length ? `\n\nStrings already found to be character-for-character identical: ${identical.map((s) => `"${s}"`).join(', ')}. Judge which are genuinely untranslated and which are names.` : ''}

3. Content that does not correspond — the target saying something the source does not, beyond ordinary translation.

Two things you must NOT report, and they matter as much as what you do:

- A problem that is present in BOTH images. If a counter runs under the same artwork in image 1, if a row is blank in image 1, if a screen has an inconsistent name in image 1, then the source has that defect and the translation did not cause it. Say nothing about it. Filing it against the translation sends the wrong team after it.

- Content that legitimately differs between captures. Carousels, rotating promo slots, countdown timers, live counters and anything showing another player's name will differ for reasons that have nothing to do with language. If a panel's content is simply different rather than a translation of the same thing, it is a different card, not a defect.

If the two images show no translation-caused difference, return an empty issues list. That is a useful and common answer.`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    if (!text.trim()) return { issues: [] };
    try {
      return JSON.parse(text);
    } catch {
      return { issues: [] };
    }
  }

  /**
   * Find one specific string on screen and say where it is.
   *
   * Used to pick a row out of a scrolling list — a language in the picker, say
   * — where there is no fixed coordinate to aim at because the list moves.
   * Deliberately narrow and cheap: one string in, one point out, so the gap
   * between looking and tapping stays short. A list that re-anchors itself a
   * second after it settles punishes anything slower.
   *
   * With `requireChecked` it answers a different question — is this row's
   * checkbox ticked? — which is how a selection is confirmed before something
   * irreversible is pressed.
   *
   * @returns {Promise<{x:number,y:number}|null>} normalized centre, or null.
   */
  async locateText(png, needle, { requireChecked = false } = {}) {
    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 2000,
      output_config: {
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['found'],
            properties: {
              found: { type: 'boolean', description: 'Is the row visible on this screenshot?' },
              x: { type: 'number', description: 'Horizontal centre of the row, 0 = left edge, 1 = right edge.' },
              y: { type: 'number', description: 'Vertical centre of the row, 0 = top edge, 1 = bottom edge.' },
              checked: { type: 'boolean', description: 'Does the row show a ticked checkbox or a selected mark?' },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: png.toString('base64') } },
            {
              type: 'text',
              text: `Find the row labelled "${needle}" in this screenshot.

Answer found:false if it is not visible, including when it is only partly visible or cut off at the top or bottom edge — a half-visible row cannot be tapped reliably, and saying it is there when it is not selects the wrong thing.

If it is visible, give the centre of the row as fractions of the image: x 0 is the left edge and 1 the right edge, y 0 is the top and 1 the bottom.${
                requireChecked
                  ? '\n\nAlso report whether that row is currently SELECTED — a ticked checkbox, a check mark, or an obvious highlight that the other rows do not have. Be strict: report checked:false if you are unsure.'
                  : ''
              }`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const p = JSON.parse(text);
      if (!p.found) return null;
      if (requireChecked && !p.checked) return null;
      if (!Number.isFinite(p.x) || !Number.isFinite(p.y)) return null;
      return { x: p.x, y: p.y };
    } catch {
      return null;
    }
  }

  /**
   * Decide, for each string read off the screen, what the sheet says about it.
   *
   * This is the comparison the whole product is for, and on a title with no
   * accessibility bridge it was not happening. The deterministic path reconciles
   * `state.texts` against the sheet — but those come from the bridge, so on a
   * Unity build that draws into a SurfaceView the list is empty, matchesFor
   * never runs, and the strings the vision pass reads are never looked up at
   * all.
   *
   * Matching them mechanically is not good enough either. A shipped string is
   * rarely byte-identical to its row: placeholders are filled in, the UI
   * upper-cases it, punctuation drifts, a counter is appended. Normalised
   * comparison then reports a correct translation as missing, and fuzzy
   * comparison reports nonsense — asked about a settings toggle it offered
   * "Light Ammo" at 46% as the nearest row.
   *
   * So retrieval stays mechanical, because narrowing 6,747 rows to a shortlist
   * is what an index is for, and the judgement moves here.
   */
  async reconcileWithSheet({ strings, target, source, screenName = '' }) {
    if (!strings || !strings.length) return [];

    const lines = strings.map((s, i) => {
      const cands = (s.candidates || []).slice(0, 3).map((c) => {
        // How it was found matters as much as what was found: without this a
        // 46% trigram guess reads exactly like an exact hit, and the model
        // adopts the junk suggestion instead of saying nothing matched.
        const how = c.score == null || c.score >= 1
          ? 'exact match'
          : `approximate, ${Math.round(c.score * 100)}% similar`;
        return `        row ${c.row} (${how}) key=${JSON.stringify(c.key || '')} `
          + `${source}=${JSON.stringify(c.source || '')} ${target}=${JSON.stringify(c.value || '')}`;
      });
      return `  [${i}] ${JSON.stringify(s.text)}\n${cands.length ? cands.join('\n') : '        (no candidate rows found)'}`;
    });

    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 8000,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: 'medium',
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['verdicts'],
            properties: {
              verdicts: {
                type: 'array',
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['index', 'status'],
                  properties: {
                    index: { type: 'number', description: 'The [n] of the on-screen string.' },
                    status: {
                      type: 'string',
                      enum: ['correct', 'untranslated', 'wrong_translation', 'not_in_sheet', 'variant'],
                      description: 'variant = matches its row allowing for placeholders, case or UI formatting.',
                    },
                    key: { type: 'string', description: 'The matching row key, or "".' },
                    row: { type: 'number', description: 'The matching row number, or 0.' },
                    expected: { type: 'string', description: 'What the sheet says it should be, or "".' },
                    severity: { type: 'string', enum: ['high', 'medium', 'low'] },
                    note: { type: 'string', description: 'One sentence, only when something is wrong.' },
                  },
                },
              },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            {
              type: 'text',
              text: `These strings were read off a game screen${screenName ? ` (${screenName})` : ''} that is running in **${target}**. Under each one are the closest rows an index could find in the localization sheet, with that row's ${source} and ${target} values.

Decide, for each string, what the sheet says about it.

${lines.join('\n')}

Statuses:

- **correct** — it matches its row's ${target} value.
- **variant** — it is that row, and the difference is mechanical rather than a mistake: a filled-in placeholder, upper-casing by the UI, a trailing counter or punctuation, a line break. This is NOT a defect and matters as much as the others: reporting these is what makes a report untrustworthy.
- **untranslated** — the row exists and has a ${target} value, but the screen is showing the ${source} text instead. Do not use this for brand names, product names or strings that are identical in both languages by nature (numbers, "OK", "HYBRID", proper nouns).
- **wrong_translation** — it matches a row, but what is on screen differs from that row's ${target} value in a way that changes the meaning. Say what the sheet expected.
- **not_in_sheet** — none of the candidate rows is this string. It is hardcoded or the key has drifted, so no translator can see or fix it. Say this only when you are confident none of the candidates is a match — a bad fuzzy suggestion is not a match.

Each candidate says how it was found. An **exact match** is reliable. An **approximate** one is a guess from a trigram index, and below about 70% it usually means nothing was found at all — treat that as no candidate rather than as the answer. Judge the string, not the suggestion.

Return a verdict for every string. Leave note empty when the status is correct or variant.`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const p = JSON.parse(text);
      return Array.isArray(p.verdicts) ? p.verdicts : [];
    } catch {
      return [];
    }
  }

  /**
   * Given a goal and the screen in front of us, decide the single next action.
   *
   * This replaces writing the steps down. A recorded procedure — tap the gear,
   * tap Account, tap Change, find the row — assumes the app is where it was
   * when someone wrote it, and on this title it usually is not: a promo covers
   * the lobby, a list re-anchors, a layout shifts, and every tap after the
   * first missed one lands somewhere arbitrary while the script reports
   * success. That failed repeatedly and in silence.
   *
   * Looking before each move costs a call and cannot fail that way. The model
   * gets the goal, the current screen and what has already been tried, and
   * returns one action. The caller executes it and looks again.
   *
   * `history` matters more than it looks: without it the model proposes the
   * same tap forever when a screen does not respond.
   */
  async decideNextAction(png, goal, history = [], hints = []) {
    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 1500,
      thinking: { type: 'adaptive' },
      output_config: {
        effort: 'medium',
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['action', 'why'],
            properties: {
              action: {
                type: 'string',
                enum: ['tap', 'swipe', 'back', 'wait', 'done', 'stuck'],
                description: 'done when the goal is visibly achieved; stuck when nothing on this screen can advance it.',
              },
              x: { type: 'number', description: 'Tap point or swipe start, 0-1 across the image.' },
              y: { type: 'number', description: 'Tap point or swipe start, 0-1 down the image.' },
              x2: { type: 'number', description: 'Swipe end, 0-1 across.' },
              y2: { type: 'number', description: 'Swipe end, 0-1 down.' },
              target: { type: 'string', description: 'What you are aiming at, in a few words.' },
              why: { type: 'string', description: 'One sentence: why this advances the goal.' },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: png.toString('base64') } },
            {
              type: 'text',
              text: `You are driving a game on a phone over adb, one action at a time, to reach a goal. This screenshot is the screen right now.

GOAL: ${goal}

${history.length
  ? `What you have already done, oldest first:\n${history.map((h, i) => `  ${i + 1}. ${h}`).join('\n')}\n\nIf the last action changed nothing, do NOT repeat it — try a different control, or scroll to bring something else into view.`
  : 'Nothing has been tried yet.'}
${hints.length ? `\nWorth knowing about this app:\n${hints.map((h) => `  - ${h}`).join('\n')}\n` : ''}
Answer with ONE action:

- tap    — give x,y at the centre of the control. Only tap something you can actually see.
- swipe  — give x,y to start and x2,y2 to end. To reveal items ABOVE the visible ones, start in the MIDDLE of the list and drag DOWNWARD; starting at the very edge of a panel usually lands on its frame and does nothing.
- back   — the system back button. Avoid it on a main or home screen, where it usually offers to quit.
- wait   — the screen is mid-animation or loading.
- done   — the goal is visibly achieved on this screen. Do not say done in the hope that it worked.
- stuck  — nothing here can advance the goal.

Coordinates are fractions of the image: x 0 is the left edge and 1 the right, y 0 is the top and 1 the bottom.

Prefer the smallest step that makes progress, and say plainly in "why" what you expect to happen — if the next screenshot does not show it, that is the signal you were wrong.`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const p = JSON.parse(text);
      if (!p || typeof p.action !== 'string') return null;
      return {
        action: p.action,
        x: Number.isFinite(p.x) ? p.x : null,
        y: Number.isFinite(p.y) ? p.y : null,
        x2: Number.isFinite(p.x2) ? p.x2 : null,
        y2: Number.isFinite(p.y2) ? p.y2 : null,
        target: String(p.target || ''),
        why: String(p.why || ''),
      };
    } catch {
      return null;
    }
  }

  /**
   * Look again at one finding, magnified, and keep it or drop it.
   *
   * Every defect about pixels rather than meaning — a clipped descender, a
   * missing tone mark, text touching its border — is decided at full-screen
   * size on a few pixels, and that is where a scan invents things. Driven by
   * hand on the Thai pass this second look retracted two findings and confirmed
   * a third; the two it retracted were correctly spelled strings whose marks
   * were simply below the resolution being judged at.
   *
   * The prompt is written to make retracting easy. A report that is mostly
   * right is worth less than a shorter one that is entirely right, because the
   * first wrong entry is where a translator stops believing the rest.
   */
  async verifyFinding(cropPng, issue, context = {}) {
    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 800,
      output_config: {
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['holds', 'why'],
            properties: {
              holds: { type: 'boolean', description: 'Is the reported defect really there at this magnification?' },
              why: { type: 'string', description: 'One sentence. If it does not hold, say what is actually there.' },
              severity: { type: 'string', enum: ['high', 'medium', 'low'], description: 'Revised severity, if it holds.' },
              text: { type: 'string', description: 'The string as it actually reads at this size, if different.' },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: cropPng.toString('base64') } },
            {
              type: 'text',
              text: `This is a magnified crop of one small part of a game screen — the area around a string that a first pass, looking at the whole screen at once, reported as a defect.

The report was:

  type:     ${issue.type}
  string:   ${JSON.stringify(issue.text || '')}
  where:    ${issue.where || 'unspecified'}
  severity: ${issue.severity}
  claim:    ${issue.message || ''}
${context.language ? `  language: ${context.language}\n` : ''}
Decide whether that defect is really there, now that you can see it properly.

Retract it — holds:false — if any of these is true:
- the string is complete and correctly spelled, and the marks the first pass thought were missing are present at this size;
- the text fits inside its container, with the descenders and diacritics inside the border, even if it is close;
- the "overlap" is a background pattern, a shadow, a border or artwork rather than another string;
- what looked like corruption is a font style, an icon, or a script you can read correctly here.

Keep it — holds:true — only if you can point at the defect in this image. Truncation means a word actually stops; overflow means glyphs cross the container edge; a missing mark means you can see the bare vowel.

Be willing to retract. The first pass was guessing from a few pixels and this is the check on it, so agreeing by default makes the check worthless. If the crop does not show enough to decide, retract rather than guess.`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const p = JSON.parse(text);
      return {
        holds: p.holds !== false,
        why: String(p.why || ''),
        severity: ['high', 'medium', 'low'].includes(p.severity) ? p.severity : null,
        text: typeof p.text === 'string' && p.text.trim() ? p.text.trim() : null,
      };
    } catch {
      return { holds: true, why: '', severity: null, text: null };   // unparseable: leave it alone
    }
  }

  /**
   * Name the screen in front of us, from a shortlist the route map supplies.
   *
   * The crawler could already do this from text — but only with the Unity
   * bridge attached. On a title that draws its UI into a SurfaceView there are
   * no strings until something has read the pixels, so the modal-dismissing
   * code was matching against an empty array and concluding, every single
   * time, that nothing was covering the screen. The promo interstitial then sat
   * there while four recorded taps fired into it and the language switch
   * silently did nothing.
   *
   * One small call fixes that, and it is far cheaper than the alternative: a
   * full analyzeScreen on a modal whose only useful property is which button
   * closes it.
   */
  async identifyScreen(png, candidates) {
    const list = (candidates || [])
      .filter((c) => c && c.name && Array.isArray(c.anyText) && c.anyText.length)
      .slice(0, 40);
    if (!list.length) return null;

    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 500,
      output_config: {
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['name'],
            properties: {
              name: { type: 'string', description: 'The matching screen name, or "" if none match.' },
              confident: { type: 'boolean', description: 'True only if the match is unambiguous.' },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: png.toString('base64') } },
            {
              type: 'text',
              text: `Which of these known screens is this screenshot showing?

${list.map((c) => `- ${c.name}: shows text such as ${c.anyText.slice(0, 6).map((t) => JSON.stringify(t)).join(', ')}`).join('\n')}

The sample texts are only hints, and may be in a different language than the screenshot — match on what the screen IS, not on an exact string. A full-screen promotional or reward panel covering the app counts as that panel, not as whatever is behind it.

Answer with the name exactly as written above. If none of them describe this screen, answer with an empty name.`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const p = JSON.parse(text);
      const name = String(p.name || '').trim();
      if (!name || p.confident === false) return null;
      return list.some((c) => c.name === name) ? name : null;
    } catch {
      return null;
    }
  }

  /**
   * Read every row currently visible in a list, top to bottom.
   *
   * locateText answers "is this one row here?", which is the wrong question
   * for a list that will not hold still. To get to an entry the list keeps
   * pulling away from, the crawler has to know what it *can* reach right now
   * and pick a stepping stone — and asking row by row would cost a call per
   * guess. One call returns the whole visible window instead.
   *
   * Rows cut off at either edge are reported with visible:false: they render,
   * but a tap on one lands on nothing, which is a failure that leaves no trace
   * on screen.
   */
  async readListRows(png) {
    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 2000,
      output_config: {
        format: {
          type: 'json_schema',
          schema: {
            type: 'object',
            additionalProperties: false,
            required: ['rows'],
            properties: {
              rows: {
                type: 'array',
                description: 'The list rows on screen, in the order they appear, top first.',
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['label', 'y', 'visible'],
                  properties: {
                    label: { type: 'string', description: 'The row text, exactly as printed.' },
                    x: { type: 'number', description: 'Horizontal centre of the row, 0 = left, 1 = right.' },
                    y: { type: 'number', description: 'Vertical centre of the row, 0 = top, 1 = bottom.' },
                    visible: { type: 'boolean', description: 'True only if the whole row is inside the list, not clipped at either edge.' },
                    checked: { type: 'boolean', description: 'Does the row show a ticked checkbox or selected mark?' },
                  },
                },
              },
            },
          },
        },
      },
      messages: [
        {
          role: 'user',
          content: [
            { type: 'image', source: { type: 'base64', media_type: 'image/png', data: png.toString('base64') } },
            {
              type: 'text',
              text: `List every selectable row visible in the scrolling list on this screenshot, in top-to-bottom order.

For each row give its label exactly as printed, the centre of the row as fractions of the image (x 0 left, 1 right; y 0 top, 1 bottom), and whether it is SELECTED — a ticked checkbox, check mark or obvious highlight the other rows lack.

Set visible:false for any row clipped at the top or bottom edge of the list, even if you can read its text. A partly visible row cannot be tapped, so reporting it as usable selects nothing at all.

Report only rows of the list itself. Ignore titles, buttons such as CANCEL or CONFIRM, and anything outside the list.`,
            },
          ],
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
    try {
      const p = JSON.parse(text);
      if (!Array.isArray(p.rows)) return [];
      return p.rows
        .filter((r) => r && typeof r.label === 'string' && Number.isFinite(r.y))
        .map((r) => ({
          label: r.label.trim(),
          x: Number.isFinite(r.x) ? r.x : 0.5,
          y: r.y,
          visible: r.visible !== false,
          checked: r.checked === true,
        }));
    } catch {
      return [];
    }
  }

  /**
   * Distil a finished run into a few lines worth keeping.
   *
   * Not a summary of the findings — that is what summarize() is for — but a
   * description of the *app*: what its screens are, what interrupts a scan,
   * which strings are brand names rather than untranslated text. It is handed
   * back as context on the next run, so each scan starts less ignorant than
   * the last.
   */
  async learn({ screens, issues, previousNotes }) {
    const lines = [
      previousNotes ? `What we already believed about this app:\n${previousNotes}\n` : '',
      `Screens captured this run (${screens.length}):`,
      ...screens.slice(0, 40).map((s) => `- ${s.id}: ${s.summary || '(no summary)'}`),
      '',
      `Recurring findings (${issues.length} total):`,
      ...[...new Set(issues.map((i) => `${i.type}: "${String(i.text || '').slice(0, 60)}"`))].slice(0, 40).map((s) => `- ${s}`),
    ].filter(Boolean);

    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 1200,
      output_config: { effort: 'low' },
      system: [{
        type: 'text',
        text: [
          'You maintain a short standing description of a mobile game, used to brief a localization QA agent before it scans the app again.',
          'Write at most 12 short lines. Cover only things that will still be true next run:',
          '- what the main screens are and roughly how they connect',
          '- anything that interrupts a scan (ads, daily rewards, login prompts) and how it is cleared',
          '- strings that look like defects but are not (brand names, deliberate English, stylised text)',
          'Merge with what we already believed; correct it where this run contradicts it.',
          'No preamble, no headings, no markdown. Plain lines only.',
        ].join('\n'),
      }],
      messages: [{ role: 'user', content: [{ type: 'text', text: lines.join('\n') }] }],
    }));

    const text = (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
    return text.slice(0, 2000);
  }

  async summarize(issues, ctx) {
    const byType = {};
    for (const i of issues) byType[i.type] = (byType[i.type] || 0) + 1;
    const top = issues
      .filter((i) => i.severity === 'high')
      .slice(0, 40)
      .map((i) => `- [${i.type}] ${i.screenId}: "${i.text}" — ${i.message}`)
      .join('\n');

    const message = await this._send(() => this._messages.stream({
      ...this._common(),
      model: this.model,
      max_tokens: 4000,
      thinking: { type: 'adaptive' },
      output_config: { effort: 'medium' },
      messages: [
        {
          role: 'user',
          content: `Localization scan of ${ctx.screens} screens in ${ctx.targetHeader}. ${issues.length} issues found.

Counts by type: ${JSON.stringify(byType)}

Highest-severity findings:
${top || '(none)'}

Write a short readout for the localization lead: what is broken, which failures share a root cause, and what to fix first. Lead with the outcome. Plain prose, no headings, under 200 words.`,
        },
      ],
    }).finalMessage());

    this._track(message.usage);
    return (message.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('').trim();
  }
}

module.exports = { ClaudeAnalyzer, ISSUE_TYPES, PRICING };
