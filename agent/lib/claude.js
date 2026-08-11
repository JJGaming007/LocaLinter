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
  constructor({ apiKey, model = 'claude-opus-5', effort = 'high', baseUrl = '' }) {
    if (!apiKey) throw new Error('No Anthropic API key configured.');
    // A base URL points the SDK at a company gateway (LiteLLM and friends)
    // that speaks the same /v1/messages format. Empty means Anthropic direct.
    this.client = new Anthropic({ apiKey, ...(baseUrl ? { baseURL: baseUrl } : {}) });
    this.baseUrl = baseUrl || '';
    this.model = model;
    this.effort = effort;
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

  /** Short natural-language wrap-up over the whole run. */
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
