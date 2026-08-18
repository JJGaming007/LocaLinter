'use strict';

// Route maps live in the agent's data directory, under routes/*.json (that is
// agent/routes from a source checkout — see lib/paths.js). Each one is what an
// earlier pass
// learned about a game: which packages it ships under, where its screens are,
// which "?" badges open flyout tooltips, how to change language, and how to
// get out of a stuck state. Serving them lets a scan start from knowledge
// instead of rediscovering the same ground by vision on every run.

const fs = require('fs');
const path = require('path');

const { ROUTES_DIR } = require('./paths');

function readAll() {
  let files = [];
  try {
    files = fs.readdirSync(ROUTES_DIR).filter((f) => f.endsWith('.json'));
  } catch {
    return [];
  }
  const out = [];
  for (const file of files) {
    try {
      const raw = fs.readFileSync(path.join(ROUTES_DIR, file), 'utf8');
      const data = JSON.parse(raw);
      out.push({ name: path.basename(file, '.json'), data });
    } catch (e) {
      console.warn(`[routes] skipping ${file}: ${e.message}`);
    }
  }
  return out;
}

// A compact shape for the picker: enough to describe a route without shipping
// every coordinate to the browser.
function summarise({ name, data }) {
  const app = data.app || {};
  const screens = Object.keys(data.screens || {});
  const badges = Object.values(data.screens || {})
    .reduce((n, s) => n + Object.keys(s.infoBadges || {}).length, 0);
  return {
    name,
    label: app.name || name,
    packages: app.packages || {},
    engine: app.engine || null,
    screens: screens.length,
    screenNames: screens,
    infoBadges: badges,
    procedures: Object.keys(data.procedures || {}),
    capabilities: data.capabilities || {},
    knownIssues: Object.entries(data.knownIssues || {})
      .filter(([k]) => k !== '$comment')
      .map(([k, v]) => ({ key: k, note: v })),
    recordedOn: app.recordedOn || null,
  };
}

function list() {
  return readAll().map(summarise);
}

function get(name) {
  const found = readAll().find((r) => r.name === name);
  return found ? found.data : null;
}

// Which package a route expects for a given environment, so a scan can check
// it is pointed at the build the route was recorded against.
function packageFor(data, pkg) {
  const packages = (data && data.app && data.app.packages) || {};
  const hit = Object.entries(packages).find(([, v]) => v === pkg);
  return hit ? hit[0] : null;
}

/**
 * Turn a route map into context for the vision pass.
 *
 * A route map is not only coordinates: it is what a person learned by playing
 * the build with their own hands — which strings this app renders with markup,
 * which mismatches are already accepted, which defects a human confirmed, and
 * which states end a run. None of that reached the model before, so every run
 * re-derived it (or, more often, did not) and reported accepted bugs as new.
 *
 * Coordinates are deliberately left out; the crawler acts on those, the model
 * only needs the judgement.
 */
function promptContext(data) {
  if (!data) return '';
  const NL = String.fromCharCode(10);
  const parts = [];
  const app = data.app || {};

  const section = (heading, entries, render) => {
    const rows = Object.entries(entries || {}).filter(([k]) => k.charAt(0) !== '$');
    if (!rows.length) return;
    parts.push(heading + NL + rows.map(render).join(NL));
  };

  if (app.name) {
    const rec = app.recordedOn || {};
    parts.push(
      'This app is ' + app.name + (rec.build ? ' (recorded against build ' + rec.build + ')' : '') + '. ' +
      'A person walked this build by hand and wrote down the following. Treat it as established fact about the app, not as a guess.'
    );
  }

  section(
    'How this app renders text, and what that means for comparing it to the sheet:',
    data.checkHints,
    ([k, v]) => '- ' + k + ': ' + v
  );

  section(
    'Already known and accepted by the team. Still report these if you see them — they are tagged and collapsed downstream — but do NOT treat them as new discoveries, and never let them crowd out the rest of your findings:',
    data.knownIssues,
    ([k, v]) => '- ' + k + ': ' + ((v && v.note) || '')
  );

  section(
    'Real defects a human already confirmed in this build. They are genuine and NOT suppressed — if you see them, report them; if you see the same pattern elsewhere, report that too. Each ends with what it teaches about where such bugs hide:',
    data.knownFindings,
    ([k, v]) => {
      const where = v.screen ? ' [' + v.screen + ']' : '';
      const teaches = v.teaches ? ' — Generalise: ' + v.teaches : '';
      return '- ' + k + where + ' (' + (v.severity || 'medium') + ', ' + (v.type || 'issue') + '): ' + (v.detail || '') + teaches;
    }
  );

  section(
    'States that can end or derail a run. If a screenshot shows one of these, say so plainly in your notes — it explains why later screens may be missing:',
    data.hazards,
    ([k, v]) => '- ' + k + ': ' + v
  );

  section(
    'How the text on these screens is reached. If a screenshot looks sparse it is usually because one of these was not done — say so rather than concluding the screen has little text:',
    data.techniques,
    ([k, v]) => '- ' + k + ': ' + v
  );

  return parts.join(NL + NL);
}

module.exports = { list, get, packageFor, promptContext, ROUTES_DIR };
