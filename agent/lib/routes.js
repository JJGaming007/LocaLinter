'use strict';

// Route maps live in agent/routes/*.json. Each one is what an earlier pass
// learned about a game: which packages it ships under, where its screens are,
// which "?" badges open flyout tooltips, how to change language, and how to
// get out of a stuck state. Serving them lets a scan start from knowledge
// instead of rediscovering the same ground by vision on every run.

const fs = require('fs');
const path = require('path');

const ROUTES_DIR = path.join(__dirname, '..', 'routes');

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

module.exports = { list, get, packageFor, ROUTES_DIR };
