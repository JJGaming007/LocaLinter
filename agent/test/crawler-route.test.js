'use strict';
process.env.LOCALINTER_DATA_DIR = require('os').tmpdir() + '/localinter-selftest';

const assert = require('assert');
const { Crawler } = require('../lib/crawler');
const baseline = require('../lib/baseline');
const route = require('../routes/indus.json');

const logs = [];
const run = { log: (m) => logs.push(String(m)), warnings: [], skipped: [], emit() {}, issues: [] };
const cfg = { blockedLabels: [], probeLabels: [], scrollSteps: 3, settleMs: 10, scrollProbe: true };

function makeCrawler(targetHeader = 'Portuguese (Brazil)') {
  const c = new Crawler({
    cfg, adb: null, bridge: null, sheet: null, analyzer: null, run, route,
    target: { header: targetHeader, code: 'pt-BR', rtl: false, sourceHeader: 'English', expectsNonLatin: false },
  });
  c.deviceSize = { width: 2400, height: 1080 };
  return c;
}

// ── 1. picker entry resolution ────────────────────────────────────────────
{
  const c = makeCrawler();
  assert.strictEqual(c.pickerEntryFor('Portuguese (Brazil)'), 'PORTUGUESE (BRAZIL)');
  assert.strictEqual(c.pickerEntryFor('portuguese (brazil)'), 'PORTUGUESE (BRAZIL)');
  assert.strictEqual(c.pickerEntryFor('German'), 'GERMAN');
  assert.strictEqual(c.pickerEntryFor('Thai'), 'THAI', 'Thai must resolve — it was missing from the old list');
  assert.strictEqual(c.pickerEntryFor('Tagalog'), 'TAGALOG');
  assert.strictEqual(c.pickerEntryFor('Klingon'), null, 'an unknown language must be refused, not guessed');
  console.log('1 ok  picker entries resolve, including the four that were missing');
}

// ── 2. auto-dismiss recognition ───────────────────────────────────────────
{
  const c = makeCrawler();
  const daily = c.identifyRouteScreen(['DAILY LOGIN REWARDS', 'Green Cloak', 'CLAIM']);
  assert.strictEqual(daily && daily.name, 'dailyLoginRewards');
  assert.ok(daily.def.autoDismiss, 'the daily-login modal must be marked dismiss-on-sight');
  assert.ok(c.resolveRef(daily.def.autoDismiss.tap), 'its close control must resolve to a point');

  const exit = c.identifyRouteScreen(['EXIT GAME', 'Are you sure you want to exit']);
  assert.strictEqual(exit && exit.name, 'exitConfirm');
  assert.ok(exit.def.autoDismiss);
  const p = c.resolveRef(exit.def.autoDismiss.tap);
  assert.ok(Math.abs(p.x - 0.434 * 2400) < 1, 'exitConfirm must dismiss via NO, not YES');
  console.log('2 ok  modals identified and their close controls resolve');
}

// ── 3. back is refused where the route says it is unsafe ──────────────────
{
  const c = makeCrawler();
  assert.strictEqual(c.hazardWarnsAgainstBack(), true, 'Indus flags back as unsafe');
  const bare = new Crawler({
    cfg, adb: null, bridge: null, sheet: null, analyzer: null, run,
    route: { screens: {}, hazards: { something: 'nothing to do with navigation' } },
    target: { header: 'X', sourceHeader: 'English' },
  });
  assert.strictEqual(bare.hazardWarnsAgainstBack(), false, 'an unrelated hazard must not disable back');
  console.log('3 ok  back is refused on Indus and allowed elsewhere');
}

// ── 4. scrollable regions come off the route map ──────────────────────────
{
  const c = makeCrawler();
  c.routeScreenFor.set('s1', { name: 'weapons.arsenal', def: route.screens['weapons.arsenal'] });
  const regions = c.routeScrollRegions('s1');
  assert.ok(regions.length, 'arsenal must offer a scroll region');
  assert.ok(regions.some((r) => r.axis === 'horizontal'), 'the weapon strip scrolls sideways');

  c.routeScreenFor.set('s2', { name: 'store.bundleDetail', def: route.screens['store.bundleDetail'] });
  const rewards = c.routeScrollRegions('s2');
  assert.ok(rewards.some((r) => r.axis === 'vertical'));

  assert.deepStrictEqual(c.routeScrollRegions('unknown'), [], 'an unrecognised screen falls back to the generic probe');
  console.log('4 ok  scroll regions and axes read from the route map');
}

// ── 5. baseline round-trip ────────────────────────────────────────────────
{
  const c = makeCrawler('English');
  assert.strictEqual(c.capturingBaseline, true, 'a run in the source language records the baseline');
  assert.strictEqual(c.routeName, 'Indus');

  const t = makeCrawler('Portuguese (Brazil)');
  assert.strictEqual(t.capturingBaseline, false);

  baseline.record('Indus', 'English', 'store.bundleDetail', {
    texts: ['Rewards', 'BUY FULL PACKAGE', '20% OFF', 'Offer Valid Till :'],
    png: Buffer.from('not-a-real-png'),
  });
  const got = baseline.get('Indus', 'English', 'store.bundleDetail');
  assert.ok(got && got.texts.includes('Rewards'));
  assert.ok(got.png, 'the screenshot must survive the round trip');

  // The strings that stayed English are the untranslated candidates.
  const identical = baseline.untranslatedCandidates(
    got.texts,
    ['Recompensas', 'BUY FULL PACKAGE', '20% OFF', 'Oferta válida até:']
  );
  assert.ok(identical.includes('BUY FULL PACKAGE'));
  assert.ok(identical.includes('20% OFF'));
  assert.ok(!identical.includes('Rewards'), 'a translated string is not a candidate');
  assert.ok(!identical.includes('Offer Valid Till :'), 'a translated string is not a candidate');
  console.log('5 ok  baseline round-trips and finds the real untranslated cluster');
}

// ── 6. noise is not mistaken for untranslated text ────────────────────────
{
  const same = ['1/1', 'x 60', '₹29.00', 'OK', '2.14.0', 'LOUD'];
  const out = baseline.untranslatedCandidates(same, same);
  assert.deepStrictEqual(out, ['LOUD'], 'numbers and short tokens are not findings; a word is a candidate');
  console.log('6 ok  numeric and short strings are not reported as untranslated');
}

// ── 7. the blocked list still guards the dangerous taps ───────────────────
{
  // Built the way a real run builds it: the config's blocked labels only. The
  // route map's own list has to be picked up by the Crawler itself. This test
  // used to translate route.blocked.labels into cfg here, which is precisely
  // the wiring production was missing — so it passed while a live scan happily
  // tapped PLAY.
  const c = new Crawler({
    cfg, adb: null, bridge: null, sheet: null, analyzer: null, run, route,
    target: { header: 'German', sourceHeader: 'English' },
  });
  // Blocked: the things that end the run or destroy the account it runs on.
  assert.strictEqual(c.isBlocked('lobby.play'), true, 'PLAY commits the account to a live match');
  assert.strictEqual(c.isBlocked('exitConfirm.yes'), true);
  assert.strictEqual(c.isBlocked('settings.resetSettings'), true);
  assert.strictEqual(c.isBlocked('settings.exitGame'), true);

  // Not blocked: spending money. The account is a test one with a test card,
  // and the purchase and confirmation dialogs carry some of the most
  // commercially visible text in the build — the screens a localization scan
  // most wants to read. Blocking them was hiding them.
  assert.strictEqual(c.isBlocked('store.offers.purchase'), false, 'purchases are in scope on a test account');
  assert.strictEqual(c.isBlocked('lobby.currencyAdd'), false, 'the currency top-up is a purchase entry point');
  assert.strictEqual(c.isBlocked('store.bundleDetail.pricePlate'), false);
  assert.strictEqual(c.isBlocked('Buy'), false, 'a button reading "Buy" is no longer refused on sight');

  // ...and the Store catalogues stay reachable.
  assert.strictEqual(c.isBlocked('store.catGems'), false, 'store categories must be explorable');
  assert.strictEqual(c.isBlocked('store.bundleStore.view'), false);
  console.log('7 ok  run-ending taps blocked, purchases and catalogues open');
}

// ── 8. the language switch reads its own result back ──────────────────────
//
// The restart at the end of setLanguage reverts the account to English
// intermittently — three of four trials on 2026-08-19. When it happens the run
// carries on and reports every string in the build as untranslated, and nothing
// about the run looks wrong. The read-back is the only thing standing between
// that and a confidently wrong report, so it is pinned here rather than left to
// whoever next edits the procedure.
{
  const steps = route.procedures.setLanguage.steps;
  const verify = steps.find((s) => s.verifyLanguageApplied);
  assert.ok(verify, 'setLanguage must read the language back after the restart');
  assert.strictEqual(verify.onMismatch, 'abort', 'a wrong language must abort, never warn-and-continue');

  const restartAt = steps.findIndex((s) => s.action === 'restart');
  assert.ok(restartAt >= 0, 'the restart step must still exist');
  assert.ok(steps.indexOf(verify) > restartAt, 'the read-back is only meaningful after the restart');

  // The picker mis-selects silently, so the checkbox read-back matters too.
  assert.ok(steps.some((s) => s.verifyChecked), 'setLanguage must verify the checkbox before confirming');

  assert.ok(route.hazards.languageMayNotSurviveRestart, 'the persistence hazard must stay recorded');
  console.log('8 ok  the language switch verifies the checkbox and reads the language back');
}

// ── 9. Thai is tested, and the next target is named ───────────────────────
{
  const entries = route.screens['settings.languagePicker'].entries;
  assert.ok(!entries.untested.includes('THAI'), 'Thai was driven on 2026-08-19');
  assert.strictEqual(route.knownIssues.thaiGlyphCoverage.status, 'PASS');
  assert.ok(entries.untested.includes('VIETNAMESE'), 'Vietnamese is the remaining script risk');
  assert.deepStrictEqual(route.knownIssues.thaiGlyphCoverage.remainingScriptRisk, ['VIETNAMESE']);
  console.log('9 ok  Thai retired as a glyph risk, Vietnamese named as the next one');
}

// ── 10. the model decides, the route map vetoes ───────────────────────────
//
// The order matters as much as the outcome. Recorded controls used to go first
// and displace anything the model proposed near them, which made the crawl a
// hand-drawn map with the model filling gaps. Now the proposals lead and the
// map only refuses, renames and tops up — so this checks both that PLAY cannot
// get through under any label, and that the model's own choices stay in front.
{
  const c = makeCrawler('German');
  c.routeScreenFor.set('screen-001', { name: 'lobby', def: route.screens.lobby });

  const index = c.routeControlIndex('screen-001');
  const play = index.find((k) => k.ref === 'lobby.play');
  const missions = index.find((k) => k.ref === 'lobby.missions');
  assert.ok(play && missions, 'the index must include controls, blocked ones included');

  const proposals = [
    { key: 'a', kind: 'tap', x: play.x, y: play.y, label: 'PLAY', priority: 'high' },
    { key: 'b', kind: 'tap', x: missions.x, y: missions.y, label: 'Missions button', priority: 'high' },
    { key: 'c', kind: 'tap', x: 5, y: 5, label: 'something only the model saw', priority: 'medium' },
  ];
  const out = c.applyRouteToProposals(proposals, 'screen-001');
  const labels = out.map((a) => a.label);

  assert.ok(!labels.includes('PLAY'), 'a proposal landing on PLAY must be refused');
  assert.ok(!out.some((a) => a.label === 'lobby.play'), 'and not readmitted under its route name');
  assert.ok(labels.includes('lobby.missions'), 'a proposal on a known control takes the route ref as its name');
  assert.ok(labels.includes('something only the model saw'), 'proposals with no counterpart survive');
  assert.strictEqual(labels[0], 'lobby.missions', 'the model\'s own proposals stay at the front');
  assert.ok(
    out.slice(2).every((a) => a.fromRoute),
    'controls the model missed are appended behind its proposals, not ahead of them'
  );
  console.log('10 ok  the model proposes, the route map vetoes and tops up');
}

console.log('\nall self-tests passed');
