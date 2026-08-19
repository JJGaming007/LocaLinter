# LocaLinter

Localization QA for mobile games. Two halves that share a UI:

1. **Sheet linting** — load a localization spreadsheet (Google Sheets or a local
   file) and report format, placeholder and coverage problems. Deterministic,
   instant, no model involved.
2. **Device Scan** — drive the real game on a real phone over adb, read every
   screen, and report what a player in that language would actually see. This is
   the part under active development.

## The goal

**The AI model does the testing.** It looks at the screen, decides what to tap,
reads the text, compares it to the sheet and judges what is wrong — the way a
human tester would. It is not a scripted crawler with the model filling gaps.

That direction was chosen after a scripted crawler was tried and kept failing
silently: recorded coordinates assume the app is where it was when someone wrote
them, and when it is not, every tap after the first miss lands somewhere
arbitrary while the procedure still reports success.

**What stays scripted, and why.** Three things a screenshot cannot tell you, all
of them in the route map:

- **veto** — PLAY commits the account to a live match with an abandon penalty;
  back on the lobby opens the exit dialog. Nothing in the pixels says so.
- **known non-defects** — so the same accepted quirk is not re-reported forever.
- **safety read-backs** — the account language reverts on restart intermittently,
  so it is confirmed from Settings whichever way the switch was driven.

Everything else — where to tap, what screen this is, what the text says, whether
it matches the sheet — is the model's job.

## Layout

```
agent/            the scanning service (Node, no build step)
  server.js       HTTP API on 127.0.0.1:8790; also serves the UI
  lib/claude.js   every model call lives here
  lib/crawler.js  the scan loop: drive, capture, analyse
  lib/zoom.js     crop + magnify a string for a second look
  lib/sheet.js    the spreadsheet index (retrieval, not judgement)
  lib/checks.js   deterministic checks — needs the Unity bridge, so it
                  contributes nothing on a SurfaceView title like Indus
  lib/paths.js    where config, route maps and runs live
  routes/         route maps + the session notes behind them
desktop/          Electron shell; runs the agent in-process
index.html        the UI, served by the agent
device-scan.js    the Device Scan panel
```

## Running it

```bash
cd desktop && npm install && npx electron .
```

The window opens on `http://127.0.0.1:8790`. To drive it programmatically, add
`--remote-debugging-port=9222` and attach with `playwright-core`
(`chromium.connectOverCDP`) — there is no other way to click its UI headlessly.

Needs, all per-user and gitignored:

- `agent/google-client.json` — OAuth client, or the sign-in gate cannot be passed
- an API key, saved through the panel (Advanced → Base URL first if it is a
  gateway key, or it fails the `sk-ant-` check)

`npm test --prefix agent` runs the self-tests. They are fast and they are the
only tests; run them before committing.

## Model calls per screen

| Call | When |
|---|---|
| `analyzeScreen` | always — read the text, judge it |
| `proposeTargets` | always — what is worth tapping |
| `reconcileWithSheet` | always — what the sheet says about each string |
| `identifyScreen` | when the recorded signatures do not match |
| `decideNextAction` | once per step of `pursue()` |
| `verifyFinding` | per rendering finding, capped at `zoomVerifyMax` |
| `readListRows`, `locateText` | list navigation, recovery |

About 36 calls and 160k input tokens for a four-screen run. `modelSheetCompare`,
`modelDrivenNavigation` and `zoomVerify` each turn off if a run needs to be
cheaper.

**Not model calls:** the sheet lint, custom rules, screen deduping via perceptual
hash, adb, and the route-map veto.

## How the pieces fit

- `pursue(goal)` — screenshot, decide one action, act, look again. Replaces
  recorded step lists. Nothing is assumed between steps.
- `applyRouteToProposals` — the model proposes; the map vetoes, renames and
  appends what was missed. The order matters: proposals lead.
- `reconcileScreenWithSheet` — retrieval by index, judgement by model. Send the
  candidates' **similarity**; without it a 46% guess reads like an exact hit.
- `confirmUnderMagnification` — the nine rendering issue types get cropped and
  re-examined. Semantic findings do not; magnification cannot answer them.

## Working on this

- **Run it against the device before believing it works.** Every serious bug this
  project has had was invisible to the tests and obvious on the first real run.
- **A wrong finding costs more than a missed one.** The first bad entry is where
  a translator stops believing the report.
- **"The control did not respond" and "the app forbids this" look identical from
  outside.** Vary the gesture before concluding anything about the app.
- Route maps are seeded into the per-user data dir. An untouched copy is
  refreshed on launch; an edited one is kept.

## State, 20 August 2026

Branch `desktop-redesign-and-agent-training`.

**Working, verified on device:** model-led control selection; model-driven
navigation (it set the language to German from a Vietnamese UI in 12 steps with
no coordinates); model sheet comparison; zoom verification; route-map veto
(refused 11 taps in one run, answered the exit dialog NO); language read-back
after restart.

**Known, not fixed:**

- The account language reverts to English on restart intermittently — 3 of 4
  trials. The read-back catches it and stops the run; the cause is unknown.
- `checks.js` contributes nothing on Indus and every deterministic finding comes
  from the sheet comparison instead.
- The Device Scan panel's blockers list keeps stale entries in a hidden
  container. Harmless, invisible to users.

**Next:**

- Korean and both Chinese variants have never been driven. Script risk is gone
  (Thai and Vietnamese both passed, CJK passed 18 Aug); the reason to run them is
  that Quad, Claim and Inventory break in most languages tested so far.
- The in-match HUD is unmapped. It needs PLAY, which is vetoed, so it wants a
  supervised run rather than a crawl.
- Purchase flows were unblocked but have not been scanned yet. The test account
  uses a test card, and the confirmation dialogs are wanted.
- `pursue()` currently drives the language switch only. The crawl itself still
  picks actions from a queue; giving it a goal is the obvious next step.
