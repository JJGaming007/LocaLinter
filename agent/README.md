# LocaLinter Device Scan — local agent

Crawls a running build of your game, screen by screen, and compares every string
it finds against the localization sheet already loaded in LocaLinter. Works
against an Android device over ADB and against the Unity Editor in Play Mode.

Everything runs on your machine. The agent binds to `127.0.0.1` only, and your
Anthropic API key never reaches the browser.

## Why this runs on your PC and not on a server

Your test device is plugged into a USB port on your machine. ADB talks to it
over that cable, so the thing driving the scan has to be on the same machine —
a server in a data centre has no cable to your phone. That is why every tester
runs their own copy of the agent, and why it never needs to be reachable from
the network.

## For testers

1. Get `LocaLinter-Agent.exe` and run it. A console window opens and stays
   open — that *is* the agent, so leave it alone while you scan.
2. Open LocaLinter, sign in, load your sheet, and go to the **Device Scan** tab.
   It should say **Connected**. If not, press **Reconnect**.
3. Paste your Anthropic API key once. It is stored on your machine only.
4. No `adb`? The panel says so and offers a **Download ADB** button, which
   fetches Google's platform-tools for you. Nothing to install by hand.

You need USB debugging enabled on the device. You do not need Node, the Android
SDK, or a checkout of this repo.

Windows will likely warn about an unrecognised app the first time, because the
executable is unsigned — *More info → Run anyway*. Signing it with a code
signing certificate is the only way to remove that prompt.

The agent keeps config, route maps and run output in
`%LOCALAPPDATA%\LocaLinter\agent`. Delete that folder to start clean.

## From a source checkout

```bash
cd agent
npm install
npm start
```

Identical behaviour, except the data directory is `agent/` itself, so an
existing `config.json` and `runs/` stay exactly where they were. Requirements:
Node 18+ (20.12+ to build the executable), `adb` on your PATH or downloaded
through the panel. `ANTHROPIC_API_KEY` in the environment works instead of
pasting a key. `LOCALINTER_DATA_DIR` overrides where state is kept, and `PORT`
moves it off 8790.

## Building the executable

```bash
cd agent
npm run build:exe        # -> agent/dist/LocaLinter-Agent.exe (~87 MB)
```

Bundles the agent with esbuild, embeds the route maps under `routes/` as an
asset, and injects the result into a copy of the Node binary you built with
(Node's Single Executable Application support). Windows only, and it produces an
executable for the machine it ran on. `postject` prints a *"signature seems
corrupted"* warning — expected, since injecting into a signed `node.exe`
invalidates Node's signature.

Hand the resulting file to testers however suits you — a GitHub release asset,
a shared drive. Set `AGENT_DOWNLOAD_URL` at the top of `device-scan.js` to that
location and the Device Scan panel will link to it directly.

## The Unity bridge — add this

`unity/LocaLinterBridge.cs` is the difference between a good scan and an
excellent one. Copy it anywhere under your project's `Assets/` folder. Nothing
else changes: it self-installs at runtime, and it is wrapped in
`#if UNITY_EDITOR || DEVELOPMENT_BUILD`, so it cannot ship in a release build.

Without it, the agent reads text out of screenshots. With it, the agent gets:

| | Screenshots only | With the bridge |
|---|---|---|
| Strings | read visually | exact, straight from the Text/TMP components |
| Truncation & overflow | judged by eye | measured — `isTextTruncated`, `preferredWidth` vs the box |
| Element positions | approximate | exact screen rects |
| Clicking | tap at a guessed coordinate | click the actual control by id |
| Scrolling lists | swipe and hope | drive each ScrollRect through its full range |
| Screen identity | perceptual image hash | the set of strings and controls on screen |

On device the agent reaches the bridge through `adb forward tcp:8791`, set up
automatically. In the Editor, just press Play.

The bridge listens on `127.0.0.1:8791`; override with the `LOCALINTER_PORT`
environment variable.

## How a scan works

1. **Capture** the current screen: screenshot plus, if the bridge is up, every
   visible string with its rect, font, and truncation state.
2. **Check** mechanically against the sheet — untranslated strings, strings from
   the wrong language column, broken placeholders, tokens like `{0}` leaking to
   the player, reversed RTL text, tofu glyphs, overflow, overlap, text pushed off
   screen, strings that appear nowhere in the sheet.
3. **Look** at the screenshot with Claude for what measurement cannot catch:
   clipping behind art, collisions with icons, unreadable contrast, layouts that
   should mirror in RTL but do not, mistranslations, and text baked into
   textures. Claude is also asked to list any on-screen text the engine did not
   report, and those strings go back through step 2.
4. **Explore.** Every control on the screen is queued. Each new state — a menu, a
   popup, a dropdown with its options open, a tooltip from a long-pressed info
   badge — is its own screen and gets the full treatment. Scroll views are
   stepped through so text below the fold is not missed.

Coverage is state-based, not screen-based, which is why flyouts and dropdowns
are covered as thoroughly as top-level menus.

## Safety

The agent never taps a control whose label matches a blocked pattern. The
default list covers purchases, subscriptions, account deletion, progress resets,
and logout, so a scan against a live build with a real payment method on file
cannot spend money. Edit the list under **Advanced → Never tap controls
matching**; anything skipped is listed in the run report so you know where
coverage stops.

## Settings

Set in the UI, stored in `agent/config.json`.

| Setting | Default | Notes |
|---|---|---|
| Max screens / actions / depth | 120 / 400 / 12 | Raise for a full sweep, lower for a smoke test |
| Model | `claude-opus-5` | `claude-sonnet-5` is faster and cheaper per screen |
| Effort | `high` | `xhigh` for the most thorough reading |
| Claude vision pass | on | Off = mechanical checks only, no API cost |
| Scroll through lists | on | Steps every ScrollRect with the bridge; swipes until the screen stops changing without it |
| Long-press for tooltips | on | Long-presses info, help, and "?" controls |
| Android package | — | Lets the agent restart the app when it cannot find its way back |

## Output

Findings appear live in the Device Scan tab, grouped by screen with the
screenshot beside them, and can be exported as JSON. Every run also writes
`agent/runs/<run-id>/` with `report.json` and a PNG per screen.

## What a scan costs

Each run reports `usage` — input, output, and cached tokens, and the number of
API calls — in `report.json` and in the run log. It is a measurement, not an
estimate. The app itself no longer tracks or displays a running spend figure;
read the token counts against your own Anthropic billing if you need dollars.

Safety classifiers occasionally decline a screenshot. Rather than losing that
screen, the agent asks the API to fall back to another model for the declined
request; if the key cannot use that feature, it stops asking and carries on.

## Troubleshooting

**"No devices found"** — check `adb devices`; accept the USB debugging prompt on
the phone.

**"No in-game bridge"** — the scan still runs on screenshots alone. To fix: the
build must be a development build with `LocaLinterBridge.cs` in the project, and
the game must be in the foreground.

**Every string is reported as untranslated** — the game is running in a different
language than the column you selected. The agent warns about this when the
bridge can read the game's locale; otherwise switch the language in-game first.

**The crawler keeps landing on the wrong screen** — set the Android package name
so it can restart the app to re-anchor.
