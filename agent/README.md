# LocaLinter Device Scan — local agent

Crawls a running build of your game, screen by screen, and compares every string
it finds against the localization sheet already loaded in LocaLinter. Works
against an Android device over ADB and against the Unity Editor in Play Mode.

Everything runs on your machine. The agent binds to `127.0.0.1` only, and your
Anthropic API key never reaches the browser.

## Setup

```bash
cd agent
npm install
npm start
```

Then open LocaLinter, load your sheet, and switch to the **Device Scan** tab.
Paste your Anthropic API key once — the agent stores it in `agent/config.json`,
which is gitignored. `ANTHROPIC_API_KEY` in the environment also works.

Requirements: Node 18+, `adb` on your PATH (or point at it in Advanced), USB
debugging enabled on the device.

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
| Scroll through lists | on | Steps every ScrollRect (bridge only) |
| Long-press for tooltips | on | Long-presses info, help, and "?" controls |
| Android package | — | Lets the agent restart the app when it cannot find its way back |

## Output

Findings appear live in the Device Scan tab, grouped by screen with the
screenshot beside them, and can be exported as JSON. Every run also writes
`agent/runs/<run-id>/` with `report.json` and a PNG per screen.

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
