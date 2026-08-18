# Indus device scan — second manual pass, 18 August 2026

Driven by hand over adb against `com.indusgame.play` **2.14.0 #11451000** on a
vivo I2217 (2400×1080, landscape), account `FAUPEHGM`, region India.

Extends the first 18 August pass rather than replacing it. Everything here is
encoded in `indus.json`; this file is the reasoning. **Where the two passes
disagree, this one won and the earlier claim is marked in the route map with a
`$correction` saying what the evidence was** — a stale route map that looks
confident is worse than one that admits a gap.

Languages driven this session: **Portuguese (Brazil) → English → German →
Japanese → English**. The account was left on **English**, which is where it was
before any of this work started.

---

## What this pass was for

Three jobs, in the order the previous session left them:

1. open the screens nobody had opened — Store catalogues, mission
   sub-categories, Weapons sub-collections, the in-match HUD;
2. run a CJK pass and a German pass, the likeliest sources of glyph and
   overflow defects;
3. make the crawler act on the parts of the route map it was ignoring.

All three are done except the in-match HUD, which is deliberately still
unmapped — see the end.

---

## The headline: three recorded facts turned out to be wrong

This is the part worth reading. Each of these had been written down confidently
by an earlier pass, and each was wrong in a way that changed what a scan should
do.

### 1. The language list has nineteen entries, not fifteen

The first pass recorded the list as "enumerated end to end, 15 entries" with
`incomplete: false`. It is **nineteen**. Between INDONESIAN and JAPANESE sit
**TAGALOG, MALAY, VIETNAMESE and THAI** — a whole page that the earlier
enumeration skipped, almost certainly on one un-settled swipe.

The irony is exact: the first pass wrote down `techniques.scrollMomentumSettle`
*because* momentum had made it mis-tap three times, and then fell to the same
problem in its reading rather than its tapping. Momentum corrupts what you
*see*, not just what you *touch*.

Consequence beyond the count: **Thai has never been tested and is now the
build's biggest glyph-coverage risk** — stacked diacritics, a script unlike
anything else shipping, and newly present in the list.

### 2. The Settings "language leak" is not what it was described as

Recorded diagnosis: a fixed set of Settings keys "keeps rendering in the
**previously** selected language."

Tested across four languages in sequence — pt-BR → English → German → Japanese —
the leaked keys rendered **Portuguese every single time**, including two
switches after Portuguese had stopped being the active language. "Previous
language" predicts English at the second step. It did not happen.

Then the decisive test: force-stop, relaunch while on Japanese. The keys came
back **in Japanese** (`設定をリセット`, `ゲーム終了`).

So the real bug is: **these strings are frozen in the language the app process
launched with, and only a restart moves them.** That is a different bug, and a
much cheaper one — and it also proves the translations exist, which the old
description implied they might not.

It is also **wider than Settings**. On that same restart the daily-login reward
item changed from `Green Cloak` to `緑のマント`. Localized *content* is frozen
at launch too, so the old `scope: "settings.*"` was too narrow.

**This has a hard consequence for the scanner**, now encoded as
`hazards.staleLanguageAfterSwitch` and a mandatory restart step in
`procedures.setLanguage`: a run that switches language and captures without
restarting will report stale strings as untranslated defects across the entire
build. Some of the "untranslated" observations in earlier sessions may be
artefacts of exactly this.

### 3. Scroll momentum was not the only thing moving the language list

The first pass blamed its three mis-taps on momentum. Momentum is real, but
there is a second mechanism: **the picker re-anchors to the currently checked
entry a couple of seconds after it settles.** Twice this session a list was
fully settled — two identical frames 2 s apart, the rule the notes prescribe —
a screenshot was taken, and by the time the tap landed the list had jumped back
to the selected language and the tap hit nothing.

Settling is necessary and **not sufficient**. The working method is: swipe,
wait ~2 s for momentum, capture **once**, locate, and tap **immediately**. A
long settle loop actively hurts, because it gives the re-anchor time to fire.
Recorded as `hazards.languagePickerReanchors`.

---

## The English baseline earned its place — mostly by deleting findings

`techniques.englishBaseline` was recorded but never exercised. Running it
changed the outcome of six candidate findings. It **proved two** and **killed
three**, and the three it killed are the more valuable demonstration.

**Proved:**

- *Rewards column word-break.* English `Rewards` sits on one line with room to
  spare. Portuguese renders `Recompen` / `sas` — split mid-word, no hyphen.
  German renders `Belohnung` / `en`, identically broken. Japanese `報酬` fits.
  Two of three localized languages fail, so the fix belongs to the **layout**,
  not to either translation. Best single exemplar in the build.
- *Season banner.* `TEMPORADA TERMINADA` runs the full card width and clips its
  final A; `SEASON ENDED` uses about the left 40%. The earlier note put the
  difference at "~73% wider"; measured against the pair it is nearer 2.2×, and
  it is a real clip rather than "no margin left".

**Killed:**

- *`Limite diário: 1/1` running under the gem art.* Filed in my notes as a
  pt-BR overflow — until English showed `Daily Limit : 1/1` doing exactly the
  same thing. A **source layout bug in every language**, and reporting it
  against pt-BR would have sent the wrong team after it.
- *The blank third row in the redemption sort dropdown.* Looked like an empty
  string key. Present in English. Not a localization defect.
- *"Bundle Store" vs "Loja Premium" vs "Bundles Store".* Looked like a textbook
  pt-BR terminology inconsistency. In English the same chain reads **Bundle
  Store → Premium Store → Bundles Store** — all three differ. A **source**
  naming defect that pt-BR merely inherited; `LOJA PREMIUM` is a correct
  translation of `PREMIUM STORE`.

The narrow pt-BR finding that survives is the one already recorded: the rail
entry `Bundle Store` alone is untranslated. German leaves it untranslated too,
which widens it from a pt-BR miss to a key with **no translation in any locale**.

One more trap worth writing down: **never diff the lobby's third promo card.**
It is a rotating carousel and showed different content on every capture
regardless of language.

---

## Findings by language

### German — 8, including the best single defect of the session

| Severity | String | What is wrong |
|---|---|---|
| high | `SPEICHERN` | The lobby **shop** button reads "to save (data)". The English noun *Store* translated in its verb sense. A German player reads it as a save-game function. The build already uses *Laden* for Store elsewhere, and Japanese renders the same key correctly as 店 — so the key is sound and this one call site is wrong. |
| high | `Belohnung`/`en` | Rewards column word-break (shared with pt-BR; layout fix). |
| med-high | `Um` | The About tab. A preposition meaning "around/at"; wants *Über*. Its seven siblings are all correct. |
| med-high | `BEANSPRUCHEN` | Overflows the lobby claim chip on **both** sides — the leading B and trailing EN sit outside the yellow plate. |
| medium | `Laden einlösen` | Redeem Store as an imperative, "redeem the store". Its rail siblings are noun phrases. |
| medium | `Prismatic Store` | Untranslated in German only; pt-BR has *Loja Prismática*. |
| medium | promo card + bundle body copy | Fall back to English; pt-BR translates both. |
| — | `SAISON BEENDET` | Fits comfortably. Worth noting the German *shortest* string on the screen where Portuguese overflows. |

### Japanese — glyph coverage passes, grammar does not

**Glyph coverage: PASS.** Every kanji and kana renders, no tofu, no fallback
boxes. The CJK font ships. That retires the old Arabic-style concern for JA/KO/ZH
— though **not** for Thai, which is a different script and untested.

| Severity | String | What is wrong |
|---|---|---|
| high | `日目4` … `日目7` | Counter order reversed. Correct Japanese is `4日目` — the numeral **precedes** 日目. The build concatenates `<label><n>` in English "Day 4" order. Repeats on all seven tiles. Fix is a positional placeholder `{n}日目`; no wording change can fix a hardcoded concatenation. |
| medium | `所有` | The "Hold" option on ADS Mode and Quick Chat. Means "to own/possess"; wants 長押し. **Exactly the same defect class as German SPEICHERN** — a short ambiguous English control word taken in the wrong sense. |
| medium | `一般的な` | The General tab: the adjectival form, grammatically incomplete standing alone. Wants 一般. Same class as German `Um`. |
| medium | `常時接続` | Gyroscope "Always On" rendered as "always **connected**", a networking term. |
| medium | `HYBRID` | Untranslated beside two translated siblings. |
| lower confidence | `在庫`, `兵器`, `請求` | Register errors: warehouse *stock* for Inventory, military *ordnance* for Weapons, *invoice/bill* for Claim. Defensible as findings but worth a native reviewer before filing as defects rather than quality notes. |

Two patterns fell out of running German and Japanese together, and both are now
`checkHints`:

- **Short ambiguous English control words are the highest-yield audit in any
  language.** *Store*, *Hold*, *Back*, *Free*, *Match*, *Save*, *Press*, *Round*,
  *Draw*, *Scope*. They produce translations that are competent renderings of a
  *different word*, so no length check or spell check will ever catch them.
- **Standalone nav labels get the wrong part of speech.** German `Um`, Japanese
  `一般的な`. Check every tab against the grammatical form its siblings take.

### Portuguese (Brazil) — new since the first pass

- **`missions.daily` carries two defects on one row.** Besides the known plural
  bug, `Obtenha 1 tiros na cabeça em uma partida em Battle Royale` concatenates
  two prepositional phrases — "in a match" + "in Battle Royale". The sibling row
  shows the clean form. Independently fixable, and a reminder not to stop reading
  a string after the first problem.
- The plural bug is confirmed to affect **count = 1 only**: `2 tiros`,
  `4 inimigos`, `6 kits` are all correct.
- **`PRÓXIMO A ENTRAR`** labels the daily-login countdown and reads as "next in
  line to enter". Wants `PRÓXIMO LOGIN EM`.
- **Redemption sort chip** overflows to two lines and escapes its container;
  English fits easily, German fits but only just.
- **Number formatting**: `₹2,499.00` in a pt-BR UI is anglophone grouping. That
  follows *region*, like currency, so it needs a product decision rather than a
  bug. What *is* reportable without one is the inconsistency on a single screen:
  grouped prices (`₹4,999.00`) beside ungrouped quantities (`x 12000`).

---

## Screens opened this pass

**Store — all seven catalogues.** `featured` (one hero card), `offers` (4×2,
scrolls vertically, five packs), `gems` (3×2, six SKUs, almost no translatable
text), `bundleStore` → `bundleDetail` behind VIEW, `prismatic` (two skins),
`redemption` (the densest screen in the build: sort chip, type rail, grid,
detail panel), `creatorCode`.

The old route map blocked `store.*` wholesale, which would have kept a scan out
of all of it — the most commercially visible text in the build, and the source
of several findings above. `blocked.labels` now names only the purchase confirms.

**Missions — all four sub-categories.** `event` (5), `daily` (4, plus a
milestone track the others lack), `weekly` (2), `loginDaily` (a reward track,
not a mission list).

**Weapons — all three sub-collections.** `arsenal` (class rail + horizontally
scrolling weapon strip + per-weapon detail; descriptors are correct pt-BR),
`evoX` (skin evolution), `melee`.

**`dailyLoginRewards` — a screen the route map did not have at all**, and the
single most disruptive thing on the device. It covers the lobby after launch
**and after every language change**, and it silently ate three separate
navigation sequences this session: the taps landed on the modal, I believed I
was in Settings, and the next captures were of the wrong screen. A crawler would
have captured, analysed and reported those with complete confidence.

---

## The in-match HUD is still unmapped, deliberately

Reaching it needs `PLAY`, which `blocked.labels` forbids for good reasons: it
commits the scan account to a live match with real players and an abandon
penalty, and the crawl has no way out of it. Asked and confirmed: leave it, and
record why. It needs a supervised match run by hand, not an autonomous crawl.

---

## What changed in the code

`lib/crawler.js` now acts on three parts of the route map it had been carrying
and ignoring, plus one that was quietly broken.

- **`autoDismiss`** — modals the map marks dismiss-on-sight are closed with
  their own close control instead of being explored. Crucially this replaces
  pressing **back**, which in this app opens the exit-game dialog: the generic
  "press back until the overlay goes away" routine would eventually quit the
  game and end the run. `dismissOverlays` now refuses back entirely when a
  hazard says so.
- **`scrollable`** — the generic probe swipes vertically up the middle of the
  screen, which does nothing to a weapon strip that scrolls sideways along the
  bottom or a rewards rail narrower than the swipe. Recorded regions are now
  swiped along their own axis, inside their own bounds.
- **`englishBaseline`** — source-language captures are written to disk keyed by
  route screen (`lib/baseline.js`) and compared when a later run recognises the
  same screen. A cheap textual pass finds strings identical to the source; a
  paired-image call finds text that no longer fits. The comparison prompt is
  written as much around **not** reporting things — anything wrong in *both*
  images belongs to the source, and carousels and timers differ for reasons that
  have nothing to do with language.
- **`procedures.setLanguage` was broken** and would have silently skipped every
  run. It still looked for per-language coordinates that the first pass had
  deliberately removed when it discovered the picker scrolls. It now locates the
  row by reading the screen, taps immediately (the re-anchor), reads the
  checkbox back before confirming, and **restarts the app afterwards**.

`npm test` covers the route-map logic that can be checked without a device:
entry resolution including the four newly found languages, modal identification
and that `exitConfirm` dismisses via **NO**, scroll axes, the baseline
round-trip, and that the blocked list still stops purchases while letting the
Store catalogues through.

---

## Still open

- **Thai** — untested, newly discovered, highest glyph risk in the build.
- Korean and both Chinese variants — the font covers CJK, but the `日目4`
  counter-order defect is exactly the kind that repeats per-language.
- The in-match HUD.
- The Japanese register calls (`在庫`, `兵器`, `請求`) want a native reviewer.
- `Pacote LOUD Classics` / `LOUD BUNDLE` — still needs the explicit product
  decision on whether LOUD product names stay English. `RASTROS DE GLÓRIA` can
  now be treated as settled: it stays Portuguese under English, German and
  Japanese alike, which reads as deliberate branding for a Brazilian-org season.
