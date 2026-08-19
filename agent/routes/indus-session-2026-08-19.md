# Indus device scan — Thai pass, 19 August 2026

Driven by hand over adb against `com.indusgame.play` **2.14.0 #11451000** on a
vivo I2217 (2400×1080, landscape), account `FAUPEHGM`, region India. Same build
and same device as both 18 August passes.

Languages driven this session: **English (baseline) → Thai → Japanese (control)
→ Thai**. The account was left on **Thai**, which is a change from how previous
sessions left it — see "State the device was left in" at the end.

This pass had one job, the one the route map itself named as the highest-value
next target: **drive Thai**, the language that was newly discovered in the
second 18 August pass, never tested, and flagged as the build's biggest
remaining glyph-coverage risk.

That question is now answered. But the more important finding is not about Thai
at all — it is about the recorded `procedures.setLanguage`, which is unsafe in a
way that can silently invalidate an entire scan.

---

## Headline: the mandatory restart can silently throw away the language

`procedures.setLanguage` ends with `action: "restart"` marked **MANDATORY**,
added by the second 18 August pass to defeat `hazards.staleLanguageAfterSwitch`.

That restart is not safe. **The account language does not reliably survive a
process restart.**

Four restarts were performed this session under a controlled protocol:

| # | Language set | Settle before restart | Language after restart |
|---|---|---|---|
| 1 | Thai | ~18 s | **reverted to English** |
| 2 | Thai | 60 s | **reverted to English** |
| 3 | Japanese | 60 s | **reverted to English** |
| 4 | Thai | ~5 min of normal navigation | **held as Thai** |

Network was healthy throughout and was checked explicitly after the first
reverts — WiFi validated, 0% packet loss, 24 ms RTT to 8.8.8.8 — so this is not
a failed server fetch caused by a dead link.

**It is intermittent, and that is worse than a deterministic bug.** A scan that
follows the recorded procedure will, some fraction of the time, restart, come
back in English, and then capture the entire build believing it is looking at
the target language. Every string would be "untranslated". The run would look
successful and be entirely worthless.

The one restart that held was also the one preceded by several minutes of
ordinary navigation (lobby, then Missions) rather than a timed wait. That is a
hypothesis, not a conclusion: 60 seconds of sitting still was not enough twice,
and roughly five minutes of moving around was enough once. One trial does not
establish the mechanism, and it should not be encoded as if it did.

### What to do about it — the cheap decisive guard

The fix is not to pick a side between "restart is mandatory" and "never
restart". Both hazards are real and they pull in opposite directions:

- skip the restart and `staleLanguageAfterSwitch` leaves parts of the UI frozen
  in the launch language;
- perform the restart and the language itself may be gone.

The guard that resolves both is to **stop trusting either and read the value
back**: after the restart, navigate to Settings → Account and confirm the
`Language` row still shows the target before capturing anything. It is one
screen, it is unambiguous, and it converts a silent catastrophic failure into a
loud recoverable one.

This is now `procedures.setLanguage.steps[].verifyLanguageApplied` and
`hazards.languageMayNotSurviveRestart`.

### This partly rewrites the 18 August conclusion

The second 18 August pass wrote: "force-stop, relaunch while on Japanese. The
keys came back in Japanese (`設定をリセット`, `ゲーム終了`)."

Those are the *leaked* keys — the small frozen set. Seeing them in Japanese
after a relaunch shows the leak model is right; it does **not** show that the
account language survived, because the leaked keys were never the thing in
question. Under this session's evidence that relaunch may well have come back
on English with Japanese leaked keys, which is exactly the confusing state the
device was found in at the start of today's session (see below).

The `staleLanguageAfterSwitch` model itself is **confirmed, not overturned** —
see the next section.

---

## The stale-language leak reproduced exactly, on arrival

The device was found with the app already running, `Language: English`, and
`サポート`, `変化`, `設定をリセット`, `ゲーム終了` and both sign-in buttons still
rendering in **Japanese**. A restart cleared all of it to English.

That is `hazards.staleLanguageAfterSwitch` reproducing precisely as recorded:
the process had launched under Japanese, the account was later moved to English,
and the frozen key set never followed. Screenshot kept as
`evidence-stale-ja-change-button.png`.

It also surfaced a Japanese defect that three passes had walked past, because it
is only visible while the leak is active:

- **`変化` for the Region / Language "CHANGE" button.** `変化` is change in the
  sense of *transformation, mutation* — something that happens to a thing. The
  button performs a change, which is `変更`. Confirmed against the same screen
  in English (`CHANGE`) so this is the Japanese rendering of that key, not an
  artefact. Same defect class as German `SPEICHERN` and Japanese `所有`.

---

## Thai: glyph coverage PASSES

**This retires the build's biggest recorded glyph risk.**

Every Thai string rendered completely — no tofu, no fallback boxes, no dropped
marks. That specifically includes the stacked diacritics that made Thai a risk
in the first place:

- `ทั่วไป` (General) — mai ek above sara a
- `ความไวต่อความรู้สึก` (Sensitivity) — sara uu with mai tri stacked
- `รีเซ็ตการตั้งค่า` (Reset Settings) — mai taikhu, and mai tho above sara a
- `อ้างสิทธิ์ทั้งหมด` (Claim All) — thanthakhat and a stacked ั้
- `ฆ่าศัตรู … ด้วย …` — mai ek and mai tho mid-sentence

The Thai font ships and composes correctly. Combined with the CJK pass on 18
August, the remaining untested scripts are Korean and both Chinese variants
(same CJK font, low risk), and Vietnamese, whose stacked tone+vowel marks are
the one script family still genuinely unproven.

**A caution that is now a technique.** Two candidate "missing tone mark" findings
were raised at full-screen resolution and **both died when cropped and
upscaled** — `ฤดูกาลสิ้นสุดแล้ว` (Season Ended) and `อ้างสิทธิ์ทั้งหมด` are
correctly spelled and correctly fitted; the marks were simply below the
resolution I was judging at. Thai diacritics are small, high-contrast and sit
right at the cap line, which makes them the single easiest thing in this build
to hallucinate a defect about. `checkHints.cropBeforeCalling` is not optional
for Thai — it is the difference between a real finding and a fabricated one.

---

## Thai findings

### The big one: `เก็บ` for STORE — a third language, same broken key

The lobby's **Store** tab reads `เก็บ`, the verb *to keep / to put away / to
collect*. A shop is `ร้านค้า`.

This is the **exact same defect as German `SPEICHERN`** — the English noun
*Store* taken in its verb sense — now confirmed in a third language. Japanese
renders the same key correctly as `店`. So:

- the key is sound and the source string is retrievable;
- German and Thai both fail it the same way;
- it is the highest-yield single key in the build for a terminology audit.

### `Claim` is mistranslated in Thai, twice, differently

| Where | Thai | Means | Should be |
|---|---|---|---|
| Lobby claim chip | `เรียกร้อง` | to demand, to call for (a right) | `รับ` / `รับรางวัล` |
| Missions "Claim All" | `อ้างสิทธิ์ทั้งหมด` | to assert a legal right | `รับทั้งหมด` |

Two separate wrong-sense renderings of the same English word in the same build,
so this is a terminology-consistency defect **on top of** the sense error.
Japanese independently fails the same key as `請求` (invoice / bill). Three
languages, three different wrong senses, one English word — `Claim` belongs on
`checkHints.ambiguousControlWords` as a worked example.

### `สีเหลียม` for QUAD — wrong sense *and* misspelled

The mode chip reads `สีเหลียม | ไลต์เฮาส์ / แบทเทิลรอยัล`.

Two independent defects in one short string:

1. **Wrong sense.** `สี่เหลี่ยม` means *quadrilateral / square* — the geometric
   shape. QUAD here is a four-player squad, which is `ทีม 4 คน` or `สควอด`.
2. **Misspelled.** As shipped it is `สีเหลียม`, missing the mai ek on both
   syllables. `สี` alone means *colour*. The correct form is `สี่เหลี่ยม`.

This is a source-string typo, not a rendering failure: the same font renders mai
ek correctly in `ทั่วไป`, `ต่อ`, `ฆ่า` and `ด้วย` on other screens. Verified by
cropping at 4×.

### Mission templates insert English spacing into a spaceless script

Every mission row spaces its interpolated tokens:

> `ฆ่าศัตรู 1 ด้วย โชคชะตา ใน แบทเทิลรอยัล`
> `เล่น 12 แมตช์ของ มินิ TDM`
> `สร้างความเสียหายรวม 1200 ใน แบทเทิลรอยัล`

**Thai does not put spaces between words.** The template is composed as
`<verb> {n} <with> {item} <in> {mode}` with English-style separators, and the
spaces survive into Thai. Native form: `ฆ่าศัตรู 1 คนด้วยโชคชะตาในแบทเทิลรอยัล`.

The first row also drops the classifier — Thai counts people with `คน`, so
`ฆ่าศัตรู 1` should be `ฆ่าศัตรู 1 คน`.

This is the Thai instance of the pt-BR `missionTemplateComposition` defect and
it affects **every mission row**, not one string. No length check or spell check
will catch it.

### `Event Missions` untranslated — and it is the active tab

The Missions rail reads `Event Missions` in English beside three correctly
translated siblings (`ภารกิจประจำวัน`, `ภารกิจรายสัปดาห์`, `เข้าสู่ระบบรายวัน`).
It is the tab selected on entry, so it is the first thing a Thai player reads on
the screen.

### Smaller Thai items

- **`สร้างเวอร์ชัน` for "Build Version"** — `สร้าง` is the verb *to build /
  create*. The row shows a version number, so it wants `เวอร์ชันบิลด์` or plain
  `เวอร์ชัน`. Same class as `เก็บ` / `SPEICHERN` / `所有`.
- **`ความไวต่อความรู้สึก` for "Sensitivity"** — literally *sensitivity to
  feeling*, the emotional sense. Control sensitivity is `ความไว`. It is also the
  longest item in the settings rail and comes close to the panel edge, though it
  does not clip.
- **`Green Cloak` untranslated** on the daily-login modal, where Japanese has
  `緑のマント`. A Thai-only gap on a key that demonstrably has translations.
- **`LOUD Classics Pack` / `THE GOAT PASS` untranslated** — consistent with the
  existing brand-string findings, no new decision needed.
- **`กราฟฟิก` for "Graphics"** — low confidence. The Royal Institute form is
  `กราฟิก` with one `ฟ`; the doubled form is common in the wild. Worth a native
  reviewer, not worth filing.

### Negative results worth keeping

Both of these are defects that repeat across languages and specifically **do
not** repeat in Thai:

- **The Japanese day-counter order bug does not occur.** Japanese renders `日目4`
  with the numeral in the wrong position; Thai renders `วันที่ 4` … `วันที่ 7`
  in correct order. The `<label><n>` concatenation happens to be right for Thai.
- **The pt-BR season-banner overflow does not occur.** `ฤดูกาลสิ้นสุดแล้ว` fits
  the card comfortably where `TEMPORADA TERMINADA` clips its final A.

---

## The language picker: the recorded technique is wrong and cost three mis-taps

`hazards.languagePickerReanchors` says the list re-anchors to the checked entry
a couple of seconds after settling, and prescribes: "swipe, wait ~2 s, capture
**once**, locate, and tap **immediately**. A long settle loop actively hurts."

Racing the re-anchor does not work, and this session lost three taps to it:

1. scrolled to THAI, waited 2 s, captured — the list had already snapped back to
   ENGLISH at the top;
2. computed THAI's position from the geometry and tapped — hit **MALAY**;
3. tapped where THAI was visible in the previous capture — the list re-anchored
   between capture and tap, and the tap hit nothing.

**The correct technique is the opposite of the recorded one: do not race the
re-anchor, wait for it to finish.** The re-anchor fires after a *scroll*, not
after a tap, and once it has fired the list is stable — a tap three seconds
later landed correctly and a deliberate 3 s wait afterwards did not move it.

Recorded as a rewritten `hazards.languagePickerReanchors` plus
`techniques.pickerTwoHopSelection`:

- The re-anchor parks the **checked** entry at the top of the viewport, so a
  distant target cannot be reached by scrolling — the list keeps pulling back.
- Tapping a row only moves the checkbox. **Nothing is applied until CONFIRM**,
  so an intermediate selection is free.
- Therefore: tap any row *near* the target, let the list re-anchor onto it, then
  tap the target from the now-stable list. Two hops, both verified.

Also corrected: **`scrollable.region`'s top of 0.24 is too generous.** A row
clipped by the viewport's top edge renders but does not accept taps — a tap at
y≈0.27 on a half-visible THAI row did nothing at all. Taps should be confined to
rows fully inside the region.

`verifyChecked` earned its place twice more. Both mis-taps were silent: the
wrong language was selected with no visual cue anywhere except the checkbox
itself. Without the read-back this session would have confirmed MALAY and then
attributed every Malay string to Thai.

---

## State the device was left in

**The account was left on Thai**, not English. Previous passes deliberately
restored English; this one did not, because the last restart was the trial that
established the persistence behaviour and reverting it would have destroyed the
observation.

Anyone picking this up should either set it back to English or account for it.
Given the intermittent revert, **read the value rather than assuming it** — that
is the same guard this session added to the procedure.

---

## Still open

- **Vietnamese** — now the only script family with genuinely unproven rendering;
  stacked tone-plus-vowel marks, same risk profile Thai had before today.
- Korean and both Chinese variants — CJK font is proven, but the Japanese
  day-counter defect is the kind that repeats per-language and is unchecked.
- Tagalog, Malay, Indonesian, Turkish, Russian, Italian, both Spanish variants,
  Portuguese (Portugal) — never driven.
- **The persistence mechanism.** Four restarts is enough to prove the revert is
  real and intermittent; it is not enough to explain it. If it turns out to be
  time- or activity-dependent the procedure can be tightened; until then the
  read-back guard covers it.
- The in-match HUD, still deliberately unmapped.
- The Japanese register calls (`在庫`, `兵器`, `請求`) still want a native
  reviewer, and Thai now adds `กราฟฟิก` to that list.
