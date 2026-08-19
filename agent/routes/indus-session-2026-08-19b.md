# Indus device scan — Vietnamese pass, 19 August 2026

Driven against `com.indusgame.play` **2.14.0 #11451000** on the vivo I2217,
account `FAUPEHGM`, region India. Same build and device as the Thai pass earlier
the same day.

Vietnamese was the last script the route map still called unproven: stacked
tone-plus-vowel marks, the same risk profile Thai had before it was tested and
passed.

Half of this pass was driven by hand and half through the desktop app, which is
now model-led — the crawler proposes what to tap and the route map only vetoes,
names and tops up.

---

## Glyph coverage: PASSES

**The last script risk is retired.** Vietnamese stacks two marks on one vowel far
more often than Thai does, and every combination on the lobby rendered
completely — no tofu, no dropped marks, no fallback boxes:

| String | Meaning | The hard part |
|---|---|---|
| `NHIỆM VỤ` | Missions | `Ệ` — circumflex *and* dot below |
| `ĐƯỢC XẾP HẠNG` | Ranked | `Ợ` horn + dot below, `Ế` circumflex + acute |
| `CỬA HÀNG` | Store | `Ử` horn + hook above |
| `KHẢ NĂNG` | Abilities | `Ả` hook above, `Ă` breve |
| `HÀNG TỒN KHO` | Inventory | `Ồ` circumflex + grave |
| `KHẲNG ĐỊNH` | Claim | `Ẳ` breve + hook, `Ị` dot below, `Đ` stroke |
| `TRẬN CHIẾN SINH TỒN` | Battle Royale | `Ậ` circumflex + dot below |
| `Cấp độ 3` | Level 3 | `ấ` and `ộ` in body text at small size |

With Thai proven earlier today and CJK proven on 18 August, **no shipping script
in this build is now untested for rendering.** What remains everywhere is
wording, not glyphs.

---

## The headline: three defects that repeat across languages

Vietnamese is the second or third language to fail the same three English keys,
which moves all of them from "a translator got this wrong" to "the source string
is ambiguous and every language is guessing".

### 1. QUAD is being read as the geometric shape — now in two languages

The mode chip reads `TỨ GIÁC | LIGHTHOUSE`. **`tứ giác` is a quadrilateral**, the
four-sided polygon. QUAD here is a four-player squad, which is `ĐỘI 4` or the
borrowed `SQUAD`.

Thai renders the identical error: `สี่เหลี่ยม`, also the geometric shape. Two
unrelated languages reaching the same wrong noun is not two translator mistakes —
it is a four-letter English string with no context attached to it.

### 2. Claim is being read as *assert a right* — now in four languages

| Language | String | What it means |
|---|---|---|
| Vietnamese | `KHẲNG ĐỊNH` | to affirm, to assert (a statement) |
| Thai (lobby) | `เรียกร้อง` | to demand, to call for (a right) |
| Thai (missions) | `อ้างสิทธิ์` | to assert a legal right |
| Japanese | `請求` | to invoice, to bill |

Not one of them means *collect your reward*. Vietnamese wants `NHẬN` or
`NHẬN THƯỞNG`. Four languages, four different wrong senses, one English word —
this is the strongest single argument in the whole route map for splitting or
annotating the source key.

### 3. Inventory is being read as warehouse stock — confirmed, not suspected

`HÀNG TỒN KHO` is the accounting term: goods remaining in a warehouse. It is what
an ERP screen says. A player's item bag is `TÚI ĐỒ` or `KHO ĐỒ`.

The 18 August pass found Japanese `在庫` — the same warehouse sense — and could
only file it as a register question wanting a native reviewer, because one
language reading a word oddly is weak evidence. **A second, unrelated language
making the identical substitution settles it.** The Japanese finding should be
upgraded from a quality note to a defect.

---

## Vietnamese findings

### `ĐƯỢC XẾP HẠNG` for RANKED — wrong part of speech

`được xếp hạng` is a passive verb phrase, "to be ranked". The mode is a noun:
`XẾP HẠNG`. The stray `ĐƯỢC` turns a menu item into a sentence fragment.

Same class as German `Um` for About and Japanese `一般的な` for General — a
standalone nav label taking the grammatical form it would have inside a sentence.
That pattern is now confirmed in four languages and is already recorded as
`checkHints.standaloneNavLabels`; this is its cleanest example yet, because the
siblings around it (`NHIỆM VỤ`, `CỬA HÀNG`, `VŨ KHÍ`) are all correctly nouns.

### `KHẢ NĂNG` for ABILITIES — defensible, lower confidence

`khả năng` is ability in the abstract sense — capability, possibility. Game
abilities are usually `KỸ NĂNG`, skills. Worth a native reviewer rather than
filing: unlike the three above there is no second language corroborating it.

### `LIGHTHOUSE` untranslated — and inconsistently so

The map name is left in English here. Thai transliterated the same name as
`ไลต์เฮาส์`. Neither is wrong on its own, but the pair shows there is no decision
recorded anywhere about whether map names are translated, transliterated or
left — so each language has made its own. That is a product decision to make
once, not a defect to file per language.

### Correct, and worth recording as such

These were checked and are right, so that a later pass does not re-raise them:

- **`CỬA HÀNG` for STORE is correct.** Notable because this is the key German
  (`SPEICHERN`) and Thai (`เก็บ`) both break by taking *Store* as a verb.
  Vietnamese gets it right, which confirms the key itself is sound and the
  source string is retrievable.
- `NHIỆM VỤ` (Missions), `VŨ KHÍ` (Weapons), `MỜI` (Invite), `THÊM BẠN BÈ` (Add
  Friends), `CHƠI` (Play), `Cấp độ` (Level), `MÙA GIẢI KẾT THÚC` (Season Ended),
  `Giới hạn hàng ngày` (Daily Limit) — all correct and all fitting their
  containers.
- **`MÙA GIẢI KẾT THÚC` fits the season banner comfortably**, where pt-BR
  `TEMPORADA TERMINADA` clips its final A. Another language that does not
  reproduce `knownFindings.seasonBannerOverflowPt`.
- The third promo card was **not** compared across languages. It is the rotating
  carousel that `knownFindings` warns never to diff, and it showed different
  content again here.

---

## What the run says about the tooling

This was the first language pass driven through the model-led crawler rather
than the recorded map, and it behaved as intended: the model proposed the
controls, the route map vetoed the dangerous ones and added what the model had
not seen.

One weakness showed up immediately and is worth fixing. **The model repeatedly
proposes the lobby's back arrow**, which on this screen opens the exit-game
dialog. The dialog is caught and answered NO every time — `exitConfirm` is marked
dismiss-on-sight — so nothing breaks, but each round trip costs an action and two
vision calls for no text. `hazards.backOnLobbyExits` already records why back is
unsafe here; the veto should extend to a *proposed tap on a back control while
the lobby is the current screen*, not only to the back key itself.

---

## Still open after this pass

- **No script risks remain.** Korean and both Chinese variants share the proven
  CJK font; every other shipping language is Latin, Thai or Vietnamese, all now
  exercised.
- Korean and the Chinese variants have still never been driven for *wording* —
  and given that QUAD, Claim and Inventory each break in most languages tested,
  they are likely to break there too.
- Tagalog, Malay, Indonesian, Turkish, Russian, Italian, both Spanish variants
  and Portuguese (Portugal) — never driven.
- The in-match HUD, still deliberately unmapped.
- `KHẢ NĂNG` and the Japanese register calls (`在庫` now upgraded, `兵器`, `請求`)
  want a native reviewer.

---

## What the automated scan added — Settings, and it is worse in there

The hand pass covered the lobby. The model-led scan got into Settings and found
a denser cluster than anything on the lobby, all of it the same defect class.
Four were verified by cropping before being written down.

### `VỠ NHẸ` for TAP — the worst string in the build

`vỡ` is to break or shatter; `nhẹ` is light or slight. Together: *lightly
shattered*. The English is **TAP**, the touch gesture.

It is the **selected** option on three separate rows — ADS Mode, QuickChat and
Emote — so it is the first thing a Vietnamese player reads in Settings, three
times. `GIỮ` for Hold beside it is correct, so this is not a row that was
mistranslated wholesale; it is one string. Wants `CHẠM` or `NHẤN`.

### `CHỈ QUẢNG CÁO` for ADS ONLY — the acronym was read as advertisements

On the Gyroscope row: `TẮT | CHỈ QUẢNG CÁO | LUÔN BẬT` — Off / **Ads only** /
Always on. `quảng cáo` means advertising. ADS here is aim-down-sights.

The detail that makes it filable rather than arguable: **the row directly above
keeps `Chế độ ADS` as an acronym.** The same three letters are treated as an
acronym in one row and expanded to "advertisements" in the next, on one screen.

### `ỦNG HỘ` for SUPPORT, `GIẢI PHÓNG` for RELEASE

- `ủng hộ` is to endorse, back or donate to something. The button opens a help
  desk. Wants `HỖ TRỢ`.
- `giải phóng` is to liberate or free up — the sense used for freeing memory, not
  for letting go of a button.

### `Về` for ABOUT — the standalone-nav-label pattern again

`về` on its own is the preposition *about / regarding*, grammatically incomplete
as a menu item. Wants `Giới thiệu` or `Thông tin`.

That makes four languages now failing this same pattern: German `Um`, Japanese
`一般的な`, and Vietnamese twice — `Về` here and `ĐƯỢC XẾP HẠNG` on the lobby.

### Terminology drift: "Input Block" four ways on one screen

The scan reported the same key rendered four different ways with inconsistent
word order on a single screen. Not a wrong-sense error — a consistency one, and
the sort that only shows up when a whole screen is read at once.

### The rest, as reported

5 untranslated strings (including `India`), 2 mixed-language strings, 10 strings
on screen that are not in the sheet at all, 3 overlapping elements and 1 vertical
overflow.

The ten out-of-sheet strings are worth a look on their own account: they are
either content the sheet does not cover or keys that have drifted, and neither is
visible from the spreadsheet side.

---

## The tooling fault this pass found and fixed

The first attempt got three screens in ten minutes and stalled. The model
proposed the lobby's **back arrow** every pass; back on the lobby opens the
exit-game dialog; `exitConfirm` answered NO; the crawl returned to the lobby and
proposed the back arrow again. Nothing broke and nothing progressed.

The model cannot see this — a back arrow looks like a back arrow, and only the
map knows the lobby is the root. `screens.lobby.backExits` now records it and a
proposal to go back from such a screen is refused. The re-run vetoed it on the
first screen and went straight into Settings, which is where everything above
was found.
