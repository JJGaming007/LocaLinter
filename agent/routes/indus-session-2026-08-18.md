# Indus device scan — manual pass, 18 August 2026

Driven by hand over adb against `com.indusgame.play` **2.14.0 #11451000** on a
vivo I2217 (2400×1080, landscape), account `FAUPEHGM`, region India.
Everything learned here is encoded in `indus.json`; this file is the reasoning
behind it.

Supersedes the 11 Aug pass where they disagree — that one ran against dev
1.12.0 and several of its coordinates and assumptions are now stale.

## What changed since 11 August

- **The language picker is a scrolling list, and it grew.** 1.12.0 was recorded
  as English / Arabic / French / Russian at fixed coordinates. 2.14.0 ships 15
  languages and no Arabic — full list below. The old per-language coordinates
  were removed rather than corrected: a scrolling list has no fixed coordinates,
  so entries must be located by reading the screen.
- **Settings is one screen with a sticky tab**, not several. It reopens on
  whichever tab was last used.
- `setLanguage` ends on **settings.account**, not the lobby, and takes ~15s.

## The known Settings issue, characterised

Switching the account language does not fully re-render Settings: a fixed set of
keys keeps rendering in the *previously* selected language.

With the account on **English**, Settings showed `SUPORTE`, `MUDAR` ×2,
`SAIR DO JOGO`, `CONFIGURAÇÕES BÁSICAS`, `REDEFINIR CONFIGURAÇÕES`,
`SELECIONE O IDIOMA`, `Entrar com o Play Games`, `Entrar com o Google` — all
Portuguese — while every other string on the same screens was English.

Switching to Portuguese rendered **all of them correctly**. That is the proof it
is a refresh bug and not a missing translation. The 11 Aug French pass hit the
same keys (`CHANGEMENT`, `SORTIE DU JEU`, `SOUTIEN`, `PARAMÈTRES DE BASE`), so
it is the same key set every time.

Per the 18 Aug decision these are **still reported**, tagged `known: true` at low
severity and collapsed — see `knownIssues.settingsLanguageLeak`, and
`markKnownIssues()` in `lib/crawler.js`. Anything in Settings *not* in that list
is a new defect and reports normally.

## How to actually get the strings out (the expensive lesson)

A crawler that screenshots each screen once sees a small fraction of this build's
text. Four techniques, all confirmed on device:

1. **Grid detail panels.** Inventory and Abilities show only the *selected*
   item's name and lore. Adjacent avatar cells gave `BRUNO PLAYHARD / From a
   mobile gaming channel…` and `ADI SERIES / Customizable Avatar…`. 19 cells are
   now walked deliberately.
2. **Row description panels.** Settings → Gameplay describes only the selected
   row. `Auto Sprint` → "Enables continuous sprinting…"; `Cosmium Trigger` →
   "Tap the button to interact with Cosmium." 7 rows now selected in turn.
3. **Info-badge flyouts.** The `?` beside an ability name opens a HOW TO USE
   panel with three lines that *occlude* the description underneath. Capture
   before and after, then dismiss.
4. **Scroll-momentum settle.** Lists keep gliding after the swipe. Aiming at
   PORTUGUESE (BRAZIL) selected GERMAN, then SPANISH (MEXICO), on two separate
   attempts. Rule: take two screenshots ~2s apart, proceed only when identical,
   then re-locate the target in the *fresh* screenshot.

Plus **englishBaseline**: capture each screen in English and in the target
language. The season-banner overflow below is only provable from the pair.

## Findings — Portuguese (Brazil), prod 2.14.0

| # | Severity | Type | String | Detail |
|---|---|---|---|---|
| 1 | high | grammar | `Obtenha 1 tiros na cabeça em uma partida` | Count substituted into a template fixed in the plural. `1 tiros` = singular number, plural noun. Should be `1 tiro`. Affects every mission whose count can reach 1. |
| 2 | high | untranslated | `Bundle Store` | The only English entry in the Store's category rail, beside Destaques, Ofertas, Gemas, Loja Prismática, Loja de Resgate, Código de Criador. The sibling pattern is `Loja de …`. |
| 3 | medium | terminology | `Linguagem` vs `SELECIONE O IDIOMA` | The label and the picker it opens, two taps apart, use different words for the same concept. pt-BR convention is *Idioma*. |
| 4 | medium | terminology | `RESGATAR` vs `REIVINDIQUE TUDO` | One action, two verbs, one screen apart (lobby claim vs missions claim-all). |
| 5 | medium | untranslated | `Portuguese (Brazil)`, `India` | Row labels translate (Região, Linguagem); their **values** stay English. Should be `Português (Brasil)`, `Índia`. The French pass found the identical pattern. |
| 6 | medium | untranslated | language picker entries | Every entry is English regardless of UI language. A language picker is the one screen that must use endonyms — a player who cannot read the current language has to find their own. |
| 7 | medium | overflow | `TEMPORADA TERMINADA` | ~73% wider than `SEASON ENDED`; runs under the character art to the season card's right edge. Not clipped on this device, but it has no margin left. |
| 8 | low | partial | `Pacote LOUD Classics` | Half-translated promo card. Probably fine if `LOUD Classics` is a product name — needs an explicit decision. |
| 9 | low (verify) | terminology | `RANQUEADAS` vs `CLASSIFICAÇÃO` | Lobby button vs destination title. May be a distinction the English makes too — check the English pair before reporting. |

Not defects, confirmed: prices render as `₹29.00` because the account **region**
is India while the UI language is Portuguese — currency follows region, not
language. Lobby chat lines and player names are user content and are never
localized.

## Screens mapped

`lobby`, `settings` (+ `general`, `gameplay`, `account`, `languagePicker`),
`inventory`, `abilities`, `missions`, `ranked`, `store`, `weapons`,
`promoInterstitial`, `exitConfirm` — 14 in all, every coordinate tapped and
confirmed on the device.

`weapons` is a **modal chooser** over the dimmed lobby (Arsenal / Evo-x /
Corpo a Corpo), not a leaf screen. `ranked` is empty between seasons
(`NOVA TEMPORADA EM BREVE!`) — that is the game's state, not a failed capture.

## The language list, in full

Enumerated end to end. **15 entries**, in this order:

ENGLISH · FRENCH · RUSSIAN · GERMAN · ITALIAN · SPANISH (SPAIN) ·
SPANISH (MEXICO) · PORTUGUESE (BRAZIL) · PORTUGUESE (PORTUGAL) · TURKISH ·
INDONESIAN · JAPANESE · KOREAN · CHINESE (TRADITIONAL) · CHINESE (SIMPLIFIED)

**Arabic is gone.** It was offered in 1.12.0 and rendered as tofu boxes; it is
not in 2.14.0 at all. `knownIssues.arabicTofu` is therefore retired rather than
stale — there is nothing left to test or suppress.

Note the four CJK entries: those are the ones to test next for glyph coverage
and for the overflow that long German and Russian strings usually cause.

## A hazard worth knowing

**Android BACK on the lobby opens the exit-game confirmation** (`SAIR DO JOGO`
/ "Tem certeza de que deseja sair do jogo?"). Its SIM button quits the app. This
was hit during recording: a back-press meant to close the Weapons modal fell
through to the lobby, opened the dialog, and the next three taps landed on it
instead of the intended screen.

A crawler that uses BACK to navigate up **will** reach this. Prefer the
in-screen back chevron at `[0.035, 0.076]`, and after any back-press check for
the `exitConfirm` signature before tapping anything else. Recorded as
`screens.exitConfirm` (with `autoDismiss`) and `hazards.backOnLobbyExits`.

## Still open

- Store category catalogues, mission sub-categories (Diárias / Semanais / Login
  Diário), the Weapons sub-collections (Arsenal / Evo-x / Corpo a Corpo) and the
  in-match HUD were not opened.
- CJK and German/Russian passes not run — the likeliest sources of glyph and
  overflow defects.
- The account was left on **Portuguese (Brazil)**; it was English at the start.
