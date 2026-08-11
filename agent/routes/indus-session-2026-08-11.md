# Indus device scan — manual pass, 11 August 2026

Recorded so the LocaLinter agent can replay this instead of rediscovering it.
Route map with coordinates: `indus.json`. Everything below was done over adb
against `com.indusgame.dev` 1.12.0 on a vivo I2217 (2400×1080, landscape).

## Why this pass was manual

The agent could not run: the configured API key is rejected by the Anthropic API
(`401 invalid x-api-key`, and `401 Invalid bearer token` for the Bearer form), so
the vision pass had no model behind it. The steps below were driven by hand over
adb with the screenshots read directly, which is exactly what the agent automates
once it has a working model endpoint.

## Environment facts worth keeping

- **No accessibility text.** `uiautomator dump` returns zero text nodes; Unity
  draws into one SurfaceView. Strings can only come from pixels — on dev,
  staging and prod alike. There is no build variant that avoids this.
- **Language is per account**, changed in-game under Settings → Account →
  Language. The device locale stays India/English and has no effect.
- **Prod and dev carry different accounts and languages.** Prod
  (`com.indusgame.play` 2.13.0) was signed in as a level 3 account in
  Portuguese; dev was a level 4 account in English.
- **EC-70001 "LOST CONNECTION"** stranded the dev build; RETRY did not clear it
  while the device had working internet (ping 24 ms). `am force-stop` plus a
  relaunch recovered it. Worth trying before reporting a scan as failed.

## Steps taken

1. `adb devices` → single device, `10BD751PKL00039`.
2. Captured the foreground screen; found the dev build stuck on LOST CONNECTION.
3. Confirmed device connectivity (WiFi validated, ping 8.8.8.8 fine) — so the
   failure was the game's backend, not the network.
4. Tapped RETRY (no effect), then force-stopped and relaunched → lobby.
5. Lobby → settings gear → General, scrolled to Advanced Settings.
6. General → Account → found Region and Language.
7. Language → Arabic → confirmed. Whole UI became tofu boxes. **Known and
   accepted for this build; not a finding.** Skip Arabic on 1.12.0.
8. Language → French → confirmed. Renders correctly; findings below.

## Findings — French, Settings → Account (one screen)

Severity is a first pass; a translator should confirm the wording calls.

| # | Type | String | Detail |
|---|---|---|---|
| 1 | mistranslation | `CHANGEMENT` | Button meaning "Change". `CHANGEMENT` is the noun *a change*; a button needs the verb — `MODIFIER` or `CHANGER`. Appears twice (Region, Language). |
| 2 | mistranslation | `SORTIE DU JEU` | "Exit of the game" as a noun phrase. Should be `QUITTER LE JEU`. |
| 3 | mistranslation | `SOUTIEN` | For a support/help button. `SOUTIEN` is moral or financial support; UI convention is `ASSISTANCE` or `SUPPORT`. |
| 4 | literal translation | `Version de construction` | Literal rendering of "Build Version" — *construction* as in building work. Should be `Version du build`. |
| 5 | untranslated | `Gameplay` | English in the French sidebar. May be an accepted loanword; worth a decision either way. |
| 6 | untranslated value | `French`, `India` | The *values* of Language and Region stay English. Should be `Français` and `Inde`. |
| 7 | overflow risk | `RÉINITIALISER LES PARAMÈTRES` | Wraps to two lines and fills its button to the edges. Compare against the English `RESET SETTINGS`, which fits on one. |
| 8 | overflow risk | `Connectez-vous avec Play Games` | Renders visibly smaller than neighbouring text, suggesting auto-shrink to fit. |

Pattern: 1–4 look like machine translation that was never reviewed by a French
speaker. Worth checking whether the whole French column came from the same pass.

## What the agent should do with this

- Load `indus.json`, drive `procedures.setLanguage` for each language under test,
  then crawl from `lobby`.
- Probe every `infoBadges` entry with a tap and a long-press — those tooltips are
  the flyout text that a coordinate-guessing crawl misses.
- Skip Arabic on 1.12.0 (`knownIssues.arabic`).
- Compare against the **Dev** sheet for `com.indusgame.dev`, and **Global Prod**
  for `com.indusgame.play`. The `.xlsx` in the repo root is not a substitute — it
  holds 310 rows and 16 Portuguese strings, and none of the on-screen strings
  checked against it were present.
