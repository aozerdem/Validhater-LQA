# PE Pattern Findings (MT vs Post-Edit diff, 232 segments)

Source: PE-PT-3050063-260608110511 — cols SourceSegment / OriginalTargetSegment /
Post-EditingTargetSegment. 184 of 232 segments were edited (TER avg 0.40).
Data is reliable-but-noisy: treat as corroborating evidence, not ground truth.

---

## Mechanical fixes — deterministic pre-check candidates (roadmap improvement #5)

### number + unit spacing (HIGHEST PRIORITY)
- Pattern: "\d%" in MT → "\d %" in PE
- Count: 9 occurrences. Triple-confirmed (goldset, production report, PE data).
- Extends to ALL units: mm, cm, m, kg, g, ml, l, W, V, A, mAh, kB, MB, GB, °C
- Example: "100% bomull" → "100 % bomull"; "14mm" → "14 mm"
- Pre-check regex: re.search(r'\d(mm|cm|m|km|kg|g|ml|l|W|V|A|mAh|kB|MB|GB|%|°C)\b', target)

### en-dash in numeric ranges
- Pattern: "\d-\d" in MT → "\d–\d" in PE
- Count: 6 occurrences.
- Pre-check regex: re.search(r'\d-\d', target) → flag if source has a numeric range

### decimal separator
- Pattern: "\d.\d" (period) in MT → "\d,\d" (comma) in PE
- Count: 2 occurrences. Less frequent but confirmed.
- Exception: version numbers keep period (ver. 4.2) — check context before flagging.

### MT duplication artifacts
- Pattern: adjacent identical token pairs e.g. "enkelt og enkelt" (easily and easily)
- Count: low but present. Core SG section 6.1 prohibits unintentional duplication.
- Pre-check: scan for identical adjacent content words (length > 3) in target.

### Bracketed section label translation (4/4 confirmed, 100% consistent)
- Pattern: source has [English Label] → MT translates it → PE reverts to English
- Examples: [Design Description]→[Designbeskrivelse] (wrong); PE reverts to [Design Description]
- Pre-check: extract bracket content from source and target; flag if they differ.
- NOTE: this is already added to the SYSTEM_PROMPT in Task 2 of this handoff.

---

## Term / idiom corrections — confirm existing lexicon entries

- "vennligst" removed as unidiomatic — 5 occurrences. Matches goldset. In prompt.
- "designet" (calque of "designed") replaced with "laget"/"lages" — 3 occurrences. Added to prompt in Task 2.
- Recurring removed MT tokens (weak candidates — noisy signal): materiale, teppe, kontakt.

---

## Structural / register edits (noted, deliberately NOT encoded as hard rules)

- LLs restructured marketing-register product titles for Norwegian naturalness,
  departing from strict source-mirroring (e.g. umbrella title fully rewritten).
- TDT Core SG says preserve source element order; LLs override this for idiomaticity
  on customer-facing-feeling segments. This tension is real but encoding it would
  spike false-flags. Stays at WARN level (advisory). Do not add as a hard rule.
- "er underlagt den virkelige tingen" → "avhenger av det virkelige produktet":
  MT calqued "is subject to the real thing"; LL rephrased naturally. Unidiomatic
  pattern, not a lexicon-level fix.

---

## Edge case — DO NOT over-translate broken source tokens

- Source: "Upper Material:COULD" (broken product attribute, "COULD" is a material code)
- MT translated "COULD" to "kan" (Norwegian for "could")
- LL PE reverted to "COULD" — garbage/metadata source tokens must be mirrored, not translated
- This is now in the SYSTEM_PROMPT (Task 2, Edit 2.2). Logged here for the pre-check
  task: a deterministic check could flag when an apparent English source token is absent
  from the target (potential over-translation of source noise).

---

## Caveat

Some MT==PE unchanged rows still contain untranslated English (e.g. "Brand Name:tomorrow-today"),
meaning the PE pass itself has missed errors. Confirms PE data is evidence, not ground truth.
