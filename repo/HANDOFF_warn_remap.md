# HANDOFF — WARN Remap + System Prompt Updates + PE Pattern Notes

> **For Claude Code.** Read CLAUDE.md first. Three tasks in order.
> Task 1 is code-only (calibrate.py). Task 2 is prompt-only (va_evaluator.py SYSTEM_PROMPT).
> Task 3 is a notes file — no code. Do not mix tasks or add features not listed here.

---

## TASK 1 — Remap WARN as advisory in calibrate.py

### Why

The first calibration run used this binary mapping to compare against the LL's Yes/No:

- OK            → predicted "Yes"
- WARN or FAIL  → predicted "No"

6 of the 13 false-flags were WARN-level (score 96–97, "slightly stiff phrasing") —
advisory observations that a Language Lead would never actually block a segment over.
Lumping WARN with FAIL dragged precision down to 74.5%. WARN should be advisory:
it appears in the report but does not count as "should not have been published."

### Change in calibrate.py

In the binary-verdict derivation step, change the rule to:

- OK   → predicted "Yes"
- WARN → predicted "Yes"   ← advisory, does not block
- FAIL → predicted "No"

The raw `eval_severity` and `eval_score` columns in the report stay unchanged.
Only the classification logic changes.

### Four-class labels (same definitions, new binary underneath)

- TRUE_PASS:  LL Yes, predicted Yes  (OK or WARN)
- TRUE_FLAG:  LL No,  predicted No   (FAIL only)
- FALSE_PASS: LL No,  predicted Yes  (OK or WARN)  ← dangerous direction
- FALSE_FLAG: LL Yes, predicted No   (FAIL only)    ← noisy direction

### Re-run

Run full goldset (77 answered rows — rows 20, 23, 36 stay excluded).
Print: overall agreement %, four-class counts, recall on No-rows (TRUE_FLAG / 40),
precision (TRUE_FLAG / (TRUE_FLAG + FALSE_FLAG)).
Expected: recall ~92–95%, precision ~84%, agreement ~87%, FALSE_PASS still 2.
Write `calibration_report_remap_<timestamp>.xlsx` (same colour coding: FALSE_PASS red,
FALSE_FLAG yellow, agreements green).

### Do NOT

- Do not change SYSTEM_PROMPT in this task.
- Do not change the 98/95 score thresholds in va_evaluator.py.
- Do not touch va_evaluator.py at all — this task is calibrate.py only.

### Report back

Four-class counts + recall/precision. Confirm whether the 2 FALSE_PASSes are still
rows 13 and 39 (both previously judged defensible — expected yes).

---

## TASK 2 — System prompt updates in va_evaluator.py (SYSTEM_PROMPT only)

Four targeted edits derived from goldset calibration + deep MT→PE diff analysis
(232 segments, PE-PT-3050063). Apply each exactly as written. No other changes.

---

### Edit 2.1 — Extend the NB-NO known-error lexicon

The lexicon was added in the previous handoff. **Replace it entirely** with this
expanded version (new entries marked with `← NEW`):

```
=== NB-NO KNOWN MT ERROR LEXICON (from LL-rated calibration + PE diff data) ===

These specific errors are produced repeatedly by the MT engine on this content and
were rejected by Language Leads across multiple data sources. Treat every occurrence
as an error requiring at minimum a WARN, and FAIL if the meaning is significantly wrong.

TERM SUBSTITUTIONS (MT uses wrong word — correct term confirmed by LL):
- "vacuum bags" (household/vacuum cleaner context) → "vakuumposer" is WRONG
  Correct: "støvsugerposer". Note: vakuumposer = vacuum-seal storage bags (e.g. sous-vide);
  støvsugerposer = vacuum cleaner bags. Distinguish by context.
- "placemats" → any variant of "bordskåler" is WRONG. Correct: "bordbrikker"
- "baseball cap" → "baseballhette" / "baseballhetten" is rejected.
  Preferred: "baseballcaps" / "baseballcapsen"
- "hat" → wrong type is an error: "lue" = knitted/beanie; "hatt" = brimmed/general.
  Check context and product type.
- "great" / "flott" in product quality context → "flott" is weak. Preferred: "god"
- "Suit for adult(s)" → "Drakt for voksne" is WRONG (drakt = costume/outfit).
  Correct: "Passer til voksne" or "Egnet for voksne"
- "designed" → "designet" is a recurring MT calque.              ← NEW
  Language Leads consistently replace it with "laget" / "lages" in product copy.
  Flag "designet" in sentences like "er designet med/for å" as style:unidiomatic.

UNIDIOMATIC PATTERNS (MT produces technically valid but unnatural Norwegian):
- "vennligst" → consistently removed as unidiomatic. Flag any occurrence.
  Natural alternative: rephrase the instruction without vennligst.
- Over-long compound chains (e.g. "motorsykkelbremsekoblingsspaker") → flag as
  style:unidiomatic; natural Norwegian splits these with spaces or hyphens.
- "På salg" for limited-time sale → anglicism. Natural NB-NO: "På tilbud" / "Tilbud".

This lexicon grows with each production cycle. Apply the same scrutiny to comparable
terms not yet listed — these entries indicate the TYPES of errors this MT engine makes.
```

---

### Edit 2.2 — New rule: bracketed section labels stay in source language

This pattern was confirmed 4/4 times in the PE diff data. Insert as a new paragraph
INSIDE the ACCURACY → Untranslated section (after the existing untranslated rule),
before the FLUENCY section:

```
Bracketed section labels [Like This] must NOT be translated.
If the source contains a structural label such as [Design Description],
[Material Description], [Product Performance], [Accessory Construction], or any
similar [Label], the target must preserve the label exactly as in the source —
including capitalisation and English text. Translating these labels (e.g.
[Designbeskrivelse], [Produktytelse]) is an accuracy:untranslated error.
These are system/structural markers, not translatable copy.
Similarly: broken source tokens that appear to be metadata values (e.g. a material
listed as "COULD" or a spec value in all-caps English) should be mirrored verbatim,
not translated. Translating them is an error.
```

---

### Edit 2.3 — Extend the typography / spacing rule to all unit symbols

Current text in the LOCALE CONVENTIONS block:

> "A missing space between a number and % (e.g. "100%") is a HARD ERROR..."

Extend the sentence to cover all measurement units. Replace that sentence with:

```
A missing space between a number and any unit symbol is a HARD ERROR — Language
Leads reject segments on this alone. This applies to ALL unit symbols:
%, mm, cm, m, km, kg, g, ml, l, W, V, A, mAh, kB, MB, GB, °C, and others.
Examples: "100%" → "100 %", "14mm" → "14 mm", "5kg" → "5 kg".
Same rule applies to en-dash in numerical ranges: "5-10" → "5–10" is a HARD ERROR.
```

---

### Edit 2.4 — Reorder output JSON (reasoning before score)

Replace the output JSON block in `=== OUTPUT FORMAT ===` with:

```json
{
  "reasoning": "<1-3 sentences max, specific, cite the exact issue and the affected text>",
  "error_category": "<one of the category strings below, or no-error>",
  "score": <integer 0-100>,
  "severity": "<OK|WARN|FAIL>"
}
```

The model generates text sequentially — committing to reasoning before the score
improves judgement quality. The parsing code in `evaluate_segment` is key-based and
needs no change, but verify after editing.

---

### Do NOT (Task 2)

- Do not change the grading scale score thresholds (98/95) — they are correct.
- Do not change the error taxonomy strings — they must stay aligned to Amazon's labels.
- Do not add few-shot examples yet — that is a separate future task.
- Do not modify calibrate.py in this task.

---

## TASK 3 — Save PE pattern findings as a notes file

Create `notes/pe_pattern_findings.md` with the content below verbatim. This is a
reference document for future tasks — no code, just create the file.

```markdown
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
```

---

## Report back after all three tasks

1. Task 1: four-class counts + recall/precision, confirm FALSE_PASS rows.
2. Task 2: confirm all four edits applied (lexicon, bracket rule, unit spacing, JSON reorder).
3. Task 3: confirm `notes/pe_pattern_findings.md` created.
4. Suggest CLAUDE.md roadmap update:
   - Mark goldset validation ✅ done with final remap numbers
   - Mark system prompt update ✅ done (lexicon v2, bracket rule, unit spacing, JSON reorder)
   - Add to Planned: "Deterministic pre-checks (improvement #5) — patterns documented
     in notes/pe_pattern_findings.md — build after remap calibration confirms baseline"
