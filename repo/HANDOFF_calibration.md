# HANDOFF — Calibration Harness + System Prompt Updates

> **For Claude Code.** Read CLAUDE.md first for full project context.
> This file specifies two tasks derived from goldset analysis done in claude.ai on 2026-06-11.
> Execute in order. Do not add features beyond what is specified here.

---

## Inputs you have

- `va_evaluator.py` — existing tool (reader / evaluator / aggregation / report)
- `goldset_parsed.json` — NEW, place in repo root. 80 goldset rows parsed from the
  LL-filled Excel, structure per row:
  ```json
  {
    "row": 2,
    "source": "...",            // EN-GB source
    "mt": "...",                // raw MT the LL judged
    "verdict": "Yes" | "No" | null,
    "pe": "...",                // LL's corrected version (if No)
    "reason": "...",            // LL's explanation (if No)
    "mt_red": ["..."],          // spans the LL marked red in the MT = the faulty parts
    "pe_red": ["..."]           // spans the LL marked red in the PE = the corrections
  }
  ```
  Stats: 37 Yes, 40 No, 3 unanswered (rows 20, 23, 36 — **exclude from calibration**).

---

## TASK 1 — Update SYSTEM_PROMPT in va_evaluator.py

Three edits, derived from LL goldset behaviour. Insert exactly as specified.

### Edit 1.1 — New section: NB-NO known-error lexicon

Insert AFTER the `=== KNOWN ESCALATION PATTERNS ===` section, BEFORE `=== GRADING SCALE ===`:

```
=== NB-NO KNOWN MT ERROR LEXICON (from LL-rated calibration data) ===

These specific errors were repeatedly produced by the MT engine on this content and
rejected by the Language Leads. Treat every occurrence as an error:

- "vacuum bags" (household/cleaner context) → "vakuumposer" is WRONG → correct: "støvsugerposer"
  (vakuumposer = vacuum-seal storage bags; støvsugerposer = vacuum cleaner bags. Decide from context.)
- "placemats" → "bordskåler" or any other variant is WRONG → correct: "bordbrikker"
- "baseball cap" → "baseballhette" / "baseballhetten" is rejected → preferred: "baseballcaps" / "baseballcapsen"
- "hat" → check context: "lue" (knitted/beanie) vs "hatt" (brimmed/general). Wrong type = error.
- "great" (quality praise) → "flott" is weak in product context → preferred: "god"
- "Suit for adult(s)" → "Drakt for voksne" is WRONG (Drakt = costume/outfit) → correct: "Passer til voksne"
- "vennligst" → flag as unidiomatic; rarely used in natural Norwegian. Prefer rephrasing.
- Over-long compound chains (e.g. "motorsykkelbremsekoblingsspaker") → flag as unidiomatic;
  natural Norwegian splits these.

This lexicon is not exhaustive — it indicates the TYPES of term-choice errors this MT engine
makes. Apply the same scrutiny to comparable terms.
```

### Edit 1.2 — Strengthen number/percent spacing rule

In the `LOCALE CONVENTIONS` block of the nb-NO appendix section, REPLACE the line:

```
- Measurements: Space between number and symbol (100 m, 50 %). Convert imperial to metric where applicable (except TVs, hard drives, laptops, bicycle tyres → keep imperial).
```

WITH:

```
- Measurements: Space between number and symbol (100 m, 50 %, 100 % bomull). A missing space
  between a number and % (e.g. "100%") is a HARD ERROR — Language Leads reject segments on
  this alone. Same for missing en-dash in numerical ranges (e.g. "5-10" should be "5–10").
  Convert imperial to metric where applicable (except TVs, hard drives, laptops, bicycle
  tyres → keep imperial). NOTE: inches are acceptable in some Norwegian product contexts —
  if the LL context allows inch usage, do not auto-fail; flag as WARN with reasoning.
```

### Edit 1.3 — Recalibrate the grading scale to the LL's actual bar

REPLACE the entire `=== GRADING SCALE ===` section with:

```
=== GRADING SCALE ===

The publishing bar is "most natural current Norwegian usage" — NOT merely "defensible
translation". Language Leads reject technically-correct translations when a more common
term exists (e.g. baseballhette → baseballcaps). Score accordingly:

Score 98–100 (OK):    Correct AND natural. The phrasing is what a Norwegian copywriter
                      would actually produce. Minor stylistic preference differences only.
Score 95–97  (WARN):  Understandable and accurate, but contains a less-common term choice,
                      a single mechanical slip (spacing, dash), or slightly stiff phrasing
                      that a careful validator would fix before publishing.
Score 0–94   (FAIL):  Wrong term, untranslated content, grammar error, missing/added
                      content, unintelligible output, or a hard typographic error
                      (number+% spacing, missing range dash). Should NOT have been published.
```

### Edit 1.4 — Reorder the output JSON (reason before score)

REPLACE the `=== OUTPUT FORMAT ===` JSON block with:

```
{
  "reasoning": "<1-3 sentences max, specific, cite the exact issue>",
  "error_category": "<one of the category strings below, or no-error>",
  "score": <integer 0-100>,
  "severity": "<OK|WARN|FAIL>"
}
```

(Generation is sequential — making the model articulate reasoning before committing to a
score improves judgement. The parsing code in `evaluate_segment` is key-based so it needs
no change, but verify.)

---

## TASK 2 — Build calibrate.py

New file. Purpose: measure evaluator agreement with the LL goldset and produce a
misses report for prompt tuning. Reuse `evaluate_segment` and `SYSTEM_PROMPT` by import
from `va_evaluator.py` — do not duplicate logic.

### Spec

```
python calibrate.py            # runs full goldset (77 answered rows)
python calibrate.py --limit 10 # first N rows only (for cheap smoke tests)
```

Behaviour:
1. Load `goldset_parsed.json`. Skip rows where verdict is null (rows 20, 23, 36).
2. For each row, call `evaluate_segment(client, source, mt)`.
3. Map results to a binary verdict for comparison:
   - Evaluator OK            → predicted "Yes" (publishable)
   - Evaluator WARN or FAIL  → predicted "No"  (should not publish)
   (Keep the raw severity too — we want to see WARN vs FAIL split in the report.)
4. Compare with LL verdict. Classify each row:
   - TRUE_PASS:  LL Yes, evaluator OK
   - TRUE_FLAG:  LL No,  evaluator WARN/FAIL
   - FALSE_PASS: LL No,  evaluator OK          ← DANGEROUS direction, minimize this
   - FALSE_FLAG: LL Yes, evaluator WARN/FAIL   ← noisy direction, acceptable in moderation
5. Print summary to console:
   - overall agreement %
   - count + % for each of the four classes
   - recall on No-rows (TRUE_FLAG / all LL-No) — headline metric
   - precision on flags (TRUE_FLAG / all evaluator flags)
6. Write `calibration_report_<timestamp>.xlsx` with one row per goldset segment:
   row | source | mt | LL_verdict | LL_reason | mt_red_spans | eval_severity |
   eval_score | eval_category | eval_reasoning | classification
   Colour-code: FALSE_PASS red, FALSE_FLAG yellow, agreements green.
7. Additionally, for every FALSE_PASS, print to console: source, mt, the LL's reason,
   and the red spans — this is the working list for the next prompt-tuning pass.

### Acceptance criteria

- Headline target after Task 1 prompt updates: **recall on No-rows ≥ 85%**, FALSE_PASS ≤ 6 rows.
- If first run lands below that, do NOT tune further automatically — output the misses
  report and stop. Prompt tuning decisions go through Ahmet.
- Secondary check: of the TRUE_FLAGs, the evaluator's error_category should broadly match
  the LL's reason. Report category agreement informally in console output (no hard gate).

### Cost note

77 segments × 1 call ≈ trivial. Use the same model as production
(`claude-sonnet-4-20250514`). Run with `--limit 5` first to verify plumbing, then full.

---

## Things NOT to do

- Do not change scoring thresholds (98/95) — the grading-scale text changed, the numeric
  bands did not.
- Do not touch the reader, aggregation, or report-writer code paths in this task.
- Do not add the deterministic pre-checks yet (improvement #5 in the roadmap) — that is a
  separate task after calibration results are in.
- Do not implement self-consistency / double-running yet.
- Rows 20, 23, 36 stay excluded — Ahmet will chase the LL for those separately.

## After running

Report back to Ahmet: the four-class counts, recall on No-rows, and the FALSE_PASS list.
Suggest updating CLAUDE.md roadmap: mark "goldset validation" in progress with the
agreement numbers.
