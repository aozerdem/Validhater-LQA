# CLAUDE.md — VA Evaluator Project Context

> **Purpose of this file:** Full project context for Claude Code working on this repository.
> Read this before making any changes. It explains what this tool is, why it exists,
> the decisions already made, and the roadmap. The human you are working with is Ahmet —
> he prefers direct, concise answers, step-by-step development, and no scope creep.

---

## 1. WHO / WHERE

- **Developer/owner:** Ahmet Özerdem, Account Quality Manager at Acclaro (localization company).
- **Ahmet's profile:** Localization expert (workflows, QA, termbase/TM management, query management).
  Foundational Python — writes practical tooling, not production software engineering. Explain
  technical decisions clearly, don't assume deep CS background, but don't dumb things down either.
- **Working style:** Methodical, checkpoint-based. Build MvP first, fundamentals before improvements.
  Make a plan, stick to it, no deviation. He will explicitly say when to move to the next step.

## 2. THE PROJECT (business context)

**Client:** Amazon (via Acclaro). Program: **TDT — Translation Data for Training.**
ASIN product-listing content (titles, bullet points, descriptions) is translated to produce
**parallel translation data used to train Amazon's MT systems**. Not customer-facing content,
but Amazon's quality bar is high and they run their own QA scorecards on deliveries.

**Current engagement:** 3.3M source words, **EN-GB → NB-NO** (Norwegian Bokmål),
RFQ# SVNO260326. Kickoff June 10, 2026. Final SLA August 17, 2026. Staggered deliveries.

**Workflow:** Validation → Full MTPE → 10% Revision, on WOL (WordsOnline / ATMS, Memsource-based).

- **Validation step:** Linguists review raw MT segment-by-segment. If the MT is publishable
  as-is, they mark it `Publish = Yes` (it skips post-editing entirely → cost/time savings).
  If not, `Publish = No` → routed to MTPE.
- **The risk:** A validator who wrongly publishes bad MT lets errors straight through to
  the customer. Amazon scores deliveries; files scoring **below ~95% fail** their QA.
- **Two NB-NO Language Leads (validators in the sample):** Birgitte Sciarretta, Monica Opdal.
  (Memory note: earlier project emails called them "Birgitte and Monique" — same people.)

**Important history (decided, do not reopen):** Ahmet argued the project should be Light MTPE
based on poor source quality; management (Mel — VP, Roxana — Quality Manager) decided **FMTPE
is final** because that's what Amazon bought. The SGs are shared with linguists **as-is**, with
explicit instruction that the Light-MTPE carve-outs in the appendix **do not apply**.
This tool exists partly to make the FMTPE bar achievable in practice.

## 3. WHAT THIS TOOL DOES

**`va_evaluator.py`** — AI-powered QA checker for validated segments.

Pipeline:
1. **Read** Galileo HO export(s) (.xlsx) → filter `Segments` sheet for `Publish == "Yes"`.
2. **Evaluate** each (source, MT-target) pair with Claude against the TDT Core SG +
   nb-NO Appendix + known Amazon escalation patterns.
3. **Score** each segment 0–100 with severity (OK / WARN / FAIL), error category
   (Amazon's scorecard taxonomy), and short reasoning.
4. **Aggregate** per validator: average score, severity counts, category breakdown,
   flagged-segment list.
5. **Write** a separate output workbook (never mutate the source export):
   `Segments` sheet (per-row, colour-coded severity) + `Validator_Summary` sheet.

**Run:** `python va_evaluator.py file1.xlsx [file2.xlsx ...]`  (needs `ANTHROPIC_API_KEY` env var)
**Dependencies:** `openpyxl`, `anthropic`. Model: `claude-sonnet-4-20250514`.

## 4. INPUT FILE FORMAT (Galileo HO export)

Three sheets: `Handoff` (client/project meta), `Subtasks` (per-task validator/wordcount),
`Segments` (the data). Segments sheet: 46 columns. The ones we use (0-based index):

| Idx | Col | Field |
|-----|-----|-------|
| 4   | E   | SourceSegment (EN-GB) |
| 5   | F   | OriginalTargetSegment — the raw MT the validator approved |
| 10  | K   | SegmentOriginType (all `mt` in sample, MatchValue 0 = no TM leverage) |
| 15  | P   | ValidationResource (validator name) |
| 16  | Q   | ValidatorGalileoID |
| 20  | U   | **Publish** — "Yes" = validated, our scope; "No" = sent to PE |
| 40  | AP  | FileName |
| 41  | AQ  | SegmentID (source CSV filename) |

Columns G/H/I (`AI-PE output`, `AI-QE evaluation`, `AI-QE score`) exist in the schema but are
empty — schema placeholders. We deliberately **do not** write into them; output is a separate file.
Sample file stats: 646 segments, 232 Publish=Yes (Birgitte 205, Monica 27).

## 5. SCORING CALIBRATION (derived from real Amazon QA data — do not change casually)

From historical Amazon scorecards (2023 TDT arc launch, EN-GB → DE/ES/FR/IT/NL):
- Amazon pass threshold ≈ **95%**. Passing files scored 96–99; failing files 80–94.
- Thresholds in code: `OK >= 98`, `WARN 95–97`, `FAIL < 95`.

Error taxonomy = Amazon's scorecard labels (use these exact strings):
`accuracy:mistranslation` `accuracy:omission` `accuracy:addition` `accuracy:untranslated`
`fluency:grammar` `fluency:spelling` `fluency:typography` `fluency:unintelligible`
`style:unidiomatic` `style:company-style` `locale-convention:number-format`
`locale-convention:measurement-format` `no-error`

Frequency/severity weighting from real data: mistranslation is the #1 fail driver;
omission, untranslated, grammar, spelling all high-frequency; literal/unidiomatic
phrasing is a consistent Amazon complaint across languages.

**Known dispute area:** ALL CAPS handling (mirror source caps vs. Norwegian rules) was
disputed between Jonckers LAs and Amazon QA with no SG resolution. The evaluator should
**not flag ALL CAPS** as an error — deliberately neutral on this.

**Escalation patterns baked into the system prompt** (from a real ES-ES escalation):
false friends ("loved one" → "amante"), literal translations, number-accord errors,
unintelligible MT validated unchanged, trailing space before period, measurement
format violations, untranslated common words.

## 6. STYLE GUIDE RULES (embedded in SYSTEM_PROMPT — source documents)

- **TDT Core SG:** parallel-translation principle (mirror source; no additions/omissions;
  keep punctuation/structure/element order unless target grammar conflicts). Core SG has
  priority over the appendix when they conflict.
- **nb-NO Appendix highlights:** decimal comma (50,5), non-breaking-space thousands (1 526),
  dates DD.MM.YYYY, 24-h time, °C only, space before unit symbols (100 m, 50 %),
  guillemets « », en-dash for ranges, "hos Amazon" not "ved Amazon", acronym plurals
  PC-ene/TV-ene, mva. with period, formal register, gender-neutral, active voice,
  no full stop at end of headings.
- Source `.md` versions of both SGs exist in the project materials (Ahmet has them).

## 7. CALIBRATION / GOLDSET (in progress)

An 80-segment goldset was sent to the LLs (via Roxana) — they fill in:
Publish Yes/No, corrected PE version if No, detailed reason if No.
**When the filled goldset comes back:** run the evaluator on those 80 segments and compare
against LL judgements. Tune the system prompt if miscalibrated. This is the acceptance
gate before trusting the tool on production exports.

An **LQA guidelines** document also exists — Ahmet will share it; consider folding relevant
parts into the system prompt at calibration time (his words: "use it however you want,
you can even skip it altogether").

## 8. ROADMAP

**Done:**
- [x] Checkpoint 1 — Excel reader, multi-file batch via CLI args
- [x] Checkpoint 2 — single-segment evaluator function + system prompt (structurally tested;
      live calibration pending API run on Ahmet's machine)
- [x] Checkpoint 3 — batch runner + per-validator aggregation
- [x] Checkpoint 4 — Excel report output (separate file, two sheets, colour-coded)

**Next (in order):**
- [ ] Live Checkpoint-2 calibration: 5 hand-picked segments, sanity-check scores
- [ ] Goldset validation: run on 80 LL-rated segments, measure agreement, tune prompt
- [ ] Process first real production HO exports (project kicked off June 10, 2026)

**Planned (do not build until Ahmet says so):**
- [ ] **Streamlit deployment** — the tool will move from CLI to a Streamlit app for LL/PM use.
      Keep the core logic (read/evaluate/aggregate/report) cleanly separable from the CLI
      so the Streamlit wrapper is thin.
- [ ] Possible later features (explicitly deferred from MvP): termbase consistency checks,
      "modification implemented?" re-check loop (Roxana asked for this), per-source-file
      scoring, PE-step quality reports (same approach, different export slice).

**Explicitly out of scope for now:** GUI beyond Streamlit plan, multi-language support
(NB-NO only for this engagement; SV-SE etc. may come later), MT-engine comparison,
time-on-task analytics.

## 9. WORKING AGREEMENTS

1. **No scope creep.** Build what's on the roadmap, in order. If a new idea comes up,
   note it under "Planned", don't implement it.
2. **Linguistic reliability over features.** When in doubt, make the evaluation stricter
   to test and more transparent, not fancier.
3. **Never mutate input files.** Reports always go to separate output files.
4. **Keep the system prompt's error taxonomy aligned to Amazon's scorecard labels** —
   the output needs to map 1:1 to how Amazon reports issues back.
5. **Folder hygiene:** repo had leftover scaffold files (`cli.py`, `excel_reader.py`) from
   an earlier session — superseded by the self-contained `va_evaluator.py`. Archive or
   delete; don't re-introduce parallel entry points.
6. Ahmet will keep this file updated with major decisions. If you make a structural change,
   suggest the corresponding CLAUDE.md update.
