"""
tmx_analyzer.py
Analyze Amazon TM (TMX) 100%/101% matches → extract term and style insights
for SYSTEM_PROMPT enrichment in va_evaluator.py.

Usage:
    python tmx_analyzer.py [file.tmx]
    (or run with no args — file picker dialog will open)

Output:
    tmx_findings_<timestamp>.md — structured findings ready for human review
"""

import sys
import json
import re
import os
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import boto3
from botocore.config import Config

BEDROCK_MODEL_ID = "us.anthropic.claude-opus-4-7"
BATCH_SIZE = 50   # segments per Claude call

# xml:lang namespace prefix used by ET
_XML_LANG = "{http://www.w3.org/XML/1998/namespace}lang"


# ─────────────────────────────────────────────
# UI HELPERS  (same pattern as va_evaluator.py)
# ─────────────────────────────────────────────

def ask_bearer_token() -> str:
    result = {"token": ""}
    win = tk.Tk()
    win.title("AWS Bedrock Token")
    win.resizable(False, False)
    win.attributes("-topmost", True)
    tk.Label(win, text="Enter your AWS Bedrock bearer token:").pack(padx=20, pady=(16, 4))
    entry = tk.Entry(win, show="*", width=56)
    entry.pack(padx=20, pady=4)
    entry.focus_set()
    def submit():
        result["token"] = re.sub(r"\s", "", entry.get())
        win.destroy()
    tk.Button(win, text="Continue", command=submit, width=12).pack(pady=(8, 16))
    win.bind("<Return>", lambda _: submit())
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.mainloop()
    return result["token"]


def pick_tmx_file() -> str:
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title="Select TMX file to analyze",
        filetypes=[("TMX files", "*.tmx"), ("All files", "*.*")],
    )
    root.destroy()
    return path


# ─────────────────────────────────────────────
# TMX PARSER
# ─────────────────────────────────────────────

def _seg_text(tuv) -> str:
    """Extract plain text from a <tuv>, stripping any inline tags (<ph>, <bpt>, etc.)."""
    seg = tuv.find("seg")
    if seg is None:
        return ""
    # ET itertext() walks text + tail of all children — gives us the text content
    # without inline tag names
    return "".join(seg.itertext()).strip()


def parse_tmx(filepath: str) -> list[dict]:
    """
    Parse a TMX file and return a list of {source, target} dicts.
    Handles both xml:lang="en-GB" and plain lang="en-GB" attribute styles.
    Skips pairs where source == target (untranslated carry-overs).
    """
    tree = ET.parse(filepath)
    root = tree.getroot()

    pairs = []
    for tu in root.iter("tu"):
        source = target = None
        for tuv in tu.findall("tuv"):
            # Try namespace-qualified first, fall back to plain attr
            lang = (tuv.get(_XML_LANG) or tuv.get("lang") or "").lower()
            text = _seg_text(tuv)
            if not text:
                continue
            if lang.startswith("en"):
                source = text
            elif lang.startswith("nb") or lang.startswith("no"):
                target = text

        if source and target and source.strip() != target.strip():
            pairs.append({"source": source, "target": target})

    return pairs


# ─────────────────────────────────────────────
# BATCH ANALYSIS
# ─────────────────────────────────────────────

_BATCH_SYSTEM = (
    "You are a Norwegian Bokmål (NB-NO) localization expert analyzing Amazon translation memory segments. "
    "These are 100%/101% TM matches — translations Amazon has approved for publication. "
    "Your output feeds into a QA evaluator prompt. Be precise, specific, and concise."
)


def analyze_batch(client, pairs: list[dict]) -> dict:
    """Send one batch of EN→NB pairs to Claude. Returns parsed findings dict."""
    segments_text = "\n".join(
        f"{i + 1}. EN: {p['source']} | NB: {p['target']}"
        for i, p in enumerate(pairs)
    )

    user_message = f"""Analyze these {len(pairs)} Amazon-approved EN-GB → NB-NO translation pairs.
Goal: extract insights that help an AI evaluator catch NB-NO quality issues in Amazon product copy.

--- SEGMENTS ---
{segments_text}
--- END SEGMENTS ---

Return ONLY valid JSON — no preamble, no markdown fences:
{{
  "term_choices": [
    {{"en": "source term", "nb_no": "approved NB-NO term", "notes": "why non-obvious or noteworthy"}}
  ],
  "style_patterns": ["specific observation with example text from the segments"],
  "conventions": ["specific confirmed convention — only if 2+ segments support it"],
  "surprising_choices": ["consistent deviation from standard NB-NO that appears intentional"]
}}

Rules:
- term_choices: only where the NB-NO choice is non-obvious or differs from a generic dictionary. Skip obvious cognates and proper nouns.
- style_patterns: cite actual segment text. Concrete, not abstract.
- conventions: number/unit formatting, punctuation, capitalisation — confirmed by multiple examples only.
- surprising_choices: anything Amazon does that overrides standard NB-NO rules.
- Return [] for any category with nothing noteworthy in this batch. No filler."""

    response = client.converse(
        modelId=BEDROCK_MODEL_ID,
        system=[{"text": _BATCH_SYSTEM}],
        messages=[{"role": "user", "content": [{"text": user_message}]}],
        inferenceConfig={"maxTokens": 2500},
    )

    raw = response["output"]["message"]["content"][0]["text"].strip()
    raw = re.sub(r"^```json\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    brace_idx = raw.find("{")
    if brace_idx > 0:
        raw = raw[brace_idx:]

    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"    [warning] JSON parse failed: {e}")
        return {"term_choices": [], "style_patterns": [], "conventions": [], "surprising_choices": []}


# ─────────────────────────────────────────────
# SYNTHESIS
# ─────────────────────────────────────────────

_SYNTHESIS_SYSTEM = (
    "You are editing a QA evaluator system prompt for Amazon EN-GB → NB-NO translation quality assessment. "
    "You have aggregated raw findings from Amazon's own translation memory (100%/101% matches). "
    "Produce clean, deduplicated, ready-to-paste prompt additions. Every sentence must add value for a quality evaluator."
)


def _dedup_strings(items: list[str]) -> list[str]:
    seen = set()
    result = []
    for item in items:
        key = item.lower().strip()
        if key not in seen:
            seen.add(key)
            result.append(item)
    return result


def synthesize(client, all_findings: list[dict], total_pairs: int) -> str:
    """Merge all batch findings into a clean, curated markdown document."""
    # Flatten
    all_terms    = [item for f in all_findings for item in f.get("term_choices", [])]
    all_patterns = [item for f in all_findings for item in f.get("style_patterns", [])]
    all_convs    = [item for f in all_findings for item in f.get("conventions", [])]
    all_surprises= [item for f in all_findings for item in f.get("surprising_choices", [])]

    # Pre-deduplicate before sending — keeps payload manageable
    seen_en: dict = {}
    for item in all_terms:
        key = item.get("en", "").lower().strip()
        if key and key not in seen_en:
            seen_en[key] = item
    deduped_terms = list(seen_en.values())

    combined = json.dumps({
        "term_choices":       deduped_terms,
        "style_patterns":     _dedup_strings(all_patterns),
        "conventions":        _dedup_strings(all_convs),
        "surprising_choices": _dedup_strings(all_surprises),
    }, ensure_ascii=False, indent=2)

    user_message = f"""The findings below were extracted in batches from {total_pairs} Amazon-approved EN-GB → NB-NO TM segments.
Deduplicate, merge near-duplicates, discard weak or obvious entries, and produce the final output.

RAW FINDINGS:
{combined}

Write a markdown document with exactly these four sections:

## New Term Choices
Markdown table: | EN term | Amazon NB-NO | Notes |
Only include entries that are non-obvious and appear in multiple segments.
Sort by frequency (most-confirmed first).

## Style Patterns (for NB-NO Language Appendix)
Bullet list of Amazon-specific patterns in product copy structure or phrasing.
Each bullet: one concrete, actionable rule a QA evaluator can apply.

## Confirmed Conventions (for Locale / Typography sections)
Bullet list of format/punctuation/capitalisation decisions confirmed by this TM data.
Exclude anything already standard NB-NO. Flag where Amazon deviates from the norm.

## House Style Overrides
Bullet list of places where Amazon consistently overrides standard NB-NO rules.
Write "None identified." if there are none.

## Reviewer Notes
2-3 sentences: what this TM data reveals about content type, reliability, and caveats
before integrating these findings into the evaluator prompt."""

    response = client.converse(
        modelId=BEDROCK_MODEL_ID,
        system=[{"text": _SYNTHESIS_SYSTEM}],
        messages=[{"role": "user", "content": [{"text": user_message}]}],
        inferenceConfig={"maxTokens": 5000},
    )

    return response["output"]["message"]["content"][0]["text"].strip()


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    if len(sys.argv) >= 2:
        tmx_path = sys.argv[1]
    else:
        print("No file specified — opening file picker...")
        tmx_path = pick_tmx_file()
        if not tmx_path:
            print("No file selected. Exiting.")
            sys.exit(0)

    if not Path(tmx_path).exists():
        print(f"Error: file not found: {tmx_path}")
        sys.exit(1)

    token = ask_bearer_token()
    if not token:
        print("No token provided. Exiting.")
        sys.exit(1)

    print(f"\n=== TMX Analyzer ===")
    print(f"File: {Path(tmx_path).name}")

    print("\nParsing TMX...")
    pairs = parse_tmx(tmx_path)
    print(f"  {len(pairs)} EN→NB translation pairs extracted")

    if not pairs:
        print("No valid EN→NB pairs found. Check that the file contains en-GB and nb-NO tuv elements.")
        sys.exit(0)

    os.environ["AWS_BEARER_TOKEN_BEDROCK"] = token
    client = boto3.client(
        service_name="bedrock-runtime",
        region_name="us-east-1",
        config=Config(read_timeout=300, connect_timeout=30),
    )

    batches = [pairs[i:i + BATCH_SIZE] for i in range(0, len(pairs), BATCH_SIZE)]
    print(f"\nRunning {len(batches)} batch(es) × up to {BATCH_SIZE} segments each...")

    all_findings = []
    for i, batch in enumerate(batches, 1):
        print(f"  Batch {i}/{len(batches)} ({len(batch)} segments)...", end=" ", flush=True)
        findings = analyze_batch(client, batch)
        all_findings.append(findings)
        n_terms    = len(findings.get("term_choices", []))
        n_patterns = len(findings.get("style_patterns", []))
        n_convs    = len(findings.get("conventions", []))
        print(f"→ {n_terms} terms | {n_patterns} patterns | {n_convs} conventions")

    print("\nSynthesizing all findings...")
    synthesis = synthesize(client, all_findings, len(pairs))

    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = Path(tmx_path).parent.parent / f"tmx_findings_{timestamp}.md"

    header = f"""# TMX Analysis Findings
**Source:** {Path(tmx_path).name}
**Segments analyzed:** {len(pairs)}
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M")}

> **Before using:** review every entry below. TM data reflects approved usage but may contain
> inconsistencies or domain-specific choices that don't generalise. Curate before adding to SYSTEM_PROMPT.

---

"""
    output_path.write_text(header + synthesis, encoding="utf-8")
    print(f"\nFindings saved: {output_path}")
    print("=== Done ===")


if __name__ == "__main__":
    main()
