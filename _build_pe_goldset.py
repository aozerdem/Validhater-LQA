"""
Job 1: build a 50-row PE calibration goldset from the linguist-labelled feedback.
Ground truth = column "Would linguist have sent to PE?" (Yes = send to PE, No = keep).

Selection: keep ALL 24 "No" rows (scarce minority, defines the keep boundary) +
random 26 of the 82 "Yes" (fixed seed for reproducibility) = 50.
"""
import json
import random
from pathlib import Path

SRC = Path(r"C:\Users\Ahmet.Ozerdem\Desktop\Scripts\TDT-LQA\linguist feedback\POST EDITING FEEDBACK_PE_va_report_PE_20260612_102519.md")
OUT = Path(r"C:\Users\Ahmet.Ozerdem\Desktop\Scripts\TDT-LQA\repo\pe_goldset_parsed.json")

lines = SRC.read_text(encoding="utf-8").splitlines()
headers, raw_rows = None, []
for line in lines:
    s = line.strip()
    if not s.startswith("|"):
        continue
    cells = [c.strip().replace("\\_", "_") for c in s.split("|")[1:-1]]
    if all(set(c) <= {"-", " ", ":"} for c in cells):
        continue
    if headers is None:
        headers = cells
        continue
    raw_rows.append(cells)

H = {h: i for i, h in enumerate(headers)}

def norm(v):
    v = v.strip().lower()
    if v in ("yes", "y", "ja"): return "Yes"
    if v in ("no", "n", "nei"): return "No"
    return None

parsed = []
for c in raw_rows:
    verdict = norm(c[H["Would linguist have sent to PE?"]]) if H["Would linguist have sent to PE?"] < len(c) else None
    if verdict is None:
        continue
    parsed.append({
        "seg": c[H["SegmentID"]],
        "source": c[H["Source"]],
        "mt": c[H["MT_Target"]],
        "ll_send_to_pe": verdict,                 # GROUND TRUTH
        "prior_ai_sev": c[H["Severity"]],
        "prior_ai_score": c[H["Score"]],
        "prior_ai_cat": c[H["ErrorCategory"]],
    })

yes = [r for r in parsed if r["ll_send_to_pe"] == "Yes"]
no  = [r for r in parsed if r["ll_send_to_pe"] == "No"]

random.seed(42)
yes_sample = random.sample(yes, 26)
selected = no + yes_sample
random.shuffle(selected)
for i, r in enumerate(selected, 1):
    r["idx"] = i

OUT.write_text(json.dumps(selected, ensure_ascii=False, indent=2), encoding="utf-8")

# Composition report
def ai_pred(sev): return "Yes" if sev == "FAIL" else "No"
yes_caught = sum(1 for r in yes_sample if ai_pred(r["prior_ai_sev"]) == "Yes")
yes_missed = sum(1 for r in yes_sample if ai_pred(r["prior_ai_sev"]) == "No")
no_caught  = sum(1 for r in no if ai_pred(r["prior_ai_sev"]) == "No")

print(f"Source pool: {len(parsed)} clear-verdict rows ({len(yes)} Yes, {len(no)} No)")
print(f"Selected: {len(selected)} rows -> {OUT.name}")
print(f"  No  (keep, AI should NOT send to PE): {len(no)}  (all kept)")
print(f"  Yes (send to PE), sampled:           {len(yes_sample)}")
print(f"     of those, old prompt CAUGHT (FAIL): {yes_caught}")
print(f"     of those, old prompt MISSED (OK/WARN): {yes_missed}")
print(f"  No rows old prompt correctly kept clean: {no_caught}/{len(no)}")
print()
print("Selected rows (idx | seg | LL | old-AI | source):")
for r in selected:
    print(f"  {r['idx']:>2} | seg {r['seg']:>4} | LL={r['ll_send_to_pe']:>3} | {r['prior_ai_sev']:>4} {r['prior_ai_score']:>3} | {r['source'][:55]}".encode('ascii','replace').decode())
