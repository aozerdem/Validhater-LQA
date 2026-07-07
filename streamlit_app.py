"""
streamlit_app.py
TDT LQA Evaluator — web UI for Language Leads
"""

import base64
import io
import os
import random
import threading
from datetime import datetime

import streamlit.components.v1 as components

import boto3
import streamlit as st
from botocore.config import Config
from streamlit.runtime.scriptrunner import add_script_run_ctx, get_script_run_ctx

from va_evaluator import (
    read_validated_segments,
    read_postedited_segments,
    run_batch,
    aggregate_by_validator,
    write_report,
    load_termbase,
)

st.set_page_config(page_title="TDT LQA Evaluator", layout="wide")


def _auto_download(report_bytes: bytes, mode: str, timestamp: str) -> None:
    """Inject a JS snippet that immediately downloads the report to the user's Downloads folder."""
    b64      = base64.b64encode(report_bytes).decode()
    filename = f"va_report_{mode}_{timestamp}.xlsx"
    mime     = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    components.html(
        f"""
        <script>
            (function() {{
                const a = document.createElement('a');
                a.href = 'data:{mime};base64,{b64}';
                a.download = '{filename}';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            }})();
        </script>
        """,
        height=0,
    )


# ─────────────────────────────────────────────
# AUTH
# ─────────────────────────────────────────────

def login_gate() -> bool:
    if st.session_state.get("logged_in"):
        return True

    _, col, _ = st.columns([1, 2, 1])
    with col:
        st.title("TDT LQA Evaluator")
        st.caption("Amazon EN-GB → NB-NO · Quality Assurance")
        st.write("")
        with st.form("login"):
            username = st.text_input("Username")
            password = st.text_input("Password", type="password")
            if st.form_submit_button("Log in", use_container_width=True):
                users = dict(st.secrets.get("users", {}))
                if username in users and users[username] == password:
                    st.session_state.logged_in = True
                    st.session_state.current_user = username
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")
    return False


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    if not login_gate():
        return

    # Header
    c1, c2 = st.columns([7, 3])
    with c1:
        st.title("TDT LQA Evaluator")
        st.caption("Amazon EN-GB → NB-NO · Quality Assurance")
    with c2:
        st.write("")
        st.write(f"Logged in as **{st.session_state.current_user}**")
        if st.button("Log out"):
            st.session_state.pop("eval_results", None)
            st.session_state.logged_in = False
            st.rerun()

    st.divider()

    # ── Settings ──────────────────────────────
    col_mode, col_scope = st.columns(2)

    with col_mode:
        mode = st.radio(
            "Check mode",
            options=["PEQA", "PU", "PE"],
            captions=[
                "Score linguists' post-edited output — flags below 95",
                "Check if published strings were right to publish",
                "Check if strings were right to send to PE",
            ],
        )

    with col_scope:
        full_check = st.radio("Scope", ["Full check", "Spot check"]) == "Full check"
        spot_n = None
        if not full_check:
            spot_n = st.number_input(
                "Segments to sample", min_value=5, max_value=500, value=50, step=5
            )

    st.divider()

    # ── File upload ───────────────────────────
    upload_label = (
        "Upload post-editing handoff export(s) (.xlsx)"
        if mode == "PEQA"
        else "Upload Galileo VA export(s) (.xlsx)"
    )
    uploaded = st.file_uploader(upload_label, type="xlsx", accept_multiple_files=True)

    if uploaded:
        st.write(f"**{len(uploaded)} file(s) ready**")
        if st.button("Run evaluation", type="primary", use_container_width=True):
            _run(uploaded, mode, spot_n)
    elif "eval_results" not in st.session_state:
        st.info("Upload one or more .xlsx files to begin.")

    # ── Results — always render from session state ─────────────────────────
    if "eval_results" in st.session_state:
        r = st.session_state["eval_results"]
        if r.get("fresh"):
            r["fresh"] = False
            _auto_download(r["report_bytes"], r["mode"], r["timestamp"])
        n_errors = sum(1 for s in r["segments"] if s.get("error_category") == "api-error")
        if n_errors:
            st.warning(
                f"{n_errors} segment(s) could not be evaluated after retries and are marked "
                f"as api-error in the report. Re-running the batch will resolve this."
            )
        st.success(
            f"Report ready — {len(r['segments'])} segments · "
            f"{r['mode']} · {r['timestamp']}"
        )
        if st.button("New evaluation (clear results)"):
            st.session_state.pop("eval_results", None)
            st.rerun()
        st.divider()
        _show_results(r["segments"], r["summary"], r["mode"],
                      r["report_bytes"], r["timestamp"])


# ─────────────────────────────────────────────
# EVALUATION
# ─────────────────────────────────────────────

def _run(uploaded_files, mode: str, spot_n):
    segments = []

    with st.status("Reading files…") as status:
        for f in uploaded_files:
            try:
                file_bytes = io.BytesIO(f.read())
                if mode == "PEQA":
                    segs = read_postedited_segments(file_bytes, source_label=f.name)
                else:
                    publish_filter = "Yes" if mode == "PU" else "No"
                    segs = read_validated_segments(
                        file_bytes, publish_filter, source_label=f.name
                    )
                st.write(f"✓ {f.name} — {len(segs)} segments")
                segments.extend(segs)
            except Exception as exc:
                st.error(f"Could not read {f.name}: {exc}")
                return
        status.update(label=f"{len(segments)} segments loaded", state="complete")

    if not segments:
        st.warning("No matching segments found in the uploaded files.")
        return

    if spot_n and spot_n < len(segments):
        segments = random.sample(segments, int(spot_n))
        st.info(f"Spot check: randomly selected {len(segments)} segments.")

    try:
        os.environ["AWS_BEARER_TOKEN_BEDROCK"] = st.secrets["aws"]["bearer_token"]
        client = boto3.client(
            service_name="bedrock-runtime",
            region_name="us-east-1",
            config=Config(read_timeout=300, connect_timeout=30),
        )
    except Exception as exc:
        st.error(f"AWS connection failed: {exc}")
        return

    termbase = load_termbase()

    progress_bar = st.progress(0.0, text="Starting evaluation…")
    status_text  = st.empty()
    ctx = get_script_run_ctx()

    def progress_fn(done, total, severity, score, category):
        add_script_run_ctx(threading.current_thread(), ctx)
        score_str = str(score) if score is not None else "---"
        sev_icon  = {"OK": "✅", "WARN": "⚠️", "FAIL": "❌"}.get(severity, "•")
        pct = done / total
        progress_bar.progress(pct, text=f"{done}/{total} segments evaluated")
        status_text.markdown(
            f"**Last:** {sev_icon} {severity} &nbsp;|&nbsp; score: `{score_str}` &nbsp;|&nbsp; {category}"
        )

    try:
        segments = run_batch(segments, client, termbase, mode=mode, progress_fn=progress_fn)
    except Exception as exc:
        st.error(f"Evaluation error: {exc}")
        return

    progress_bar.progress(1.0, text="Done")
    status_text.empty()

    summary      = aggregate_by_validator(segments)
    timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_buf   = write_report(segments, summary, output_path=None, mode=mode, return_bytes=True)
    report_bytes = report_buf.read()   # BytesIO → bytes; required for base64 encoding

    st.session_state["eval_results"] = {
        "segments":     segments,
        "summary":      summary,
        "mode":         mode,
        "timestamp":    timestamp,
        "report_bytes": report_bytes,
        "fresh":        True,
    }
    st.rerun()


# ─────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────

def _show_results(segments: list, summary: dict, mode: str, report_bytes: bytes, timestamp: str):
    st.subheader("Resource Performance")

    perf_rows = []
    for v, data in summary.items():
        sc = data["severity_counts"]
        n = data["total_segments"]
        if mode == "PEQA":
            flagged = sum(
                1 for s in segments
                if s["validator_name"] == v
                and s.get("score") is not None and s["score"] < 95
            )
            rate_label = "Below-95 rate"
        else:
            flagged = sc.get("FAIL" if mode == "PU" else "OK", 0)
            rate_label = "Questionable rate"
        perf_rows.append({
            "Linguist":   v,
            "Segments":   n,
            "Avg score":  data["avg_score"],
            "OK":         sc.get("OK", 0),
            "WARN":       sc.get("WARN", 0),
            "FAIL":       sc.get("FAIL", 0),
            rate_label:   f"{flagged / n * 100:.1f}%" if n else "—",
        })

    st.dataframe(perf_rows, use_container_width=True, hide_index=True)

    st.subheader("Segments")

    fc1, fc2 = st.columns(2)
    with fc1:
        sev_filter = st.multiselect(
            "Severity", ["OK", "WARN", "FAIL"], default=["WARN", "FAIL"]
        )
    with fc2:
        names = list(summary.keys())
        name_filter = st.multiselect("Linguist", names, default=names)

    filtered = [
        s for s in segments
        if s.get("severity") in sev_filter and s.get("validator_name") in name_filter
    ]
    st.caption(f"{len(filtered)} of {len(segments)} segments shown")

    rows = []
    for s in filtered:
        src, tgt = s["source"], s["mt_target"]
        rows.append({
            "Seg ID":    s.get("segment_id", ""),
            "Linguist":  s["validator_name"],
            "Source":    src[:100] + "…" if len(src) > 100 else src,
            "Target":    tgt[:100] + "…" if len(tgt) > 100 else tgt,
            "Score":     s.get("score"),
            "Severity":  s.get("severity", ""),
            "Category":  s.get("error_category", ""),
            "Reasoning": s.get("reasoning", ""),
        })

    st.dataframe(rows, use_container_width=True, hide_index=True)

    st.divider()
    st.download_button(
        "⬇ Download Excel report",
        data=report_bytes,
        file_name=f"va_report_{mode}_{timestamp}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )


if __name__ == "__main__":
    main()
