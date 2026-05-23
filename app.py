"""
StatementSense — Streamlit App
Friendly, simple bank statement analyzer for SBI PDFs.
"""
 
import io
import tempfile
import zipfile
from pathlib import Path
 
import pandas as pd
import streamlit as st
 
from expense_tracker_OriginalID import (
    build_summary,
    classify_transaction,
    merge_pdfs,
    process_pdf,
    write_excel,
)
 
# ─────────────────────────────────────────────
# Page config
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="StatementSense",
    page_icon="💰",
    layout="centered",
    initial_sidebar_state="collapsed",
)
 
# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
 
html, body, [class*="css"], .stApp {
    font-family: 'Plus Jakarta Sans', sans-serif;
    background-color: #F7F8FC;
    color: #1A1D2E;
}
 
#MainMenu, footer, header { visibility: hidden; }
 
.topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 1rem 0 2rem;
    border-bottom: 1.5px solid #E8EAF2;
    margin-bottom: 2.5rem;
}
.topbar-logo {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    font-weight: 700;
    font-size: 1.1rem;
    color: #1A1D2E;
}
.topbar-logo .dot { color: #4361EE; }
.topbar-tag {
    background: #EEF2FF;
    color: #4361EE;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.06em;
    padding: 0.25rem 0.7rem;
    border-radius: 99px;
    text-transform: uppercase;
}
 
.steps {
    display: flex;
    align-items: center;
    gap: 0;
    margin-bottom: 2.5rem;
    justify-content: center;
}
.step-item { display: flex; align-items: center; gap: 0.5rem; }
.step-circle {
    width: 32px; height: 32px; border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 0.8rem; font-weight: 700; flex-shrink: 0;
    border: 2px solid #D1D5E8; color: #9CA3C8; background: white;
}
.step-circle.active { background: #4361EE; border-color: #4361EE; color: white; }
.step-circle.done   { background: #10B981; border-color: #10B981; color: white; }
.step-label { font-size: 0.78rem; font-weight: 600; color: #9CA3C8; white-space: nowrap; }
.step-label.active  { color: #4361EE; }
.step-label.done    { color: #10B981; }
.step-line { width: 48px; height: 2px; background: #E2E5F0; margin: 0 0.3rem; }
.step-line.done { background: #10B981; }
 
.section-head { margin-bottom: 0.5rem; }
.section-head h2 { font-size: 1.45rem; font-weight: 700; color: #1A1D2E; margin: 0 0 0.3rem; }
.section-head p { color: #6B7280; font-size: 0.9rem; margin: 0; font-weight: 400; }
 
[data-testid="stFileUploadDropzone"] {
    background: white !important;
    border: 2px dashed #C7CDE8 !important;
    border-radius: 14px !important;
    padding: 2.5rem !important;
}
[data-testid="stFileUploadDropzone"]:hover { border-color: #4361EE !important; }
[data-testid="stFileUploadDropzone"] p { font-family: 'Plus Jakarta Sans', sans-serif !important; color: #6B7280 !important; }
 
.tip-box {
    background: #EEF2FF; border-left: 3px solid #4361EE;
    border-radius: 0 10px 10px 0; padding: 0.75rem 1rem;
    font-size: 0.83rem; color: #4361EE; margin-top: 1rem; line-height: 1.55;
}
.tip-box strong { font-weight: 600; }
 
.file-badge {
    display: inline-flex; align-items: center; gap: 0.4rem;
    background: white; border: 1.5px solid #E2E5F0; border-radius: 8px;
    padding: 0.35rem 0.75rem; font-size: 0.8rem; color: #374151;
    font-weight: 500; margin: 0.3rem 0.2rem;
}
 
.stButton > button {
    background: #4361EE !important; color: white !important;
    border: none !important; border-radius: 10px !important;
    padding: 0.7rem 2rem !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important; font-size: 1rem !important;
    width: 100%; letter-spacing: 0.01em !important;
}
.stButton > button:hover { background: #3451D1 !important; }
 
.stProgress > div > div { background: linear-gradient(90deg, #4361EE, #7C9CFF) !important; border-radius: 99px !important; }
.stProgress > div { background: #E2E5F0 !important; border-radius: 99px !important; }
 
.stat-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(130px, 1fr)); gap: 0.75rem; margin: 1.5rem 0; }
.stat-card { background: white; border: 1.5px solid #E8EAF2; border-radius: 12px; padding: 1rem 1.1rem; text-align: center; }
.stat-card .stat-val { font-size: 1.5rem; font-weight: 700; color: #1A1D2E; display: block; line-height: 1.2; }
.stat-card .stat-lbl { font-size: 0.72rem; color: #9CA3AF; font-weight: 500; text-transform: uppercase; letter-spacing: 0.07em; margin-top: 0.2rem; display: block; }
 
.success-bar {
    background: #ECFDF5; border: 1.5px solid #A7F3D0; border-radius: 10px;
    padding: 0.85rem 1.2rem; color: #065F46; font-weight: 600; font-size: 0.9rem;
    display: flex; align-items: center; gap: 0.6rem; margin-bottom: 1rem;
}
 
.stDownloadButton > button {
    background: white !important; color: #1A1D2E !important;
    border: 1.5px solid #D1D5E8 !important; border-radius: 10px !important;
    padding: 0.55rem 1.2rem !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important; font-size: 0.88rem !important; width: 100%;
}
.stDownloadButton > button:hover { border-color: #4361EE !important; color: #4361EE !important; background: #EEF2FF !important; }
 
.streamlit-expanderHeader {
    background: white !important; border: 1.5px solid #E8EAF2 !important;
    border-radius: 10px !important;
    font-family: 'Plus Jakarta Sans', sans-serif !important;
    font-weight: 600 !important; font-size: 0.88rem !important; color: #374151 !important;
}
 
.cat-row { display: flex; align-items: center; justify-content: space-between; padding: 0.6rem 0; border-bottom: 1px solid #F3F4F6; font-size: 0.88rem; }
.cat-row:last-child { border-bottom: none; }
.cat-pill { display: inline-flex; align-items: center; gap: 0.35rem; background: #F3F4F6; border-radius: 6px; padding: 0.2rem 0.65rem; font-size: 0.78rem; font-weight: 600; color: #374151; }
.cat-amount { font-weight: 700; color: #EF4444; font-size: 0.88rem; }
 
.footer { text-align: center; color: #D1D5DB; font-size: 0.75rem; padding: 2rem 0 1rem; margin-top: 3rem; border-top: 1px solid #F3F4F6; }
</style>
""", unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
 
CATEGORY_ICONS = {
    "Food & Dining":   "🍽️",
    "Groceries":       "🛒",
    "Transport":       "🚗",
    "Shopping":        "🛍️",
    "Utilities":       "💡",
    "Health":          "💊",
    "Entertainment":   "🎬",
    "Education":       "📚",
    "Finance":         "🏦",
    "Miscellaneous":   "📦",
}
 
def fmt_inr(val):
    try:
        val = float(val)
        if val >= 1_00_000:
            return f"₹{val/1_00_000:.1f}L"
        if val >= 1_000:
            return f"₹{val/1_000:.1f}K"
        return f"₹{val:,.0f}"
    except Exception:
        return "₹0"
 
def step_indicator(current: int):
    steps = ["Upload", "Analyzing", "Results"]
    parts = []
    for i, label in enumerate(steps, start=1):
        if i < current:
            circ = '<div class="step-circle done">✓</div>'
            lbl  = f'<span class="step-label done">{label}</span>'
        elif i == current:
            circ = f'<div class="step-circle active">{i}</div>'
            lbl  = f'<span class="step-label active">{label}</span>'
        else:
            circ = f'<div class="step-circle">{i}</div>'
            lbl  = f'<span class="step-label">{label}</span>'
        parts.append(f'<div class="step-item">{circ}{lbl}</div>')
        if i < len(steps):
            line_cls = "done" if i < current else ""
            parts.append(f'<div class="step-line {line_cls}"></div>')
    st.markdown(f'<div class="steps">{"".join(parts)}</div>', unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────────
# Top bar
# ─────────────────────────────────────────────
st.markdown("""
<div class="topbar">
    <div class="topbar-logo">💰 Statement<span class="dot">Sense</span></div>
    <div class="topbar-tag">SBI · Free</div>
</div>
""", unsafe_allow_html=True)
 
 
# ─────────────────────────────────────────────
# State
# ─────────────────────────────────────────────
if "results" not in st.session_state:
    st.session_state.results = None
 
 
# ─────────────────────────────────────────────
# Step 1 — Upload
# ─────────────────────────────────────────────
if st.session_state.results is None:
    step_indicator(1)
 
    st.markdown("""
    <div class="section-head">
        <h2>Upload your bank statements</h2>
        <p>Works with any SBI account statement PDF — single month or the whole year.</p>
    </div>
    """, unsafe_allow_html=True)
 
    uploaded_files = st.file_uploader(
        "Drop PDFs here or click to browse",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )
 
    if uploaded_files:
        badges = "".join(
            f'<span class="file-badge">📄 {f.name}</span>'
            for f in uploaded_files
        )
        st.markdown(f"<div style='margin:0.8rem 0'>{badges}</div>", unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="tip-box">
            <strong>Where to get your statement?</strong><br>
            Log in to SBI YONO or Net Banking → Accounts → Statement → Download PDF.
            Password-protected PDFs are not supported yet.
        </div>
        """, unsafe_allow_html=True)
 
    st.markdown("<div style='height:1.2rem'></div>", unsafe_allow_html=True)
 
    if uploaded_files:
        if st.button(f"✨  Analyze {len(uploaded_files)} statement{'s' if len(uploaded_files) > 1 else ''}"):
 
            step_indicator(2)
            progress = st.progress(0, text="Starting…")
            status   = st.empty()
 
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
 
                status.markdown("<p style='color:#6B7280;font-size:0.85rem;text-align:center'>Saving files…</p>", unsafe_allow_html=True)
                progress.progress(10, text="Saving files…")
                pdf_paths = []
                for f in uploaded_files:
                    dest = tmpdir / f.name
                    dest.write_bytes(f.read())
                    pdf_paths.append(str(dest))
 
                status.markdown("<p style='color:#6B7280;font-size:0.85rem;text-align:center'>Merging PDFs…</p>", unsafe_allow_html=True)
                progress.progress(25, text="Merging PDFs…")
                merged_pdf_path = str(tmpdir / "combined_statements.pdf")
                merge_pdfs(pdf_paths, merged_pdf_path)
 
                all_dfs = []
                for i, pdf_path in enumerate(pdf_paths):
                    pct = 30 + int(40 * (i + 1) / len(pdf_paths))
                    msg = f"Reading file {i+1} of {len(pdf_paths)}…"
                    status.markdown(f"<p style='color:#6B7280;font-size:0.85rem;text-align:center'>{msg}</p>", unsafe_allow_html=True)
                    progress.progress(pct, text=msg)
                    df = process_pdf(pdf_path)
                    if not df.empty:
                        all_dfs.append(df)
 
                if not all_dfs:
                    progress.empty()
                    status.empty()
                    st.error("❌  No transactions found. Make sure the PDF is a text-based SBI statement (not a scanned image).")
                    st.stop()
 
                status.markdown("<p style='color:#6B7280;font-size:0.85rem;text-align:center'>Categorizing transactions…</p>", unsafe_allow_html=True)
                progress.progress(75, text="Categorizing transactions…")
                merged_df = pd.concat(all_dfs, ignore_index=True)
                merged_df["category"] = merged_df["merchant"].apply(classify_transaction)
                summary = build_summary(merged_df)
 
                status.markdown("<p style='color:#6B7280;font-size:0.85rem;text-align:center'>Building your Excel report…</p>", unsafe_allow_html=True)
                progress.progress(90, text="Building Excel report…")
                excel_path = str(tmpdir / "transaction_summary.xlsx")
                write_excel(summary, merged_df, excel_path)
 
                progress.progress(100, text="Done!")
                status.empty()
                progress.empty()
 
                zip_buf = io.BytesIO()
                excel_bytes = Path(excel_path).read_bytes()
                pdf_bytes   = Path(merged_pdf_path).read_bytes()
                with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr("transaction_summary.xlsx", excel_bytes)
                    zf.writestr("combined_statements.pdf",  pdf_bytes)
                zip_buf.seek(0)
 
                st.session_state.results = {
                    "df":          merged_df,
                    "summary":     summary,
                    "excel_bytes": excel_bytes,
                    "pdf_bytes":   pdf_bytes,
                    "zip_bytes":   zip_buf.getvalue(),
                }
                st.rerun()
 
 
# ─────────────────────────────────────────────
# Step 3 — Results
# ─────────────────────────────────────────────
else:
    res       = st.session_state.results
    df        = res["df"]
    debit_df  = df[df["txn_type"] == "Debit"]
    credit_df = df[df["txn_type"] == "Credit"]
    months    = df["month_name"].nunique()
 
    step_indicator(3)
 
    st.markdown(f"""
    <div class="success-bar">
        ✅ &nbsp; Done! Found <strong>{len(df)} transactions</strong> across <strong>{months} month{'s' if months != 1 else ''}</strong>.
    </div>
    """, unsafe_allow_html=True)
 
    # Stat cards
    st.markdown(f"""
    <div class="stat-grid">
        <div class="stat-card">
            <span class="stat-val">{len(df)}</span>
            <span class="stat-lbl">Transactions</span>
        </div>
        <div class="stat-card">
            <span class="stat-val" style="color:#EF4444">{fmt_inr(debit_df['amount'].sum())}</span>
            <span class="stat-lbl">Total Spent</span>
        </div>
        <div class="stat-card">
            <span class="stat-val" style="color:#10B981">{fmt_inr(credit_df['amount'].sum())}</span>
            <span class="stat-lbl">Total Received</span>
        </div>
        <div class="stat-card">
            <span class="stat-val">{df['merchant'].nunique()}</span>
            <span class="stat-lbl">Merchants</span>
        </div>
        <div class="stat-card">
            <span class="stat-val">{months}</span>
            <span class="stat-lbl">Month{'s' if months != 1 else ''}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    # Downloads
    st.markdown("<p style='font-weight:600; font-size:0.9rem; color:#374151; margin:1.2rem 0 0.6rem'>Download your report</p>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("📊  Excel Report", res["excel_bytes"],
            "transaction_summary.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.download_button("📄  Merged PDF", res["pdf_bytes"],
            "combined_statements.pdf", "application/pdf")
    with c3:
        st.download_button("📦  Download All", res["zip_bytes"],
            "statementsense_report.zip", "application/zip")
 
    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
 
    # Category breakdown
    with st.expander("📂  Where did your money go?", expanded=True):
        cat_grp = (
            debit_df.groupby("category")["amount"].sum()
            .sort_values(ascending=False).reset_index()
        )
        total_spend = debit_df["amount"].sum()
        rows_html = ""
        for _, row in cat_grp.iterrows():
            icon = CATEGORY_ICONS.get(row["category"], "📦")
            pct  = (row["amount"] / total_spend * 100) if total_spend else 0
            rows_html += f"""
            <div class="cat-row">
                <span class="cat-pill">{icon} {row['category']}</span>
                <div style="flex:1;margin:0 1rem;background:#F3F4F6;border-radius:99px;height:6px;overflow:hidden">
                    <div style="width:{pct:.1f}%;background:#4361EE;height:100%;border-radius:99px"></div>
                </div>
                <span class="cat-amount">- {fmt_inr(row['amount'])}</span>
            </div>"""
        st.markdown(
            f"<div style='background:white;border:1.5px solid #E8EAF2;border-radius:12px;padding:0.8rem 1.1rem'>{rows_html}</div>",
            unsafe_allow_html=True)
 
    # Top merchants
    with st.expander("🏪  Top 10 merchants by spend"):
        top = (
            debit_df.groupby("merchant")["amount"].sum()
            .sort_values(ascending=False).head(10).reset_index()
        )
        top.columns = ["Merchant", "Total Spent"]
        top["Total Spent"] = top["Total Spent"].apply(fmt_inr)
        st.dataframe(top, use_container_width=True, hide_index=True)
 
    # Month-wise (only if multi-month)
    if months > 1:
        with st.expander("📅  Month-wise spending"):
            mth = (
                debit_df.groupby("month_name")["amount"].sum()
                .reset_index().rename(columns={"amount": "Spent", "month_name": "Month"})
            )
            mth["Spent"] = mth["Spent"].apply(fmt_inr)
            st.dataframe(mth, use_container_width=True, hide_index=True)
 
    st.markdown("<div style='height:1rem'></div>", unsafe_allow_html=True)
    if st.button("🔄  Analyze another statement"):
        st.session_state.results = None
        st.rerun()
 
 
# ─────────────────────────────────────────────
# Footer
# ─────────────────────────────────────────────
st.markdown("""
<div class="footer">
    Your files are processed locally and never stored &nbsp;·&nbsp; SBI statements only (for now)
</div>
""", unsafe_allow_html=True)