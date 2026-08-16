# =========================================================
# DAILY REPORT RFL — CLEAN 1-PAGE EXPORT CONSOLE
# =========================================================
import io
import os
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="Daily Production & HR Report | RFL",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def load_css(file_name="style.css"):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


load_css("style.css")

if "app_launched" not in st.session_state:
    st.session_state["app_launched"] = False

EXCEL_SIZES = ["160", "90", "120", "250", "270", "280", "380", "330", "470", "530", "800", "428"]


@st.cache_data
def load_and_parse_data(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)
    if "Details" in xls.sheet_names:
        df_det = pd.read_excel(xls, sheet_name="Details")
        df_det["DateClean"] = pd.to_datetime(df_det["Date"], errors="coerce").dt.strftime("%d-%m-%Y")
        df_det["DateObj"] = pd.to_datetime(df_det["Date"], errors="coerce")
    else:
        df_det = pd.DataFrame()
    return df_det


def compute_daily_size_summary(df_day_details):
    records = []
    for sz in EXCEL_SIZES:
        grp = df_day_details[df_day_details["Size"].astype(str).str.replace(".0", "", regex=False) == sz]
        if grp.empty:
            records.append({
                "MC Size": sz,
                "MC QTY": 0,
                "CT Avg": 0.0,
                "Run Hr Avg": 0.0,
                "Total Cap (Pcs)": 0,
                "Total Prod (Pcs)": 0,
                "Remarks": "Stopped",
                "% Achievement": 0.0,
                "Cap Ton": 0.0,
                "Prod Ton": 0.0,
            })
            continue

        mc_qty = grp["MC SL"].nunique()
        tot_runtime = grp["Total Run Time"].sum()
        run_hr_avg = tot_runtime / mc_qty if mc_qty > 0 else 0.0
        avg_ct = (grp["CT"] * grp["Total Run Time"]).sum() / tot_runtime if tot_runtime > 0 else grp["CT"].mean()

        def calc_row_cap(r):
            active_shifts = (1.0 if pd.notna(r["A Good"]) and r["A Good"] > 0 else 0.0) + \
                            (1.0 if pd.notna(r["B Good"]) and r["B Good"] > 0 else 0.0)
            if active_shifts == 0 and (r.get("T-Bad", 0) > 0):
                active_shifts = 1.0
            return r["STD Cap/Shift"] * active_shifts

        tot_cap_pcs = grp.apply(calc_row_cap, axis=1).sum()
        tot_prod_pcs = grp["T-Good"].sum()
        cap_ton = ((grp.apply(calc_row_cap, axis=1) * grp["Unit Wt"]) / 1000.0).sum()
        prod_ton = ((grp["T-Good"] * grp["Unit Wt"]) / 1000.0).sum()
        ach_pct = (tot_prod_pcs / tot_cap_pcs * 100) if tot_cap_pcs > 0 else 0.0

        if mc_qty == 0 or tot_prod_pcs == 0:
            remarks = "Stopped"
        elif ach_pct >= 90.0 or run_hr_avg >= 20.0:
            remarks = "High Ach."
        elif run_hr_avg < 10.0 and tot_prod_pcs > 0:
            remarks = "Low Hours"
        else:
            remarks = "-"

        records.append({
            "MC Size": sz,
            "MC QTY": int(mc_qty),
            "CT Avg": round(avg_ct, 1),
            "Run Hr Avg": round(run_hr_avg, 1),
            "Total Cap (Pcs)": int(round(tot_cap_pcs)),
            "Total Prod (Pcs)": int(round(tot_prod_pcs)),
            "Remarks": remarks,
            "% Achievement": round(ach_pct, 1),
            "Cap Ton": round(cap_ton, 2),
            "Prod Ton": round(prod_ton, 2),
        })
    return pd.DataFrame(records)


def compute_shiftwise_productivity(df_day_details, day_hr, night_hr):
    a_good = df_day_details["A Good"].sum()
    b_good = df_day_details["B Good"].sum()
    a_tot = df_day_details["A Total"].fillna(0).sum()
    b_tot = df_day_details["B Total"].fillna(0).sum()

    a_bad = max(0.0, a_tot - a_good) if a_tot > 0 else (df_day_details["T-Bad"].sum() * (a_good / (a_good + b_good))) if (a_good + b_good) > 0 else 0.0
    b_bad = max(0.0, b_tot - b_good) if b_tot > 0 else (df_day_details["T-Bad"].sum() - a_bad)

    a_ton = ((df_day_details["A Good"] * df_day_details["Unit Wt"]) / 1000.0).sum()
    b_ton = ((df_day_details["B Good"] * df_day_details["Unit Wt"]) / 1000.0).sum()

    a_per_hr_pcs = (a_good / day_hr) if day_hr > 0 else 0.0
    b_per_hr_pcs = (b_good / night_hr) if night_hr > 0 else 0.0
    a_per_hr_kg = ((a_ton * 1000.0) / day_hr) if day_hr > 0 else 0.0
    b_per_hr_kg = ((b_ton * 1000.0) / night_hr) if night_hr > 0 else 0.0

    tot_hr = day_hr + night_hr
    tot_good = a_good + b_good
    tot_bad = a_bad + b_bad
    tot_output = tot_good + tot_bad
    tot_ton = a_ton + b_ton

    tot_per_hr_pcs = (tot_good / tot_hr) if tot_hr > 0 else 0.0
    tot_per_hr_kg = ((tot_ton * 1000.0) / tot_hr) if tot_hr > 0 else 0.0

    a_rej_rate = (a_bad / (a_good + a_bad) * 100) if (a_good + a_bad) > 0 else 0.0
    b_rej_rate = (b_bad / (b_good + b_bad) * 100) if (b_good + b_bad) > 0 else 0.0
    tot_rej_rate = (tot_bad / tot_output * 100) if tot_output > 0 else 0.0

    records = [
        {
            "Shift Name": "Day Shift",
            "HR Count": f"{day_hr} Nos",
            "Total Output": f"{int(round(a_good + a_bad)):,}",
            "Good Output": f"{int(round(a_good)):,}",
            "Rejection": f"{int(round(a_bad)):,}",
            "Rejection %": f"{a_rej_rate:.2f}%",
            "Good Ton": f"{a_ton:.3f} T",
            "Per HR Output": f"{a_per_hr_pcs:,.1f} Pcs/HR",
            "Per HR Tonnage": f"{a_per_hr_kg:.2f} kg/HR",
        },
        {
            "Shift Name": "Night Shift",
            "HR Count": f"{night_hr} Nos",
            "Total Output": f"{int(round(b_good + b_bad)):,}",
            "Good Output": f"{int(round(b_good)):,}",
            "Rejection": f"{int(round(b_bad)):,}",
            "Rejection %": f"{b_rej_rate:.2f}%",
            "Good Ton": f"{b_ton:.3f} T",
            "Per HR Output": f"{b_per_hr_pcs:,.1f} Pcs/HR",
            "Per HR Tonnage": f"{b_per_hr_kg:.2f} kg/HR",
        },
        {
            "Shift Name": "Total / Overall",
            "HR Count": f"{tot_hr} Nos",
            "Total Output": f"{int(round(tot_output)):,}",
            "Good Output": f"{int(round(tot_good)):,}",
            "Rejection": f"{int(round(tot_bad)):,}",
            "Rejection %": f"{tot_rej_rate:.2f}%",
            "Good Ton": f"{tot_ton:.3f} T",
            "Per HR Output": f"{tot_per_hr_pcs:,.1f} Pcs/HR",
            "Per HR Tonnage": f"{tot_per_hr_kg:.2f} kg/HR",
        },
    ]
    return pd.DataFrame(records)


# =========================================================
# 1. SETUP / UPLOAD SCREEN
# =========================================================
if not st.session_state["app_launched"]:
    st.markdown("## 📊 **DAILY REPORT RFL SETUP**")
    st.markdown("##### Upload your production workbook to launch.")
    st.divider()

    uploaded_file = st.file_uploader("Select Excel File (.xlsx, .xls)", type=["xlsx", "xls"], key="daily_upload")
    if uploaded_file is not None:
        if st.button("🚀 Launch Executive Report", type="primary", use_container_width=True):
            st.session_state["file_bytes"] = uploaded_file.getvalue()
            st.session_state["app_launched"] = True
            st.rerun()

# =========================================================
# 2. ONE-PAGE REPORT DASHBOARD
# =========================================================
else:
    df_details = load_and_parse_data(st.session_state["file_bytes"])
    all_dates = sorted([d for d in df_details["DateClean"].dropna().unique() if d != "nan"])

    # Top Minimal Control Bar
    st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
    c_date, c_day, c_night, c_snap, c_act = st.columns([1.3, 0.9, 0.9, 1.1, 0.7], gap="small")

    with c_date:
        sel_date = st.selectbox("📅 Date", all_dates, index=len(all_dates) - 1)
    with c_day:
        day_hr = st.number_input("☀️ Day HR", min_value=1, value=65, step=1)
    with c_night:
        night_hr = st.number_input("🌙 Night HR", min_value=1, value=60, step=1)
    with c_snap:
        st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
        components.html(
            f"""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <script>
            function captureReport() {{
                const target = window.parent.document.querySelector('#report-card');
                html2canvas(target, {{scale: 2, useCORS: true, backgroundColor: '#ffffff'}}).then(canvas => {{
                    const link = document.createElement('a');
                    link.download = 'Daily_Production_Report_{sel_date}.jpg';
                    link.href = canvas.toDataURL('image/jpeg', 0.95);
                    link.click();
                }});
            }}
            </script>
            <button onclick="captureReport()" style="background:#2563eb; color:white; border:none; padding:6px 12px; border-radius:6px; font-weight:700; cursor:pointer; font-family:Inter,sans-serif; font-size:11px; width:100%;">
                📸 Download 1-Page JPG
            </button>
            """,
            height=32,
        )
    with c_act:
        st.markdown("<div style='margin-top: 1rem;'></div>", unsafe_allow_html=True)
        if st.button("🔄 File", use_container_width=True):
            st.session_state["app_launched"] = False
            st.session_state.pop("file_bytes", None)
            st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # Process Metrics
    df_day = df_details[df_details["DateClean"] == sel_date].copy()
    df_size = compute_daily_size_summary(df_day)
    df_shift = compute_shiftwise_productivity(df_day, day_hr, night_hr)

    total_prod = df_size["Total Prod (Pcs)"].sum()
    total_cap = df_size["Total Cap (Pcs)"].sum()
    active_mcs = df_size[df_size["MC QTY"] > 0]["MC QTY"].sum()
    total_ton = df_size["Prod Ton"].sum()
    overall_eff = (total_prod / total_cap * 100) if total_cap > 0 else 0.0

    total_hr = day_hr + night_hr
    hr_output = (total_prod / total_hr) if total_hr > 0 else 0.0
    hr_per_mc = (total_hr / active_mcs) if active_mcs > 0 else 0.0

    active_grp = df_size[df_size["MC QTY"] > 0]
    avg_ct = (active_grp["CT Avg"] * active_grp["MC QTY"]).sum() / active_mcs if active_mcs > 0 else 0.0
    avg_run_hr = (active_grp["Run Hr Avg"] * active_grp["MC QTY"]).sum() / active_mcs if active_mcs > 0 else 0.0

    # Build Narrative Highlights
    running_df = df_size[df_size["Total Prod (Pcs)"] > 0].sort_values("Total Prod (Pcs)", ascending=False)
    top_row = running_df.iloc[0] if not running_df.empty else None

    stopped_mcs = df_size[(df_size["MC QTY"] == 0) | (df_size["Total Prod (Pcs)"] == 0)]["MC Size"].tolist()
    low_hr_mcs = df_size[(df_size["Run Hr Avg"] > 0) & (df_size["Run Hr Avg"] < 10) & (df_size["Total Prod (Pcs)"] > 0)]
    high_ach_mcs = df_size[(df_size["% Achievement"] >= 90.0) | (df_size["Run Hr Avg"] >= 20.0)]

    top_contrib_text = ""
    if top_row is not None:
        share_pct = (top_row["Total Prod (Pcs)"] / total_prod * 100) if total_prod > 0 else 0.0
        top_contrib_text = (
            f"MC Size {top_row['MC Size']} generated the highest output of <b>{top_row['Total Prod (Pcs)']:,} Pcs</b> "
            f"(approx. <b>{share_pct:.1f}%</b> of overall factory production) with <b>{top_row['% Achievement']:.0f}%</b> achievement "
            f"and <b>{top_row['Run Hr Avg']} Run Hours Avg</b>."
        )

    areas_improvement = []
    if stopped_mcs:
        areas_improvement.append(f"• MC Sizes {', '.join(stopped_mcs)} were completely stopped (0% achievement).")
    for _, r in low_hr_mcs.iterrows():
        areas_improvement.append(f"• MC Size {r['MC Size']} recorded low run hours ({r['Run Hr Avg']} Hours Avg) and output ({r['% Achievement']:.0f}% achievement).")
    for _, r in high_ach_mcs.iterrows():
        areas_improvement.append(f"• MC Size {r['MC Size']} performed exceptionally well with {r['% Achievement']:.0f}% achievement and {r['Run Hr Avg']} Run Hours.")

    improvement_block_text = "<br>".join(areas_improvement) if areas_improvement else "• Operations ran smoothly with no major bottlenecks detected."

    # Build Table 1 Rows (HTML)
    t1_rows = ""
    for _, r in df_size.iterrows():
        rem_cls = "badge-stopped" if r["Remarks"] == "Stopped" else ("badge-high" if r["Remarks"] == "High Ach." else ("badge-low" if r["Remarks"] == "Low Hours" else ""))
        t1_rows += f"""
        <tr>
            <td><b>{r['MC Size']}</b></td>
            <td>{r['MC QTY']}</td>
            <td>{r['CT Avg']:.0f}</td>
            <td>{r['Run Hr Avg']:.1f}</td>
            <td>{r['Total Cap (Pcs)']:,}</td>
            <td>{r['Total Prod (Pcs)']:,}</td>
            <td class="{rem_cls}">{r['Remarks']}</td>
            <td><b>{r['% Achievement']:.0f}%</b></td>
        </tr>
        """
    t1_rows += f"""
    <tr class="subtotal-row">
        <td>Sub Total</td>
        <td>{active_mcs}</td>
        <td>{avg_ct:.0f}</td>
        <td>{avg_run_hr:.1f}</td>
        <td>{int(total_cap):,}</td>
        <td>{int(total_prod):,}</td>
        <td>-</td>
        <td>{overall_eff:.0f}%</td>
    </tr>
    """

    # Build Table 2 Rows (HTML)
    t2_rows = ""
    for _, r in df_shift.iterrows():
        is_sub = "subtotal-row" if "Total" in r["Shift Name"] else ""
        t2_rows += f"""
        <tr class="{is_sub}">
            <td><b>{r['Shift Name']}</b></td>
            <td>{r['HR Count']}</td>
            <td>{r['Total Output']}</td>
            <td>{r['Good Output']}</td>
            <td>{r['Rejection']}</td>
            <td>{r['Rejection %']}</td>
            <td>{r['Good Ton']}</td>
            <td><b>{r['Per HR Output']}</b></td>
            <td>{r['Per HR Tonnage']}</td>
        </tr>
        """

    # RENDER EXACT 1-PAGE REPORT CONTAINER
    st.markdown(
        f"""
        <div id="report-card">
            <!-- Header -->
            <div class="rep-header">
                <div>
                    <span style="color: #60a5fa; font-size: 0.65rem; font-weight: 800; text-transform: uppercase;">✦ OPERATIONAL ANALYTICS - DAILY SUMMARY</span>
                    <h3>Daily Production & HR Report</h3>
                    <p>Comprehensive Operational Efficiency & Machine Performance Dashboard &nbsp;|&nbsp; 📅 <b>Report Date:</b> {sel_date}</p>
                </div>
                <div class="rep-badge">
                    <div class="val">{overall_eff:.0f}%</div>
                    <div class="lbl">Overall Efficiency</div>
                </div>
            </div>

            <!-- Mid Section: Table 1 (Left) + Key Analysis (Right) -->
            <div style="display: grid; grid-template-columns: 1.4fr 1fr; gap: 10px; margin-bottom: 10px;">
                <div class="section-box">
                    <h4>⚙️ MACHINE WISE PRODUCTION BREAKDOWN</h4>
                    <table class="clean-table">
                        <thead>
                            <tr>
                                <th>MC Size</th>
                                <th>MC QTY</th>
                                <th>CT Avg</th>
                                <th>Run Hr</th>
                                <th>Total Cap (Pcs)</th>
                                <th>Total Prod (Pcs)</th>
                                <th>Remarks</th>
                                <th>% Ach</th>
                            </tr>
                        </thead>
                        <tbody>
                            {t1_rows}
                        </tbody>
                    </table>
                </div>
                <div class="section-box">
                    <h4>🎯 KEY PERFORMANCE ANALYSIS</h4>
                    <div class="narrative-body">
                        <h5>🎯 Overall Target Achievement</h5>
                        <p style="margin:0 0 4px 0;">Total production reached <b>{total_prod:,} Pcs</b> against target of <b>{total_cap:,} Pcs</b> (<b style="color: #10b981;">{overall_eff:.0f}% Efficiency</b>, {total_ton:.2f} Ton).</p>
                        
                        <h5>👥 Manpower Productivity (HR Output)</h5>
                        <p style="margin:0 0 4px 0;">With <b>{total_hr} HR</b> on <b>{active_mcs} MCs</b>: <b>{int(round(hr_output)):,} Pcs/person</b> & <b>{hr_per_mc:.1f} HR/machine</b>.</p>
                        
                        <h5>🏆 Top Contributing Machine</h5>
                        <p style="margin:0 0 4px 0;">{top_contrib_text}</p>
                        
                        <h5>⚠️ Area for Improvement & Highlights</h5>
                        <p style="margin:0;">{improvement_block_text}</p>
                    </div>
                </div>
            </div>

            <!-- Bottom Section: Table 2 -->
            <div class="section-box">
                <h4>👥 SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN</h4>
                <table class="clean-table">
                    <thead>
                        <tr>
                            <th>Shift Name</th>
                            <th>HR Count</th>
                            <th>Total Output</th>
                            <th>Good Output</th>
                            <th>Rejection</th>
                            <th>Rejection %</th>
                            <th>Good Ton</th>
                            <th>Per HR Output</th>
                            <th>Per HR Tonnage</th>
                        </tr>
                    </thead>
                    <tbody>
                        {t2_rows}
                    </tbody>
                </table>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
