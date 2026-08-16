# =========================================================
# DAILY REPORT RFL — COMPLETE EXECUTIVE CONSOLE
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

# =========================================================
# DATA PARSING ENGINE
# =========================================================
EXCEL_SIZES = ["160", "90", "120", "250", "270", "280", "380", "330", "470", "530", "800", "428"]


@st.cache_data
def load_and_parse_data(file_bytes):
    """Parses raw Excel workbook containing 'Details'."""
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
    """Computes Machine Size performance table with active-shift capacity scaling."""
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

        if tot_runtime > 0:
            avg_ct = (grp["CT"] * grp["Total Run Time"]).sum() / tot_runtime
        else:
            avg_ct = grp["CT"].mean()

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
    """Computes Shift A vs Shift B labor productivity and scrap rate breakdown."""
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
            "HR Count (Persons)": f"{day_hr} Nos",
            "Total Output (Pcs)": f"{int(round(a_good + a_bad)):,}",
            "Good Output (Pcs)": f"{int(round(a_good)):,}",
            "Rejection (Pcs)": f"{int(round(a_bad)):,}",
            "Rejection Rate": f"{a_rej_rate:.2f}%",
            "Good Tonnage": f"{a_ton:.3f} Tons",
            "Per HR Good Output": f"{a_per_hr_pcs:,.2f} Pcs/HR",
            "Per HR Tonnage": f"{a_per_hr_kg:.2f} kg/HR",
            "per_hr_pcs_raw": a_per_hr_pcs,
            "per_hr_kg_raw": a_per_hr_kg,
            "rej_raw": a_rej_rate,
            "bad_raw": a_bad,
        },
        {
            "Shift Name": "Night Shift",
            "HR Count (Persons)": f"{night_hr} Nos",
            "Total Output (Pcs)": f"{int(round(b_good + b_bad)):,}",
            "Good Output (Pcs)": f"{int(round(b_good)):,}",
            "Rejection (Pcs)": f"{int(round(b_bad)):,}",
            "Rejection Rate": f"{b_rej_rate:.2f}%",
            "Good Tonnage": f"{b_ton:.3f} Tons",
            "Per HR Good Output": f"{b_per_hr_pcs:,.2f} Pcs/HR",
            "Per HR Tonnage": f"{b_per_hr_kg:.2f} kg/HR",
            "per_hr_pcs_raw": b_per_hr_pcs,
            "per_hr_kg_raw": b_per_hr_kg,
            "rej_raw": b_rej_rate,
            "bad_raw": b_bad,
        },
        {
            "Shift Name": "Total / Overall",
            "HR Count (Persons)": f"{tot_hr} Nos",
            "Total Output (Pcs)": f"{int(round(tot_output)):,}",
            "Good Output (Pcs)": f"{int(round(tot_good)):,}",
            "Rejection (Pcs)": f"{int(round(tot_bad)):,}",
            "Rejection Rate": f"{tot_rej_rate:.2f}%",
            "Good Tonnage": f"{tot_ton:.3f} Tons",
            "Per HR Good Output": f"{tot_per_hr_pcs:,.2f} Pcs/HR",
            "Per HR Tonnage": f"{tot_per_hr_kg:.2f} kg/HR",
            "per_hr_pcs_raw": tot_per_hr_pcs,
            "per_hr_kg_raw": tot_per_hr_kg,
            "rej_raw": tot_rej_rate,
            "bad_raw": tot_bad,
        },
    ]
    return pd.DataFrame(records)


# =========================================================
# SECTION 1: SIMPLIFIED UPLOAD SCREEN
# =========================================================
if not st.session_state["app_launched"]:
    st.markdown("## 📊 **DAILY REPORT RFL SETUP**")
    st.markdown("##### Upload your production workbook to launch the live console.")
    st.divider()

    st.markdown(
        '<div style="background:#ffffff; padding:1.75rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #2563eb; max-width:850px; margin:0 auto; box-shadow: 0 4px 12px rgba(15,23,42,0.05);">'
        '<h3 style="margin-top:0; color:#0f172a;">📂 Upload Production Workbook</h3>'
        '<p style="color:#64748b !important;">Select the Excel file (.xlsx) containing the production logs</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)

    col_up, _ = st.columns([2, 1])
    with col_up:
        uploaded_file = st.file_uploader("Select Excel File (.xlsx, .xls)", type=["xlsx", "xls"], key="daily_upload")

        if uploaded_file is not None:
            if st.button("🚀 Launch Executive Dashboard", type="primary", use_container_width=True):
                st.session_state["file_bytes"] = uploaded_file.getvalue()
                st.session_state["file_name"] = uploaded_file.name
                st.session_state["app_launched"] = True
                st.rerun()

# =========================================================
# SECTION 2: FULL EXECUTIVE DASHBOARD VIEW
# =========================================================
else:
    df_details = load_and_parse_data(st.session_state["file_bytes"])
    all_dates = sorted([d for d in df_details["DateClean"].dropna().unique() if d != "nan"])

    # 1. TOP CONTROL BAR
    st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
    c_date, c_day, c_night, c_snap, c_act = st.columns([1.5, 1, 1, 1.2, 0.8], gap="small")

    with c_date:
        sel_date = st.selectbox("📅 **Operational Date**", all_dates, index=len(all_dates) - 1)

    with c_day:
        day_hr = st.number_input("☀️ **Day Shift HR**", min_value=1, value=65, step=1)

    with c_night:
        night_hr = st.number_input("🌙 **Night Shift HR**", min_value=1, value=60, step=1)

    with c_snap:
        st.markdown("<div style='margin-top: 1.65rem;'></div>", unsafe_allow_html=True)
        components.html(
            f"""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <script>
            function captureReport() {{
                const target = window.parent.document.querySelector('#export-report-container');
                html2canvas(target, {{scale: 2, useCORS: true, backgroundColor: '#ffffff'}}).then(canvas => {{
                    const link = document.createElement('a');
                    link.download = 'Daily_Production_Report_{sel_date}.jpg';
                    link.href = canvas.toDataURL('image/jpeg', 0.95);
                    link.click();
                }});
            }}
            </script>
            <button onclick="captureReport()" style="background:#2563eb; color:white; border:none; padding:9px 14px; border-radius:8px; font-weight:700; cursor:pointer; font-family:Inter, sans-serif; font-size:12px; width:100%; box-shadow: 0 2px 6px rgba(37,99,235,0.25);">
                📸 Download JPG
            </button>
            """,
            height=42,
        )

    with c_act:
        st.markdown("<div style='margin-top: 1.65rem;'></div>", unsafe_allow_html=True)
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

    # 2. TOP HEADER BANNER
    st.markdown(
        f"""
        <div class="report-header-banner">
            <div>
                <span style="color: #60a5fa; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.08em; text-transform: uppercase;">✦ OPERATIONAL ANALYTICS - DAILY SUMMARY</span>
                <h2>Daily Production & HR Report</h2>
                <p>Comprehensive Operational Efficiency & Machine Performance Dashboard &nbsp;|&nbsp; 📅 <b>Report Date:</b> {sel_date}</p>
            </div>
            <div class="efficiency-badge-large">
                <div class="value">{overall_eff:.0f}%</div>
                <div class="label">Overall Efficiency</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 3. KPI METRICS CARDS ROW
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(f'<div class="kpi-card blue"><div class="kpi-title">TOTAL PROD</div><div class="kpi-val">{total_prod:,}</div><div class="kpi-sub">Pcs Output</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card purple"><div class="kpi-title">TOTAL CAP</div><div class="kpi-val">{total_cap:,}</div><div class="kpi-sub">Target Pcs</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card yellow"><div class="kpi-title">ACTIVE MC</div><div class="kpi-val">{active_mcs}</div><div class="kpi-sub">Operating MC</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card indigo"><div class="kpi-title">TOTAL HR</div><div class="kpi-val">{total_hr}</div><div class="kpi-sub">Manpower ({day_hr}D + {night_hr}N)</div></div>', unsafe_allow_html=True)
    k5.markdown(f'<div class="kpi-card teal"><div class="kpi-title">HR OUTPUT</div><div class="kpi-val">{int(round(hr_output)):,}</div><div class="kpi-sub">Pcs / Person</div></div>', unsafe_allow_html=True)
    k6.markdown(f'<div class="kpi-card pink"><div class="kpi-title">HR PER MC</div><div class="kpi-val">{hr_per_mc:.1f}</div><div class="kpi-sub">Persons / MC</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom: 1.15rem;'></div>", unsafe_allow_html=True)

    # 4. MID SECTION: SIZE BREAKDOWN (LEFT) + NARRATIVE ANALYSIS (RIGHT)
    col_left, col_right = st.columns([1.45, 1.05], gap="medium")

    running_df = df_size[df_size["Total Prod (Pcs)"] > 0].sort_values("Total Prod (Pcs)", ascending=False)
    top_row = running_df.iloc[0] if not running_df.empty else None

    stopped_mcs = df_size[(df_size["MC QTY"] == 0) | (df_size["Total Prod (Pcs)"] == 0)]["MC Size"].tolist()
    low_hr_mcs = df_size[(df_size["Run Hr Avg"] > 0) & (df_size["Run Hr Avg"] < 10) & (df_size["Total Prod (Pcs)"] > 0)]
    high_ach_mcs = df_size[(df_size["% Achievement"] >= 90.0) | (df_size["Run Hr Avg"] >= 20.0)]

    with col_left:
        st.markdown('<div class="panel-card"><h4>⚙️ MACHINE WISE PRODUCTION BREAKDOWN</h4>', unsafe_allow_html=True)

        subtotal_row = pd.DataFrame([{
            "MC Size": "Sub Total",
            "MC QTY": int(active_mcs),
            "CT Avg": round(avg_ct, 0),
            "Run Hr Avg": round(avg_run_hr, 1),
            "Total Cap (Pcs)": f"{int(total_cap):,}",
            "Total Prod (Pcs)": f"{int(total_prod):,}",
            "Remarks": "-",
            "% Achievement": f"{overall_eff:.0f}%",
        }])

        df_display = df_size[["MC Size", "MC QTY", "CT Avg", "Run Hr Avg", "Total Cap (Pcs)", "Total Prod (Pcs)", "Remarks", "% Achievement"]].copy()
        df_display["Total Cap (Pcs)"] = df_display["Total Cap (Pcs)"].apply(lambda x: f"{x:,}")
        df_display["Total Prod (Pcs)"] = df_display["Total Prod (Pcs)"].apply(lambda x: f"{x:,}")
        df_display["% Achievement"] = df_display["% Achievement"].apply(lambda x: f"{x:.0f}%")

        df_final_table = pd.concat([df_display, subtotal_row], ignore_index=True)

        st.dataframe(df_final_table, use_container_width=True, hide_index=True, height=430)
        st.markdown("</div>", unsafe_allow_html=True)

    with col_right:
        top_contrib_text = ""
        if top_row is not None:
            share_pct = (top_row["Total Prod (Pcs)"] / total_prod * 100) if total_prod > 0 else 0.0
            top_contrib_text = (
                f"MC Size {top_row['MC Size']} generated the highest output of {top_row['Total Prod (Pcs)']:,} Pcs "
                f"(approx. {share_pct:.1f}% of overall factory production) with {top_row['% Achievement']:.0f}% achievement "
                f"and {top_row['Run Hr Avg']} Run Hours Avg."
            )

        areas_improvement = []
        if stopped_mcs:
            areas_improvement.append(f"• MC Sizes {', '.join(stopped_mcs)} were completely stopped (0% achievement).")
        for _, r in low_hr_mcs.iterrows():
            areas_improvement.append(f"• MC Size {r['MC Size']} recorded low run hours ({r['Run Hr Avg']} Hours Avg) and output ({r['% Achievement']:.0f}% achievement).")
        for _, r in high_ach_mcs.iterrows():
            areas_improvement.append(f"• MC Size {r['MC Size']} performed exceptionally well with {r['% Achievement']:.0f}% achievement and {r['Run Hr Avg']} Run Hours.")

        improvement_block_text = "\n".join(areas_improvement) if areas_improvement else "• Operations ran smoothly with no major bottlenecks detected."

        raw_narrative_text = f"""Dear Sir,

🎯 Overall Target Achievement
Total production reached {total_prod:,} Pcs against a target capacity of {total_cap:,} Pcs, achieving an overall plant efficiency of {overall_eff:.0f}% ({total_ton:.2f} Ton produced).

👥 Manpower Productivity (HR Output)
With {total_hr} HR personnel deployed across {active_mcs} active machines, average productivity was {int(round(hr_output)):,} Pcs/person and {hr_per_mc:.1f} HR/machine.

🏆 Top Contributing Machine
{top_contrib_text}

⚠️ Area for Improvement & Highlights
{improvement_block_text}"""

        st.markdown(
            f"""<div class="panel-card">
                <h4>🎯 KEY PERFORMANCE ANALYSIS</h4>
                <div class="narrative-block">
                    <h5>🎯 Overall Target Achievement</h5>
                    <p>Total production reached <b>{total_prod:,} Pcs</b> against a target capacity of <b>{total_cap:,} Pcs</b>, achieving an overall plant efficiency of <b style="color: #10b981;">{overall_eff:.0f}%</b> ({total_ton:.2f} Ton produced).</p>
                    <h5>👥 Manpower Productivity (HR Output)</h5>
                    <p>With <b>{total_hr} HR</b> personnel deployed across <b>{active_mcs} active machines</b>, average productivity was <b>{int(round(hr_output)):,} Pcs/person</b> and <b>{hr_per_mc:.1f} HR/machine</b>.</p>
                    <h5>🏆 Top Contributing Machine</h5>
                    <p>{top_contrib_text}</p>
                    <h5>⚠️ Area for Improvement & Highlights</h5>
                    <p style="white-space: pre-line; margin: 0;">{improvement_block_text}</p>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        with st.expander("📋 Copy Plain Text Report (for WhatsApp / Email)"):
            st.text_area("Report Text", value=raw_narrative_text, height=160, label_visibility="collapsed")

    # 5. BOTTOM SECTION: SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN
    st.markdown('<div class="panel-card"><h4>👥 SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN</h4>', unsafe_allow_html=True)

    table_cols = ["Shift Name", "HR Count (Persons)", "Total Output (Pcs)", "Good Output (Pcs)", "Rejection (Pcs)", "Rejection Rate", "Good Tonnage", "Per HR Good Output", "Per HR Tonnage"]
    st.dataframe(df_shift[table_cols], use_container_width=True, hide_index=True)

    day_row = df_shift.iloc[0]
    night_row = df_shift.iloc[1]
    tot_shift_row = df_shift.iloc[2]

    pcs_diff_pct = ((day_row["per_hr_pcs_raw"] - night_row["per_hr_pcs_raw"]) / night_row["per_hr_pcs_raw"] * 100) if night_row["per_hr_pcs_raw"] > 0 else 0.0
    kg_diff_pct = ((day_row["per_hr_kg_raw"] - night_row["per_hr_kg_raw"]) / night_row["per_hr_kg_raw"] * 100) if night_row["per_hr_kg_raw"] > 0 else 0.0

    c_hl1, c_hl2 = st.columns(2, gap="medium")
    with c_hl1:
        st.markdown(
            f"""<div class="callout-card green">
                <h5>👥 Labor Efficiency Highlights</h5>
                Day Shift achieved <b>{pcs_diff_pct:+.2f}%</b> piece output per HR ({day_row['per_hr_pcs_raw']:,.2f} vs {night_row['per_hr_pcs_raw']:,.2f} Pcs) and <b>{kg_diff_pct:+.2f}%</b> tonnage per HR ({day_row['per_hr_kg_raw']:.2f} vs {night_row['per_hr_kg_raw']:.2f} kg) compared to Night Shift.
            </div>""",
            unsafe_allow_html=True,
        )

    with c_hl2:
        st.markdown(
            f"""<div class="callout-card red">
                <h5>⚠️ Rejection & Scrap Control</h5>
                Overall rejection was maintained at <b>{tot_shift_row['rej_raw']:.2f}%</b> ({int(round(tot_shift_row['bad_raw'])):,} Bad Pcs). Day Shift recorded <b>{day_row['rej_raw']:.2f}%</b> while Night Shift recorded <b>{night_row['rej_raw']:.2f}%</b>.
            </div>""",
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)

    # =========================================================
    # 6. DEDICATED 1-PAGE EXPORT TEMPLATE (Captured by Download Button)
    # =========================================================
    t1_rows_html = ""
    for _, r in df_size.iterrows():
        rem_color = "#dc2626" if r["Remarks"] == "Stopped" else ("#16a34a" if r["Remarks"] == "High Ach." else ("#d97706" if r["Remarks"] == "Low Hours" else "#0f172a"))
        t1_rows_html += f"""
        <tr>
            <td style="padding:4px 6px; border:1px solid #e2e8f0; font-weight:700;">{r['MC Size']}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0;">{r['MC QTY']}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0;">{r['CT Avg']:.0f}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0;">{r['Run Hr Avg']:.1f}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0;">{r['Total Cap (Pcs)']:,}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0;">{r['Total Prod (Pcs)']:,}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0; color:{rem_color}; font-weight:700;">{r['Remarks']}</td>
            <td style="padding:4px 6px; border:1px solid #e2e8f0; font-weight:700;">{r['% Achievement']:.0f}%</td>
        </tr>
        """
    t1_rows_html += f"""
    <tr style="background:#f1f5f9; font-weight:800; border-top:2px solid #cbd5e1;">
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">Sub Total</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">{active_mcs}</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">{avg_ct:.0f}</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">{avg_run_hr:.1f}</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">{int(total_cap):,}</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">{int(total_prod):,}</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">-</td>
        <td style="padding:5px 6px; border:1px solid #e2e8f0;">{overall_eff:.0f}%</td>
    </tr>
    """

    t2_rows_html = ""
    for _, r in df_shift.iterrows():
        is_sub = "background:#f1f5f9; font-weight:800;" if "Total" in r["Shift Name"] else ""
        t2_rows_html += f"""
        <tr style="{is_sub}">
            <td style="padding:5px 6px; border:1px solid #e2e8f0; font-weight:700;">{r['Shift Name']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['HR Count (Persons)']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['Total Output (Pcs)']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['Good Output (Pcs)']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['Rejection (Pcs)']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['Rejection Rate']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['Good Tonnage']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0; font-weight:700; color:#16a34a;">{r['Per HR Good Output']}</td>
            <td style="padding:5px 6px; border:1px solid #e2e8f0;">{r['Per HR Tonnage']}</td>
        </tr>
        """

    # Hidden dedicated capture card for html2canvas
    st.markdown(
        f"""
        <div style="position: absolute; left: -9999px; top: -9999px;">
            <div id="export-report-container" style="width: 1120px; background: #ffffff; padding: 24px; font-family: Inter, sans-serif; color: #0f172a;">
                <!-- Header -->
                <div style="background: linear-gradient(135deg, #091e3a 0%, #102a4e 100%); border-radius: 10px; padding: 14px 20px; color: #ffffff; display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px;">
                    <div>
                        <span style="color: #60a5fa; font-size: 11px; font-weight: 800; text-transform: uppercase;">✦ OPERATIONAL ANALYTICS - DAILY SUMMARY</span>
                        <h2 style="margin: 2px 0 0 0; font-size: 20px; font-weight: 800; color: #ffffff;">Daily Production & HR Report</h2>
                        <p style="margin: 2px 0 0 0; font-size: 12px; color: #94a3b8;">📅 Report Date: <b>{sel_date}</b></p>
                    </div>
                    <div style="background: #10b981; border-radius: 8px; padding: 6px 16px; text-align: center;">
                        <div style="font-size: 22px; font-weight: 900; line-height: 1; color: #ffffff;">{overall_eff:.0f}%</div>
                        <div style="font-size: 9px; font-weight: 700; text-transform: uppercase; color: #ffffff;">Overall Efficiency</div>
                    </div>
                </div>

                <!-- Mid Grid -->
                <div style="display: grid; grid-template-columns: 1.45fr 1fr; gap: 12px; margin-bottom: 12px;">
                    <div style="border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px;">
                        <h4 style="margin: 0 0 8px 0; font-size: 13px; font-weight: 800;">⚙️ MACHINE WISE PRODUCTION BREAKDOWN</h4>
                        <table style="width: 100%; border-collapse: collapse; font-size: 11px; text-align: center;">
                            <thead>
                                <tr style="background: #0f172a; color: #ffffff;">
                                    <th style="padding: 5px;">MC Size</th>
                                    <th style="padding: 5px;">MC QTY</th>
                                    <th style="padding: 5px;">CT Avg</th>
                                    <th style="padding: 5px;">Run Hr</th>
                                    <th style="padding: 5px;">Total Cap</th>
                                    <th style="padding: 5px;">Total Prod</th>
                                    <th style="padding: 5px;">Remarks</th>
                                    <th style="padding: 5px;">% Ach</th>
                                </tr>
                            </thead>
                            <tbody>
                                {t1_rows_html}
                            </tbody>
                        </table>
                    </div>

                    <div style="border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px;">
                        <h4 style="margin: 0 0 8px 0; font-size: 13px; font-weight: 800;">🎯 KEY PERFORMANCE ANALYSIS</h4>
                        <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 10px; font-size: 11.5px; line-height: 1.45;">
                            <h5 style="margin: 0 0 2px 0; font-size: 12px; font-weight: 700;">🎯 Overall Target Achievement</h5>
                            <p style="margin: 0 0 6px 0;">Total production reached <b>{total_prod:,} Pcs</b> against target of <b>{total_cap:,} Pcs</b> (<b style="color: #10b981;">{overall_eff:.0f}% Efficiency</b>, {total_ton:.2f} Ton).</p>
                            
                            <h5 style="margin: 6px 0 2px 0; font-size: 12px; font-weight: 700;">👥 Manpower Productivity (HR Output)</h5>
                            <p style="margin: 0 0 6px 0;">With <b>{total_hr} HR</b> on <b>{active_mcs} MCs</b>: <b>{int(round(hr_output)):,} Pcs/person</b> & <b>{hr_per_mc:.1f} HR/machine</b>.</p>
                            
                            <h5 style="margin: 6px 0 2px 0; font-size: 12px; font-weight: 700;">🏆 Top Contributing Machine</h5>
                            <p style="margin: 0 0 6px 0;">{top_contrib_text}</p>
                            
                            <h5 style="margin: 6px 0 2px 0; font-size: 12px; font-weight: 700;">⚠️ Area for Improvement & Highlights</h5>
                            <p style="margin: 0; white-space: pre-line;">{improvement_block_text}</p>
                        </div>
                    </div>
                </div>

                <!-- Bottom Section -->
                <div style="border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px;">
                    <h4 style="margin: 0 0 8px 0; font-size: 13px; font-weight: 800;">👥 SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN</h4>
                    <table style="width: 100%; border-collapse: collapse; font-size: 11px; text-align: center;">
                        <thead>
                            <tr style="background: #0f172a; color: #ffffff;">
                                <th style="padding: 5px;">Shift Name</th>
                                <th style="padding: 5px;">HR Count</th>
                                <th style="padding: 5px;">Total Output</th>
                                <th style="padding: 5px;">Good Output</th>
                                <th style="padding: 5px;">Rejection</th>
                                <th style="padding: 5px;">Rejection %</th>
                                <th style="padding: 5px;">Good Ton</th>
                                <th style="padding: 5px;">Per HR Output</th>
                                <th style="padding: 5px;">Per HR Tonnage</th>
                            </tr>
                        </thead>
                        <tbody>
                            {t2_rows_html}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
