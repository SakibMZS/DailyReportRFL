# =========================================================
# DAILY REPORT RFL — ONE-PAGE EXECUTIVE CONSOLE
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
    """Computes Machine Size performance table scaled by active shift capacity."""
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
            "HR Count": f"{day_hr} Nos",
            "Total Output": f"{int(round(a_good + a_bad)):,}",
            "Good Output": f"{int(round(a_good)):,}",
            "Rejection": f"{int(round(a_bad)):,}",
            "Rejection %": f"{a_rej_rate:.2f}%",
            "Good Ton": f"{a_ton:.3f} T",
            "Per HR Output": f"{a_per_hr_pcs:,.1f} Pcs/HR",
            "Per HR Tonnage": f"{a_per_hr_kg:.2f} kg/HR",
            "per_hr_pcs_raw": a_per_hr_pcs,
            "per_hr_kg_raw": a_per_hr_kg,
            "rej_raw": a_rej_rate,
            "bad_raw": a_bad,
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
            "per_hr_pcs_raw": b_per_hr_pcs,
            "per_hr_kg_raw": b_per_hr_kg,
            "rej_raw": b_rej_rate,
            "bad_raw": b_bad,
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
        '<div style="background:#ffffff; padding:1.5rem; border-radius:10px; border:1px solid #e2e8f0; border-top:4px solid #2563eb; max-width:850px; margin:0 auto; box-shadow: 0 2px 8px rgba(15,23,42,0.04);">'
        '<h3 style="margin-top:0; color:#0f172a;">📂 Upload Production Workbook</h3>'
        '<p style="color:#64748b !important;">Select the Excel file (.xlsx) containing the production logs</p></div>',
        unsafe_allow_html=True,
    )
    st.markdown("<div style='margin-bottom: 0.8rem;'></div>", unsafe_allow_html=True)

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
# SECTION 2: ONE-PAGE LIVE DASHBOARD
# =========================================================
else:
    df_details = load_and_parse_data(st.session_state["file_bytes"])
    all_dates = sorted([d for d in df_details["DateClean"].dropna().unique() if d != "nan"])

    # 1. TOP COMPACT CONTROL BAR
    st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
    c_date, c_day, c_night, c_snap, c_act = st.columns([1.4, 1, 1, 1.4, 0.8], gap="small")

    with c_date:
        sel_date = st.selectbox("📅 Operational Date", all_dates, index=len(all_dates) - 1)

    with c_day:
        day_hr = st.number_input("☀️ Day Shift HR", min_value=1, value=65, step=1)

    with c_night:
        night_hr = st.number_input("🌙 Night Shift HR", min_value=1, value=60, step=1)

    with c_snap:
        st.markdown("<div style='margin-top: 1.15rem;'></div>", unsafe_allow_html=True)
        components.html(
            f"""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
            <script>
            function captureReport() {{
                const target = window.parent.document.querySelector('.main .block-container');
                html2canvas(target, {{scale: 2, useCORS: true, backgroundColor: '#f4f7fb'}}).then(canvas => {{
                    const link = document.createElement('a');
                    link.download = 'Daily_Production_Report_{sel_date}.jpg';
                    link.href = canvas.toDataURL('image/jpeg', 0.95);
                    link.click();
                }});
            }}
            function printReport() {{
                window.parent.print();
            }}
            </script>
            <div style="display:flex; gap:6px;">
                <button onclick="captureReport()" style="background:#2563eb; color:white; border:none; padding:7px 10px; border-radius:6px; font-weight:700; cursor:pointer; font-family:Inter,sans-serif; font-size:11px; flex:1; box-shadow: 0 1px 4px rgba(37,99,235,0.25);">
                    📸 Save JPG
                </button>
                <button onclick="printReport()" style="background:#0f172a; color:white; border:none; padding:7px 10px; border-radius:6px; font-weight:700; cursor:pointer; font-family:Inter,sans-serif; font-size:11px; flex:1;">
                    🖨️ PDF / Print
                </button>
            </div>
            """,
            height=34,
        )

    with c_act:
        st.markdown("<div style='margin-top: 1.15rem;'></div>", unsafe_allow_html=True)
        if st.button("🔄 File", use_container_width=True):
            st.session_state["app_launched"] = False
            st.session_state.pop("file_bytes", None)
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    # Filter data for selected date
    df_day = df_details[df_details["DateClean"] == sel_date].copy()
    df_size = compute_daily_size_summary(df_day)
    df_shift = compute_shiftwise_productivity(df_day, day_hr, night_hr)

    # Global KPI Calculations
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
                <span style="color: #60a5fa; font-size: 0.65rem; font-weight: 800; letter-spacing: 0.06em; text-transform: uppercase;">✦ OPERATIONAL ANALYTICS - DAILY SUMMARY</span>
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

    # 3. KPI METRIC CARDS ROW
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.markdown(f'<div class="kpi-card blue"><div class="kpi-title">TOTAL PROD</div><div class="kpi-val">{total_prod:,}</div><div class="kpi-sub">Pcs Output</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card purple"><div class="kpi-title">TOTAL CAP</div><div class="kpi-val">{total_cap:,}</div><div class="kpi-sub">Target Pcs</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card yellow"><div class="kpi-title">ACTIVE MC</div><div class="kpi-val">{active_mcs}</div><div class="kpi-sub">Operating MC</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card indigo"><div class="kpi-title">TOTAL HR</div><div class="kpi-val">{total_hr}</div><div class="kpi-sub">Manpower ({day_hr}D + {night_hr}N)</div></div>', unsafe_allow_html=True)
    k5.markdown(f'<div class="kpi-card teal"><div class="kpi-title">HR OUTPUT</div><div class="kpi-val">{int(round(hr_output)):,}</div><div class="kpi-sub">Pcs / Person</div></div>', unsafe_allow_html=True)
    k6.markdown(f'<div class="kpi-card pink"><div class="kpi-title">HR PER MC</div><div class="kpi-val">{hr_per_mc:.1f}</div><div class="kpi-sub">Persons / MC</div></div>', unsafe_allow_html=True)

    st.markdown("<div style='margin-bottom: 0.5rem;'></div>", unsafe_allow_html=True)

    # 4. MID SECTION: SIZE BREAKDOWN TABLE (LEFT) + NARRATIVE ANALYSIS (RIGHT)
    col_left, col_right = st.columns([1.45, 1.05], gap="small")

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

        st.dataframe(df_final_table, use_container_width=True, hide_index=True, height=280)
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
                    <p style="margin-bottom:0.3rem;">Total production reached <b>{total_prod:,} Pcs</b> against target of <b>{total_cap:,} Pcs</b> (<b style="color: #10b981;">{overall_eff:.0f}% Efficiency</b>, {total_ton:.2f} Ton).</p>
                    <h5>👥 Manpower Productivity (HR Output)</h5>
                    <p style="margin-bottom:0.3rem;">With <b>{total_hr} HR</b> on <b>{active_mcs} MCs</b>: <b>{int(round(hr_output)):,} Pcs/person</b> & <b>{hr_per_mc:.1f} HR/machine</b>.</p>
                    <h5>🏆 Top Contributing Machine</h5>
                    <p style="margin-bottom:0.3rem;">{top_contrib_text}</p>
                    <h5>⚠️ Area for Improvement & Highlights</h5>
                    <p style="white-space: pre-line; margin: 0;">{improvement_block_text}</p>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        with st.expander("📋 Copy Plain Text Report (WhatsApp/Email)"):
            st.text_area("Report Text", value=raw_narrative_text, height=110, label_visibility="collapsed")

    # 5. BOTTOM SECTION: SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN
    st.markdown('<div class="panel-card"><h4>👥 SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN</h4>', unsafe_allow_html=True)

    table_cols = ["Shift Name", "HR Count", "Total Output", "Good Output", "Rejection", "Rejection %", "Good Ton", "Per HR Output", "Per HR Tonnage"]
    st.dataframe(df_shift[table_cols], use_container_width=True, hide_index=True, height=135)

    day_row = df_shift.iloc[0]
    night_row = df_shift.iloc[1]
    tot_shift_row = df_shift.iloc[2]

    pcs_diff_pct = ((day_row["per_hr_pcs_raw"] - night_row["per_hr_pcs_raw"]) / night_row["per_hr_pcs_raw"] * 100) if night_row["per_hr_pcs_raw"] > 0 else 0.0
    kg_diff_pct = ((day_row["per_hr_kg_raw"] - night_row["per_hr_kg_raw"]) / night_row["per_hr_kg_raw"] * 100) if night_row["per_hr_kg_raw"] > 0 else 0.0

    c_hl1, c_hl2 = st.columns(2, gap="small")
    with c_hl1:
        st.markdown(
            f"""<div class="callout-card green">
                <h5>👥 Labor Efficiency Highlights</h5>
                Day Shift achieved <b>{pcs_diff_pct:+.2f}%</b> piece output per HR ({day_row['per_hr_pcs_raw']:,.1f} vs {night_row['per_hr_pcs_raw']:,.1f} Pcs) and <b>{kg_diff_pct:+.2f}%</b> tonnage per HR ({day_row['per_hr_kg_raw']:.2f} vs {night_row['per_hr_kg_raw']:.2f} kg) compared to Night Shift.
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
