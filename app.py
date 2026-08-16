# =========================================================
# OPERATIONS CONSOLE — FULL-WIDTH ENTERPRISE PLATFORM
# =========================================================
import io
import os
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches

st.set_page_config(
    page_title="Operations Console | Daily Report RFL",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="collapsed",
)


def load_css(file_name="style.css"):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)


load_css("style.css")

if "active_view" not in st.session_state:
    st.session_state["active_view"] = "hub_home"

EXCEL_SIZES = ["160", "90", "120", "250", "270", "280", "380", "330", "470", "530", "800", "428"]


# =========================================================
# MODULE 1: COMPUTATIONS & BACKEND GRAPHICS ENGINE
# =========================================================
@st.cache_data
def m1_parse_workbook(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)

    target_sheet = "Details" if "Details" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=target_sheet)

    if "Date" in df.columns:
        df["DateClean"] = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%d-%m-%Y")
    else:
        df["DateClean"] = "Default"
    return df


def m1_compute_size_summary(df_day_details):
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
            active_shifts = (1.0 if pd.notna(r.get("A Good")) and r.get("A Good", 0) > 0 else 0.0) + \
                            (1.0 if pd.notna(r.get("B Good")) and r.get("B Good", 0) > 0 else 0.0)
            if active_shifts == 0 and (r.get("T-Bad", 0) > 0):
                active_shifts = 1.0
            return r.get("STD Cap/Shift", 0) * active_shifts

        tot_cap_pcs = grp.apply(calc_row_cap, axis=1).sum()
        tot_prod_pcs = grp["T-Good"].sum()
        cap_ton = ((grp.apply(calc_row_cap, axis=1) * grp.get("Unit Wt", 0)) / 1000.0).sum()
        prod_ton = ((grp["T-Good"] * grp.get("Unit Wt", 0)) / 1000.0).sum()

        ach_pct = (tot_prod_pcs / tot_cap_pcs * 100) if tot_cap_pcs > 0 else 0.0

        if mc_qty == 0 or tot_prod_pcs == 0:
            remarks = "Stopped"
        elif ach_pct >= 90.0 or run_hr_avg >= 20.0:
            remarks = "High Ach."
        elif run_hr_avg < 10.0 and tot_prod_pcs > 0:
            remarks = "Low Hours"
        elif ach_pct < 65.0:
            remarks = "Low Ach."
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


def m1_compute_shiftwise_productivity(df_day_details, day_hr, night_hr):
    a_good = df_day_details["A Good"].sum() if "A Good" in df_day_details.columns else 0.0
    b_good = df_day_details["B Good"].sum() if "B Good" in df_day_details.columns else 0.0

    a_tot = df_day_details["A Total"].fillna(0).sum() if "A Total" in df_day_details.columns else a_good
    b_tot = df_day_details["B Total"].fillna(0).sum() if "B Total" in df_day_details.columns else b_good

    t_bad = df_day_details["T-Bad"].sum() if "T-Bad" in df_day_details.columns else 0.0

    a_bad = max(0.0, a_tot - a_good) if a_tot > 0 else (t_bad * (a_good / (a_good + b_good))) if (a_good + b_good) > 0 else 0.0
    b_bad = max(0.0, b_tot - b_good) if b_tot > 0 else (t_bad - a_bad)

    unit_wt = df_day_details.get("Unit Wt", 0)
    a_ton = ((a_good * unit_wt) / 1000.0).sum()
    b_ton = ((b_good * unit_wt) / 1000.0).sum()

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


def m1_generate_executive_jpg(df_size, total_prod, total_cap, active_mc, total_hr, day_hr, night_hr, hr_output, hr_per_mc, overall_eff, sel_date, top_row, share_pct, stopped_mcs, low_hr_mcs, high_ach_mcs):
    fig, ax = plt.subplots(figsize=(16, 9.8), dpi=220)
    fig.patch.set_facecolor('#f4f7fc')
    ax.set_facecolor('#f4f7fc')
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Banner
    banner = patches.FancyBboxPatch((2, 85), 96, 12.5, boxstyle="round,pad=0.3,rounding_size=1.2", facecolor='#091e3a', edgecolor='none')
    ax.add_patch(banner)
    ax.text(4, 94.5, "OPERATIONAL ANALYTICS - DAILY SUMMARY", color='#5ba4fc', fontsize=10, fontweight='bold')
    ax.text(4, 90.5, "Daily Production & HR Report", color='#ffffff', fontsize=19, fontweight='bold')
    ax.text(4, 87.2, f"Comprehensive Operational Efficiency & Machine Performance Dashboard   |   Report Date: {sel_date}", color='#94a3b8', fontsize=9.2)

    # Efficiency Badge
    eff_badge = patches.FancyBboxPatch((82.0, 86.2), 14.0, 10, boxstyle="round,pad=0.2,rounding_size=1", facecolor='#10b981', edgecolor='none')
    ax.add_patch(eff_badge)
    ax.text(89.0, 92.2, f"{overall_eff:.0f}%", color='#ffffff', fontsize=23, fontweight='bold', ha='center', va='center')
    ax.text(89.0, 88.2, "OVERALL EFFICIENCY", color='#ffffff', fontsize=7, fontweight='bold', ha='center', va='center')

    # KPI Cards
    kpi_data = [
        ("TOTAL PROD", f"{total_prod:,}", "Pcs Output", "#2563eb"),
        ("TOTAL CAP", f"{total_cap:,}", "Target Pcs", "#8b5cf6"),
        ("ACTIVE MC", f"{active_mc}", "Operating MC", "#f59e0b"),
        ("TOTAL HR", f"{total_hr}", f"Manpower ({day_hr}D + {night_hr}N)", "#6366f1"),
        ("HR OUTPUT", f"{int(round(hr_output)):,}", "Pcs / Person", "#06b6d4"),
        ("HR PER MC", f"{hr_per_mc:.1f}", "Persons / MC", "#ec4899"),
    ]

    kpi_w = 15.0
    kpi_gap = 1.2
    for i, (title, val, sub, col_bar) in enumerate(kpi_data):
        x0 = 2 + i * (kpi_w + kpi_gap)
        card = patches.FancyBboxPatch((x0, 73.5), kpi_w, 9.5, boxstyle="round,pad=0.2,rounding_size=0.8", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
        ax.add_patch(card)
        top_bar = patches.FancyBboxPatch((x0 + 0.1, 82.2), kpi_w - 0.2, 0.6, boxstyle="round,pad=0.05,rounding_size=0.3", facecolor=col_bar, edgecolor='none')
        ax.add_patch(top_bar)
        ax.text(x0 + kpi_w/2, 80.8, title, color='#64748b', fontsize=7.8, fontweight='bold', ha='center')
        ax.text(x0 + kpi_w/2, 77.2, val, color='#0f172a', fontsize=14, fontweight='bold', ha='center')
        ax.text(x0 + kpi_w/2, 74.8, sub, color='#94a3b8', fontsize=7.0, ha='center')

    # Left & Right Panels
    left_card = patches.FancyBboxPatch((2, 2.5), 57.5, 69.0, boxstyle="round,pad=0.3,rounding_size=1", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
    ax.add_patch(left_card)
    ax.text(4, 68.5, "MACHINE WISE PRODUCTION BREAKDOWN", color='#0f172a', fontsize=11, fontweight='bold')

    right_card = patches.FancyBboxPatch((61.0, 2.5), 37.0, 69.0, boxstyle="round,pad=0.3,rounding_size=1", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
    ax.add_patch(right_card)
    ax.text(63.0, 68.5, "KEY PERFORMANCE ANALYSIS", color='#0f172a', fontsize=11, fontweight='bold')

    col_names = ["MC Size", "MC QTY", "CT Avg", "Run Hr Avg", "Total Cap (Pcs)", "Total Prod (Pcs)", "Remarks", "% Achievement"]
    col_xs = [6.0, 11.5, 17.0, 23.5, 32.5, 41.5, 50.0, 56.5]

    tbl_hdr = patches.Rectangle((3.5, 63.8), 54.5, 3.2, facecolor='#0f172a', edgecolor='none')
    ax.add_patch(tbl_hdr)
    for name, cx in zip(col_names, col_xs):
        ax.text(cx, 65.4, name, color='#ffffff', fontsize=7.5, fontweight='bold', ha='center', va='center')

    row_y = 61.2
    row_step = 4.25

    for r_i, (_, r) in enumerate(df_size.iterrows()):
        bg_c = '#f8fafc' if r_i % 2 == 1 else '#ffffff'
        row_bg = patches.Rectangle((3.5, row_y - 1.5), 54.5, row_step, facecolor=bg_c, edgecolor='none')
        ax.add_patch(row_bg)
        
        rh_bg = patches.Rectangle((20.0, row_y - 1.5), 7.0, row_step, facecolor='#edf4ff', edgecolor='none')
        ax.add_patch(rh_bg)
        ax.plot([3.5, 58.0], [row_y - 1.5, row_y - 1.5], color='#e2e8f0', linewidth=0.6)

        r_data = [
            str(r["MC Size"]),
            str(r["MC QTY"]),
            f"{r['CT Avg']:.0f}",
            f"{r['Run Hr Avg']:.1f}",
            f"{int(r['Total Cap (Pcs)']):,}",
            f"{int(r['Total Prod (Pcs)']):,}",
            str(r["Remarks"]),
            f"{r['% Achievement']:.0f}%"
        ]

        for c_i, (val, cx) in enumerate(zip(r_data, col_xs)):
            font_w = 'bold' if c_i in [0, 7] else 'normal'
            t_col = '#0f172a'
            if c_i == 6:
                if val == "Stopped": t_col = '#ef4444'
                elif "High" in val: t_col = '#10b981'
                elif "Top" in val: t_col = '#2563eb'
                elif "Low" in val: t_col = '#f59e0b'
            elif c_i == 7:
                num = int(val.replace('%',''))
                if num == 0: t_col = '#94a3b8'
                elif num >= 85: t_col = '#10b981'
                elif num >= 70: t_col = '#f59e0b'
                else: t_col = '#ef4444'
            ax.text(cx, row_y + 0.6, val, color=t_col, fontsize=7.6, fontweight=font_w, ha='center', va='center')
        row_y -= row_step

    sub_bg = patches.Rectangle((3.5, row_y - 1.5), 54.5, row_step, facecolor='#f1f5f9', edgecolor='none')
    ax.add_patch(sub_bg)
    active_grp = df_size[df_size["MC QTY"] > 0]
    avg_ct = (active_grp["CT Avg"] * active_grp["MC QTY"]).sum() / active_mc if active_mc > 0 else 0.0
    avg_run_hr = (active_grp["Run Hr Avg"] * active_grp["MC QTY"]).sum() / active_mc if active_mc > 0 else 0.0

    sub_vals = ["Sub Total", str(active_mc), f"{avg_ct:.0f}", f"{avg_run_hr:.1f}", f"{int(total_cap):,}", f"{int(total_prod):,}", "-", f"{overall_eff:.0f}%"]
    for c_i, (val, cx) in enumerate(zip(sub_vals, col_xs)):
        t_col = '#10b981' if c_i == 7 else '#0f172a'
        ax.text(cx, row_y + 0.6, val, color=t_col, fontsize=8.0, fontweight='bold', ha='center', va='center')
    ax.plot([3.5, 58.0], [row_y + row_step - 1.5, row_y + row_step - 1.5], color='#cbd5e1', linewidth=1.2)

    # Narrative on Right
    narr_y = 65.0
    ax.text(63.0, narr_y, "Overall Target Achievement", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.6
    p1 = f"Total production reached {total_prod:,} Pcs against a target capacity of\n{total_cap:,} Pcs, achieving an overall plant efficiency of {overall_eff:.0f}%."
    ax.text(63.0, narr_y, p1, color='#475569', fontsize=7.8, linespacing=1.45, va='top')

    narr_y -= 7.5
    ax.text(63.0, narr_y, "Manpower Productivity (HR Output)", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.6
    p2 = f"With {total_hr} HR personnel deployed ({day_hr} Day + {night_hr} Night) across\n{active_mc} active machines, average productivity was {int(round(hr_output)):,} Pcs/person\nand {hr_per_mc:.1f} HR/machine."
    ax.text(63.0, narr_y, p2, color='#475569', fontsize=7.8, linespacing=1.45, va='top')

    narr_y -= 7.5
    ax.text(63.0, narr_y, "Top Contributing Machine", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.6
    if top_row is not None:
        p3 = f"MC Size {top_row['MC Size']} generated highest output of {top_row['Total Prod (Pcs)']:,} Pcs\n(approx. {share_pct:.1f}% of factory production) with {top_row['% Achievement']:.0f}% achievement and\n{top_row['Run Hr Avg']} Run Hours Avg."
        ax.text(63.0, narr_y, p3, color='#475569', fontsize=7.8, linespacing=1.45, va='top')

    narr_y -= 7.5
    ax.text(63.0, narr_y, "Area for Improvement & Highlights", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.6

    bullets = []
    if stopped_mcs:
        bullets.append((f"• MC Sizes {', '.join(stopped_mcs)} were completely stopped (0% achievement).", '#ef4444'))
    for _, r in low_hr_mcs.iterrows():
        bullets.append((f"• MC Size {r['MC Size']} recorded lower output ({r['% Achievement']:.0f}% achievement,\n  {r['Run Hr Avg']} Run Hours Avg).", '#ef4444'))
    for _, r in high_ach_mcs.iterrows():
        bullets.append((f"• MC Size {r['MC Size']} performed exceptionally well with {r['% Achievement']:.0f}% achievement\n  and {r['Run Hr Avg']} Run Hours.", '#10b981'))

    for b_text, b_col in bullets[:5]:
        ax.text(63.0, narr_y, b_text, color=b_col, fontsize=7.6, linespacing=1.35, va='top')
        narr_y -= 3.6 if '\n' not in b_text else 5.2

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = io.BytesIO()
    plt.savefig(buf, format='jpg', facecolor=fig.get_facecolor(), edgecolor='none', dpi=220)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# FULL-WIDTH VIEW ROUTING
# =========================================================

# ---------------------------------------------------------
# VIEW 1: HUB HOME / OVERVIEW
# ---------------------------------------------------------
if st.session_state["active_view"] == "hub_home":
    c_brand, c_meta = st.columns([3, 1])
    with c_brand:
        st.markdown("## 🏭 **OPERATIONS CONSOLE & REPORTING HUB**")
        st.caption("Centralized Industrial Engineering & Daily Operational Intelligence")
    with c_meta:
        if st.button("🗑️ Reset All Sessions / Clear Cache", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    st.divider()

    st.markdown("### 📋 **Select a Reporting Module to Launch**")

    # Top Row of Module Cards
    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #2563eb;">
                <div>
                    <span style="background: #e0f2fe; color: #0284c7; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Active Module</span>
                    <h3 style="margin-top: 0.75rem;">📊 Size-Wise Performance & HR</h3>
                    <p>
                        Daily factory efficiency briefs, machine-size capacity breakdown, shift manpower productivity indexing (Pcs/HR & kg/HR), WhatsApp copy text, and pixel-perfect 1-page visual report exports.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("🚀 Launch Size-Wise Module", type="primary", use_container_width=True):
            st.session_state["active_view"] = "mod_size_wise"
            st.rerun()

    with c2:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #8b5cf6;">
                <div>
                    <span style="background: #f3e8ff; color: #7c3aed; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Workspace Ready</span>
                    <h3 style="margin-top: 0.75rem;">⏱️ Mold Changeover & Downtime</h3>
                    <p>
                        Track SMED mold changeover benchmarks, mechanical toggle system conversions vs hydraulic cylinders, tool setup times, and line stoppage losses.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⚙️ Launch Changeover Module", use_container_width=True):
            st.session_state["active_view"] = "mod_changeover"
            st.rerun()

    st.markdown("<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True)

    # Bottom Row of Module Cards
    c3, c4 = st.columns(2, gap="large")

    with c3:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #ef4444;">
                <div>
                    <span style="background: #fee2e2; color: #dc2626; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Workspace Ready</span>
                    <h3 style="margin-top: 0.75rem;">📉 Daily Scrap & Defect Analytics</h3>
                    <p>
                        Analyze item-level rejection quantities, resin purge losses, masterbatch color change scrap, and Six Sigma quality defect distributions.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⚙️ Launch Scrap Module", use_container_width=True):
            st.session_state["active_view"] = "mod_scrap"
            st.rerun()

    with c4:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #10b981;">
                <div>
                    <span style="background: #dcfce7; color: #16a34a; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Workspace Ready</span>
                    <h3 style="margin-top: 0.75rem;">📈 Monthly Trends & OEE Analytics</h3>
                    <p>
                        Month-to-Date (MTD) cumulative production curves, availability / performance / quality OEE factoring, and plant capacity utilization insights.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⚙️ Launch OEE Module", use_container_width=True):
            st.session_state["active_view"] = "mod_oee"
            st.rerun()


# ---------------------------------------------------------
# VIEW 2: MODULE 1 — SIZE-WISE PERFORMANCE & HR
# ---------------------------------------------------------
elif st.session_state["active_view"] == "mod_size_wise":

    # Top Navigation Breadcrumb Bar
    c_back, c_title, c_act = st.columns([1.2, 3, 1.2], vertical_alignment="center")
    with c_back:
        if st.button("⬅️ Back to Operations Hub", use_container_width=True):
            st.session_state["active_view"] = "hub_home"
            st.rerun()
    with c_title:
        st.markdown("<h3 style='margin:0; text-align:center; font-weight:800; color:#0f172a;'>📊 SIZE-WISE PERFORMANCE & HR CONSOLE</h3>", unsafe_allow_html=True)
    with c_act:
        if "m1_file_bytes" in st.session_state:
            if st.button("🔄 Change Excel File", use_container_width=True):
                st.session_state.pop("m1_file_bytes", None)
                st.session_state.pop("m1_file_name", None)
                st.rerun()

    st.divider()

    # Step A: Ingestion if no file uploaded
    if "m1_file_bytes" not in st.session_state:
        c_up, _ = st.columns([2, 1])
        with c_up:
            st.markdown(
                '<div style="background:#ffffff; padding:1.75rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #2563eb; box-shadow: 0 4px 12px rgba(15,23,42,0.05);">'
                '<h3 style="margin-top:0; color:#0f172a;">📂 Upload Daily Production Workbook</h3>'
                '<p style="color:#64748b !important;">Select the Excel workbook (.xlsx, .xls) containing the production details.</p></div>',
                unsafe_allow_html=True,
            )
            st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)

            uploaded_file = st.file_uploader("Select Excel File (.xlsx, .xls)", type=["xlsx", "xls"], key="m1_uploader")

            if uploaded_file is not None:
                if st.button("🚀 Ingest Workbook & Generate Dashboard", type="primary", use_container_width=True):
                    st.session_state["m1_file_bytes"] = uploaded_file.getvalue()
                    st.session_state["m1_file_name"] = uploaded_file.name
                    st.rerun()

    # Step B: Live Full-Width Workspace
    else:
        df_details = m1_parse_workbook(st.session_state["m1_file_bytes"])
        all_dates = sorted([d for d in df_details["DateClean"].dropna().unique() if d != "nan"])

        # Control Bar
        st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
        c_date, c_day, c_night, c_snap = st.columns([1.5, 1, 1, 1.4], gap="small")

        with c_date:
            sel_date = st.selectbox("📅 **Operational Date**", all_dates, index=len(all_dates) - 1)

        with c_day:
            day_hr = st.number_input("☀️ **Day Shift HR**", min_value=1, value=73, step=1)

        with c_night:
            night_hr = st.number_input("🌙 **Night Shift HR**", min_value=1, value=61, step=1)

        # Process metrics
        df_day = df_details[df_details["DateClean"] == sel_date].copy()
        df_size = m1_compute_size_summary(df_day)
        df_shift = m1_compute_shiftwise_productivity(df_day, day_hr, night_hr)

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

        running_df = df_size[df_size["Total Prod (Pcs)"] > 0].sort_values("Total Prod (Pcs)", ascending=False)
        top_row = running_df.iloc[0] if not running_df.empty else None
        share_pct = (top_row["Total Prod (Pcs)"] / total_prod * 100) if (top_row is not None and total_prod > 0) else 0.0

        stopped_mcs = df_size[(df_size["MC QTY"] == 0) | (df_size["Total Prod (Pcs)"] == 0)]["MC Size"].tolist()
        low_hr_mcs = df_size[(df_size["Run Hr Avg"] > 0) & (df_size["Run Hr Avg"] < 14) & (df_size["% Achievement"] < 70) & (df_size["Total Prod (Pcs)"] > 0)]
        high_ach_mcs = df_size[(df_size["% Achievement"] >= 84.0) | (df_size["Run Hr Avg"] >= 20.0)]

        # Generate Image Bytes in backend
        jpg_bytes = m1_generate_executive_jpg(
            df_size, total_prod, total_cap, active_mcs, total_hr, day_hr, night_hr,
            hr_output, hr_per_mc, overall_eff, sel_date, top_row, share_pct,
            stopped_mcs, low_hr_mcs, high_ach_mcs
        )

        with c_snap:
            st.markdown("<div style='margin-top: 1.65rem;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📸 Download 1-Page JPG",
                data=jpg_bytes,
                file_name=f"Daily_Production_Report_{sel_date}.jpg",
                mime="image/jpeg",
                use_container_width=True,
            )

        st.markdown("</div>", unsafe_allow_html=True)

        # Header Banner
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

        # KPI Metrics Cards
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f'<div class="kpi-card blue"><div class="kpi-title">TOTAL PROD</div><div class="kpi-val">{total_prod:,}</div><div class="kpi-sub">Pcs Output</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card purple"><div class="kpi-title">TOTAL CAP</div><div class="kpi-val">{total_cap:,}</div><div class="kpi-sub">Target Pcs</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card yellow"><div class="kpi-title">ACTIVE MC</div><div class="kpi-val">{active_mcs}</div><div class="kpi-sub">Operating MC</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="kpi-card indigo"><div class="kpi-title">TOTAL HR</div><div class="kpi-val">{total_hr}</div><div class="kpi-sub">Manpower ({day_hr}D + {night_hr}N)</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="kpi-card teal"><div class="kpi-title">HR OUTPUT</div><div class="kpi-val">{int(round(hr_output)):,}</div><div class="kpi-sub">Pcs / Person</div></div>', unsafe_allow_html=True)
        k6.markdown(f'<div class="kpi-card pink"><div class="kpi-title">HR PER MC</div><div class="kpi-val">{hr_per_mc:.1f}</div><div class="kpi-sub">Persons / MC</div></div>', unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom: 1.15rem;'></div>", unsafe_allow_html=True)

        # Mid Section: Table (Left) + Analysis (Right)
        col_left, col_right = st.columns([1.45, 1.05], gap="medium")

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
                top_contrib_text = (
                    f"MC Size {top_row['MC Size']} generated the highest output of {top_row['Total Prod (Pcs)']:,} Pcs "
                    f"(approx. {share_pct:.1f}% of overall factory production) with {top_row['% Achievement']:.0f}% achievement "
                    f"and {top_row['Run Hr Avg']} Run Hours Avg."
                )

            areas_improvement = []
            if stopped_mcs:
                areas_improvement.append(f"• MC Sizes {', '.join(stopped_mcs)} were completely stopped (0% achievement).")
            for _, r in low_hr_mcs.iterrows():
                areas_improvement.append(f"• MC Size {r['MC Size']} recorded lower output achievement ({r['% Achievement']:.0f}% achievement, {r['Run Hr Avg']} Run Hours Avg).")
            for _, r in high_ach_mcs.iterrows():
                areas_improvement.append(f"• MC Size {r['MC Size']} performed exceptionally well with {r['% Achievement']:.0f}% achievement and {r['Run Hr Avg']} Run Hours.")

            improvement_block_text = "\n".join(areas_improvement) if areas_improvement else "• Operations ran smoothly with no major bottlenecks detected."

            raw_narrative_text = f"""Dear Sir,

🎯 Overall Target Achievement
Total production reached {total_prod:,} Pcs against a target capacity of {total_cap:,} Pcs, achieving an overall plant efficiency of {overall_eff:.0f}% ({total_ton:.2f} Ton produced).

👥 Manpower Productivity (HR Output)
With {total_hr} HR personnel deployed ({day_hr} Day + {night_hr} Night) across {active_mcs} active machines, average productivity was {int(round(hr_output)):,} Pcs/person and {hr_per_mc:.1f} HR/machine.

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
                        <p>With <b>{total_hr} HR</b> personnel deployed ({day_hr} Day + {night_hr} Night) across <b>{active_mcs} active machines</b>, average productivity was <b>{int(round(hr_output)):,} Pcs/person</b> and <b>{hr_per_mc:.1f} HR/machine</b>.</p>
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

        # Bottom Section: Shiftwise Breakdown
        st.markdown('<div class="panel-card"><h4>👥 SHIFTWISE PRODUCTIVITY & SCRAP BREAKDOWN</h4>', unsafe_allow_html=True)

        table_cols = ["Shift Name", "HR Count (Persons)", "Total Output (Pcs)", "Good Output (Pcs)", "Rejection (Pcs)", "Rejection Rate", "Good Tonnage", "Per HR Good Output", "Per HR Tonnage"]
        st.dataframe(df_shift[table_cols], use_container_width=True, hide_index=True)

        day_row = df_shift.iloc[0]
        night_row = df_shift.iloc[1]
        tot_shift_row = df_shift.iloc[2]

        pcs_diff_pct = ((night_row["per_hr_pcs_raw"] - day_row["per_hr_pcs_raw"]) / day_row["per_hr_pcs_raw"] * 100) if day_row["per_hr_pcs_raw"] > 0 else 0.0
        kg_diff_pct = ((night_row["per_hr_kg_raw"] - day_row["per_hr_kg_raw"]) / day_row["per_hr_kg_raw"] * 100) if day_row["per_hr_kg_raw"] > 0 else 0.0

        c_hl1, c_hl2 = st.columns(2, gap="medium")
        with c_hl1:
            st.markdown(
                f"""<div class="callout-card green">
                    <h5>👥 Labor Efficiency Highlights</h5>
                    Night Shift achieved <b>{pcs_diff_pct:+.2f}%</b> piece output per HR ({night_row['per_hr_pcs_raw']:,.2f} vs {day_row['per_hr_pcs_raw']:,.2f} Pcs) and <b>{kg_diff_pct:+.2f}%</b> tonnage per HR ({night_row['per_hr_kg_raw']:.2f} vs {day_row['per_hr_kg_raw']:.2f} kg) compared to Day Shift.
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


# ---------------------------------------------------------
# VIEW 3, 4, 5: FUTURE RESERVED WORKSPACES
# ---------------------------------------------------------
elif st.session_state["active_view"] == "mod_changeover":
    if st.button("⬅️ Back to Operations Hub"):
        st.session_state["active_view"] = "hub_home"
        st.rerun()
    st.divider()
    st.markdown("## ⏱️ **MOLD CHANGEOVER & DOWNTIME MODULE**")
    st.info("🛠️ This full-width module workspace is ready for your changeover metrics and dataset structure.")

elif st.session_state["active_view"] == "mod_scrap":
    if st.button("⬅️ Back to Operations Hub"):
        st.session_state["active_view"] = "hub_home"
        st.rerun()
    st.divider()
    st.markdown("## 📉 **DAILY SCRAP & DEFECT ANALYTICS MODULE**")
    st.info("🛠️ This full-width module workspace is ready for your scrap tracking metrics and dataset structure.")

elif st.session_state["active_view"] == "mod_oee":
    if st.button("⬅️ Back to Operations Hub"):
        st.session_state["active_view"] = "hub_home"
        st.rerun()
    st.divider()
    st.markdown("## 📈 **MONTHLY TRENDS & OEE ANALYTICS MODULE**")
    st.info("🛠️ This full-width module workspace is ready for your OEE tracking metrics and dataset structure.")
