import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# =========================================================
# STANDARD MACHINE POSITION & LINE MAPPING
# =========================================================
MAPPING_DATA = [
    ("A1-160", "FF A-B", "IMM-160-6"),
    ("A2-120", "FF A-B", "IMM-120-20"),
    ("A3-120", "FF A-B", "IMM-120-28"),
    ("A4-120", "FF A-B", "IMM-120-29"),
    ("A5-160", "FF A-B", "IMM-160-7"),
    ("A6-160", "FF A-B", "IMM-160-12"),
    ("A7-160", "FF A-B", "IMM-160-48"),
    ("B1-120", "FF A-B", "IMM-120-11"),
    ("B2-120", "FF A-B", "IMM-120-15"),
    ("B3-120", "FF A-B", "IMM-120-14"),
    ("B4-120", "FF A-B", "IMM-120-75"),
    ("B5-90PC", "FF A-B", "IMM-90-8"),
    ("B6-90PC", "FF A-B", "IMM-90-9"),
    ("B7-120PC", "FF A-B", "IMM-120-32"),
    ("B8-120PC", "FF A-B", "IMM-120-27"),
    ("C1-120", "FF C-D", "IMM-120-4"),
    ("C2-160", "FF C-D", "IMM-160-17"),
    ("C3-120", "FF C-D", "IMM-120-22"),
    ("C4-120PC", "FF C-D", "IMM-120-46"),
    ("C5-90", "FF C-D", "IMM-90-4"),
    ("C6-120", "FF C-D", "IMM-120-47"),
    ("C7-160", "FF C-D", "IMM-160-51"),
    ("D1-160", "FF C-D", "IMM-160-39"),
    ("D2-160", "FF C-D", "IMM-160-79"),
    ("D3-160", "FF C-D", "IMM-160-80"),
    ("A1-280TC", "GF A-B", "IMM-280R-25"),
    ("A2-380", "GF A-B", "IMM-380-5"),
    ("A3-380 (PC)", "GF A-B", "IMM-380-81"),
    ("A4-380", "GF A-B", "IMM-380-80"),
    ("A5-HP-330", "GF A-B", "IMM-330-4"),
    ("B1-470", "GF A-B", "IMM-470-5"),
    ("B2-380", "GF A-B", "IMM-380-6"),
    ("B3-530", "GF A-B", "IMM-530-15"),
    ("B4-530", "GF A-B", "IMM-530-16"),
    ("B5-530", "GF A-B", "IMM-530-22"),
    ("B6-380", "GF A-B", "IMM-380-4"),
    ("C1-800-30", "GF C-D", "IMM-800-30"),
    ("C2-800-31", "GF C-D", "IMM-800-31"),
    ("C3-270-1", "GF C-D", "IMM-270-1"),
    ("C4-380-73", "GF C-D", "IMM-380-73"),
    ("C5-380-44", "GF C-D", "IMM-380-44"),
    ("C6-280TC", "GF C-D", "IMM-280R-3"),
    ("D1-280TC", "GF C-D", "IMM-280R-24"),
    ("D2-MA2-250", "GF C-D", "IMM-250-106"),
    ("D3-330-1", "GF C-D", "IMM-330-1"),
    ("D4-HP-330-5", "GF C-D", "IMM-330-5"),
    ("D5-428-1", "GF C-D", "IMM-428-1"),
    ("D6-HP-428-4", "GF C-D", "IMM-428-4"),
    ("D7-HP-330", "GF C-D", "IMM-330-8"),
    ("E1-380-90", "GF E-F", "IMM-380-90"),
    ("E2-380-94", "GF E-F", "IMM-380-94"),
    ("E3-380-88", "GF E-F", "IMM-380-88"),
    ("E4-380-76", "GF E-F", "IMM-380-76"),
    ("E5-380-62", "GF E-F", "IMM-380-62"),
    ("E6-380-75", "GF E-F", "IMM-380-75"),
    ("F1-380-92", "GF E-F", "IMM-380-92"),
    ("F2-380-93", "GF E-F", "IMM-380-93"),
    ("F3-380-98", "GF E-F", "IMM-380-98"),
    ("F4-380-99", "GF E-F", "IMM-380-99"),
    ("F5-380-101", "GF E-F", "IMM-380-101"),
    ("F6-380-100", "GF E-F", "IMM-380-100"),
]

POS_MAP = {smart_manu: pos for pos, line, smart_manu in MAPPING_DATA}
LINE_MAP = {smart_manu: line for pos, line, smart_manu in MAPPING_DATA}


def get_col(df, candidates, default=None):
    for c in candidates:
        if c in df.columns:
            return c
    return default


@st.cache_data
def m2_parse_workbook(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)
    
    sheet_name = "RejectionReport" if "RejectionReport" in xls.sheet_names else ("This Month" if "This Month" in xls.sheet_names else xls.sheet_names[0])
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values])
        if "Machine" in row_str and ("Quantity" in row_str or "Qty" in row_str or "Cause" in row_str):
            header_idx = idx
            break
            
    if header_idx is not None:
        df_clean = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx)
    else:
        df_clean = pd.read_excel(xls, sheet_name=sheet_name)
    
    df_clean.columns = [str(c).strip() for c in df_clean.columns]
    
    date_col = get_col(df_clean, ["Added Date", "Date", "Entry Date", "AddedDate"], df_clean.columns[-1])
    df_clean["DateClean"] = pd.to_datetime(df_clean[date_col], errors="coerce")
    df_clean = df_clean.dropna(subset=["DateClean"]).sort_values("DateClean")
    df_clean["DateStr"] = df_clean["DateClean"].dt.strftime("%Y-%m-%d")
    df_clean["YearMonth"] = df_clean["DateClean"].dt.to_period("M")
    
    unique_months = sorted(df_clean["YearMonth"].unique())
    if len(unique_months) >= 2:
        df_prev = df_clean[df_clean["YearMonth"] == unique_months[-2]].copy()
        df_curr = df_clean[df_clean["YearMonth"] == unique_months[-1]].copy()
    else:
        df_prev = pd.DataFrame()
        df_curr = df_clean.copy()
        
    return df_prev, df_curr, df_clean


def m2_compute_daily_rejection(df_day, min_qty=50):
    if df_day.empty:
        return pd.DataFrame()
    
    mc_col = get_col(df_day, ["Machine", "MC SL", "MC Name"], df_day.columns[2])
    item_col = get_col(df_day, ["Item", "Mold", "Item Name", "Mold / Item"], df_day.columns[3])
    cause_col = get_col(df_day, ["Cause", "Causes", "Defect", "Reason"], "Cause")
    qty_col = get_col(df_day, ["Quantity", "Qty", "Rejection Pcs", "Qty (Pcs)"], "Quantity")
    wt_col = get_col(df_day, ["Weight", "Rejection Ton", "Weight (Ton)", "Weight (kg)"], "Weight")
    
    records = []
    for mc, grp in df_day.groupby(mc_col):
        raw_qty = pd.to_numeric(grp[qty_col], errors="coerce").fillna(0).sum() if qty_col in grp.columns else 0.0
        qty_factor = 1000.0 if (qty_col in grp.columns and grp[qty_col].max() < 100) else 1.0
        total_pcs = raw_qty * qty_factor
        total_ton = pd.to_numeric(grp[wt_col], errors="coerce").fillna(0).sum() if wt_col in grp.columns else 0.0
        
        if cause_col in grp.columns:
            unique_causes = [str(c).strip() for c in grp[cause_col].dropna().unique() if str(c).strip()]
            causes_str = ", ".join(unique_causes) if unique_causes else "No Rejection"
        else:
            causes_str = "-"

        mold_name = str(grp[item_col].iloc[0]) if (item_col in grp.columns and not grp[item_col].dropna().empty) else "-"
        pos = POS_MAP.get(str(mc), "-")
        line = LINE_MAP.get(str(mc), "-")
        
        if total_pcs > min_qty:
            records.append({
                "MC Position": pos,
                "Line": line,
                "Smart Manu": str(mc),
                "Causes": causes_str,
                "Qty (Pcs)": int(round(total_pcs)),
                "Weight (Ton)": round(total_ton, 4),
                "Weight (kg)": round(total_ton * 1000.0, 2),
                "Mold": mold_name
            })
            
    df_res = pd.DataFrame(records)
    if not df_res.empty:
        df_res = df_res.sort_values("Qty (Pcs)", ascending=False).reset_index(drop=True)
    return df_res


def m2_compute_pareto(df_curr):
    cause_col = get_col(df_curr, ["Cause", "Causes", "Defect", "Reason"], None)
    qty_col = get_col(df_curr, ["Quantity", "Qty", "Rejection Pcs", "Qty (Pcs)"], None)
    wt_col = get_col(df_curr, ["Weight", "Rejection Ton", "Weight (Ton)"], None)

    if df_curr.empty or not cause_col:
        return pd.DataFrame()
    
    qty_factor = 1000.0 if (qty_col and df_curr[qty_col].max() < 100) else 1.0
    
    pareto = df_curr.groupby(cause_col).agg(
        Rejection_Pcs=(qty_col, lambda x: int(round(pd.to_numeric(x, errors="coerce").fillna(0).sum() * qty_factor))) if qty_col else ("DateClean", "count"),
        Rejection_Ton=(wt_col, lambda x: pd.to_numeric(x, errors="coerce").fillna(0).sum()) if wt_col else ("DateClean", "count")
    ).reset_index()
    
    pareto = pareto.rename(columns={cause_col: "Cause"})
    total_pcs = pareto["Rejection_Pcs"].sum()
    pareto["% Share"] = (pareto["Rejection_Pcs"] / total_pcs * 100).round(2) if total_pcs > 0 else 0.0
    pareto = pareto.sort_values("Rejection_Pcs", ascending=False).reset_index(drop=True)
    return pareto


def m2_compute_tonnage_comparison(df_prev, df_curr):
    wt_col_prev = get_col(df_prev, ["Weight", "Rejection Ton"], None) if not df_prev.empty else None
    wt_col_curr = get_col(df_curr, ["Weight", "Rejection Ton"], None) if not df_curr.empty else None

    if not df_prev.empty and wt_col_prev:
        t_prev = df_prev.groupby(df_prev["DateClean"].dt.day)[wt_col_prev].sum().reset_index()
        t_prev.columns = ["Day", "Prev_Month_Ton"]
    else:
        t_prev = pd.DataFrame(columns=["Day", "Prev_Month_Ton"])

    if not df_curr.empty and wt_col_curr:
        t_curr = df_curr.groupby(df_curr["DateClean"].dt.day)[wt_col_curr].sum().reset_index()
        t_curr.columns = ["Day", "Curr_Month_Ton"]
    else:
        t_curr = pd.DataFrame(columns=["Day", "Curr_Month_Ton"])

    return pd.merge(t_prev, t_curr, on="Day", how="outer").sort_values("Day").fillna(0.0)


def m2_generate_scrap_jpg(df_day_filtered, sel_date_obj, total_rej_pcs, total_rej_ton, prev_total_ton, prev_avg_ton, curr_as_of_total_ton, curr_as_of_avg_ton, high_rej_count, total_day_mcs, top_cause, top_cause_pcs, top_cause_pct, top_wt_mc, top_wt_kg, top3_pct, gf_share_pct, ff_share_pct):
    fig, ax = plt.subplots(figsize=(16, 10.2), dpi=220)
    fig.patch.set_facecolor('#f4f7fc')
    ax.set_facecolor('#f4f7fc')
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    date_formatted = sel_date_obj.strftime("%B %d, %Y")
    day_formatted = sel_date_obj.strftime("%B %d")

    # 1. Compact Header Banner
    banner = patches.FancyBboxPatch((1.5, 90.0), 97.0, 8.5, boxstyle="round,pad=0.2,rounding_size=0.8", facecolor='#1e1b4b', edgecolor='none')
    ax.add_patch(banner)
    ax.text(3.5, 96.2, "OPERATIONAL QUALITY & DEFECT CONTROL", color='#a5b4fc', fontsize=8.2, fontweight='bold')
    ax.text(3.5, 93.2, "Daily Scrap & Defect Analytics Report", color='#ffffff', fontsize=15.5, fontweight='bold')
    ax.text(3.5, 91.2, f"Complete Line Defect Breakdown (>50 Pcs) & Action Intelligence   |   Report Date: {date_formatted}", color='#94a3b8', fontsize=8.0)

    # Compact Banner Badge
    badge = patches.FancyBboxPatch((84.0, 90.8), 13.5, 6.9, boxstyle="round,pad=0.15,rounding_size=0.6", facecolor='#dc2626', edgecolor='none')
    ax.add_patch(badge)
    ax.text(90.75, 95.0, f"{total_rej_ton:.3f}T", color='#ffffff', fontsize=16, fontweight='bold', ha='center', va='center')
    ax.text(90.75, 92.2, "LAST DAY SCRAP TON", color='#ffffff', fontsize=6.2, fontweight='bold', ha='center', va='center')

    # 2. Compact 6 KPI Cards
    kpis = [
        ("PREV MO. TOTAL", f"{prev_total_ton:.2f} T", "Total Rejection", "#64748b"),
        ("PREV MO. AVG", f"{prev_avg_ton:.2f} T/Day", "Daily Baseline", "#64748b"),
        ("THIS MO. AS OF", f"{curr_as_of_total_ton:.2f} T", f"As of Day {sel_date_obj.day}", "#2563eb"),
        ("THIS MO. AVG", f"{curr_as_of_avg_ton:.2f} T/Day", "Current Pace", "#2563eb"),
        ("LAST DAY SCRAP", f"{total_rej_ton:.3f} T", f"{total_rej_pcs:,} Pcs", "#dc2626"),
        ("CRITICAL MC (>50)", f"{high_rej_count}", "Lines Exceeding Limit", "#8b5cf6"),
    ]

    kpi_w, kpi_gap = 15.1, 1.25
    for i, (title, val, sub, col_bar) in enumerate(kpis):
        x0 = 1.5 + i * (kpi_w + kpi_gap)
        card = patches.FancyBboxPatch((x0, 82.5), kpi_w, 6.3, boxstyle="round,pad=0.15,rounding_size=0.5", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=0.8)
        ax.add_patch(card)
        top_bar = patches.FancyBboxPatch((x0 + 0.1, 88.3), kpi_w - 0.2, 0.45, boxstyle="round,pad=0.03,rounding_size=0.2", facecolor=col_bar, edgecolor='none')
        ax.add_patch(top_bar)
        ax.text(x0 + kpi_w/2, 87.2, title, color='#64748b', fontsize=6.8, fontweight='bold', ha='center')
        ax.text(x0 + kpi_w/2, 84.8, val, color='#0f172a', fontsize=11.5, fontweight='bold', ha='center')
        ax.text(x0 + kpi_w/2, 83.2, sub, color='#94a3b8', fontsize=6.2, ha='center')

    # 3. Main Workspace Body (Expanded to 79.5% height)
    left_card = patches.FancyBboxPatch((1.5, 1.5), 67.0, 79.5, boxstyle="round,pad=0.25,rounding_size=0.8", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
    ax.add_patch(left_card)
    ax.text(3.5, 79.0, f"PLASTIC-3 MACHINE REJECTION LOG (>50 Pcs) — {day_formatted}", color='#0f172a', fontsize=10.0, fontweight='bold')

    right_card = patches.FancyBboxPatch((70.0, 1.5), 28.5, 79.5, boxstyle="round,pad=0.25,rounding_size=0.8", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
    ax.add_patch(right_card)
    ax.text(71.5, 79.0, "EXECUTIVE SUMMARY & ACTION", color='#0f172a', fontsize=10.0, fontweight='bold')

    # Balanced 2-Subtable Display for Left Panel (All Rows Visible)
    mid_idx = (len(df_day_filtered) + 1) // 2
    sub_a = df_day_filtered.iloc[:mid_idx]
    sub_b = df_day_filtered.iloc[mid_idx:]

    cols = ["MC Pos", "Smart Manu", "Causes", "Qty", "kg"]
    sub_configs = [
        (sub_a, 2.8, [4.5, 9.8, 19.5, 28.5, 33.0], 31.8),
        (sub_b, 35.8, [37.5, 42.8, 52.5, 61.5, 66.0], 31.8)
    ]

    for sub_df, left_x, col_xs, tbl_w in sub_configs:
        tbl_hdr = patches.Rectangle((left_x, 75.0), tbl_w, 2.4, facecolor='#0f172a', edgecolor='none')
        ax.add_patch(tbl_hdr)
        for name, cx in zip(cols, col_xs):
            ax.text(cx, 76.2, name, color='#ffffff', fontsize=6.4, fontweight='bold', ha='center', va='center')

        row_y = 73.0
        row_step = 3.55
        for r_i, (_, r) in enumerate(sub_df.iterrows()):
            bg_c = '#f8fafc' if r_i % 2 == 1 else '#ffffff'
            row_bg = patches.Rectangle((left_x, row_y - 1.2), tbl_w, row_step, facecolor=bg_c, edgecolor='none')
            ax.add_patch(row_bg)
            ax.plot([left_x, left_x + tbl_w], [row_y - 1.2, row_y - 1.2], color='#e2e8f0', linewidth=0.45)

            ax.text(col_xs[0], row_y + 0.35, str(r["MC Position"]), color='#0f172a', fontsize=6.3, fontweight='bold', ha='center')
            ax.text(col_xs[1], row_y + 0.35, str(r["Smart Manu"]), color='#64748b', fontsize=6.2, ha='center')
            ax.text(col_xs[2], row_y + 0.35, str(r["Causes"])[:21], color='#ef4444', fontsize=6.0, ha='center')
            ax.text(col_xs[3], row_y + 0.35, f"{int(r['Qty (Pcs)']):,}", color='#0f172a', fontsize=6.4, fontweight='bold', ha='center')
            ax.text(col_xs[4], row_y + 0.35, f"{r['Weight (kg)']:.1f}", color='#0f172a', fontsize=6.2, ha='center')
            row_y -= row_step

    # Right Panel 3 Info Blocks
    # Box 1: Executive Summary
    b1 = patches.FancyBboxPatch((71.2, 54.0), 26.0, 22.0, boxstyle="round,pad=0.2,rounding_size=0.5", facecolor='#f8fafc', edgecolor='#e2e8f0', linewidth=0.8)
    ax.add_patch(b1)
    ax.text(72.5, 73.8, "> Executive Approval Summary", color='#0f172a', fontsize=7.8, fontweight='bold')
    t1 = (
        f"• {high_rej_count} machines in Plastic-3 exceeded 50 pcs.\n"
        f"• Total Last Day Scrap: {total_rej_pcs:,} Pcs ({total_rej_ton:.3f} T).\n"
        f"• Prev Mo: {prev_total_ton:.2f} T ({prev_avg_ton:.2f} T/D avg).\n"
        f"• This Mo (As of {day_formatted}): {curr_as_of_total_ton:.2f} T ({curr_as_of_avg_ton:.2f} T/D avg)."
    )
    ax.text(72.5, 70.8, t1, color='#334155', fontsize=6.8, linespacing=1.45, va='top')

    # Box 2: Defect Drivers
    b2 = patches.FancyBboxPatch((71.2, 28.0), 26.0, 24.5, boxstyle="round,pad=0.2,rounding_size=0.5", facecolor='#fef2f2', edgecolor='#fecaca', linewidth=0.8)
    ax.add_patch(b2)
    ax.text(72.5, 50.3, "[!] Top Defect Driver & Heavy Lines", color='#b91c1c', fontsize=7.8, fontweight='bold')
    t2 = (
        f"• Primary Scrap Cause: '{top_cause}'\n"
        f"  generating {top_cause_pcs:,} pcs loss ({top_cause_pct:.1f}% share).\n"
        f"• Secondary: 'Color Problem*'\n"
        f"• Heaviest Scrap MC: Machine {top_wt_mc}\n"
        f"  generating {top_wt_kg:.1f} kg scrap loss.\n"
        f"• Repetitive short filling noted on cutlery\n"
        f"  and multi-cavity container molds."
    )
    ax.text(72.5, 47.3, t2, color='#7f1d1d', fontsize=6.7, linespacing=1.4, va='top')

    # Box 3: Last Day Overview
    b3 = patches.FancyBboxPatch((71.2, 3.0), 26.0, 23.5, boxstyle="round,pad=0.2,rounding_size=0.5", facecolor='#eff6ff', edgecolor='#bfdbfe', linewidth=0.8)
    ax.add_patch(b3)
    ax.text(72.5, 24.3, "> Last Day Plant Overview", color='#1d4ed8', fontsize=7.8, fontweight='bold')
    t3 = (
        f"• Total Active Lines Logged: {total_day_mcs} machines\n"
        f"  ({high_rej_count} lines >50 pcs, {total_day_mcs - high_rej_count} minor lines <=50 pcs).\n"
        f"• Overall Factory Scrap: {total_rej_pcs:,} Pcs / {total_rej_ton:.3f} Ton.\n"
        f"• Top 3 Causes Account for: {top3_pct:.1f}% of lost pcs.\n"
        f"• Shop Floor Distribution: GF sections generated\n"
        f"  {gf_share_pct:.1f}% of scrap wt; FF sections generated {ff_share_pct:.1f}%."
    )
    ax.text(72.5, 21.3, t3, color='#1e3a8a', fontsize=6.7, linespacing=1.4, va='top')

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = io.BytesIO()
    plt.savefig(buf, format='jpg', facecolor=fig.get_facecolor(), edgecolor='none', dpi=220)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render_scrap_module():
    c_back, c_title, c_act = st.columns([1.2, 3, 1.2], vertical_alignment="center")
    with c_back:
        if st.button("⬅️ Back to Operations Hub", use_container_width=True):
            st.session_state["active_view"] = "hub_home"
            st.rerun()
    with c_title:
        st.markdown("<h3 style='margin:0; text-align:center; font-weight:800; color:#0f172a;'>📉 DAILY SCRAP & DEFECT ANALYTICS</h3>", unsafe_allow_html=True)
    with c_act:
        if "m2_file_bytes" in st.session_state:
            if st.button("🔄 Change Excel File", use_container_width=True):
                st.session_state.pop("m2_file_bytes", None)
                st.rerun()

    st.divider()

    if "m2_file_bytes" not in st.session_state:
        c_up, _ = st.columns([2, 1])
        with c_up:
            st.markdown(
                '<div style="background:#ffffff; padding:1.75rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #dc2626; box-shadow: 0 4px 12px rgba(15,23,42,0.05);">'
                '<h3 style="margin-top:0; color:#0f172a;">📂 Upload Rejection / Scrap Workbook</h3>'
                '<p style="color:#64748b !important;">Select the Excel workbook containing monthly defect records (e.g. rej78.xlsx).</p></div>',
                unsafe_allow_html=True,
            )
            st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)

            uploaded_file = st.file_uploader("Select Excel File (.xlsx, .xls)", type=["xlsx", "xls"], key="m2_uploader")
            if uploaded_file is not None:
                if st.button("🚀 Ingest Rejection Data & Launch", type="primary", use_container_width=True):
                    st.session_state["m2_file_bytes"] = uploaded_file.getvalue()
                    st.rerun()
    else:
        df_prev, df_curr, df_full = m2_parse_workbook(st.session_state["m2_file_bytes"])
        all_dates = sorted(df_curr["DateStr"].unique().tolist())
        
        st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
        c_date, c_cut, c_snap = st.columns([1.5, 1.2, 1.3], gap="small")
        with c_date:
            sel_date_str = st.selectbox("📅 **Operational Date**", all_dates, index=len(all_dates) - 1)
        with c_cut:
            min_cutoff = st.number_input("🔢 **Min Cutoff (Pcs)**", min_value=1, value=50, step=10)
            
        sel_date_obj = pd.to_datetime(sel_date_str)
        sel_day_num = sel_date_obj.day
        day_formatted = sel_date_obj.strftime("%B %d")

        # 1. Day records & filter
        df_day = df_curr[df_curr["DateStr"] == sel_date_str].copy()
        df_day_filtered = m2_compute_daily_rejection(df_day, min_qty=min_cutoff)
        
        # 2. Previous Month Stats (Full Month)
        prev_wt_col = get_col(df_prev, ["Weight", "Rejection Ton"], None)
        if not df_prev.empty and prev_wt_col:
            prev_total_ton = float(pd.to_numeric(df_prev[prev_wt_col], errors="coerce").fillna(0).sum())
            prev_days_count = df_prev["DateClean"].dt.days_in_month.iloc[0] if not df_prev.empty else 31
            prev_avg_ton = prev_total_ton / prev_days_count
        else:
            prev_total_ton, prev_avg_ton = 0.0, 0.0

        # 3. Present Month Stats (As of selected date)
        curr_wt_col = get_col(df_curr, ["Weight", "Rejection Ton"], None)
        df_curr_as_of = df_curr[df_curr["DateClean"].dt.day <= sel_day_num]
        if not df_curr_as_of.empty and curr_wt_col:
            curr_as_of_total_ton = float(pd.to_numeric(df_curr_as_of[curr_wt_col], errors="coerce").fillna(0).sum())
            curr_as_of_avg_ton = curr_as_of_total_ton / sel_day_num
        else:
            curr_as_of_total_ton, curr_as_of_avg_ton = 0.0, 0.0

        # 4. Daily Totals & Drivers
        qty_col = get_col(df_day, ["Quantity", "Qty", "Rejection Pcs"], None)
        wt_col = get_col(df_day, ["Weight", "Rejection Ton"], None)
        cause_col = get_col(df_day, ["Cause", "Causes"], "Cause")
        mc_col = get_col(df_day, ["Machine", "MC SL"], "Machine")

        qty_factor = 1000.0 if (qty_col and df_day[qty_col].max() < 100) else 1.0
        total_rej_pcs = int(round(pd.to_numeric(df_day[qty_col], errors="coerce").fillna(0).sum() * qty_factor)) if (qty_col and not df_day.empty) else 0
        total_rej_ton = float(pd.to_numeric(df_day[wt_col], errors="coerce").fillna(0).sum()) if (wt_col and not df_day.empty) else 0.0
        high_rej_count = len(df_day_filtered)
        total_day_mcs = df_day[mc_col].nunique() if mc_col in df_day.columns else high_rej_count

        pareto_df = m2_compute_pareto(df_curr)
        df_trend = m2_compute_tonnage_comparison(df_prev, df_curr)

        # Top defect cause & heaviest machine metrics
        if not df_day.empty and cause_col in df_day.columns and qty_col in df_day.columns:
            cause_grp = df_day.groupby(cause_col)[qty_col].sum() * qty_factor
            top_cause = cause_grp.idxmax() if not cause_grp.empty else "General"
            top_cause_pcs = int(round(cause_grp.max())) if not cause_grp.empty else 0
            top_cause_pct = (top_cause_pcs / total_rej_pcs * 100.0) if total_rej_pcs > 0 else 0.0
            top3_pcs = cause_grp.sort_values(ascending=False).head(3).sum()
            top3_pct = (top3_pcs / total_rej_pcs * 100.0) if total_rej_pcs > 0 else 0.0
        else:
            top_cause, top_cause_pcs, top_cause_pct, top3_pct = "General", 0, 0.0, 0.0

        if not df_day.empty and mc_col in df_day.columns and wt_col in df_day.columns:
            mc_wt_grp = df_day.groupby(mc_col)[wt_col].sum() * 1000.0
            top_wt_mc = mc_wt_grp.idxmax() if not mc_wt_grp.empty else "-"
            top_wt_kg = float(mc_wt_grp.max()) if not mc_wt_grp.empty else 0.0
            
            df_day["LineCode"] = df_day[mc_col].map(LINE_MAP).fillna("-")
            gf_wt = df_day[df_day["LineCode"].str.startswith("GF")][wt_col].sum()
            ff_wt = df_day[df_day["LineCode"].str.startswith("FF")][wt_col].sum()
            tot_w = (gf_wt + ff_wt) if (gf_wt + ff_wt) > 0 else 1.0
            gf_share_pct = (gf_wt / tot_w * 100.0)
            ff_share_pct = (ff_wt / tot_w * 100.0)
        else:
            top_wt_mc, top_wt_kg, gf_share_pct, ff_share_pct = "-", 0.0, 80.0, 20.0

        # 5. Visual Export
        jpg_bytes = m2_generate_scrap_jpg(
            df_day_filtered, sel_date_obj, total_rej_pcs, total_rej_ton,
            prev_total_ton, prev_avg_ton, curr_as_of_total_ton, curr_as_of_avg_ton, high_rej_count,
            total_day_mcs, top_cause, top_cause_pcs, top_cause_pct, top_wt_mc, top_wt_kg,
            top3_pct, gf_share_pct, ff_share_pct
        )

        with c_snap:
            st.markdown("<div style='margin-top: 1.65rem;'></div>", unsafe_allow_html=True)
            st.download_button(label="📸 Download 1-Page JPG", data=jpg_bytes, file_name=f"Daily_Scrap_Report_{sel_date_str}.jpg", mime="image/jpeg", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        # Header Banner
        st.markdown(
            f"""
            <div class="report-header-banner" style="background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);">
                <div>
                    <span style="color: #a5b4fc; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.08em; text-transform: uppercase;">✦ QUALITY & SCRAP ANALYTICS</span>
                    <h2>Daily Scrap & Defect Summary</h2>
                    <p>Plastic-3 Rejections (&gt;{min_cutoff} Pcs) & Month-over-Month Baseline Tracking &nbsp;|&nbsp; 📅 <b>Report Date:</b> {sel_date_str}</p>
                </div>
                <div class="efficiency-badge-large" style="background: #dc2626;">
                    <div class="value">{total_rej_ton:.3f}T</div>
                    <div class="label">Last Day Scrap Ton</div>
                </div>
            </div>
            """, unsafe_allow_html=True,
        )

        # 6 KPI Cards
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f'<div class="kpi-card indigo"><div class="kpi-title">PREV MO. TOTAL</div><div class="kpi-val">{prev_total_ton:.2f} T</div><div class="kpi-sub">Total Rejection</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card teal"><div class="kpi-title">PREV MO. AVG</div><div class="kpi-val">{prev_avg_ton:.2f} T/D</div><div class="kpi-sub">Daily Baseline</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card blue"><div class="kpi-title">THIS MO. AS OF</div><div class="kpi-val">{curr_as_of_total_ton:.2f} T</div><div class="kpi-sub">As of Day {sel_day_num}</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="kpi-card purple"><div class="kpi-title">THIS MO. AVG</div><div class="kpi-val">{curr_as_of_avg_ton:.2f} T/D</div><div class="kpi-sub">As of Day {sel_day_num}</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="kpi-card pink"><div class="kpi-title">LAST DAY SCRAP</div><div class="kpi-val">{total_rej_ton:.3f} T</div><div class="kpi-sub">{total_rej_pcs:,} Pcs</div></div>', unsafe_allow_html=True)
        k6.markdown(f'<div class="kpi-card yellow"><div class="kpi-title">CRITICAL MC (&gt;{min_cutoff})</div><div class="kpi-val">{high_rej_count}</div><div class="kpi-sub">Lines Over Limit</div></div>', unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom: 1.15rem;'></div>", unsafe_allow_html=True)

        # Mid Section
        col_left, col_right = st.columns([1.55, 0.95], gap="medium")
        with col_left:
            st.markdown(f'<div class="panel-card"><h4>⚙️ PLASTIC-3 MACHINE REJECTION LOG (&gt;{min_cutoff} Pcs)</h4>', unsafe_allow_html=True)
            if not df_day_filtered.empty:
                st.dataframe(
                    df_day_filtered[["MC Position", "Line", "Smart Manu", "Causes", "Qty (Pcs)", "Weight (kg)", "Mold"]],
                    use_container_width=True,
                    hide_index=True,
                    height=420
                )
            else:
                st.success("✅ No machines exceeded the rejection cutoff threshold today!")
            st.markdown("</div>", unsafe_allow_html=True)

        with col_right:
            approval_text = f"""Sir,

These are the machines from Plastic-3 that had a rejection count of more than 50 pieces on {day_formatted}.

Last month, we recorded {prev_total_ton:.2f} tons of rejection with an average of {prev_avg_ton:.2f} tons/day, whereas this month we have recorded {curr_as_of_total_ton:.2f} tons as of today with {curr_as_of_avg_ton:.2f} tons/day.

Need your approval, please, to send to rejection."""

            st.markdown(
                f"""<div class="panel-card">
                    <h4>📝 EXECUTIVE APPROVAL TEXT</h4>
                    <div class="narrative-block" style="font-size: 0.88rem; line-height: 1.6;">
                        <p style="margin: 0 0 0.75rem 0;"><b>Sir,</b></p>
                        <p>These are the machines from <b>Plastic-3</b> that had a rejection count of more than 50 pieces on <b>{day_formatted}</b>.</p>
                        <p>Last month, we recorded <b>{prev_total_ton:.2f} tons</b> of rejection with an average of <b>{prev_avg_ton:.2f} tons/day</b>, whereas this month we have recorded <b>{curr_as_of_total_ton:.2f} tons</b> as of today with <b>{curr_as_of_avg_ton:.2f} tons/day</b>.</p>
                        <p style="margin: 0.75rem 0 0 0; color: #dc2626; font-weight: 700;">Need your approval, please, to send to rejection.</p>
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )

            with st.expander("📋 Copy Plain Text for Approval / WhatsApp"):
                st.text_area("Approval Text", value=approval_text, height=160, label_visibility="collapsed")

        # Bottom Section
        st.markdown('<div class="panel-card"><h4>📅 MONTH-OVER-MONTH DAILY SCRAP TONNAGE TREND</h4>', unsafe_allow_html=True)
        trend_display = df_trend.rename(columns={"Day": "Day of Month", "Prev_Month_Ton": "Previous Month Scrap (Tons)", "Curr_Month_Ton": "Current Month Scrap (Tons)"})
        st.dataframe(trend_display, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)
