import io
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# Standard machine position mapping lookup table
POSITION_MAP = {
    "IMM-280R-25": "A1-280TC", "IMM-380-5": "A2-380", "IMM-380-81": "A3-380 (PC)",
    "IMM-380-82": "A4-380 (PC)", "IMM-380-88": "A5-380", "IMM-380-7": "A6-380",
    "IMM-380-60": "B1-380", "IMM-380-6": "B2-380", "IMM-380-58": "B3-380",
    "IMM-380-59": "B4-380", "IMM-380-76": "B5-380", "IMM-380-75": "B6-380",
    "IMM-380-57": "C1-380", "IMM-380-36": "C2-380", "IMM-380-45": "C3-380",
    "IMM-380-43": "C4-380", "IMM-380-44": "C5-380-44", "IMM-380-74": "C6-380",
    "IMM-330-4": "D1-330", "IMM-330-3": "D2-330", "IMM-330-5": "D3-330",
    "IMM-330-6": "D4-330", "IMM-428-2": "D5-HP-428-2", "IMM-428-4": "D6-HP-428-4",
    "IMM-428-3": "D7-HP-428-3", "IMM-428-1": "D8-HP-428-1", "IMM-160-5": "E1-160",
    "IMM-160-7": "E2-160", "IMM-160-8": "E3-160", "IMM-160-9": "E4-160",
    "IMM-160-12": "E5-160", "IMM-160-14": "E6-160", "IMM-160-10": "F1-160",
    "IMM-160-11": "F2-160", "IMM-160-13": "F3-160", "IMM-160-15": "F4-160",
    "IMM-160-16": "F5-160", "IMM-160-6": "F6-160", "IMM-120-20": "G1-120",
    "IMM-120-4": "G2-120", "IMM-120-19": "G3-120", "IMM-120-14": "G4-120",
    "IMM-90-2": "G5-90", "IMM-90-1": "G6-90", "IMM-250-7": "H1-250",
    "IMM-250-8": "H2-250", "IMM-250-6": "H3-250", "IMM-250-9": "H4-250",
    "IMM-250-10": "H5-250", "IMM-270-1": "H6-270", "IMM-470-3": "I1-470",
    "IMM-470-4": "I2-470", "IMM-470-5": "I3-470", "IMM-470-6": "I4-470",
    "IMM-530-4": "I5-530", "IMM-800-1": "I6-800",
}


@st.cache_data
def m2_parse_workbook(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)
    
    # 1. Parse Raw Records (This Month)
    sheet_raw = "This Month" if "This Month" in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(xls, sheet_name=sheet_raw)
    date_col = "Added Date" if "Added Date" in df_raw.columns else ("Date" if "Date" in df_raw.columns else df_raw.columns[0])
    df_raw["DateClean"] = pd.to_datetime(df_raw[date_col], errors="coerce")
    df_raw = df_raw.dropna(subset=["DateClean"]).sort_values("DateClean")
    df_raw["DateStr"] = df_raw["DateClean"].dt.strftime("%Y-%m-%d")
    df_raw["YearMonth"] = df_raw["DateClean"].dt.to_period("M")
    
    # Identify unique months
    unique_months = df_raw["YearMonth"].unique()
    if len(unique_months) >= 2:
        df_prev = df_raw[df_raw["YearMonth"] == unique_months[-2]].copy()
        df_curr = df_raw[df_raw["YearMonth"] == unique_months[-1]].copy()
    else:
        df_prev = pd.DataFrame()
        df_curr = df_raw.copy()

    # 2. Check Comparison Sheet for Historical Baseline if present
    df_comp = None
    if "Comparison" in xls.sheet_names:
        df_comp = pd.read_excel(xls, sheet_name="Comparison")
        
    return df_prev, df_curr, df_raw, df_comp


def m2_compute_daily_rejection(df_day, min_qty=50):
    if df_day.empty:
        return pd.DataFrame()
    
    mc_col = "Machine" if "Machine" in df_day.columns else df_day.columns[2]
    item_col = "Item" if "Item" in df_day.columns else df_day.columns[3]
    cause_col = "Cause" if "Cause" in df_day.columns else "Causes"
    
    records = []
    for mc, grp in df_day.groupby(mc_col):
        # In Excel raw data: 0.10 quantity = 100 pcs (multiply by 1000)
        qty_factor = 1000.0 if grp["Quantity"].max() < 100 else 1.0
        total_pcs = grp["Quantity"].sum() * qty_factor
        total_ton = grp["Weight"].sum()
        
        # Combine unique causes and items
        unique_causes = [str(c).strip() for c in grp[cause_col].dropna().unique() if str(c).strip()]
        causes_str = ", ".join(unique_causes) if unique_causes else "No Rejection"
        mold_name = str(grp[item_col].iloc[0]) if not grp[item_col].dropna().empty else "-"
        pos = POSITION_MAP.get(mc, grp.get("Position", pd.Series(["-"])).iloc[0] if "Position" in grp.columns else "-")
        
        if total_pcs >= min_qty:
            records.append({
                "Position": pos,
                "Machine": mc,
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
    cause_col = "Cause" if "Cause" in df_curr.columns else "Causes"
    if df_curr.empty or cause_col not in df_curr.columns:
        return pd.DataFrame()
    
    qty_factor = 1000.0 if df_curr["Quantity"].max() < 100 else 1.0
    pareto = df_curr.groupby(cause_col).agg(
        Rejection_Pcs=("Quantity", lambda x: int(round(x.sum() * qty_factor))),
        Rejection_Ton=("Weight", "sum")
    ).reset_index()
    
    total_pcs = pareto["Rejection_Pcs"].sum()
    pareto["% Share"] = (pareto["Rejection_Pcs"] / total_pcs * 100).round(2) if total_pcs > 0 else 0.0
    pareto = pareto.sort_values("Rejection_Pcs", ascending=False).reset_index(drop=True)
    return pareto


def m2_compute_tonnage_comparison(df_prev, df_curr, df_comp=None):
    if df_comp is not None and "Date" in df_comp.columns and "Rejection Ton" in df_comp.columns:
        # Pull directly from Comparison sheet
        c_prev = df_comp[["Date", "Rejection Ton"]].dropna().copy()
        c_prev["Day"] = pd.to_datetime(c_prev["Date"], errors="coerce").dt.day
        c_prev = c_prev.rename(columns={"Rejection Ton": "Prev_Month_Ton"})[["Day", "Prev_Month_Ton"]]
        
        c_curr_cols = [c for c in df_comp.columns if "Rejection Ton." in c or c == "Rejection Ton.1"]
        if c_curr_cols:
            c_curr = df_comp[["Date.1", c_curr_cols[0]]].dropna().copy()
            c_curr["Day"] = pd.to_datetime(c_curr["Date.1"], errors="coerce").dt.day
            c_curr = c_curr.rename(columns={c_curr_cols[0]: "Curr_Month_Ton"})[["Day", "Curr_Month_Ton"]]
            df_trend = pd.merge(c_prev, c_curr, on="Day", how="outer").sort_values("Day").fillna(0.0)
            return df_trend

    # Fallback to computing from raw logs
    t_prev = df_prev.groupby(df_prev["DateClean"].dt.day)["Weight"].sum().reset_index() if not df_prev.empty else pd.DataFrame(columns=["Day", "Weight"])
    t_curr = df_curr.groupby(df_curr["DateClean"].dt.day)["Weight"].sum().reset_index() if not df_curr.empty else pd.DataFrame(columns=["Day", "Weight"])
    t_prev.columns = ["Day", "Prev_Month_Ton"]
    t_curr.columns = ["Day", "Curr_Month_Ton"]
    return pd.merge(t_prev, t_curr, on="Day", how="outer").sort_values("Day").fillna(0.0)


def m2_generate_scrap_jpg(df_day_filtered, sel_date, total_rej_pcs, total_rej_ton, top_cause, mtd_ton, prev_avg_ton, high_rej_count):
    fig, ax = plt.subplots(figsize=(16, 9.8), dpi=220)
    fig.patch.set_facecolor('#f4f7fc')
    ax.set_facecolor('#f4f7fc')
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

    # Banner
    banner = patches.FancyBboxPatch((2, 85), 96, 12.5, boxstyle="round,pad=0.3,rounding_size=1.2", facecolor='#1e1b4b', edgecolor='none')
    ax.add_patch(banner)
    ax.text(4, 94.5, "OPERATIONAL QUALITY & DEFECT CONTROL", color='#a5b4fc', fontsize=10, fontweight='bold')
    ax.text(4, 90.5, "Daily Scrap & Defect Analytics Report", color='#ffffff', fontsize=19, fontweight='bold')
    ax.text(4, 87.2, f"Line-Level Defect Isolation (>50 Pcs Threshold) & Pareto Distribution   |   Report Date: {sel_date}", color='#94a3b8', fontsize=9.2)

    # Scrap Badge
    badge = patches.FancyBboxPatch((82.0, 86.2), 14.0, 10, boxstyle="round,pad=0.2,rounding_size=1", facecolor='#dc2626', edgecolor='none')
    ax.add_patch(badge)
    ax.text(89.0, 92.2, f"{total_rej_ton:.3f}T", color='#ffffff', fontsize=21, fontweight='bold', ha='center', va='center')
    ax.text(89.0, 88.2, "DAILY SCRAP TON", color='#ffffff', fontsize=7, fontweight='bold', ha='center', va='center')

    # 6 KPI Cards
    kpis = [
        ("TOTAL REJ PCS", f"{total_rej_pcs:,}", "Pieces Lost", "#dc2626"),
        ("DAILY SCRAP TON", f"{total_rej_ton:.3f} T", f"{total_rej_ton*1000:.1f} kg", "#f59e0b"),
        ("CRITICAL MC (>50)", f"{high_rej_count}", "Lines Over Limit", "#8b5cf6"),
        ("TOP DEFECT CAUSE", str(top_cause)[:15], "Primary Scrap Driver", "#2563eb"),
        ("MTD TOTAL SCRAP", f"{mtd_ton:.2f} T", "Month-To-Date", "#06b6d4"),
        ("PREV MO. AVG", f"{prev_avg_ton:.3f} T", "Daily Baseline", "#64748b"),
    ]

    kpi_w, kpi_gap = 15.0, 1.2
    for i, (title, val, sub, col_bar) in enumerate(kpis):
        x0 = 2 + i * (kpi_w + kpi_gap)
        card = patches.FancyBboxPatch((x0, 73.5), kpi_w, 9.5, boxstyle="round,pad=0.2,rounding_size=0.8", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
        ax.add_patch(card)
        top_bar = patches.FancyBboxPatch((x0 + 0.1, 82.2), kpi_w - 0.2, 0.6, boxstyle="round,pad=0.05,rounding_size=0.3", facecolor=col_bar, edgecolor='none')
        ax.add_patch(top_bar)
        ax.text(x0 + kpi_w/2, 80.8, title, color='#64748b', fontsize=7.6, fontweight='bold', ha='center')
        ax.text(x0 + kpi_w/2, 77.2, val, color='#0f172a', fontsize=13.5, fontweight='bold', ha='center')
        ax.text(x0 + kpi_w/2, 74.8, sub, color='#94a3b8', fontsize=6.8, ha='center')

    # Left & Right Panels
    left_card = patches.FancyBboxPatch((2, 2.5), 62.0, 69.0, boxstyle="round,pad=0.3,rounding_size=1", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
    ax.add_patch(left_card)
    ax.text(4, 68.5, "MACHINE WISE CRITICAL REJECTION LOG (>50 Pcs)", color='#0f172a', fontsize=11, fontweight='bold')

    right_card = patches.FancyBboxPatch((65.5, 2.5), 32.5, 69.0, boxstyle="round,pad=0.3,rounding_size=1", facecolor='#ffffff', edgecolor='#e2e8f0', linewidth=1)
    ax.add_patch(right_card)
    ax.text(67.5, 68.5, "DEFECT ANALYSIS & CORRECTIVE FOCUS", color='#0f172a', fontsize=11, fontweight='bold')

    col_names = ["Position", "Machine", "Causes", "Rej Pcs", "Weight (kg)"]
    col_xs = [6.0, 16.0, 34.0, 49.0, 58.0]
    
    tbl_hdr = patches.Rectangle((3.5, 63.8), 59.0, 3.2, facecolor='#0f172a', edgecolor='none')
    ax.add_patch(tbl_hdr)
    for name, cx in zip(col_names, col_xs):
        ax.text(cx, 65.4, name, color='#ffffff', fontsize=7.5, fontweight='bold', ha='center', va='center')

    row_y = 61.2
    row_step = 3.9
    for r_i, (_, r) in enumerate(df_day_filtered.head(14).iterrows()):
        bg_c = '#f8fafc' if r_i % 2 == 1 else '#ffffff'
        row_bg = patches.Rectangle((3.5, row_y - 1.4), 59.0, row_step, facecolor=bg_c, edgecolor='none')
        ax.add_patch(row_bg)
        ax.plot([3.5, 62.5], [row_y - 1.4, row_y - 1.4], color='#e2e8f0', linewidth=0.6)

        ax.text(col_xs[0], row_y + 0.5, str(r["Position"])[:10], color='#0f172a', fontsize=7.2, fontweight='bold', ha='center')
        ax.text(col_xs[1], row_y + 0.5, str(r["Machine"]), color='#0f172a', fontsize=7.2, ha='center')
        ax.text(col_xs[2], row_y + 0.5, str(r["Causes"])[:28], color='#ef4444', fontsize=7.0, ha='center')
        ax.text(col_xs[3], row_y + 0.5, f"{int(r['Qty (Pcs)']):,}", color='#0f172a', fontsize=7.4, fontweight='bold', ha='center')
        ax.text(col_xs[4], row_y + 0.5, f"{r['Weight (kg)']:.1f}", color='#0f172a', fontsize=7.2, ha='center')
        row_y -= row_step

    narr_y = 64.0
    ax.text(67.5, narr_y, "Daily Scrap Performance", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.4
    ax.text(67.5, narr_y, f"Total line scrap recorded was {total_rej_pcs:,} Pcs\n({total_rej_ton:.3f} Tons). A total of {high_rej_count} machines\nexceeded the 50 pcs rejection threshold.", color='#475569', fontsize=7.8, linespacing=1.4, va='top')

    narr_y -= 7.5
    ax.text(67.5, narr_y, "Primary Defect Driver", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.4
    ax.text(67.5, narr_y, f"Highest contributing defect was '{top_cause}'.\nImmediate process tuning and mold maintenance\nare required on primary affected lines.", color='#475569', fontsize=7.8, linespacing=1.4, va='top')

    narr_y -= 7.5
    ax.text(67.5, narr_y, "Month-to-Date Comparison", color='#0f172a', fontsize=9.5, fontweight='bold')
    narr_y -= 2.4
    ax.text(67.5, narr_y, f"Cumulative MTD Scrap: {mtd_ton:.2f} Tons\nBaseline Daily Average: {prev_avg_ton:.3f} Tons/Day.", color='#475569', fontsize=7.8, linespacing=1.4, va='top')

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
                '<p style="color:#64748b !important;">Select the Excel workbook containing monthly defect records (e.g. RejectionMCwise.xlsx).</p></div>',
                unsafe_allow_html=True,
            )
            st.markdown("<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True)

            uploaded_file = st.file_uploader("Select Excel File (.xlsx, .xls)", type=["xlsx", "xls"], key="m2_uploader")
            if uploaded_file is not None:
                if st.button("🚀 Ingest Rejection Data & Launch", type="primary", use_container_width=True):
                    st.session_state["m2_file_bytes"] = uploaded_file.getvalue()
                    st.rerun()
    else:
        df_prev, df_curr, df_full, df_comp = m2_parse_workbook(st.session_state["m2_file_bytes"])
        all_dates = sorted(df_curr["DateStr"].unique().tolist())
        
        st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
        c_date, c_cut, c_snap = st.columns([1.5, 1.2, 1.3], gap="small")
        with c_date:
            sel_date = st.selectbox("📅 **Operational Date**", all_dates, index=len(all_dates) - 1)
        with c_cut:
            min_cutoff = st.number_input("🔢 **Min Cutoff (Pcs)**", min_value=1, value=50, step=10)
            
        df_day = df_curr[df_curr["DateStr"] == sel_date].copy()
        df_day_filtered = m2_compute_daily_rejection(df_day, min_qty=min_cutoff)
        pareto_df = m2_compute_pareto(df_curr)
        df_trend = m2_compute_tonnage_comparison(df_prev, df_curr, df_comp)

        qty_factor = 1000.0 if df_day["Quantity"].max() < 100 else 1.0
        total_rej_pcs = int(round(df_day["Quantity"].sum() * qty_factor)) if not df_day.empty else 0
        total_rej_ton = df_day["Weight"].sum() if not df_day.empty else 0.0
        high_rej_count = len(df_day_filtered)
        top_cause = pareto_df.iloc[0]["Cause"] if not pareto_df.empty else "None"
        mtd_ton = df_curr["Weight"].sum() if not df_curr.empty else 0.0
        prev_avg_ton = (df_prev["Weight"].sum() / df_prev["DateStr"].nunique()) if not df_prev.empty and df_prev["DateStr"].nunique() > 0 else (df_trend["Prev_Month_Ton"].mean() if not df_trend.empty else 0.0)

        jpg_bytes = m2_generate_scrap_jpg(df_day_filtered, sel_date, total_rej_pcs, total_rej_ton, top_cause, mtd_ton, prev_avg_ton, high_rej_count)

        with c_snap:
            st.markdown("<div style='margin-top: 1.65rem;'></div>", unsafe_allow_html=True)
            st.download_button(label="📸 Download 1-Page JPG", data=jpg_bytes, file_name=f"Daily_Scrap_Report_{sel_date}.jpg", mime="image/jpeg", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown(
            f"""
            <div class="report-header-banner" style="background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);">
                <div>
                    <span style="color: #a5b4fc; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.08em; text-transform: uppercase;">✦ QUALITY & SCRAP ANALYTICS</span>
                    <h2>Daily Scrap & Defect Summary</h2>
                    <p>Line-level rejection isolation (&ge;{min_cutoff} Pcs) & Monthly Pareto Analysis &nbsp;|&nbsp; 📅 <b>Report Date:</b> {sel_date}</p>
                </div>
                <div class="efficiency-badge-large" style="background: #dc2626;">
                    <div class="value">{total_rej_ton:.3f}T</div>
                    <div class="label">Daily Scrap Ton</div>
                </div>
            </div>
            """, unsafe_allow_html=True,
        )

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f'<div class="kpi-card blue"><div class="kpi-title">TOTAL REJ PCS</div><div class="kpi-val">{total_rej_pcs:,}</div><div class="kpi-sub">Total Pcs Lost</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card pink"><div class="kpi-title">DAILY SCRAP TON</div><div class="kpi-val">{total_rej_ton:.3f} T</div><div class="kpi-sub">{total_rej_ton*1000:.1f} kg</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card purple"><div class="kpi-title">CRITICAL MC (>{min_cutoff})</div><div class="kpi-val">{high_rej_count}</div><div class="kpi-sub">Lines Over Limit</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="kpi-card yellow"><div class="kpi-title">TOP DEFECT CAUSE</div><div class="kpi-val" style="font-size:1.05rem;">{top_cause[:14]}</div><div class="kpi-sub">Primary Scrap Driver</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="kpi-card teal"><div class="kpi-title">MTD TOTAL SCRAP</div><div class="kpi-val">{mtd_ton:.2f} T</div><div class="kpi-sub">Month-to-Date</div></div>', unsafe_allow_html=True)
        k6.markdown(f'<div class="kpi-card indigo"><div class="kpi-title">PREV MO. AVG</div><div class="kpi-val">{prev_avg_ton:.3f} T</div><div class="kpi-sub">Daily Benchmark</div></div>', unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom: 1.15rem;'></div>", unsafe_allow_html=True)

        col_left, col_right = st.columns([1.5, 1.0], gap="medium")
        with col_left:
            st.markdown(f'<div class="panel-card"><h4>⚙️ MACHINE WISE CRITICAL DEFECT LOG (&ge;{min_cutoff} Pcs)</h4>', unsafe_allow_html=True)
            if not df_day_filtered.empty:
                st.dataframe(df_day_filtered[["Position", "Machine", "Causes", "Qty (Pcs)", "Weight (kg)", "Mold"]], use_container_width=True, hide_index=True, height=420)
            else:
                st.success("✅ No machines exceeded the rejection cutoff threshold today!")
            st.markdown("</div>", unsafe_allow_html=True)

        with col_right:
            st.markdown('<div class="panel-card"><h4>📊 MTD DEFECT PARETO (TOP CAUSES)</h4>', unsafe_allow_html=True)
            st.dataframe(pareto_df.head(6), use_container_width=True, hide_index=True, height=180)
            
            whatsapp_scrap_text = f"""Dear Sir,

🚨 *Daily Scrap & Defect Report ({sel_date})*
• *Total Rejection Output:* {total_rej_pcs:,} Pcs ({total_rej_ton:.3f} Tons)
• *Critical Lines (>{min_cutoff} Pcs):* {high_rej_count} Machines
• *Top Scrap Cause:* {top_cause}
• *MTD Scrap Total:* {mtd_ton:.2f} Tons (Prev Month Avg: {prev_avg_ton:.3f} T/Day)"""

            with st.expander("📋 Copy Plain Text Scrap Brief"):
                st.text_area("WhatsApp Scrap Brief", value=whatsapp_scrap_text, height=140, label_visibility="collapsed")
            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown('<div class="panel-card"><h4>📅 MONTH-OVER-MONTH DAILY SCRAP TONNAGE TREND</h4>', unsafe_allow_html=True)
        trend_display = df_trend.rename(columns={"Day": "Day of Month", "Prev_Month_Ton": "Previous Month Scrap (Tons)", "Curr_Month_Ton": "Current Month Scrap (Tons)"})
        st.dataframe(trend_display, use_container_width=True, hide_index=True)
        st.markdown("</div>", unsafe_allow_html=True)
