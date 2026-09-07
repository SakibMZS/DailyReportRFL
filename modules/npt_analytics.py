import io
import textwrap
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches

from config import (
    TOTAL_PLANT_MCS,
    DAILY_AVAILABLE_HRS,
    POS_MAP,
    LINE_MAP,
    MAINTENANCE_CAUSES,
    EXCEL_SIZES,
)


def get_col(df, candidates, default=None):
    for c in candidates:
        if c in df.columns:
            return c
    return default


def extract_mc_size(pos_val, mc_val):
    text = f"{pos_val} {mc_val}".upper()
    for sz in sorted(EXCEL_SIZES, key=len, reverse=True):
        if sz in text:
            return sz
    return "Other"


# =========================================================
# 1. PARSING ENGINES (DOWNTIME & SERVICE MAINTENANCE)
# =========================================================
@st.cache_data
def m3_parse_downtime_workbook(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)
    sheet_name = "DowntimeReport" if "DowntimeReport" in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values])
        if "Machine" in row_str and ("Duration" in row_str or "Cause" in row_str):
            header_idx = idx
            break

    df_clean = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx) if header_idx is not None else pd.read_excel(xls, sheet_name=sheet_name)
    df_clean.columns = [str(c).strip() for c in df_clean.columns]

    # Exclude open-ended stoppages without end timestamp
    to_col = get_col(df_clean, ["To Time", "ToTime", "End Time", "EndTime"])
    if to_col:
        df_clean = df_clean[df_clean[to_col].notna()].copy()

    # Base operational date on Cause Added Date
    date_col = get_col(df_clean, ["Cause Added Date", "CauseAddedDate", "Date", "Added Date"], df_clean.columns[-1])
    df_clean["DateClean"] = pd.to_datetime(df_clean[date_col], errors="coerce")
    df_clean = df_clean.dropna(subset=["DateClean"]).sort_values("DateClean")
    df_clean["DateStr"] = df_clean["DateClean"].dt.strftime("%Y-%m-%d")
    df_clean["MonthName"] = df_clean["DateClean"].dt.strftime("%B")
    df_clean["DayNum"] = df_clean["DateClean"].dt.day
    df_clean["YearMonth"] = df_clean["DateClean"].dt.to_period("M")

    # Duration in Hours
    sec_col = get_col(df_clean, ["Duration (In Second)", "Duration(In Second)", "Duration(s)", "Seconds"])
    dur_col = get_col(df_clean, ["Duration", "Duration (Hrs)", "Hours"])
    if sec_col in df_clean.columns:
        df_clean["Hours"] = pd.to_numeric(df_clean[sec_col], errors="coerce").fillna(0) / 3600.0
    elif dur_col in df_clean.columns:
        df_clean["Hours"] = pd.to_timedelta(df_clean[dur_col], errors="coerce").dt.total_seconds() / 3600.0
    else:
        df_clean["Hours"] = 0.0

    df_clean = df_clean[df_clean["Hours"] > 0].copy()

    # Machine Positions and Sizes
    mc_col = get_col(df_clean, ["Machine", "MC SL"], "Machine")
    df_clean["Position"] = df_clean[mc_col].astype(str).map(POS_MAP).fillna("-")
    df_clean["Line"] = df_clean[mc_col].astype(str).map(LINE_MAP).fillna("-")
    df_clean["Size"] = df_clean.apply(lambda r: extract_mc_size(r["Position"], r[mc_col]), axis=1)

    # Clean Cause and Maintenance tagging
    cause_col = get_col(df_clean, ["Cause", "Causes", "Reason", "Defect"], "Cause")
    df_clean["CauseClean"] = df_clean[cause_col].astype(str).str.replace("*", "", regex=False).str.strip()
    df_clean["Is_Maintenance"] = df_clean[cause_col].isin(MAINTENANCE_CAUSES)
    df_clean["Is_SMED"] = df_clean[cause_col].astype(str).str.strip() == "Mold Change*"

    return df_clean


@st.cache_data
def m3_parse_service_maintenance(file_bytes):
    if not file_bytes:
        return pd.DataFrame()
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)
    sheet_name = "ServiceMaintenanceHistoryReport" if "ServiceMaintenanceHistoryReport" in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    hdr_idx = None
    for idx, r in df_raw.iterrows():
        r_str = " ".join([str(v) for v in r.values])
        if "TicketId" in r_str and "Maintenance Type" in r_str:
            hdr_idx = idx
            break

    df_clean = pd.read_excel(xls, sheet_name=sheet_name, skiprows=hdr_idx) if hdr_idx is not None else pd.read_excel(xls, sheet_name=sheet_name)
    df_clean.columns = [str(c).strip() for c in df_clean.columns]

    from_col = get_col(df_clean, ["From", "Start Date", "Created Date"], "From")
    df_clean["DateClean"] = pd.to_datetime(df_clean[from_col], errors="coerce")
    df_clean["MonthName"] = df_clean["DateClean"].dt.strftime("%B")
    df_clean["DayNum"] = df_clean["DateClean"].dt.day
    df_clean["YearMonth"] = df_clean["DateClean"].dt.to_period("M")

    mc_col = get_col(df_clean, ["Machine", "MC SL"], "Machine")
    df_clean["Position"] = df_clean[mc_col].astype(str).map(POS_MAP).fillna("-")
    df_clean["Size"] = df_clean.apply(lambda r: extract_mc_size(r["Position"], r[mc_col]), axis=1)

    return df_clean


# =========================================================
# 2. COMPUTATION LOGIC
# =========================================================
def m3_compute_month_comparison(df_downtime, active_month, months_order, cutoff_day, mode="like_for_like"):
    causes = sorted([c for c in df_downtime["Cause"].dropna().unique() if "Server Error" not in str(c)])
    records = []

    for c in causes:
        row = {"Cause": c}
        for m_period in months_order:
            m_df = df_downtime[df_downtime["YearMonth"] == m_period]
            m_name = m_df["MonthName"].iloc[0] if not m_df.empty else str(m_period)
            
            if mode == "like_for_like":
                scope_df = m_df[m_df["DayNum"] <= cutoff_day]
            else:
                scope_df = m_df

            c_hrs = scope_df[scope_df["Cause"] == c]["Hours"].sum()
            row[f"{m_name} (Hrs)"] = round(c_hrs, 2)

        records.append(row)

    comp_df = pd.DataFrame(records)
    if comp_df.empty:
        return comp_df

    # Calculate percentages
    for m_period in months_order:
        m_df = df_downtime[df_downtime["YearMonth"] == m_period]
        m_name = m_df["MonthName"].iloc[0] if not m_df.empty else str(m_period)
        tot_hrs = comp_df[f"{m_name} (Hrs)"].sum()
        comp_df[f"{m_name} (%)"] = ((comp_df[f"{m_name} (Hrs)"] / tot_hrs * 100).round(2) if tot_hrs > 0 else 0.0).apply(lambda x: f"{x:.2f}%")

    curr_m_name = df_downtime[df_downtime["YearMonth"] == active_month]["MonthName"].iloc[0]
    comp_df = comp_df.sort_values(f"{curr_m_name} (Hrs)", ascending=False).reset_index(drop=True)
    return comp_df


def m3_compute_smed_summary(df_scope):
    smed_df = df_scope[df_scope["Is_SMED"]].copy()
    if smed_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    daily_smed = (
        smed_df.groupby(["DateStr", "DateClean"])
        .agg(
            Mold_Change_Qty=("Hours", "count"),
            Total_Time_Hrs=("Hours", "sum"),
            Involved_Machines=("Position", lambda x: ", ".join(sorted(set(str(v) for v in x if v != "-")))),
        )
        .reset_index()
    )
    daily_smed["Avg_SMED_Min"] = (daily_smed["Total_Time_Hrs"] / daily_smed["Mold_Change_Qty"] * 60.0).round(2)
    daily_smed["Total_Time_Hrs"] = daily_smed["Total_Time_Hrs"].round(2)

    size_smed = (
        smed_df.groupby("Size")
        .agg(
            Mold_Change_Qty=("Hours", "count"),
            Total_Time_Hrs=("Hours", "sum"),
        )
        .reset_index()
    )
    size_smed["Avg_SMED_Min"] = (size_smed["Total_Time_Hrs"] / size_smed["Mold_Change_Qty"] * 60.0).round(2)
    size_smed["Total_Time_Hrs"] = size_smed["Total_Time_Hrs"].round(2)
    size_smed = size_smed.sort_values("Mold_Change_Qty", ascending=False).reset_index(drop=True)

    return daily_smed, size_smed


def m3_compute_maintenance_ticket_correlation(df_downtime_month, df_service_month):
    records = []
    for sz in EXCEL_SIZES:
        sz_dt = df_downtime_month[(df_downtime_month["Size"] == sz) & (df_downtime_month["Cause"] == "Machine Problem*")]
        dt_hrs = sz_dt["Hours"].sum()

        if not df_service_month.empty:
            type_col = get_col(df_service_month, ["Type", "Maintenance Type", "Type "], "Type")
            sz_sm = df_service_month[(df_service_month["Size"] == sz) & (df_service_month[type_col].astype(str).str.contains("Machine Maintenance", case=False, na=False))]
            ticket_qty = len(sz_sm)
        else:
            ticket_qty = 0

        records.append({
            "MC Size": sz,
            "SMS Token Qty": int(ticket_qty),
            "Downtime (Hrs)": round(dt_hrs, 2),
        })

    corr_df = pd.DataFrame(records)
    corr_df = corr_df[(corr_df["SMS Token Qty"] > 0) | (corr_df["Downtime (Hrs)"] > 0)].sort_values("Downtime (Hrs)", ascending=False).reset_index(drop=True)
    return corr_df


# =========================================================
# 3. EXECUTIVE 1-PAGE REPORT GENERATOR (MATPLOTLIB)
# =========================================================
def m3_generate_executive_jpg(
    df_day,
    sel_date_obj,
    total_day_hrs,
    maint_hrs,
    oper_hrs,
    loss_pct,
    maint_loss_pct,
    curr_as_of_total_hrs,
    curr_as_of_avail_hrs,
    as_of_loss_pct,
    curr_as_of_maint_hrs,
    as_of_maint_loss_pct,
    as_of_maint_share_pct,
    gf_share_pct,
    ff_share_pct,
    top_cause_name,
    top_cause_hrs,
    top_cause_pct,
    top_maint_pos,
    top_maint_mc,
    top_maint_hrs,
    prev_total_hrs,
    prev_avg_hrs,
    curr_as_of_avg_hrs,
    top_causes_list,
):
    fig, ax = plt.subplots(figsize=(18, 10.5), dpi=220)
    fig.patch.set_facecolor("#f1f5f9")
    ax.set_facecolor("#f1f5f9")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    date_formatted = sel_date_obj.strftime("%B %d, %Y")
    day_formatted = sel_date_obj.strftime("%B %d")
    day_num = sel_date_obj.day

    # Header
    ax.text(1.5, 98.4, "DAILY NON-PRODUCTIVE TIME (NPT) ANALYTICS", color="#0f172a", fontsize=16.0, fontweight="bold", va="top")
    ax.text(1.5, 95.8, f"Plastic-3 Stoppage & Machine Downtime Log ({TOTAL_PLANT_MCS} IMMs Baseline: 1,464 H/Day)  |  Report Date: {date_formatted}", color="#64748b", fontsize=8.8, va="top")
    ax.text(98.5, 97.2, "PLASTIC-3 OPERATIONS", color="#2563eb", fontsize=8.8, fontweight="bold", ha="right", va="top")

    # KPI Row
    kpis = [
        ("PREV MO. TOTAL", f"{prev_total_hrs:.1f} H", f"Daily Avg: {prev_avg_hrs:.1f} H/D", "#64748b"),
        ("THIS MO. AS OF", f"{curr_as_of_total_hrs:.1f} H", f"Wasted: {as_of_loss_pct:.1f}% Capacity", "#2563eb"),
        ("MTD MAINT. LOSS", f"{curr_as_of_maint_hrs:.1f} H", f"{as_of_maint_loss_pct:.1f}% Cap ({as_of_maint_share_pct:.1f}% NPT)", "#7c3aed"),
        ("LAST DAY NPT", f"{total_day_hrs:.1f} H", f"{loss_pct:.1f}% Available Lost", "#dc2626"),
        ("DAY MAINT. LOSS", f"{maint_hrs:.1f} H", f"{maint_loss_pct:.1f}% Available", "#9333ea"),
        ("TOP NPT DRIVER", f"{top_cause_name[:15]}", f"{top_cause_hrs:.1f} H ({top_cause_pct:.1f}%)", "#059669"),
    ]

    kpi_w, kpi_gap = 15.1, 1.25
    for i, (title, val, sub, col_bar) in enumerate(kpis):
        x0 = 1.5 + i * (kpi_w + kpi_gap)
        card = patches.FancyBboxPatch((x0, 87.0), kpi_w, 7.2, boxstyle="round,pad=0.15,rounding_size=0.5", facecolor="#ffffff", edgecolor="#cbd5e1", linewidth=0.8)
        ax.add_patch(card)
        top_bar = patches.FancyBboxPatch((x0 + 0.1, 93.75), kpi_w - 0.2, 0.45, boxstyle="round,pad=0.03,rounding_size=0.2", facecolor=col_bar, edgecolor="none")
        ax.add_patch(top_bar)
        ax.text(x0 + kpi_w / 2, 92.4, title, color="#64748b", fontsize=7.6, fontweight="bold", ha="center")
        ax.text(x0 + kpi_w / 2, 89.6, val, color="#0f172a", fontsize=13.0, fontweight="bold", ha="center")
        ax.text(x0 + kpi_w / 2, 87.8, sub, color="#94a3b8", fontsize=6.8, ha="center")

    # Containers
    left_card = patches.FancyBboxPatch((1.5, 1.5), 74.0, 84.0, boxstyle="round,pad=0.25,rounding_size=0.8", facecolor="#ffffff", edgecolor="#cbd5e1", linewidth=1)
    ax.add_patch(left_card)
    ax.text(3.5, 83.5, f"DAILY NPT MACHINE LOG (WITH POSITION) — {day_formatted}", color="#0f172a", fontsize=11.0, fontweight="bold")
    ax.text(73.5, 83.5, f"Plant Loss: {loss_pct:.1f}% of Available Capacity", color="#64748b", fontsize=8.0, ha="right")

    right_card = patches.FancyBboxPatch((76.5, 1.5), 22.0, 84.0, boxstyle="round,pad=0.25,rounding_size=0.8", facecolor="#ffffff", edgecolor="#cbd5e1", linewidth=1)
    ax.add_patch(right_card)
    ax.text(78.0, 83.5, "EXECUTIVE BRIEFING", color="#0f172a", fontsize=11.0, fontweight="bold")

    # Table
    df_sorted = df_day.sort_values("Hours", ascending=False).head(20)
    left_x = 2.6
    tbl_w = 71.8
    tbl_hdr = patches.Rectangle((left_x, 79.5), tbl_w, 2.6, facecolor="#1e293b", edgecolor="none")
    ax.add_patch(tbl_hdr)
    ax.text(left_x + 1.0, 80.8, "POS", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
    ax.text(left_x + 9.0, 80.8, "MACHINE", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
    ax.text(left_x + 23.0, 80.8, "STOPPAGE CAUSE", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
    ax.text(left_x + 47.0, 80.8, "CATEGORY", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
    ax.text(left_x + 58.0, 80.8, "START TIME", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
    ax.text(left_x + 70.0, 80.8, "LOST (HRS)", color="#ffffff", fontsize=7.2, fontweight="bold", ha="right", va="center")

    row_y = 77.2
    row_step = min(3.8, 74.0 / max(1, len(df_sorted)))
    for r_i, (_, r) in enumerate(df_sorted.iterrows()):
        bg_c = "#f8fafc" if r_i % 2 == 1 else "#ffffff"
        row_bg = patches.Rectangle((left_x, row_y - 1.2), tbl_w, row_step, facecolor=bg_c, edgecolor="none")
        ax.add_patch(row_bg)
        ax.plot([left_x, left_x + tbl_w], [row_y - 1.2, row_y - 1.2], color="#e2e8f0", linewidth=0.45)

        cat_str = "Maintenance" if r["Is_Maintenance"] else "Operational"
        cat_col = "#7c3aed" if r["Is_Maintenance"] else "#475569"
        from_str = str(r.get("From Time", "-"))[:16]

        ax.text(left_x + 1.0, row_y + 0.35, str(r["Position"]), color="#0f172a", fontsize=6.8, fontweight="bold", va="center")
        ax.text(left_x + 9.0, row_y + 0.35, str(r["Machine"]), color="#64748b", fontsize=6.6, va="center")
        ax.text(left_x + 23.0, row_y + 0.35, str(r["Cause"])[:24], color="#b91c1c" if r["Is_Maintenance"] else "#0f172a", fontsize=6.6, va="center")
        ax.text(left_x + 47.0, row_y + 0.35, cat_str, color=cat_col, fontsize=6.6, fontweight="bold", va="center")
        ax.text(left_x + 58.0, row_y + 0.35, from_str, color="#64748b", fontsize=6.4, va="center")
        ax.text(left_x + 70.0, row_y + 0.35, f"{r['Hours']:.2f} h", color="#dc2626" if r["Hours"] >= 4.0 else "#0f172a", fontsize=7.0, fontweight="bold", ha="right", va="center")
        row_y -= row_step

    # Right Card 1: Pareto Drivers
    c1 = patches.FancyBboxPatch((77.5, 42.5), 20.0, 39.0, boxstyle="round,pad=0.2,rounding_size=0.5", facecolor="#fff7f7", edgecolor="#fecaca", linewidth=0.8)
    ax.add_patch(c1)
    ax.text(78.6, 78.8, "NPT Pareto Root Causes", color="#b91c1c", fontsize=9.6, fontweight="bold")
    top_causes_txt = "\n".join([f"  {idx+1}. {c[:17]}: {h:.1f}h ({p:.1f}%)" for idx, (c, h, p) in enumerate(top_causes_list[:4])])
    t1 = (
        f"• Top Contributing Stoppages:\n{top_causes_txt}\n\n"
        f"• Critical Maintenance Line:\n  {top_maint_pos} ({top_maint_mc})\n  Loss: {top_maint_hrs:.1f} Hours.\n\n"
        f"• Corrective Directives:\n  - Prioritize technician assignment\n    on recurring breakdown lines.\n  - Monitor spare parts buffer."
    )
    ax.text(78.6, 75.2, t1, color="#7f1d1d", fontsize=8.0, linespacing=1.45, va="top")

    # Right Card 2: Plant & Month-To-Date Overview
    c2 = patches.FancyBboxPatch((77.5, 2.5), 20.0, 38.5, boxstyle="round,pad=0.2,rounding_size=0.5", facecolor="#f0fdf4", edgecolor="#bbf7d0", linewidth=0.8)
    ax.add_patch(c2)
    ax.text(78.6, 38.2, "Month-to-Date & Maintenance Loss", color="#15803d", fontsize=9.6, fontweight="bold")
    t2 = (
        f"• MTD Wasted Capacity (Day 1–{day_num}):\n  {curr_as_of_total_hrs:,.1f} H lost of {curr_as_of_avail_hrs:,.0f} H total\n  ({as_of_loss_pct:.1f}% Plant Loss MTD).\n\n"
        f"• MTD Maintenance Breakdown:\n  - Wasted: {curr_as_of_maint_hrs:,.1f} Hours\n  - Plant Loss: {as_of_maint_loss_pct:.1f}% of Available\n  - NPT Share: {as_of_maint_share_pct:.1f}% of Total NPT\n\n"
        f"• Stoppage Hours by Floor:\n  - Ground Floor: {gf_share_pct:.1f}% of loss\n  - First Floor: {ff_share_pct:.1f}% of loss"
    )
    ax.text(78.6, 34.6, t2, color="#166534", fontsize=8.0, linespacing=1.45, va="top")

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = io.BytesIO()
    plt.savefig(buf, format="jpg", facecolor=fig.get_facecolor(), edgecolor="none", dpi=220)
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


# =========================================================
# 4. STREAMLIT RENDER ENTRY POINT
# =========================================================
def render_npt_module():
    c_back, c_title, c_act = st.columns([1.5, 3.5, 1.5], vertical_alignment="center")
    with c_back:
        if st.button("⬅️ Back to Operations Hub", use_container_width=True):
            st.session_state["active_view"] = "hub_home"
            st.rerun()
    with c_title:
        st.markdown("<h3 style='margin:0; text-align:center; font-weight:800; color:#0f172a;'>⏱️ NON-PRODUCTIVE TIME (NPT) ANALYTICS</h3>", unsafe_allow_html=True)
    with c_act:
        if "m3_file_bytes" in st.session_state:
            if st.button("🔄 Change Files", use_container_width=True):
                st.session_state.pop("m3_file_bytes", None)
                st.session_state.pop("m3_sm_bytes", None)
                st.rerun()

    st.divider()

    if "m3_file_bytes" not in st.session_state:
        c_up1, c_up2 = st.columns(2, gap="medium")
        with c_up1:
            st.markdown(
                '<div style="background:#ffffff; padding:1.5rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #8b5cf6;">'
                '<h4 style="margin-top:0; color:#0f172a;">📂 1. Primary Downtime Report (Required)</h4>'
                '<p style="color:#64748b !important; font-size:0.85rem;">Upload the multi-month ERP Downtime workbook (e.g. DowntimeReport.xlsx).</p></div>',
                unsafe_allow_html=True,
            )
            up_dt = st.file_uploader("Select Downtime Workbook", type=["xlsx", "xls"], key="up_dt_file")

        with c_up2:
            st.markdown(
                '<div style="background:#ffffff; padding:1.5rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #10b981;">'
                '<h4 style="margin-top:0; color:#0f172a;">🛠️ 2. Service Maintenance Report (Optional)</h4>'
                '<p style="color:#64748b !important; font-size:0.85rem;">Upload engineering workshop ticket logs for SMS token audit.</p></div>',
                unsafe_allow_html=True,
            )
            up_sm = st.file_uploader("Select Maintenance Ticket Workbook", type=["xlsx", "xls"], key="up_sm_file")

        if up_dt is not None:
            if st.button("🚀 Ingest Workbooks & Launch Console", type="primary", use_container_width=True):
                st.session_state["m3_file_bytes"] = up_dt.getvalue()
                st.session_state["m3_sm_bytes"] = up_sm.getvalue() if up_sm is not None else None
                st.rerun()

    else:
        df_downtime = m3_parse_downtime_workbook(st.session_state["m3_file_bytes"])
        df_service = m3_parse_service_maintenance(st.session_state.get("m3_sm_bytes"))

        all_months = sorted(df_downtime["YearMonth"].unique())
        active_month = all_months[-1]
        active_m_df = df_downtime[df_downtime["YearMonth"] == active_month]
        all_dates = sorted(active_m_df["DateStr"].unique().tolist())

        st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
        c_date, c_mode, c_snap = st.columns([1.5, 1.5, 1.2], gap="medium")
        with c_date:
            sel_date_str = st.selectbox("📅 **Operational Date**", all_dates, index=len(all_dates) - 1)
        with c_mode:
            comp_mode = st.radio("📊 **3-Month Comparison Scope**", ["Like-for-Like (Day 1–N)", "Full Month Baseline"], horizontal=True)

        sel_date_obj = pd.to_datetime(sel_date_str)
        sel_day_num = sel_date_obj.day
        day_formatted = sel_date_obj.strftime("%B %d")

        # Day & MTD filtered frames
        df_day = active_m_df[active_m_df["DateStr"] == sel_date_str].copy()
        df_as_of = active_m_df[active_m_df["DayNum"] <= sel_day_num].copy()

        # Day Metrics
        total_day_hrs = float(df_day["Hours"].sum())
        maint_hrs = float(df_day[df_day["Is_Maintenance"]]["Hours"].sum())
        oper_hrs = total_day_hrs - maint_hrs
        loss_pct = (total_day_hrs / DAILY_AVAILABLE_HRS) * 100.0
        maint_loss_pct = (maint_hrs / DAILY_AVAILABLE_HRS) * 100.0

        # MTD Metrics
        curr_as_of_avail_hrs = sel_day_num * DAILY_AVAILABLE_HRS
        curr_as_of_total_hrs = float(df_as_of["Hours"].sum())
        curr_as_of_maint_hrs = float(df_as_of[df_as_of["Is_Maintenance"]]["Hours"].sum())
        curr_as_of_avg_hrs = curr_as_of_total_hrs / sel_day_num
        as_of_loss_pct = (curr_as_of_total_hrs / curr_as_of_avail_hrs) * 100.0
        as_of_maint_loss_pct = (curr_as_of_maint_hrs / curr_as_of_avail_hrs) * 100.0
        as_of_maint_share_pct = (curr_as_of_maint_hrs / curr_as_of_total_hrs * 100.0) if curr_as_of_total_hrs > 0 else 0.0

        # Previous Month Metrics
        if len(all_months) >= 2:
            prev_m_df = df_downtime[df_downtime["YearMonth"] == all_months[-2]]
            prev_total_hrs = float(prev_m_df["Hours"].sum())
            prev_days_count = prev_m_df["DateClean"].dt.days_in_month.iloc[0]
            prev_avg_hrs = prev_total_hrs / prev_days_count
        else:
            prev_total_hrs, prev_avg_hrs = 0.0, 0.0

        # Floor distribution
        gf_hrs = df_day[df_day["Line"].str.startswith("GF")]["Hours"].sum()
        ff_hrs = df_day[df_day["Line"].str.startswith("FF")]["Hours"].sum()
        tot_flr = (gf_hrs + ff_hrs) if (gf_hrs + ff_hrs) > 0 else 1.0
        gf_share_pct = (gf_hrs / tot_flr) * 100.0
        ff_share_pct = (ff_hrs / tot_flr) * 100.0

        # Top cause and maintenance line
        df_cause_day = (
            df_day.groupby("Cause")["Hours"].sum().reset_index().sort_values("Hours", ascending=False)
        )
        if not df_cause_day.empty:
            top_cause_name = df_cause_day.iloc[0]["Cause"]
            top_cause_hrs = df_cause_day.iloc[0]["Hours"]
            top_cause_pct = (top_cause_hrs / total_day_hrs * 100.0) if total_day_hrs > 0 else 0.0
            top_causes_list = [(r["Cause"], r["Hours"], (r["Hours"] / total_day_hrs * 100)) for _, r in df_cause_day.head(4).iterrows()]
        else:
            top_cause_name, top_cause_hrs, top_cause_pct = "None", 0.0, 0.0
            top_causes_list = []

        maint_day_lines = df_day[df_day["Is_Maintenance"]].groupby(["Position", "Machine"])["Hours"].sum().reset_index().sort_values("Hours", ascending=False)
        if not maint_day_lines.empty:
            top_maint_pos = maint_day_lines.iloc[0]["Position"]
            top_maint_mc = maint_day_lines.iloc[0]["Machine"]
            top_maint_hrs = maint_day_lines.iloc[0]["Hours"]
        else:
            top_maint_pos, top_maint_mc, top_maint_hrs = "-", "None", 0.0

        # JPG Visual Generation
        jpg_bytes = m3_generate_executive_jpg(
            df_day, sel_date_obj, total_day_hrs, maint_hrs, oper_hrs, loss_pct, maint_loss_pct,
            curr_as_of_total_hrs, curr_as_of_avail_hrs, as_of_loss_pct,
            curr_as_of_maint_hrs, as_of_maint_loss_pct, as_of_maint_share_pct,
            gf_share_pct, ff_share_pct, top_cause_name, top_cause_hrs, top_cause_pct,
            top_maint_pos, top_maint_mc, top_maint_hrs, prev_total_hrs, prev_avg_hrs,
            curr_as_of_avg_hrs, top_causes_list
        )

        with c_snap:
            st.markdown("<div style='margin-top: 1.6rem;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📸 Download 1-Page JPG Report",
                data=jpg_bytes,
                file_name=f"Daily_NPT_Report_{sel_date_str}.jpg",
                mime="image/jpeg",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        # 6 Web KPI Cards
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(f'<div class="kpi-card indigo"><div class="kpi-title">PREV MO. TOTAL</div><div class="kpi-val">{prev_total_hrs:.1f} H</div><div class="kpi-sub">Daily Avg: {prev_avg_hrs:.1f} H/D</div></div>', unsafe_allow_html=True)
        k2.markdown(f'<div class="kpi-card blue"><div class="kpi-title">THIS MO. AS OF</div><div class="kpi-val">{curr_as_of_total_hrs:.1f} H</div><div class="kpi-sub">Wasted: {as_of_loss_pct:.1f}% Cap</div></div>', unsafe_allow_html=True)
        k3.markdown(f'<div class="kpi-card purple"><div class="kpi-title">MTD MAINT. LOSS</div><div class="kpi-val">{curr_as_of_maint_hrs:.1f} H</div><div class="kpi-sub">{as_of_maint_loss_pct:.1f}% Cap ({as_of_maint_share_pct:.1f}% NPT)</div></div>', unsafe_allow_html=True)
        k4.markdown(f'<div class="kpi-card pink"><div class="kpi-title">LAST DAY NPT</div><div class="kpi-val">{total_day_hrs:.1f} H</div><div class="kpi-sub">{loss_pct:.1f}% Available Lost</div></div>', unsafe_allow_html=True)
        k5.markdown(f'<div class="kpi-card yellow"><div class="kpi-title">DAY MAINT. LOSS</div><div class="kpi-val">{maint_hrs:.1f} H</div><div class="kpi-sub">{maint_loss_pct:.1f}% Available</div></div>', unsafe_allow_html=True)
        k6.markdown(f'<div class="kpi-card teal"><div class="kpi-title">TOP DRIVER</div><div class="kpi-val">{top_cause_hrs:.1f} H</div><div class="kpi-sub">{top_cause_name[:14]}</div></div>', unsafe_allow_html=True)

        st.markdown("<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True)

        # Mid Section: Log Table & WhatsApp Note
        col_left, col_right = st.columns([1.55, 0.95], gap="large")
        with col_left:
            st.markdown(f"#### ⚙️ MACHINE NPT INCIDENTS LOG — {day_formatted}")
            display_cols = ["Position", "Machine", "Line", "Cause", "Hours", "From Time", "To Time", "Cause Added By"]
            st.dataframe(df_day[display_cols].sort_values("Hours", ascending=False), use_container_width=True, hide_index=True, height=380)

        with col_right:
            npt_brief_text = f"""📋 *PLASTIC-3 DAILY NPT & DOWNTIME BRIEF*
📅 *Date:* {day_formatted}

Dear Sir,

🎯 *Operational Availability & Stoppage Summary*
Total downtime reached *{total_day_hrs:.1f} Hours* ({loss_pct:.1f}% of total plant capacity across 61 IMMs).
Month-to-Date (Day 1–{sel_day_num}) has lost *{curr_as_of_total_hrs:,.1f} Hours* (*{as_of_loss_pct:.1f}%* of total potential capacity).

🔧 *Maintenance Loss Summary (MTD & Daily)*
• MTD Maintenance Wasted: *{curr_as_of_maint_hrs:.1f} Hours* ({as_of_maint_loss_pct:.1f}% of plant capacity, {as_of_maint_share_pct:.1f}% of total NPT)
• Last Day Maintenance: *{maint_hrs:.1f} Hours* ({maint_loss_pct:.1f}% of available time)
• Last Day Operational Loss: *{oper_hrs:.1f} Hours*

🏆 *Top Stoppage Contributor*
*{top_cause_name}* accounted for *{top_cause_hrs:.1f} Hours* ({top_cause_pct:.1f}% of day NPT).
Heaviest breakdown line: *{top_maint_pos}* ({top_maint_mc}) with *{top_maint_hrs:.1f} Hours* lost."""

            st.markdown("#### 📝 EXECUTIVE BRIEFING TEXT")
            st.markdown(
                f"""<div class="narrative-block">
                    <p style="margin: 0 0 0.5rem 0; font-weight: 800; color: #1e293b;">📋 PLASTIC-3 DAILY NPT & DOWNTIME BRIEF</p>
                    <p style="margin: 0 0 0.75rem 0; color: #64748b; font-size: 0.82rem;">📅 <b>Date:</b> {day_formatted}</p>
                    <h5>🎯 Operational Availability & Stoppage Summary</h5>
                    <p>Total downtime reached <b>{total_day_hrs:.1f} Hours</b> (<b style="color: #dc2626;">{loss_pct:.1f}%</b> of total plant capacity across 61 IMMs).<br>
                    Month-to-Date (Day 1–{sel_day_num}) has lost <b>{curr_as_of_total_hrs:,.1f} Hours</b> (<b>{as_of_loss_pct:.1f}%</b> of total potential capacity).</p>
                    <h5>🔧 Maintenance Loss Summary (MTD & Daily)</h5>
                    <p>• MTD Maintenance Wasted: <b style="color: #7c3aed;">{curr_as_of_maint_hrs:.1f} Hours</b> (<b>{as_of_maint_loss_pct:.1f}%</b> of capacity, <b>{as_of_maint_share_pct:.1f}%</b> of total NPT)<br>
                    • Last Day Maintenance: <b>{maint_hrs:.1f} Hours</b> ({maint_loss_pct:.1f}% of available time)<br>
                    • Last Day Operational Loss: <b>{oper_hrs:.1f} Hours</b></p>
                    <h5>🏆 Top Stoppage Contributor</h5>
                    <p><b>{top_cause_name}</b> accounted for <b>{top_cause_hrs:.1f} Hours</b> ({top_cause_pct:.1f}% of day NPT).<br>
                    Heaviest breakdown line: <b>{top_maint_pos}</b> ({top_maint_mc}) with <b>{top_maint_hrs:.1f} Hours</b> lost.</p>
                </div>""",
                unsafe_allow_html=True,
            )
            with st.expander("📋 Copy Plain Text Brief"):
                st.text_area("Brief Text", value=npt_brief_text, height=190, label_visibility="collapsed")

        st.divider()

        # =========================================================
        # 4 TABS: COMPARISON, SMED, MAINTENANCE AUDIT, DATE SUMMARY
        # =========================================================
        tab_comp, tab_smed, tab_maint, tab_date = st.tabs([
            "📊 3-Month Trend & Comparison",
            "⏱️ SMED (Mold Changeover)",
            "🛠️ Maintenance & SMS Audit",
            "📅 Date-Wise Plant Downtime",
        ])

        with tab_comp:
            mode_param = "like_for_like" if "Like-for-Like" in comp_mode else "full_month"
            comp_df = m3_compute_month_comparison(df_downtime, active_month, all_months, sel_day_num, mode=mode_param)
            st.markdown(f"#### 📈 3-MONTH CAUSE-WISE BREAKDOWN ({comp_mode} | Cutoff Day {sel_day_num})")
            st.dataframe(comp_df, use_container_width=True, hide_index=True)

        with tab_smed:
            st.markdown(f"#### ⏱️ MOLD CHANGEOVER (SMED) PERFORMANCE — {active_m_df['MonthName'].iloc[0]}")
            c_smed1, c_smed2 = st.columns([1.6, 1.0], gap="large")
            daily_smed, size_smed = m3_compute_smed_summary(active_m_df)

            with c_smed1:
                st.markdown("##### 📅 Daily SMED Execution Log")
                if not daily_smed.empty:
                    st.dataframe(daily_smed, use_container_width=True, hide_index=True)
                else:
                    st.info("No Mold Change* events logged in this month.")

            with c_smed2:
                st.markdown("##### 📏 Size-Wise SMED Duration")
                if not size_smed.empty:
                    st.dataframe(size_smed, use_container_width=True, hide_index=True)
                else:
                    st.info("No data available.")

        with tab_maint:
            st.markdown("#### 🛠️ ENGINEERING WORKSHOP SMS TOKEN & DOWNTIME CORRELATION")
            df_sm_month = df_service[df_service["YearMonth"] == active_month] if not df_service.empty else pd.DataFrame()
            corr_df = m3_compute_maintenance_ticket_correlation(active_m_df, df_sm_month)
            if not corr_df.empty:
                st.dataframe(corr_df, use_container_width=True, hide_index=True)
            else:
                st.info("Upload the Service Maintenance History Report on initial screen to populate SMS tokens.")

        with tab_date:
            st.markdown("#### 📅 DAILY PLANT DOWNTIME SUMMARY")
            date_summary = (
                active_m_df.groupby(["DateStr", "DateClean"])
                .agg(
                    Total_Hours=("Hours", "sum"),
                    Maint_Hours=("Hours", lambda x: x[active_m_df.loc[x.index, "Is_Maintenance"]].sum()),
                    Total_Logs=("Hours", "count"),
                    Active_MCs=("Machine", "nunique"),
                )
                .reset_index()
            )
            date_summary["Loss_Pct"] = ((date_summary["Total_Hours"] / DAILY_AVAILABLE_HRS) * 100).round(2).apply(lambda x: f"{x:.2f}%")
            date_summary["Total_Hours"] = date_summary["Total_Hours"].round(2)
            date_summary["Maint_Hours"] = date_summary["Maint_Hours"].round(2)
            st.dataframe(date_summary, use_container_width=True, hide_index=True)
