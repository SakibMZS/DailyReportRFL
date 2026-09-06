import io
import textwrap
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches

from config import (
    TOTAL_PLANT_MCS,
    DAILY_AVAILABLE_HRS,
    POS_MAP,
    LINE_MAP,
    MAINTENANCE_CAUSES,
)


def get_col(df, candidates, default=None):
    for c in candidates:
        if c in df.columns:
            return c
    return default


@st.cache_data
def m3_parse_workbook(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)

    sheet_name = (
        "DowntimeReport"
        if "DowntimeReport" in xls.sheet_names
        else xls.sheet_names[0]
    )
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values])
        if "Machine" in row_str and ("Duration" in row_str or "Cause" in row_str):
            header_idx = idx
            break

    if header_idx is not None:
        df_clean = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx)
    else:
        df_clean = pd.read_excel(xls, sheet_name=sheet_name)

    df_clean.columns = [str(c).strip() for c in df_clean.columns]

    # Exclude open-ended stoppages without end timestamp
    to_col = get_col(df_clean, ["To Time", "ToTime", "End Time", "EndTime"])
    if to_col:
        df_clean = df_clean[df_clean[to_col].notna()].copy()

    # Base operational date on Cause Added Date
    date_col = get_col(
        df_clean,
        ["Cause Added Date", "CauseAddedDate", "Date", "Added Date"],
        df_clean.columns[-1],
    )
    df_clean["DateClean"] = pd.to_datetime(df_clean[date_col], errors="coerce")
    df_clean = df_clean.dropna(subset=["DateClean"]).sort_values("DateClean")
    df_clean["DateStr"] = df_clean["DateClean"].dt.strftime("%Y-%m-%d")
    df_clean["YearMonth"] = df_clean["DateClean"].dt.to_period("M")

    # Duration Hours
    sec_col = get_col(
        df_clean,
        ["Duration (In Second)", "Duration(In Second)", "Duration(s)", "Seconds"],
    )
    dur_col = get_col(df_clean, ["Duration", "Duration (Hrs)", "Hours"])

    if sec_col in df_clean.columns:
        df_clean["Hours"] = (
            pd.to_numeric(df_clean[sec_col], errors="coerce").fillna(0) / 3600.0
        )
    elif dur_col in df_clean.columns:
        df_clean["Hours"] = (
            pd.to_timedelta(df_clean[dur_col], errors="coerce").dt.total_seconds()
            / 3600.0
        )
    else:
        df_clean["Hours"] = 0.0

    df_clean = df_clean[df_clean["Hours"] > 0].copy()

    # Shop floor mapping via central config
    mc_col = get_col(df_clean, ["Machine", "MC SL"], "Machine")
    df_clean["Position"] = df_clean[mc_col].astype(str).map(POS_MAP).fillna("-")
    df_clean["Line"] = df_clean[mc_col].astype(str).map(LINE_MAP).fillna("-")

    # Maintenance Tagging (5 Selected Categories)
    cause_col = get_col(df_clean, ["Cause", "Causes", "Reason", "Defect"], "Cause")
    df_clean["CauseClean"] = (
        df_clean[cause_col].astype(str).str.replace("*", "", regex=False).str.strip()
    )
    df_clean["Is_Maintenance"] = df_clean[cause_col].isin(MAINTENANCE_CAUSES)

    unique_months = sorted(df_clean["YearMonth"].unique())
    if len(unique_months) >= 2:
        df_prev = df_clean[df_clean["YearMonth"] == unique_months[-2]].copy()
        df_curr = df_clean[df_clean["YearMonth"] == unique_months[-1]].copy()
    else:
        df_prev = pd.DataFrame()
        df_curr = df_clean.copy()

    return df_prev, df_curr, df_clean


def m3_compute_cause_summary(df_scope):
    if df_scope.empty:
        return pd.DataFrame()

    total_hrs = df_scope["Hours"].sum()
    res = (
        df_scope.groupby("Cause")
        .agg(
            Hours=("Hours", "sum"),
            Logs_Count=("Hours", "count"),
            MC_Count=("Machine", "nunique"),
            Is_Maintenance=("Is_Maintenance", "first"),
        )
        .reset_index()
    )

    res["Percentage"] = (res["Hours"] / total_hrs * 100.0) if total_hrs > 0 else 0.0
    res = res.sort_values("Hours", ascending=False).reset_index(drop=True)
    return res


def m3_compute_date_summary(df_curr):
    if df_curr.empty:
        return pd.DataFrame()

    res = (
        df_curr.groupby(["DateStr", "DateClean"])
        .agg(
            Total_Hours=("Hours", "sum"),
            Maint_Hours=(
                "Hours",
                lambda x: x[df_curr.loc[x.index, "Is_Maintenance"]].sum(),
            ),
            Operational_Hours=(
                "Hours",
                lambda x: x[~df_curr.loc[x.index, "Is_Maintenance"]].sum(),
            ),
            Total_Logs=("Hours", "count"),
            Active_MCs=("Machine", "nunique"),
        )
        .reset_index()
    )

    res["Plant_Loss_Pct"] = (res["Total_Hours"] / DAILY_AVAILABLE_HRS) * 100.0
    res["Maint_Loss_Pct"] = (res["Maint_Hours"] / DAILY_AVAILABLE_HRS) * 100.0
    res = res.sort_values("DateClean").reset_index(drop=True)
    return res


def m3_compute_machine_maintenance(df_scope):
    maint_df = df_scope[df_scope["Is_Maintenance"]].copy()
    if maint_df.empty:
        return pd.DataFrame()

    res = (
        maint_df.groupby(["Position", "Machine", "Line", "Cause"])
        .agg(Hours=("Hours", "sum"), Incidents=("Hours", "count"))
        .reset_index()
    )
    res = res.sort_values("Hours", ascending=False).reset_index(drop=True)
    return res


def m3_generate_npt_jpg(
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

    # 1. Header
    ax.text(
        1.5,
        98.4,
        "DAILY NON-PRODUCTIVE TIME (NPT) ANALYTICS",
        color="#0f172a",
        fontsize=16.0,
        fontweight="bold",
        va="top",
    )
    ax.text(
        1.5,
        95.8,
        f"Plastic-3 Stoppage & Machine Downtime Log ({TOTAL_PLANT_MCS} IMMs Baseline: 1,464 H/Day)  |  Report Date: {date_formatted}",
        color="#64748b",
        fontsize=8.8,
        va="top",
    )
    ax.text(
        98.5,
        97.2,
        "PLASTIC-3 OPERATIONS",
        color="#2563eb",
        fontsize=8.8,
        fontweight="bold",
        ha="right",
        va="top",
    )

    # 2. Top 6 KPI Cards
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
        card = patches.FancyBboxPatch(
            (x0, 87.0),
            kpi_w,
            7.2,
            boxstyle="round,pad=0.15,rounding_size=0.5",
            facecolor="#ffffff",
            edgecolor="#cbd5e1",
            linewidth=0.8,
        )
        ax.add_patch(card)
        top_bar = patches.FancyBboxPatch(
            (x0 + 0.1, 93.75),
            kpi_w - 0.2,
            0.45,
            boxstyle="round,pad=0.03,rounding_size=0.2",
            facecolor=col_bar,
            edgecolor="none",
        )
        ax.add_patch(top_bar)
        ax.text(
            x0 + kpi_w / 2, 92.4, title, color="#64748b", fontsize=7.6, fontweight="bold", ha="center"
        )
        ax.text(
            x0 + kpi_w / 2, 89.6, val, color="#0f172a", fontsize=13.0, fontweight="bold", ha="center"
        )
        ax.text(
            x0 + kpi_w / 2, 87.8, sub, color="#94a3b8", fontsize=6.8, ha="center"
        )

    # 3. Main Workspace Containers
    left_card = patches.FancyBboxPatch(
        (1.5, 1.5),
        74.0,
        84.0,
        boxstyle="round,pad=0.25,rounding_size=0.8",
        facecolor="#ffffff",
        edgecolor="#cbd5e1",
        linewidth=1,
    )
    ax.add_patch(left_card)
    ax.text(
        3.5,
        83.5,
        f"DAILY NPT MACHINE LOG (WITH POSITION) — {day_formatted}",
        color="#0f172a",
        fontsize=11.0,
        fontweight="bold",
    )
    ax.text(
        73.5,
        83.5,
        f"Plant Loss: {loss_pct:.1f}% of Available Capacity",
        color="#64748b",
        fontsize=8.0,
        ha="right",
    )

    right_card = patches.FancyBboxPatch(
        (76.5, 1.5),
        22.0,
        84.0,
        boxstyle="round,pad=0.25,rounding_size=0.8",
        facecolor="#ffffff",
        edgecolor="#cbd5e1",
        linewidth=1,
    )
    ax.add_patch(right_card)
    ax.text(
        78.0,
        83.5,
        "EXECUTIVE BRIEFING",
        color="#0f172a",
        fontsize=11.0,
        fontweight="bold",
    )

    # Left Card Table: Machine & Local Position Tag
    df_sorted = df_day.sort_values("Hours", ascending=False).head(20)
    left_x = 2.6
    tbl_w = 71.8
    tbl_hdr = patches.Rectangle(
        (left_x, 79.5), tbl_w, 2.6, facecolor="#1e293b", edgecolor="none"
    )
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
        row_bg = patches.Rectangle(
            (left_x, row_y - 1.2), tbl_w, row_step, facecolor=bg_c, edgecolor="none"
        )
        ax.add_patch(row_bg)
        ax.plot(
            [left_x, left_x + tbl_w],
            [row_y - 1.2, row_y - 1.2],
            color="#e2e8f0",
            linewidth=0.45,
        )

        cat_str = "Maintenance" if r["Is_Maintenance"] else "Operational"
        cat_col = "#7c3aed" if r["Is_Maintenance"] else "#475569"
        from_str = str(r.get("From Time", "-"))[:16]

        ax.text(left_x + 1.0, row_y + 0.35, str(r["Position"]), color="#0f172a", fontsize=6.8, fontweight="bold", va="center")
        ax.text(left_x + 9.0, row_y + 0.35, str(r["Machine"]), color="#64748b", fontsize=6.6, va="center")
        ax.text(left_x + 23.0, row_y + 0.35, str(r["Cause"])[:24], color="#b91c1c" if r["Is_Maintenance"] else "#0f172a", fontsize=6.6, va="center")
        ax.text(left_x + 47.0, row_y + 0.35, cat_str, color=cat_col, fontsize=6.6, fontweight="bold", va="center")
        ax.text(left_x + 58.0, row_y + 0.35, from_str, color="#64748b", fontsize=6.4, va="center")
        ax.text(
            left_x + 70.0,
            row_y + 0.35,
            f"{r['Hours']:.2f} h",
            color="#dc2626" if r["Hours"] >= 4.0 else "#0f172a",
            fontsize=7.0,
            fontweight="bold",
            ha="right",
            va="center",
        )
        row_y -= row_step

    # Right Card 1: Pareto Drivers
    c1 = patches.FancyBboxPatch(
        (77.5, 42.5),
        20.0,
        39.0,
        boxstyle="round,pad=0.2,rounding_size=0.5",
        facecolor="#fff7f7",
        edgecolor="#fecaca",
        linewidth=0.8,
    )
    ax.add_patch(c1)
    ax.text(
        78.6,
        78.8,
        "NPT Pareto Root Causes",
        color="#b91c1c",
        fontsize=9.6,
        fontweight="bold",
    )

    top_causes_txt = "\n".join(
        [
            f"  {idx+1}. {c[:17]}: {h:.1f}h ({p:.1f}%)"
            for idx, (c, h, p) in enumerate(top_causes_list[:4])
        ]
    )
    t1 = (
        f"• Top Contributing Stoppages:\n"
        f"{top_causes_txt}\n\n"
        f"• Critical Maintenance Line:\n"
        f"  {top_maint_pos} ({top_maint_mc})\n"
        f"  Loss: {top_maint_hrs:.1f} Hours.\n\n"
        f"• Corrective Directives:\n"
        f"  - Prioritize technician assignment\n"
        f"    on recurring breakdown lines.\n"
        f"  - Monitor spare parts buffer."
    )
    ax.text(78.6, 75.2, t1, color="#7f1d1d", fontsize=8.0, linespacing=1.45, va="top")

    # Right Card 2: Plant & Month-To-Date Overview (Including MTD Maintenance)
    c2 = patches.FancyBboxPatch(
        (77.5, 2.5),
        20.0,
        38.5,
        boxstyle="round,pad=0.2,rounding_size=0.5",
        facecolor="#f0fdf4",
        edgecolor="#bbf7d0",
        linewidth=0.8,
    )
    ax.add_patch(c2)
    ax.text(
        78.6,
        38.2,
        "Month-to-Date & Maintenance Loss",
        color="#15803d",
        fontsize=9.6,
        fontweight="bold",
    )
    t2 = (
        f"• MTD Wasted Capacity (Day 1–{day_num}):\n"
        f"  {curr_as_of_total_hrs:,.1f} H lost of {curr_as_of_avail_hrs:,.0f} H total\n"
        f"  ({as_of_loss_pct:.1f}% Plant Loss MTD).\n\n"
        f"• MTD Maintenance Breakdown:\n"
        f"  - Wasted: {curr_as_of_maint_hrs:,.1f} Hours\n"
        f"  - Plant Loss: {as_of_maint_loss_pct:.1f}% of Available\n"
        f"  - NPT Share: {as_of_maint_share_pct:.1f}% of Total NPT\n\n"
        f"• Stoppage Hours by Floor:\n"
        f"  - Ground Floor: {gf_share_pct:.1f}% of loss\n"
        f"  - First Floor: {ff_share_pct:.1f}% of loss"
    )
    ax.text(78.6, 34.6, t2, color="#166534", fontsize=8.0, linespacing=1.45, va="top")

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = io.BytesIO()
    plt.savefig(
        buf,
        format="jpg",
        facecolor=fig.get_facecolor(),
        edgecolor="none",
        dpi=220,
    )
    plt.close(fig)
    buf.seek(0)
    return buf.getvalue()


def render_npt_module():
    c_back, c_title, c_act = st.columns([1.5, 3.5, 1.5], vertical_alignment="center")
    with c_back:
        if st.button("⬅️ Back to Operations Hub", use_container_width=True):
            st.session_state["active_view"] = "hub_home"
            st.rerun()
    with c_title:
        st.markdown(
            "<h3 style='margin:0; text-align:center; font-weight:800; color:#0f172a;'>⏱️ NON-PRODUCTIVE TIME (NPT) ANALYTICS</h3>",
            unsafe_allow_html=True,
        )
    with c_act:
        if "m3_file_bytes" in st.session_state:
            if st.button("🔄 Change Excel File", use_container_width=True):
                st.session_state.pop("m3_file_bytes", None)
                st.rerun()

    st.divider()

    if "m3_file_bytes" not in st.session_state:
        c_up, _ = st.columns([2, 1])
        with c_up:
            st.markdown(
                '<div style="background:#ffffff; padding:1.75rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #8b5cf6; box-shadow: 0 4px 12px rgba(15,23,42,0.05);">'
                '<h3 style="margin-top:0; color:#0f172a;">📂 Upload Downtime / NPT Report</h3>'
                "<p style=\"color:#64748b !important;\">Select the downtime report workbook (.xlsx) containing machine stoppage logs.</p></div>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div style='margin-bottom: 1rem;'></div>", unsafe_allow_html=True
            )

            uploaded_file = st.file_uploader(
                "Select Excel File (.xlsx, .xls)",
                type=["xlsx", "xls"],
                key="m3_uploader",
            )
            if uploaded_file is not None:
                if st.button(
                    "🚀 Ingest NPT Data & Launch",
                    type="primary",
                    use_container_width=True,
                ):
                    st.session_state["m3_file_bytes"] = uploaded_file.getvalue()
                    st.rerun()
    else:
        df_prev, df_curr, df_full = m3_parse_workbook(st.session_state["m3_file_bytes"])
        all_dates = sorted(df_curr["DateStr"].unique().tolist())

        # Control Bar
        st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
        c_date, c_blank, c_snap = st.columns([1.5, 1.2, 1.5], gap="medium")
        with c_date:
            sel_date_str = st.selectbox(
                "📅 **Operational Date**", all_dates, index=len(all_dates) - 1
            )

        sel_date_obj = pd.to_datetime(sel_date_str)
        sel_day_num = sel_date_obj.day
        day_formatted = sel_date_obj.strftime("%B %d")

        # Day and MTD filtered sets
        df_day = df_curr[df_curr["DateStr"] == sel_date_str].copy()
        df_as_of = df_curr[df_curr["DateClean"].dt.day <= sel_day_num].copy()

        # Day metrics
        total_day_hrs = float(df_day["Hours"].sum())
        maint_hrs = float(df_day[df_day["Is_Maintenance"]]["Hours"].sum())
        oper_hrs = total_day_hrs - maint_hrs
        loss_pct = (total_day_hrs / DAILY_AVAILABLE_HRS) * 100.0
        maint_loss_pct = (maint_hrs / DAILY_AVAILABLE_HRS) * 100.0

        # Present Month MTD Metrics & Maintenance Wasted Percentage
        curr_as_of_avail_hrs = sel_day_num * DAILY_AVAILABLE_HRS
        if not df_as_of.empty:
            curr_as_of_total_hrs = float(df_as_of["Hours"].sum())
            curr_as_of_maint_hrs = float(df_as_of[df_as_of["Is_Maintenance"]]["Hours"].sum())
            curr_as_of_avg_hrs = curr_as_of_total_hrs / sel_day_num
            as_of_loss_pct = (curr_as_of_total_hrs / curr_as_of_avail_hrs) * 100.0
            as_of_maint_loss_pct = (curr_as_of_maint_hrs / curr_as_of_avail_hrs) * 100.0
            as_of_maint_share_pct = (
                (curr_as_of_maint_hrs / curr_as_of_total_hrs * 100.0)
                if curr_as_of_total_hrs > 0
                else 0.0
            )
        else:
            curr_as_of_total_hrs, curr_as_of_maint_hrs, curr_as_of_avg_hrs = 0.0, 0.0, 0.0
            as_of_loss_pct, as_of_maint_loss_pct, as_of_maint_share_pct = 0.0, 0.0, 0.0

        # Previous Month Metrics
        if not df_prev.empty:
            prev_total_hrs = float(df_prev["Hours"].sum())
            prev_days_count = df_prev["DateClean"].dt.days_in_month.iloc[0]
            prev_avg_hrs = prev_total_hrs / prev_days_count
        else:
            prev_total_hrs, prev_avg_hrs = 0.0, 0.0

        # Shop Floor Distribution
        gf_hrs = df_day[df_day["Line"].str.startswith("GF")]["Hours"].sum()
        ff_hrs = df_day[df_day["Line"].str.startswith("FF")]["Hours"].sum()
        tot_flr = (gf_hrs + ff_hrs) if (gf_hrs + ff_hrs) > 0 else 1.0
        gf_share_pct = (gf_hrs / tot_flr) * 100.0
        ff_share_pct = (ff_hrs / tot_flr) * 100.0

        # Cause breakdowns
        df_cause_day = m3_compute_cause_summary(df_day)
        df_cause_asof = m3_compute_cause_summary(df_as_of)
        df_date_summary = m3_compute_date_summary(df_curr)
        df_maint_day = m3_compute_machine_maintenance(df_day)
        df_maint_asof = m3_compute_machine_maintenance(df_as_of)

        # Top cause
        if not df_cause_day.empty:
            top_cause_row = df_cause_day.iloc[0]
            top_cause_name = str(top_cause_row["Cause"])
            top_cause_hrs = float(top_cause_row["Hours"])
            top_cause_pct = float(top_cause_row["Percentage"])
            top_causes_list = [
                (r["Cause"], r["Hours"], r["Percentage"])
                for _, r in df_cause_day.head(4).iterrows()
            ]
        else:
            top_cause_name, top_cause_hrs, top_cause_pct = "None", 0.0, 0.0
            top_causes_list = []

        # Top maintenance machine & position
        if not df_maint_day.empty:
            top_maint_row = df_maint_day.iloc[0]
            top_maint_pos = str(top_maint_row["Position"])
            top_maint_mc = str(top_maint_row["Machine"])
            top_maint_hrs = float(top_maint_row["Hours"])
        else:
            top_maint_pos, top_maint_mc, top_maint_hrs = "-", "None", 0.0

        # Visual JPG generation
        jpg_bytes_npt = m3_generate_npt_jpg(
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
        )

        with c_snap:
            st.markdown(
                "<div style='margin-top: 1.6rem;'></div>", unsafe_allow_html=True
            )
            st.download_button(
                label="📸 Download 1-Page JPG Report",
                data=jpg_bytes_npt,
                file_name=f"Daily_NPT_Report_{sel_date_str}.jpg",
                mime="image/jpeg",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        # 6 KPI Cards in Web Dashboard (Highlighting MTD Maintenance)
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(
            f'<div class="kpi-card indigo"><div class="kpi-title">PREV MO. TOTAL</div><div class="kpi-val">{prev_total_hrs:.1f} H</div><div class="kpi-sub">Daily Avg: {prev_avg_hrs:.1f} H/D</div></div>',
            unsafe_allow_html=True,
        )
        k2.markdown(
            f'<div class="kpi-card blue"><div class="kpi-title">THIS MO. AS OF</div><div class="kpi-val">{curr_as_of_total_hrs:.1f} H</div><div class="kpi-sub">Wasted: {as_of_loss_pct:.1f}% Cap</div></div>',
            unsafe_allow_html=True,
        )
        k3.markdown(
            f'<div class="kpi-card purple"><div class="kpi-title">MTD MAINT. LOSS</div><div class="kpi-val">{curr_as_of_maint_hrs:.1f} H</div><div class="kpi-sub">{as_of_maint_loss_pct:.1f}% Cap ({as_of_maint_share_pct:.1f}% NPT)</div></div>',
            unsafe_allow_html=True,
        )
        k4.markdown(
            f'<div class="kpi-card pink"><div class="kpi-title">LAST DAY NPT</div><div class="kpi-val">{total_day_hrs:.1f} H</div><div class="kpi-sub">{loss_pct:.1f}% Available Lost</div></div>',
            unsafe_allow_html=True,
        )
        k5.markdown(
            f'<div class="kpi-card yellow"><div class="kpi-title">DAY MAINT. LOSS</div><div class="kpi-val">{maint_hrs:.1f} H</div><div class="kpi-sub">{maint_loss_pct:.1f}% Available</div></div>',
            unsafe_allow_html=True,
        )
        k6.markdown(
            f'<div class="kpi-card teal"><div class="kpi-title">TOP DRIVER</div><div class="kpi-val">{top_cause_hrs:.1f} H</div><div class="kpi-sub">{top_cause_name[:14]}</div></div>',
            unsafe_allow_html=True,
        )

        st.markdown(
            "<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True
        )

        # Mid Section: Log Table & WhatsApp Brief
        col_left, col_right = st.columns([1.55, 0.95], gap="large")
        with col_left:
            st.markdown(f"#### ⚙️ MACHINE NPT INCIDENTS LOG — {day_formatted}")
            if not df_day.empty:
                display_cols = [
                    "Position",
                    "Machine",
                    "Line",
                    "Cause",
                    "Hours",
                    "From Time",
                    "To Time",
                    "Cause Added By",
                ]
                st.dataframe(
                    df_day[display_cols].sort_values("Hours", ascending=False),
                    use_container_width=True,
                    hide_index=True,
                    height=390,
                )
            else:
                st.success("✅ Zero downtime logged for this date!")

        with col_right:
            npt_brief_text = f"""📋 *PLASTIC-3 DAILY NPT & DOWNTIME BRIEF*
📅 *Date:* {day_formatted}

Dear Sir,

🎯 *Operational Availability & Stoppage Summary*
Total downtime reached *{total_day_hrs:.1f} Hours* ({loss_pct:.1f}% of total plant capacity across 61 IMMs).
Month-to-Date (Day 1–{sel_day_num}) has lost *{curr_as_of_total_hrs:,.1f} Hours* (*{as_of_loss_pct:.1f}%* of total potential capacity).

🔧 *Maintenance Loss Summary (MTD & Daily)*
• MTD Maintenance Wasted: *{curr_as_of_maint_hrs:.1f} Hours* ({as_of_maint_loss_pct:.1f}% of available plant time, {as_of_maint_share_pct:.1f}% of total NPT)
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
                    <p>• MTD Maintenance Wasted: <b style="color: #7c3aed;">{curr_as_of_maint_hrs:.1f} Hours</b> (<b>{as_of_maint_loss_pct:.1f}%</b> of available plant capacity, <b>{as_of_maint_share_pct:.1f}%</b> of total NPT)<br>
                    • Last Day Maintenance: <b>{maint_hrs:.1f} Hours</b> ({maint_loss_pct:.1f}% of available time)<br>
                    • Last Day Operational Loss: <b>{oper_hrs:.1f} Hours</b></p>
                    <h5>🏆 Top Stoppage Contributor</h5>
                    <p><b>{top_cause_name}</b> accounted for <b>{top_cause_hrs:.1f} Hours</b> ({top_cause_pct:.1f}% of day NPT).<br>
                    Heaviest breakdown line: <b>{top_maint_pos}</b> ({top_maint_mc}) with <b>{top_maint_hrs:.1f} Hours</b> lost.</p>
                </div>""",
                unsafe_allow_html=True,
            )

            with st.expander("📋 Copy Plain Text Brief"):
                st.text_area(
                    "Brief Text",
                    value=npt_brief_text,
                    height=200,
                    label_visibility="collapsed",
                )

        st.divider()

        # Section 3: Cause-Wise Breakdown Tabs
        st.markdown("#### 🔍 CAUSE-WISE DOWNTIME ANALYSIS")
        tab_cause_day, tab_cause_asof = st.tabs(
            [
                f"📅 Selected Date ({day_formatted})",
                f"📈 Month-to-Date Cumulative (Day 1 – {sel_day_num})",
            ]
        )

        display_cause_cols = [
            "Cause",
            "Hours",
            "Percentage",
            "Logs_Count",
            "MC_Count",
            "Is_Maintenance",
        ]
        with tab_cause_day:
            st.dataframe(
                df_cause_day[display_cause_cols],
                use_container_width=True,
                hide_index=True,
            )
        with tab_cause_asof:
            st.dataframe(
                df_cause_asof[display_cause_cols],
                use_container_width=True,
                hide_index=True,
            )

        st.divider()

        # Section 4: Maintenance Stoppages Sorted (Day & MTD)
        st.markdown("#### 🛠️ MAINTENANCE ISSUE BREAKDOWN (WITH POSITIONS)")
        tab_maint_day, tab_maint_asof = st.tabs(
            [
                f"📅 Selected Date Breakdown ({day_formatted})",
                f"📈 Month-to-Date Breakdowns (Day 1 – {sel_day_num})",
            ]
        )

        with tab_maint_day:
            if not df_maint_day.empty:
                st.dataframe(df_maint_day, use_container_width=True, hide_index=True)
            else:
                st.info("No maintenance stoppages occurred on this date.")

        with tab_maint_asof:
            if not df_maint_asof.empty:
                st.dataframe(df_maint_asof, use_container_width=True, hide_index=True)
            else:
                st.info("No maintenance stoppages recorded for this month.")

        st.divider()

        # Section 5: Date-Wise Daily Trend
        st.markdown("#### 📅 DATE-WISE PLANT DOWNTIME SUMMARY")
        st.dataframe(df_date_summary, use_container_width=True, hide_index=True)
