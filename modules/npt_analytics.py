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

# Standard Plastic-3 Machine Quantity Distribution by Size
SIZE_NOS_MAP = {
    "160": 9,
    "120": 13,
    "90": 3,
    "280": 3,
    "380": 19,
    "330": 4,
    "470": 1,
    "530": 3,
    "800": 2,
    "270": 1,
    "250": 1,
    "428": 2,
}

DEFAULT_TOP_10_CAUSES = [
    "Manpower Short*",
    "No Demand*",
    "Mold Problem*",
    "Machine Problem*",
    "Robot Problem*",
    "Sample + Mold Test (RND)*",
    "Power Breakdown (Unscheduled)*",
    "Product Jam*",
    "Mold Change*",
    "Color variation*",
]


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
# 1. PARSING ENGINES
# =========================================================
@st.cache_data
def m3_parse_downtime_workbook(file_bytes):
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

    df_clean = (
        pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx)
        if header_idx is not None
        else pd.read_excel(xls, sheet_name=sheet_name)
    )
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
    df_clean["MonthName"] = df_clean["DateClean"].dt.strftime("%B")
    df_clean["DayNum"] = df_clean["DateClean"].dt.day
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

    # Machine Positions and Sizes
    mc_col = get_col(df_clean, ["Machine", "MC SL"], "Machine")
    df_clean["Position"] = df_clean[mc_col].astype(str).map(POS_MAP).fillna("-")
    df_clean["Line"] = df_clean[mc_col].astype(str).map(LINE_MAP).fillna("-")
    df_clean["Size"] = df_clean.apply(
        lambda r: extract_mc_size(r["Position"], r[mc_col]), axis=1
    )

    # Clean Cause & Maintenance flags
    cause_col = get_col(df_clean, ["Cause", "Causes", "Reason", "Defect"], "Cause")
    df_clean["CauseClean"] = (
        df_clean[cause_col].astype(str).str.replace("*", "", regex=False).str.strip()
    )
    df_clean["Is_Maintenance"] = df_clean[cause_col].isin(MAINTENANCE_CAUSES)
    df_clean["Is_SMED"] = (
        df_clean[cause_col].astype(str).str.strip() == "Mold Change*"
    )

    return df_clean


@st.cache_data
def m3_parse_service_maintenance(file_bytes):
    if not file_bytes:
        return pd.DataFrame()
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)
    sheet_name = (
        "ServiceMaintenanceHistoryReport"
        if "ServiceMaintenanceHistoryReport" in xls.sheet_names
        else xls.sheet_names[0]
    )
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    hdr_idx = None
    for idx, r in df_raw.iterrows():
        r_str = " ".join([str(v) for v in r.values])
        if "TicketId" in r_str and "Maintenance Type" in r_str:
            hdr_idx = idx
            break

    df_clean = (
        pd.read_excel(xls, sheet_name=sheet_name, skiprows=hdr_idx)
        if hdr_idx is not None
        else pd.read_excel(xls, sheet_name=sheet_name)
    )
    df_clean.columns = [str(c).strip() for c in df_clean.columns]

    from_col = get_col(df_clean, ["From", "Start Date", "Created Date"], "From")
    df_clean["DateClean"] = pd.to_datetime(df_clean[from_col], errors="coerce")
    df_clean["MonthName"] = df_clean["DateClean"].dt.strftime("%B")
    df_clean["DayNum"] = df_clean["DateClean"].dt.day
    df_clean["YearMonth"] = df_clean["DateClean"].dt.to_period("M")

    mc_col = get_col(df_clean, ["Machine", "MC SL"], "Machine")
    df_clean["Position"] = df_clean[mc_col].astype(str).map(POS_MAP).fillna("-")
    df_clean["Size"] = df_clean.apply(
        lambda r: extract_mc_size(r["Position"], r[mc_col]), axis=1
    )

    return df_clean


# =========================================================
# 2. COMPUTATION HELPERS
# =========================================================
def m3_compute_size_wise_npt(df_scope, cutoff_days):
    records = []
    tot_hrs_all = df_scope["Hours"].sum()

    for sz in EXCEL_SIZES:
        nos = SIZE_NOS_MAP.get(sz, 0)
        sz_hrs = df_scope[df_scope["Size"] == sz]["Hours"].sum()
        avail_hrs = nos * 24.0 * max(1, cutoff_days)
        cap_pct = (sz_hrs / avail_hrs * 100.0) if avail_hrs > 0 else 0.0

        records.append({
            "MC Size": sz,
            "Nos": nos,
            "NPT hrs": round(sz_hrs, 2),
            "NPT %": round(cap_pct, 2),
        })

    df_res = pd.DataFrame(records)
    tot_avail = TOTAL_PLANT_MCS * 24.0 * max(1, cutoff_days)
    summary_pct = (tot_hrs_all / tot_avail * 100.0) if tot_avail > 0 else 0.0

    return df_res, tot_hrs_all, summary_pct


def m3_compute_smed_table(df_scope, max_days=7):
    smed_df = df_scope[df_scope["Is_SMED"]].copy()
    if smed_df.empty:
        return pd.DataFrame()

    daily = (
        smed_df.groupby(["DateStr", "DateClean"])
        .agg(
            Mold_Change_Qty=("Hours", "count"),
            Total_Time=("Hours", "sum"),
            Involved_Mcs=(
                "Position",
                lambda x: ", ".join(sorted(set(str(v) for v in x if v != "-"))),
            ),
        )
        .reset_index()
    )
    daily["Avg_SMED"] = (daily["Total_Time"] / daily["Mold_Change_Qty"] * 60.0).round(2)
    daily["Date"] = daily["DateClean"].dt.strftime("%d-%b")
    daily["Total_Time"] = daily["Total_Time"].round(2)

    # Cap to last 7 days if n > 7
    if len(daily) > max_days:
        daily = daily.tail(max_days).reset_index(drop=True)

    return daily[["Date", "Mold_Change_Qty", "Total_Time", "Avg_SMED", "Involved_Mcs"]]


def m3_compute_maint_daily_table(df_scope, max_days=7):
    maint_causes = [
        "Machine Problem*",
        "Robot Problem*",
        "Controller Problem*",
        "RMCS Problem*",
        "Oil or water Leakage*",
    ]
    dates = sorted(df_scope["DateClean"].dropna().unique())
    if len(dates) > max_days:
        dates = dates[-max_days:]

    records = []
    for dt in dates:
        dt_df = df_scope[df_scope["DateClean"] == dt]
        dt_tot = dt_df["Hours"].sum()
        row = {"Date": dt.strftime("%d-%b")}

        maint_sum = 0.0
        for mc_cause in maint_causes:
            c_hrs = dt_df[dt_df["Cause"] == mc_cause]["Hours"].sum()
            row[mc_cause] = round(c_hrs, 2)
            maint_sum += c_hrs

        row["Total Share"] = (
            f"{(maint_sum / dt_tot * 100):.0f}%" if dt_tot > 0 else "0%"
        )
        records.append(row)

    return pd.DataFrame(records)


def m3_compute_consolidated_daily_log(df_day):
    if df_day.empty:
        return pd.DataFrame()

    records = []
    for (pos, mc), grp in df_day.groupby(["Position", "Machine"]):
        causes = ", ".join(
            sorted(
                set(
                    str(c).replace("*", "").strip()
                    for c in grp["Cause"].dropna().unique()
                )
            )
        )
        line = grp["Line"].iloc[0]
        hrs = grp["Hours"].sum()
        start_t = str(grp["From Time"].min())[:16]

        records.append({
            "Position": pos,
            "Machine": mc,
            "Line": line,
            "Combined Causes": causes,
            "Total Hours": round(hrs, 2),
            "Start Time": start_t,
        })

    res = (
        pd.DataFrame(records)
        .sort_values("Total Hours", ascending=False)
        .reset_index(drop=True)
    )
    return res


# =========================================================
# 3. 2×2 GRID EXECUTIVE 1-PAGE JPG REPORT
# =========================================================
def m3_generate_2x2_executive_jpg(
    sel_date_obj,
    cutoff_day,
    top_10_causes,
    curr_share_dict,
    prev_share_dict,
    curr_month_name,
    prev_month_name,
    df_size_grid,
    size_tot_hrs,
    size_summary_pct,
    df_smed_grid,
    df_maint_grid,
):
    fig, ax = plt.subplots(figsize=(19.0, 11.2), dpi=220)
    fig.patch.set_facecolor("#f8fafc")
    ax.set_facecolor("#f8fafc")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    date_str = sel_date_obj.strftime("%d-%m-%Y")

    # Header Bar
    banner = patches.FancyBboxPatch(
        (1.5, 93.0),
        97.0,
        6.0,
        boxstyle="round,pad=0.2,rounding_size=0.4",
        facecolor="#0f172a",
        edgecolor="none",
    )
    ax.add_patch(banner)
    ax.text(
        3.5,
        96.2,
        "OPERATIONAL ANALYTICS — NON-PRODUCTIVE TIME (NPT) 4-GRID REPORT",
        color="#ffffff",
        fontsize=14.0,
        fontweight="bold",
        va="center",
    )
    ax.text(
        96.5,
        96.2,
        f"Cutoff: Day 1–{cutoff_day}  |  Operational Date: {date_str}",
        color="#94a3b8",
        fontsize=9.0,
        ha="right",
        va="center",
    )

    # 4 Grid Outer Panels
    w_box, h_box = 47.8, 43.5
    p1 = patches.FancyBboxPatch(
        (1.5, 48.0),
        w_box,
        h_box,
        boxstyle="round,pad=0.2,rounding_size=0.6",
        facecolor="#ffffff",
        edgecolor="#cbd5e1",
        linewidth=1,
    )
    p2 = patches.FancyBboxPatch(
        (50.7, 48.0),
        w_box,
        h_box,
        boxstyle="round,pad=0.2,rounding_size=0.6",
        facecolor="#ffffff",
        edgecolor="#cbd5e1",
        linewidth=1,
    )
    p3 = patches.FancyBboxPatch(
        (1.5, 3.0),
        w_box,
        h_box,
        boxstyle="round,pad=0.2,rounding_size=0.6",
        facecolor="#ffffff",
        edgecolor="#cbd5e1",
        linewidth=1,
    )
    p4 = patches.FancyBboxPatch(
        (50.7, 3.0),
        w_box,
        h_box,
        boxstyle="round,pad=0.2,rounding_size=0.6",
        facecolor="#ffffff",
        edgecolor="#cbd5e1",
        linewidth=1,
    )

    for p in [p1, p2, p3, p4]:
        ax.add_patch(p)

    # Headers for Grids
    ax.text(
        3.0,
        89.6,
        f"GRID 1: TOP 10 NPT SHARE IMPACT COMPARISON ({curr_month_name} vs {prev_month_name})",
        color="#0f172a",
        fontsize=10.0,
        fontweight="bold",
    )
    ax.text(
        52.2,
        89.6,
        f"GRID 2: MC SIZE-WISE NPT CAPACITY LOSS (Day 1–{cutoff_day})",
        color="#0f172a",
        fontsize=10.0,
        fontweight="bold",
    )
    ax.text(
        3.0,
        44.6,
        "GRID 3: MOLD CHANGEOVER (SMED) EXECUTION LOG",
        color="#0f172a",
        fontsize=10.0,
        fontweight="bold",
    )
    ax.text(
        52.2,
        44.6,
        "GRID 4: TECHNICAL & MAINTENANCE STATIONS IMPACT",
        color="#0f172a",
        fontsize=10.0,
        fontweight="bold",
    )

    # -------------------------------------------------------------
    # GRID 1: IMPACT COMPARISON CHART (NPT SHARE %)
    # -------------------------------------------------------------
    y_g1 = 86.5
    y_step_g1 = 3.6
    max_share = (
        max(
            [curr_share_dict.get(c, 0.0) for c in top_10_causes]
            + [prev_share_dict.get(c, 0.0) for c in top_10_causes]
            + [35.0]
        )
    )

    for c in top_10_causes[:10]:
        c_clean = c.replace("*", "")[:24]
        val_curr = curr_share_dict.get(c, 0.0)
        val_prev = prev_share_dict.get(c, 0.0)

        ax.text(
            3.2,
            y_g1 - 0.2,
            c_clean,
            color="#0f172a",
            fontsize=7.2,
            fontweight="bold",
            va="center",
        )

        # Bar Current
        w_curr = (val_curr / max_share) * 22.0
        bar_c = patches.Rectangle(
            (19.0, y_g1 - 0.7), w_curr, 1.2, facecolor="#dc2626", edgecolor="none"
        )
        ax.add_patch(bar_c)
        ax.text(
            19.5 + w_curr,
            y_g1 - 0.1,
            f"{val_curr:.1f}%",
            color="#dc2626",
            fontsize=6.8,
            fontweight="bold",
            va="center",
        )

        # Bar Previous
        w_prev = (val_prev / max_share) * 22.0
        bar_p = patches.Rectangle(
            (19.0, y_g1 - 2.1), w_prev, 1.2, facecolor="#94a3b8", edgecolor="none"
        )
        ax.add_patch(bar_p)
        ax.text(
            19.5 + w_prev,
            y_g1 - 1.5,
            f"{val_prev:.1f}%",
            color="#64748b",
            fontsize=6.8,
            va="center",
        )

        y_g1 -= y_step_g1

    # Grid 1 Legend
    ax.add_patch(
        patches.Rectangle(
            (32.0, 89.2), 2.0, 1.0, facecolor="#dc2626", edgecolor="none"
        )
    )
    ax.text(
        34.5,
        89.7,
        f"{curr_month_name} Share",
        color="#0f172a",
        fontsize=7.2,
        va="center",
    )
    ax.add_patch(
        patches.Rectangle(
            (41.0, 89.2), 2.0, 1.0, facecolor="#94a3b8", edgecolor="none"
        )
    )
    ax.text(
        43.5,
        89.7,
        f"{prev_month_name} Share",
        color="#0f172a",
        fontsize=7.2,
        va="center",
    )

    # -------------------------------------------------------------
    # GRID 2: MC SIZE-WISE TABLE (Exact Match image_406e3c.png)
    # -------------------------------------------------------------
    x_g2 = [52.5, 62.0, 73.0, 86.0]
    hdr_box_g2 = patches.Rectangle(
        (52.0, 85.5), 45.0, 2.4, facecolor="#0f172a", edgecolor="none"
    )
    ax.add_patch(hdr_box_g2)
    ax.text(
        x_g2[0],
        86.7,
        "Mc Size",
        color="#ffffff",
        fontsize=7.5,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g2[1],
        86.7,
        "Nos",
        color="#ffffff",
        fontsize=7.5,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g2[2],
        86.7,
        "NPT hrs",
        color="#ffffff",
        fontsize=7.5,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g2[3],
        86.7,
        "NPT %",
        color="#ffffff",
        fontsize=7.5,
        fontweight="bold",
        va="center",
    )

    y_g2 = 83.2
    step_g2 = 2.7
    for idx, r in df_size_grid.iterrows():
        bg_c = "#f8fafc" if idx % 2 == 1 else "#ffffff"
        ax.add_patch(
            patches.Rectangle(
                (52.0, y_g2 - 1.0),
                45.0,
                step_g2,
                facecolor=bg_c,
                edgecolor="none",
            )
        )
        ax.plot(
            [52.0, 97.0],
            [y_g2 - 1.0, y_g2 - 1.0],
            color="#e2e8f0",
            linewidth=0.4,
        )

        ax.text(
            x_g2[0],
            y_g2 + 0.3,
            str(r["MC Size"]),
            color="#0f172a",
            fontsize=7.2,
            fontweight="bold",
            va="center",
        )
        ax.text(
            x_g2[1],
            y_g2 + 0.3,
            str(r["Nos"]),
            color="#64748b",
            fontsize=7.2,
            va="center",
        )
        ax.text(
            x_g2[2],
            y_g2 + 0.3,
            f"{r['NPT hrs']:.1f}",
            color="#0f172a",
            fontsize=7.2,
            va="center",
        )
        ax.text(
            x_g2[3],
            y_g2 + 0.3,
            f"{r['NPT %']:.1f}%",
            color="#b91c1c" if r["NPT %"] >= 20 else "#0f172a",
            fontsize=7.2,
            fontweight="bold" if r["NPT %"] >= 20 else "normal",
            va="center",
        )
        y_g2 -= step_g2

    # Summary Row
    ax.add_patch(
        patches.Rectangle(
            (52.0, y_g2 - 1.0),
            45.0,
            step_g2,
            facecolor="#fef2f2",
            edgecolor="none",
        )
    )
    ax.text(
        x_g2[0],
        y_g2 + 0.3,
        "Summary >>",
        color="#dc2626",
        fontsize=7.6,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g2[1],
        y_g2 + 0.3,
        f"{TOTAL_PLANT_MCS}",
        color="#0f172a",
        fontsize=7.6,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g2[2],
        y_g2 + 0.3,
        f"{size_tot_hrs:,.1f}",
        color="#0f172a",
        fontsize=7.6,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g2[3],
        y_g2 + 0.3,
        f"{size_summary_pct:.1f}%",
        color="#dc2626",
        fontsize=7.6,
        fontweight="bold",
        va="center",
    )

    # -------------------------------------------------------------
    # GRID 3: SMED TABLE (Exact Match image_4013a3.png)
    # -------------------------------------------------------------
    hdr_box_g3 = patches.Rectangle(
        (3.0, 40.8), 44.5, 2.4, facecolor="#0f172a", edgecolor="none"
    )
    ax.add_patch(hdr_box_g3)
    ax.text(
        3.5,
        42.0,
        "Date",
        color="#ffffff",
        fontsize=7.0,
        fontweight="bold",
        va="center",
    )
    ax.text(
        9.0,
        42.0,
        "Mold Qty",
        color="#ffffff",
        fontsize=7.0,
        fontweight="bold",
        va="center",
    )
    ax.text(
        15.5,
        42.0,
        "Total (Hr)",
        color="#ffffff",
        fontsize=7.0,
        fontweight="bold",
        va="center",
    )
    ax.text(
        22.0,
        42.0,
        "Avg (Min)",
        color="#ffffff",
        fontsize=7.0,
        fontweight="bold",
        va="center",
    )
    ax.text(
        29.0,
        42.0,
        "Involved Machines",
        color="#ffffff",
        fontsize=7.0,
        fontweight="bold",
        va="center",
    )

    y_g3 = 38.5
    step_g3 = 4.8
    for idx, r in df_smed_grid.iterrows():
        bg_c = "#f8fafc" if idx % 2 == 1 else "#ffffff"
        ax.add_patch(
            patches.Rectangle(
                (3.0, y_g3 - 2.0),
                44.5,
                step_g3,
                facecolor=bg_c,
                edgecolor="none",
            )
        )
        ax.plot(
            [3.0, 47.5],
            [y_g3 - 2.0, y_g3 - 2.0],
            color="#e2e8f0",
            linewidth=0.4,
        )

        inv_wrap = "\n".join(textwrap.wrap(str(r["Involved_Mcs"]), width=28))
        ax.text(
            3.5,
            y_g3 + 0.3,
            str(r["Date"]),
            color="#0f172a",
            fontsize=6.8,
            fontweight="bold",
            va="center",
        )
        ax.text(
            9.5,
            y_g3 + 0.3,
            str(r["Mold_Change_Qty"]),
            color="#0f172a",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            16.0,
            y_g3 + 0.3,
            f"{r['Total_Time']:.2f}",
            color="#0f172a",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            22.5,
            y_g3 + 0.3,
            f"{r['Avg_SMED']:.1f}",
            color="#2563eb",
            fontsize=6.8,
            fontweight="bold",
            va="center",
        )
        ax.text(
            29.0,
            y_g3 + 0.3,
            inv_wrap,
            color="#475569",
            fontsize=6.2,
            va="center",
        )
        y_g3 -= step_g3

    # -------------------------------------------------------------
    # GRID 4: MAINTENANCE BREAKDOWN (Exact Match image_407125.png)
    # -------------------------------------------------------------
    hdr_box_g4 = patches.Rectangle(
        (52.0, 40.8), 45.0, 2.4, facecolor="#059669", edgecolor="none"
    )
    ax.add_patch(hdr_box_g4)
    x_g4 = [52.5, 59.0, 66.5, 74.5, 82.0, 89.5, 94.0]
    ax.text(
        x_g4[0],
        42.0,
        "Date",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g4[1],
        42.0,
        "Machine*",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g4[2],
        42.0,
        "Robot*",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g4[3],
        42.0,
        "Controller*",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g4[4],
        42.0,
        "RMCS*",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g4[5],
        42.0,
        "Oil/Water*",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )
    ax.text(
        x_g4[6],
        42.0,
        "Share",
        color="#ffffff",
        fontsize=6.8,
        fontweight="bold",
        va="center",
    )

    y_g4 = 38.5
    step_g4 = 4.8
    for idx, r in df_maint_grid.iterrows():
        bg_c = "#f8fafc" if idx % 2 == 1 else "#ffffff"
        ax.add_patch(
            patches.Rectangle(
                (52.0, y_g4 - 2.0),
                45.0,
                step_g4,
                facecolor=bg_c,
                edgecolor="none",
            )
        )
        ax.plot(
            [52.0, 97.0],
            [y_g4 - 2.0, y_g4 - 2.0],
            color="#e2e8f0",
            linewidth=0.4,
        )

        ax.text(
            x_g4[0],
            y_g4 + 0.3,
            str(r["Date"]),
            color="#0f172a",
            fontsize=6.8,
            fontweight="bold",
            va="center",
        )
        ax.text(
            x_g4[1],
            y_g4 + 0.3,
            f"{r['Machine Problem*']:.1f}",
            color="#2563eb",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            x_g4[2],
            y_g4 + 0.3,
            f"{r['Robot Problem*']:.1f}",
            color="#2563eb",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            x_g4[3],
            y_g4 + 0.3,
            f"{r['Controller Problem*']:.1f}",
            color="#2563eb",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            x_g4[4],
            y_g4 + 0.3,
            f"{r['RMCS Problem*']:.1f}",
            color="#2563eb",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            x_g4[5],
            y_g4 + 0.3,
            f"{r['Oil or water Leakage*']:.1f}",
            color="#2563eb",
            fontsize=6.8,
            va="center",
        )
        ax.text(
            x_g4[6],
            y_g4 + 0.3,
            str(r["Total Share"]),
            color="#0f172a",
            fontsize=6.8,
            fontweight="bold",
            va="center",
        )
        y_g4 -= step_g4

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


# =========================================================
# 4. STREAMLIT RENDER CONSOLE
# =========================================================
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
                '<p style="color:#64748b !important; font-size:0.85rem;">Upload the multi-month (or 2-month) ERP Downtime workbook.</p></div>',
                unsafe_allow_html=True,
            )
            up_dt = st.file_uploader(
                "Select Downtime Workbook (.xlsx)",
                type=["xlsx", "xls"],
                key="up_dt_file",
            )

        with c_up2:
            st.markdown(
                '<div style="background:#ffffff; padding:1.5rem; border-radius:12px; border:1px solid #e2e8f0; border-top:4px solid #10b981;">'
                '<h4 style="margin-top:0; color:#0f172a;">🛠️ 2. Service Maintenance Report (Optional)</h4>'
                '<p style="color:#64748b !important; font-size:0.85rem;">Upload engineering workshop ticket logs for SMS token audit.</p></div>',
                unsafe_allow_html=True,
            )
            up_sm = st.file_uploader(
                "Select Maintenance Ticket Workbook (.xlsx)",
                type=["xlsx", "xls"],
                key="up_sm_file",
            )

        if up_dt is not None:
            if st.button(
                "🚀 Ingest Workbooks & Launch Console",
                type="primary",
                use_container_width=True,
            ):
                st.session_state["m3_file_bytes"] = up_dt.getvalue()
                st.session_state["m3_sm_bytes"] = (
                    up_sm.getvalue() if up_sm is not None else None
                )
                st.rerun()

    else:
        df_downtime = m3_parse_downtime_workbook(st.session_state["m3_file_bytes"])
        df_service = m3_parse_service_maintenance(
            st.session_state.get("m3_sm_bytes")
        )

        all_months = sorted(df_downtime["YearMonth"].unique())
        active_month = all_months[-1]
        active_m_df = df_downtime[df_downtime["YearMonth"] == active_month]

        # Explicit Cutoff Selection (Requirement 1)
        avail_cutoff_dates = sorted(active_m_df["DateStr"].unique().tolist())

        st.markdown('<div class="control-bar-card">', unsafe_allow_html=True)
        c_date, c_causes, c_snap = st.columns([1.3, 2.2, 1.3], gap="small")

        with c_date:
            sel_cutoff_str = st.selectbox(
                "📅 **Select Last Operational Day (Cutoff)**",
                avail_cutoff_dates,
                index=len(avail_cutoff_dates) - 1,
            )

        sel_date_obj = pd.to_datetime(sel_cutoff_str)
        cutoff_day = sel_date_obj.day
        day_formatted = sel_date_obj.strftime("%d-%b")
        curr_month_name = active_m_df["MonthName"].iloc[0]

        # Customizable Top 10 Causes (Requirement 5)
        all_present_causes = sorted(
            [
                c
                for c in df_downtime["Cause"].dropna().unique()
                if "Server Error" not in str(c)
            ]
        )
        default_selected = [
            c for c in DEFAULT_TOP_10_CAUSES if c in all_present_causes
        ]

        with c_causes:
            selected_top_causes = st.multiselect(
                "🎯 **Top 10 NPT Reasons (Editable)**",
                all_present_causes,
                default=default_selected,
            )
            if not selected_top_causes:
                selected_top_causes = default_selected

        # Scoped DataFrames
        df_last_day = active_m_df[active_m_df["DateStr"] == sel_cutoff_str].copy()
        df_mtd = active_m_df[active_m_df["DayNum"] <= cutoff_day].copy()

        # Previous Month identification
        if len(all_months) >= 2:
            prev_month = all_months[-2]
            prev_m_df = df_downtime[df_downtime["YearMonth"] == prev_month]
            prev_month_name = prev_m_df["MonthName"].iloc[0]
            prev_m_mtd = prev_m_df[prev_m_df["DayNum"] <= cutoff_day]
        else:
            prev_month = active_month
            prev_m_df = active_m_df
            prev_month_name = "Prior"
            prev_m_mtd = df_mtd

        # Share Dictionaries for Grid 1 Chart
        tot_mtd_hrs = df_mtd["Hours"].sum()
        tot_prev_hrs = prev_m_mtd["Hours"].sum()
        curr_share_dict = (
            (df_mtd.groupby("Cause")["Hours"].sum() / tot_mtd_hrs * 100).to_dict()
            if tot_mtd_hrs > 0
            else {}
        )
        prev_share_dict = (
            (
                prev_m_mtd.groupby("Cause")["Hours"].sum()
                / tot_prev_hrs
                * 100
            ).to_dict()
            if tot_prev_hrs > 0
            else {}
        )

        # Grid 2, 3, 4 Data Computations
        df_size_grid, size_tot_hrs, size_summary_pct = m3_compute_size_wise_npt(
            df_mtd, cutoff_day
        )
        df_smed_grid = m3_compute_smed_table(df_mtd, max_days=7)
        df_maint_grid = m3_compute_maint_daily_table(df_mtd, max_days=7)

        # 4-Grid JPG Generation
        jpg_bytes = m3_generate_2x2_executive_jpg(
            sel_date_obj,
            cutoff_day,
            selected_top_causes,
            curr_share_dict,
            prev_share_dict,
            curr_month_name,
            prev_month_name,
            df_size_grid,
            size_tot_hrs,
            size_summary_pct,
            df_smed_grid,
            df_maint_grid,
        )

        with c_snap:
            st.markdown(
                "<div style='margin-top: 1.65rem;'></div>",
                unsafe_allow_html=True,
            )
            st.download_button(
                label="📸 Download 2×2 JPG Report",
                data=jpg_bytes,
                file_name=f"NPT_4Grid_Report_{sel_cutoff_str}.jpg",
                mime="image/jpeg",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        # 4 Core KPI Cards
        k1, k2, k3, k4 = st.columns(4)
        k1.markdown(
            f'<div class="kpi-card blue"><div class="kpi-title">MTD TOTAL NPT (Day 1–{cutoff_day})</div><div class="kpi-val">{tot_mtd_hrs:,.1f} H</div><div class="kpi-sub">Pace: {tot_mtd_hrs/cutoff_day:.1f} H/Day</div></div>',
            unsafe_allow_html=True,
        )
        k2.markdown(
            f'<div class="kpi-card purple"><div class="kpi-title">PLANT CAPACITY NPT %</div><div class="kpi-val">{size_summary_pct:.2f}%</div><div class="kpi-sub">Of {TOTAL_PLANT_MCS*24*cutoff_day:,.0f} H Total Available</div></div>',
            unsafe_allow_html=True,
        )
        last_day_hrs = df_last_day["Hours"].sum()
        k3.markdown(
            f'<div class="kpi-card pink"><div class="kpi-title">LAST DAY NPT ({day_formatted})</div><div class="kpi-val">{last_day_hrs:.1f} H</div><div class="kpi-sub">{(last_day_hrs/DAILY_AVAILABLE_HRS*100):.1f}% Day Capacity</div></div>',
            unsafe_allow_html=True,
        )
        maint_last_hrs = df_last_day[df_last_day["Is_Maintenance"]][
            "Hours"
        ].sum()
        k4.markdown(
            f'<div class="kpi-card yellow"><div class="kpi-title">LAST DAY MAINT. IMPACT</div><div class="kpi-val">{maint_last_hrs:.1f} H</div><div class="kpi-sub">{(maint_last_hrs/last_day_hrs*100 if last_day_hrs>0 else 0):.1f}% NPT Share</div></div>',
            unsafe_allow_html=True,
        )

        st.markdown(
            "<div style='margin-bottom: 1.25rem;'></div>",
            unsafe_allow_html=True,
        )

        # Consolidated Machine Table & WhatsApp Brief (Requirements 3 & 4)
        col_left, col_right = st.columns([1.5, 1.1], gap="large")

        with col_left:
            st.markdown(
                f"#### ⚙️ MACHINE NPT INCIDENTS LOG — {sel_date_obj.strftime('%B %d')}"
            )
            st.caption(
                "Consolidated: Single combined entry per machine for the day."
            )
            df_cons_log = m3_compute_consolidated_daily_log(df_last_day)
            if not df_cons_log.empty:
                st.dataframe(
                    df_cons_log,
                    use_container_width=True,
                    hide_index=True,
                    height=400,
                )
            else:
                st.success("✅ Zero downtime logged for this date!")

        with col_right:
            # Build 3-Part WhatsApp Brief exactly to spec
            top_mtd_causes = (
                df_mtd.groupby("Cause")["Hours"]
                .sum()
                .sort_values(ascending=False)
                .head(4)
            )
            top_causes_lines = []
            for c_n, c_h in top_mtd_causes.items():
                c_clean = c_n.replace("*", "").strip()
                c_pct = (c_h / tot_mtd_hrs * 100) if tot_mtd_hrs > 0 else 0.0
                top_causes_lines.append(
                    f"{c_clean}: {c_h:,.2f} Hrs ({c_pct:.2f}% share)"
                )

            # Maintenance values on last day
            maint_items = [
                "Machine Problem*",
                "Robot Problem*",
                "Controller Problem*",
                "RMCS Problem*",
                "Oil or water Leakage*",
            ]
            maint_lines = []
            tot_tech_day = 0.0
            for mi in maint_items:
                mi_clean = mi.replace("*", "").strip()
                mi_h = df_last_day[df_last_day["Cause"] == mi]["Hours"].sum()
                tot_tech_day += mi_h
                maint_lines.append(f"{mi_clean}: {mi_h:.2f} Hrs")

            # SMED on last day
            smed_last_df = df_last_day[df_last_day["Is_SMED"]]
            smed_setups = len(smed_last_df)
            smed_hrs = smed_last_df["Hours"].sum()
            smed_avg_min = (
                (smed_hrs / smed_setups * 60.0) if smed_setups > 0 else 0.0
            )
            smed_mcs_str = (
                ", ".join(
                    sorted(
                        set(
                            str(v)
                            for v in smed_last_df["Position"].unique()
                            if v != "-"
                        )
                    )
                )
                if smed_setups > 0
                else "None"
            )

            # MTD SMED
            smed_mtd_df = df_mtd[df_mtd["Is_SMED"]]
            mtd_smed_setups = len(smed_mtd_df)
            mtd_smed_hrs = smed_mtd_df["Hours"].sum()
            mtd_smed_avg = (
                (mtd_smed_hrs / mtd_smed_setups * 60.0)
                if mtd_smed_setups > 0
                else 0.0
            )

            # Prior months comparison string
            comp_str_list = []
            for p_m in all_months[:-1]:
                p_df = df_downtime[df_downtime["YearMonth"] == p_m]
                p_name = p_df["MonthName"].iloc[0]
                p_hrs = p_df["Hours"].sum()
                comp_str_list.append(f"{p_hrs:,.2f} Hrs in {p_name}")
            comp_str = " and ".join(comp_str_list) if comp_str_list else "-"

            whatsapp_msg = f"""📅 *Date:* {sel_date_obj.strftime('%d-%m-%Y')}

Dear Sir,

*1. Overall Monthly NPT Analysis (MTD Comparison)*
Current Month Total NPT: *{tot_mtd_hrs:,.2f} Hours* (vs. {comp_str}).

*Top Contributing Causes ({curr_month_name}):*
{chr(10).join(top_causes_lines)}

*2. Machine Maintenance & Technical NPT (Last Day: {day_formatted})*
{chr(10).join(maint_lines)}
*Total Maintenance Impact:* ~{tot_tech_day:.2f} Hrs ({(tot_tech_day/last_day_hrs*100 if last_day_hrs>0 else 0):.1f}% day NPT share)

*3. Mold Change & SMED Performance (Last Day: {day_formatted})*
• Mold Changes Completed: *{smed_setups} setups*
• Total Setup Time: *{smed_hrs:.2f} Hours*
• Average SMED: *{smed_avg_min:.2f} Min/change*
• Involved Machines: {smed_mcs_str}
• Month-to-Date SMED: *{mtd_smed_setups} setups* completed totaling *{mtd_smed_hrs:.2f} Hours* (MTD Avg: *{mtd_smed_avg:.2f} Min*)"""

            st.markdown("#### 📝 EXECUTIVE BRIEFING TEXT")
            st.markdown(
                f"""<div class="narrative-block">
                    <p style="margin:0 0 0.5rem 0; font-weight:800; color:#1e293b;">📋 PLASTIC-3 DAILY NPT & DOWNTIME BRIEF</p>
                    <p style="margin:0 0 0.75rem 0; color:#64748b; font-size:0.82rem;">📅 <b>Date:</b> {sel_date_obj.strftime('%d-%m-%Y')}</p>
                    <h5>1. Overall Monthly NPT Analysis</h5>
                    <p>Current Month Total NPT: <b>{tot_mtd_hrs:,.2f} Hours</b> (vs. {comp_str}).</p>
                    <p style="margin:0.25rem 0 0.5rem 0;">{'<br>'.join(top_causes_lines)}</p>
                    <h5>2. Machine Maintenance & Technical NPT</h5>
                    <p style="margin:0.25rem 0 0.5rem 0;">{'<br>'.join(maint_lines)}<br><b>Total Impact:</b> ~{tot_tech_day:.2f} Hrs</p>
                    <h5>3. Mold Change & SMED Performance</h5>
                    <p>• Changes: <b>{smed_setups} setups</b> ({smed_hrs:.2f} Hrs | Avg: <b>{smed_avg_min:.2f} Min</b>)<br>
                    • Machines: {smed_mcs_str}<br>
                    • MTD: <b>{mtd_smed_setups} setups</b> ({mtd_smed_hrs:.2f} Hrs | Avg: <b>{mtd_smed_avg:.2f} Min</b>)</p>
                </div>""",
                unsafe_allow_html=True,
            )

            with st.expander("📋 Copy Plain Text Brief"):
                st.text_area(
                    "Brief Text",
                    value=whatsapp_msg,
                    height=200,
                    label_visibility="collapsed",
                )

        st.divider()

        # Detailed Analysis Tabs
        tab_grid2, tab_smed, tab_maint = st.tabs([
            "📏 MC Size-Wise Capacity Loss",
            "⏱️ SMED (Mold Changeover)",
            "🛠️ Maintenance & SMS Audit",
        ])

        with tab_grid2:
            st.markdown(
                f"#### 📏 MC SIZE-WISE NPT BREAKDOWN (Day 1–{cutoff_day})"
            )
            st.dataframe(df_size_grid, use_container_width=True, hide_index=True)

        with tab_smed:
            st.markdown(
                f"#### ⏱️ SMED CHANGEOVER PERFORMANCE (Day 1–{cutoff_day})"
            )
            st.dataframe(df_smed_grid, use_container_width=True, hide_index=True)

        with tab_maint:
            st.markdown(
                f"#### 🛠️ TECHNICAL & BREAKDOWN TREND (Day 1–{cutoff_day})"
            )
            st.dataframe(
                df_maint_grid, use_container_width=True, hide_index=True
            )
