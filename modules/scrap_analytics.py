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
            causes_list = [str(c).strip().replace("*", "") for c in grp[cause_col].dropna().unique() if str(c).strip()]
            causes_str = ", ".join(causes_list) if causes_list else "No Rejection"
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


def m2_generate_scrap_jpg(
    df_day_filtered,
    sel_date_obj,
    total_rej_pcs,
    total_rej_ton,
    prev_total_ton,
    prev_avg_ton,
    curr_as_of_total_ton,
    curr_as_of_avg_ton,
    high_rej_count,
    total_day_mcs,
    top_cause,
    top_cause_pcs,
    top_cause_pct,
    top_wt_mc,
    top_wt_kg,
    top3_pct,
    gf_share_pct,
    ff_share_pct,
):
    """
    Generate a professional one-page JPG management report.

    Design:
    - Compact title area
    - 6 KPI cards
    - ONE full-width rejection table
    - No Line column
    - Causes are wrapped, never truncated
    - Dynamic row heights
    - Dynamic image height based on number of machines
    - Executive summary/action cards below the table
    """

    import math
    import textwrap

    # ---------------------------------------------------------
    # BASIC PREPARATION
    # ---------------------------------------------------------
    df_report = df_day_filtered.copy()

    # Ensure stable ordering by rejection quantity.
    if not df_report.empty and "Qty (Pcs)" in df_report.columns:
        df_report = df_report.sort_values(
            "Qty (Pcs)",
            ascending=False
        ).reset_index(drop=True)

    n_rows = len(df_report)

    date_formatted = sel_date_obj.strftime("%B %d, %Y")
    day_formatted = sel_date_obj.strftime("%B %d")

    # ---------------------------------------------------------
    # WRAPPING SETTINGS
    # ---------------------------------------------------------
    # Causes receive the majority of the table width.
    #
    # The value is deliberately based on characters rather than
    # pixels so that the same logic works for different report
    # sizes.
    CAUSE_WRAP = 52

    def wrap_cause(value):
        """
        Wrap cause text without cutting it.
        Removes '*' because the report is intended for management.
        """
        if value is None:
            return "-"

        text = str(value).strip()

        if not text:
            return "-"

        text = text.replace("*", "")

        # Clean repeated spaces.
        text = " ".join(text.split())

        return "\n".join(
            textwrap.wrap(
                text,
                width=CAUSE_WRAP,
                break_long_words=False,
                break_on_hyphens=False,
            )
        )

    # ---------------------------------------------------------
    # DETERMINE ROW HEIGHT
    # ---------------------------------------------------------
    # Most rows will be one line.
    # Long cause descriptions become two or three lines.
    row_line_counts = []

    if not df_report.empty:
        for _, row in df_report.iterrows():
            wrapped = wrap_cause(row.get("Causes", "-"))
            line_count = max(1, wrapped.count("\n") + 1)
            row_line_counts.append(line_count)

    # Minimum row height.
    BASE_ROW_HEIGHT = 0.62

    # Additional height for wrapped cause lines.
    EXTRA_LINE_HEIGHT = 0.30

    row_heights = [
        BASE_ROW_HEIGHT + max(0, lines - 1) * EXTRA_LINE_HEIGHT
        for lines in row_line_counts
    ]

    # If there are no rows, still reserve some table space.
    if not row_heights:
        row_heights = [0.70]

    table_height = sum(row_heights) + 0.72

    # ---------------------------------------------------------
    # EXECUTIVE SUMMARY AREA
    # ---------------------------------------------------------
    summary_height = 4.15

    # ---------------------------------------------------------
    # TOTAL REPORT HEIGHT
    # ---------------------------------------------------------
    #
    # Top title       ~1.15
    # KPI row         ~1.45
    # spacing         ~0.25
    # table           dynamic
    # summary        ~4.15
    #
    # This makes the image taller when many machines qualify
    # instead of shrinking the table text.
    # ---------------------------------------------------------
    top_area = 3.0

    fig_height = max(
        11.5,
        top_area + table_height + summary_height
    )

    fig, ax = plt.subplots(
        figsize=(18, fig_height),
        dpi=220
    )

    fig.patch.set_facecolor("#f4f7fb")
    ax.set_facecolor("#f4f7fb")

    ax.set_xlim(0, 100)
    ax.set_ylim(0, fig_height)
    ax.axis("off")

    # ---------------------------------------------------------
    # COORDINATE SYSTEM
    # ---------------------------------------------------------
    y_top = fig_height - 0.55

    # ---------------------------------------------------------
    # COMPACT REPORT TITLE
    # ---------------------------------------------------------
    ax.text(
        2.0,
        y_top,
        "DAILY SCRAP & DEFECT ANALYTICS REPORT",
        color="#111827",
        fontsize=17,
        fontweight="bold",
        va="top",
    )

    ax.text(
        2.0,
        y_top - 0.48,
        f"Plastic-3  |  Report Date: {date_formatted}  |  "
        f"Rejection Threshold: >{50} Pcs",
        color="#64748b",
        fontsize=8.5,
        va="top",
    )

    # Small report identifier on right.
    ax.text(
        98.0,
        y_top - 0.08,
        "OPERATIONAL QUALITY & DEFECT CONTROL",
        color="#6366f1",
        fontsize=7.2,
        fontweight="bold",
        ha="right",
        va="top",
    )

    # ---------------------------------------------------------
    # KPI CARDS
    # ---------------------------------------------------------
    kpis = [
        (
            "PREV MO. TOTAL",
            f"{prev_total_ton:.2f} T",
            "Total Rejection",
            "#64748b",
        ),
        (
            "PREV MO. AVG",
            f"{prev_avg_ton:.2f} T/Day",
            "Daily Baseline",
            "#64748b",
        ),
        (
            "THIS MO. AS OF",
            f"{curr_as_of_total_ton:.2f} T",
            f"As of Day {sel_date_obj.day}",
            "#2563eb",
        ),
        (
            "THIS MO. AVG",
            f"{curr_as_of_avg_ton:.2f} T/Day",
            "Current Pace",
            "#2563eb",
        ),
        (
            "LAST DAY SCRAP",
            f"{total_rej_ton:.3f} T",
            f"{total_rej_pcs:,} Pcs",
            "#dc2626",
        ),
        (
            f"CRITICAL MC (>{50})",
            f"{high_rej_count}",
            "Lines Exceeding Limit",
            "#8b5cf6",
        ),
    ]

    kpi_y = y_top - 1.35
    kpi_h = 1.12
    gap = 0.65
    left_margin = 2.0
    total_width = 96.0

    kpi_w = (
        total_width - gap * (len(kpis) - 1)
    ) / len(kpis)

    for i, (title, value, subtitle, accent) in enumerate(kpis):

        x = left_margin + i * (kpi_w + gap)

        card = patches.FancyBboxPatch(
            (x, kpi_y - kpi_h),
            kpi_w,
            kpi_h,
            boxstyle="round,pad=0.10,rounding_size=0.18",
            facecolor="#ffffff",
            edgecolor="#dbe2ea",
            linewidth=0.8,
        )

        ax.add_patch(card)

        # Accent line.
        ax.add_patch(
            patches.FancyBboxPatch(
                (x + 0.12, kpi_y - 0.13),
                kpi_w - 0.24,
                0.09,
                boxstyle="round,pad=0.01,rounding_size=0.04",
                facecolor=accent,
                edgecolor="none",
            )
        )

        ax.text(
            x + kpi_w / 2,
            kpi_y - 0.32,
            title,
            color="#64748b",
            fontsize=6.7,
            fontweight="bold",
            ha="center",
            va="center",
        )

        ax.text(
            x + kpi_w / 2,
            kpi_y - 0.62,
            value,
            color="#111827",
            fontsize=12.2,
            fontweight="bold",
            ha="center",
            va="center",
        )

        ax.text(
            x + kpi_w / 2,
            kpi_y - 0.91,
            subtitle,
            color="#94a3b8",
            fontsize=6.2,
            ha="center",
            va="center",
        )

    # ---------------------------------------------------------
    # TABLE POSITION
    # ---------------------------------------------------------
    table_top = kpi_y - kpi_h - 0.55

    table_left = 2.0
    table_width = 96.0

    # Header height.
    header_h = 0.72

    table_bottom = table_top - table_height

    # Outer table container.
    table_card = patches.FancyBboxPatch(
        (table_left, table_bottom),
        table_width,
        table_height,
        boxstyle="round,pad=0.12,rounding_size=0.20",
        facecolor="#ffffff",
        edgecolor="#dbe2ea",
        linewidth=0.8,
    )

    ax.add_patch(table_card)

    # ---------------------------------------------------------
    # TABLE TITLE
    # ---------------------------------------------------------
    ax.text(
        table_left + 1.0,
        table_top - 0.42,
        f"PLASTIC-3 MACHINE REJECTION LOG (>50 Pcs) — {day_formatted}",
        color="#111827",
        fontsize=9.8,
        fontweight="bold",
        va="center",
    )

    # ---------------------------------------------------------
    # TABLE HEADER
    # ---------------------------------------------------------
    header_y = table_top - 0.82

    header = patches.FancyBboxPatch(
        (table_left + 0.8, header_y - header_h),
        table_width - 1.6,
        header_h,
        boxstyle="round,pad=0.02,rounding_size=0.06",
        facecolor="#0f172a",
        edgecolor="none",
    )

    ax.add_patch(header)

    # ---------------------------------------------------------
    # COLUMN DEFINITIONS
    # ---------------------------------------------------------
    #
    # No Line column.
    #
    # Causes gets the largest share of the width.
    #
    # MC Position     13%
    # Smart Manu      18%
    # Causes          49%
    # Qty             10%
    # Weight          10%
    # ---------------------------------------------------------

    table_x = table_left + 0.8
    table_w = table_width - 1.6

    col_widths = [
        0.13,
        0.18,
        0.49,
        0.10,
        0.10,
    ]

    col_x = [table_x]

    for width in col_widths[:-1]:
        col_x.append(
            col_x[-1] + table_w * width
        )

    headers = [
        "MC Position",
        "Smart Manu",
        "Causes",
        "Qty (Pcs)",
        "Weight (kg)",
    ]

    for i, title in enumerate(headers):

        x0 = col_x[i]
        x1 = (
            col_x[i + 1]
            if i < len(col_x) - 1
            else table_x + table_w
        )

        ax.text(
            (x0 + x1) / 2,
            header_y - header_h / 2,
            title,
            color="#ffffff",
            fontsize=6.8,
            fontweight="bold",
            ha="center",
            va="center",
        )

    # ---------------------------------------------------------
    # TABLE ROWS
    # ---------------------------------------------------------
    if df_report.empty:

        ax.text(
            table_left + table_width / 2,
            header_y - header_h - 0.65,
            "No machines exceeded the rejection threshold.",
            color="#64748b",
            fontsize=8.5,
            ha="center",
            va="center",
        )

    else:

        current_y = header_y - header_h

        for row_idx, (_, row) in enumerate(df_report.iterrows()):

            row_h = row_heights[row_idx]

            row_top = current_y
            row_bottom = current_y - row_h

            # Alternating row background.
            row_bg = (
                "#ffffff"
                if row_idx % 2 == 0
                else "#f8fafc"
            )

            ax.add_patch(
                patches.Rectangle(
                    (
                        table_x,
                        row_bottom,
                    ),
                    table_w,
                    row_h,
                    facecolor=row_bg,
                    edgecolor="none",
                )
            )

            # Bottom separator.
            ax.plot(
                [table_x, table_x + table_w],
                [row_bottom, row_bottom],
                color="#e2e8f0",
                linewidth=0.45,
            )

            # -------------------------------------------------
            # VALUES
            # -------------------------------------------------
            mc_position = str(
                row.get("MC Position", "-")
            )

            smart_manu = str(
                row.get("Smart Manu", "-")
            )

            causes = wrap_cause(
                row.get("Causes", "-")
            )

            qty = row.get("Qty (Pcs)", 0)

            weight = row.get("Weight (kg)", 0)

            try:
                qty_text = f"{int(round(float(qty))):,}"
            except Exception:
                qty_text = str(qty)

            try:
                weight_text = f"{float(weight):,.1f}"
            except Exception:
                weight_text = str(weight)

            # -------------------------------------------------
            # VERTICAL CENTERING
            # -------------------------------------------------
            center_y = (
                row_top + row_bottom
            ) / 2

            # -------------------------------------------------
            # MC POSITION
            # -------------------------------------------------
            x0 = col_x[0]
            x1 = col_x[1]

            ax.text(
                (x0 + x1) / 2,
                center_y,
                mc_position,
                color="#111827",
                fontsize=6.5,
                fontweight="bold",
                ha="center",
                va="center",
            )

            # -------------------------------------------------
            # SMART MANU
            # -------------------------------------------------
            x0 = col_x[1]
            x1 = col_x[2]

            ax.text(
                (x0 + x1) / 2,
                center_y,
                smart_manu,
                color="#64748b",
                fontsize=6.3,
                ha="center",
                va="center",
            )

            # -------------------------------------------------
            # CAUSES
            # -------------------------------------------------
            x0 = col_x[2]
            x1 = col_x[3]

            ax.text(
                x0 + 0.55,
                center_y,
                causes,
                color="#dc4b4b",
                fontsize=6.3,
                ha="left",
                va="center",
                linespacing=1.15,
            )

            # -------------------------------------------------
            # QUANTITY
            # -------------------------------------------------
            x0 = col_x[3]
            x1 = col_x[4]

            ax.text(
                (x0 + x1) / 2,
                center_y,
                qty_text,
                color="#111827",
                fontsize=6.5,
                fontweight="bold",
                ha="center",
                va="center",
            )

            # -------------------------------------------------
            # WEIGHT
            # -------------------------------------------------
            x0 = col_x[4]
            x1 = table_x + table_w

            ax.text(
                (x0 + x1) / 2,
                center_y,
                weight_text,
                color="#111827",
                fontsize=6.3,
                ha="center",
                va="center",
            )

            current_y = row_bottom

    # ---------------------------------------------------------
    # EXECUTIVE SUMMARY SECTION
    # ---------------------------------------------------------
    summary_top = table_bottom - 0.55

    # Section title.
    ax.text(
        2.0,
        summary_top,
        "EXECUTIVE SUMMARY & ACTION ANALYSIS",
        color="#111827",
        fontsize=9.8,
        fontweight="bold",
        va="top",
    )

    box_gap = 1.0
    box_y = summary_top - 0.48
    box_h = 3.15

    box_w = (
        (96.0 - 2 * box_gap) / 3
    )

    # ---------------------------------------------------------
    # SUMMARY BOX 1
    # ---------------------------------------------------------
    x1 = 2.0

    box1 = patches.FancyBboxPatch(
        (x1, box_y - box_h),
        box_w,
        box_h,
        boxstyle="round,pad=0.14,rounding_size=0.18",
        facecolor="#f8fafc",
        edgecolor="#dbe2ea",
        linewidth=0.8,
    )

    ax.add_patch(box1)

    ax.text(
        x1 + 0.8,
        box_y - 0.48,
        "EXECUTIVE SUMMARY",
        color="#111827",
        fontsize=7.3,
        fontweight="bold",
        va="center",
    )

    summary_text = (
        f"• {high_rej_count} machines in Plastic-3\n"
        f"  exceeded 50 pcs.\n"
        f"• Last Day: {total_rej_pcs:,} pcs "
        f"({total_rej_ton:.3f} T).\n"
        f"• Previous Month: {prev_total_ton:.2f} T "
        f"({prev_avg_ton:.2f} T/Day).\n"
        f"• Current Month as of {day_formatted}: "
        f"{curr_as_of_total_ton:.2f} T "
        f"({curr_as_of_avg_ton:.2f} T/Day)."
    )

    ax.text(
        x1 + 0.8,
        box_y - 0.95,
        summary_text,
        color="#334155",
        fontsize=6.5,
        linespacing=1.45,
        va="top",
    )

    # ---------------------------------------------------------
    # SUMMARY BOX 2
    # ---------------------------------------------------------
    x2 = x1 + box_w + box_gap

    box2 = patches.FancyBboxPatch(
        (x2, box_y - box_h),
        box_w,
        box_h,
        boxstyle="round,pad=0.14,rounding_size=0.18",
        facecolor="#fef2f2",
        edgecolor="#fecaca",
        linewidth=0.8,
    )

    ax.add_patch(box2)

    ax.text(
        x2 + 0.8,
        box_y - 0.48,
        "TOP DEFECT DRIVER & HEAVY LINE",
        color="#b91c1c",
        fontsize=7.3,
        fontweight="bold",
        va="center",
    )

    driver_text = (
        f"• Primary Scrap Cause: "
        f"'{str(top_cause).replace('*', '')}'\n"
        f"  generating {top_cause_pcs:,} pcs "
        f"({top_cause_pct:.1f}% share).\n"
        f"• Heaviest Scrap Machine: "
        f"{top_wt_mc}\n"
        f"  generating {top_wt_kg:.1f} kg scrap.\n"
        f"• Top 3 Causes account for "
        f"{top3_pct:.1f}% of rejected pieces."
    )

    ax.text(
        x2 + 0.8,
        box_y - 0.95,
        driver_text,
        color="#7f1d1d",
        fontsize=6.5,
        linespacing=1.45,
        va="top",
    )

    # ---------------------------------------------------------
    # SUMMARY BOX 3
    # ---------------------------------------------------------
    x3 = x2 + box_w + box_gap

    box3 = patches.FancyBboxPatch(
        (x3, box_y - box_h),
        box_w,
        box_h,
        boxstyle="round,pad=0.14,rounding_size=0.18",
        facecolor="#eff6ff",
        edgecolor="#bfdbfe",
        linewidth=0.8,
    )

    ax.add_patch(box3)

    ax.text(
        x3 + 0.8,
        box_y - 0.48,
        "LAST DAY PLANT OVERVIEW",
        color="#1d4ed8",
        fontsize=7.3,
        fontweight="bold",
        va="center",
    )

    overview_text = (
        f"• Total Logged Lines: {total_day_mcs} machines\n"
        f"  ({high_rej_count} >50 pcs; "
        f"{max(total_day_mcs - high_rej_count, 0)} ≤50 pcs).\n"
        f"• Overall Factory Scrap: "
        f"{total_rej_pcs:,} pcs / {total_rej_ton:.3f} T.\n"
        f"• Top 3 Causes: {top3_pct:.1f}% "
        f"of lost pieces.\n"
        f"• Floor Split: GF {gf_share_pct:.1f}% "
        f"vs FF {ff_share_pct:.1f}% scrap weight."
    )

    ax.text(
        x3 + 0.8,
        box_y - 0.95,
        overview_text,
        color="#1e3a8a",
        fontsize=6.5,
        linespacing=1.45,
        va="top",
    )

    # ---------------------------------------------------------
    # FINAL LAYOUT
    # ---------------------------------------------------------
    plt.subplots_adjust(
        left=0,
        right=1,
        top=1,
        bottom=0,
    )

    buf = io.BytesIO()

    plt.savefig(
        buf,
        format="jpg",
        facecolor=fig.get_facecolor(),
        edgecolor="none",
        dpi=220,
        bbox_inches="tight",
        pad_inches=0.08,
    )

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
