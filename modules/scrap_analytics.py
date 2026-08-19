import io
import re
import textwrap
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


def clean_mold_name(val):
    if not val or pd.isna(val):
        return "-"
    text = str(val).strip()
    text = re.sub(r"\(.*?\)", "", text).strip()
    return text if text else "-"


def get_col(df, candidates, default=None):
    for c in candidates:
        if c in df.columns:
            return c
    return default


@st.cache_data
def m2_parse_workbook(file_bytes):
    file_stream = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(file_stream)

    sheet_name = (
        "RejectionReport"
        if "RejectionReport" in xls.sheet_names
        else ("This Month" if "This Month" in xls.sheet_names else xls.sheet_names[0])
    )
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values])
        if "Machine" in row_str and (
            "Quantity" in row_str or "Qty" in row_str or "Cause" in row_str
        ):
            header_idx = idx
            break

    if header_idx is not None:
        df_clean = pd.read_excel(xls, sheet_name=sheet_name, skiprows=header_idx)
    else:
        df_clean = pd.read_excel(xls, sheet_name=sheet_name)

    df_clean.columns = [str(c).strip() for c in df_clean.columns]

    date_col = get_col(
        df_clean, ["Added Date", "Date", "Entry Date", "AddedDate"], df_clean.columns[-1]
    )
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
        raw_qty = (
            pd.to_numeric(grp[qty_col], errors="coerce").fillna(0).sum()
            if qty_col in grp.columns
            else 0.0
        )
        qty_factor = 1000.0 if (qty_col in grp.columns and grp[qty_col].max() < 100) else 1.0
        total_pcs = raw_qty * qty_factor
        total_ton = (
            pd.to_numeric(grp[wt_col], errors="coerce").fillna(0).sum()
            if wt_col in grp.columns
            else 0.0
        )

        if cause_col in grp.columns:
            causes_list = [
                str(c).strip().replace("*", "")
                for c in grp[cause_col].dropna().unique()
                if str(c).strip()
            ]
            causes_str = ", ".join(causes_list) if causes_list else "No Rejection"
        else:
            causes_str = "-"

        raw_mold = (
            str(grp[item_col].iloc[0])
            if (item_col in grp.columns and not grp[item_col].dropna().empty)
            else "-"
        )
        mold_name = clean_mold_name(raw_mold)
        pos = POS_MAP.get(str(mc), "-")
        line = LINE_MAP.get(str(mc), "-")

        if total_pcs >= min_qty:
            records.append({
                "Position": pos,
                "Line": line,
                "Machine": str(mc),
                "Causes": causes_str,
                "Qty": int(round(total_pcs)),
                "Weight (Ton)": round(total_ton, 4),
                "Weight (kg)": round(total_ton * 1000.0, 2),
                "Mold": mold_name,
            })

    df_res = pd.DataFrame(records)
    if not df_res.empty:
        df_res = df_res.sort_values("Qty", ascending=False).reset_index(drop=True)
    return df_res


def m2_compute_cause_breakdown(df_scope):
    if df_scope.empty:
        return pd.DataFrame()
    qty_col = get_col(df_scope, ["Quantity", "Qty", "Rejection Pcs", "Qty (Pcs)"], None)
    wt_col = get_col(df_scope, ["Weight", "Rejection Ton", "Weight (Ton)"], None)
    cause_col = get_col(df_scope, ["Cause", "Causes", "Defect", "Reason"], "Cause")
    mc_col = get_col(df_scope, ["Machine", "MC SL"], "Machine")

    qty_factor = 1000.0 if (qty_col and df_scope[qty_col].max() < 100) else 1.0

    res = (
        df_scope.groupby(cause_col)
        .agg(
            Rej_Pcs=(
                qty_col,
                lambda x: int(round(pd.to_numeric(x, errors="coerce").fillna(0).sum() * qty_factor)),
            )
            if qty_col
            else (cause_col, "count"),
            Rej_Kg=(
                wt_col,
                lambda x: round(pd.to_numeric(x, errors="coerce").fillna(0).sum() * 1000.0, 1),
            )
            if wt_col
            else (cause_col, "count"),
            Rej_Ton=(
                wt_col,
                lambda x: round(pd.to_numeric(x, errors="coerce").fillna(0).sum(), 4),
            )
            if wt_col
            else (cause_col, "count"),
            Entries_Count=(cause_col, "count"),
            MC_Count=(mc_col, "nunique") if mc_col in df_scope.columns else (cause_col, "count"),
        )
        .reset_index()
    )

    res["Cause"] = res[cause_col].astype(str).str.replace("*", "", regex=False).str.strip()
    tot_pcs = res["Rej_Pcs"].sum()
    res["% Share"] = (res["Rej_Pcs"] / tot_pcs * 100.0).round(1) if tot_pcs > 0 else 0.0
    res = res.sort_values("Rej_Pcs", ascending=False).reset_index(drop=True)
    return res[["Cause", "Rej_Pcs", "Rej_Kg", "Rej_Ton", "Entries_Count", "MC_Count", "% Share"]]


def m2_compute_lineman_breakdown(df_scope):
    if df_scope.empty:
        return pd.DataFrame()
    lineman_col = get_col(df_scope, ["Added By", "AddedBy", "Lineman", "Operator", "Added_By"], None)
    if not lineman_col or lineman_col not in df_scope.columns:
        return pd.DataFrame()

    qty_col = get_col(df_scope, ["Quantity", "Qty", "Rejection Pcs", "Qty (Pcs)"], None)
    wt_col = get_col(df_scope, ["Weight", "Rejection Ton", "Weight (Ton)"], None)
    mc_col = get_col(df_scope, ["Machine", "MC SL"], "Machine")

    qty_factor = 1000.0 if (qty_col and df_scope[qty_col].max() < 100) else 1.0

    res = (
        df_scope.groupby(lineman_col)
        .agg(
            Rej_Pcs=(
                qty_col,
                lambda x: int(round(pd.to_numeric(x, errors="coerce").fillna(0).sum() * qty_factor)),
            )
            if qty_col
            else (lineman_col, "count"),
            Rej_Kg=(
                wt_col,
                lambda x: round(pd.to_numeric(x, errors="coerce").fillna(0).sum() * 1000.0, 1),
            )
            if wt_col
            else (lineman_col, "count"),
            Rej_Ton=(
                wt_col,
                lambda x: round(pd.to_numeric(x, errors="coerce").fillna(0).sum(), 4),
            )
            if wt_col
            else (lineman_col, "count"),
            Logged_Entries=(lineman_col, "count"),
            Machines_Covered=(mc_col, "nunique") if mc_col in df_scope.columns else (lineman_col, "count"),
        )
        .reset_index()
    )

    tot_pcs = res["Rej_Pcs"].sum()
    tot_ton = res["Rej_Ton"].sum()
    res["% Pcs Share"] = (res["Rej_Pcs"] / tot_pcs * 100.0).round(1) if tot_pcs > 0 else 0.0
    res["% Ton Share"] = (res["Rej_Ton"] / tot_ton * 100.0).round(1) if tot_ton > 0 else 0.0
    res = res.sort_values("Rej_Pcs", ascending=False).reset_index(drop=True)
    res = res.rename(columns={lineman_col: "Lineman (Added By)"})
    return res


def m2_compute_tonnage_comparison(df_prev, df_curr):
    wt_col_prev = (
        get_col(df_prev, ["Weight", "Rejection Ton"], None) if not df_prev.empty else None
    )
    wt_col_curr = (
        get_col(df_curr, ["Weight", "Rejection Ton"], None) if not df_curr.empty else None
    )

    if not df_prev.empty and wt_col_prev:
        t_prev = (
            df_prev.groupby(df_prev["DateClean"].dt.day)[wt_col_prev].sum().reset_index()
        )
        t_prev.columns = ["Day", "Prev_Month_Ton"]
    else:
        t_prev = pd.DataFrame(columns=["Day", "Prev_Month_Ton"])

    if not df_curr.empty and wt_col_curr:
        t_curr = (
            df_curr.groupby(df_curr["DateClean"].dt.day)[wt_col_curr].sum().reset_index()
        )
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
    top3_summary_list,
):
    fig, ax = plt.subplots(figsize=(18, 10.5), dpi=220)
    fig.patch.set_facecolor("#f1f5f9")
    ax.set_facecolor("#f1f5f9")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis("off")

    date_formatted = sel_date_obj.strftime("%B %d, %Y")
    day_formatted = sel_date_obj.strftime("%B %d")

    # 1. Header
    ax.text(
        1.5,
        98.4,
        "DAILY REJECTION & DEFECT ANALYTICS REPORT",
        color="#0f172a",
        fontsize=16.0,
        fontweight="bold",
        va="top",
    )
    ax.text(
        1.5,
        95.8,
        f"Plastic-3 Machine Rejection Log (>50 Pcs) & Plant Summary  |  Report Date: {date_formatted}",
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

    # 2. KPI Cards Row
    kpis = [
        ("PREV MO. TOTAL", f"{prev_total_ton:.2f} T", "Total Rejection", "#64748b"),
        ("PREV MO. AVG", f"{prev_avg_ton:.2f} T/Day", "Daily Baseline", "#64748b"),
        ("THIS MO. AS OF", f"{curr_as_of_total_ton:.2f} T", f"As of {day_formatted}", "#2563eb"),
        ("THIS MO. AVG", f"{curr_as_of_avg_ton:.2f} T/Day", "Current MTD Pace", "#2563eb"),
        ("LAST DAY REJECTION", f"{total_rej_ton:.3f} T", f"{total_rej_pcs:,} Pcs Lost", "#dc2626"),
        ("CRITICAL MC (>50)", f"{high_rej_count} MCs", "Lines Exceeding Limit", "#7c3aed"),
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
        ax.text(x0 + kpi_w / 2, 87.8, sub, color="#94a3b8", fontsize=6.8, ha="center")

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
        f"PLASTIC-3 MACHINE REJECTION LOG (>50 Pcs) — {day_formatted}",
        color="#0f172a",
        fontsize=11.0,
        fontweight="bold",
    )
    ax.text(
        73.5,
        83.5,
        f"{high_rej_count} Machines Active Above Threshold",
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
    ax.text(78.0, 83.5, "EXECUTIVE ANALYSIS", color="#0f172a", fontsize=11.0, fontweight="bold")

    # Dynamic Table
    n_count = len(df_day_filtered)

    if n_count <= 30:
        left_x = 2.6
        tbl_w = 71.8
        tbl_hdr = patches.Rectangle((left_x, 79.5), tbl_w, 2.6, facecolor="#1e293b", edgecolor="none")
        ax.add_patch(tbl_hdr)
        ax.text(left_x + 1.0, 80.8, "POSITION", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
        ax.text(left_x + 8.5, 80.8, "MACHINE", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
        ax.text(left_x + 18.0, 80.8, "DEFECT CAUSES", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")
        ax.text(left_x + 44.0, 80.8, "QTY", color="#ffffff", fontsize=7.2, fontweight="bold", ha="right", va="center")
        ax.text(left_x + 46.0, 80.8, "MOLD / ITEM", color="#ffffff", fontsize=7.2, fontweight="bold", va="center")

        row_y = 77.2
        row_step = min(3.8, 74.0 / max(1, n_count))
        for r_i, (_, r) in enumerate(df_day_filtered.iterrows()):
            bg_c = "#f8fafc" if r_i % 2 == 1 else "#ffffff"
            row_bg = patches.Rectangle((left_x, row_y - 1.2), tbl_w, row_step, facecolor=bg_c, edgecolor="none")
            ax.add_patch(row_bg)
            ax.plot([left_x, left_x + tbl_w], [row_y - 1.2, row_y - 1.2], color="#e2e8f0", linewidth=0.45)

            ax.text(left_x + 1.0, row_y + 0.35, str(r["Position"]), color="#0f172a", fontsize=6.8, fontweight="bold", va="center")
            ax.text(left_x + 8.5, row_y + 0.35, str(r["Machine"]), color="#64748b", fontsize=6.6, va="center")

            cause_wrap = "\n".join(textwrap.wrap(str(r["Causes"]), width=38))
            ax.text(left_x + 18.0, row_y + 0.35, cause_wrap, color="#b91c1c", fontsize=6.4, va="center")

            ax.text(left_x + 44.0, row_y + 0.35, f"{int(r['Qty']):,}", color="#0f172a", fontsize=7.0, fontweight="bold", ha="right", va="center")

            mold_wrap = "\n".join(textwrap.wrap(str(r["Mold"]), width=38))
            ax.text(left_x + 46.0, row_y + 0.35, mold_wrap, color="#334155", fontsize=6.4, va="center")
            row_y -= row_step

    else:
        mid_idx = (n_count + 1) // 2
        sub_a = df_day_filtered.iloc[:mid_idx]
        sub_b = df_day_filtered.iloc[mid_idx:]

        sub_configs = [(sub_a, 2.6, 35.8), (sub_b, 39.0, 35.8)]

        for sub_df, left_x, tbl_w in sub_configs:
            tbl_hdr = patches.Rectangle((left_x, 79.5), tbl_w, 2.6, facecolor="#1e293b", edgecolor="none")
            ax.add_patch(tbl_hdr)
            ax.text(left_x + 0.8, 80.8, "POS", color="#ffffff", fontsize=7.0, fontweight="bold", va="center")
            ax.text(left_x + 4.8, 80.8, "MACHINE", color="#ffffff", fontsize=7.0, fontweight="bold", va="center")
            ax.text(left_x + 10.4, 80.8, "DEFECT CAUSES", color="#ffffff", fontsize=7.0, fontweight="bold", va="center")
            ax.text(left_x + 23.0, 80.8, "QTY", color="#ffffff", fontsize=7.0, fontweight="bold", ha="right", va="center")
            ax.text(left_x + 24.0, 80.8, "MOLD / ITEM", color="#ffffff", fontsize=7.0, fontweight="bold", va="center")

            row_y = 77.2
            row_step = 3.75
            for r_i, (_, r) in enumerate(sub_df.iterrows()):
                bg_c = "#f8fafc" if r_i % 2 == 1 else "#ffffff"
                row_bg = patches.Rectangle((left_x, row_y - 1.2), tbl_w, row_step, facecolor=bg_c, edgecolor="none")
                ax.add_patch(row_bg)
                ax.plot([left_x, left_x + tbl_w], [row_y - 1.2, row_y - 1.2], color="#e2e8f0", linewidth=0.45)

                ax.text(left_x + 0.8, row_y + 0.35, str(r["Position"]), color="#0f172a", fontsize=6.7, fontweight="bold", va="center")
                ax.text(left_x + 4.8, row_y + 0.35, str(r["Machine"]), color="#64748b", fontsize=6.5, va="center")

                cause_wrap = "\n".join(textwrap.wrap(str(r["Causes"]), width=20))
                ax.text(left_x + 10.4, row_y + 0.35, cause_wrap, color="#b91c1c", fontsize=6.2, va="center")

                ax.text(left_x + 23.0, row_y + 0.35, f"{int(r['Qty']):,}", color="#0f172a", fontsize=6.8, fontweight="bold", ha="right", va="center")

                mold_wrap = "\n".join(textwrap.wrap(str(r["Mold"]), width=20))
                ax.text(left_x + 24.0, row_y + 0.35, mold_wrap, color="#334155", fontsize=6.2, va="center")
                row_y -= row_step

    # Right Executive Brief: 2 Full-Height Cards
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
    ax.text(78.6, 78.8, "Rejection Pareto & Top Causes", color="#b91c1c", fontsize=9.6, fontweight="bold")

    top3_lines = "\n".join([f"  {idx+1}. {c}: {p:,} pcs ({pct:.1f}%)" for idx, (c, p, pct) in enumerate(top3_summary_list)])
    t1 = (
        f"• Top 3 Causes ({top3_pct:.1f}% of loss):\n"
        f"{top3_lines}\n\n"
        f"• Heaviest Loss Machine:\n"
        f"  {top_wt_mc} ({top_wt_kg:.1f} kg loss).\n\n"
        f"• Critical Observations:\n"
        f"  Repetitive short filling noted\n"
        f"  on cutlery and box lid molds.\n"
        f"  Color change purging requires\n"
        f"  strict standardization."
    )
    ax.text(78.6, 75.2, t1, color="#7f1d1d", fontsize=8.2, linespacing=1.45, va="top")

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
    ax.text(78.6, 38.2, "Shop Floor & Plant Distribution", color="#15803d", fontsize=9.6, fontweight="bold")
    t2 = (
        f"• Total Plant Logged: {total_day_mcs} MCs\n"
        f"  - {high_rej_count} Lines > 50 pcs (Critical)\n"
        f"  - {total_day_mcs - high_rej_count} Lines <= 50 pcs (Controlled)\n\n"
        f"• Last Day Output Lost:\n"
        f"  {total_rej_pcs:,} Pcs / {total_rej_ton:.3f} Ton.\n\n"
        f"• Weight Share by Shop Floor:\n"
        f"  - GF Lines: {gf_share_pct:.1f}% of loss wt\n"
        f"  - FF Lines: {ff_share_pct:.1f}% of loss wt\n\n"
        f"• Monthly Rejection Pace:\n"
        f"  {curr_as_of_avg_ton:.2f} T/Day (vs {prev_avg_ton:.2f} Prev Mo)."
    )
    ax.text(78.6, 34.6, t2, color="#166534", fontsize=8.2, linespacing=1.45, va="top")

    plt.subplots_adjust(left=0, right=1, top=1, bottom=0)
    buf = io.BytesIO()
    plt.savefig(buf, format="jpg", facecolor=fig.get_facecolor(), edgecolor="none", dpi=220)
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
        st.markdown(
            "<h3 style='margin:0; text-align:center; font-weight:800; color:#0f172a;'>📉 DAILY REJECTION & DEFECT ANALYTICS</h3>",
            unsafe_allow_html=True,
        )
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

            uploaded_file = st.file_uploader(
                "Select Excel File (.xlsx, .xls)", type=["xlsx", "xls"], key="m2_uploader"
            )
            if uploaded_file is not None:
                if st.button(
                    "🚀 Ingest Rejection Data & Launch", type="primary", use_container_width=True
                ):
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
        df_as_of = df_curr[df_curr["DateClean"].dt.day <= sel_day_num].copy()
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
        if not df_as_of.empty and curr_wt_col:
            curr_as_of_total_ton = float(pd.to_numeric(df_as_of[curr_wt_col], errors="coerce").fillna(0).sum())
            curr_as_of_avg_ton = curr_as_of_total_ton / sel_day_num
        else:
            curr_as_of_total_ton, curr_as_of_avg_ton = 0.0, 0.0

        # Calculate Variance Metrics
        diff_ton = curr_as_of_avg_ton - prev_avg_ton
        pct_diff = (diff_ton / prev_avg_ton * 100.0) if prev_avg_ton > 0 else 0.0

        if diff_ton > 0:
            variance_line_plain = f"⚠️ Variance: Unfortunately, we are producing +{diff_ton:.2f} Tons/Day (+{pct_diff:.1f}%) more rejection compared to last month."
            variance_line_html = f'<p style="margin: 0 0 0.75rem 0; color: #dc2626; font-size: 0.85rem;">⚠️ <b>Variance:</b> Unfortunately, we are producing <b>+{diff_ton:.2f} Tons/Day (+{pct_diff:.1f}%)</b> more rejection compared to last month.</p>'
        elif diff_ton < 0:
            variance_line_plain = f"✅ Variance: We are producing {abs(diff_ton):.2f} Tons/Day ({abs(pct_diff):.1f}%) less rejection compared to last month."
            variance_line_html = f'<p style="margin: 0 0 0.75rem 0; color: #16a34a; font-size: 0.85rem;">✅ <b>Variance:</b> We are producing <b>{abs(diff_ton):.2f} Tons/Day ({abs(pct_diff):.1f}%)</b> less rejection compared to last month.</p>'
        else:
            variance_line_plain = "ℹ️ Variance: Daily rejection rate is on par with last month's baseline."
            variance_line_html = '<p style="margin: 0 0 0.75rem 0; color: #64748b; font-size: 0.85rem;">ℹ️ <b>Variance:</b> Daily rejection rate is on par with last month\'s baseline.</p>'

        # 4. Daily Totals & Drivers
        qty_col = get_col(df_day, ["Quantity", "Qty", "Rejection Pcs"], None)
        wt_col = get_col(df_day, ["Weight", "Rejection Ton"], None)
        cause_col = get_col(df_day, ["Cause", "Causes"], "Cause")
        mc_col = get_col(df_day, ["Machine", "MC SL"], "Machine")

        qty_factor = 1000.0 if (qty_col and df_day[qty_col].max() < 100) else 1.0
        total_rej_pcs = (
            int(round(pd.to_numeric(df_day[qty_col], errors="coerce").fillna(0).sum() * qty_factor))
            if (qty_col and not df_day.empty)
            else 0
        )
        total_rej_ton = (
            float(pd.to_numeric(df_day[wt_col], errors="coerce").fillna(0).sum())
            if (wt_col and not df_day.empty)
            else 0.0
        )
        high_rej_count = len(df_day_filtered)
        total_day_mcs = df_day[mc_col].nunique() if mc_col in df_day.columns else high_rej_count

        # Compute breakdowns
        df_cause_day = m2_compute_cause_breakdown(df_day)
        df_cause_as_of = m2_compute_cause_breakdown(df_as_of)
        df_lineman_day = m2_compute_lineman_breakdown(df_day)
        df_lineman_as_of = m2_compute_lineman_breakdown(df_as_of)
        df_trend = m2_compute_tonnage_comparison(df_prev, df_curr)

        # Top defect causes list & heaviest machine
        top3_summary_list = []
        if not df_day.empty and cause_col in df_day.columns and qty_col in df_day.columns:
            cause_grp = df_day.groupby(cause_col)[qty_col].sum() * qty_factor
            top_cause = cause_grp.idxmax() if not cause_grp.empty else "General"
            top_cause_pcs = int(round(cause_grp.max())) if not cause_grp.empty else 0
            top_cause_pct = (top_cause_pcs / total_rej_pcs * 100.0) if total_rej_pcs > 0 else 0.0

            top3_sorted = cause_grp.sort_values(ascending=False).head(3)
            top3_pcs = top3_sorted.sum()
            top3_pct = (top3_pcs / total_rej_pcs * 100.0) if total_rej_pcs > 0 else 0.0
            for c_name, c_qty in top3_sorted.items():
                c_clean = str(c_name).replace("*", "").strip()
                top3_summary_list.append(
                    (
                        c_clean,
                        int(round(c_qty)),
                        (c_qty / total_rej_pcs * 100.0) if total_rej_pcs > 0 else 0.0,
                    )
                )
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
            gf_share_pct = gf_wt / tot_w * 100.0
            ff_share_pct = ff_wt / tot_w * 100.0
        else:
            top_wt_mc, top_wt_kg, gf_share_pct, ff_share_pct = "-", 0.0, 80.0, 20.0

        # 5. Visual Export Button
        jpg_bytes = m2_generate_scrap_jpg(
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
            top3_summary_list,
        )

        with c_snap:
            st.markdown("<div style='margin-top: 1.65rem;'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📸 Download 1-Page JPG",
                data=jpg_bytes,
                file_name=f"Daily_Rejection_Report_{sel_date_str}.jpg",
                mime="image/jpeg",
                use_container_width=True,
            )
        st.markdown("</div>", unsafe_allow_html=True)

        # 6 KPI Cards in Web Dashboard
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.markdown(
            f'<div class="kpi-card indigo"><div class="kpi-title">PREV MO. TOTAL</div><div class="kpi-val">{prev_total_ton:.2f} T</div><div class="kpi-sub">Total Rejection</div></div>',
            unsafe_allow_html=True,
        )
        k2.markdown(
            f'<div class="kpi-card teal"><div class="kpi-title">PREV MO. AVG</div><div class="kpi-val">{prev_avg_ton:.2f} T/D</div><div class="kpi-sub">Daily Baseline</div></div>',
            unsafe_allow_html=True,
        )
        k3.markdown(
            f'<div class="kpi-card blue"><div class="kpi-title">THIS MO. AS OF</div><div class="kpi-val">{curr_as_of_total_ton:.2f} T</div><div class="kpi-sub">As of Day {sel_day_num}</div></div>',
            unsafe_allow_html=True,
        )
        k4.markdown(
            f'<div class="kpi-card purple"><div class="kpi-title">THIS MO. AVG</div><div class="kpi-val">{curr_as_of_avg_ton:.2f} T/D</div><div class="kpi-sub">As of Day {sel_day_num}</div></div>',
            unsafe_allow_html=True,
        )
        k5.markdown(
            f'<div class="kpi-card pink"><div class="kpi-title">LAST DAY REJECTION</div><div class="kpi-val">{total_rej_ton:.3f} T</div><div class="kpi-sub">{total_rej_pcs:,} Pcs</div></div>',
            unsafe_allow_html=True,
        )
        k6.markdown(
            f'<div class="kpi-card yellow"><div class="kpi-title">CRITICAL MC (&gt;{min_cutoff})</div><div class="kpi-val">{high_rej_count}</div><div class="kpi-sub">Lines Over Limit</div></div>',
            unsafe_allow_html=True,
        )

        st.markdown("<div style='margin-bottom: 1.15rem;'></div>", unsafe_allow_html=True)

        # Mid Section: Rejection Log Table & WhatsApp Note
        col_left, col_right = st.columns([1.55, 0.95], gap="medium")
        with col_left:
            st.markdown(f"### ⚙️ PLASTIC-3 MACHINE REJECTION LOG (&gt;{min_cutoff} Pcs) — {day_formatted}")
            if not df_day_filtered.empty:
                st.dataframe(
                    df_day_filtered[["Position", "Machine", "Causes", "Qty", "Weight (kg)", "Mold"]],
                    use_container_width=True,
                    hide_index=True,
                    height=380,
                )
            else:
                st.success("✅ No machines exceeded the rejection cutoff threshold today!")

        with col_right:
            approval_text = f"""📋 *PLASTIC-3 DAILY SCRAP & REJECTION BRIEF*
📅 *Date:* {day_formatted}

Dear Sir,

These are the line records from *Plastic-3* where rejection exceeded *50 pieces*:

🔹 *Prev. Month Total:* {prev_total_ton:.2f} Tons ({prev_avg_ton:.2f} T/Day)
🔹 *Present Month (As of Today):* {curr_as_of_total_ton:.2f} Tons ({curr_as_of_avg_ton:.2f} T/Day)
{variance_line_plain}

📌 *Please grant your approval to send these items for rejection clearance.*"""

            st.markdown("### 📝 EXECUTIVE APPROVAL TEXT")
            st.markdown(
                f"""<div class="narrative-block" style="font-size: 0.88rem; line-height: 1.6;">
                    <p style="margin: 0 0 0.5rem 0; font-weight: 800; color: #1e293b;">📋 PLASTIC-3 DAILY SCRAP & REJECTION BRIEF</p>
                    <p style="margin: 0 0 0.75rem 0; color: #64748b; font-size: 0.82rem;">📅 <b>Date:</b> {day_formatted}</p>
                    <p style="margin: 0 0 0.5rem 0;"><b>Dear Sir,</b></p>
                    <p>These are the line records from <b>Plastic-3</b> where rejection exceeded <b>50 pieces</b>:</p>
                    <p style="margin: 0.5rem 0 0.2rem 0;">🔹 <b>Prev. Month Total:</b> {prev_total_ton:.2f} Tons ({prev_avg_ton:.2f} T/Day)</p>
                    <p style="margin: 0 0 0.2rem 0;">🔹 <b>Present Month (As of Today):</b> {curr_as_of_total_ton:.2f} Tons ({curr_as_of_avg_ton:.2f} T/Day)</p>
                    {variance_line_html}
                    <p style="margin: 0.75rem 0 0 0; color: #dc2626; font-weight: 700;">📌 Please grant your approval to send these items for rejection clearance.</p>
                </div>""",
                unsafe_allow_html=True,
            )

            with st.expander("📋 Copy Plain Text for Approval / WhatsApp"):
                st.text_area(
                    "Approval Text", value=approval_text, height=180, label_visibility="collapsed"
                )

        st.divider()

        # Section 3: Cause-Wise Rejection Defect Analysis
        st.markdown("### 🔍 CAUSE-WISE REJECTION DEFECT ANALYSIS")
        tab_cause_day, tab_cause_asof = st.tabs([
            f"📅 Selected Date Breakdown ({day_formatted})",
            f"📈 As of Month-to-Date Defect Pareto (Day 1 – {sel_day_num})",
        ])
        with tab_cause_day:
            st.dataframe(df_cause_day, use_container_width=True, hide_index=True)
        with tab_cause_asof:
            st.dataframe(df_cause_as_of, use_container_width=True, hide_index=True)

        st.divider()

        # Section 4: Lineman-Wise Analysis (Column I - Added By)
        st.markdown("### 👷 LINEMAN-WISE REJECTION LOG ANALYSIS (ADDED BY)")
        tab_line_day, tab_line_asof = st.tabs([
            f"📅 Selected Date Linemen Activity ({day_formatted})",
            f"📈 As of Month-to-Date Linemen Overview (Day 1 – {sel_day_num})",
        ])
        with tab_line_day:
            if not df_lineman_day.empty:
                st.dataframe(df_lineman_day, use_container_width=True, hide_index=True)
            else:
                st.info("No lineman entries logged for this date.")
        with tab_line_asof:
            if not df_lineman_as_of.empty:
                st.dataframe(df_lineman_as_of, use_container_width=True, hide_index=True)
            else:
                st.info("No lineman entries logged for the current month.")

        st.divider()

        # Section 5: Month-over-Month Daily Trend
        st.markdown("### 📅 MONTH-OVER-MONTH DAILY REJECTION TONNAGE TREND")
        trend_display = df_trend.rename(
            columns={
                "Day": "Day of Month",
                "Prev_Month_Ton": "Previous Month Rejection (Tons)",
                "Curr_Month_Ton": "Current Month Rejection (Tons)",
            }
        )
        st.dataframe(trend_display, use_container_width=True, hide_index=True)
