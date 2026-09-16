# =========================================================
# OPERATIONS CONSOLE — ENTERPRISE MANUFACTURING PORTAL
# =========================================================
import importlib
import os
import streamlit as st

import modules.npt_analytics as npt_analytics
import modules.scrap_analytics as scrap_analytics

importlib.reload(scrap_analytics)
importlib.reload(npt_analytics)

st.set_page_config(
    page_title="RFL Operations Intelligence Portal | RIP-DPL",
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

# =========================================================
# ROUTING CONTROLLER
# =========================================================

# ---------------------------------------------------------
# VIEW 1: EXECUTIVE COMMAND CENTER (LANDING)
# ---------------------------------------------------------
if st.session_state["active_view"] == "hub_home":

    # Industrial Executive Dark Banner
    st.markdown(
        """
        <div style="background: #081225; border: 1px solid #1e293b; border-radius: 10px; padding: 1.25rem 2rem; margin-bottom: 1.5rem; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 4px 20px rgba(0,0,0,0.15);">
            <div>
                <div style="color: #38bdf8; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 0.2rem;">
                    PRAN-RFL GROUP &bull; INDUSTRIAL ENGINEERING DIVISION
                </div>
                <div style="color: #ffffff; font-size: 1.6rem; font-weight: 800; letter-spacing: -0.02em;">
                    PLANT OPERATIONS ANALYTICS PORTAL
                </div>
            </div>
            <div style="text-align: right;">
                <div style="color: #94a3b8; font-size: 0.8rem; font-weight: 600; letter-spacing: 0.05em; margin-bottom: 0.25rem;">
                    DIVISION: <span style="color: #f8fafc; font-weight: 700;">RIP &gt; DPL</span>
                </div>
                <div style="color: #22c55e; font-size: 0.78rem; font-weight: 700; letter-spacing: 0.08em; display: inline-flex; align-items: center; gap: 6px;">
                    <span style="display:inline-block; width:8px; height:8px; background:#22c55e; border-radius:50%; box-shadow: 0 0 8px #22c55e;"></span> SYSTEM OPERATIONAL
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Sub-header Status Toolbar
    c_sub, c_rst = st.columns([3.5, 1.0], vertical_alignment="center")
    with c_sub:
        st.markdown(
            """
            <div style="color: #475569; font-size: 0.9rem; font-weight: 500;">
                Autonomous Plant Floor Analytics Engine &bull; Multi-Section Auto-Resolution (Plastic-3, 6.1, 6.2, 7.1, 7.2, 7.3)
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c_rst:
        if st.button("Reset Session Cache", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.session_state["active_view"] = "hub_home"
            st.rerun()

    st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

    # 2 Core Operational Suites (Rendered inside native Streamlit border containers)
    col_npt, col_scrap = st.columns(2, gap="large")

    with col_npt:
        with st.container(border=True):
            st.markdown(
                """
                <div style="border-top: 4px solid #0284c7; padding-top: 1rem; margin-top: -0.5rem;">
                    <div style="color: #0284c7; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 0.35rem;">
                        OPERATIONAL EFFICIENCY
                    </div>
                    <h2 style="color: #0f172a; font-size: 1.4rem; font-weight: 800; margin: 0 0 0.75rem 0;">
                        Non-Productive Time (NPT) Analytics
                    </h2>
                    <p style="color: #475569; font-size: 0.88rem; line-height: 1.55; margin-bottom: 1.25rem;">
                        Precision machine breakdown accounting, standardized 8:00 AM &ndash; 8:00 AM operational day slicing, 
                        technical vs. non-technical stoppage attribution, mold setup (SMED) tracking, and section-specific capacity loss analysis.
                    </p>
                    <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 1.5rem;">
                        <span style="background: #f1f5f9; border: 1px solid #cbd5e1; color: #334155; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">8 AM &ndash; 8 AM Window</span>
                        <span style="background: #f1f5f9; border: 1px solid #cbd5e1; color: #334155; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">Auto-Section Capacity</span>
                        <span style="background: #f1f5f9; border: 1px solid #cbd5e1; color: #334155; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">SMED Benchmarking</span>
                        <span style="background: #f1f5f9; border: 1px solid #cbd5e1; color: #334155; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">2&times;2 Executive Export</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Launch NPT Analytics Console", type="primary", key="btn_npt", use_container_width=True):
                st.session_state["active_view"] = "mod_npt"
                st.rerun()

    with col_scrap:
        with st.container(border=True):
            st.markdown(
                """
                <div style="border-top: 4px solid #b91c1c; padding-top: 1rem; margin-top: -0.5rem;">
                    <div style="color: #b91c1c; font-size: 0.72rem; font-weight: 800; letter-spacing: 0.1em; text-transform: uppercase; margin-bottom: 0.35rem;">
                        QUALITY CONTROL & RECOVERY
                    </div>
                    <h2 style="color: #0f172a; font-size: 1.4rem; font-weight: 800; margin: 0 0 0.75rem 0;">
                        Daily Rejection & Defect Analytics
                    </h2>
                    <p style="color: #475569; font-size: 0.88rem; line-height: 1.55; margin-bottom: 1.25rem;">
                        Item-level defect volume indexing, scrap tonnage reconciliation, 80/20 Pareto root-cause identification, 
                        lineman floor logging compliance audits, and executive clearance reporting.
                    </p>
                    <div style="display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 1.5rem;">
                        <span style="background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">Threshold Audits</span>
                        <span style="background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">Section Tonnage (T)</span>
                        <span style="background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">Defect Pareto</span>
                        <span style="background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; font-size: 0.75rem; font-weight: 600; padding: 4px 10px; border-radius: 4px;">Excel & JPG Clearance</span>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("Launch Rejection Analytics Console", type="primary", key="btn_scrap", use_container_width=True):
                st.session_state["active_view"] = "mod_scrap"
                st.rerun()

# ---------------------------------------------------------
# VIEW 2: MODULE 1 — NON-PRODUCTIVE TIME (NPT) ANALYTICS
# ---------------------------------------------------------
elif st.session_state["active_view"] == "mod_npt":
    npt_analytics.render_npt_module()

# ---------------------------------------------------------
# VIEW 3: MODULE 2 — DAILY SCRAP & REJECTION ANALYTICS
# ---------------------------------------------------------
elif st.session_state["active_view"] == "mod_scrap":
    scrap_analytics.render_scrap_module()
