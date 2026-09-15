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

    # Top Industrial Corporate Header
    st.markdown(
        """
        <div class="top-nav-banner">
            <div class="brand-block">
                <div class="division-tag">PRAN-RFL GROUP &bull; INDUSTRIAL ENGINEERING DIVISION</div>
                <div class="main-title">PLANT OPERATIONS ANALYTICS PORTAL</div>
            </div>
            <div class="status-block">
                <div class="org-hierarchy">DIVISION: <span>RIP &gt; DPL</span></div>
                <div class="system-status"><span class="pulse-indicator"></span>SYSTEM OPERATIONAL</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Sub-Header Info Strip
    c_meta, c_action = st.columns([3.5, 1.0], vertical_alignment="center")
    with c_meta:
        st.markdown(
            """
            <div style="color: #475569; font-size: 0.88rem; font-weight: 500;">
                Autonomous Plant Floor Analytics Engine &bull; Automatic Multi-Section Resolution (Plastic-3, 6.1, 6.2, 7.1, 7.2, 7.3)
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c_action:
        if st.button("Reset Session Cache", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.session_state["active_view"] = "hub_home"
            st.rerun()

    st.markdown("<div style='margin-bottom: 1.75rem;'></div>", unsafe_allow_html=True)

    # 2 Core Operational Suites
    col_npt, col_scrap = st.columns(2, gap="large")

    with col_npt:
        st.markdown(
            """
            <div class="exec-card npt-card">
                <div class="card-kicker">OPERATIONAL EFFICIENCY</div>
                <h2 class="card-heading">Non-Productive Time (NPT) Analytics</h2>
                <p class="card-summary">
                    Precision machine breakdown accounting, standardized 8:00 AM &ndash; 8:00 AM operational day slicing, 
                    technical vs. non-technical stoppage attribution, mold setup (SMED) tracking, and section-specific capacity loss analysis.
                </p>
                <div class="metric-tags">
                    <span class="tag">8 AM &ndash; 8 AM Window</span>
                    <span class="tag">Auto-Section Capacity</span>
                    <span class="tag">SMED Benchmarking</span>
                    <span class="tag">2&times;2 Executive Export</span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Launch NPT Analytics Console", type="primary", key="btn_npt", use_container_width=True):
            st.session_state["active_view"] = "mod_npt"
            st.rerun()

    with col_scrap:
        st.markdown(
            """
            <div class="exec-card scrap-card">
                <div class="card-kicker">QUALITY CONTROL & RECOVERY</div>
                <h2 class="card-heading">Daily Rejection & Defect Analytics</h2>
                <p class="card-summary">
                    Item-level defect volume indexing, scrap tonnage reconciliation, 80/20 Pareto root-cause identification, 
                    lineman floor logging compliance audits, and executive clearance reporting.
                </p>
                <div class="metric-tags">
                    <span class="tag">Threshold Audits</span>
                    <span class="tag">Section Tonnage (T)</span>
                    <span class="tag">Defect Pareto</span>
                    <span class="tag">Excel & JPG Clearance</span>
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
