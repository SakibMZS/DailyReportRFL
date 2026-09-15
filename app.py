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
    page_title="Operations Console | Industrial Engineering",
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

if "selected_floor" not in st.session_state:
    st.session_state["selected_floor"] = "Plastic-3"

# =========================================================
# ROUTING CONTROLLER
# =========================================================

# ---------------------------------------------------------
# VIEW 1: ENTERPRISE HUB HOME
# ---------------------------------------------------------
if st.session_state["active_view"] == "hub_home":

    # Top Brand Navigation Bar
    st.markdown(
        """
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 0.8rem 1.5rem; background: #091e3a; border-radius: 8px; margin-bottom: 1.5rem;">
            <div>
                <div style="color: #38bdf8; font-size: 0.75rem; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase;">Enterprise Industrial Intelligence</div>
                <div style="color: #ffffff; font-size: 1.4rem; font-weight: 800; letter-spacing: -0.01em;">OPERATIONS ANALYTICS PORTAL</div>
            </div>
            <div style="color: #94a3b8; font-size: 0.85rem; font-weight: 500;">
                Injection Molding Division
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Global Context Controls
    c_context, c_action = st.columns([3, 1], vertical_alignment="center")
    with c_context:
        floor_options = ["Plastic-3"]  # Expanded automatically when multi-floor config is plugged in
        selected_fl = st.selectbox(
            "Select Manufacturing Section / Floor Baseline",
            floor_options,
            index=floor_options.index(st.session_state["selected_floor"]),
            key="plant_floor_select",
        )
        st.session_state["selected_floor"] = selected_fl

    with c_action:
        if st.button("Clear Cache & Session", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.session_state["active_view"] = "hub_home"
            st.session_state["selected_floor"] = "Plastic-3"
            st.rerun()

    st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)

    # 2 Core Operational Suites
    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.markdown(
            """
            <div class="enterprise-module-card">
                <div class="card-status-pill status-npt">Operational Suite</div>
                <h3 class="card-title">Non-Productive Time (NPT) Analytics</h3>
                <p class="card-desc">
                    Comprehensive machine breakdown tracking, operational day (8 AM–8 AM) duration slicing, 
                    technical vs. non-technical downtime attribution, mold setup (SMED) benchmarking, and plant capacity impact analysis.
                </p>
                <div class="card-data-requirement">
                    <span class="req-label">Data Ingestion Requirement:</span><br>
                    Please upload the relevant file containing data from the start of the last month to the current month's latest operational date.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Launch NPT Analytics Module", type="primary", key="btn_launch_npt", use_container_width=True):
            st.session_state["active_view"] = "mod_npt"
            st.rerun()

    with col2:
        st.markdown(
            """
            <div class="enterprise-module-card">
                <div class="card-status-pill status-scrap">Quality Control Suite</div>
                <h3 class="card-title">Rejection & Scrap Analytics</h3>
                <p class="card-desc">
                    Item-level defect volume indexing, scrap weight tonnage reconciliation, Pareto root-cause ranking, 
                    lineman logging audit, and like-for-like month-over-month variance clearance reporting.
                </p>
                <div class="card-data-requirement">
                    <span class="req-label">Data Ingestion Requirement:</span><br>
                    Please upload the relevant file containing data from the start of the last month to the current month's latest operational date.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Launch Rejection Analytics Module", type="primary", key="btn_launch_scrap", use_container_width=True):
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
