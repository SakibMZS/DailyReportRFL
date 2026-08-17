# =========================================================
# OPERATIONS CONSOLE — FULL-WIDTH ENTERPRISE PLATFORM
# =========================================================
import os
import streamlit as st

# Import isolated modules
from modules.size_wise import render_size_wise_module

st.set_page_config(
    page_title="Operations Console | Daily Report RFL",
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
# VIEW ROUTING
# =========================================================

# ---------------------------------------------------------
# VIEW 1: HUB HOME / OVERVIEW
# ---------------------------------------------------------
if st.session_state["active_view"] == "hub_home":
    c_brand, c_meta = st.columns([3, 1])
    with c_brand:
        st.markdown("## 🏭 **OPERATIONS CONSOLE & REPORTING HUB**")
        st.caption("Centralized Industrial Engineering & Daily Operational Intelligence")
    with c_meta:
        if st.button("🗑️ Reset All Sessions / Clear Cache", use_container_width=True):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

    st.divider()

    st.markdown("### 📋 **Select a Reporting Module to Launch**")

    # Top Row of Module Cards
    c1, c2 = st.columns(2, gap="large")

    with c1:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #2563eb;">
                <div>
                    <span style="background: #e0f2fe; color: #0284c7; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Active Module</span>
                    <h3 style="margin-top: 0.75rem;">📊 Size-Wise Performance & HR</h3>
                    <p>
                        Daily factory efficiency briefs, machine-size capacity breakdown, shift manpower productivity indexing (Pcs/HR & kg/HR), WhatsApp copy text, and pixel-perfect 1-page visual report exports.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("🚀 Launch Size-Wise Module", type="primary", use_container_width=True):
            st.session_state["active_view"] = "mod_size_wise"
            st.rerun()

    with c2:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #8b5cf6;">
                <div>
                    <span style="background: #f3e8ff; color: #7c3aed; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Workspace Ready</span>
                    <h3 style="margin-top: 0.75rem;">⏱️ Mold Changeover & Downtime</h3>
                    <p>
                        Track SMED mold changeover benchmarks, mechanical toggle system conversions vs hydraulic cylinders, tool setup times, and line stoppage losses.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⚙️ Launch Changeover Module", use_container_width=True):
            st.session_state["active_view"] = "mod_changeover"
            st.rerun()

    st.markdown("<div style='margin-bottom: 1.25rem;'></div>", unsafe_allow_html=True)

    # Bottom Row of Module Cards
    c3, c4 = st.columns(2, gap="large")

    with c3:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #ef4444;">
                <div>
                    <span style="background: #fee2e2; color: #dc2626; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Workspace Ready</span>
                    <h3 style="margin-top: 0.75rem;">📉 Daily Scrap & Defect Analytics</h3>
                    <p>
                        Analyze item-level rejection quantities, resin purge losses, masterbatch color change scrap, and Six Sigma quality defect distributions.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⚙️ Launch Scrap Module", use_container_width=True):
            st.session_state["active_view"] = "mod_scrap"
            st.rerun()

    with c4:
        st.markdown(
            """
            <div class="hub-card" style="border-top: 5px solid #10b981;">
                <div>
                    <span style="background: #dcfce7; color: #16a34a; padding: 4px 10px; border-radius: 6px; font-size: 0.72rem; font-weight: 800; text-transform: uppercase;">Workspace Ready</span>
                    <h3 style="margin-top: 0.75rem;">📈 Monthly Trends & OEE Analytics</h3>
                    <p>
                        Month-to-Date (MTD) cumulative production curves, availability / performance / quality OEE factoring, and plant capacity utilization insights.
                    </p>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("⚙️ Launch OEE Module", use_container_width=True):
            st.session_state["active_view"] = "mod_oee"
            st.rerun()


# ---------------------------------------------------------
# VIEW 2: MODULE 1 — SIZE-WISE PERFORMANCE & HR
# ---------------------------------------------------------
elif st.session_state["active_view"] == "mod_size_wise":
    render_size_wise_module()


# ---------------------------------------------------------
# VIEW 3, 4, 5: FUTURE RESERVED WORKSPACES
# ---------------------------------------------------------
elif st.session_state["active_view"] == "mod_changeover":
    if st.button("⬅️ Back to Operations Hub"):
        st.session_state["active_view"] = "hub_home"
        st.rerun()
    st.divider()
    st.markdown("## ⏱️ **MOLD CHANGEOVER & DOWNTIME MODULE**")
    st.info("🛠️ This full-width module workspace is ready for your changeover metrics and dataset structure.")

elif st.session_state["active_view"] == "mod_scrap":
    if st.button("⬅️ Back to Operations Hub"):
        st.session_state["active_view"] = "hub_home"
        st.rerun()
    st.divider()
    st.markdown("## 📉 **DAILY SCRAP & DEFECT ANALYTICS MODULE**")
    st.info("🛠️ This full-width module workspace is ready for your scrap tracking metrics and dataset structure.")

elif st.session_state["active_view"] == "mod_oee":
    if st.button("⬅️ Back to Operations Hub"):
        st.session_state["active_view"] = "hub_home"
        st.rerun()
    st.divider()
    st.markdown("## 📈 **MONTHLY TRENDS & OEE ANALYTICS MODULE**")
    st.info("🛠️ This full-width module workspace is ready for your OEE tracking metrics and dataset structure.")
