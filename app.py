import io
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Operations Console",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="locked",
)

def load_css(file_name="style.css"):
    if os.path.exists(file_name):
        with open(file_name) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css("style.css")

# Session state initialization
if "app_launched" not in st.session_state:
    st.session_state["app_launched"] = False

# ============================================
# LANDING SETUP SCREEN
# ============================================
if not st.session_state["app_launched"]:
    st.markdown("## 🏭 **OPERATIONS CONSOLE SETUP**")
    st.markdown("##### Upload your source file to launch.")
    st.divider()

    st.markdown(
        '<div class="setup-card"><h3>📂 Source Data Entry</h3><p style="color:#64748b !important;">Select the Excel / CSV data file to begin</p></div>',
        unsafe_allow_html=True,
    )
    uploaded_file = st.file_uploader("Upload Data File (.xlsx, .csv)", type=["xlsx", "xls", "csv"], key="init_upload")

    st.divider()

    c_btn, _ = st.columns([1, 3])
    with c_btn:
        if st.button("🚀 Launch Dashboard", type="primary", use_container_width=True):
            if uploaded_file is None:
                st.error("Please upload a data file to launch.")
            else:
                st.session_state["file_bytes"] = uploaded_file.getvalue()
                st.session_state["file_name"] = uploaded_file.name
                st.session_state["app_launched"] = True
                st.rerun()

# ============================================
# MAIN APPLICATION CONSOLE
# ============================================
else:
    # Sidebar layout with logo and branding
    with st.sidebar:
        col_logo, col_text = st.columns([1, 2.3], gap="small", vertical_alignment="center")
        with col_logo:
            if os.path.exists("logo.png"):
                st.image("logo.png", use_container_width=True)
            else:
                st.markdown("🏭")
        with col_text:
            st.markdown("### **OPERATIONS CONSOLE**")
            st.caption("Active Session")

        st.divider()

        nav_choice = st.radio(
            "📍 **Select Module:**",
            [
                "📊 Summary Overview",
                "📦 Detailed Log",
            ],
        )

        st.divider()

        if st.button("⚙️ Change Uploaded File", use_container_width=True):
            st.session_state["app_launched"] = False
            st.session_state.pop("file_bytes", None)
            st.session_state.pop("file_name", None)
            st.rerun()

    # Load data
    file_stream = io.BytesIO(st.session_state["file_bytes"])
    if st.session_state["file_name"].endswith(".csv"):
        df_raw = pd.read_csv(file_stream)
    else:
        df_raw = pd.read_excel(file_stream)

    st.markdown(f"## {nav_choice}")
    st.divider()

    # Starter Display Area
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Rows", f"{len(df_raw):,}")
    c2.metric("Total Columns", f"{len(df_raw.columns):,}")
    c3.metric("Status", "🟢 Ready")

    st.divider()
    st.markdown("### 📋 Preview Table")
    st.dataframe(df_raw, use_container_width=True, hide_index=True)
