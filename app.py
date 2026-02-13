import streamlit as st
from utils.styles import apply_custom_css
from utils.auth import init_session_state, show_login, get_menu_options, logout
from features.mvr_all_trans import run_mvr_all_trans
from features.hdvi_mvr import run_hdvi_mvr
from features.pdf_tools import run_pdf_maker, run_pdf_play
from features.body_gpt import run_body_gpt
from features.cause_gpt import run_cause_gpt
from features.accident_gpt import run_accident_gpt
from features.custom_gpt import run_custom_gpt

# --- APP CONFIGURATION ---
st.set_page_config(
    page_title="Clicknget AI Tools",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- APPLY AESTHETICS ---
apply_custom_css()

# --- INITIALIZE SESSION ---
init_session_state()

# --- AUTHENTICATION ---
if not st.session_state["authenticated"]:
    show_login()
    st.stop()

# --- SIDEBAR NAVIGATION ---
with st.sidebar:
    st.markdown(f"### 👋 Welcome, **{st.session_state['username']}**")
    st.markdown(f"**Role:** `{st.session_state['role']}`")
    st.markdown("---")

    menu_options = get_menu_options(st.session_state["role"])
    if menu_options:
        # Custom header for the menu
        st.markdown("### 📋 Navigation")
        menu = st.radio("Select a Tool", menu_options, label_visibility="collapsed")
    else:
        st.warning("No menu options available for your role.")
        menu = None

    st.markdown("---")
    logout()
    st.caption("v2.0.0 | Built with ❤️ by Yogaraj")

# --- MAIN CONTENT ROUTING ---
if menu:
    # Use a container for better layout
    main_container = st.container()
    
    with main_container:
        try:
            if menu == "MVR All Trans":
                run_mvr_all_trans()
            elif menu == "HDVI-MVR":
                run_hdvi_mvr()
            elif menu == "PDF Maker":
                run_pdf_maker()
            elif menu == "PDF Play":
                run_pdf_play()
            elif menu == "Body GPT":
                run_body_gpt()
            elif menu == "Cause GPT":
                run_cause_gpt()
            elif menu == "Accident GPT":
                run_accident_gpt()
            elif menu == "Custom GPT" or menu == "Categorization":
                run_custom_gpt()
        except Exception as e:
            st.error(f"### ❌ An unexpected error occurred in {menu}")
            st.error(str(e))
            with st.expander("Show Traceback"):
                import traceback
                st.code(traceback.format_exc())
else:
    st.info("Please select a tool from the sidebar to begin.")

# --- FOOTER / STATUS ---
st.markdown("---")
st.markdown(
    """
    <div style="display: flex; justify-content: space-between; align-items: center; opacity: 0.6; font-size: 0.8rem;">
        <div>Status: <span class="status-indicator status-operational"></span> Operational</div>
        <div>System ID: CLICKNGET-PRD-01</div>
    </div>
    """, 
    unsafe_allow_html=True
)