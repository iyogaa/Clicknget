import streamlit as st
from ui.components.styles import apply_custom_css
from ui.components.auth import init_session_state, show_login, get_menu_options, logout
from ui.pages.welcome_page import run_welcome
from ui.pages.mvr_summary_page import run_mvr_summary
from ui.pages.pdf_tools_consolidated import run_pdf_tools_consolidated

st.set_page_config(
    page_title="Clicknget AI Tools",
    page_icon="📄",
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

# --- SIMPLE SIDEBAR NAVIGATION ---
with st.sidebar:
    st.markdown("""
    <div style="padding: 1.5rem 0 1rem 0; text-align: center;">
        <h2 style="margin: 0;"> </h2>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Navigation
    main_menu = st.radio(
        "Navigation",
        ["App", "Summary Report", "PDF Tools"],
        label_visibility="collapsed"
    )
    
    st.markdown("---")
    
    # Logout
    if st.button("🚪 Logout", use_container_width=True):
        logout()

# --- MAIN CONTENT AREA ---
main_container = st.container()

with main_container:
    if main_menu:
        if main_menu == "App":
            run_welcome()
        elif main_menu == "Summary Report":
            run_mvr_summary()
        elif main_menu == "PDF Tools":
            run_pdf_tools_consolidated()

# --- FOOTER ---
st.markdown("---")
st.markdown("""
<div style="text-align: center; opacity: 0.6; font-size: 0.8rem; padding: 1rem 0;">
    Insight Board v2.0 | © 2024
</div>
""", unsafe_allow_html=True)