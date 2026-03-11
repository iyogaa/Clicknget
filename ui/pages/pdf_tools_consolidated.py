import streamlit as st
from ui.pages.pdf_tools_page import run_pdf_maker, run_pdf_play

def run_pdf_tools_consolidated():
    """Consolidated PDF Tools page with modern layout"""
    
    # Page header
    
    st.markdown("")
    
    # Tool selector with better layout
    col1, col2, col3 = st.columns([2, 3, 2])
    
    with col2:
        pdf_tool = st.selectbox(
            "Choose PDF Tool:",
            options=["PDF Maker", "PDF Play"],
            label_visibility="collapsed",
            key="pdf_tool_selector"
        )
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("")
    
    # Display selected tool with better header
    try:
        if pdf_tool == "PDF Maker":
            run_pdf_maker()
            
        elif pdf_tool == "PDF Play":
            run_pdf_play()
            
    except Exception as e:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.error(f"❌ Error loading PDF tool: {str(e)}")
        with col2:
            if st.button("Show Details", key="pdf_error_details"):
                with st.expander("🔍 Technical Details"):
                    import traceback
                    st.code(traceback.format_exc())
