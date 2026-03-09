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
            options=["PDF Maker", "PDF Play", "Excel to PDF Converter"],
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
            
        elif pdf_tool == "Excel to PDF Converter":
            st.markdown("""
            <div style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.1), rgba(16, 185, 129, 0.05)); border-left: 4px solid #10B981; border-radius: var(--radius-md); padding: 1rem; margin-bottom: 1.5rem; border: 1px solid rgba(16, 185, 129, 0.3);">
                <div style="display: flex; align-items: center; gap: 0.75rem;">
                    <div style="font-size: 1.75rem;">📊</div>
                    <div>
                        <h3 style="margin: 0; color: #10B981;">Excel to PDF Converter</h3>
                        <p style="margin: 0.25rem 0 0 0; color: var(--text-muted); font-size: 0.85rem;">Convert Excel spreadsheets to professional PDF documents</p>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.info("📄 Excel to PDF conversion is powered by PDF Play tools. Select 'PDF Play' above to get started.")
            
    except Exception as e:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.error(f"❌ Error loading PDF tool: {str(e)}")
        with col2:
            if st.button("Show Details", key="pdf_error_details"):
                with st.expander("🔍 Technical Details"):
                    import traceback
                    st.code(traceback.format_exc())
