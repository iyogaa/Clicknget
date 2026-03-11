import streamlit as st
import io
from core.processors.alltrans_processor import Alltrans
from openpyxl import load_workbook

def run_mvr_all_trans():
    
    tab1, tab2 = st.tabs(["Upload & Process", " "])
    
    with tab1:
        st.markdown("Upload Client Excel")
        lookup_file = st.file_uploader("Upload Client Excel", type=["xlsx", "xls", "csv"], key="alltrans_client", label_visibility="collapsed")
        st.markdown("Upload Output Excel")
        main_file = st.file_uploader("Upload MVR File", type=["xlsx", "xls", "csv"], key="alltrans_mvr", label_visibility="collapsed")
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        if main_file and lookup_file:
            
            chosen_sheet = None
            try:
                # Try to see if it has sheets (Excel)
                lookup_wb = load_workbook(io.BytesIO(lookup_file.getvalue()), read_only=True)
                sheets = lookup_wb.sheetnames
                if len(sheets) > 1:
                    chosen_sheet = st.selectbox("Lookup workbook has multiple sheets. Choose one:", options=sheets, label_visibility="collapsed")
                else:
                    chosen_sheet = sheets[0]
                    st.info(f"✅ Detected sheet: **{chosen_sheet}**")
            except Exception:
                # Likely a CSV or not a standard Excel file
                st.info("📄 File format: CSV/Plain Text (No sheets)")
                chosen_sheet = None
            
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Process button
            col1, col2, col3 = st.columns([2, 1.5, 2])
            with col2:
                if st.button("Generate Report", use_container_width=True):
                    try:
                        main_bytes = main_file.read()
                        lookup_bytes = lookup_file.read()

                        gen = Alltrans(template_path="Template.xlsx",
                                    alltrans_sheet="All Trans",
                                    alltrans_header_row=4,#Fixed constant row for Column Headers
                                    mvr_sheet_name="MVR")
                        
                        with st.spinner("Processing Alltrans..."):
                            out = gen.run(main_bytes, lookup_bytes, chosen_lookup_sheet=chosen_sheet, preview_rows=8)

                        original_name = getattr(main_file, "name", None)
                        base = original_name.rsplit(".", 1)[0] if original_name else "Final_Report"
                        out_name = f"{base}.xlsx"
                        
                        #st.success("Report generated successfully!")
                        st.download_button(
                            "Download", 
                            data=out, 
                            file_name=out_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ Error: {e}")
                        if st.checkbox("Show technical details", key="alltrans_show_details"):
                            import traceback
                            st.code(traceback.format_exc())
    
    with tab2:
        st.markdown("""
        ### 📖 How to Use Alltrans Processing
        
        **Step 1: Prepare Your Files**
        - **MVR File**: Motor Vehicle Record data (Excel or CSV)
        - **Client File**: Customer lookup file with related data
        
        **Step 2: Upload Files**
        1. Upload your MVR file
        2. Upload your client Excel file
        3. If Excel has multiple sheets, select the correct one
        
        **Step 3: Generate Report**
        - Click the "Generate Report" button
        - Wait for processing to complete
        - Download the consolidated report
        
        **Step 4: Download Results**
        - The file combines MVR and client data
        - Ready for analysis and distribution
        
        ### 💡 Tips
        - Ensure both files have matching identifier columns
        - MVR sheet should be consistently formatted
        - Lookup file should have unique client records
        """)

