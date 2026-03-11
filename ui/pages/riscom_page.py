import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from core.processors.riscom_processor import process_riscom_mvr_data
import os

def run_riscom_mvr():
    """Riscom MVR Dashboard with modern metrics and layout"""
    
    tab1, tab2 = st.tabs(["📤 Upload & Process", " "])
    
    with tab1:
        st.markdown("Upload Output File")
        uploaded_file = st.file_uploader("Choose an Excel file (report_file.xlsx)", type=["xlsx"], key="riscom_upload", label_visibility="collapsed")
        
        st.markdown("</div>", unsafe_allow_html=True)
        
        if uploaded_file is not None:
            # Process button
            col1, col2, col3 = st.columns([2, 1.5, 2])
            with col2:
                if st.button("Process Report", use_container_width=True):
                    try:
                        with st.spinner("Processing MVR Report..."):
                            
                            # Get Original File Name
                            original_filename = uploaded_file.name
                            file_name_without_ext, _ = os.path.splitext(original_filename)

                            # Remove "report_" prefix if exists
                            if file_name_without_ext.lower().startswith("report_"):
                                file_name_without_ext = file_name_without_ext[len("report_"):]

                            dynamic_output_name = f"{file_name_without_ext}.xlsx"

                            # Process File
                            file_content = uploaded_file.getvalue()
                            mvr_file_buffer = io.BytesIO(file_content)
                            
                            try:
                                original_wb = load_workbook(io.BytesIO(file_content))
                            except Exception:
                                st.error("❌ Invalid Excel file format. Please upload a valid .xlsx file.")
                                st.stop()
                            
                            buffer, processed_data = process_riscom_mvr_data(mvr_file_buffer, original_wb)
                            
                            df = pd.DataFrame(processed_data)
                            
                            # --- SUMMARY METRICS SECTION ---
                            
                            #total_drivers = len(df)
                            #approved_count = len(df[df['Status'].astype(str).str.lower() == 'approved'])
                            pending_count = len(df[df['Status'].astype(str).str.lower() == 'pending'])
                            
                            missing_mvr_count = len(df[df.get('MVR Received', 'FALSE').astype(str).str.upper() == 'FALSE'])
                            
                            medical_issues_count = len(
                                df[df.get('Comments', '').astype(str).str.contains("Medical", case=False, na=False)]
                            )
                            
                            has_violations = df.apply(lambda x: (
                                pd.to_numeric(x.get('Minor Count', 0), errors='coerce') > 0 or 
                                pd.to_numeric(x.get('Major Count', 0), errors='coerce') > 0 or
                                pd.to_numeric(x.get('Accident Count', 0), errors='coerce') > 0
                            ), axis=1)

                            violations_count = len(df[has_violations])

                            # Render metrics with modern cards
                            col1, col2, col3 = st.columns(3)
                            
                            
                            col4, col5, col6 = st.columns(3)
                            
                            
                            
                            #st.dataframe(df, use_container_width=True)
                            
                            #st.markdown("")
                            # Download button
                            col1, col2, col3 = st.columns([2, 1.5, 2])
                            st.download_button(
                                label="Download",
                                data=buffer,
                                file_name=dynamic_output_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            
                            st.success("Processing complete!")
                            
                    except Exception as e:
                        st.error(f"❌ An error occurred during processing: {str(e)}")
                        if st.checkbox("Show error details", key="riscom_show_details"):
                            with st.expander("Technical Details"):
                                import traceback
                                st.code(traceback.format_exc())