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
                if st.button("⚡ Process Report", use_container_width=True):
                    try:
                        with st.spinner("🔄 Processing MVR Report..."):
                            
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
                            st.markdown("""
                            <div style="margin: 2rem 0 1.5rem 0; padding-bottom: 1rem; border-bottom: 2px solid var(--border-color);">
                                <h3 style="margin: 0;">📊 Summary Metrics</h3>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            total_drivers = len(df)
                            approved_count = len(df[df['Status'].astype(str).str.lower() == 'approved'])
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
                            
                            with col1:
                                st.markdown(f"""
                                <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 1.5rem; text-align: center;">
                                    <p style="margin: 0; color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; font-weight: 600;">Total Processed</p>
                                    <p style="margin: 0.5rem 0 0 0; font-size: 2.5rem; font-weight: 700; color: #6366F1;">{total_drivers}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col2:
                                st.markdown(f"""
                                <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 1.5rem; text-align: center;">
                                    <p style="margin: 0; color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; font-weight: 600;">Approved</p>
                                    <p style="margin: 0.5rem 0 0 0; font-size: 2.5rem; font-weight: 700; color: #10B981;">{approved_count}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col3:
                                st.markdown(f"""
                                <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 1.5rem; text-align: center;">
                                    <p style="margin: 0; color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; font-weight: 600;">Pending</p>
                                    <p style="margin: 0.5rem 0 0 0; font-size: 2.5rem; font-weight: 700; color: #F59E0B;">{pending_count}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            col4, col5, col6 = st.columns(3)
                            
                            with col4:
                                st.markdown(f"""
                                <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 1.5rem; text-align: center;">
                                    <p style="margin: 0; color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; font-weight: 600;">Missing MVR</p>
                                    <p style="margin: 0.5rem 0 0 0; font-size: 2.5rem; font-weight: 700; color: #EF4444;">{missing_mvr_count}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col5:
                                st.markdown(f"""
                                <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 1.5rem; text-align: center;">
                                    <p style="margin: 0; color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; font-weight: 600;">Medical Issues</p>
                                    <p style="margin: 0.5rem 0 0 0; font-size: 2.5rem; font-weight: 700; color: #06B6D4;">{medical_issues_count}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            with col6:
                                st.markdown(f"""
                                <div style="background: var(--bg-card); border: 1px solid var(--border-color); border-radius: var(--radius-lg); padding: 1.5rem; text-align: center;">
                                    <p style="margin: 0; color: var(--text-muted); font-size: 0.75rem; text-transform: uppercase; font-weight: 600;">With Violations</p>
                                    <p style="margin: 0.5rem 0 0 0; font-size: 2.5rem; font-weight: 700; color: #EC4899;">{violations_count}</p>
                                </div>
                                """, unsafe_allow_html=True)
                            
                            st.markdown("")
                            
                            # --- DATA SECTION ---
                            st.markdown("""
                            <div style="margin: 2rem 0 1.5rem 0; padding-bottom: 1rem; border-bottom: 2px solid var(--border-color);">
                                <h3 style="margin: 0;">📋 Processed Data</h3>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            st.dataframe(df, use_container_width=True)
                            
                            st.markdown("")
                            
                            # Download button
                            col1, col2, col3 = st.columns([2, 1.5, 2])
                            with col2:
                                st.download_button(
                                    label="📥 Download Excel",
                                    data=buffer,
                                    file_name=dynamic_output_name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                            
                            st.success("✅ Processing complete!")
                            
                    except Exception as e:
                        st.error(f"❌ An error occurred during processing: {str(e)}")
                        if st.checkbox("Show error details", key="riscom_show_details"):
                            with st.expander("Technical Details"):
                                import traceback
                                st.code(traceback.format_exc())