import streamlit as st
import pandas as pd
import io
import os

from core.processors.riscom_renewal import download_report_mvr_renewal_riscom_test2


def run_riscom_renewal_mvr():
    """RISCOM Renewal MVR Dashboard"""

    tab1, tab2 = st.tabs(["📤 Upload & Process", " "])

    with tab1:

        st.markdown("Upload Renewal ZIP File")

        uploaded_file = st.file_uploader(
            "Upload ZIP file (report_mvr + drivers)",
            type=["zip"],
            key="riscom_renewal_upload",
            label_visibility="collapsed"
        )

        if uploaded_file is not None:

            col1, col2, col3 = st.columns([2, 1.5, 2])

            with col2:
                if st.button("Process Renewal Report", use_container_width=True):

                    try:

                        with st.spinner("Processing Renewal MVR Report..."):

                            # Original file name
                            original_filename = uploaded_file.name
                            file_name_without_ext, _ = os.path.splitext(original_filename)

                            # Output file name
                            dynamic_output_name = f"{file_name_without_ext}.xlsx"

                            # Process ZIP
                            result_buffer = download_report_mvr_renewal_riscom_test2(uploaded_file)

                            # Some versions return tuple
                            if isinstance(result_buffer, (list, tuple)):
                                result_buffer = result_buffer[0]

                            # Optional preview (future use)
                            # df = pd.read_excel(result_buffer)

                            col1, col2, col3 = st.columns([2, 1.5, 2])

                            st.download_button(
                                label="Download",
                                data=result_buffer,
                                file_name=dynamic_output_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                            st.success("Processing complete!")

                    except Exception as e:

                        st.error(f"❌ An error occurred during processing: {str(e)}")

                        if st.checkbox("Show error details", key="renewal_show_details"):
                            with st.expander("Technical Details"):
                                import traceback
                                st.code(traceback.format_exc())