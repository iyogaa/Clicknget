import pandas as pd
import streamlit as st
import numpy as np
import io
from datetime import timedelta

CONFIG = {
    'column_map': {
        'task_name': 'task_name',
        'task_id': 'task_id', 
        'user_name': 'user_name',
        'making_user': 'making_user',
        'class_name': 'class_name',
        'creation_date': 'creation_date',
        'creation_time': 'creation_time',
        'pick_time': 'pick_time',
        'making_start_time': 'making_start_time',
        'making_end_time': 'making_end_time',
        'completed_time': 'completed_time'
    },
    'ca_clients': ['HDVI_HITL_70', 'RISCOM_HITL_88', 'AllTrans_HITL_87'],
    'wc_clients': ['Method_HITL_67', 'Foresight_HITL_127'],
    'workflow_mapping': {
        'HDVI_HITL_70': ['MVR', 'IFTA'],
        'AllTrans_HITL_87': ['MVR'],
        'RISCOM_HITL_88': ['MVR'],
        'Method_HITL_67': ['SUPPLEMENT_WORKERCOMP'],
        'Foresight_HITL_127': ['SUPPLEMENT_WORKERCOMP']
    }
}

def load_and_combine_excel(file_path, file_type='both'):
    try:
        excel_file = pd.ExcelFile(file_path)
        sheets_data = []
        allowed_clients = []
        
        if file_type == 'ca':
            allowed_clients = CONFIG['ca_clients']
        elif file_type == 'wc':
            allowed_clients = CONFIG['wc_clients']
        else:
            allowed_clients = CONFIG['ca_clients'] + CONFIG['wc_clients']
        
        for sheet_name in excel_file.sheet_names:
            if sheet_name in allowed_clients:
                df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
                df_sheet['client'] = sheet_name
                if sheet_name in CONFIG['ca_clients']:
                    df_sheet['file_type'] = 'CA'
                elif sheet_name in CONFIG['wc_clients']:
                    df_sheet['file_type'] = 'WC'
                sheets_data.append(df_sheet)
        
        if not sheets_data:
            return pd.DataFrame()
        
        return pd.concat(sheets_data, ignore_index=True)
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return pd.DataFrame()

def load_and_combine_excel_from_bytes(file_bytes, file_type='both'):
    try:
        excel_file = pd.ExcelFile(io.BytesIO(file_bytes))
        sheets_data = []
        allowed_clients = []
        
        if file_type == 'ca':
            allowed_clients = CONFIG['ca_clients']
        elif file_type == 'wc':
            allowed_clients = CONFIG['wc_clients']
        else:
            allowed_clients = CONFIG['ca_clients'] + CONFIG['wc_clients']
        
        for sheet_name in excel_file.sheet_names:
            if sheet_name in allowed_clients:
                df_sheet = pd.read_excel(excel_file, sheet_name=sheet_name)
                df_sheet['client'] = sheet_name
                if sheet_name in CONFIG['ca_clients']:
                    df_sheet['file_type'] = 'CA'
                elif sheet_name in CONFIG['wc_clients']:
                    df_sheet['file_type'] = 'WC'
                sheets_data.append(df_sheet)
        
        if not sheets_data:
            return pd.DataFrame()
        
        return pd.concat(sheets_data, ignore_index=True)
    except Exception as e:
        st.error(f"Error loading Excel file: {str(e)}")
        return pd.DataFrame()

def normalize_columns(df, column_map=None):
    if column_map is None:
        column_map = CONFIG['column_map']
    result_df = df.copy()
    rename_dict = {v: k for k, v in column_map.items() if v in df.columns}
    result_df = result_df.rename(columns=rename_dict)
    expected_cols = list(column_map.keys()) + ['client']
    if 'file_type' in result_df.columns:
        expected_cols.append('file_type')
    for col in expected_cols:
        if col not in result_df.columns:
            result_df[col] = pd.NA
    return result_df[expected_cols]

def parse_datetimes(df, cols):
    result_df = df.copy()
    for col in cols:
        if col in result_df.columns:
            date_col = col.replace('_time', '_date')
            if date_col in result_df.columns and col in ['creation_time', 'pick_time', 'completed_time']:
                result_df[col] = pd.to_datetime(result_df[date_col].astype(str) + ' ' + result_df[col].astype(str), errors='coerce')
            else:
                result_df[col] = pd.to_datetime(result_df[col], errors='coerce')
    return result_df

def assign_workflow(row):
    client = row.get('client', '')
    class_name = str(row.get('class_name', '')).upper()
    
    workflows = CONFIG['workflow_mapping'].get(client, [])
    
    if client == 'HDVI_HITL_70':
        if 'IFTA' in class_name:
            return 'IFTA'
        else:
            return 'MVR'
    elif client in CONFIG['ca_clients']:
        return 'MVR'
    elif client in CONFIG['wc_clients']:
        return 'SUPPLEMENT_WORKERCOMP'
    
    for workflow in ['MVR', 'IFTA', 'SUPPLEMENT_WORKERCOMP']:
        if workflow in class_name:
            return workflow
    
    return 'UNKNOWN'

def compute_derived_metrics(df):
    result_df = df.copy()
    
    if 'workflow' not in result_df.columns:
        result_df['workflow'] = result_df.apply(assign_workflow, axis=1)
    
    if 'file_type' not in result_df.columns:
        result_df['file_type'] = result_df['client'].apply(
            lambda x: 'CA' if x in CONFIG['ca_clients'] else ('WC' if x in CONFIG['wc_clients'] else 'UNKNOWN')
        )
    
    result_df['parsing_error'] = False
    result_df['duplicate_task_id'] = False
    result_df['missing_making_record'] = False
    result_df['error_flag'] = 'none'
    
    timestamp_cols = ['creation_time', 'pick_time', 'making_start_time', 'making_end_time', 'completed_time']
    has_parsing_error = result_df[timestamp_cols].isna().any(axis=1)
    result_df.loc[has_parsing_error, 'parsing_error'] = True
    
    task_id_counts = result_df['task_id'].value_counts()
    duplicate_ids = task_id_counts[task_id_counts > 1].index
    result_df['duplicate_task_id'] = result_df['task_id'].isin(duplicate_ids)
    
    result_df['missing_making_record'] = (result_df['making_start_time'].isna() | result_df['making_end_time'].isna())
    
    conditions = [
        result_df['parsing_error'],
        result_df['duplicate_task_id'],
        (result_df['pick_time'].isna() | result_df['completed_time'].isna()),
        (result_df['completed_time'] < result_df['pick_time']),
        (result_df['making_end_time'] < result_df['making_start_time'])
    ]
    choices = ['parsing_error', 'duplicate_task_id', 'missing_pick_or_complete', 'completion_before_pick', 'making_end_before_start']
    result_df['error_flag'] = np.select(conditions, choices, default='none')
    
    result_df['date'] = pd.NaT
    mask_pick_time_valid = result_df['pick_time'].notna()
    result_df.loc[mask_pick_time_valid, 'date'] = result_df.loc[mask_pick_time_valid, 'pick_time'].dt.date
    mask_pick_time_missing = result_df['pick_time'].isna() & result_df['creation_time'].notna()
    result_df.loc[mask_pick_time_missing, 'date'] = result_df.loc[mask_pick_time_missing, 'creation_time'].dt.date
    
    mask = ~result_df['pick_time'].isna() & ~result_df['completed_time'].isna()
    result_df.loc[mask, 'task_duration_min'] = (result_df.loc[mask, 'completed_time'] - result_df.loc[mask, 'pick_time']).dt.total_seconds() / 60
    
    mask = ~result_df['making_start_time'].isna() & ~result_df['making_end_time'].isna()
    result_df.loc[mask, 'making_time_min'] = (result_df.loc[mask, 'making_end_time'] - result_df.loc[mask, 'making_start_time']).dt.total_seconds() / 60
    
    mask = ~result_df['pick_time'].isna() & ~result_df['making_start_time'].isna()
    result_df.loc[mask, 'idle_before_making_min'] = (result_df.loc[mask, 'making_start_time'] - result_df.loc[mask, 'pick_time']).dt.total_seconds() / 60
    
    mask = ~result_df['making_end_time'].isna() & ~result_df['completed_time'].isna()
    result_df.loc[mask, 'idle_after_making_min'] = (result_df.loc[mask, 'completed_time'] - result_df.loc[mask, 'making_end_time']).dt.total_seconds() / 60
    
    result_df['total_idle_min'] = result_df['task_duration_min'] - result_df['making_time_min']
    
    mask = ~result_df['creation_time'].isna() & ~result_df['pick_time'].isna()
    result_df.loc[mask, 'task_aging_min'] = (result_df.loc[mask, 'pick_time'] - result_df.loc[mask, 'creation_time']).dt.total_seconds() / 60
    
    mask = ~result_df['creation_time'].isna() & ~result_df['completed_time'].isna()
    result_df.loc[mask, 'full_lifecycle_min'] = (result_df.loc[mask, 'completed_time'] - result_df.loc[mask, 'creation_time']).dt.total_seconds() / 60
    
    result_df['efficiency_ratio'] = result_df['making_time_min'] / result_df['task_duration_min']
    
    negative_duration_mask = ((result_df['task_duration_min'] < 0) | (result_df['making_time_min'] < 0) | (result_df['idle_before_making_min'] < 0) | (result_df['idle_after_making_min'] < 0))
    result_df.loc[negative_duration_mask & (result_df['error_flag'] == 'none'), 'error_flag'] = 'negative_task_duration'
    
    return result_df

def create_user_detailed_view(user, filtered_df):
    user_tasks = filtered_df[filtered_df['making_user'] == user].copy()
    if user_tasks.empty:
        return
    
    col1, col2, col3, col4 = st.columns(4)
    
    completed_tasks = user_tasks[user_tasks['completed_time'].notna()]
    total_tasks = len(user_tasks)
    completed_count = len(completed_tasks)
    
    with col1:
        st.metric("Total Tasks", total_tasks)
    with col2:
        st.metric("Completed Tasks", completed_count)
    with col3:
        avg_duration = completed_tasks['task_duration_min'].mean() if not completed_tasks.empty else 0
        st.metric("Avg Duration", f"{avg_duration:.1f} min" if not pd.isna(avg_duration) else "N/A")
    with col4:
        avg_efficiency = completed_tasks['efficiency_ratio'].mean() * 100 if not completed_tasks.empty else 0
        st.metric("Avg Efficiency", f"{avg_efficiency:.1f}%" if not pd.isna(avg_efficiency) else "N/A")
    
    if 'workflow' in user_tasks.columns:
        st.subheader("📊 Workflow Breakdown")
        workflow_stats = user_tasks.groupby('workflow').agg({
            'task_id': 'count',
            'task_duration_min': 'mean',
            'completed_time': lambda x: x.notna().sum()
        }).reset_index()
        workflow_stats.columns = ['Workflow', 'Total Tasks', 'Avg Duration (min)', 'Completed']
        st.dataframe(workflow_stats, use_container_width=True)
    
    st.subheader("📅 Task Timeline")
    user_tasks_sorted = user_tasks[user_tasks['pick_time'].notna()].sort_values('pick_time', ascending=False)
    if not user_tasks_sorted.empty:
        display_cols = ['task_id', 'workflow', 'file_type', 'client', 'pick_time', 'completed_time', 'task_duration_min']
        display_cols = [col for col in display_cols if col in user_tasks_sorted.columns]
        st.dataframe(user_tasks_sorted[display_cols].head(50), use_container_width=True)
    else:
        st.info("No tasks with pickup time available")

def create_user_summary_table(filtered_df):
    users = sorted(filtered_df['making_user'].dropna().unique())
    user_summary_data = []
    
    for user in users:
        user_tasks = filtered_df[filtered_df['making_user'] == user]
        completed_tasks = user_tasks[user_tasks['completed_time'].notna()]
        
        workflows = ', '.join(user_tasks['workflow'].unique()) if 'workflow' in user_tasks.columns else 'N/A'
        file_types = ', '.join(user_tasks['file_type'].unique()) if 'file_type' in user_tasks.columns else 'N/A'
        
        user_summary_data.append({
            'User': user,
            'Total Tasks': len(user_tasks),
            'Completed': len(completed_tasks),
            'Workflows': workflows,
            'File Types': file_types,
            'First Pickup': user_tasks['pick_time'].min() if user_tasks['pick_time'].notna().any() else 'N/A',
            'Last Pickup': user_tasks['pick_time'].max() if user_tasks['pick_time'].notna().any() else 'N/A',
            'Avg Duration (min)': completed_tasks['task_duration_min'].mean() if not completed_tasks.empty and completed_tasks['task_duration_min'].notna().any() else 'N/A'
        })
    
    return pd.DataFrame(user_summary_data)

def run_user_wise_dashboard():
    st.title("User-wise Task Performance Dashboard")
    st.markdown("### Process CA and WC Excel Files for User Insights")
    
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'file_processed' not in st.session_state:
        st.session_state.file_processed = False
    
    with st.sidebar:
        st.header("📁 Data Input")
        st.subheader("CA Excel File")
        uploaded_ca_file = st.file_uploader("Upload CA Excel file", type=['xlsx', 'xls'], key="ca_file")
        ca_file_path = st.text_input("Or enter CA file path:", value="", key="ca_path")
        
        st.subheader("WC Excel File")
        uploaded_wc_file = st.file_uploader("Upload WC Excel file", type=['xlsx', 'xls'], key="wc_file")
        wc_file_path = st.text_input("Or enter WC file path:", value="", key="wc_path")
        
        st.header("⚙️ Settings")
        analysis_period = st.selectbox("Period", ["Last 7 days", "Last 30 days", "Last 90 days", "Custom", "All data"], index=1)
    
    has_ca_file = uploaded_ca_file is not None or ca_file_path
    has_wc_file = uploaded_wc_file is not None or wc_file_path
    
    if not has_ca_file and not has_wc_file:
        st.info("👆 Please upload at least one Excel file (CA or WC) to get started")
        st.markdown("""
        **Expected Files:**
        - **CA Excel**: Should contain sheets: HDVI_HITL_70, RISCOM_HITL_88, AllTrans_HITL_87
        - **WC Excel**: Should contain sheets: Method_HITL_67, Foresight_HITL_127
        """)
        return
    
    process_clicked = st.sidebar.button("🔄 Process Data", type="primary")
    
    if process_clicked or (not st.session_state.file_processed and (has_ca_file or has_wc_file)):
        with st.spinner("Processing Excel files..."):
            all_dataframes = []
            
            if has_ca_file:
                try:
                    if uploaded_ca_file:
                        ca_df = load_and_combine_excel_from_bytes(uploaded_ca_file.getvalue(), file_type='ca')
                    else:
                        ca_df = load_and_combine_excel(ca_file_path, file_type='ca')
                    
                    if not ca_df.empty:
                        all_dataframes.append(ca_df)
                        st.success(f"✅ CA File: Loaded {len(ca_df)} rows")
                    else:
                        st.warning("⚠️ CA File: No matching sheets found")
                except Exception as e:
                    st.error(f"❌ Error processing CA file: {str(e)}")
            
            if has_wc_file:
                try:
                    if uploaded_wc_file:
                        wc_df = load_and_combine_excel_from_bytes(uploaded_wc_file.getvalue(), file_type='wc')
                    else:
                        wc_df = load_and_combine_excel(wc_file_path, file_type='wc')
                    
                    if not wc_df.empty:
                        all_dataframes.append(wc_df)
                        st.success(f"✅ WC File: Loaded {len(wc_df)} rows")
                    else:
                        st.warning("⚠️ WC File: No matching sheets found")
                except Exception as e:
                    st.error(f"❌ Error processing WC file: {str(e)}")
            
            if all_dataframes:
                raw_df = pd.concat(all_dataframes, ignore_index=True)
                
                normalized_df = normalize_columns(raw_df)
                datetime_cols = ['creation_time', 'pick_time', 'making_start_time', 'making_end_time', 'completed_time']
                parsed_df = parse_datetimes(normalized_df, datetime_cols)
                enriched_df = compute_derived_metrics(parsed_df)
                
                st.session_state.processed_data = enriched_df
                st.session_state.file_processed = True
                
                st.success(f"✅ Total: {len(raw_df)} rows processed from {len(all_dataframes)} file(s)")
                
                if len(normalized_df) > 0:
                    with st.expander("🔍 Data Diagnostics"):
                        st.write(f"**Total rows after normalization:** {len(normalized_df)}")
                        st.write(f"**File types:** {enriched_df['file_type'].value_counts().to_dict() if 'file_type' in enriched_df.columns else 'N/A'}")
                        st.write(f"**Workflows:** {enriched_df['workflow'].value_counts().to_dict() if 'workflow' in enriched_df.columns else 'N/A'}")
                        st.write(f"**Clients:** {enriched_df['client'].value_counts().to_dict()}")
                        st.write(f"**Columns found:** {list(normalized_df.columns)}")
                        date_cols = ['creation_time', 'pick_time']
                        for col in date_cols:
                            if col in normalized_df.columns:
                                non_null_count = normalized_df[col].notna().sum()
                                st.write(f"**{col}:** {non_null_count}/{len(normalized_df)} non-null values")
            else:
                st.error("❌ No data could be loaded from the provided files")
                return
    
    if st.session_state.processed_data is None:
        st.warning("Click 'Process Data' to load and process your files")
        return
        
    enriched_df = st.session_state.processed_data
    
    if enriched_df.empty:
        st.warning("No data available after processing. Please check your files.")
        return
    
    today = pd.Timestamp.now().date()
    
    if analysis_period == "Last 7 days":
        start_date = today - timedelta(days=7)
        end_date = today
    elif analysis_period == "Last 30 days":
        start_date = today - timedelta(days=30)
        end_date = today
    elif analysis_period == "Last 90 days":
        start_date = today - timedelta(days=90)
        end_date = today
    elif analysis_period == "Custom":
        col1, col2 = st.sidebar.columns(2)
        with col1:
            start_date = st.date_input("Start Date", value=today - timedelta(days=30))
        with col2:
            end_date = st.date_input("End Date", value=today)
    else:
        available_dates = enriched_df['date'].dropna()
        if not available_dates.empty:
            start_date = available_dates.min()
            end_date = available_dates.max()
        else:
            start_date = today - timedelta(days=365)
            end_date = today
    
    start_dt = pd.to_datetime(start_date)
    end_dt = pd.to_datetime(end_date) + pd.Timedelta(days=1)
    
    has_dates = enriched_df['date'].notna().any()
    
    if has_dates and analysis_period != "All data":
        date_mask = (
            (enriched_df['date'] >= start_date) & (enriched_df['date'] <= end_date)
        )
        date_filtered_df = enriched_df[date_mask].copy()
        
        if date_filtered_df.empty and not enriched_df.empty:
            dates_in_range = enriched_df['date'].notna() & (
                (enriched_df['date'] >= start_date) & (enriched_df['date'] <= end_date)
            )
            if not dates_in_range.any():
                date_filtered_df = enriched_df.copy()
                st.info(f"⚠️ No data found in selected date range ({start_date} to {end_date}). Showing all available data.")
            else:
                date_mask_with_nan = date_mask | enriched_df['date'].isna()
                date_filtered_df = enriched_df[date_mask_with_nan].copy()
    else:
        date_filtered_df = enriched_df.copy()

    with st.sidebar:
        st.header("🔍 Filters")
        filter_source_df = date_filtered_df
        
        users = sorted(filter_source_df['making_user'].dropna().unique())
        st.subheader("👤 Select User(s)")
        selected_users = st.multiselect("Users", users, default=users if users else [], key="user_filter")
        
        if 'workflow' in filter_source_df.columns:
            workflows = sorted(filter_source_df['workflow'].dropna().unique())
            selected_workflows = st.multiselect("Workflows", workflows, default=workflows if workflows else [], key="workflow_filter")
        else:
            selected_workflows = []
        
        if 'file_type' in filter_source_df.columns:
            file_types = sorted(filter_source_df['file_type'].dropna().unique())
            selected_file_types = st.multiselect("File Types (CA/WC)", file_types, default=file_types if file_types else [], key="file_type_filter")
        else:
            selected_file_types = []
        
        clients = sorted(filter_source_df['client'].dropna().unique())
        selected_clients = st.multiselect("Clients", clients, default=clients if clients else [], key="client_filter")
        
        filtered_df = date_filtered_df.copy()
        if selected_users and len(selected_users) > 0:
            filtered_df = filtered_df[filtered_df['making_user'].isin(selected_users)]
        if selected_workflows and len(selected_workflows) > 0:
            filtered_df = filtered_df[filtered_df['workflow'].isin(selected_workflows)]
        if selected_file_types and len(selected_file_types) > 0:
            filtered_df = filtered_df[filtered_df['file_type'].isin(selected_file_types)]
        if selected_clients and len(selected_clients) > 0:
            filtered_df = filtered_df[filtered_df['client'].isin(selected_clients)]
    
    if filtered_df.empty:
        st.warning("No data available after filtering. Please adjust your filters.")
        return
    
    st.header("👥 User-Wise Dashboard")
    
    user_summary = create_user_summary_table(filtered_df)
    
    selected_users_list = sorted(filtered_df['making_user'].dropna().unique())
    if len(selected_users_list) == 1:
        selected_user = selected_users_list[0]
        st.subheader(f"📊 Detailed View: {selected_user}")
        create_user_detailed_view(selected_user, filtered_df)
    else:
        st.subheader("📊 User Summary Table")
        st.dataframe(user_summary, use_container_width=True)
        
        st.markdown("---")
        st.subheader("🔍 View Individual User Details")
        user_for_detail = st.selectbox("Select a user to view detailed information:", 
                                     options=[''] + selected_users_list, 
                                     key="user_detail_select")
        
        if user_for_detail:
            create_user_detailed_view(user_for_detail, filtered_df)
    
    st.markdown("---")
    st.subheader("Overview Metrics")
    col1, col2, col3, col4 = st.columns(4)
    
    completed_tasks = filtered_df[filtered_df['completed_time'].notna()]
    with col1:
        st.metric("Total Tasks", len(filtered_df))
    with col2:
        st.metric("Completed Tasks", len(completed_tasks))
    with col3:
        avg_duration = completed_tasks['task_duration_min'].mean() if not completed_tasks.empty else 0
        st.metric("Avg Duration", f"{avg_duration:.1f} min" if not pd.isna(avg_duration) else "N/A")
    with col4:
        avg_efficiency = completed_tasks['efficiency_ratio'].mean() * 100 if not completed_tasks.empty else 0
        st.metric("Avg Efficiency", f"{avg_efficiency:.1f}%" if not pd.isna(avg_efficiency) else "N/A")
    
    col1, col2 = st.columns(2)
    with col1:
        if 'workflow' in filtered_df.columns:
            st.subheader("Tasks by Workflow")
            workflow_counts = filtered_df['workflow'].value_counts()
            st.bar_chart(workflow_counts)
    
    with col2:
        if 'file_type' in filtered_df.columns:
            st.subheader("Tasks by File Type")
            file_type_counts = filtered_df['file_type'].value_counts()
            st.bar_chart(file_type_counts)
    
    st.markdown("---")
    st.subheader("Task Details")
    display_cols_data = ['task_id', 'making_user', 'workflow', 'file_type', 'client', 
                        'pick_time', 'completed_time', 'task_duration_min']
    display_cols_data = [col for col in display_cols_data if col in filtered_df.columns]
    st.dataframe(filtered_df[display_cols_data], use_container_width=True)
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        csv_data = filtered_df.to_csv(index=False)
        st.download_button("Download Full Data (CSV)", data=csv_data, 
                         file_name="user_metrics_data.csv", mime="text/csv")
    with col2:
        csv_summary = user_summary.to_csv(index=False)
        st.download_button("Download User Summary (CSV)", data=csv_summary,
                         file_name="user_summary.csv", mime="text/csv")

if __name__ == "__main__":
    run_user_wise_dashboard()