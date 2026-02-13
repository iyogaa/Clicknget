import os
import sys
import threading
import pandas as pd
import streamlit as st
import concurrent.futures
from typing import List, Tuple
from fuzzywuzzy import fuzz
from functools import lru_cache
import uuid
import io
import time

from stqdm import stqdm
from streamlit.runtime.scriptrunner import get_script_run_ctx, add_script_run_ctx
from pillm import litellmclient as lli

BATCH_SIZE = 20
MAX_RETRIES = 3

def categorize_accidents(descriptions, batch_start, system_prompt, filename, session_id, ctx=None, st_bar=None):
    if ctx:
        add_script_run_ctx(threading.currentThread(), ctx)
    
    # Prepare the input for GPT
    input_text = "SNo Description\n" + "\n".join(
        [f"{batch_start + i+1}. {desc[0]}" for i, desc in enumerate(descriptions)]
    )
    
    try:
        tags = [filename]
        prompt_name = "categorize_body_hierarchy_1_standard"
        response = lli.process_input(input_text, tags, prompt_name, session_id)
        if st_bar:
            st_bar.update(1)
        return response.content
    except Exception as e:
        # Logging error but not calling st.error here to avoid thread issues
        print(f"Error in API call for batch starting at {batch_start}: {str(e)}")
        return None

def process_descriptions(df_descriptions, column_name, system_prompt, filename, session_id, batch_size=20):
    all_results = {}
    total_rows = len(df_descriptions)
    batches = [(i, i + batch_size) for i in range(0, total_rows, batch_size)]
    remaining_batch_indices = list(range(len(batches)))
    
    ctx = get_script_run_ctx()
    st_bar = stqdm(total=len(batches), desc="Processing batches")

    for attempt in range(MAX_RETRIES):
        if not remaining_batch_indices:
            break
            
        if attempt > 0:
            st.info(f"Retrying {len(remaining_batch_indices)} failed batches (Attempt {attempt + 1}/{MAX_RETRIES})...")
            time.sleep(1) # Subtle delay before retry

        with concurrent.futures.ThreadPoolExecutor() as executor:
            future_to_batch_idx = {
                executor.submit(
                    categorize_accidents, 
                    [desc for desc in df_descriptions.iloc[batches[idx][0] : batches[idx][1]][[column_name]].itertuples(index=False, name=None)], 
                    batches[idx][0], 
                    system_prompt, 
                    filename,
                    session_id,
                    ctx, 
                    st_bar
                ): idx
                for idx in remaining_batch_indices
            }

            for future in concurrent.futures.as_completed(future_to_batch_idx):
                idx = future_to_batch_idx[future]
                try:
                    result = future.result()
                    if result:
                        all_results[idx] = result.strip("```").strip("json")
                except Exception as e:
                    print(f"Batch index {idx} generated an exception: {e}")

        # Update remaining batches
        remaining_batch_indices = [idx for idx in remaining_batch_indices if idx not in all_results]

    if remaining_batch_indices:
        st.error(f"Failed to process {len(remaining_batch_indices)} batches after {MAX_RETRIES} attempts.")
    
    # Return results in the original order to be safe, though SNo mapping is used later
    return [all_results[i] for i in sorted(all_results.keys())]

def parse_gpt_response(response):
    lines = response.strip().split("\n")
    parsed = []
    for line in lines:
        parts = line.split("|")
        if len(parts) >= 2:
            # Handle cases where GPT might include more pipes or extra whitespace
            sno = parts[0].strip().replace(".", "")
            category = parts[1].strip()
            parsed.append(
                {
                    "SNo": sno,
                    "Body Part - Hierarchy 1": category,
                }
            )
    return parsed

@lru_cache(maxsize=None)
def match_with_confidence(input_keyword: str, words: Tuple[str, ...], threshold: float = 0.9) -> List[Tuple[str, float]]:
    matched_words = []
    
    for word in words:
        if word == input_keyword:
            matched_words = [(word, 1.0)]
            break

        similarity = fuzz.ratio(input_keyword, word) / 100.0
        
        if similarity >= threshold:
            matched_words.append((word, similarity))
    
    matched_words.sort(key=lambda x: x[1], reverse=True)
    return matched_words

def process_and_parse_descriptions(df, column_name, system_prompt, filename, session_id, batch_size):
    results = process_descriptions(df, column_name, system_prompt, filename, session_id, batch_size=batch_size)

    all_categorized = []
    for result in results:
        try:
            parsed = parse_gpt_response(result)
            all_categorized.extend(parsed)
        except Exception as e:
            st.error(f"Error parsing GPT response: {str(e)}")

    # Create a mapping dictionary
    categorization_dict = {
        item["SNo"]: item["Body Part - Hierarchy 1"] for item in all_categorized
    }

    valid_categories = ("Head", "Neck", "Upper Extremities", "Trunk", "Lower Extremities", "Multiple Body Parts", "Misc")

    def get_valid_category(sno):
        category = categorization_dict.get(str(sno), "")
        if not category:
            return ""
        matches = match_with_confidence(category, valid_categories)
        return matches[0][0] if matches else category # Fallback to raw if no fuzzy match above threshold

    df["Body Part - Hierarchy 1"] = df["SNo"].apply(get_valid_category)
    return df

def run():
    # Page state management
    if "body_gpt_processed_df" not in st.session_state:
        st.session_state.body_gpt_processed_df = None
    if "body_gpt_current_file" not in st.session_state:
        st.session_state.body_gpt_current_file = None

    uploaded_file = st.file_uploader("Choose a CSV file", type="xlsx")
    
    if uploaded_file is not None:
        # Reset state if a new file is uploaded
        if st.session_state.body_gpt_current_file != uploaded_file.name:
            st.session_state.body_gpt_processed_df = None
            st.session_state.body_gpt_current_file = uploaded_file.name

        try:
            filename = uploaded_file.name
            xls = pd.ExcelFile(uploaded_file)
            df_input = pd.read_excel(uploaded_file, sheet_name="lossrun_data", dtype=str)
        except Exception as e:
            st.error(f"Error reading the Excel file: {str(e)}")
            return

        column_name = st.selectbox(
            "Select the column to process",
            options=["LossDescription", "ResultingInjuryDesc", "PartInjuredDesc"]
        )

        if column_name not in df_input.columns:
            st.error(f"The selected column '{column_name}' is not in the data.")
            return

        if "SNo" not in df_input.columns:
            df_input["SNo"] = [str(i) for i in range(1, len(df_input) + 1)]

        # Process button
        if st.button("Process"):
            session_id = str(uuid.uuid4())
            system_prompt = ""

            with st.status("Predicting Categories..."):
                processed_df = process_and_parse_descriptions(
                    df_input.copy(), 
                    column_name, 
                    system_prompt, 
                    filename, 
                    session_id, 
                    batch_size=BATCH_SIZE
                )
                st.session_state.body_gpt_processed_df = processed_df

        # Results display and download
        if st.session_state.body_gpt_processed_df is not None:
            df = st.session_state.body_gpt_processed_df
            st.write("### Categorized Results")
            st.dataframe(df)

            # Prepare downloadable file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet_name in xls.sheet_names:
                    if sheet_name == "lossrun_data":
                        # Save the processed one
                        df.drop(columns=["SNo"] if "SNo" in df.columns else []).to_excel(writer, index=False, sheet_name=sheet_name)
                    else:
                        original_df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
                        original_df.to_excel(writer, index=False, sheet_name=sheet_name)
            
            output.seek(0)
            processed_data = output.getvalue()

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download categorized data as Excel",
                    data=processed_data,
                    file_name=f"{os.path.splitext(filename)[0]}_categorized.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            with col2:
                csv = df.drop(columns=["SNo"] if "SNo" in df.columns else []).to_csv(index=False)
                st.download_button(
                    label="Download categorized data as CSV",
                    data=csv,
                    file_name=f"{os.path.splitext(filename)[0]}_categorized.csv",
                    mime="text/csv",
                )

