import os
import threading
import pandas as pd
import streamlit as st
import concurrent.futures
from typing import List, Tuple
from fuzzywuzzy import fuzz
from functools import lru_cache
import uuid
import io

from stqdm import stqdm
from streamlit.runtime.scriptrunner import get_script_run_ctx, add_script_run_ctx
from pillm import litellmclient as lli




# Initialize OpenAI client
BATCH_SIZE = 20

def categorize_accidents(descriptions, batch_start, system_prompt, filename, session_id, ctx=None, st_bar=None):
    if ctx:
        add_script_run_ctx(threading.currentThread(), ctx)
    # Prepare the input for GPT
    input_text = "SNo Description\n" + "\n".join(
        [f"{batch_start + i+1}. {desc[0]}" for i, desc in enumerate(descriptions)]
    )

    try:
        tags = [filename]
        prompt_name = "categorize_cause_hierarchy_1_standard"
        response = lli.process_input(input_text, tags, prompt_name, session_id)
        st_bar.update(1)
        return response.content
    except Exception as e:
        st.error(f"Error in API call: {str(e)}")
        return None


def process_descriptions(df_descriptions, column_name, system_prompt, filename, session_id, batch_size=20, st_bar=None):
    results = []
    total_batches = (len(df_descriptions) + batch_size - 1) // batch_size  # Ceiling division
    st_bar = stqdm(total=total_batches, desc="Processing batches")
    ctx = get_script_run_ctx()

    with concurrent.futures.ThreadPoolExecutor() as executor:
        future_to_batch = {
            executor.submit(
                categorize_accidents, 
                [desc for desc in df_descriptions.iloc[i : i + batch_size][[column_name]].itertuples(index=False, name=None)], 
                i, 
                system_prompt, 
                filename,
                session_id,
                ctx, 
                st_bar
            ): i
            for i in range(0, len(df_descriptions), batch_size)
        }

        for future in concurrent.futures.as_completed(future_to_batch):
            batch_index = future_to_batch[future]
            try:
                result = future.result()
                if result:
                    result = result.strip("```").strip("json")
                    results.append(result)
            except Exception as e:
                print(f"Batch at index {batch_index} generated an exception: {e}")
    return results


def parse_gpt_response(response):
    lines = response.strip().split("\n")
    parsed = []
    for line in lines:
        parts = line.split("|")
        if len(parts) == 2:
            parsed.append(
                {
                    "SNo": parts[0].strip(),
                    "Cause - Hierarchy 1": parts[1].strip(),
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

@st.cache_data
def process_and_parse_descriptions(df, column_name, system_prompt, filename, session_id, batch_size):
    # Process descriptions
    results = process_descriptions(df, column_name, system_prompt, filename, session_id, batch_size=batch_size)

    # Parse and combine results
    all_categorized = []
    for result in results:
        try:
            all_categorized.extend(parse_gpt_response(result))
            all_categorized = [
                {**item, "SNo": item["SNo"].replace(".", "").strip()} for item in all_categorized
            ]
        except Exception as e:
            st.error(f"Error parsing GPT response: {str(e)} for result {result}")

    # Create a dictionary mapping descriptions to categories
    categorization_dict = {
        item["SNo"]: item["Cause - Hierarchy 1"] for item in all_categorized
    }

    valid_categories = (
        "Burn or Scald - Heat or Cold Exposures - Contact With",
        "Includes Freezing",
        "Caught In, Under or Between",
        "Cut, Puncture, Scrape Injured by",
        "Fall, Slip or Trip Injury",
        "Motor Vehicle",
        "Strain or Injury by",
        "Striking Against or Stepping on",
        "Struck or Injured by",
        "Rubbed or Abraded by",
        "Misc"
    )

    def get_valid_category(sno):
        category = categorization_dict.get(sno, "")
        matches = match_with_confidence(category, valid_categories)
        return matches[0][0] if matches else ""

    # Add the AccidentCategory column to the dataframe
    df["Cause - Hierarchy 1"] = df["SNo"].map(get_valid_category)

    df = df.drop(columns=["SNo"])
    return df

def run():
    # File uploader
    uploaded_file = st.file_uploader("Choose a CSV file", type="xlsx")
    if uploaded_file is not None:
        # Read the Excel file
        try:
            session_id = str(uuid.uuid4())
            xls = pd.ExcelFile(uploaded_file)
            filename = uploaded_file.name
            df = pd.read_excel(uploaded_file, sheet_name="lossrun_data", dtype=str)
        except Exception as e:
            st.error(f"Error reading the Excel file: {str(e)}")
            return

        # Select column to process
        column_name = st.selectbox(
            "Select the column to process",
            options=["LossDescription", "ResultingInjuryDesc", "PartInjuredDesc"]
        )

        if column_name not in df.columns:
            st.error(f"The selected column '{column_name}' is not in the data.")
            return

        if "SNo" not in df.columns:
            df["SNo"] = [str(i) for i in range(1, len(df) + 1)]

        # Process button
        if st.button("Process"):
            # Read the prompt from the file
            # with open("cause_gpt/prompt.txt", "r") as file:
            #     system_prompt = file.read()
            system_prompt=""

            # Process descriptions
            with st.status("Predicting Categories"):
                df = process_and_parse_descriptions(df, column_name, system_prompt, filename, session_id, batch_size=BATCH_SIZE)

            # Display the results
            st.write(df)
            # Convert the DataFrame to an Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for sheet_name in xls.sheet_names:
                    original_df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
                    original_df.to_excel(writer, index=False, sheet_name=sheet_name)
                
                df.to_excel(writer, index=False, sheet_name='lossrun_data')

            output.seek(0)  
            processed_data = output.getvalue()

            # Download button for the updated Excel file
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="Download categorized data as Excel",
                    data=processed_data,
                    file_name=os.path.splitext(uploaded_file.name)[0] + "_categorized.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with col2:
                csv = df.to_csv(index=False)
                st.download_button(
                    label="Download categorized data as CSV",
                    data=csv,
                    file_name=os.path.splitext(uploaded_file.name)[0] + ".csv",
                    mime="text/csv",
                )


if __name__ == "__main__":
    run()
