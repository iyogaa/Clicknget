# Testing Instructions

This document outlines how to verify that each feature of the Clicknget app works correctly after the refactoring.

## Prerequisites

1.  **Environment**: Ensure you are in the correct Python environment.
    ```bash
    # Activate venv if needed
    .\venv\Scripts\activate
    ```
2.  **Dependencies**: Install updated requirements.
    ```bash
    pip install -r requirements.txt
    ```
3.  **Secrets**: Verify `.streamlit/secrets.toml` exists with valid credentials.
    - Default credentials for testing:
      - **Admin**: `admin` / `123`
      - **User**: `user` / `user`

## Running the App

Start the app locally:
```bash
streamlit run app.py
```

## Feature Verification Checklist

### 1. Authentication
- [ ] Launch the app. You should see a login sidebar.
- [ ] Try logging in with `admin` / `123`.
- [ ] Verify you see the full menu with all GPT tools.
- [ ] Log out and try logging in with `user` / `user`.
- [ ] Verify you see a restricted menu (MVR tools only, no GPT).

### 2. MVR All Trans
- [ ] Navigate to **MVR All Trans**.
- [ ] Upload a sample MVR file (`.xlsx`) and a Client Lookup file (`.xlsx`).
- [ ] Click **Process**.
- [ ] Verify that a `Final_Report.xlsx` is generated and downloadable.

### 3. HDVI-MVR
- [ ] Navigate to **HDVI-MVR**.
- [ ] Upload a Client Excel/CSV.
- [ ] Upload an Output Excel containing an `MVR` sheet.
- [ ] Click **Generate HDVI Report**.
- [ ] Verify the success message and download the output.

### 4. PDF Tools (Maker & Play)
- [ ] Navigate to **PDF Maker**.
- [ ] Upload a PDF and click **Process PDF**. Verify the flattened PDF download.
- [ ] Navigate to **PDF Play**.
- [ ] Test **Word → PDF** with a sample `.docx`.
- [ ] Test **Merge PDFs** by uploading two PDFs and merging them.

### 5. AI Features (Body/Cause/Accident/Custom GPT)
- [ ] Navigate to **Body GPT** (requires Admin/QA role).
- [ ] Upload a sample `.xlsx` with a `lossrun_data` sheet.
- [ ] Select a column (e.g., `LossDescription`).
- [ ] Click **Process**.
- [ ] **Note**: Since this uses `pillm` and `litellm`, ensure you have the necessary API keys or local model server running. If `pillm` is missing, you may need to install it from your private source.

## Troubleshooting

- **ModuleNotFoundError**: If you see `No module named 'pillm'`, ensure the package is installed. If `pillm` is a local folder not in the repo, add it to `PYTHONPATH` or copy it to the root.
- **Constants Error**: If you see errors related to `constants`, ensure you have migrated all specific config to `.streamlit/secrets.toml`.
- **"Connection refused" for AI**: Ensure your LLM backend is reachable.

## Production Deployment

1.  Set `debug=false` in `.streamlit/config.toml` (if it exists).
2.  Ensure `secrets.toml` is **NOT** committed to public version control if it contains real keys. Use environment variables or a secure secrets manager in production.
3.  Run with:
    ```bash
    streamlit run app.py --server.port 8501 --server.address 0.0.0.0
    ```
