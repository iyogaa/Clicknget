# How to Run the App

## Development Mode

1. **Activate Virtual Environment**:
   ```powershell
   .\venv\Scripts\activate
   ```

2. **Install Dependencies**:
   ```powershell
   pip install -r requirements.txt
   ```

3. **Run Streamlit**:
   ```powershell
   streamlit run app.py
   ```

## Production Mode

1. **Environment Setup**:
   Ensure Python 3.9+ is installed and isolate dependencies in `venv`.

2. **Configuration**:
   - Create `.streamlit/secrets.toml` with: lead
     ```toml
     [credentials]
     # Your secure credentials here
     ```
   - Ensure `server.port` is configured in `.streamlit/config.toml` or via CLI.

3. **Start Server**:
   ```bash
   streamlit run app.py --server.port 8080 --browser.serverAddress 0.0.0.0
   ```

4. **Monitoring**:
   - Logs are printed to stdout/stderr.
   - Use a process manager like `supervisord` or `systemd` to keep the app running.

## Directory Structure

- **app.py**: Main entry point and router.
- **features/**: Modular logic for each tool (MVR, HDVI, PDF, AI).
- **utils/**: Shared utilities (auth, styles).
- **pages_legacy/**: Old page files (safe to verify against but not used by main app).
- **constants.py**: REMOVED. Use `st.secrets` or `.streamlit/secrets.toml`.
