# Excel-to-PDF Master 🚀

A production-grade Streamlit application for converting Excel files to high-quality PDFs, optimized for Streamlit Cloud.

## Features
- **Clean Architecture**: Decoupled core logic, utilities, and UI.
- **Batch Processing**: Upload multiple files and download them as a ZIP.
- **Premium Design**: Sleek dark theme with responsive UI.
- **Robust Error Handling**: Custom exceptions and structured logging.
- **Secure File Handling**: Context-managed temporary files and size validation.
- **PDF Preview**: Automatic first-page preview of converted files.

## Setup & Deployment

### Local Development
1. Clone the repository.
2. Navigate to the `app/` directory.
3. Install dependencies: `pip install -r requirements.txt`.
4. Run the app: `streamlit run app.py`.

### Streamlit Cloud Deployment
1. Push the code to a GitHub repository.
2. Connect the repository to Streamlit Cloud.
3. The `packages.txt` and `requirements.txt` will automatically handle system and Python dependencies.

## Project Structure
```text
app/
├── .streamlit/
│   └── config.toml         # Theme and server configuration
├── src/
│   ├── config/
│   │   └── settings.py     # Centralized app configuration
│   ├── core/
│   │   ├── converter.py    # Main Excel-to-PDF logic
│   │   ├── exceptions.py   # Custom exception classes
│   │   └── validators.py   # File validation logic
│   └── utils/
│       ├── file_handler.py # Temporary file management
│       └── logger.py       # Structured JSON logging
├── app.py                  # Streamlit entry point
├── packages.txt            # System dependencies (LibreOffice, Fonts)
├── requirements.txt        # Python dependencies
└── runtime.txt             # Python version specification
```
