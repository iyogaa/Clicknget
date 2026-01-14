# How to Run the Clicknget Application

## Quick Start Command

```powershell
python -m streamlit run streamlit_app.py
```

## Alternative Commands

### Using the virtual environment directly:
```powershell
.\.venv\Scripts\streamlit.exe run streamlit_app.py
```

### Using Python from virtual environment:
```powershell
.\.venv\Scripts\python.exe -m streamlit run streamlit_app.py
```

## Access the Application

Once running, open your browser and go to:
- **Local URL:** http://localhost:8501
- **Network URL:** http://10.129.183.154:8501 (accessible from other devices on your network)

## Default Login Credentials

### Admin User
- **Username:** admin
- **Password:** admin123
- **Access:** Full access to all features

### QA User
- **Username:** qa
- **Password:** qa123
- **Access:** All features

### Maker User
- **Username:** maker
- **Password:** maker123
- **Access:** Limited features

> ⚠️ **Security Note:** Change these default passwords in `.streamlit/secrets.toml` for production use!

## Stop the Application

Press `Ctrl + C` in the terminal to stop the server.

## Troubleshooting

### If you get "Python was not found":
1. Make sure Python is installed
2. Use the virtual environment command: `.\.venv\Scripts\streamlit.exe run streamlit_app.py`

### If you get "streamlit not found":
```powershell
pip install -r requirements.txt
```

### If the port is already in use:
```powershell
python -m streamlit run streamlit_app.py --server.port 8502
```

## New Feature: DOB Validation

The app now automatically validates Driver Date of Birth (DOB) fields:
- ✅ Detects invalid/corrupted DOB values (e.g., "XX XX X003")
- ✅ Automatically fetches correct DOB from Client Excel
- ✅ All existing functionality preserved
