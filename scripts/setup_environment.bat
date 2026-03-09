@echo off
REM ============================================================================
REM CLICKNGET - Clean Environment Setup Script
REM This script completely rebuilds the Python virtual environment
REM ============================================================================

echo.
echo ========================================
echo CLICKNGET Environment Setup
echo ========================================
echo.

REM Step 1: Remove old virtual environment
echo [1/5] Removing old virtual environment...
if exist venv (
    rmdir /s /q venv
    echo     - Old venv removed
) else (
    echo     - No old venv found
)

if exist .venv (
    rmdir /s /q .venv
    echo     - Old .venv removed
) else (
    echo     - No old .venv found
)

echo.

REM Step 2: Create fresh virtual environment
echo [2/5] Creating fresh virtual environment...
python -m venv venv
if %errorlevel% neq 0 (
    echo ERROR: Failed to create virtual environment
    echo Please ensure Python 3.10+ is installed
    pause
    exit /b 1
)
echo     - Virtual environment created successfully
echo.

REM Step 3: Activate virtual environment
echo [3/5] Activating virtual environment...
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)
echo     - Virtual environment activated
echo.

REM Step 4: Upgrade pip
echo [4/5] Upgrading pip to latest version...
python -m pip install --upgrade pip
echo     - pip upgraded successfully
echo.

REM Step 5: Install dependencies
echo [5/5] Installing dependencies from requirements.txt...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
echo     - All dependencies installed successfully
echo.

echo ========================================
echo Setup Complete!
echo ========================================
echo.
echo To run the application:
echo   1. Activate environment: venv\Scripts\activate
echo   2. Run app: streamlit run app.py
echo.
pause
