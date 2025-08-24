@echo off
echo ============================================
echo Setting up Capitaline Consolidator
echo ============================================

REM Step 1: Check if Python is available
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python is not installed or not on PATH.
    echo Please install Python 3.8+ from https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

REM Step 2: Create virtual environment (.venv)
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv
) else (
    echo Virtual environment already exists. Skipping creation.
)

REM Step 3: Activate venv and upgrade pip
call .venv\Scripts\activate
python -m pip install --upgrade pip

REM Step 4: Install required dependencies
pip install pandas openpyxl xlsxwriter

echo ============================================
echo Setup complete.
echo To run the script:
echo.
echo     call .venv\Scripts\activate
echo     python main.py
echo.
echo Or use debug mode:
echo     python main.py --debug
echo ============================================
pause
