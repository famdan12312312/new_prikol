@echo off
echo ===================================================
echo   Starting Document Automation System
echo ===================================================
echo.

rem 1. Check virtual environment
if not exist venv (
    echo [1/4] Creating virtual environment venv...
    python -m venv venv
    if errorlevel 1 (
        echo Error: Failed to create venv. Make sure Python is installed and added to PATH.
        pause
        exit /b
    )
) else (
    echo [1/4] Virtual environment venv already exists.
)

rem 2. Activate virtual environment
echo [2/4] Activating venv...
call venv\Scripts\activate.bat

rem 3. Install requirements
echo [3/4] Installing dependencies from requirements.txt...
pip install -r requirements.txt
if errorlevel 1 (
    echo Error: Failed to install dependencies.
    pause
    exit /b
)

rem 4. Initialize database
echo [4/4] Initializing SQLite database...
python db_setup.py
if errorlevel 1 (
    echo Error: Failed to initialize database.
    pause
    exit /b
)

rem Run queries test
echo.
echo Running queries test:
python queries.py
echo.

echo ===================================================
echo   Launching Streamlit App...
echo ===================================================
python -m streamlit run app.py
pause
