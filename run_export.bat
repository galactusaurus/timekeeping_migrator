@echo off
REM Timekeeping Migrator - Automated CSV Export Script
REM This batch file automates the process of exporting Access data to SQLite and then to CSV

setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ========================================
echo Timekeeping Migrator - CSV Export Tool
echo ========================================
echo.

REM ========================================
REM Step 1: Verify Python is installed
REM ========================================
echo [1/7] Checking for Python installation...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please contact your administrator to install Python 3.8 or higher
    echo and add it to your system PATH.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo %PYTHON_VERSION% found

REM ========================================
REM Step 2: Verify Pip is installed
REM ========================================
echo.
echo [2/7] Checking for Pip installation...
pip --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Pip is not installed or not in PATH
    echo.
    echo Please contact your administrator to install Pip
    echo and add it to your system PATH.
    echo.
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('pip --version') do set PIP_VERSION=%%i
echo %PIP_VERSION% found

REM ========================================
REM Step 3: Setup Virtual Environment
REM ========================================
echo.
echo [3/7] Setting up Python virtual environment...
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo Virtual environment created
) else (
    echo Virtual environment already exists
)

REM Activate virtual environment
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)
echo Virtual environment activated

REM ========================================
REM Step 4: Install Required Packages
REM ========================================
echo.
echo [4/7] Installing required Python packages...
echo This may take a moment...
echo.

REM List of packages to install (excluding built-in modules)
set "PACKAGES=pandas pywin32 PyYAML"

for %%p in (%PACKAGES%) do (
    echo Installing %%p...
    pip install %%p --quiet
    if errorlevel 1 (
        echo WARNING: Failed to install %%p
    )
)

echo Packages installation complete

REM ========================================
REM Step 5: Verify Config File
REM ========================================
echo.
echo [5/7] Verifying configuration...
if not exist "config.yaml" (
    echo ERROR: config.yaml not found in project root
    echo Please ensure config.yaml exists with the required parameters
    pause
    exit /b 1
)
echo config.yaml found

REM ========================================
REM Step 6: Run Export to SQLite
REM ========================================
echo.
echo [6/7] Exporting Access database to SQLite...
echo This may take several minutes depending on data volume...
echo.

python scripts\export_to_sqlite.py --filter-project > "%temp%\export_output.txt" 2>&1
set EXPORT_EXIT_CODE=%errorlevel%

REM Display export output
type "%temp%\export_output.txt"

REM Check for success message
findstr /M "Output saved to:" "%temp%\export_output.txt" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Export to SQLite failed
    echo Please review the output above for details
    echo.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Export to SQLite completed successfully

REM ========================================
REM Step 7: Run Query to CSV
REM ========================================
echo.
echo [7/7] Generating CSV report from latest database...
echo.

python scripts\query_to_csv.py --latest > "%temp%\query_output.txt" 2>&1
set QUERY_EXIT_CODE=%errorlevel%

REM Display query output
type "%temp%\query_output.txt"

if !QUERY_EXIT_CODE! neq 0 (
    echo.
    echo ERROR: CSV generation failed
    echo Please review the output above for details
    echo.
    pause
    exit /b 1
)

REM ========================================
REM Completion
REM ========================================
echo.
echo ========================================
echo SUCCESS: Process completed!
echo ========================================
echo.
echo Your CSV file has been generated in the project root directory
echo with a timestamp in the filename (results_YYYYMMDD_HHMMSS.csv)
echo.
pause
exit /b 0
