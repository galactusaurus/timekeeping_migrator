@echo off
chcp 65001 > nul

echo.
echo   +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo   ^|                                                           ^|
echo   ^|   ███████╗  ██████╗   ███████╗                            ^|
echo   ^|   ██╔══██║ ██╔════╝   ██╔════╝                            ^|
echo   ^|   ███████║ ██║  ███╗  █████╗                              ^|
echo   ^|   ██║  ██║ ██║   ██║  ██╔══╝                              ^|
echo   ^|   ██║  ██║ ╚██████╔╝  ███████╗                            ^|
echo   ^|   ╚═╝  ╚═╝  ╚═════╝   ╚══════╝                            ^|
echo   ^|                                                           ^|
echo   ^|          TIMEKEEPING MIGRATOR - CSV EXPORT TOOL           ^|
echo   ^|                     AGE ENGINEERING                       ^|
echo   ^|                                                           ^|
echo   +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+
echo.
echo.
echo Timekeeping Migrator - Automated CSV Export Script
echo This batch file automates the process of exporting Access data to SQLite and then to CSV

setlocal enabledelayedexpansion
cd /d "%~dp0"

REM ========================================
REM Capture Start Time (with seconds)
REM ========================================
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set start_date=%%c-%%a-%%b)
for /f "tokens=1-3 delims=:." %%a in ("!time!") do (set start_time=%%a:%%b:%%c)
set START_TIME=!start_date! !start_time!
set /a START_SECONDS=0
REM Convert time to seconds for duration calculation
for /f "tokens=1-3 delims=:." %%a in ("!time!") do (
    set /a START_SECONDS=%%a*3600+%%b*60+%%c
)

echo.
echo ========================================
echo Timekeeping Migrator - CSV Export Tool
echo ========================================
echo Started: !START_TIME!
echo.

REM ========================================
REM CONFIGURATION STAGE ----------------------------
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
REM - pandas: Data manipulation and Excel/CSV export
REM - pywin32: Windows COM support for Access database
REM - PyYAML: YAML config file parsing
REM - openpyxl: Excel file creation (required by pandas.to_excel)
set "PACKAGES=pandas pywin32 PyYAML openpyxl"

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
REM ---- END CONFIGURATION STAGE ----------------------------
REM ========================================

REM ========================================
REM EXTRACT STAGE ----------------------------
REM ========================================

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
REM END EXTRACT STAGE ----------------------------
REM ========================================

REM ========================================
REM TRANSFORM STAGE ----------------------------
REM ========================================

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
REM ========================================
REM Completion - Capture End Time and Calculate Duration
REM ========================================
for /f "tokens=2-4 delims=/ " %%a in ('date /t') do (set end_date=%%c-%%a-%%b)
for /f "tokens=1-3 delims=:." %%a in ("!time!") do (set end_time=%%a:%%b:%%c)
set END_TIME=!end_date! !end_time!

REM Calculate elapsed time in seconds
set /a END_SECONDS=0
for /f "tokens=1-3 delims=:." %%a in ("!time!") do (
    set /a END_SECONDS=%%a*3600+%%b*60+%%c
)

REM Handle day wraparound (if end time is before start time, add 24 hours)
if !END_SECONDS! LSS !START_SECONDS! (
    set /a END_SECONDS=!END_SECONDS!+86400
)

set /a DURATION=!END_SECONDS!-!START_SECONDS!
set /a HOURS=!DURATION!/3600
set /a MINUTES=(!DURATION! %%3600)/60
set /a SECONDS=!DURATION! %%60


REM ========================================
REM TRANSFORM STAGE ----------------------------
REM ========================================

echo.
echo ========================================
echo SUCCESS: Process completed!
echo ========================================
echo.
echo Started:  !START_TIME!
echo Ended:    !END_TIME!
echo Duration: !HOURS!h !MINUTES!m !SECONDS!s
echo.
echo Your CSV file has been generated in the project root directory
echo with a timestamp in the filename (results_YYYYMMDD_HHMMSS.csv)
echo.
pause
exit /b 0
