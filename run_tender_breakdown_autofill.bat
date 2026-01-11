@echo off
REM ====================================================
REM  AUTOFILL LAUNCHER
REM ====================================================

REM 1. Ensure we are running in the project directory
cd /d "%~dp0"

REM 2. Activate the virtual environment
REM    (Checks if the environment exists first to prevent errors)
if exist "myenv\Scripts\activate.bat" (
    call myenv\Scripts\activate
) else (
    echo [ERROR] Virtual environment 'myenv' not found!
    echo Please make sure 'myenv' exists in this folder.
    pause
    exit /b
)

REM 3. Run the Python script
echo Starting Autofill Process...
echo ------------------------------------------
python src\autofill.py

REM 4. Pause execution so you can read the report
echo.
echo ------------------------------------------
echo Process Finished.
pause