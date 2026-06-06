@echo off
setlocal

cd /d "%~dp0"

if not exist logs mkdir logs

echo ================================================== >> logs\console.log
echo Step 2 started %date% %time% >> logs\console.log

if not exist ".venv\Scripts\python.exe" (
    echo Creating Python virtual environment...
    py -3 -m venv .venv >> logs\setup.log 2>&1
    if errorlevel 1 (
        python -m venv .venv >> logs\setup.log 2>&1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment creation failed. Check logs\setup.log
    echo Virtual environment creation failed. Check logs\setup.log >> logs\console.log
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

python -m pip install --upgrade pip >> logs\setup.log 2>&1
python -m pip install -r requirements.txt >> logs\setup.log 2>&1
if errorlevel 1 (
    echo Dependency installation failed. Check logs\setup.log
    echo Dependency installation failed. Check logs\setup.log >> logs\console.log
    pause
    exit /b 1
)

python "2-6-2026_Final_Working.py" --mode full >> logs\console.log 2>&1
set EXIT_CODE=%errorlevel%

if not "%EXIT_CODE%"=="0" (
    echo Step 2 failed. Check logs\console.log and the latest logs\run_*.log
) else (
    echo Step 2 completed successfully.
)

echo Step 2 finished %date% %time% with exit code %EXIT_CODE% >> logs\console.log
pause
exit /b %EXIT_CODE%
