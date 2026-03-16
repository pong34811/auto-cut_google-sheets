@echo off
echo ========================================
echo Video Clip Cutter Runner
echo ========================================
echo.

echo Activating virtual environment...
call venv\Scripts\activate.bat
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to activate virtual environment
    echo Please run setup.bat first
    pause
    exit /b 1
)

echo.
echo Running clip cutter...
python clip_cutter.py

echo.
echo Done!
pause
