@echo off
echo ========================================
echo Video Clip Cutter Setup
echo ========================================
echo.

echo [1/4] Setting up Python virtual environment...
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
    if %ERRORLEVEL% neq 0 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
) else (
    echo Virtual environment already exists
)

echo Activating virtual environment...
call venv\Scripts\activate.bat
if %ERRORLEVEL% neq 0 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)

echo Installing requirements...
pip install -r requirements.txt
if %ERRORLEVEL% neq 0 (
    echo WARNING: Failed to install requirements, continuing...
)

echo.
echo [2/4] Installing FFmpeg via winget...
winget install --id=Gyan.FFmpeg -e --accept-package-agreements --accept-source-agreements
if %ERRORLEVEL% neq 0 (
    echo FFmpeg already installed or installation failed, continuing...
)

echo.
echo [3/4] Copying ffmpeg.exe to current directory...
for /f "tokens=*" %%i in ('where ffmpeg') do (
    set FFMPEG_PATH=%%i
    goto :found
)
:found
if defined FFMPEG_PATH (
    copy "%FFMPEG_PATH%" .
    echo Successfully copied ffmpeg.exe to current directory
) else (
    echo ERROR: Could not find ffmpeg.exe after installation
    pause
    exit /b 1
)

echo.
echo [4/4] Verifying installation...
ffmpeg -version >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo FFmpeg installation successful!
    echo.
    echo You can now run the clip cutter script:
    echo python clip_cutter.py
) else (
    echo ERROR: FFmpeg verification failed
    pause
    exit /b 1
)

echo.
echo Setup complete!
echo Note: Virtual environment is activated. Run 'venv\Scripts\activate.bat' to activate it again later.
pause
