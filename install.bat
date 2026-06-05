@echo off
setlocal enabledelayedexpansion
title CDR Forensic Analysis Tool v3.1 - Installer
mode con: cols=80 lines=45
color 0F

:: ============================================================
:: BANNER
:: ============================================================
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════════════╗
echo  ║                                                                    ║
echo  ║     ██████╗██████╗ ██████╗     ████████╗ ██████╗  ██████╗ ██╗     ║
echo  ║    ██╔════╝██╔══██╗██╔══██╗    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ║
echo  ║    ██║     ██║  ██║██████╔╝       ██║   ██║   ██║██║   ██║██║     ║
echo  ║    ██║     ██║  ██║██╔══██╗       ██║   ██║   ██║██║   ██║██║     ║
echo  ║    ╚██████╗██████╔╝██║  ██║       ██║   ╚██████╔╝╚██████╔╝███████╗║
echo  ║     ╚═════╝╚═════╝ ╚═╝  ╚═╝       ╚═╝    ╚═════╝  ╚═════╝╚══════╝║
echo  ║                                                                    ║
echo  ║          FORENSIC ANALYSIS TOOL v3.1 - LAW ENFORCEMENT             ║
echo  ║                    i9/JAS Compatible Format                        ║
echo  ║                                                                    ║
echo  ╚══════════════════════════════════════════════════════════════════════╝
echo.
echo   Installer will set up everything needed to run the CDR Analysis Tool.
echo.
echo   ┌──────────────────────────────────────────────────────────────────┐
echo   │  INSTALLATION STEPS:                                            │
echo   │                                                                  │
echo   │   [1] Check / Install Python                                    │
echo   │   [2] Install Required Packages (pandas, openpyxl)             │
echo   │   [3] Create Desktop Shortcut                                   │
echo   │   [4] Launch Tool                                               │
echo   │                                                                  │
echo   └──────────────────────────────────────────────────────────────────┘
echo.
pause

:: ============================================================
:: STEP 1 - PYTHON CHECK
:: ============================================================
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════════════╗
echo  ║  STEP 1 of 4  :  CHECKING PYTHON INSTALLATION                     ║
echo  ╚══════════════════════════════════════════════════════════════════════╝
echo.
echo   Checking if Python is installed...
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo   ┌──────────────────────────────────────────────────────────────────┐
    echo   │  [!!] Python is NOT installed                                   │
    echo   └──────────────────────────────────────────────────────────────────┘
    echo.
    echo   Downloading Python 3.12 installer... Please wait.
    echo.

    powershell -Command "Write-Host '   Downloading...' -NoNewline; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%TEMP%\python_installer.exe'; Write-Host ' Done!'"

    if not exist "%TEMP%\python_installer.exe" (
        echo.
        echo   ┌──────────────────────────────────────────────────────────────────┐
        echo   │  [FAILED] Could not download Python automatically.              │
        echo   │                                                                  │
        echo   │  Please install Python 3.10+ manually from:                     │
        echo   │  https://www.python.org/downloads/                              │
        echo   │                                                                  │
        echo   │  >>> IMPORTANT: Check "Add Python to PATH" during install! <<<  │
        echo   │                                                                  │
        echo   │  Then run this installer again.                                 │
        echo   └──────────────────────────────────────────────────────────────────┘
        echo.
        pause
        exit /b 1
    )

    echo   Installing Python 3.12...
    echo   (This will take 1-2 minutes, please wait)
    echo.
    echo   ░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░  Installing...
    
    "%TEMP%\python_installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1

    set "PATH=%LOCALAPPDATA%\Programs\Python\Python312;%LOCALAPPDATA%\Programs\Python\Python312\Scripts;%PATH%"

    timeout /t 3 >nul

    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo.
        echo   ┌──────────────────────────────────────────────────────────────────┐
        echo   │  [!!] Python installed but PATH not updated yet.                │
        echo   │                                                                  │
        echo   │  Please RESTART your computer and run this installer again.     │
        echo   └──────────────────────────────────────────────────────────────────┘
        echo.
        pause
        exit /b 1
    )

    echo.
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do (
        echo   ┌──────────────────────────────────────────────────────────────────┐
        echo   │  [OK] %%i installed successfully!                       │
        echo   └──────────────────────────────────────────────────────────────────┘
    )
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do (
        echo   ┌──────────────────────────────────────────────────────────────────┐
        echo   │  [OK] %%i detected!                                     │
        echo   └──────────────────────────────────────────────────────────────────┘
    )
)

echo.
echo   Step 1 complete. Proceeding...
timeout /t 2 >nul

:: ============================================================
:: STEP 2 - INSTALL PACKAGES
:: ============================================================
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════════════╗
echo  ║  STEP 2 of 4  :  INSTALLING PYTHON PACKAGES                       ║
echo  ╚══════════════════════════════════════════════════════════════════════╝
echo.
echo   Installing required packages...
echo.
echo   ┌──────────────────────────────────────────────────────────────────┐
echo   │  Package          Status                                        │
echo   │  ─────────────    ──────────────────────────────                │

:: Upgrade pip silently
pip install --upgrade pip >nul 2>&1

:: Install pandas
echo   │  pandas           Installing...                                 │
pip install pandas >nul 2>&1
if %errorlevel% neq 0 (
    pip install --user pandas >nul 2>&1
)

:: Install openpyxl
echo   │  openpyxl         Installing...                                 │
pip install openpyxl >nul 2>&1
if %errorlevel% neq 0 (
    pip install --user openpyxl >nul 2>&1
)

echo   └──────────────────────────────────────────────────────────────────┘
echo.

:: Verify
echo   Verifying installation...
echo.
python -c "import pandas; print('   [OK] pandas         v' + pandas.__version__)"
python -c "import openpyxl; print('   [OK] openpyxl       v' + openpyxl.__version__)"

if %errorlevel% neq 0 (
    echo.
    echo   ┌──────────────────────────────────────────────────────────────────┐
    echo   │  [FAILED] Package installation failed.                          │
    echo   │                                                                  │
    echo   │  Try running manually in Command Prompt:                        │
    echo   │     pip install pandas openpyxl                                 │
    echo   └──────────────────────────────────────────────────────────────────┘
    echo.
    pause
    exit /b 1
)

echo.
echo   ┌──────────────────────────────────────────────────────────────────┐
echo   │  [OK] All packages installed successfully!                      │
echo   └──────────────────────────────────────────────────────────────────┘
echo.
echo   Step 2 complete. Proceeding...
timeout /t 2 >nul

:: ============================================================
:: STEP 3 - DESKTOP SHORTCUT
:: ============================================================
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════════════╗
echo  ║  STEP 3 of 4  :  CREATING DESKTOP SHORTCUT                        ║
echo  ╚══════════════════════════════════════════════════════════════════════╝
echo.

set "TOOL_DIR=%~dp0"

:: Create launcher bat
echo   Creating launcher script...
(
    echo @echo off
    echo title CDR Forensic Analysis Tool v3.1
    echo cd /d "%TOOL_DIR%"
    echo python gui.py
    echo if %%errorlevel%% neq 0 ^(
    echo     echo.
    echo     echo   [ERROR] Tool failed to start. Check Python installation.
    echo     pause
    echo ^)
) > "%TOOL_DIR%CDR_Tool.bat"

:: Create Desktop shortcut
echo   Creating Desktop shortcut...
echo.

powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'CDR Analysis Tool.lnk')); $s.TargetPath = '%TOOL_DIR%CDR_Tool.bat'; $s.WorkingDirectory = '%TOOL_DIR%'; $s.IconLocation = 'shell32.dll,21'; $s.Description = 'CDR Forensic Analysis Tool v3.1 - Law Enforcement Edition'; $s.Save()"

if exist "%USERPROFILE%\Desktop\CDR Analysis Tool.lnk" (
    echo   ┌──────────────────────────────────────────────────────────────────┐
    echo   │  [OK] Desktop shortcut created!                                 │
    echo   │                                                                  │
    echo   │   Look for "CDR Analysis Tool" icon on your Desktop            │
    echo   └──────────────────────────────────────────────────────────────────┘
) else (
    echo   ┌──────────────────────────────────────────────────────────────────┐
    echo   │  [!!] Shortcut creation failed (permissions issue)              │
    echo   │                                                                  │
    echo   │   You can still run the tool by double-clicking:               │
    echo   │   %TOOL_DIR%CDR_Tool.bat
    echo   └──────────────────────────────────────────────────────────────────┘
)

echo.
echo   Step 3 complete. Proceeding...
timeout /t 2 >nul

:: ============================================================
:: STEP 4 - DONE
:: ============================================================
cls
echo.
echo  ╔══════════════════════════════════════════════════════════════════════╗
echo  ║                                                                    ║
echo  ║     ██████╗██████╗ ██████╗     ████████╗ ██████╗  ██████╗ ██╗     ║
echo  ║    ██╔════╝██╔══██╗██╔══██╗    ╚══██╔══╝██╔═══██╗██╔═══██╗██║     ║
echo  ║    ██║     ██║  ██║██████╔╝       ██║   ██║   ██║██║   ██║██║     ║
echo  ║    ██║     ██║  ██║██╔══██╗       ██║   ██║   ██║██║   ██║██║     ║
echo  ║    ╚██████╗██████╔╝██║  ██║       ██║   ╚██████╔╝╚██████╔╝███████╗║
echo  ║     ╚═════╝╚═════╝ ╚═╝  ╚═╝       ╚═╝    ╚═════╝  ╚═════╝╚══════╝║
echo  ║                                                                    ║
echo  ╚══════════════════════════════════════════════════════════════════════╝
echo.
echo   ┌──────────────────────────────────────────────────────────────────┐
echo   │                                                                  │
echo   │          INSTALLATION COMPLETED SUCCESSFULLY!                   │
echo   │                                                                  │
echo   ├──────────────────────────────────────────────────────────────────┤
echo   │                                                                  │
echo   │   [OK] Python             Installed                             │
echo   │   [OK] pandas             Installed                             │
echo   │   [OK] openpyxl           Installed                             │
echo   │   [OK] Desktop Shortcut   Created                              │
echo   │                                                                  │
echo   ├──────────────────────────────────────────────────────────────────┤
echo   │                                                                  │
echo   │   HOW TO USE:                                                   │
echo   │                                                                  │
echo   │    1. Double-click "CDR Analysis Tool" on Desktop               │
echo   │    2. Select your CDR CSV file                                  │
echo   │    3. Click "GENERATE FORENSIC REPORT"                          │
echo   │    4. Report opens automatically!                               │
echo   │                                                                  │
echo   ├──────────────────────────────────────────────────────────────────┤
echo   │                                                                  │
echo   │   Tool Location: %TOOL_DIR%
echo   │                                                                  │
echo   └──────────────────────────────────────────────────────────────────┘
echo.
echo.

set /p LAUNCH="   Launch the tool now? (Y/N): "
if /i "%LAUNCH%"=="Y" (
    echo.
    echo   Starting CDR Analysis Tool...
    cd /d "%TOOL_DIR%"
    start "" python gui.py
)

echo.
echo   You can close this window now.
echo.
pause
