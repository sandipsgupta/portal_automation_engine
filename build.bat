@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  Portal Automation Engine — Build Script
REM  Run this once from the project folder to generate the .exe
REM  Requirements: Python, venv activated, pyinstaller installed
REM ─────────────────────────────────────────────────────────────────────

echo.
echo  ====================================================
echo   Portal Automation Engine — PyInstaller Build
echo  ====================================================
echo.

REM Step 1 — Install / update dependencies
echo [1/4] Installing dependencies...
pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: pip install failed. Make sure venv is activated.
    pause
    exit /b 1
)

REM Step 2 — Install Playwright browsers (Chromium only)
echo [2/4] Installing Playwright Chromium browser...
playwright install chromium
if errorlevel 1 (
    echo ERROR: Playwright browser install failed.
    pause
    exit /b 1
)

REM Step 3 — Clean previous build
echo [3/4] Cleaning previous build...
if exist dist\PortalAutomationEngine.exe (
    del /q dist\PortalAutomationEngine.exe
)
if exist build (
    rmdir /s /q build
)

REM Step 4 — Build the .exe
echo [4/4] Building executable...
pyinstaller portal_engine.spec --noconfirm
if errorlevel 1 (
    echo ERROR: PyInstaller build failed. Check output above.
    pause
    exit /b 1
)

echo.
echo  ====================================================
echo   BUILD COMPLETE
echo   Output: dist\PortalAutomationEngine.exe
echo  ====================================================
echo.
echo  IMPORTANT — Before sending to client:
echo  1. Copy dist\PortalAutomationEngine.exe
echo  2. Copy .env (with credentials) to same folder
echo  3. Client double-clicks PortalAutomationEngine.exe
echo  4. Client uses Browse to select their daily sheet
echo.
pause