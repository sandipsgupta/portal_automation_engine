@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  Portal Automation Engine — Build Script
REM  Run from project root with venv activated: .\build.bat
REM ─────────────────────────────────────────────────────────────────────

echo.
echo  ====================================================
echo   Portal Automation Engine -- Build
echo  ====================================================
echo.

REM ── Step 1: Install / update dependencies ────────────────────────────
echo [1/3] Installing dependencies...
python -m pip install -r requirements.txt -q
if errorlevel 1 ( echo ERROR: pip install failed. & pause & exit /b 1 )

REM ── Step 2: Clean previous build ─────────────────────────────────────
echo [2/3] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

REM ── Step 3: Build .exe with PyInstaller ──────────────────────────────
echo [3/3] Building PortalAutomationEngine.exe...
python -m PyInstaller PortalAutomationEngine.spec --noconfirm
if errorlevel 1 ( echo ERROR: PyInstaller build failed. & pause & exit /b 1 )

echo.
echo  ====================================================
echo   BUILD COMPLETE
echo   Output: dist\PortalAutomationEngine.exe
echo  ====================================================
echo.
echo  NEXT -- Create Windows Installer:
echo    1. Open Inno Setup (jrsoftware.org if not installed)
echo    2. Open installer.iss
echo    3. Press F9 to compile
echo    4. Output: installer\Setup_PortalEngine_v1.0.0.exe
echo.
pause
