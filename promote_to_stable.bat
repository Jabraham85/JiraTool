@echo off
setlocal enabledelayedexpansion

echo ============================================
echo   Promote Experimental Release to Stable
echo ============================================
echo.

:: Check that version_experimental.json exists
if not exist "version_experimental.json" (
    echo ERROR: version_experimental.json not found.
    echo Run this from the project root directory.
    pause
    exit /b 1
)

:: Read experimental version using PowerShell (handles JSON properly)
for /f "usebackq delims=" %%V in (`powershell -NoProfile -Command "(Get-Content 'version_experimental.json' | ConvertFrom-Json).version"`) do set "EXP_VER=%%V"
for /f "usebackq delims=" %%U in (`powershell -NoProfile -Command "(Get-Content 'version_experimental.json' | ConvertFrom-Json).download_url"`) do set "EXP_URL=%%U"
for /f "usebackq delims=" %%C in (`powershell -NoProfile -Command "(Get-Content 'version_experimental.json' | ConvertFrom-Json).changelog"`) do set "EXP_LOG=%%C"

:: Read current stable version
for /f "usebackq delims=" %%S in (`powershell -NoProfile -Command "(Get-Content 'version.json' | ConvertFrom-Json).version"`) do set "STABLE_VER=%%S"

echo Current stable:        v%STABLE_VER%
echo Experimental to promote: v%EXP_VER%
echo.

if "%EXP_VER%"=="%STABLE_VER%" (
    echo WARNING: Experimental and stable are already the same version.
)

echo This will:
echo   1. Copy version_experimental.json to version.json
echo   2. Commit and push to GitHub
echo.

set /p CONFIRM="Proceed? (y/n): "
if /i not "%CONFIRM%"=="y" (
    echo Cancelled.
    pause
    exit /b 0
)

echo.
echo Copying experimental manifest to stable...
copy /y "version_experimental.json" "version.json" >nul

echo Staging changes...
git add version.json

echo Committing...
git commit -m "Promote experimental v%EXP_VER% to stable"

echo Pushing to GitHub...
git push origin main

echo.
echo ============================================
echo   Done! Stable is now v%EXP_VER%
echo ============================================
echo.
echo Users on the Stable channel will now receive this update.
pause
