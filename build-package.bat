@echo off
setlocal

cd /d "%~dp0"

set "VERSION=1.0.0"
set "SKIPBUILD="

if /I "%~1"=="-SkipBuild" set "SKIPBUILD=-SkipBuild"
if /I "%~1"=="skip" set "SKIPBUILD=-SkipBuild"
if not "%~2"=="" set "VERSION=%~2"

echo.
echo ======================================
echo Licorp_CombineCAD Package Script
echo ======================================
echo Version: %VERSION%
if defined SKIPBUILD (
    echo Mode: Skip build, package from existing bin outputs
) else (
    echo Mode: Build then package
)
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-package.ps1" -Version "%VERSION%" %SKIPBUILD%
if errorlevel 1 (
    echo.
    echo ======================================
    echo PACKAGE FAILED
    echo ======================================
    echo.
    pause
    exit /b 1
)

echo.
echo ======================================
echo PACKAGE COMPLETED
echo ======================================
echo.
echo ZIP:
echo   %~dp0artifacts\Licorp_CombineCAD_Setup_%VERSION%.zip
echo.
echo INSTALLER FOLDER:
echo   %~dp0artifacts\release\%VERSION%\installer
echo.
pause
exit /b 0
