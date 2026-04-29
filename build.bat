@echo off
REM ============================================================
REM Licorp_CombiCAD Build Script
REM Builds Revit add-in (R2020 + R2025 + R2027) + AutoCAD plugin
REM ============================================================

setlocal enabledelayedexpansion

echo.
echo ======================================
echo Licorp_CombiCAD Build Script
echo ======================================
echo.

cd /d "%~dp0"

REM ========================================
REM SECTION 1: Find AutoCAD Installation
REM ========================================
echo [Info] Checking AutoCAD installation...

set AUTOCAD_PATH=
set AUTOCAD_VER=

REM Check Registry for AutoCAD 2024
for /f "tokens=*" %%i in ('reg query "HKLM\SOFTWARE\Autodesk\AutoCAD\R24.0" /v AcadLocation 2^>nul') do (
    echo %%i | find "AcadLocation" >nul
    if !errorlevel!==0 (
        for /f "tokens=2*" %%a in ("%%i") do set AUTOCAD_PATH=%%b
        set AUTOCAD_VER=2024
    )
)

REM Check Registry for AutoCAD 2025 if 2024 not found
if not defined AUTOCAD_PATH (
    for /f "tokens=*" %%i in ('reg query "HKLM\SOFTWARE\Autodesk\AutoCAD\R25.0" /v AcadLocation 2^>nul') do (
        echo %%i | find "AcadLocation" >nul
        if !errorlevel!==0 (
            for /f "tokens=2*" %%a in ("%%i") do set AUTOCAD_PATH=%%b
            set AUTOCAD_VER=2025
        )
    )
)

REM Check common installation paths if registry not found
if not defined AUTOCAD_PATH (
    if exist "C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe" (
        set AUTOCAD_PATH=C:\Program Files\Autodesk\AutoCAD 2024
        set AUTOCAD_VER=2024
    ) else if exist "C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe" (
        set AUTOCAD_PATH=C:\Program Files\Autodesk\AutoCAD 2025
        set AUTOCAD_VER=2025
    )
)

if defined AUTOCAD_PATH (
    echo [Info] Found AutoCAD !AUTOCAD_VER! at: !AUTOCAD_PATH!
) else (
    echo [Warning] AutoCAD not detected - AutoCAD plugin will NOT be built
)

echo.

REM ========================================
REM SECTION 2: Clean Previous Build
REM ========================================
echo [Build] Cleaning previous builds...
dotnet clean src\Licorp_CombineCAD.R2020\Licorp_CombineCAD.R2020.csproj -c Release --nologo -v q 2>nul
dotnet clean src\Licorp_CombineCAD.R2025\Licorp_CombineCAD.R2025.csproj -c Release --nologo -v q 2>nul
dotnet clean src\Licorp_CombineCAD.R2027\Licorp_CombineCAD.R2027.csproj -c Release --nologo -v q 2>nul
echo [Clean] Done.

REM ========================================
REM SECTION 3: Build Revit R2020 (.NET Framework 4.8)
REM ========================================
echo.
echo [Build] Building Revit R2020 add-in (Revit 2020-2024)...
dotnet build src\Licorp_CombineCAD.R2020\Licorp_CombineCAD.R2020.csproj -c Release --nologo
if errorlevel 1 (
    echo.
    echo [ERROR] Revit R2020 build FAILED!
    goto :error
)
echo [Build] Revit R2020: bin\R2020\Release\Licorp_CombineCAD.dll

REM ========================================
REM SECTION 4: Build Revit R2025 (.NET 8)
REM ========================================
echo.
echo [Build] Building Revit R2025 add-in (Revit 2025-2026)...
dotnet build src\Licorp_CombineCAD.R2025\Licorp_CombineCAD.R2025.csproj -c Release --nologo
if errorlevel 1 (
    echo.
    echo [ERROR] Revit R2025 build FAILED!
    goto :error
)
echo [Build] Revit R2025: bin\R2025\Release\Licorp_CombineCAD.dll

REM ========================================
REM SECTION 5: Build Revit R2027 (.NET 8)
REM ========================================
echo.
echo [Build] Building Revit R2027 add-in (Revit 2027+)...
dotnet build src\Licorp_CombineCAD.R2027\Licorp_CombineCAD.R2027.csproj -c Release --nologo
if errorlevel 1 (
    echo.
    echo [ERROR] Revit R2027 build FAILED!
    goto :error
)
echo [Build] Revit R2027: bin\R2027\Release\Licorp_CombineCAD.dll

REM ========================================
REM SECTION 6: Build AutoCAD Plugin
REM ========================================
    echo.
    echo [Build] Building AutoCAD plugin (cross-compiling for .NET 4.8 and .NET 8.0)...

    dotnet build src.acad\Licorp_MergeSheets\Licorp_MergeSheets.csproj -c Release --nologo

    if errorlevel 1 (
        echo.
        echo [WARNING] AutoCAD plugin build FAILED
    ) else (
        echo [Build] AutoCAD plugin (2024): bin\acad\Release\net48\Licorp_MergeSheets.dll
        echo [Build] AutoCAD plugin (2025): bin\acad\Release\net8.0-windows\Licorp_MergeSheets.dll
    )

REM ========================================
REM SUCCESS
REM ========================================
echo.
echo ======================================
echo BUILD COMPLETED SUCCESSFULLY
echo ======================================
echo.
echo Output Files:
echo   R2020 (Revit 2020-2024): bin\R2020\Release\Licorp_CombineCAD.dll
echo   R2025 (Revit 2025-2026): bin\R2025\Release\Licorp_CombineCAD.dll
echo   R2027 (Revit 2027+):     bin\R2027\Release\Licorp_CombineCAD.dll
if defined AUTOCAD_PATH (
    echo   ACAD: bin\acad\Release\Licorp_MergeSheets.dll
)
echo.
echo Next Steps:
echo   1. Run deploy-bundle.bat to deploy to all Revit versions
echo   2. Restart Revit
echo   3. Tab Licorp -^> Combine CAD panel should appear
echo.
pause
exit /b 0

:error
echo.
echo ======================================
echo BUILD FAILED!
echo ======================================
echo.
pause
exit /b 1
