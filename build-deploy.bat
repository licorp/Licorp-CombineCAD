@echo off
setlocal enabledelayedexpansion

REM ============================================================
REM  Licorp CombineCAD - BUILD + DEPLOY (1 click)
REM  Build R2020-R2027 + AutoCAD, deploy:
REM    Revit:   %APPDATA%\Autodesk\Revit\Addins\{version}\
REM    AutoCAD: C:\ProgramData\Autodesk\ApplicationPlugins\
REM ============================================================

title Licorp CombineCAD - Build + Deploy
cd /d "%~dp0"
set "ROOT=%~dp0"
set "ACAD_BUNDLE=C:\ProgramData\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle"

echo.
echo  ========================================
echo   LICORP COMBINECAD - BUILD + DEPLOY
echo  ========================================
echo.

REM ========================================
REM  STEP 1: CLEAN
REM ========================================
echo [1] Cleaning previous builds...
dotnet clean Licorp_CombineCAD.sln -c Release --nologo -v q 2>nul
echo      Done.
echo.

REM ========================================
REM  STEP 2: BUILD ALL REVIT VERSIONS
REM ========================================
for %%v in (R2020 R2021 R2022 R2023 R2024 R2025 R2026 R2027) do (
    echo [2] Building Revit %%v...
    dotnet build src\Licorp_CombineCAD.%%v\Licorp_CombineCAD.%%v.csproj -c Release --nologo
    if errorlevel 1 (
        echo      [ERROR] %%v FAILED!
        goto :build_err
    )
)
echo      All 8 Revit versions built OK!
echo.

REM ========================================
REM  STEP 3: BUILD AUTOCAD PLUGIN
REM ========================================
echo [3] Building AutoCAD MergeSheets plugin...
dotnet build src.acad\Licorp_MergeSheets\Licorp_MergeSheets.csproj -c Release --nologo
if errorlevel 1 (
    echo      [WARNING] AutoCAD plugin build FAILED - skipping
    set "ACAD_SKIP=1"
) else (
    echo      AutoCAD plugin built OK!
    set "ACAD_SKIP=0"
)
echo.

REM ========================================
REM STEP 4: DEPLOY REVIT ADD-IN
REM Revit doc .addin tu: C:\ProgramData\Autodesk\ApplicationPlugins\
REM ========================================
echo [4] Deploying Revit add-in to C:\ProgramData\Autodesk\ApplicationPlugins\...
echo.

for %%v in (R2020 R2021 R2022 R2023 R2024 R2025 R2026 R2027) do (
    set "RV=%%v"
    set "RV=!RV:R=!"
    set "SRC=%ROOT%bin\%%v\Release"
    set "DST=C:\ProgramData\Autodesk\ApplicationPlugins\Licorp_CombineCAD"

    if not exist "!DST!" mkdir "!DST!"

    if exist "!SRC!\Licorp_CombineCAD.dll" (
        REM Copy ALL files (DLLs, deps.json, runtimes, etc.)
        xcopy "!SRC!\*" "!DST!\" /E /Y /Q /I >nul 2>&1

REM Create .addin manifest
    (
    echo ^<?xml version="1.0" encoding="utf-8" standalone="no"?^>
    echo ^<RevitAddIns^>
    echo ^<AddIn Type="Application"^>
    echo ^<Name^>Licorp CombineCAD^</Name^>
    echo ^<Assembly^>C:\ProgramData\Autodesk\ApplicationPlugins\Licorp_CombineCAD\Licorp_CombineCAD.dll^</Assembly^>
    echo ^<AddInId^>F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6^</AddInId^>
    echo ^<FullClassName^>Licorp_CombineCAD.App^</FullClassName^>
    echo ^<VendorId^>LICORP^</VendorId^>
    echo ^<VendorDescription^>Licorp - Export Revit sheets to DWG with multi-layout merge^</VendorDescription^>
    echo ^</AddIn^>
    echo ^</RevitAddIns^>
    ) > "!DST!\Licorp_CombineCAD.addin"

        echo       Revit !RV! - DLLs + .addin + libs
    ) else (
        echo       Revit !RV! - WARNING: Build output not found
    )
)

REM ========================================
REM  STEP 5: DEPLOY AUTOCAD PLUGIN
REM  AutoCAD doc tu: C:\ProgramData\Autodesk\ApplicationPlugins\
REM ========================================
echo.
echo [5] Deploying AutoCAD plugin to %ACAD_BUNDLE%...
echo.

if exist "%ACAD_BUNDLE%" rd /s /q "%ACAD_BUNDLE%" 2>nul
mkdir "%ACAD_BUNDLE%"

if "%ACAD_SKIP%"=="0" (
    set "ACAD_SRC=%ROOT%bin\acad\Release"

if exist "!ACAD_SRC!\Licorp_MergeSheets.dll" (
    REM Create Contents subfolder and copy DLLs
    if not exist "%ACAD_BUNDLE%\Contents" mkdir "%ACAD_BUNDLE%\Contents"
    xcopy "!ACAD_SRC!\Licorp_MergeSheets.dll" "%ACAD_BUNDLE%\Contents\" /Y /Q >nul 2>&1
    xcopy "!ACAD_SRC!\Newtonsoft.Json.dll" "%ACAD_BUNDLE%\Contents\" /Y /Q >nul 2>&1

    REM Copy PackageContents.xml from source
    copy /Y "!ROOT!src.acad\Licorp_MergeSheets\PackageContents.xml" "%ACAD_BUNDLE%\" >nul 2>&1

    echo AutoCAD - MergeSheets.dll + Newtonsoft.Json.dll + PackageContents.xml
)
) else (
    echo       AutoCAD - Skipped (build failed)
)

REM ========================================
REM  DONE!
REM ========================================
echo.
echo  ========================================
echo   BUILD + DEPLOY THANH CONG!
echo  ========================================
echo.
echo Revit add-in deployed to:
echo C:\ProgramData\Autodesk\ApplicationPlugins\Licorp_CombineCAD\
echo.
echo  AutoCAD plugin deployed to:
echo    %ACAD_BUNDLE%\
echo.
echo  Khoi dong lai Revit - Tab Licorp - Combine CAD
echo.
pause
exit /b 0

:build_err
echo.
echo  ========================================
echo   BUILD FAILED! Khong deploy.
echo  ========================================
echo.
pause
exit /b 1
