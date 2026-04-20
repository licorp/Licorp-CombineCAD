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
set "ACAD_BUNDLE=C:\ProgramData\Autodesk\ApplicationPlugins\Licorp_CombineCAD.bundle"

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
REM  STEP 4: DEPLOY REVIT ADD-IN
REM  Revit doc .addin tu: %APPDATA%\Autodesk\Revit\Addins\{version}\
REM  Mapping: R2020->2020, R2021->2021, ... R2027->2027
REM ========================================
echo [4] Deploying Revit add-in to %APPDATA%\Autodesk\Revit\Addins\...
echo.

for %%v in (R2020 R2021 R2022 R2023 R2024 R2025 R2026 R2027) do (
    set "RV=%%v"
    set "RV=!RV:R=!"
    set "SRC=%ROOT%bin\%%v\Release"
    set "DST=%APPDATA%\Autodesk\Revit\Addins\!RV!"

    if not exist "!DST!" mkdir "!DST!"

    if exist "!SRC!\Licorp_CombineCAD.dll" (
        REM Copy ALL files (DLLs, deps.json, runtimes, etc.)
        xcopy "!SRC!\*" "!DST!\" /E /Y /Q /I >nul 2>&1

        REM Create .addin manifest
        (
            echo ^<?xml version="1.0" encoding="utf-8" standalone="no"?^>
            echo ^<RevitAddIns^>
            echo   ^<AddIn Type="Application"^>
            echo     ^<Name^>Licorp CombineCAD^</Name^>
            echo     ^<Assembly^>!DST!\Licorp_CombineCAD.dll^</Assembly^>
            echo     ^<AddInId^>F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6^</AddInId^>
            echo     ^<FullClassName^>Licorp_CombineCAD.App^</FullClassName^>
            echo     ^<VendorId^>LICORP^</VendorId^>
            echo     ^<VendorDescription^>Licorp - Export Revit sheets to DWG with multi-layout merge^</VendorDescription^>
            echo   ^</AddIn^>
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
        xcopy "!ACAD_SRC!\*" "%ACAD_BUNDLE%\" /E /Y /Q /I >nul 2>&1

        (
            echo ^<?xml version="1.0" encoding="utf-8"?^>
            echo ^<ApplicationPackage SchemaVersion="1.0" AppVersion="1.0.0"
            echo   ProductCode="{B8C9D4E5-1F2A-3B4C-5D6E-7F8A9B0C1D2E}"
            echo   Name="Licorp CombineCAD"
            echo   Description="Merge multiple DWG layouts into a single DWG"^>
            echo   ^<CompanyDetails Name="Licorp" /^>
            echo   ^<Components^>
            echo     ^<ComponentEntry AppName="Licorp MergeSheets"
            echo       ModuleName="./Licorp_MergeSheets.dll"
            echo       LoadOnAutoCADStartup="True" /^>
            echo   ^</Components^>
            echo ^</ApplicationPackage^>
        ) > "%ACAD_BUNDLE%\PackageContents.xml"

        echo       AutoCAD - MergeSheets.dll + libs + PackageContents.xml
    ) else (
        echo       AutoCAD - WARNING: DLL not found
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
echo  Revit add-in deployed to:
for %%v in (2020 2021 2022 2023 2024 2025 2026 2027) do (
    echo    %%v: %APPDATA%\Autodesk\Revit\Addins\%%v\
)
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
