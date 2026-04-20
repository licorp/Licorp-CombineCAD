# Deploy script - Build va deploy Licorp_CombineCAD
# Chay: .\deploy-bundle.ps1

$ErrorActionPreference = "Continue"
$rootDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Join-Path $rootDir "src"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " LICORP COMBINECAD - DEPLOY" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Build R2020 (Revit 2020-2024, net48)
Write-Host "`n[1/4] Building R2020 (Revit 2020-2024)..." -ForegroundColor Yellow
dotnet build "$sourceDir\Licorp_CombineCAD.R2020\Licorp_CombineCAD.R2020.csproj" -c Debug
if ($LASTEXITCODE -ne 0) { Write-Host "Build R2020 FAILED!" -ForegroundColor Red; exit 1 }
Write-Host "Build R2020 OK!" -ForegroundColor Green

# Build R2025 (Revit 2025-2026, net8.0)
Write-Host "`n[2/4] Building R2025 (Revit 2025-2026)..." -ForegroundColor Yellow
dotnet build "$sourceDir\Licorp_CombineCAD.R2025\Licorp_CombineCAD.R2025.csproj" -c Debug
if ($LASTEXITCODE -ne 0) { Write-Host "Build R2025 FAILED!" -ForegroundColor Red; exit 1 }
Write-Host "Build R2025 OK!" -ForegroundColor Green

# Build R2027 (Revit 2027, net8.0)
Write-Host "`n[3/4] Building R2027 (Revit 2027)..." -ForegroundColor Yellow
dotnet build "$sourceDir\Licorp_CombineCAD.R2027\Licorp_CombineCAD.R2027.csproj" -c Debug
if ($LASTEXITCODE -ne 0) { Write-Host "Build R2027 FAILED!" -ForegroundColor Red; exit 1 }
Write-Host "Build R2027 OK!" -ForegroundColor Green

# Build AutoCAD MergeSheets plugin
Write-Host "`n[4/4] Building AutoCAD MergeSheets plugin..." -ForegroundColor Yellow
dotnet build "$sourceDir\..\src.acad\Licorp_MergeSheets\Licorp_MergeSheets.csproj" -c Debug
if ($LASTEXITCODE -ne 0) { Write-Host "Build MergeSheets FAILED!" -ForegroundColor Red; exit 1 }
Write-Host "Build MergeSheets OK!" -ForegroundColor Green

# Deploy to output folder
Write-Host "`nDeploying to output..." -ForegroundColor Yellow

$addinContent = @"
<?xml version="1.0" encoding="utf-8" standalone="no"?>
<RevitAddIns>
<AddIn Type="Application">
<Name>Licorp CombineCAD</Name>
<Assembly>Licorp_CombineCAD.dll</Assembly>
<AddInId>F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6</AddInId>
<FullClassName>Licorp_CombineCAD.App</FullClassName>
<VendorId>LICORP</VendorId>
<VendorDescription>Licorp - Export Revit sheets to DWG with multi-layout merge</VendorDescription>
</AddIn>
</RevitAddIns>
"@

# R2020 -> Revit 2020-2024 (net48)
# R2025 -> Revit 2025-2026 (net8.0)
# R2027 -> Revit 2027+ (net8.0)
$versions = @(
    @{ Year = "2020"; Source = "R2020" },
    @{ Year = "2021"; Source = "R2020" },
    @{ Year = "2022"; Source = "R2020" },
    @{ Year = "2023"; Source = "R2020" },
    @{ Year = "2024"; Source = "R2020" },
    @{ Year = "2025"; Source = "R2025" },
    @{ Year = "2026"; Source = "R2025" },
    @{ Year = "2027"; Source = "R2027" }
)

foreach ($v in $versions) {
    $project = $v.Source
    $year = $v.Year
    $srcDir = Join-Path $rootDir "bin\$project\Debug"
    $destDir = "$env:APPDATA\Autodesk\Revit\Addins\$year"

    Write-Host " Deploying Revit $year ($project)..." -ForegroundColor Cyan
    Write-Host " Source: $srcDir" -ForegroundColor Gray

    if (-not (Test-Path $srcDir)) {
        Write-Host " ERROR: Source not found!" -ForegroundColor Red
        continue
    }

    if (-not (Test-Path $destDir)) {
        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
    }

    $dllCount = 0
    Get-ChildItem "$srcDir\*.dll" | Where-Object { $_.Name -notlike "*.resources.dll" } | ForEach-Object {
        Copy-Item $_.FullName $destDir -Force
        $dllCount++
    }

    Get-ChildItem "$srcDir\runtimes" -Recurse -Filter "*.dll" -ErrorAction SilentlyContinue | ForEach-Object {
        $runtimeDest = Join-Path $destDir "runtimes\$($_.Directory.Name)\"
        if (-not (Test-Path $runtimeDest)) {
            New-Item -ItemType Directory -Path $runtimeDest -Force | Out-Null
        }
        Copy-Item $_.FullName $runtimeDest -Force
        $dllCount++
    }

    $addinContent | Out-File "$destDir\Licorp_CombineCAD.addin" -Encoding utf8 -NoNewline

    Write-Host " Copied $dllCount files + .addin manifest" -ForegroundColor Green
    Write-Host " Revit $year OK" -ForegroundColor White
}

# Deploy AutoCAD MergeSheets plugin to C:\ProgramData\Autodesk\ApplicationPlugins
$bundleName = "Licorp_MergeSheets.bundle"
$bundleDir = Join-Path "C:\ProgramData\Autodesk\ApplicationPlugins" $bundleName
$mergeSheetsSrc = Join-Path $rootDir "bin\acad\Debug"

Write-Host "`n Deploying AutoCAD MergeSheets plugin..." -ForegroundColor Cyan
Write-Host " Source: $mergeSheetsSrc" -ForegroundColor Gray
Write-Host " Target: $bundleDir" -ForegroundColor Gray

if (Test-Path $mergeSheetsSrc) {
    if (-not (Test-Path $bundleDir)) {
        New-Item -ItemType Directory -Path $bundleDir -Force | Out-Null
    }

    $bundleDll = Get-ChildItem "$mergeSheetsSrc\Licorp_MergeSheets.dll" -ErrorAction SilentlyContinue
    if ($bundleDll) {
        Copy-Item $bundleDll.FullName $bundleDir -Force
        Write-Host " Copied Licorp_MergeSheets.dll" -ForegroundColor Green
    }
    else {
        Write-Host " WARNING: Licorp_MergeSheets.dll not found in build output" -ForegroundColor Yellow
    }

    $pkgContent = @"
<?xml version="1.0" encoding="utf-8"?>
<ApplicationPackage SchemaVersion="1.0" AppVersion="1.0.0" ProductCode="{B8C9D4E5-1F2A-3B4C-5D6E-7F8A9B0C1D2E}" Name="Licorp MergeSheets" Description="Merge multiple DWG layouts into a single DWG">
<CompanyDetails Name="Licorp" />
<Components>
<ComponentEntry AppName="Licorp MergeSheets" ModuleName="./Licorp_MergeSheets.dll" LoadOnAutoCADStartup="True" />
</Components>
</ApplicationPackage>
"@
    $pkgContent | Out-File "$bundleDir\PackageContents.xml" -Encoding utf8 -NoNewline
    Write-Host " Created PackageContents.xml" -ForegroundColor Green

    Get-ChildItem "$mergeSheetsSrc\*.dll" | Where-Object { $_.Name -ne "Licorp_MergeSheets.dll" -and $_.Name -notlike "*.resources.dll" } | ForEach-Object {
        Copy-Item $_.FullName $bundleDir -Force
        Write-Host " Copied $($_.Name)" -ForegroundColor Gray
    }

    Write-Host " MergeSheets bundle deployed OK" -ForegroundColor White
}
else {
    Write-Host " WARNING: MergeSheets build output not found" -ForegroundColor Yellow
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host " DEPLOY THANH CONG!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "`nRevit 2020-2024: net48 (R2020)" -ForegroundColor Cyan
Write-Host "Revit 2025-2026: net8.0 (R2025)" -ForegroundColor Cyan
Write-Host "Revit 2027+: net8.0 (R2027)" -ForegroundColor Cyan
Write-Host "`nRevit Addin: $env:APPDATA\Autodesk\Revit\Addins\{version}\" -ForegroundColor Yellow
Write-Host "AutoCAD Bundle: C:\ProgramData\Autodesk\ApplicationPlugins\$bundleName\" -ForegroundColor Yellow
Read-Host "`nNhan Enter de dong"
