# Build Script for Licorp_CombiCAD
# Builds Revit add-in (R2020-R2027) and AutoCAD plugin

$ErrorActionPreference = "Stop"

$ProjectRoot = $PSScriptRoot
$SolutionFile = Join-Path $ProjectRoot "Licorp_CombineCAD.sln"

Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host " Licorp_CombiCAD Build Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (!$dotnet) {
    Write-Host "ERROR: .NET SDK not found. Please install .NET 8 SDK." -ForegroundColor Red
    exit 1
}

Write-Host "Using: $($dotnet.Source)" -ForegroundColor Gray
Write-Host ""

# Clean
Write-Host "[Build] Cleaning..." -ForegroundColor Cyan
dotnet clean $SolutionFile -c Release --nologo -v q
Write-Host ""

# Build all Revit versions (2020-2024: .NET Framework 4.8)
$net48Versions = @("R2020", "R2021", "R2022", "R2023", "R2024")
foreach ($ver in $net48Versions) {
    $revitYear = $ver.Substring(1)
    Write-Host "[Build] Building $ver (Revit $revitYear, .NET Framework 4.8)..." -ForegroundColor Cyan
    $csproj = Join-Path $ProjectRoot "src\Licorp_CombineCAD.$ver\Licorp_CombineCAD.$ver.csproj"
    dotnet build $csproj -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host " FAILED: $ver" -ForegroundColor Red
        exit 1
    }
    Write-Host ""
}

# Build all Revit versions (2025-2026: .NET 8)
$net8Versions = @("R2025", "R2026")
foreach ($ver in $net8Versions) {
    $revitYear = $ver.Substring(1)
    Write-Host "[Build] Building $ver (Revit $revitYear, .NET 8)..." -ForegroundColor Cyan
    $csproj = Join-Path $ProjectRoot "src\Licorp_CombineCAD.$ver\Licorp_CombineCAD.$ver.csproj"
    dotnet build $csproj -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        Write-Host " FAILED: $ver" -ForegroundColor Red
        exit 1
    }
    Write-Host ""
}

# Build Revit R2027 (.NET 10) - Revit 2027+
Write-Host "[Build] Building R2027 (Revit 2027+, .NET 10)..." -ForegroundColor Cyan
$csproj2027 = Join-Path $ProjectRoot "src\Licorp_CombineCAD.R2027\Licorp_CombineCAD.R2027.csproj"
dotnet build $csproj2027 -c Release --nologo
if ($LASTEXITCODE -ne 0) {
    Write-Host " FAILED: R2027" -ForegroundColor Red
    exit 1
}
Write-Host ""

# Build AutoCAD Plugin
Write-Host "[Build] Building AutoCAD Plugin..." -ForegroundColor Cyan
$acadProject = Join-Path $ProjectRoot "src.acad\Licorp_MergeSheets\Licorp_MergeSheets.csproj"
if (Test-Path $acadProject) {
    dotnet build $acadProject -c Release --nologo
    if ($LASTEXITCODE -ne 0) { exit 1 }
} else {
    Write-Host "  AutoCAD project not found, skipping..." -ForegroundColor Yellow
}
Write-Host ""

Write-Host "Build completed successfully!" -ForegroundColor Green
Write-Host ""
Write-Host "Output locations:" -ForegroundColor Cyan
Write-Host "  R2020 (Revit 2020): $ProjectRoot\bin\R2020\Release\" -ForegroundColor White
Write-Host "  R2021 (Revit 2021): $ProjectRoot\bin\R2021\Release\" -ForegroundColor White
Write-Host "  R2022 (Revit 2022): $ProjectRoot\bin\R2022\Release\" -ForegroundColor White
Write-Host "  R2023 (Revit 2023): $ProjectRoot\bin\R2023\Release\" -ForegroundColor White
Write-Host "  R2024 (Revit 2024): $ProjectRoot\bin\R2024\Release\" -ForegroundColor White
Write-Host "  R2025 (Revit 2025): $ProjectRoot\bin\R2025\Release\" -ForegroundColor White
Write-Host "  R2026 (Revit 2026): $ProjectRoot\bin\R2026\Release\" -ForegroundColor White
Write-Host "  R2027 (Revit 2027): $ProjectRoot\bin\R2027\Release\" -ForegroundColor White
Write-Host "  AutoCAD Plugin:     $ProjectRoot\bin\acad\Release\" -ForegroundColor White
Write-Host ""
