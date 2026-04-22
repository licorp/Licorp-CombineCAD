# Licorp_CombiCAD Deploy Script
# Deploys Revit add-in and AutoCAD plugin to ProgramData

param(
    [switch]$Uninstall,
    [switch]$AutoCADOnly,
    [switch]$RevitOnly
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path $PSScriptRoot -Parent
$BinFolder = Join-Path $ProjectRoot "bin"

function Deploy-RevitAddin {
    Write-Host "[Deploy] Installing Revit add-in..." -ForegroundColor Cyan

    $dllSource = Join-Path $BinFolder "R2025\Release\Licorp_CombineCAD.dll"
    $addinSource = Join-Path $PSScriptRoot "Licorp_CombineCAD.addin"

    $bundleDest = "$env:PROGRAMDATA\Autodesk\ApplicationPlugins\Licorp_CombineCAD"
    $wrongBundle = "$env:PROGRAMDATA\Autodesk\ApplicationPlugins\Licorp_CombineCAD.bundle"

    # Clean up wrong folder if exists
    if (Test-Path $wrongBundle) {
        Write-Host " Cleaning up wrong folder: $wrongBundle" -ForegroundColor Yellow
        Remove-Item $wrongBundle -Recurse -Force
    }

    if (!$Uninstall) {
        if (!(Test-Path $bundleDest)) {
            New-Item -ItemType Directory -Path $bundleDest -Force | Out-Null
        }

        if (Test-Path $dllSource) {
            Copy-Item $dllSource -Destination $bundleDest -Force
            Write-Host " DLL: $dllSource" -ForegroundColor Gray
        } else {
            Write-Host " DLL not found: $dllSource (build R2025 project first)" -ForegroundColor Red
            return
        }

        if (Test-Path $addinSource) {
            Copy-Item $addinSource -Destination $bundleDest -Force
            Write-Host " ADDIN: $addinSource" -ForegroundColor Gray
        }

        Write-Host " Installed: $bundleDest" -ForegroundColor Green
    } else {
        if (Test-Path $bundleDest) {
            Remove-Item $bundleDest -Recurse -Force
            Write-Host " Uninstalled: $bundleDest" -ForegroundColor Yellow
        }
    }
}

function Deploy-AutoCADPlugin {
    Write-Host "[Deploy] Installing AutoCAD plugin..." -ForegroundColor Cyan

    $dllSource = Join-Path $BinFolder "acad\Release\Licorp_MergeSheets.dll"
    $bundleSource = Join-Path $ProjectRoot "src.acad\Licorp_MergeSheets"

    $bundleDest = "$env:PROGRAMDATA\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle"
    $contentsDest = Join-Path $bundleDest "Contents"

    # Clean up wrong folder if exists (in case old structure exists)
    $wrongPath = "$env:PROGRAMDATA\Autodesk\ApplicationPlugins\Licorp_MergeSheets"
    if (Test-Path $wrongPath -PathType Container) {
        $items = Get-ChildItem $wrongPath -Force
        if ($items.Count -eq 0 -or ($items | Where-Object { $_.Name -eq "Contents" }).Count -eq 0) {
            Write-Host " Cleaning up wrong folder: $wrongPath" -ForegroundColor Yellow
            Remove-Item $wrongPath -Recurse -Force
        }
    }

    if (!$Uninstall) {
        if (!(Test-Path $bundleDest)) {
            New-Item -ItemType Directory -Path $bundleDest -Force | Out-Null
        }

        if (!(Test-Path $contentsDest)) {
            New-Item -ItemType Directory -Path $contentsDest -Force | Out-Null
        }

if (Test-Path $dllSource) {
        Copy-Item $dllSource -Destination $contentsDest -Force
        Write-Host " DLL: $dllSource" -ForegroundColor Gray
    } else {
        Write-Host " DLL not found: $dllSource (build acad project first)" -ForegroundColor Red
        return
    }

    $newtonsoftSource = Join-Path $BinFolder "acad\Release\Newtonsoft.Json.dll"
    if (Test-Path $newtonsoftSource) {
        Copy-Item $newtonsoftSource -Destination $contentsDest -Force
        Write-Host " Newtonsoft.Json.dll: $newtonsoftSource" -ForegroundColor Gray
    } else {
        Write-Host " Newtonsoft.Json.dll not found: $newtonsoftSource" -ForegroundColor Yellow
    }

        $packageContents = Join-Path $bundleSource "PackageContents.xml"
        if (Test-Path $packageContents) {
            Copy-Item $packageContents -Destination $bundleDest -Force
            Write-Host " PackageContents.xml: $packageContents" -ForegroundColor Gray
        }

        Write-Host " Installed: $bundleDest" -ForegroundColor Green
    } else {
        if (Test-Path $bundleDest) {
            Remove-Item $bundleDest -Recurse -Force
            Write-Host " Uninstalled: $bundleDest" -ForegroundColor Yellow
        }
    }
}

# Main
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Licorp_CombiCAD Deploy Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

if ($Uninstall) {
    Write-Host "Mode: UNINSTALL" -ForegroundColor Yellow
} else {
    Write-Host "Mode: INSTALL" -ForegroundColor Green
}
Write-Host ""
Write-Host "Target: $env:PROGRAMDATA\Autodesk\ApplicationPlugins" -ForegroundColor Gray
Write-Host ""

if ($RevitOnly) {
    Deploy-RevitAddin
} elseif ($AutoCADOnly) {
    Deploy-AutoCADPlugin
} else {
    Deploy-RevitAddin
    Write-Host ""
    Deploy-AutoCADPlugin
}

Write-Host ""
Write-Host "Done!" -ForegroundColor Cyan