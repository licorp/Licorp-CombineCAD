# Licorp_CombiCAD Deploy Script
# Deploys Revit add-in and AutoCAD plugin

param(
    [switch]$Uninstall,
    [switch]$AutoCADOnly,
    [switch]$RevitOnly
)

$ErrorActionPreference = "Stop"

$ProjectRoot = $PSScriptRoot
$DeployFolder = Join-Path $ProjectRoot "deploy"
$BinFolder = Join-Path $ProjectRoot "bin"

function Deploy-RevitAddin {
    Write-Host "[Deploy] Installing Revit add-in..." -ForegroundColor Cyan

    $addinSource = Join-Path $DeployFolder "Licorp_CombineCAD.addin"
    $dllSource = Join-Path $BinFolder "R2025\Release\Licorp_CombineCAD.dll"

    # Find Revit addins folder
    $revitVersions = @("2026", "2025", "2024", "2023", "2022", "2021", "2020")
    $installedRevit = @()

    foreach ($ver in $revitVersions) {
        $addinsPath = "$env:APPDATA\Autodesk\REVIT\Addins\$ver"
        if (Test-Path $addinsPath) {
            $installedRevit += @{ Version = $ver; Path = $addinsPath }
            Write-Host "  Found Revit $ver at $addinsPath"
        }
    }

    if ($installedRevit.Count -eq 0) {
        Write-Host "[Deploy] No Revit installation found!" -ForegroundColor Red
        return
    }

    foreach ($revit in $installedRevit) {
        $addinDest = Join-Path $revit.Path "Licorp_CombineCAD.addin"
        $dllDest = Join-Path $revit.Path "Licorp_CombineCAD.dll"

        try {
            if (!$Uninstall) {
                Copy-Item $addinSource -Destination $addinDest -Force
                Copy-Item $dllSource -Destination $dllDest -Force
                Write-Host "  Installed for Revit $($revit.Version)" -ForegroundColor Green
            } else {
                Remove-Item $addinDest -ErrorAction SilentlyContinue
                Remove-Item $dllDest -ErrorAction SilentlyContinue
                Write-Host "  Uninstalled for Revit $($revit.Version)" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "  Failed for Revit $($revit.Version): $_" -ForegroundColor Red
        }
    }
}

function Deploy-AutoCADPlugin {
    Write-Host "[Deploy] Installing AutoCAD plugin..." -ForegroundColor Cyan

    $bundleSource = Join-Path $ProjectRoot "src.acad\Licorp_MergeSheets"
    $dllSource = Join-Path $BinFolder "acad\Release\Licorp_MergeSheets.dll"

    $bundleDest = "$env:PROGRAMDATA\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle"

    if (!$Uninstall) {
        if (!(Test-Path $bundleDest)) {
            New-Item -ItemType Directory -Path $bundleDest -Force | Out-Null
        }

        Copy-Item (Join-Path $bundleSource "PackageContents.xml") -Destination $bundleDest -Force

        $contentsDest = Join-Path $bundleDest "Contents"
        if (!(Test-Path $contentsDest)) {
            New-Item -ItemType Directory -Path $contentsDest -Force | Out-Null
        }

        if (Test-Path $dllSource) {
            Copy-Item $dllSource -Destination $contentsDest -Force
            Write-Host "  Plugin installed to $bundleDest" -ForegroundColor Green
        } else {
            Write-Host "  Plugin DLL not found: $dllSource (build acad project first)" -ForegroundColor Yellow
        }
    } else {
        if (Test-Path $bundleDest) {
            Remove-Item $bundleDest -Recurse -Force
            Write-Host "  Plugin uninstalled" -ForegroundColor Yellow
        }
    }
}

# Main
Write-Host ""
Write-Host "======================================" -ForegroundColor Cyan
Write-Host " Licorp_CombiCAD Deploy Script" -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan
Write-Host ""

if ($Uninstall) {
    Write-Host "Mode: UNINSTALL" -ForegroundColor Yellow
} else {
    Write-Host "Mode: INSTALL" -ForegroundColor Green
}
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