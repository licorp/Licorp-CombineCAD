param(
    [string]$Version = "1.0.0",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$artifactRoot = Join-Path $root "artifacts"
$releaseRoot = Join-Path $artifactRoot "release\$Version"
$sessionId = Get-Date -Format "yyyyMMdd_HHmmss"
$workRoot = Join-Path $releaseRoot "work\$sessionId"
$stagingRoot = Join-Path $workRoot "staging"
$revitStage = Join-Path $stagingRoot "revit"
$autocadStage = Join-Path $stagingRoot "autocad"
$installerDir = Join-Path $releaseRoot "installer"
$installerTempDir = Join-Path $workRoot "installer_tmp"
$installerScript = Join-Path $root "installer\Licorp_CombineCAD.iss"
$iscc = "C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

$revitVersions = @("R2020","R2021","R2022","R2023","R2024","R2025","R2026","R2027")

function Ensure-Directory([string]$Path) {
    if (!(Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Remove-DirectoryContents([string]$Path) {
    if (!(Test-Path -LiteralPath $Path)) {
        return
    }

    $items = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)
    foreach ($item in $items) {
        if ($null -eq $item -or [string]::IsNullOrWhiteSpace($item.FullName)) {
            continue
        }

        try {
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
        }
        catch {
            Start-Sleep -Milliseconds 250
            Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
        }
    }
}

function New-RevitAddinManifest([string]$assemblyPath) {
@"
<?xml version="1.0" encoding="utf-8" standalone="no"?>
<RevitAddIns>
  <AddIn Type="Application">
    <Name>Licorp CombineCAD</Name>
    <Assembly>$assemblyPath</Assembly>
    <AddInId>F7A3B2C1-4D5E-6F78-9A0B-C1D2E3F4A5B6</AddInId>
    <FullClassName>Licorp_CombineCAD.App</FullClassName>
    <VendorId>LICORP</VendorId>
    <VendorDescription>Licorp - Export Revit sheets to DWG with multi-layout merge</VendorDescription>
  </AddIn>
</RevitAddIns>
"@
}

function Get-RevitBundleAssemblyPath([string]$versionTag) {
    $programData = [Environment]::GetFolderPath("CommonApplicationData")
    return (Join-Path $programData "Autodesk\ApplicationPlugins\Licorp_CombineCAD\$versionTag\Licorp_CombineCAD.dll")
}

function Stage-RevitFiles {
    foreach ($versionTag in $revitVersions) {
        $year = $versionTag.Substring(1)
        $sourceDir = Join-Path $root "bin\$versionTag\Release"
        if (!(Test-Path -LiteralPath $sourceDir)) {
            throw "Missing Revit build output: $sourceDir"
        }

        $bundleDir = Join-Path $revitStage $versionTag
        Ensure-Directory $bundleDir
        Copy-Item -Path (Join-Path $sourceDir "*") -Destination $bundleDir -Recurse -Force

        $addinDir = Join-Path $revitStage "Addins\$year"
        Ensure-Directory $addinDir
        $manifestPath = Join-Path $addinDir "Licorp_CombineCAD.addin"
        $assemblyPath = Get-RevitBundleAssemblyPath $versionTag
        Set-Content -LiteralPath $manifestPath -Value (New-RevitAddinManifest $assemblyPath) -Encoding UTF8
    }
}

function Stage-AutoCADFiles {
    $sourceBase = Join-Path $root "bin\acad\Release"
    $sourcePackage = Join-Path $root "src.acad\Licorp_MergeSheets\PackageContents.xml"

    if (!(Test-Path -LiteralPath (Join-Path $sourceBase "net48\Licorp_MergeSheets.dll"))) {
        throw "Missing AutoCAD net48 build output."
    }

    if (!(Test-Path -LiteralPath (Join-Path $sourceBase "net8.0-windows\Licorp_MergeSheets.dll"))) {
        throw "Missing AutoCAD net8.0 build output."
    }

    Ensure-Directory (Join-Path $autocadStage "Contents\2024")
    Ensure-Directory (Join-Path $autocadStage "Contents\2025")

    Copy-Item -LiteralPath (Join-Path $sourceBase "net48\Licorp_MergeSheets.dll") -Destination (Join-Path $autocadStage "Contents\2024") -Force
    Copy-Item -LiteralPath (Join-Path $sourceBase "net48\Newtonsoft.Json.dll") -Destination (Join-Path $autocadStage "Contents\2024") -Force
    Copy-Item -LiteralPath (Join-Path $sourceBase "net8.0-windows\Licorp_MergeSheets.dll") -Destination (Join-Path $autocadStage "Contents\2025") -Force
    Copy-Item -LiteralPath (Join-Path $sourceBase "net8.0-windows\Newtonsoft.Json.dll") -Destination (Join-Path $autocadStage "Contents\2025") -Force
    Copy-Item -LiteralPath $sourcePackage -Destination $autocadStage -Force
}

function Build-Projects {
    & dotnet build "Licorp_CombineCAD.sln" -c Release --nologo
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet build failed."
    }
}

Ensure-Directory $artifactRoot
Ensure-Directory $releaseRoot
Ensure-Directory $installerDir
Ensure-Directory $workRoot
Ensure-Directory $installerTempDir

Ensure-Directory $stagingRoot
Ensure-Directory $revitStage
Ensure-Directory $autocadStage

if (-not $SkipBuild) {
    Build-Projects
}

Stage-RevitFiles
Stage-AutoCADFiles

if (!(Test-Path -LiteralPath $iscc)) {
    throw "Inno Setup compiler not found: $iscc"
}

$tempSetupExe = Join-Path $installerTempDir "Licorp_CombineCAD_Setup_$Version.exe"
$setupExe = Join-Path $installerDir "Licorp_CombineCAD_Setup_$Version.exe"
$escapedStagingRoot = $stagingRoot.Replace("\", "\\")
& $iscc "/DMyStagingRoot=$escapedStagingRoot" "/O$installerTempDir" "/FLicorp_CombineCAD_Setup_$Version" $installerScript
if ($LASTEXITCODE -ne 0) {
    throw "Inno Setup compilation failed."
}

if (!(Test-Path -LiteralPath $tempSetupExe)) {
    throw "Installer was not generated: $tempSetupExe"
}

$targetLocked = $false
try {
    if (Test-Path -LiteralPath $setupExe) {
        try {
            Remove-Item -LiteralPath $setupExe -Force -ErrorAction Stop
        }
        catch {
            $targetLocked = $true
        }
    }
}
catch {
    $targetLocked = $true
}

if ($targetLocked) {
    $setupExe = Join-Path $installerDir ("Licorp_CombineCAD_Setup_{0}_{1:yyyyMMdd_HHmmss}.exe" -f $Version, (Get-Date))
}

if (!(Test-Path -LiteralPath $setupExe)) {
    Copy-Item -LiteralPath $tempSetupExe -Destination $setupExe -Force
}

$zipPath = Join-Path $artifactRoot "Licorp_CombineCAD_Setup_$Version.zip"
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -LiteralPath $setupExe -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "Release folder: $releaseRoot" -ForegroundColor Green
Write-Host "Installer: $setupExe" -ForegroundColor Green
Write-Host "Zip: $zipPath" -ForegroundColor Green
