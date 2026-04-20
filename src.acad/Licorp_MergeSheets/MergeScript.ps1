# Licorp_MergeSheets AutoCAD Script
# Dùng để merge DWG files thông qua AutoCAD scripting
# Không cần .NET plugin - chỉ cần AutoCAD

param(
    [Parameter(Mandatory=$true)]
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $ConfigPath)) {
    Write-Host "[Merge] Config file not found: $ConfigPath" -ForegroundColor Red
    exit 1
}

# Read config JSON
$config = Get-Content $ConfigPath | ConvertFrom-Json

Write-Host "[Merge] Mode: $($config.Mode)"
Write-Host "[Merge] Output: $($config.OutputPath)"
Write-Host "[Merge] Files: $($config.SourceFiles.Count)"

$accoreconsole = "C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe"
if (-not (Test-Path $accoreconsole)) {
    $accoreconsole = "C:\Program Files\Autodesk\AutoCAD 2025\accoreconsole.exe"
}

if (-not (Test-Path $accoreconsole)) {
    Write-Host "[Merge] AcCoreConsole not found!" -ForegroundColor Red
    exit 1
}

# Tạo script để chạy trong AutoCAD
$scriptContent = @"
SECURELOAD 0
NETLOAD "C:\ProgramData\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle\Contents\Licorp_MergeSheets.dll"
LICORP_MERGESHEETS "$ConfigPath"
QUIT
Y
"@

$scrPath = [System.IO.Path]::GetTempFileName() + ".scr"
$scriptContent | Out-File -FilePath $scrPath -Encoding ASCII

Write-Host "[Merge] Running AcCoreConsole..."
$proc = Start-Process -FilePath $accoreconsole -ArgumentList "/i `"$($config.OutputPath)`" /s `"$scrPath`"" -Wait -PassThru -NoNewWindow

Remove-Item $scrPath -ErrorAction SilentlyContinue

if ($proc.ExitCode -eq 0) {
    Write-Host "[Merge] Success!" -ForegroundColor Green
} else {
    Write-Host "[Merge] Failed with exit code: $($proc.ExitCode)" -ForegroundColor Red
}

exit $proc.ExitCode