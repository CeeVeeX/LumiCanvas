param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "1.0.0",
    [string]$OutputDir = "artifacts",
    [string]$AppName = "LumiCanvas"
)

$ErrorActionPreference = "Stop"

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$publishDir = Join-Path $projectRoot "bin\$Configuration\net8.0-windows10.0.19041.0\$Runtime"
$outDir = Join-Path $projectRoot $OutputDir
$installerPath = Join-Path $projectRoot "installer\LumiCanvas.nsi"
$outFile = Join-Path $outDir "$AppName-Setup-$Version-$Runtime.exe"

if (-not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

Write-Host "Building app..." -ForegroundColor Cyan

dotnet build "$projectRoot\LumiCanvas.csproj" `
    -c $Configuration `
    -r $Runtime `
    -p:WindowsPackageType=None `
    -p:PublishTrimmed=false `
    -p:PublishReadyToRun=false

if (-not (Test-Path $publishDir)) {
    throw "Build output not found: $publishDir"
}

$makensis = Get-Command makensis -ErrorAction SilentlyContinue
if (-not $makensis) {
    throw "makensis not found. Please install NSIS and add it to PATH."
}

Write-Host "Building NSIS installer..." -ForegroundColor Cyan
$nsisArgs = @(
    "/DAPP_NAME=$AppName"
    "/DAPP_VERSION=$Version"
    "/DPUBLISH_DIR=$publishDir"
    "/DOUT_FILE=$outFile"
    "$installerPath"
)

& $makensis.Source @nsisArgs

if ($LASTEXITCODE -ne 0) {
    throw "NSIS build failed with exit code $LASTEXITCODE"
}

Write-Host "Done: $outFile" -ForegroundColor Green
