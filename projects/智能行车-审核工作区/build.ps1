param([Parameter(Mandatory=$true)][ValidatePattern('^[A-Za-z][A-Za-z0-9_-]{0,63}$')][string]$AppName)
$ErrorActionPreference="Stop"
$projectRoot=Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot
$env:AUTOTESTING_APP_NAME=$AppName
uv run --with "PySide6==6.8.3" --with "openpyxl>=3.1,<4" --with "docxtpl>=0.20,<1" --with "pyinstaller>=6.10,<7" pyinstaller --noconfirm --clean application.spec
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }
$exe=Join-Path $projectRoot "dist\$AppName\$AppName.exe"
if (-not (Test-Path -LiteralPath $exe)) { throw "Expected EXE was not created: $exe" }
Write-Host "Build complete: $exe"
