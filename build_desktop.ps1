$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $Root
$AppName = "卷烟投放策略转excel工具"
$ReleaseRoot = Join-Path $ProjectRoot "dist"
if ($env:STRATEGY_TOOL_RELEASE_DIR) {
    $ReleaseRoot = Join-Path $ProjectRoot $env:STRATEGY_TOOL_RELEASE_DIR
}
$DistRoot = Join-Path $ReleaseRoot $AppName
$ZipPath = Join-Path $ReleaseRoot "$AppName.zip"

Set-Location $Root

try {
    python -m PyInstaller --version | Out-Null
} catch {
    Write-Host "PyInstaller is not installed. Cannot build the exe." -ForegroundColor Yellow
    Write-Host "Install it first: python -m pip install pyinstaller" -ForegroundColor Yellow
    Write-Host "Then rerun: powershell -ExecutionPolicy Bypass -File strategy_excel_tool\build_desktop.ps1" -ForegroundColor Yellow
    exit 1
}

if (Test-Path $DistRoot) {
    Get-ChildItem -LiteralPath $DistRoot -Force | Where-Object { $_.Name -notin @("outputs", "private_templates") } | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force
    }
} else {
    New-Item -ItemType Directory -Force -Path $DistRoot | Out-Null
}

if (Test-Path (Join-Path $Root "dist")) {
    Remove-Item -LiteralPath (Join-Path $Root "dist") -Recurse -Force
}

if (Test-Path (Join-Path $Root "build")) {
    Remove-Item -LiteralPath (Join-Path $Root "build") -Recurse -Force
}

python -m PyInstaller `
    --noconfirm `
    --onedir `
    --name $AppName `
    --windowed `
    --clean `
    desktop_launcher.py

New-Item -ItemType Directory -Force -Path $DistRoot | Out-Null
Copy-Item -Path (Join-Path $Root "dist\$AppName\*") -Destination $DistRoot -Recurse -Force

$TemplateDocsOut = Join-Path $DistRoot "templates"
$PrivateTemplateRoot = Join-Path $ProjectRoot "private_templates"
$PrivateTemplateOut = Join-Path $DistRoot "private_templates"
$OutputOut = Join-Path $DistRoot "outputs"
New-Item -ItemType Directory -Force -Path $TemplateDocsOut | Out-Null
New-Item -ItemType Directory -Force -Path $PrivateTemplateOut | Out-Null
New-Item -ItemType Directory -Force -Path $OutputOut | Out-Null

if (Test-Path (Join-Path $Root "templates\README.md")) {
    Copy-Item -LiteralPath (Join-Path $Root "templates\README.md") -Destination $TemplateDocsOut -Force
}

if (Test-Path $PrivateTemplateRoot) {
    Get-ChildItem -Path $PrivateTemplateRoot -File -Force | Where-Object { $_.Name -notlike "~$*" } | ForEach-Object {
        try {
            Copy-Item -LiteralPath $_.FullName -Destination $PrivateTemplateOut -Force -ErrorAction Stop
        } catch {
            Write-Host "Private template in use, kept existing copy: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
} else {
    Write-Host "Private template directory not found: $PrivateTemplateRoot" -ForegroundColor Yellow
}

@"
Double-click 卷烟投放策略转excel工具.exe to start.

private_templates: contains the private .xls/.xlsx templates and 商品内码实例.xls.
templates: public documentation only; do not place business templates here.
outputs: exported .xls files are saved here.

If the browser does not open automatically, use the http://127.0.0.1:PORT/ address shown in the program window.
Use the Exit tool button at the top right of the webpage to stop the program.
Closing the last tool webpage stops the background service after about 10 seconds.
Refreshing or keeping another tool webpage open does not stop the service.
Active operations and an open Save As dialog postpone exit until they finish.
Export uses the system Save As dialog. Open file and Show in folder refer to your selected file, not the outputs backup.
v2.0: Customer allocation query exports filtered summaries, details, and issues as .xlsx.
System import remains .xls. Customer workbooks are not bundled; upload them locally in the query tab.
Large query exports split worksheets automatically. Do not terminate the application while export is running.
"@ | Set-Content -Encoding UTF8 -Path (Join-Path $DistRoot "README.txt")

@"
@echo off
start "" "%~dp0卷烟投放策略转excel工具.exe"
"@ | Set-Content -Encoding Default -Path (Join-Path $DistRoot "start.bat")

if (Test-Path $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
}
Compress-Archive -LiteralPath $DistRoot -DestinationPath $ZipPath -CompressionLevel Optimal

Write-Host "Build complete:" -ForegroundColor Green
Write-Host $DistRoot
Write-Host $ZipPath
