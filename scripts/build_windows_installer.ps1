param(
    [string]$InnoSetupPath = ""
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $Root

$Python = Join-Path $Root ".tkvenv\Scripts\python.exe"
if (-not (Test-Path $Python)) {
    throw "Missing .tkvenv. Run install.bat first."
}

Write-Host "Installing build dependencies..."
& $Python -m pip install -r requirements-build.txt

Write-Host "Generating application icon..."
& $Python scripts\create_icon.py

Write-Host "Building standalone application with PyInstaller..."
& $Python -m PyInstaller --noconfirm --clean packaging\FileOptimizerForAI.spec

if (-not $InnoSetupPath) {
    $Candidates = @(
        "C:\tmp\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 7\ISCC.exe",
        "C:\Program Files\Inno Setup 7\ISCC.exe"
    )
    foreach ($Candidate in $Candidates) {
        if (Test-Path $Candidate) {
            $InnoSetupPath = $Candidate
            break
        }
    }
}

if (-not $InnoSetupPath -or -not (Test-Path $InnoSetupPath)) {
    throw "Inno Setup Compiler was not found. Install Inno Setup 6 or pass -InnoSetupPath."
}

Write-Host "Compiling installer with Inno Setup..."
& $InnoSetupPath packaging\FileOptimizerForAI.iss

Write-Host ""
Write-Host "Installer created in: $Root\installer"
