param(
    [string]$ProjectDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

Set-Location $ProjectDir

Write-Host "[1/5] Checking Python..."
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    throw "Python is not found in PATH. Install Python 3.10+ and run the script again."
}

Write-Host "[2/5] Creating virtual environment..."
if (-not (Test-Path ".venv")) {
    python -m venv .venv
}

$venvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$venvPip = Join-Path $ProjectDir ".venv\Scripts\pip.exe"

Write-Host "[3/5] Installing Python dependencies..."
& $venvPython -m pip install --upgrade pip
& $venvPip install -r requirements.txt

Write-Host "[4/5] Checking Tesseract OCR..."
$tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
if (-not $tesseractCmd -and -not (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe")) {
    Write-Warning "Tesseract was not found. Install it with: winget install UB-Mannheim.TesseractOCR"
} else {
    Write-Host "Tesseract is available."
}

Write-Host "[5/5] Creating launcher script..."
$runScript = @(
    '@echo off'
    'cd /d %~dp0'
    'call .venv\Scripts\activate.bat'
    'python main.py'
) -join [Environment]::NewLine
Set-Content -Path "run_parser.bat" -Value $runScript -Encoding ASCII

Write-Host "Done. Start the app using run_parser.bat"
