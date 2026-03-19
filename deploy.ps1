param(
    [string]$ProjectDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

Set-Location $ProjectDir

function New-Venv {
    if (Test-Path ".venv") {
        return
    }

    try {
        python -m venv .venv
    } catch {
        $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
        if (-not $pyLauncher) {
            throw "Failed to create .venv with 'python -m venv'. Python launcher 'py' is also unavailable."
        }
        & py -3 -m venv .venv
    }
}

function Get-VenvPythonPath {
    $candidates = @(
        (Join-Path $ProjectDir ".venv\Scripts\python.exe"),
        (Join-Path $ProjectDir ".venv\python.exe"),
        (Join-Path $ProjectDir ".venv\bin\python"),
        (Join-Path $ProjectDir ".venv\bin\python3")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }
    return $null
}

Write-Host "[1/5] Checking Python..."
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
$pyCmd = Get-Command py -ErrorAction SilentlyContinue
if (-not $pythonCmd -and -not $pyCmd) {
    throw "Python is not found in PATH. Install Python 3.10+ and run the script again."
}

Write-Host "[2/5] Creating virtual environment..."
New-Venv
$venvPython = Get-VenvPythonPath
if (-not $venvPython) {
    throw "Virtual environment was created, but Python executable was not found in .venv."
}

Write-Host "[3/5] Installing Python dependencies..."
& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r requirements.txt

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
