param(
    [string]$ProjectDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

Set-Location $ProjectDir

Write-Host "[1/5] Проверка Python..."
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    throw "Python не найден в PATH. Установите Python 3.10+ и запустите скрипт снова."
}

Write-Host "[2/5] Создание виртуального окружения..."
if (-not (Test-Path ".venv")) {
    python -m venv .venv
}

$venvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$venvPip = Join-Path $ProjectDir ".venv\Scripts\pip.exe"

Write-Host "[3/5] Установка зависимостей Python..."
& $venvPython -m pip install --upgrade pip
& $venvPip install -r requirements.txt

Write-Host "[4/5] Проверка Tesseract OCR..."
$tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
if (-not $tesseractCmd -and -not (Test-Path "C:\Program Files\Tesseract-OCR\tesseract.exe")) {
    Write-Warning "Tesseract не найден. Установите его: winget install UB-Mannheim.TesseractOCR"
} else {
    Write-Host "Tesseract найден."
}

Write-Host "[5/5] Создание скрипта запуска..."
$runScript = @"
@echo off
cd /d %~dp0
call .venv\Scripts\activate.bat
python main.py
"@
Set-Content -Path "run_parser.bat" -Value $runScript -Encoding ASCII

Write-Host "Готово. Запускайте приложение через run_parser.bat"
