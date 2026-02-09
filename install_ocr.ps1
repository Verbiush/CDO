# Script para instalar Tesseract OCR (Requerido para PDFs de imagen SOS)
Write-Host "Iniciando instalación de dependencias OCR..." -ForegroundColor Cyan

# 1. Instalar pytesseract (Librería Python)
Write-Host "Instalando pytesseract..."
pip install pytesseract
if ($LASTEXITCODE -ne 0) {
    Write-Host "Error instalando pytesseract via pip." -ForegroundColor Red
}

# 2. Instalar Tesseract OCR Engine (Binario) via Winget
Write-Host "Buscando Tesseract-OCR en el sistema..."
$tesseractPath = "C:\Program Files\Tesseract-OCR\tesseract.exe"
if (Test-Path $tesseractPath) {
    Write-Host "Tesseract ya está instalado en $tesseractPath" -ForegroundColor Green
    exit 0
}

Write-Host "Tesseract no encontrado. Intentando instalar via Winget..."
winget install UB-Mannheim.TesseractOCR --accept-source-agreements --accept-package-agreements --silent

if ($LASTEXITCODE -eq 0) {
    Write-Host "Instalación completada exitosamente." -ForegroundColor Green
    Write-Host "Por favor reinicia la aplicación para que detecte los cambios." -ForegroundColor Yellow
} else {
    Write-Host "La instalación automática falló." -ForegroundColor Red
    Write-Host "Por favor descarga e instala manualmente desde: https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor Yellow
}

Pause
