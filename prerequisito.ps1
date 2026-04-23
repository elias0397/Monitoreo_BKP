# ==============================================================================
# Script de Preparación de Entorno Python para Descarga de Backups
# ==============================================================================

# 1. Bypass temporal de políticas de ejecución
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force

# 2. Definición de variables
$BackupDir = "C:\bkp"
$ProjectDir = "C:\Scripts\MonitorBackups" # Cambia esta ruta donde guardes tu script Python
$VenvDir = "$ProjectDir\venv"

Write-Host "Iniciando preparación del entorno..." -ForegroundColor Cyan

# 3. Creación del directorio de destino de los correos
if (!(Test-Path -Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir | Out-Null
    Write-Host "[OK] Carpeta creada: $BackupDir" -ForegroundColor Green
} else {
    Write-Host "[INFO] La carpeta $BackupDir ya existe." -ForegroundColor Yellow
}

# 4. Creación del directorio del proyecto (si no existe)
if (!(Test-Path -Path $ProjectDir)) {
    New-Item -ItemType Directory -Path $ProjectDir | Out-Null
    Write-Host "[OK] Carpeta del proyecto creada: $ProjectDir" -ForegroundColor Green
}

Set-Location -Path $ProjectDir

# 5. Creación del entorno virtual (VENV)
Write-Host "Creando entorno virtual de Python en $VenvDir..." -ForegroundColor Cyan
python -m venv venv

if (Test-Path -Path "$VenvDir\Scripts\activate.ps1") {
    Write-Host "[OK] Entorno virtual creado exitosamente." -ForegroundColor Green
    
    # 6. Activación e instalación de pywin32
    Write-Host "Activando entorno virtual e instalando pywin32..." -ForegroundColor Cyan
    & "$VenvDir\Scripts\python.exe" -m pip install --upgrade pip | Out-Null
    & "$VenvDir\Scripts\pip.exe" install pywin32
    
    Write-Host "[OK] pywin32 instalado correctamente." -ForegroundColor Green
} else {
    Write-Host "[ERROR] No se pudo crear el entorno virtual. Verifica que Python esté en el PATH." -ForegroundColor Red
}

Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "Entorno listo. Para ejecutar tu script, usa el comando:"
Write-Host "$VenvDir\Scripts\python.exe .\tu_script.py" -ForegroundColor Yellow
Write-Host "======================================================" -ForegroundColor Cyan