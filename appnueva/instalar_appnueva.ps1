$ErrorActionPreference = "Stop"

# Directorio del proyecto (carpeta padre de este script).
$ProjectRoot = Join-Path $PSScriptRoot ".."
Set-Location $ProjectRoot

function Write-Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Write-Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }

function Ensure-Python {
    $candidates = @(
        "C:\Program Files\Python313\python.exe",
        "C:\Program Files\Python312\python.exe",
        "python"
    )
    foreach ($p in $candidates) {
        try {
            $ver = & $p -c "import sys; print(sys.version)" 2>$null
            if ($LASTEXITCODE -eq 0) {
                Write-Info "Usando Python: $p (version $ver)"
                return $p
            }
        } catch { }
    }

    Write-Warn "Python no encontrado. Intentando instalar Python 3.13 con winget..."
    try {
        winget install -e --id Python.Python.3.13 -h
    } catch {
        throw "No se pudo instalar Python. Instala manualmente Python 3.13 y vuelve a correr este script."
    }

    $python = "C:\Program Files\Python313\python.exe"
    if (-not (Test-Path $python)) {
        throw "Python 3.13 no quedo disponible despues de la instalacion. Revisa la instalacion manualmente."
    }
    return $python
}

$python = Ensure-Python

# Crear y activar entorno virtual
$venvPath = Join-Path $ProjectRoot ".venv"
if (-not (Test-Path $venvPath)) {
    Write-Info "Creando entorno virtual en $venvPath"
    & $python -m venv $venvPath
}
$venvPython = Join-Path $venvPath "Scripts\python.exe"

Write-Info "Actualizando pip..."
& $venvPython -m pip install --upgrade pip

Write-Info "Instalando dependencias..."
& $venvPython -m pip install PySide6 pandas matplotlib openpyxl

# Crear lanzadores actualizados (.bat y .ps1) que usan el venv.
$batPath = Join-Path $ProjectRoot "appnueva\lanzar_appnueva.bat"
$ps1Path = Join-Path $ProjectRoot "appnueva\lanzar_appnueva.ps1"

$batContent = @"
@echo off
setlocal
pushd "%~dp0.."
set "PY=%~dp0..\.venv\Scripts\python.exe"
if not exist "%PY%" (
    echo No se encontro %PY%. Ejecutaste instalar_appnueva.ps1?
    pause
    exit /b 1
)
"%PY%" -m ventas_app.main --folder ventas
pause
"@

$ps1Content = @"
$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")
$python = Join-Path $PSScriptRoot "..\.venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "No se encontro el entorno virtual. Corre primero instalar_appnueva.ps1."
}
& $python -m ventas_app.main --folder ventas
"@

Set-Content -Path $batPath -Value $batContent -Encoding UTF8
Set-Content -Path $ps1Path -Value $ps1Content -Encoding UTF8

# (Opcional) Crear acceso directo en el escritorio.
try {
    $shell = New-Object -ComObject WScript.Shell
    $desktop = [Environment]::GetFolderPath("Desktop")
    $shortcut = $shell.CreateShortcut((Join-Path $desktop "Ventas App.lnk"))
    $shortcut.TargetPath = $batPath
    $shortcut.WorkingDirectory = (Join-Path $ProjectRoot "appnueva")
    $shortcut.Save()
    Write-Info "Acceso directo creado en el escritorio."
} catch {
    Write-Warn "No se pudo crear el acceso directo: $_"
}

Write-Host "" 
Write-Host "Instalacion completa." -ForegroundColor Green
Write-Host "Para iniciar: doble click en 'Ventas App' del escritorio o en appnueva\lanzar_appnueva.bat"
