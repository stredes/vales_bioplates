Param(
    [ValidateSet('setup','run','package','clean','all')]
    [string]$Task = 'run',
    [string]$VenvPath = '.venv',
    [string]$Entry = 'vale_consumo_bioplates.py',
    [switch]$RecreateVenv
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

function Get-PythonPath {
    param([string]$VenvPath)
    $venvPy = Join-Path $VenvPath 'Scripts/python.exe'
    if (Test-Path $venvPy) { return (Resolve-Path $venvPy).Path }
    if (Get-Command python -ErrorAction SilentlyContinue) { return 'python' }
    if (Get-Command py -ErrorAction SilentlyContinue) { return 'py' }
    throw 'Python no encontrado en PATH ni en el venv.'
}

function New-Venv {
    param([string]$VenvPath, [switch]$Recreate)
    if ($Recreate -and (Test-Path $VenvPath)) { Remove-Item -Recurse -Force $VenvPath }
    if (-not (Test-Path $VenvPath)) {
        if (Get-Command python -ErrorAction SilentlyContinue) {
            python -m venv $VenvPath
        } elseif (Get-Command py -ErrorAction SilentlyContinue) {
            py -m venv $VenvPath
        } else {
            throw 'Python no encontrado para crear el entorno virtual.'
        }
    }
}

function Install-Dependencies {
    param([string]$Py)
    & $Py -m pip install --upgrade pip
    $pkgs = @(
        'pandas',
        'reportlab',
        'openpyxl',   # lectura xlsx
        'xlrd'        # compatibilidad xls antiguos
    )
    & $Py -m pip install @pkgs
    if ($IsWindows) {
        try { & $Py -m pip install pywin32 } catch { Write-Warning 'pywin32 opcional; continuando.' }
    }
}

function New-OutputFolders {
    if (-not (Test-Path 'Vales_Historial')) { New-Item -ItemType Directory -Force -Path 'Vales_Historial' | Out-Null }
}

function Start-Application {
    param([string]$Py, [string]$Entry)
    New-OutputFolders
    & $Py $Entry
}

function New-AppPackage {
    param([string]$Py, [string]$Entry)
    & $Py -m pip install --upgrade pyinstaller
    $name = 'ValeConsumoBioplates'
    if (-not (Test-Path $Entry)) { throw "No se encontro el entrypoint: $Entry" }
    $argsList = @('-m','PyInstaller','--noconfirm','--clean','--name', $name,'--windowed','--onefile','--distpath','dist','--workpath','build','--specpath','build', $Entry)
    Write-Host ("Ejecutando: {0} {1}" -f $Py, ($argsList -join ' ')) -ForegroundColor Cyan
    & $Py @argsList
    $code = $LASTEXITCODE
    if ($code -ne 0) {
        $pyinstallerExe = Join-Path (Split-Path $Py) 'pyinstaller.exe'
        if (-not (Test-Path $pyinstallerExe)) { $pyinstallerExe = 'pyinstaller' }
        Write-Warning "python -m PyInstaller fallo (codigo $code). Probando ejecutable: $pyinstallerExe"
        & $pyinstallerExe '--noconfirm' '--clean' '--name' $name '--windowed' '--onefile' '--distpath' 'dist' '--workpath' 'build' '--specpath' 'build' $Entry
        $code = $LASTEXITCODE
        if ($code -ne 0) { throw "PyInstaller fallo con codigo $code" }
    }
    $distRoot = 'dist'
    $histSrc = 'Vales_Historial'
    $histDst = Join-Path $distRoot 'Vales_Historial'
    if (!(Test-Path $histDst)) { New-Item -ItemType Directory -Force -Path $histDst | Out-Null }
    if (Test-Path $histSrc) { Copy-Item -Recurse -Force "$histSrc\*" $histDst -ErrorAction SilentlyContinue }
    $settingsSrc = 'app_settings.json'
    if (Test-Path $settingsSrc) { Copy-Item -Force $settingsSrc (Join-Path $distRoot $settingsSrc) -ErrorAction SilentlyContinue }
    Write-Host "Paquete generado en: dist/$name.exe" -ForegroundColor Green
}

function Remove-BuildArtifacts {
    if (Test-Path 'build') { Remove-Item -Recurse -Force 'build' }
    if (Test-Path 'dist') { Remove-Item -Recurse -Force 'dist' }
    Get-ChildItem -Filter '*.spec' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
}

switch ($Task) {
    'setup' {
        New-Venv -VenvPath $VenvPath -Recreate:$RecreateVenv
        $py = Get-PythonPath -VenvPath $VenvPath
        Install-Dependencies -Py $py
    }
    'run' {
        New-Venv -VenvPath $VenvPath
        $py = Get-PythonPath -VenvPath $VenvPath
        Install-Dependencies -Py $py
        Start-Application -Py $py -Entry $Entry
    }
    'package' {
        New-Venv -VenvPath $VenvPath
        $py = Get-PythonPath -VenvPath $VenvPath
        Install-Dependencies -Py $py
        New-AppPackage -Py $py -Entry $Entry
    }
    'clean' {
        Remove-BuildArtifacts
    }
    'all' {
        New-Venv -VenvPath $VenvPath -Recreate:$RecreateVenv
        $py = Get-PythonPath -VenvPath $VenvPath
        Install-Dependencies -Py $py
        New-AppPackage -Py $py -Entry $Entry
    }
    default { throw "Tarea desconocida: $Task" }
}

