Param(
    [ValidateSet('setup','run','package','clean','all')]
    [string]$Task = 'package',
    [string]$VenvPath = '.venv',
    [string]$Entry = 'run_app.py',
    [switch]$RecreateVenv,
    # Opciones estilo build.ps1 de referencia
    [string]$Name = 'ValeConsumoBioplates',
    [switch]$Console,
    [string]$Icon,
    [switch]$OneDir,
    [switch]$NoClean,
    [switch]$InstallMissing
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
        'pillow',     # PIL para reportlab/imagenes
        'openpyxl',   # lectura xlsx
        'xlrd',       # compatibilidad xls antiguos
        'pypdf'       # unificacion de PDFs
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

function Initialize-ToolsAndDeps {
    param([string]$Py)
    if ($InstallMissing) {
        Write-Host "[INFO] Verificando e instalando herramientas (-InstallMissing)" -ForegroundColor Cyan
        # PyInstaller
        & $Py -c "import importlib.util,sys;sys.exit(0 if importlib.util.find_spec('PyInstaller') else 1)" | Out-Null
        if ($LASTEXITCODE -ne 0) { & $Py -m pip install -U pyinstaller pyinstaller-hooks-contrib }
        # Dependencias runtime
        & $Py -c "import importlib.util as u,sys;mods=['pandas','reportlab','PIL','openpyxl','xlrd','win32api','win32print','pypdf','PyPDF2'];missing=[m for m in mods if u.find_spec(m) is None];sys.exit(0 if not missing else 1)" | Out-Null
        if ($LASTEXITCODE -ne 0) { & $Py -m pip install -U pandas reportlab pillow openpyxl xlrd pypdf pywin32 }
    }
    # Verificar PyInstaller
    & $Py -c "import importlib.util as u,sys;sys.exit(0 if u.find_spec('PyInstaller') else 1)" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller no estÃƒÂ¡ instalado. Ejecute con -InstallMissing o instale manualmente." }
}

function New-AppPackage {
    param([string]$Py, [string]$Entry)
    Initialize-ToolsAndDeps -Py $Py
    if (-not (Test-Path $Entry)) { throw "No se encontrÃƒÂ³ el entrypoint: $Entry" }

    # Construir lista de argumentos estilo referencia
    $argsList = @()
    $argsList += @($Entry)
    $argsList += @('--name', $Name)
    if (-not $OneDir) { $argsList += @('--onefile') }
    if (-not $NoClean) { $argsList += @('--clean') }
    $argsList += @('--noconfirm')
    if ($Console) { $argsList += @('--console') } else { $argsList += @('--windowed','--noconsole') }
    if ($Icon -and (Test-Path $Icon)) { $argsList += @('--icon', $Icon) }
    
# Hidden-imports necesarios
    $argsList += @('--hidden-import=reportlab')
    $argsList += @('--hidden-import=PIL')
    $argsList += @('--hidden-import=openpyxl')
    $argsList += @('--hidden-import=xlrd')
    $argsList += @('--hidden-import=win32api')
    $argsList += @('--hidden-import=win32print')
    $argsList += @('--hidden-import=pypdf')
    $argsList += @('--hidden-import=PyPDF2')
    $argsList += @('--collect-all','reportlab')
    $argsList += @('--collect-all','PIL')
    $argsList += @('--collect-all','openpyxl')
    $argsList += @('--collect-all','pypdf')
    Write-Host ("Ejecutando: {0} -m PyInstaller {1}" -f $Py, ($argsList -join ' ')) -ForegroundColor Cyan
    & $Py -m PyInstaller @argsList
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller fallÃƒÂ³ con cÃƒÂ³digo $LASTEXITCODE" }

    # Post-copia a dist
    $distRoot = Join-Path $PSScriptRoot 'dist'
    $histSrc = 'Vales_Historial'
    $histDst = Join-Path $distRoot 'Vales_Historial'
    if (!(Test-Path $histDst)) { New-Item -ItemType Directory -Force -Path $histDst | Out-Null }
    if (Test-Path $histSrc) { Copy-Item -Recurse -Force "$histSrc\*" $histDst -ErrorAction SilentlyContinue }
    $settingsSrc = 'app_settings.json'
    if (Test-Path $settingsSrc) { Copy-Item -Force $settingsSrc (Join-Path $distRoot $settingsSrc) -ErrorAction SilentlyContinue }
    Write-Host "Paquete generado en: dist/$Name.exe" -ForegroundColor Green
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




