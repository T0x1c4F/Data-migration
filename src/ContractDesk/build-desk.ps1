$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$python = if ($env:PYTHON) { $env:PYTHON } else { 'python' }
$dependencyPath = Join-Path $projectRoot 'appdeps'
if (Test-Path -LiteralPath $dependencyPath) { $env:PYTHONPATH = $dependencyPath }
$buildPath = Join-Path $projectRoot 'build'
New-Item -ItemType Directory -Force -Path $buildPath | Out-Null
$env:TEMP = $buildPath
$env:TMP = $buildPath
& $python -m PyInstaller --noconfirm --onefile --windowed --name 'Contract Desk' `
    --exclude-module numpy --exclude-module lxml `
    --distpath (Join-Path $projectRoot 'dist/Contract Desk') `
    --workpath (Join-Path $buildPath 'contract-desk') --specpath $buildPath `
    (Join-Path $PSScriptRoot 'contract_desk.py')
if ($LASTEXITCODE -ne 0) { throw 'Contract Desk packaging failed.' }
