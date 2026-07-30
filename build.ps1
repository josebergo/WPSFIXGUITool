param(
    [switch]$SkipInstall
)

$ErrorActionPreference = 'Stop'
$project = $PSScriptRoot
$command = Get-Command python.exe -ErrorAction Stop
$python = $command.Source

$toolPath = Join-Path $project '.build-tools'
if (-not $SkipInstall) {
    & $python -m pip install --disable-pip-version-check --target $toolPath -r (Join-Path $project 'requirements-dev.txt')
    if ($LASTEXITCODE -ne 0) { throw '依赖安装失败。' }
}

$env:PYTHONPATH = $toolPath
$icon = Join-Path $project 'assets\wpsfix.ico'
$versionInfo = Join-Path $project 'version_info.txt'
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name 'WPSFIXGUITool' `
    --icon $icon `
    --version-file $versionInfo `
    --add-data "$icon;assets" `
    --distpath (Join-Path $project 'dist') `
    --workpath (Join-Path $project 'build') `
    --specpath $project `
    --paths $project `
    --hidden-import PIL._tkinter_finder `
    (Join-Path $project 'app.py')
if ($LASTEXITCODE -ne 0) { throw 'EXE 打包失败。' }

Write-Host "完成：$(Join-Path $project 'dist\WPSFIXGUITool.exe')"
