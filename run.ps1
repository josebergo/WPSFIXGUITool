$ErrorActionPreference = 'Stop'
$command = Get-Command pythonw.exe -ErrorAction SilentlyContinue
if ($null -eq $command) {
    $command = Get-Command python.exe -ErrorAction SilentlyContinue
}
if ($null -eq $command) {
    throw '未找到 Python。请先安装 Python 3.10+ 并安装 requirements.txt。'
}
$python = $command.Source
Start-Process -FilePath $python -ArgumentList (Join-Path $PSScriptRoot 'app.py') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
