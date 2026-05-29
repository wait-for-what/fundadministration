# 用途: Windows PowerShell 手工执行 fundadmin CLI 的薄封装。
$ErrorActionPreference = "Stop"
Set-Location (Join-Path $PSScriptRoot "..")
if (Test-Path ".venv\Scripts\Activate.ps1") {
    . .venv\Scripts\Activate.ps1
}
python -m fundadmin.cli @args
