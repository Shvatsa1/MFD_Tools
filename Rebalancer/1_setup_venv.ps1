Set-Location $PSScriptRoot
if (-not (Test-Path '.venv\Scripts\python.exe')) { py -3 -m venv .venv }
if (-not (Test-Path '.venv\Scripts\python.exe')) { python -m venv .venv }
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
& '.\.venv\Scripts\Activate.ps1'
python -m pip install --upgrade pip | Out-Null
pip install -r requirements.txt
if (-not (Test-Path 'mfd_pack.ini')) { Copy-Item mfd_pack.example.ini mfd_pack.ini }
Write-Host ''
Write-Host 'Next: add data\latestNAV_Reports.xlsx, client files under data\clients\by_client\, edit mfd_pack.ini, then: python build_mfd_pack.py' -ForegroundColor Cyan
