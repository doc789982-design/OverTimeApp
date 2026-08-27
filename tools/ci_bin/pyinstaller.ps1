$ErrorActionPreference = "Stop"
& python (Join-Path $PSScriptRoot "..\ci_pyinstaller.py") @args
exit $LASTEXITCODE
