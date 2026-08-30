# Run the full local quality gate: pytest.
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
python -m pytest tests
exit $LASTEXITCODE
