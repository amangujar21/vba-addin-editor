# Build VBA Add-in Editor: onedir for debugging, then onefile release.
# Usage: pwsh scripts/build.ps1 [-Onefile]
param(
    [switch]$Onefile
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

pip install -r requirements-build.txt
if ($Onefile) {
    pyinstaller packaging/VBAAddinEditor.spec -- --onefile --noconfirm
} else {
    pyinstaller packaging/VBAAddinEditor.spec --noconfirm
}

# Package smoke test (plan 36) on the bundled exe.
$exe = if ($Onefile) { "dist/VBAAddinEditor.exe" } else { "dist/VBAAddinEditor/VBAAddinEditor.exe" }
$fixture = "tests/fixtures/xlam/Demo.xlam"
& $exe --self-test $fixture
if ($LASTEXITCODE -ne 0) { throw "packaged self-test failed" }
& $exe --self-roundtrip $fixture
if ($LASTEXITCODE -ne 0) { throw "packaged round-trip failed" }
Write-Host "Build OK: $exe"
if ($Onefile) {
    $hash = (Get-FileHash $exe -Algorithm SHA256).Hash
    "$hash  $exe" | Out-File "dist/VBAAddinEditor.exe.sha256"
    Write-Host "SHA-256: $hash"
}
