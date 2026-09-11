# Build VBA Add-in Editor: onedir for debugging, then onefile release.
# Usage: pwsh scripts/build.ps1 [-Onefile] [-Release]
param(
    [switch]$Onefile,
    [switch]$Release
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

function Invoke-Checked {
    param([string]$Label, [scriptblock]$Command)
    & $Command
    if ($LASTEXITCODE -ne 0) { throw "$Label failed with exit $LASTEXITCODE" }
}

if ($Release) {
    $status = git status --porcelain
    if ($LASTEXITCODE -ne 0) { throw "git status failed" }
    if ($status) { throw "Release builds require a clean git tree." }
    $env:VBAAE_RELEASE = "1"
}

pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw "pip install failed" }
if ($Onefile) {
    pyinstaller packaging/VBAAddinEditor.spec -- --onefile --noconfirm
} else {
    pyinstaller packaging/VBAAddinEditor.spec --noconfirm
}
if ($LASTEXITCODE -ne 0) { throw "pyinstaller failed" }

$exe = if ($Onefile) { "dist/VBAAddinEditor.exe" } else { "dist/VBAAddinEditor/VBAAddinEditor.exe" }
$fixtureCandidates = @(
    "tests/fixtures/xlam/Demo.xlam",
    (Join-Path $root "tests/fixtures/xlam/Demo.xlam")
)
$fixture = $fixtureCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $fixture) {
    Write-Warning "Demo fixture missing; packaged smoke tests skipped. Live qualification is incomplete."
} else {
    $fixture = (Resolve-Path $fixture).Path
    & $exe --self-test $fixture
    if ($LASTEXITCODE -ne 0) { throw "packaged self-test failed" }
    & $exe --self-roundtrip $fixture
    if ($LASTEXITCODE -ne 0) { throw "packaged round-trip failed" }
}
Write-Host "Build OK: $exe"
if ($Onefile) {
    $hash = (Get-FileHash $exe -Algorithm SHA256).Hash
    "$hash  $exe" | Out-File "dist/VBAAddinEditor.exe.sha256"
    Write-Host "SHA-256: $hash"
    $manifest = @{
        executable = $exe
        sha256 = $hash
        mode = "onefile"
        identityFile = "dist/build-identity.json"
    } | ConvertTo-Json
    $manifest | Out-File "dist/release-manifest.json"
} else {
    $hash = (Get-FileHash $exe -Algorithm SHA256).Hash
    $manifest = @{
        executable = $exe
        sha256 = $hash
        mode = "onedir"
        identityFile = "dist/build-identity.json"
    } | ConvertTo-Json
    $manifest | Out-File "dist/release-manifest-onedir.json"
}
