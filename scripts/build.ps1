# Build VBA Add-in Editor: onedir for debugging, then onefile release.
# Usage: pwsh scripts/build.ps1 [-Onefile] [-Release]
#
# -Release is a strict qualification gate. It fails if tests, lint, fixtures,
# packaged smoke tests, or identity evidence are missing. It never prints
# "Build OK" or writes a release manifest in that case.
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
    if ($null -ne $LASTEXITCODE -and $LASTEXITCODE -ne 0) {
        throw "$Label failed with exit $LASTEXITCODE"
    }
}

function Get-GitPorcelain {
    $status = git status --porcelain
    if ($LASTEXITCODE -ne 0) { throw "git status failed" }
    return $status
}

$dirty = [bool](Get-GitPorcelain)
if ($Release -and $dirty) {
    throw "Release builds require a clean git tree."
}
if ($Release) {
    $env:VBAAE_RELEASE = "1"
} else {
    $env:VBAAE_RELEASE = "0"
    if ($dirty) {
        Write-Warning "Development/dirty build: git working tree is not clean. This cannot satisfy release qualification."
    }
}

Invoke-Checked "pip install" { pip install -r requirements-build.txt }
if ($Release) {
    Invoke-Checked "pip install dev extras" { pip install -e ".[dev]" }
    Invoke-Checked "pytest" { python -m pytest tests -q }
    Invoke-Checked "ruff" { python -m ruff check . }
    Invoke-Checked "pyright" { python -m pyright }
}

$demo = Join-Path $root "tests\fixtures\xlam\Demo.xlam"
$realXlam = Join-Path $root "tests\fixtures\xlam\RealAddin.xlam"
$realPpam = Join-Path $root "tests\fixtures\ppam\RealAddin.ppam"
$realPptm = Join-Path $root "tests\fixtures\pptm\RealPresentation.pptm"
$missing = @()
if (-not (Test-Path $demo)) { $missing += $demo }
if ($Release) {
    if (-not (Test-Path $realXlam)) { $missing += $realXlam }
    if (-not (Test-Path $realPpam)) { $missing += $realPpam }
    if (-not (Test-Path $realPptm)) { $missing += $realPptm }
}
if ($missing.Count -gt 0) {
    $list = $missing -join "; "
    if ($Release) {
        throw "Release builds require qualification fixtures. Missing: $list"
    }
    Write-Warning "Fixture(s) missing ($list); packaged smoke tests skipped. Live qualification is incomplete."
}

$modes = @()
if ($Release) {
    $modes = @("onedir", "onefile")
} elseif ($Onefile) {
    $modes = @("onefile")
} else {
    $modes = @("onedir")
}

foreach ($mode in $modes) {
    if ($mode -eq "onefile") {
        Invoke-Checked "pyinstaller onefile" { pyinstaller packaging/VBAAddinEditor.spec -- --onefile --noconfirm }
        $exe = Join-Path $root "dist\VBAAddinEditor.exe"
        $manifestName = "release-manifest.json"
    } else {
        Invoke-Checked "pyinstaller onedir" { pyinstaller packaging/VBAAddinEditor.spec --noconfirm }
        $exe = Join-Path $root "dist\VBAAddinEditor\VBAAddinEditor.exe"
        $manifestName = "release-manifest-onedir.json"
    }
    if (-not (Test-Path $exe)) { throw "Expected executable missing: $exe" }

    if (Test-Path $demo) {
        Invoke-Checked "packaged self-test ($mode)" { & $exe --self-test $demo }
        Invoke-Checked "packaged round-trip ($mode)" { & $exe --self-roundtrip $demo }
    } elseif ($Release) {
        throw "Release builds cannot skip packaged smoke tests."
    }

    $hash = (Get-FileHash $exe -Algorithm SHA256).Hash
    $label = if ($Release) { "qualified-release" } elseif ($dirty) { "development-dirty" } else { "development" }
    $manifest = @{
        executable = $exe
        sha256 = $hash
        mode = $mode
        identityFile = "dist/build-identity.json"
        qualification = $label
        release = [bool]$Release
        dirtyTree = $dirty
    } | ConvertTo-Json
    if ($Release) {
        $manifest | Out-File (Join-Path $root "dist\$manifestName")
        if ($mode -eq "onefile") {
            "$hash  $exe" | Out-File (Join-Path $root "dist\VBAAddinEditor.exe.sha256")
        }
        Write-Host "Release artifact $mode SHA-256: $hash"
    } else {
        $devName = if ($mode -eq "onefile") { "development-manifest.json" } else { "development-manifest-onedir.json" }
        $manifest | Out-File (Join-Path $root "dist\$devName")
        Write-Host "Development build $mode SHA-256: $hash ($label)"
    }
}

if ($Release) {
    Write-Host "Build OK: qualified onedir and onefile from the same clean commit."
} else {
    Write-Host "Build OK: development package. Not a qualified release."
}
