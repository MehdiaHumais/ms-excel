<#
.SYNOPSIS
  Builds the BritSheet Windows desktop app (.exe + .msi + NSIS installer).

.DESCRIPTION
  Wraps `npx tauri build` with the environment this machine needs:

  - MinGW-w64 on PATH. rustc's bundled dlltool cannot create import libraries on
    Windows (it passes a --temp-prefix containing a colon, which is not a legal
    filename), so we use a real binutils from $MinGWHome and tell rustc not to
    use its self-contained one via -C link-self-contained=no.
  - TEMP/TMP and CARGO_TARGET_DIR on a drive with free space. Rust link steps and
    cargo throw off large temporaries and would otherwise fill the system drive.

.PARAMETER DebugBuild
  Build the debug profile instead of release.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File scripts\build-desktop.ps1
#>
[CmdletBinding()]
param([switch]$DebugBuild)

$ErrorActionPreference = 'Stop'

$MinGWHome = 'D:\mingw64\mingw64'
if (-not (Test-Path -LiteralPath "$MinGWHome\bin\gcc.exe")) {
  # Fall back to a known MinGW-w64 installation (D:\mingw64 can be wiped by
  # drive cleanups), or one supplied via the MINGW_HOME env var.
  $candidates = @($env:MINGW_HOME, 'D:\britpdf-tools\mingw64')
  foreach ($c in $candidates) {
    if ($c -and (Test-Path -LiteralPath "$c\bin\gcc.exe")) { $MinGWHome = $c; break }
  }
}
$Scratch   = 'D:\build-tmp'
$repoRoot  = Split-Path -Parent $PSScriptRoot

if (-not (Test-Path -LiteralPath "$MinGWHome\bin\gcc.exe")) {
  throw "MinGW-w64 not found (tried $MinGWHome). Set `$MinGWHome / MINGW_HOME, or extract it there."
}

New-Item -ItemType Directory -Path $Scratch -Force | Out-Null

$env:PATH               = "$MinGWHome\bin;$env:PATH"
$env:TEMP               = $Scratch
$env:TMP                = $Scratch
# Use the MinGW binutils/gcc on PATH instead of rustc's bundled copies.
$env:RUSTFLAGS          = '-C link-self-contained=no'
# Must NOT contain spaces: MinGW's windres (used to compile the Windows
# resource file) fails on paths with spaces, and this repo lives under
# "...\msofficeclone\ms excel".
$env:CARGO_TARGET_DIR   = "$Scratch\tauri-target"

$releaseDir   = Join-Path $env:CARGO_TARGET_DIR 'release'
$stagedDll    = Join-Path $repoRoot 'src-tauri\WebView2Loader.dll'

function Invoke-TauriBuild([string[]]$Extra) {
  & npx @('tauri', 'build') @Extra
  if ($LASTEXITCODE -ne 0) { throw "tauri build failed with exit code $LASTEXITCODE" }
}

Write-Host "MinGW:   $MinGWHome" -ForegroundColor DarkGray
Write-Host "Target:  $env:CARGO_TARGET_DIR" -ForegroundColor DarkGray
Write-Host ''

Push-Location $repoRoot
try {
  # Pass 1: compile only. webview2-com-sys drops WebView2Loader.dll into the
  # target dir, and the installers do NOT ship it, so an installed copy dies
  # with "WebView2Loader.dll was not found". Stage it into src-tauri/ where
  # bundle.resources can pick it up.
  $bundleArgs = if ($DebugBuild) { @('--debug') } else { @() }
  Invoke-TauriBuild (@('--no-bundle') + $bundleArgs)

  $builtDll = Join-Path $releaseDir 'WebView2Loader.dll'
  if (-not (Test-Path -LiteralPath $builtDll)) {
    throw "Expected WebView2Loader.dll in $releaseDir but it was not produced."
  }
  Copy-Item -LiteralPath $builtDll -Destination $stagedDll -Force
  Write-Host "Staged WebView2Loader.dll -> src-tauri\" -ForegroundColor DarkGray

  # Pass 2: bundle, now including the loader.
  Invoke-TauriBuild $bundleArgs
}
finally {
  Pop-Location
}

$bundle = Join-Path $env:CARGO_TARGET_DIR 'release\bundle'
Write-Host ''
Write-Host 'Build artifacts:' -ForegroundColor Green
Get-ChildItem $releaseDir -Filter '*.exe' -ErrorAction SilentlyContinue |
  ForEach-Object { Write-Host "  $($_.FullName)  ($([math]::Round($_.Length/1MB,1)) MB)" -ForegroundColor Green }
if (Test-Path -LiteralPath $bundle) {
  Get-ChildItem $bundle -Recurse -Include '*.msi','*.exe' -ErrorAction SilentlyContinue |
    ForEach-Object { Write-Host "  $($_.FullName)  ($([math]::Round($_.Length/1MB,1)) MB)" -ForegroundColor Green }
}