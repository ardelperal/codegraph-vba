# Installs the canonical `codegraph` shim for the codegraph-vba fork (Windows).
#
# Tools that expect the canonical `codegraph` command (for example the
# gentle-ai installer) keep working: `codegraph` resolves to the
# @aroman22/codegraph-vba npm install through scripts/compat/codegraph-launcher.go.
#
#   pwsh scripts/compat/install-codegraph-shim.ps1
#
# Requires Go on PATH to build the launcher. Environment:
#   CODEGRAPH_SHIM_DIR  install location (default: %LOCALAPPDATA%\codegraph-vba-compat\bin)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command go -ErrorAction SilentlyContinue)) {
  throw 'codegraph shim: Go is required on PATH to build the launcher.'
}

$source = Join-Path $PSScriptRoot 'codegraph-launcher.go'
$binDir = if ($env:CODEGRAPH_SHIM_DIR) { $env:CODEGRAPH_SHIM_DIR } else { Join-Path $env:LOCALAPPDATA 'codegraph-vba-compat\bin' }
New-Item -ItemType Directory -Force -Path $binDir | Out-Null

$target = Join-Path $binDir 'codegraph.exe'
& go build -o $target $source
if ($LASTEXITCODE -ne 0) { throw 'codegraph shim: go build failed.' }
Write-Host "Built $target"

$userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
if (($userPath -split ';') -notcontains $binDir) {
  [Environment]::SetEnvironmentVariable('Path', "$binDir;$userPath", 'User')
  Write-Host "Added $binDir to your PATH (restart your terminal to pick it up)."
}

Write-Host 'Check: codegraph --version'
