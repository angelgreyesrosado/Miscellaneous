Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = git rev-parse --show-toplevel 2>$null
if (-not $repoRoot) {
    throw "Run this script from inside a Git repository."
}

$hookPath = Join-Path $repoRoot ".githooks"
if (-not (Test-Path $hookPath)) {
    New-Item -ItemType Directory -Path $hookPath -Force | Out-Null
}

git -C $repoRoot config core.hooksPath .githooks
Write-Host "Git hooks are configured to use $hookPath"
