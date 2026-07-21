param(
    [string]$Range,
    [switch]$UseWorkingTree,
    [string]$ChangelogPath = "CHANGELOG.md"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-RepoRoot {
    $root = git rev-parse --show-toplevel 2>$null
    if (-not $root) {
        throw "Run this script from inside a Git repository."
    }

    return $root.Trim()
}

function Get-ChangeEntries {
    param(
        [string]$RepoRoot,
        [string]$Range,
        [switch]$UseWorkingTree
    )

    $entries = @()

    if ($UseWorkingTree) {
        $rawLines = git -C $RepoRoot status --porcelain
        foreach ($line in $rawLines) {
            if ([string]::IsNullOrWhiteSpace($line)) {
                continue
            }

            $status = $line.Substring(0, 2)
            $path = $line.Substring(3)
            $changeType = "Modified"

            if ($status -match '\?\?') {
                $changeType = "Added"
            }
            elseif ($status -match '^D') {
                $changeType = "Deleted"
            }
            elseif ($status -match '^R') {
                $changeType = "Renamed"
            }

            $entries += [pscustomobject]@{
                Path = $path
                ChangeType = $changeType
            }
        }

        return $entries
    }

    if ($Range) {
        $rawLines = git -C $RepoRoot diff --name-status --relative $Range
    }
    else {
        $rawLines = git -C $RepoRoot diff --name-status --relative HEAD~1..HEAD 2>$null
    if (-not $rawLines) {
        $rawLines = git -C $RepoRoot status --porcelain
    }
    }

    foreach ($line in $rawLines) {
        if ([string]::IsNullOrWhiteSpace($line)) {
            continue
        }

        $parts = $line -split "`t"
        if ($parts.Count -eq 0) {
            continue
        }

        $changeType = switch ($parts[0]) {
            "A" { "Added" }
            "D" { "Deleted" }
            "R" { "Renamed" }
            default { "Modified" }
        }

        $path = if ($parts.Count -ge 2) { $parts[-1] } else { $line }
        $entries += [pscustomobject]@{
            Path = $path
            ChangeType = $changeType
        }
    }

    return $entries
}

function Get-ChangeGroup {
    param([string]$Path)

    $segments = @($Path -split '/').Where({ $_ -ne '' })
    if ($segments.Count -eq 0) {
        return "Root"
    }

    if ($segments.Count -eq 1) {
        return $segments[0]
    }

    return "$($segments[0]) > $($segments[1])"
}

$repoRoot = Get-RepoRoot
$targetPath = Join-Path $repoRoot $ChangelogPath
$entries = Get-ChangeEntries -RepoRoot $repoRoot -Range $Range -UseWorkingTree:$UseWorkingTree

if ($entries.Count -eq 0) {
    Write-Host "No file changes detected. Changelog was not updated."
    return
}

$grouped = $entries | Group-Object -Property { Get-ChangeGroup $_.Path }
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
$dateLabel = Get-Date -Format "yyyy-MM-dd"

$sectionLines = [System.Collections.Generic.List[string]]::new()
$sectionLines.Add("## $dateLabel")
$sectionLines.Add("")
$sectionLines.Add("### Summary")
$sectionLines.Add("- Generated from Git changes at $timestamp.")
$sectionLines.Add("")

foreach ($group in ($grouped | Sort-Object Name)) {
    $sectionLines.Add("### $($group.Name)")
    foreach ($entry in ($group.Group | Sort-Object Path)) {
        $sectionLines.Add("- $($entry.ChangeType): [$($entry.Path)]($($entry.Path))")
    }
    $sectionLines.Add("")
}

$sectionText = ($sectionLines -join "`n").TrimEnd()

if (Test-Path $targetPath) {
    $existing = [System.IO.File]::ReadAllText($targetPath, [System.Text.Encoding]::UTF8)
}
else {
    $existing = "# Changelog`n"
}

if ($existing -match "(?ms)^## $([regex]::Escape($dateLabel))\s*$") {
    $pattern = "(?ms)^## $([regex]::Escape($dateLabel)).*?(?=^## |\z)"
    $updated = [regex]::Replace($existing, $pattern, $sectionText, 1)
}
else {
    $updated = $existing.TrimEnd()
    if ($updated.Length -gt 0) {
        $updated = "$updated`n`n"
    }
    else {
        $updated = "# Changelog`n"
    }

    $updated = "$updated$sectionText`n"
}

[System.IO.File]::WriteAllText($targetPath, $updated, [System.Text.Encoding]::UTF8)
Write-Host "Updated $ChangelogPath with $($entries.Count) change entries."
