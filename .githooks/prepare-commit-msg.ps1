param(
    [Parameter(Position = 0)]
    [string]$CommitMessageFile,

    [Parameter(Position = 1)]
    [string]$Source
)

$ErrorActionPreference = 'Stop'

if (-not $CommitMessageFile) {
    exit 0
}

$repoRoot = git rev-parse --show-toplevel 2>$null
if (-not $repoRoot) {
    exit 0
}

Push-Location $repoRoot
try {
    $existingCommitMessage = ''
    $shouldGenerateMessage = $true

    if (Test-Path $CommitMessageFile) {
        $existingCommitMessage = Get-Content -Path $CommitMessageFile -Raw
        if (-not [string]::IsNullOrWhiteSpace($existingCommitMessage) -or $Source) {
            $shouldGenerateMessage = $false
        }
    }

    $changedFiles = git diff --cached --name-only --diff-filter=ACMRTUXB -- . | Where-Object { $_ -and $_.Trim() -ne 'CHANGELOG.md' }
    if ($changedFiles.Count -eq 0) {
        if ($shouldGenerateMessage -and -not [string]::IsNullOrWhiteSpace($existingCommitMessage)) {
            Set-Content -Path $CommitMessageFile -Value $existingCommitMessage
        }
        exit 0
    }

    $diffStat = git diff --cached --stat -- . | Out-String
    $diffStat = $diffStat.Trim()
    $aiMessage = ''
    $aiSummary = ''

    $provider = $env:AI_PROVIDER
    $apiKey = $env:AI_API_KEY
    $apiUrl = $env:AI_API_URL
    $model = if ($env:AI_MODEL) { $env:AI_MODEL } else { 'gpt-4o-mini' }

    if ($provider -and $apiKey -and $apiUrl) {
        try {
            $prompt = @"
You are an AI assistant that writes concise git commit subjects and short changelog summaries.
Use the staged git diff summary and changed file list below.
Return a short subject line on the first line and a one-paragraph summary on the following lines.
Do not include markdown formatting.

Changed files:
$($changedFiles -join "`n")

Diff summary:
$diffStat
"@

            $payload = [ordered]@{
                model = $model
                messages = @(
                    [ordered]@{ role = 'system'; content = 'You write concise Git commit subjects and short changelog summaries.' },
                    [ordered]@{ role = 'user'; content = $prompt }
                )
                max_tokens = 400
            } | ConvertTo-Json -Depth 8

            $response = Invoke-RestMethod -Method Post -Uri $apiUrl -Headers @{ Authorization = "Bearer $apiKey"; 'Content-Type' = 'application/json' } -Body $payload
            $raw = $null

            if ($response.choices -and $response.choices[0].message.content) {
                $raw = $response.choices[0].message.content.Trim()
            }
            elseif ($response.choices -and $response.choices[0].text) {
                $raw = $response.choices[0].text.Trim()
            }
            elseif ($response.text) {
                $raw = $response.text.Trim()
            }

            if ($raw) {
                try {
                    $json = $raw | ConvertFrom-Json
                    if ($json.subject) { $aiMessage = $json.subject.Trim() }
                    if ($json.summary) { $aiSummary = $json.summary.Trim() }
                }
                catch {
                    $lines = $raw -split '\r?\n'
                    $aiMessage = ($lines[0] -replace '\s+', ' ').Trim()
                    if ($lines.Count -gt 1) {
                        $aiSummary = ($lines[1..($lines.Count - 1)] -join "`n").Trim()
                    }
                    else {
                        $aiSummary = $raw
                    }
                }
            }
        }
        catch {
            $aiMessage = ''
            $aiSummary = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($aiMessage)) {
        $aiMessage = 'chore: update repository changes'
    }
    if ([string]::IsNullOrWhiteSpace($aiSummary)) {
        $aiSummary = $diffStat
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $branchName = git rev-parse --abbrev-ref HEAD 2>$null
    if ([string]::IsNullOrWhiteSpace($branchName)) { $branchName = 'unknown' }

    $changelogFile = Join-Path $repoRoot 'CHANGELOG.md'
    if (-not (Test-Path $changelogFile)) {
        '# Changelog' | Set-Content -Path $changelogFile
    }

    $existingChangelog = Get-Content -Path $changelogFile -Raw
    if ($existingChangelog -notmatch '^# Changelog') {
        $existingChangelog = "# Changelog`n`n$existingChangelog"
    }

    $folderGroups = [ordered]@{}
    $extensionGroups = [ordered]@{}
    foreach ($file in $changedFiles) {
        $parts = $file -split '[\\/]'
        $topFolder = if ($parts.Count -gt 1) { $parts[0] } else { '.' }
        if (-not $folderGroups.ContainsKey($topFolder)) { $folderGroups[$topFolder] = @() }
        $folderGroups[$topFolder] += $file

        $extension = [IO.Path]::GetExtension($file).ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($extension)) { $extension = '(no ext)' }
        if (-not $extensionGroups.ContainsKey($extension)) { $extensionGroups[$extension] = 0 }
        $extensionGroups[$extension]++
    }

    $entry = @()
    $entry += "## $timestamp ($branchName)"
    $entry += ''
    $entry += '### Summary'
    $entry += ''
    $entry += $aiSummary
    $entry += ''
    $entry += '### Files changed'
    $entry += ''
    foreach ($folder in $folderGroups.Keys) {
        $entry += "- $folder ($($folderGroups[$folder].Count) files)"
        foreach ($file in $folderGroups[$folder]) {
            $entry += "  - $file"
        }
        $entry += ''
    }

    $entry += '### File types'
    $entry += ''
    foreach ($extension in $extensionGroups.Keys) {
        $entry += "- ${extension}: $($extensionGroups[$extension])"
    }
    $entry += ''
    $entry += '### Git diff summary'
    $entry += ''
    $entry += '```text'
    $entry += $diffStat
    $entry += '```'
    $entry += ''

    $newBody = @($existingChangelog.TrimEnd(), ($entry -join "`n")) -join "`n"
    Set-Content -Path $changelogFile -Value $newBody.TrimEnd()
    git add --force -- $changelogFile

    if ($shouldGenerateMessage) {
        $messageLines = @(
            $aiMessage,
            '',
            'Auto-generated changelog entry: CHANGELOG.md'
        )
        Set-Content -Path $CommitMessageFile -Value ($messageLines -join "`n")
    }
    else {
        Set-Content -Path $CommitMessageFile -Value $existingCommitMessage
    }
}
finally {
    Pop-Location
}
