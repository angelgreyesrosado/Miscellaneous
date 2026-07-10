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
    if (Test-Path $CommitMessageFile) {
        if ((Get-Item $CommitMessageFile).Length -gt 0 -or $Source) {
            exit 0
        }
    }

    $targetDir = 'VisualBasic'
    $changedFiles = @(git diff --cached --name-only -- $targetDir | Where-Object { $_ -and $_.Trim() })
    if ($changedFiles.Count -eq 0) {
        Set-Content -Path $CommitMessageFile -Value "chore: update $targetDir"
        exit 0
    }

    $summary = git diff --cached --stat -- $targetDir
    $aiMessage = ''

    $provider = $env:AI_PROVIDER
    $apiKey = $env:AI_API_KEY
    $apiUrl = $env:AI_API_URL
    $model = if ($env:AI_MODEL) { $env:AI_MODEL } else { 'gpt-4o-mini' }

    if ($provider -and $apiKey -and $apiUrl) {
        try {
            $payload = [ordered]@{
                model = $model
                messages = @(
                    [ordered]@{ role = 'system'; content = 'You write concise Git commit messages.' },
                    [ordered]@{ role = 'user'; content = @(
                        'Generate a concise commit message for these code changes. Return only one short subject line and no markdown.',
                        '',
                        'Changed files:',
                        ($changedFiles -join [Environment]::NewLine),
                        '',
                        'Diff summary:',
                        $summary
                    ) -join [Environment]::NewLine }
                )
            } | ConvertTo-Json -Depth 8

            $response = Invoke-RestMethod -Method Post -Uri $apiUrl -Headers @{ Authorization = "Bearer $apiKey"; 'Content-Type' = 'application/json' } -Body $payload
            if ($response.choices -and $response.choices[0].message.content) {
                $aiMessage = ($response.choices[0].message.content -replace '\s+', ' ').Trim()
            }
            elseif ($response.text) {
                $aiMessage = ($response.text -replace '\s+', ' ').Trim()
            }
        }
        catch {
            $aiMessage = ''
        }
    }

    if ([string]::IsNullOrWhiteSpace($aiMessage)) {
        $aiMessage = "chore: update $targetDir"
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $changelogFile = Join-Path $repoRoot 'VisualBasic/CHANGELOG.md'
    $changelogDir = Split-Path -Parent $changelogFile
    New-Item -ItemType Directory -Path $changelogDir -Force | Out-Null

    $changelogBody = if (Test-Path $changelogFile) {
        $existing = Get-Content -Path $changelogFile -Raw
        if ([string]::IsNullOrWhiteSpace($existing)) { '# Changelog' } else { $existing.TrimEnd() }
    }
    else {
        '# Changelog'
    }

    if ($changelogBody -notmatch '^# Changelog') {
        $changelogBody = @('# Changelog', '', $changelogBody) -join [Environment]::NewLine
    }

    $entryLines = @(
        '',
        "## $timestamp",
        '',
        '### Commit',
        '',
        "- **Subject:** $aiMessage",
        '- **Scope:** VisualBasic',
        '- **Files:**',
        ''
    )

    foreach ($file in $changedFiles) {
        $entryLines += "  - $file"
    }

    $entryLines += @('', '### Summary', '', '```text', $summary.Trim(), '```', '')
    $changelogBody = @($changelogBody.TrimEnd(), ($entryLines -join [Environment]::NewLine)) -join [Environment]::NewLine
    Set-Content -Path $changelogFile -Value $changelogBody.TrimEnd()

    $content = @(
        $aiMessage,
        '',
        'Auto-generated summary:',
        '',
        $summary.Trim(),
        '',
        "Markdown changelog: VisualBasic/CHANGELOG.md"
    )

    $messageDir = Split-Path -Parent $CommitMessageFile
    if ($messageDir) {
        New-Item -ItemType Directory -Path $messageDir -Force | Out-Null
    }

    Set-Content -Path $CommitMessageFile -Value ($content -join [Environment]::NewLine)
}
finally {
    Pop-Location
}
