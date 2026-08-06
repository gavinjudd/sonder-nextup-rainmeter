[CmdletBinding()]
param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutputPath) {
    $OutputPath = Join-Path $repoRoot 'sonder-nextup-rainmeter.zip'
}
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$repoRoot = [IO.Path]::GetFullPath($repoRoot)

if ([IO.Path]::GetExtension($OutputPath) -ne '.zip') {
    throw 'OutputPath must name a .zip file.'
}

$gitRoot = [IO.Path]::GetFullPath((Join-Path $repoRoot '.git'))
if ($OutputPath.StartsWith(($gitRoot + [IO.Path]::DirectorySeparatorChar), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Refusing to write a package inside .git.'
}

# The archive is built from HEAD. Refuse staged or unstaged tracked changes so
# the bytes we inspect below are exactly the bytes that git archive will emit.
$trackedChanges = @(git -C $repoRoot status --porcelain=v1 --untracked-files=no)
if ($LASTEXITCODE -ne 0) { throw 'git status failed; package creation requires Git.' }
if ($trackedChanges.Count -gt 0) {
    throw 'Refusing to package a dirty tracked worktree. Commit or restore tracked changes first.'
}

$allowlist = @(
    '.gitattributes',
    '.gitignore',
    'CHANGELOG.md',
    'CONTRIBUTING.md',
    'LICENSE',
    'LICENSES/CC-BY-NC-SA-3.0.txt',
    'NOTICE.md',
    'README.md',
    'SECURITY.md',
    'Scripts/install.ps1',
    'screenshots/preview.svg',
    'overlay/@Resources/gcal_fetch.ps1',
    'overlay/@Resources/secrets.example.ini',
    'overlay/Calendar/Calendar.ini',
    'overlay/Clock/Clock.ini'
)

$forbiddenPathPattern = '(?i)(^|/)(secrets(?:[.][^/]+)?[.]ini|gcal_(events|log|tmp)([.].*)?|[.]gcal_events[.].*|gcal_to_rainmeter[.]ps1)$'
$allowedExamplePath = 'overlay/@Resources/secrets.example.ini'
$tracked = @(git -C $repoRoot ls-files)
if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed; package creation requires Git.' }

foreach ($relativePath in $tracked) {
    $trackedPath = [IO.Path]::GetFullPath((Join-Path $repoRoot $relativePath))
    if ([string]::Equals($OutputPath, $trackedPath, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to overwrite tracked source file: $relativePath"
    }
}

$forbiddenTracked = @($tracked | Where-Object {
    $normalized = $_ -replace '\\','/'
    $normalized -ne $allowedExamplePath -and $normalized -match $forbiddenPathPattern
})
if ($forbiddenTracked.Count -gt 0) {
    throw "Refusing to package prohibited tracked paths: $($forbiddenTracked -join ', ')"
}

foreach ($relativePath in $allowlist) {
    if ($tracked -notcontains $relativePath) {
        throw "Allowlisted path is not tracked: $relativePath"
    }
}

$privatePattern = 'calendar[.]google[.]com/calendar/ical/.+/private-[^/]+/basic[.]ics'
$realIcsSettingPattern = '(?im)^\s*(?:ICS_URLS?|PrimaryUrl|SecondaryUrls?)\s*=\s*(?![^\r\n]*(?:example[.]invalid|calendar[.]example))[^\r\n#;]+'
$personalEmailPattern = '(?i)\b[A-Z0-9._%+-]+@(?!(?:example[.]invalid|users[.]noreply[.]github[.]com)\b)[A-Z0-9.-]+\.[A-Z]{2,}\b'
$textExtensions = @('.ini','.md','.ps1','.txt','.yml','.yaml','.json','.xml','.svg')
foreach ($relativePath in $tracked) {
    if ($textExtensions -notcontains [IO.Path]::GetExtension($relativePath).ToLowerInvariant()) { continue }
    $fullPath = Join-Path $repoRoot $relativePath
    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        $content = [IO.File]::ReadAllText($fullPath)
        if ($content -match $privatePattern -or $content -match $realIcsSettingPattern) {
            throw "Refusing to package a non-example calendar credential found in: $relativePath"
        }
        if ($content -match $personalEmailPattern) {
            throw "Refusing to package a non-example email address found in: $relativePath"
        }
    }
}

if (Test-Path -LiteralPath $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

Push-Location $repoRoot
try {
    & git archive --format=zip --output=$OutputPath HEAD -- @allowlist
    if ($LASTEXITCODE -ne 0) { throw "git archive failed with exit code $LASTEXITCODE" }
} finally {
    Pop-Location
}

Write-Host "Created safe overlay package: $OutputPath"
