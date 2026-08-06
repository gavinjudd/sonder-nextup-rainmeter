[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

function Add-Failure {
    param([string]$Message)
    [void]$failures.Add($Message)
}

$required = @(
    'README.md',
    'LICENSE',
    'LICENSES/CC-BY-NC-SA-3.0.txt',
    'NOTICE.md',
    'CHANGELOG.md',
    'CONTRIBUTING.md',
    'SECURITY.md',
    '.github/workflows/validate.yml',
    'screenshots/preview.svg',
    'Scripts/install.ps1',
    'Scripts/pack.ps1',
    'overlay/@Resources/gcal_fetch.ps1',
    'overlay/@Resources/secrets.example.ini',
    'overlay/Calendar/Calendar.ini',
    'overlay/Clock/Clock.ini',
    'Tests/test_fetcher.ps1',
    'Tests/test_installer.ps1'
)

$tracked = @(git -C $repoRoot ls-files)
if ($LASTEXITCODE -ne 0) { throw 'git ls-files failed.' }
$untracked = @(git -C $repoRoot ls-files --others --exclude-standard)
if ($LASTEXITCODE -ne 0) { throw 'git untracked-file scan failed.' }

foreach ($path in $required) {
    if ($tracked -notcontains $path) { Add-Failure "Required tracked path is missing: $path" }
}

$forbiddenPathPattern = '(?i)(^|/)(secrets(?:[.][^/]+)?[.]ini|gcal_(events|log|tmp)([.].*)?|[.]gcal_events[.].*|gcal_to_rainmeter[.]ps1)$'
$allowedExamplePath = 'overlay/@Resources/secrets.example.ini'
$allowedResourcePaths = @(
    'overlay/@Resources/gcal_fetch.ps1',
    $allowedExamplePath
)
foreach ($path in $tracked) {
    $normalized = $path -replace '\\','/'
    if ($normalized -ne $allowedExamplePath -and $normalized -match $forbiddenPathPattern) { Add-Failure "Prohibited path is tracked: $path" }
    if ($normalized.StartsWith('Sonder/', [StringComparison]::OrdinalIgnoreCase)) { Add-Failure "Legacy Sonder tree is tracked: $path" }
    if ($normalized.StartsWith('overlay/@Resources/', [StringComparison]::OrdinalIgnoreCase) -and $allowedResourcePaths -notcontains $normalized) {
        Add-Failure "Unexpected public resource path: $path"
    }
    if ([IO.Path]::GetExtension($normalized).ToLowerInvariant() -in @('.exe','.dll','.ttf','.otf')) { Add-Failure "Third-party binary is tracked: $path" }
    if ([IO.Path]::GetExtension($normalized).ToLowerInvariant() -eq '.ics' -and -not $normalized.StartsWith('Tests/', [StringComparison]::OrdinalIgnoreCase)) {
        Add-Failure "Raw calendar data is tracked outside Tests: $path"
    }
}

foreach ($path in $untracked) {
    $normalized = $path -replace '\\','/'
    if ($normalized -ne $allowedExamplePath -and $normalized -match $forbiddenPathPattern) { Add-Failure "Prohibited path is waiting to be added: $path" }
    if ([IO.Path]::GetExtension($normalized).ToLowerInvariant() -in @('.exe','.dll','.ttf','.otf')) { Add-Failure "Untracked binary is present: $path" }
}

$privateUrlPattern = 'calendar[.]google[.]com/calendar/ical/.+/private-[^/]+/basic[.]ics'
$realIcsSettingPattern = '(?im)^\s*(?:ICS_URLS?|PrimaryUrl|SecondaryUrls?)\s*=\s*(?![^\r\n]*(?:example[.]invalid|calendar[.]example))[^\r\n#;]+'
$personalEmailPattern = '(?i)\b[A-Z0-9._%+-]+@(?!(?:example[.]invalid|users[.]noreply[.]github[.]com)\b)[A-Z0-9.-]+\.[A-Z]{2,}\b'
$textExtensions = @('.ini','.inc','.md','.ps1','.txt','.yml','.yaml','.json','.xml','.svg')

function Test-TextContent {
    param(
        [string]$Context,
        [string]$Content
    )

    if ($Content -match $privateUrlPattern) { Add-Failure "Private Google Calendar URL pattern found in: $Context" }
    if ($Content -match $realIcsSettingPattern) { Add-Failure "Non-example ICS setting found in: $Context" }
    if ($Content -match $personalEmailPattern) { Add-Failure "Non-example email address found in: $Context" }
    if ($Content -match '(?i)C:\\Users\\[^\\\r\n]+' -or $Content -match '(?i)B:\\sonder-nextup-rainmeter') {
        Add-Failure "Hardcoded personal path found in: $Context"
    }
}

foreach ($path in $tracked) {
    if ($textExtensions -notcontains [IO.Path]::GetExtension($path).ToLowerInvariant()) { continue }
    $fullPath = Join-Path $repoRoot $path
    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        Test-TextContent "working tree $path" ([IO.File]::ReadAllText($fullPath))
    }

    $indexContent = ((& git -C $repoRoot show ":$path" 2>$null) -join "`n")
    if ($LASTEXITCODE -ne 0) {
        Add-Failure "Unable to inspect staged content: $path"
    } else {
        Test-TextContent "index $path" $indexContent
    }
}

foreach ($path in $untracked) {
    if ($textExtensions -notcontains [IO.Path]::GetExtension($path).ToLowerInvariant()) { continue }
    $fullPath = Join-Path $repoRoot $path
    if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
        Test-TextContent "untracked $path" ([IO.File]::ReadAllText($fullPath))
    }
}

foreach ($path in $tracked) {
    $fullPath = Join-Path $repoRoot $path
    if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf)) { continue }
    $bytes = [IO.File]::ReadAllBytes($fullPath)
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A) {
        Add-Failure "PE executable content found in: $path"
    }
    if ($bytes.Length -ge 4) {
        $magic = [BitConverter]::ToString($bytes[0..3])
        if ($magic -in @('89-50-4E-47','4F-54-54-4F','00-01-00-00')) {
            Add-Failure "Unexpected binary asset content found in: $path"
        }
    }
}

foreach ($scriptPath in Get-ChildItem -LiteralPath $repoRoot -Recurse -Filter '*.ps1' -File) {
    if ($scriptPath.FullName -match '[\\/][.]git[\\/]') { continue }
    $tokens = $null
    $parseErrors = $null
    [void][Management.Automation.Language.Parser]::ParseFile($scriptPath.FullName, [ref]$tokens, [ref]$parseErrors)
    foreach ($parseError in $parseErrors) {
        Add-Failure "PowerShell syntax error in $($scriptPath.FullName.Substring($repoRoot.Length + 1)): $($parseError.Message)"
    }
}

foreach ($relativePath in @('overlay/Calendar/Calendar.ini','overlay/Clock/Clock.ini')) {
    $fullPath = Join-Path $repoRoot $relativePath
    $content = [IO.File]::ReadAllText($fullPath)
    $sections = @([regex]::Matches($content, '(?m)^\[([^\]]+)\]\s*$') | ForEach-Object { $_.Groups[1].Value })
    $duplicates = @($sections | Group-Object | Where-Object { $_.Count -gt 1 })
    foreach ($duplicate in $duplicates) { Add-Failure "Duplicate Rainmeter section in ${relativePath}: $($duplicate.Name)" }
}

$calendar = [IO.File]::ReadAllText((Join-Path $repoRoot 'overlay/Calendar/Calendar.ini'))
foreach ($marker in @('[MeasureRunScript]','Plugin=RunCommand','OnWakeAction=','[MeasureFetchTimer]','UpdateDivider=5')) {
    if (-not $calendar.Contains($marker)) { Add-Failure "Calendar automation marker is missing: $marker" }
}
if ($calendar -notmatch '(?i)License=Creative Commons Attribution-NonCommercial-ShareAlike 3[.]0') {
    Add-Failure 'Calendar must retain its upstream CC BY-NC-SA 3.0 license notice.'
}
foreach ($marker in @('Original source: https://github.com/mpurses/Sonder/','Modified by Gavin Judd','NOTICE.md')) {
    if (-not $calendar.Contains($marker)) { Add-Failure "Calendar attribution marker is missing: $marker" }
}
if ($calendar -match '(?m)^@include') { Add-Failure 'Calendar overlay must not depend on redistributed Sonder resource files.' }

$clock = [IO.File]::ReadAllText((Join-Path $repoRoot 'overlay/Clock/Clock.ini'))
if ($clock -match '(?i)Streamster|Big John|Open Sans|Clock bg[.]png|@include') {
    Add-Failure 'Clean-room clock references a legacy Sonder/Mond resource.'
}
foreach ($marker in @('License=MIT','[Canvas]','[DividerAccent]','[Bridge]','[PeriodDisplay]')) {
    if (-not $clock.Contains($marker)) { Add-Failure "Clock redesign marker is missing: $marker" }
}

$example = [IO.File]::ReadAllText((Join-Path $repoRoot 'overlay/@Resources/secrets.example.ini'))
if ($example -notmatch 'example[.]invalid') { Add-Failure 'Safe secrets template must use example.invalid.' }

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Error $_ }
    exit 1
}

Write-Host "Repository validation passed ($($tracked.Count) tracked files)." -ForegroundColor Green
