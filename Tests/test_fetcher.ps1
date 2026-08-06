[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = Split-Path -Parent $PSScriptRoot
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = [IO.Path]::GetFullPath((Join-Path $tempBase ("sonder-nextup-test-" + [Guid]::NewGuid().ToString('N'))))
New-Item -ItemType Directory -Path $tempRoot | Out-Null

function New-IcsEvent {
    param(
        [string]$Uid,
        [DateTime]$StartUtc,
        [string]$Summary
    )

    @"
BEGIN:VEVENT
UID:$Uid@example.invalid
DTSTART:$($StartUtc.ToUniversalTime().ToString('yyyyMMddTHHmmssZ'))
DTEND:$($StartUtc.ToUniversalTime().AddMinutes(30).ToString('yyyyMMddTHHmmssZ'))
SUMMARY:$Summary
END:VEVENT
"@
}

function New-Calendar {
    param([string[]]$Events)
    "BEGIN:VCALENDAR`r`nVERSION:2.0`r`nPRODID:-//sonder-nextup-rainmeter//offline test//EN`r`n" +
        ($Events -join "`r`n") + "`r`nEND:VCALENDAR`r`n"
}

try {
    Copy-Item -LiteralPath (Join-Path $repoRoot 'overlay\@Resources\gcal_fetch.ps1') -Destination $tempRoot

    $base = [DateTime]::UtcNow.AddMinutes(10)
    $primary = New-Calendar @(
        (New-IcsEvent 'primary-later' $base.AddHours(4) 'Primary later'),
        (New-IcsEvent 'primary-after' $base.AddHours(5) 'Primary after'),
        (New-IcsEvent 'primary-last' $base.AddHours(6) 'Primary last')
    )
    $secondary = New-Calendar @(
        (New-IcsEvent 'secondary-sooner' $base.AddHours(1) 'Secondary sooner'),
        (New-IcsEvent 'secondary-next' $base.AddHours(2) 'Secondary next')
    )

    [IO.File]::WriteAllText((Join-Path $tempRoot 'primary.ics'), $primary, (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $tempRoot 'secondary.ics'), $secondary, (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $tempRoot 'secrets.ini'), "ICS_URLS=https://example.invalid/primary.ics;https://example.invalid/secondary.ics`r`n", (New-Object Text.UTF8Encoding($false)))

    $harness = @'
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
function Invoke-WebRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [switch]$UseBasicParsing,
        [int]$TimeoutSec,
        [hashtable]$Headers
    )
    $fixture = if ($Uri -match 'secondary') { 'secondary.ics' } else { 'primary.ics' }
    [pscustomobject]@{ Content = [IO.File]::ReadAllText((Join-Path $PSScriptRoot $fixture)) }
}
& (Join-Path $PSScriptRoot 'gcal_fetch.ps1')
exit $LASTEXITCODE
'@
    [IO.File]::WriteAllText((Join-Path $tempRoot 'harness.ps1'), $harness, (New-Object Text.UTF8Encoding($false)))

    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $testOutput = @(& $powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $tempRoot 'harness.ps1') 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Fetcher fixture run failed: $($testOutput -join ' ')" }

    $lines = @([IO.File]::ReadAllLines((Join-Path $tempRoot 'gcal_events.txt'), [Text.Encoding]::Unicode))
    if ($lines.Count -ne 4) { throw "Expected 4 output lines; got $($lines.Count)." }

    $expected = @('Secondary sooner','Secondary next','Primary later','Primary after')
    for ($i = 0; $i -lt $expected.Count; $i++) {
        if ($lines[$i] -notmatch [regex]::Escape($expected[$i])) {
            throw "Chronological selection failed at row $($i + 1)."
        }
    }

    $log = [IO.File]::ReadAllText((Join-Path $tempRoot 'gcal_log.txt'))
    foreach ($privateValue in @('example.invalid','Secondary sooner','Secondary next','Primary later','Primary after',$tempRoot)) {
        if ($log.Contains($privateValue)) { throw 'Fetcher log exposed fixture URL, title, or local path.' }
    }

    Write-Host 'Offline multi-feed chronology test passed.' -ForegroundColor Green

    $lastKnownGood = 'Last known good fixture output'
    [IO.File]::WriteAllText((Join-Path $tempRoot 'gcal_events.txt'), $lastKnownGood, [Text.Encoding]::Unicode)

    $partialHarness = @'
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
function Invoke-WebRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$Uri,
        [switch]$UseBasicParsing,
        [int]$TimeoutSec,
        [hashtable]$Headers
    )
    if ($Uri -match 'secondary') { throw 'Synthetic source failure' }
    [pscustomobject]@{ Content = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'primary.ics')) }
}
& (Join-Path $PSScriptRoot 'gcal_fetch.ps1')
exit $LASTEXITCODE
'@
    [IO.File]::WriteAllText((Join-Path $tempRoot 'partial-harness.ps1'), $partialHarness, (New-Object Text.UTF8Encoding($false)))
    $partialOutput = @(& $powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $tempRoot 'partial-harness.ps1') 2>&1)
    if ($LASTEXITCODE -ne 1) { throw "Expected partial-source failure exit code 1; got $LASTEXITCODE. $($partialOutput -join ' ')" }
    if ([IO.File]::ReadAllText((Join-Path $tempRoot 'gcal_events.txt'), [Text.Encoding]::Unicode) -ne $lastKnownGood) {
        throw 'Partial-source failure replaced the last-known-good output.'
    }
    Write-Host 'Partial-source preservation test passed.' -ForegroundColor Green

    Remove-Item -LiteralPath (Join-Path $tempRoot 'secrets.ini') -Force
    $missingConfigOutput = @(& $powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $tempRoot 'gcal_fetch.ps1') 2>&1)
    if ($LASTEXITCODE -ne 2) { throw "Expected missing-config exit code 2; got $LASTEXITCODE. $($missingConfigOutput -join ' ')" }
    if ([IO.File]::ReadAllText((Join-Path $tempRoot 'gcal_events.txt'), [Text.Encoding]::Unicode) -ne $lastKnownGood) {
        throw 'Missing configuration replaced the last-known-good output.'
    }
    Write-Host 'Missing-config preservation test passed.' -ForegroundColor Green

    [IO.File]::WriteAllText((Join-Path $tempRoot 'primary.ics'), (New-Calendar @()), (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $tempRoot 'secrets.ini'), "ICS_URL=https://example.invalid/primary.ics`r`n", (New-Object Text.UTF8Encoding($false)))
    Remove-Item -LiteralPath (Join-Path $tempRoot 'gcal_events.txt') -Force
    $emptyOutput = @(& $powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $tempRoot 'harness.ps1') 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Expected empty-calendar exit code 0; got $LASTEXITCODE. $($emptyOutput -join ' ')" }
    if ([IO.File]::ReadAllText((Join-Path $tempRoot 'gcal_events.txt'), [Text.Encoding]::Unicode).Length -ne 0) {
        throw 'Empty calendar emitted a display sentinel instead of an empty output.'
    }
    Write-Host 'Empty-calendar output test passed.' -ForegroundColor Green
} finally {
    if ($tempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $tempRoot -Leaf) -like 'sonder-nextup-test-*') {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        throw "Refusing to remove unexpected test path: $tempRoot"
    }
}
