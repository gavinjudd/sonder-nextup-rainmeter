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

function ConvertTo-IcsUtc {
    param([DateTime]$Value)
    $Value.ToUniversalTime().ToString('yyyyMMddTHHmmssZ')
}

function New-RecurringMaster {
    param(
        [string]$Uid,
        [DateTime]$StartUtc,
        [string]$Summary
    )

    @"
BEGIN:VEVENT
UID:$Uid@example.invalid
DTSTART:$(ConvertTo-IcsUtc $StartUtc)
DTEND:$(ConvertTo-IcsUtc $StartUtc.AddMinutes(30))
RRULE:FREQ=WEEKLY;COUNT=4
SUMMARY:$Summary
END:VEVENT
"@
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

    function Invoke-PrimaryFixture {
        param(
            [string[]]$Events,
            [string[]]$SensitiveValues = @()
        )

        [IO.File]::WriteAllText((Join-Path $tempRoot 'primary.ics'), (New-Calendar $Events), (New-Object Text.UTF8Encoding($false)))
        [IO.File]::WriteAllText((Join-Path $tempRoot 'secrets.ini'), "ICS_URL=https://example.invalid/primary.ics`r`n", (New-Object Text.UTF8Encoding($false)))

        $fixtureOutput = @(& $powershell -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $tempRoot 'harness.ps1') 2>&1)
        if ($LASTEXITCODE -ne 0) { throw "Fetcher recurrence fixture failed: $($fixtureOutput -join ' ')" }

        $fixtureLog = [IO.File]::ReadAllText((Join-Path $tempRoot 'gcal_log.txt'))
        foreach ($privateValue in @('example.invalid', $tempRoot) + $SensitiveValues) {
            if ($fixtureLog.Contains($privateValue)) { throw 'Fetcher recurrence log exposed a fixture URL, title, or local path.' }
        }

        return @([IO.File]::ReadAllLines((Join-Path $tempRoot 'gcal_events.txt'), [Text.Encoding]::Unicode))
    }

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

    # Regression: Google exports a deleted occurrence as an EXDATE using the
    # series TZID. It must disappear on a successful refresh.
    $recurrenceUtc = [DateTime]::UtcNow.AddHours(2)
    $recurrenceUtc = $recurrenceUtc.AddTicks(-($recurrenceUtc.Ticks % [TimeSpan]::TicksPerSecond))
    $eastern = [TimeZoneInfo]::FindSystemTimeZoneById('Eastern Standard Time')
    $recurrenceEastern = [TimeZoneInfo]::ConvertTimeFromUtc([DateTime]::SpecifyKind($recurrenceUtc, [DateTimeKind]::Utc), $eastern)
    $easternStamp = $recurrenceEastern.ToString('yyyyMMddTHHmmss')
    $excludedTitle = 'Weekly exclusion regression'
    $excludedMaster = @"
BEGIN:VEVENT
UID:weekly-exdate@example.invalid
DTSTART;TZID=America/New_York:$easternStamp
DTEND;TZID=America/New_York:$($recurrenceEastern.AddMinutes(30).ToString('yyyyMMddTHHmmss'))
RRULE:FREQ=WEEKLY;COUNT=4
EXDATE;TZID=America/New_York:$easternStamp
SUMMARY:$excludedTitle
END:VEVENT
"@
    $exdateLines = @(Invoke-PrimaryFixture -Events @(
        $excludedMaster,
        (New-IcsEvent 'exdate-sentinel' $recurrenceUtc.AddHours(1) 'EXDATE sentinel')
    ) -SensitiveValues @($excludedTitle, 'EXDATE sentinel'))
    if (@($exdateLines | Where-Object { $_ -match [regex]::Escape($excludedTitle) }).Count -ne 0) {
        throw 'EXDATE failed to suppress the deleted recurring occurrence.'
    }
    if (@($exdateLines | Where-Object { $_ -match 'EXDATE sentinel' }).Count -ne 1) {
        throw 'EXDATE fixture did not write a fresh replacement output.'
    }
    Write-Host 'Recurring EXDATE deletion test passed.' -ForegroundColor Green

    # A STATUS:CANCELLED exception must suppress its original recurrence even
    # when the master appears first in the feed.
    $cancelTitle = 'Cancelled recurrence regression'
    $cancelMaster = New-RecurringMaster 'weekly-cancel' $recurrenceUtc $cancelTitle
    $cancelOverride = @"
BEGIN:VEVENT
UID:weekly-cancel@example.invalid
RECURRENCE-ID:$(ConvertTo-IcsUtc $recurrenceUtc)
DTSTART:$(ConvertTo-IcsUtc $recurrenceUtc)
STATUS:CANCELLED
SEQUENCE:1
SUMMARY:$cancelTitle
END:VEVENT
"@
    $cancelLines = @(Invoke-PrimaryFixture -Events @(
        $cancelMaster,
        $cancelOverride,
        (New-IcsEvent 'cancel-sentinel' $recurrenceUtc.AddHours(1) 'Cancel sentinel')
    ) -SensitiveValues @($cancelTitle, 'Cancel sentinel'))
    if (@($cancelLines | Where-Object { $_ -match [regex]::Escape($cancelTitle) }).Count -ne 0) {
        throw 'STATUS:CANCELLED failed to suppress the recurring occurrence.'
    }
    Write-Host 'Recurring cancellation test passed.' -ForegroundColor Green

    # A moved exception replaces the original slot exactly once.
    $movedTitle = 'Moved recurrence regression'
    $movedUtc = $recurrenceUtc.AddHours(3)
    $movedMaster = New-RecurringMaster 'weekly-move' $recurrenceUtc $movedTitle
    $movedOverride = @"
BEGIN:VEVENT
UID:weekly-move@example.invalid
RECURRENCE-ID:$(ConvertTo-IcsUtc $recurrenceUtc)
DTSTART:$(ConvertTo-IcsUtc $movedUtc)
DTEND:$(ConvertTo-IcsUtc $movedUtc.AddMinutes(30))
STATUS:CONFIRMED
SEQUENCE:1
SUMMARY:$movedTitle
END:VEVENT
"@
    $movedLines = @(Invoke-PrimaryFixture -Events @($movedMaster, $movedOverride) -SensitiveValues @($movedTitle))
    $movedMatches = @($movedLines | Where-Object { $_ -match [regex]::Escape($movedTitle) })
    if ($movedMatches.Count -ne 1) { throw "Expected one moved occurrence; got $($movedMatches.Count)." }
    $movedPrefix = $movedUtc.ToLocalTime().ToString('ddd M/d') + " $([char]0x2022) " + $movedUtc.ToLocalTime().ToString('h:mm tt')
    if ($movedMatches[0] -notlike "$movedPrefix*") {
        throw "Moved occurrence did not use the replacement start time. Expected prefix '$movedPrefix'; got '$($movedMatches[0])'."
    }
    Write-Host 'Moved recurrence test passed.' -ForegroundColor Green

    # Restore the two-source configuration used by the failure-preservation test.
    [IO.File]::WriteAllText((Join-Path $tempRoot 'secrets.ini'), "ICS_URLS=https://example.invalid/primary.ics;https://example.invalid/secondary.ics`r`n", (New-Object Text.UTF8Encoding($false)))

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
