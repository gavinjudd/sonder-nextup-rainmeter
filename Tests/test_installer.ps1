[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = Split-Path -Parent $PSScriptRoot
$installer = Join-Path $repoRoot 'Scripts\install.ps1'
$overlayRoot = Join-Path $repoRoot 'overlay'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$tempRoot = [IO.Path]::GetFullPath((Join-Path $tempBase ('sonder-nextup-installer-' + [Guid]::NewGuid().ToString('N'))))
$whatIfRoot = [IO.Path]::GetFullPath((Join-Path $tempBase ('sonder-nextup-installer-whatif-' + [Guid]::NewGuid().ToString('N'))))

function Initialize-FakeSonder {
    param([string]$Root)

    New-Item -ItemType Directory -Path (Join-Path $Root '@Resources') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $Root 'Calendar') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $Root 'Clock') -Force | Out-Null
    [IO.File]::WriteAllText((Join-Path $Root '@Resources\Variables.inc'), '[Variables]', (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $Root 'Calendar\Calendar.ini'), 'original-calendar', (New-Object Text.UTF8Encoding($false)))
}

try {
    Initialize-FakeSonder $tempRoot
    [IO.File]::WriteAllText((Join-Path $tempRoot 'Clock\Clock.ini'), 'original-clock', (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $tempRoot '@Resources\gcal_fetch.ps1'), 'original-fetcher', (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $tempRoot '@Resources\secrets.example.ini'), 'original-example', (New-Object Text.UTF8Encoding($false)))
    $privateSentinel = '[Calendar]' + "`r`n" + 'ICS_URLS=https://example.invalid/preserve-this-value'
    [IO.File]::WriteAllText((Join-Path $tempRoot '@Resources\secrets.ini'), $privateSentinel, (New-Object Text.UTF8Encoding($false)))

    & $installer -SonderPath $tempRoot

    foreach ($relativePath in @('Calendar\Calendar.ini','Clock\Clock.ini','@Resources\gcal_fetch.ps1','@Resources\secrets.example.ini')) {
        $expected = (Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $overlayRoot $relativePath)).Hash
        $actual = (Get-FileHash -Algorithm SHA256 -LiteralPath (Join-Path $tempRoot $relativePath)).Hash
        if ($actual -ne $expected) { throw "Installed file hash mismatch: $relativePath" }
    }

    if ([IO.File]::ReadAllText((Join-Path $tempRoot '@Resources\secrets.ini')) -ne $privateSentinel) {
        throw 'Installer replaced an existing secrets.ini.'
    }

    $backupRoots = @(Get-ChildItem -LiteralPath (Join-Path $tempRoot '@Resources\Backups') -Directory)
    if ($backupRoots.Count -ne 1) { throw "Expected one timestamped backup; got $($backupRoots.Count)." }
    foreach ($relativePath in @('Calendar\Calendar.ini','Clock\Clock.ini','@Resources\gcal_fetch.ps1','@Resources\secrets.example.ini')) {
        if (-not (Test-Path -LiteralPath (Join-Path $backupRoots[0].FullName $relativePath) -PathType Leaf)) {
            throw "Backup is missing: $relativePath"
        }
    }

    Initialize-FakeSonder $whatIfRoot
    & $installer -SonderPath $whatIfRoot -WhatIf
    if (Test-Path -LiteralPath (Join-Path $whatIfRoot 'Clock\Clock.ini') -PathType Leaf) {
        throw 'Installer -WhatIf wrote an overlay file.'
    }
    if (Test-Path -LiteralPath (Join-Path $whatIfRoot '@Resources\Backups')) {
        throw 'Installer -WhatIf created a backup directory.'
    }

    Write-Host 'Installer fixture and -WhatIf test passed.' -ForegroundColor Green
} finally {
    foreach ($candidate in @($tempRoot, $whatIfRoot)) {
        if ((Test-Path -LiteralPath $candidate) -and
            $candidate.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
            (Split-Path $candidate -Leaf) -like 'sonder-nextup-installer-*') {
            Remove-Item -LiteralPath $candidate -Recurse -Force -ErrorAction SilentlyContinue
        } elseif (Test-Path -LiteralPath $candidate) {
            throw "Refusing to remove unexpected test path: $candidate"
        }
    }
}
