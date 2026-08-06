[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$SonderPath = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Rainmeter\Skins\Sonder'),
    [switch]$SkipBackup
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$repoRoot = Split-Path -Parent $PSScriptRoot
$sourceRoot = Join-Path $repoRoot 'overlay'
$targetRoot = [IO.Path]::GetFullPath($SonderPath)

$officialMarkers = @(
    (Join-Path $targetRoot '@Resources\Variables.inc'),
    (Join-Path $targetRoot 'Calendar\Calendar.ini')
)

foreach ($marker in $officialMarkers) {
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) {
        throw "Official Sonder was not found at '$targetRoot'. Install it first or pass -SonderPath."
    }
}

$overlayFiles = @(
    'Calendar\Calendar.ini',
    'Clock\Clock.ini',
    '@Resources\gcal_fetch.ps1',
    '@Resources\secrets.example.ini'
)

foreach ($relativePath in $overlayFiles) {
    $source = Join-Path $sourceRoot $relativePath
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "Overlay source is missing: $source"
    }
}

$backupRoot = $null
if (-not $SkipBackup) {
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $targetRoot ("@Resources\Backups\sonder-nextup-$stamp")
}

foreach ($relativePath in $overlayFiles) {
    $source = Join-Path $sourceRoot $relativePath
    $destination = Join-Path $targetRoot $relativePath
    $destinationDirectory = Split-Path -Parent $destination

    if (-not (Test-Path -LiteralPath $destinationDirectory)) {
        if ($PSCmdlet.ShouldProcess($destinationDirectory, 'Create directory')) {
            New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
        }
    }

    if ($backupRoot -and (Test-Path -LiteralPath $destination -PathType Leaf)) {
        $backupDestination = Join-Path $backupRoot $relativePath
        $backupDirectory = Split-Path -Parent $backupDestination
        if ($PSCmdlet.ShouldProcess($destination, "Back up to $backupDestination")) {
            New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null
            Copy-Item -LiteralPath $destination -Destination $backupDestination -Force
        }
    }

    if ($PSCmdlet.ShouldProcess($destination, 'Install Sonder Next Up overlay file')) {
        Copy-Item -LiteralPath $source -Destination $destination -Force
    }
}

$secretsPath = Join-Path $targetRoot '@Resources\secrets.ini'
$examplePath = Join-Path $targetRoot '@Resources\secrets.example.ini'
if (-not (Test-Path -LiteralPath $secretsPath -PathType Leaf)) {
    if ($PSCmdlet.ShouldProcess($secretsPath, 'Create local calendar configuration')) {
        Copy-Item -LiteralPath $examplePath -Destination $secretsPath
    }
}

Write-Host ''
Write-Host 'Sonder Next Up installed successfully.' -ForegroundColor Cyan
Write-Host "Target: $targetRoot"
if ($backupRoot) { Write-Host "Backup: $backupRoot" }
Write-Host "Next: edit '$secretsPath', then refresh Sonder\Calendar and Sonder\Clock in Rainmeter."
