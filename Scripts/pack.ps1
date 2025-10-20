$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$zip = Join-Path $root 'sonder-nextup-rainmeter.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }

# --- Preferred: use git archive (packages only tracked files; respects .gitignore) ---
$gitOk = $false
try {
  $null = git --version 2>$null
  if ($LASTEXITCODE -eq 0) { $gitOk = $true }
} catch {}

if ($gitOk) {
  git archive --format=zip --output $zip HEAD
  Write-Host "Packed via git archive -> $zip"
  exit 0
}

# --- Fallback: Compress-Archive with absolute paths ---
$excludes = @(
  'Sonder\@Resources\secrets.ini',
  'Sonder\@Resources\gcal_events.txt',
  'Sonder\@Resources\gcal_log.txt'
)

# Build list, skipping .git and excluded runtime files
$files = Get-ChildItem -Recurse -File | Where-Object {
  $rel = $_.FullName.Substring($root.Length + 1)
  ($excludes -notcontains $rel) -and ($rel -notmatch '^[.]git\\')
}

# IMPORTANT: pass FullName to Compress-Archive (not just the filename)
$fullPaths = $files | Select-Object -ExpandProperty FullName
Compress-Archive -Path $fullPaths -DestinationPath $zip -Force
Write-Host "Packed via Compress-Archive -> $zip"
