# ----------------- CONFIG -----------------
# Paste your Google "Secret address in iCal format" (basic.ics) between the quotes:
$icsUrl = "https://calendar.google.com/calendar/ical/gavinjudd22%40gmail.com/private-9600e68796d473d8b85e34e8a6cdb2a8/basic.ics"

# Output files:
$root   = "$env:USERPROFILE\Documents\Rainmeter\Skins\Sonder\@Resources"
$outTxt = Join-Path $root 'gcal_events.txt'
$log    = Join-Path $root 'gcal_log.txt'
# -----------------------------------------

# Ensure TLS1.2 for Invoke-WebRequest on older PS builds
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log($msg) {
  $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  "$ts  $msg" | Add-Content -Path $log -Encoding utf8
}

try {
  New-Item -ItemType Directory -Force -Path $root | Out-Null
} catch { }

Write-Log "----- Run start -----"
Write-Log "Using ICS URL length: $($icsUrl.Length)"

# 1) Download ICS
$content = ""
try {
  $resp = Invoke-WebRequest -Uri $icsUrl -UseBasicParsing -TimeoutSec 20
  $content = ($resp.Content -replace "`r","")
  Write-Log "Downloaded ICS bytes: $($resp.RawContentLength)"
} catch {
  Write-Log "Invoke-WebRequest failed: $($_.Exception.Message)"
  try {
    $wc = New-Object System.Net.WebClient
    $content = ($wc.DownloadString($icsUrl) -replace "`r","")
    Write-Log "WebClient fallback succeeded. Length: $($content.Length)"
  } catch {
    Write-Log "WebClient failed: $($_.Exception.Message)"
    $content = ""
  }
}

if ([string]::IsNullOrWhiteSpace($content)) {
  Write-Log "Empty ICS content. Writing placeholder line."
  "No upcoming events" | Set-Content -Path $outTxt -Encoding utf8
  Write-Log "----- Run end (empty) -----"
  exit
}

# 2) Parse VEVENT blocks
$blocks = ($content -split "BEGIN:VEVENT") | Select-Object -Skip 1
$events = @()
foreach ($b in $blocks) {
  $summary = $null; $start = $null
  foreach ($line in ($b -split "`n")) {
    if ($line -like "SUMMARY:*") { $summary = $line.Substring(8).Trim() }
    elseif ($line -like "DTSTART*") {
      $val = ($line -split ":",2)[1].Trim()
      $fmts = @("yyyyMMdd'T'HHmmss'Z'","yyyyMMdd'T'HHmm'Z'","yyyyMMdd'T'HHmmss","yyyyMMdd'T'HHmm","yyyyMMdd")
      foreach ($f in $fmts) {
        try {
          $dt = [datetime]::ParseExact($val,$f,$null)
          if ($val.EndsWith("Z")) { $dt = [datetime]::SpecifyKind($dt,[DateTimeKind]::Utc).ToLocalTime() }
          $start = $dt; break
        } catch { }
      }
    }
  }
  if ($start -ne $null -and $summary) {
    $events += [pscustomobject]@{ start=$start; summary=$summary }
  }
}

Write-Log "Parsed events: $($events.Count)"

# 3) Keep next 8
$next = $events | Where-Object { $_.start -ge (Get-Date).AddHours(-1) } |
        Sort-Object start | Select-Object -First 8

# 4) Output for Rainmeter
if ($next.Count -gt 0) {
  $lines = $next | ForEach-Object { "{0:ddd M/d h:mm tt} — {1}" -f $_.start, $_.summary }
} else {
  $lines = @("No upcoming events")
}

$lines -join "`n" | Set-Content -Path $outTxt -Encoding utf8
Write-Log "Wrote $($lines.Count) lines to gcal_events.txt"
Write-Log "----- Run end -----"