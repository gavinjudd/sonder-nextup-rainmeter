# gcal_fetch.ps1 (PowerShell 5.1 Compatible)
# Fetches Google Calendar ICS feed and formats up to 4 events for Rainmeter display

# ===================== CONFIG =====================
$IcsUrl = ''   # inline fallback only (used if secrets.ini is empty)
$urls   = @()  # populated after Write-Log() via secrets.ini

$DaysAhead       = 7
$MaxItems        = 4
$IncludeLocation = $false  # Set to $true to append " — Location" after event title

# IANA to Windows timezone mapping
$WindowsTzMap = @{
    'America/New_York'      = 'Eastern Standard Time'
    'America/Detroit'       = 'Eastern Standard Time'
    'America/Toronto'       = 'Eastern Standard Time'
    'America/Chicago'       = 'Central Standard Time'
    'America/Denver'        = 'Mountain Standard Time'
    'America/Phoenix'       = 'US Mountain Standard Time'
    'America/Los_Angeles'   = 'Pacific Standard Time'
    'America/Vancouver'     = 'Pacific Standard Time'
    'America/Anchorage'     = 'Alaskan Standard Time'
    'America/Honolulu'      = 'Hawaiian Standard Time'
    'Europe/London'         = 'GMT Standard Time'
    'Europe/Paris'          = 'Romance Standard Time'
    'Europe/Berlin'         = 'W. Europe Standard Time'
    'Asia/Tokyo'            = 'Tokyo Standard Time'
    'Australia/Sydney'      = 'AUS Eastern Standard Time'
    'UTC'                   = 'UTC'
    'Etc/UTC'              = 'UTC'
}
# =================== END CONFIG ===================

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutFile   = Join-Path $ScriptDir 'gcal_events.txt'
$LogFile   = Join-Path $ScriptDir 'gcal_log.txt'

function Write-Log {
    param([string]$msg)
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    Add-Content -Path $LogFile -Value "[$ts] $msg" -Encoding UTF8
}

function Get-SafeErrorMessage {
    param([System.Exception]$Exception)
    if ($null -eq $Exception) { return 'Unknown error' }
    # Google private ICS URLs are bearer secrets. Never write one to disk via logs.
    $message = ($Exception.Message + '') -replace '(?i)\b(?:https?|webcal)://\S+', '[calendar URL redacted]'
    $message = $message -replace '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b', '[email redacted]'
    $message = $message -replace '(?i)\b[A-Z]:\\[^\r\n]+', '[local path redacted]'
    return $message
}

function Write-CalendarOutput {
    param([string]$Text)

    $directory = Split-Path -Parent $OutFile
    $tempName = '.gcal_events.{0}.{1}.tmp' -f $PID, ([Guid]::NewGuid().ToString('N'))
    $tempPath = Join-Path $directory $tempName
    $backupName = '.gcal_events.{0}.{1}.bak' -f $PID, ([Guid]::NewGuid().ToString('N'))
    $backupPath = Join-Path $directory $backupName

    try {
        [System.IO.File]::WriteAllText($tempPath, $Text, [System.Text.Encoding]::Unicode)
        if (Test-Path -LiteralPath $OutFile) {
            # Same-directory replacement prevents WebParser from seeing a truncated file.
            [System.IO.File]::Replace($tempPath, $OutFile, $backupPath, $true)
        } else {
            [System.IO.File]::Move($tempPath, $OutFile)
        }
    } finally {
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $backupPath) {
            Remove-Item -LiteralPath $backupPath -Force -ErrorAction SilentlyContinue
        }
    }
}
function Unescape-IcsText {
    param([string]$t)
    if ([string]::IsNullOrWhiteSpace($t)) { return '' }
    $t = $t -replace '\\n', ' ' -replace '\\N', ' '
    $t = $t -replace '\\,', ',' -replace '\\;', ';' -replace '\\\\', '\'
    return $t.Trim()
}

function Get-WindowsTimeZone {
    param([string]$iana)
    if ([string]::IsNullOrWhiteSpace($iana)) { return $null }
    if ($WindowsTzMap.ContainsKey($iana)) { return $WindowsTzMap[$iana] }
    Write-Log 'Unknown IANA timezone; using local fallback'
    return $null
}

function Parse-IcsDateTime {
  param(
    [string]$RawValue,
    [hashtable]$Params
  )

  $result = @{
    Ok       = $false
    Start    = $null
    IsAllDay = $false
    Reason   = ''
  }

  if ([string]::IsNullOrWhiteSpace($RawValue)) {
    $result.Reason = 'empty'
    return [pscustomobject]$result
  }

  # VALUE=DATE (all-day) or bare yyyymmdd
  $isAllDay = $false
  if ($Params -and $Params.ContainsKey('VALUE')) {
    if (($Params['VALUE'] + '').ToUpper() -eq 'DATE') { $isAllDay = $true }
  }
  if ($isAllDay -or $RawValue -match '^\d{8}$') {
    try {
      $d = [datetime]::ParseExact($RawValue, 'yyyyMMdd', $null)
      # midnight local
      $result.Start    = Get-Date -Year $d.Year -Month $d.Month -Day $d.Day -Hour 0 -Minute 0 -Second 0
      $result.IsAllDay = $true
      $result.Ok       = $true
      $result.Reason   = 'all-day'
      return [pscustomobject]$result
    } catch {
      $result.Reason = 'allday-parse-failed'
      return [pscustomobject]$result
    }
  }

  # Helper: parse naive local date-time (no TZ, no Z)
  function _ParseLocal([string]$s) {
    $fmt = $null
    if     ($s.Length -eq 15) { $fmt = 'yyyyMMddTHHmmss' }
    elseif ($s.Length -eq 13) { $fmt = 'yyyyMMddTHHmm' }
    else { return $null }
    try { return [datetime]::ParseExact($s, $fmt, $null) } catch { return $null }
  }

  # Helper: parse UTC with trailing Z
  function _ParseUtc([string]$s) {
    $fmt = $null
    if     ($s.Length -eq 16) { $fmt = 'yyyyMMddTHHmmss' }
    elseif ($s.Length -eq 14) { $fmt = 'yyyyMMddTHHmm' }
    else { return $null }
    try {
      # Strip the literal Z before parsing, then set Kind explicitly. On Windows
      # PowerShell 5.1, parsing a format that includes Z can return Kind=Local;
      # passing that value to ConvertTimeFromUtc rejects every UTC event.
      $rawUtc = [datetime]::ParseExact($s.Substring(0, $s.Length - 1), $fmt, [Globalization.CultureInfo]::InvariantCulture)
      $utc = [datetime]::SpecifyKind($rawUtc, [DateTimeKind]::Utc)
      return [TimeZoneInfo]::ConvertTimeFromUtc($utc, [TimeZoneInfo]::Local)
    } catch { return $null }
  }

  # TZID handling
  $tzid = $null
  if ($Params -and $Params.ContainsKey('TZID')) { $tzid = $Params['TZID'] }

  try {
    if ($RawValue -match 'Z$') {
      # UTC → local
      $loc = _ParseUtc $RawValue
      if ($loc) {
        $result.Start  = $loc
        $result.Ok     = $true
        $result.Reason = 'utc-z'
        return [pscustomobject]$result
      } else {
        $result.Reason = 'utc-z-parse-failed'
        return [pscustomobject]$result
      }
    }

    if ($tzid) {
      $win = $tzid
      if ($WindowsTzMap.ContainsKey($tzid)) { $win = $WindowsTzMap[$tzid] }
      # Fallback to local if unknown TZID
      $srcTz = $null
      try { $srcTz = [TimeZoneInfo]::FindSystemTimeZoneById($win) } catch { $srcTz = $null }

      $naive = _ParseLocal $RawValue
      if ($naive) {
        if ($srcTz) {
          $utc   = [TimeZoneInfo]::ConvertTimeToUtc($naive, $srcTz)
          $local = [TimeZoneInfo]::ConvertTimeFromUtc($utc, [TimeZoneInfo]::Local)
          $result.Start  = $local
          $result.Ok     = $true
          $result.Reason = 'tzid'
          return [pscustomobject]$result
        } else {
          # Unknown TZID → treat as local
          $result.Start  = $naive
          $result.Ok     = $true
          $result.Reason = 'tzid-fallback-local'
          return [pscustomobject]$result
        }
      } else {
        $result.Reason = 'tzid-parse-failed'
        return [pscustomobject]$result
      }
    }

    # Floating (no TZ, no Z) → local
    $loc2 = _ParseLocal $RawValue
    if ($loc2) {
      $result.Start  = $loc2
      $result.Ok     = $true
      $result.Reason = 'floating-local'
      return [pscustomobject]$result
    } else {
      $result.Reason = 'floating-parse-failed'
      return [pscustomobject]$result
    }
  } catch {
    $result.Reason = 'exception'
    return [pscustomobject]$result
  }
}

# --- RRULE expansion (DAILY / WEEKLY / YEARLY with BYDAY / INTERVAL / UNTIL / COUNT / EXDATE) ---
function Expand-RRule {
  param(
    [datetime]$StartLocal,
    [string]  $RRule,
    [datetime]$WindowStart,
    [datetime]$WindowEnd,
    [string[]]$ExDates,
    [bool]    $IsAllDay = $false
  )

  $occ = New-Object System.Collections.ArrayList
  if ([string]::IsNullOrWhiteSpace($RRule)) { return $occ }

  # ---- Parse RRULE into a map ----
  $map = @{}
  foreach ($p in $RRule.Split(';')) {
    $kv = $p.Split('=',2)
    if ($kv.Length -eq 2) { $map[$kv[0].ToUpperInvariant()] = $kv[1] }
  }

  $freq = ""
  if ($map.ContainsKey('FREQ')) { $freq = ($map['FREQ'] + "").ToUpperInvariant() }

  # interval default = 1 (no ternary in PS5)
  $interval = 1
  if ($map.ContainsKey('INTERVAL') -and -not [string]::IsNullOrWhiteSpace($map['INTERVAL'])) {
    $interval = [int]$map['INTERVAL']
  }

  # BYDAY list (MO,WE,FR, ...)
  $byday = @()
  if ($map.ContainsKey('BYDAY')) {
    foreach ($d in $map['BYDAY'].Split(',')) { $byday += $d.ToUpper().Trim() }
  }

  # UNTIL (may be UTC Z or local)
  $until = $null
  if ($map.ContainsKey('UNTIL')) {
    $u = $map['UNTIL']
    try {
      if ($u -match 'Z$') {
        if ($u.Length -eq 16) { $fmt = 'yyyyMMddTHHmmss' } else { $fmt = 'yyyyMMddTHHmm' }
        $rawUtc = [datetime]::ParseExact($u.Substring(0, $u.Length - 1), $fmt, [Globalization.CultureInfo]::InvariantCulture)
        $utc = [datetime]::SpecifyKind($rawUtc, [DateTimeKind]::Utc)
        $until = [TimeZoneInfo]::ConvertTimeFromUtc($utc, [TimeZoneInfo]::Local)
      } else {
        if     ($u.Length -eq 15) { $fmt = 'yyyyMMddTHHmmss' }
        elseif ($u.Length -eq 8)  { $fmt = 'yyyyMMdd' }
        else                      { $fmt = 'yyyyMMddTHHmm' }
        $until = [datetime]::ParseExact($u, $fmt, $null)
      }
    } catch { $until = $null }
  }

  # COUNT
  $count = $null
  if ($map.ContainsKey('COUNT')) { $count = [int]$map['COUNT'] }

  # Build EXDATE set (local times)
  $ex = @{}
  foreach ($exraw in ($ExDates | Where-Object { $_ })) {
    foreach ($part in ($exraw -split ',')) {
      $t = $part.Trim()
      if (-not $t) { continue }
      try {
        if ($t -match 'Z$') {
          if ($t.Length -eq 16) { $fmt = 'yyyyMMddTHHmmss' } else { $fmt = 'yyyyMMddTHHmm' }
          $rawUtc = [datetime]::ParseExact($t.Substring(0, $t.Length - 1), $fmt, [Globalization.CultureInfo]::InvariantCulture)
          $utc = [datetime]::SpecifyKind($rawUtc, [DateTimeKind]::Utc)
          $ld  = [TimeZoneInfo]::ConvertTimeFromUtc($utc, [TimeZoneInfo]::Local)
        } else {
          if     ($t.Length -eq 15) { $fmt = 'yyyyMMddTHHmmss' }
          elseif ($t.Length -eq 8)  { $fmt = 'yyyyMMdd' }
          else                       { $fmt = 'yyyyMMddTHHmm' }
          $ld  = [datetime]::ParseExact($t, $fmt, $null)
        }
        if ($IsAllDay) { $h = 0; $m = 0 } else { $h = $StartLocal.Hour; $m = $StartLocal.Minute }
        $key = (Get-Date -Year $ld.Year -Month $ld.Month -Day $ld.Day -Hour $h -Minute $m -Second 0).ToString('o')
        $ex[$key] = $true
      } catch { }
    }
  }

  function Add-If-In-Window([datetime]$dt) {
    if ($dt -ge $WindowStart -and $dt -lt $WindowEnd) {
      $k = $dt.ToString('o')
      if (-not $ex.ContainsKey($k)) { [void]$occ.Add($dt) }
    }
  }

  switch ($freq) {
    'DAILY' {
      $dt = $StartLocal; $added = 0
      while ($dt -lt $WindowEnd) {
        if ($until -and $dt -gt $until) { break }
        if ($dt -ge $WindowStart) { Add-If-In-Window $dt }
        $dt = $dt.AddDays($interval)
        if ($count) { $added++; if ($added -ge $count) { break } }
      }
    }
    'WEEKLY' {
      $mapDow = @{ 'SU'=0; 'MO'=1; 'TU'=2; 'WE'=3; 'TH'=4; 'FR'=5; 'SA'=6 }
      if ($byday.Count -gt 0) {
        $days = @()
        foreach ($b in $byday) { if ($mapDow.ContainsKey($b)) { $days += $mapDow[$b] } }
      } else {
        $days = @($StartLocal.DayOfWeek.value__)
      }

      # Sunday-based week start
      $weekStart = $StartLocal.Date.AddDays(-1 * [int]$StartLocal.DayOfWeek)
      $weeks = 0; $added = 0
      while ($true) {
        $base = $weekStart.AddDays($weeks * 7 * $interval)
        if ($until -and $base -gt $until) { break }
        foreach ($d in $days) {
          $dt = $base.AddDays($d).AddHours($StartLocal.Hour).AddMinutes($StartLocal.Minute)
          if ($until -and $dt -gt $until) { continue }
          if ($dt -ge $WindowEnd) { break }
          if ($dt -lt $StartLocal) { continue }
          Add-If-In-Window $dt
          if ($count) { $added++; if ($added -ge $count) { break } }
        }
        if ($count -and $added -ge $count) { break }
        if ($base -ge $WindowEnd) { break }
        $weeks++
      }
    }
    'MONTHLY' {
      $monthCursor = Get-Date -Year $StartLocal.Year -Month $StartLocal.Month -Day 1 -Hour 0 -Minute 0 -Second 0
      $generated = 0
      $done = $false
      $mapDow = @{ 'SU'=0; 'MO'=1; 'TU'=2; 'WE'=3; 'TH'=4; 'FR'=5; 'SA'=6 }

      while (-not $done -and $monthCursor -lt $WindowEnd) {
        $candidates = New-Object System.Collections.ArrayList
        $daysInMonth = [DateTime]::DaysInMonth($monthCursor.Year, $monthCursor.Month)

        if ($map.ContainsKey('BYMONTHDAY')) {
          foreach ($rawDay in $map['BYMONTHDAY'].Split(',')) {
            $monthDay = 0
            if (-not [int]::TryParse($rawDay.Trim(), [ref]$monthDay) -or $monthDay -eq 0) { continue }
            if ($monthDay -lt 0) { $monthDay = $daysInMonth + $monthDay + 1 }
            if ($monthDay -lt 1 -or $monthDay -gt $daysInMonth) { continue }
            $candidate = Get-Date -Year $monthCursor.Year -Month $monthCursor.Month -Day $monthDay -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second
            [void]$candidates.Add($candidate)
          }
        } elseif ($byday.Count -gt 0) {
          foreach ($token in $byday) {
            if ($token -notmatch '^(?<ord>[+-]?\d+)?(?<dow>SU|MO|TU|WE|TH|FR|SA)$') { continue }
            $targetDow = $mapDow[$matches['dow']]
            $ordinalText = $matches['ord']

            if ($ordinalText) {
              $ordinal = [int]$ordinalText
              if ($ordinal -gt 0) {
                $first = Get-Date -Year $monthCursor.Year -Month $monthCursor.Month -Day 1 -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second
                $offset = ($targetDow - [int]$first.DayOfWeek + 7) % 7
                $monthDay = 1 + $offset + (($ordinal - 1) * 7)
              } else {
                $last = Get-Date -Year $monthCursor.Year -Month $monthCursor.Month -Day $daysInMonth -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second
                $offset = ([int]$last.DayOfWeek - $targetDow + 7) % 7
                $monthDay = $daysInMonth - $offset + (($ordinal + 1) * 7)
              }
              if ($monthDay -ge 1 -and $monthDay -le $daysInMonth) {
                [void]$candidates.Add((Get-Date -Year $monthCursor.Year -Month $monthCursor.Month -Day $monthDay -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second))
              }
            } else {
              for ($monthDay = 1; $monthDay -le $daysInMonth; $monthDay++) {
                $candidate = Get-Date -Year $monthCursor.Year -Month $monthCursor.Month -Day $monthDay -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second
                if ([int]$candidate.DayOfWeek -eq $targetDow) { [void]$candidates.Add($candidate) }
              }
            }
          }
        } elseif ($StartLocal.Day -le $daysInMonth) {
          [void]$candidates.Add((Get-Date -Year $monthCursor.Year -Month $monthCursor.Month -Day $StartLocal.Day -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second))
        }

        foreach ($dt in @($candidates | Sort-Object -Unique)) {
          if ($dt -lt $StartLocal) { continue }
          if ($until -and $dt -gt $until) { $done = $true; break }
          $generated++
          Add-If-In-Window $dt
          if ($count -and $generated -ge $count) { $done = $true; break }
        }

        $monthCursor = $monthCursor.AddMonths($interval)
      }
    }
    'YEARLY' {
      $year = $StartLocal.Year
      $generated = 0
      while ($year -le $WindowEnd.Year) {
        $dt = $null
        try {
          $dt = Get-Date -Year $year -Month $StartLocal.Month -Day $StartLocal.Day -Hour $StartLocal.Hour -Minute $StartLocal.Minute -Second $StartLocal.Second
        } catch {
          # Invalid dates such as Feb 29 in a non-leap year are skipped.
        }

        if ($dt) {
          if ($until -and $dt -gt $until) { break }
          if ($dt -ge $StartLocal) {
            $generated++
            Add-If-In-Window $dt
            if ($count -and $generated -ge $count) { break }
          }
        }

        $year += $interval
      }
    }
    default {
      Add-If-In-Window $StartLocal
    }
  }

  return $occ
}

# =================== MAIN SCRIPT ===================

# --- Collect settings once (secrets.ini preferred; fall back to $IcsUrl) ---
$urls = @()
$HideSummaryDay  = @()
$HideSummaryFrom = @()
try {
  $SecretIni = Join-Path $ScriptDir 'secrets.ini'
  if (Test-Path $SecretIni) {
    foreach ($raw in (Get-Content $SecretIni -Encoding UTF8)) {
      $t = $raw.Trim()
      if ($t -like '#*' -or [string]::IsNullOrWhiteSpace($t)) { continue }
      $i = $t.IndexOf('=')
      if ($i -lt 1) { continue }
      $k = $t.Substring(0,$i).Trim().ToUpperInvariant()
      $v = $t.Substring($i+1).Trim()
      switch ($k) {
        'ICS_URLS' { $urls += ($v -split '[,;]' | ForEach-Object { ($_ -replace '\s+#.*$','').Trim() } | Where-Object { $_ }) }
        'ICS_URL'  { $urls += $v.Trim() }
        'ICS_URL2' { $urls += $v.Trim() }
        default {
          if     ($k -like 'HIDE_SUMMARY_DAY*')  { $HideSummaryDay  += $v }
          elseif ($k -like 'HIDE_SUMMARY_FROM*') { $HideSummaryFrom += $v }
        }
      }
    }
  }
} catch {
  Write-Log ("Failed to read secrets.ini: {0}" -f (Get-SafeErrorMessage $_.Exception))
}

# Build optional local hide lists:
# - HIDE_SUMMARY_DAY=Title|YYYY-MM-DD   (hide only that day)
# - HIDE_SUMMARY_FROM=Title|YYYY-MM-DD  (hide from that date forward)
if (-not $urls -or $urls.Count -eq 0) {
  if (-not [string]::IsNullOrWhiteSpace($IcsUrl) -and $IcsUrl -notmatch 'REPLACE|INSERT|YOUR') {
    $urls = @($IcsUrl)
  }
}
$urls = @($urls | Select-Object -Unique)

# Normalize to a fast lookup set (day-level)
$HideSet = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($h in $HideSummaryDay) {
  $norm = ($h + '').Trim().ToLower()
  if ($norm) { [void]$HideSet.Add($norm) }
}

# Normalize "hide from" rules to: title(lower) -> boundary UTC
$HideFromBoundary = @{}
foreach ($h in $HideSummaryFrom) {
  $parts = ($h + '').Split('|',2)
  if ($parts.Length -eq 2) {
    $t = $parts[0].Trim().ToLower()
    $d = $parts[1].Trim()
    try {
      $bdLocal = [datetime]::ParseExact($d,'yyyy-MM-dd',$null)
      $HideFromBoundary[$t] = $bdLocal.ToUniversalTime()
    } catch { }
  }
}

if (-not $urls -or $urls.Count -eq 0) {
  Write-Log "No ICS URLs found (secrets.ini or inline)."
  if (-not (Test-Path -LiteralPath $OutFile)) {
    Write-CalendarOutput ''
  }
  Write-Log "[DONE] Exit code: 2"
  exit 2
}

"[START] gcal_fetch.ps1" | Set-Content -Path $LogFile -Encoding UTF8

# Enable TLS 1.2
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
} catch {
    Write-Log "Warning: Could not set TLS 1.2"
}

$now = Get-Date
$windowEnd = $now.AddDays($DaysAhead)

Write-Log "Window: $($now.ToString('yyyy-MM-dd HH:mm')) -> $($windowEnd.ToString('yyyy-MM-dd HH:mm')) (Local TZ: $([System.TimeZoneInfo]::Local.Id))"

# Fetch & unfold ALL ICS sources (merge into one $lines list)
$lines = New-Object System.Collections.ArrayList
$srcIdx = 0
$successfulSources = 0
$failedSources = 0
$hadError = $false
$requestHeaders = @{
  'Cache-Control' = 'no-cache, no-store, max-age=0'
  'Pragma' = 'no-cache'
}

foreach ($u in $urls) {
  $srcIdx++
  $fetched = $false

  for ($attempt = 1; $attempt -le 2; $attempt++) {
    try {
        # ICS URLs are bearer secrets, so never copy them into the log.
        Write-Log ("Fetching calendar source {0} of {1} (attempt {2})" -f $srcIdx, $urls.Count, $attempt)
        $response = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 20 -Headers $requestHeaders -ErrorAction Stop
        $icsText  = $response.Content
        if ([string]::IsNullOrWhiteSpace($icsText)) { throw "Empty ICS response" }
        if ($icsText -notmatch '(?m)^BEGIN:VCALENDAR\s*$') { throw "Response is not an iCalendar feed" }
        Write-Log "Fetched ICS successfully (length: $($icsText.Length) chars)"

        # Unfold RFC 5545 continuation lines for this source
        $rawLines   = $icsText -split "`r?`n"
        $localLines = New-Object System.Collections.ArrayList
        foreach ($line in $rawLines) {
            if ($line -match '^[ \t]' -and $localLines.Count -gt 0) {
                $localLines[$localLines.Count - 1] += ($line -replace '^[ \t]+', '')
            } else {
                [void]$localLines.Add($line)
            }
        }
        Write-Log "Unfolded $($rawLines.Count) raw lines to $($localLines.Count) logical lines"

        [void]$lines.Add("X-SOURCE-INDEX:$srcIdx")
        foreach ($L in $localLines) { [void]$lines.Add($L) }
        $successfulSources++
        $fetched = $true
        break
    } catch {
        $safeError = Get-SafeErrorMessage $_.Exception
        if ($attempt -lt 2) {
          Write-Log ("Calendar source {0} attempt {1} failed; retrying: {2}" -f $srcIdx, $attempt, $safeError)
          Start-Sleep -Seconds 2
        } else {
          Write-Log ("Calendar source {0} failed after {1} attempts: {2}" -f $srcIdx, $attempt, $safeError)
        }
    }
  }

  if (-not $fetched) {
    $failedSources++
    $hadError = $true
  }
}

if ($failedSources -gt 0) {
  Write-Log ("{0} of {1} calendar sources failed; preserving the last-known-good display." -f $failedSources, $urls.Count)
  if (-not (Test-Path -LiteralPath $OutFile)) {
    Write-CalendarOutput ''
  }
  Write-Log "[DONE] Exit code: 1"
  exit 1
}

# Parse events
$events = New-Object System.Collections.ArrayList
$inEvent = $false
$current = @{}
$eventCount = 0
$parseErrors = 0

# ---- Parse VEVENTs (with EXDATE + RECURRENCE-ID cancel/override) ----
$totalEvents = 0
$parseErrors = 0
$events = New-Object System.Collections.ArrayList
$inEvent = $false
$current = @{}

# Per-instance edits/cancels keyed by UID|origStartUtcIso
$cancels   = @{}
$overrides = @{}
$cancelsBySummaryDay = @{}   # tolerate VALUE=DATE/TZ quirks by summary+day
$cancelAfter = @{}           # UID -> boundary UTC (RECURRENCE-ID;RANGE=THISANDFUTURE)
$cancelAfterBySummary = @{}  # summary(lower) -> boundary UTC (handles UID changes)

$currentSource = 1  # preserved from the merge stage; leave as-is if you already set this earlier

foreach ($line in $lines) {
    # keep source markers if you added them earlier
    if ($line -match '^X-SOURCE-INDEX:(\d+)$') { $currentSource = [int]$matches[1]; continue }

    if ($line -eq 'BEGIN:VEVENT') {
        $inEvent = $true
        $current = @{}
        $totalEvents++
        continue
    }

    if ($line -eq 'END:VEVENT') {
        if ($inEvent) {
            # Params maps for DTSTART / RECURRENCE-ID
            $dtParams  = @{}
            if ($current.ContainsKey('DTSTART_PARAMS')) { $dtParams = $current['DTSTART_PARAMS'] }
            $ridParams = @{}
            if ($current.ContainsKey('RECURRENCE-ID_PARAMS')) { $ridParams = $current['RECURRENCE-ID_PARAMS'] }

            $uid     = if ($current.ContainsKey('UID')) { $current['UID'] } else { $null }
            $summary = if ($current.ContainsKey('SUMMARY'))  { Unescape-IcsText ($current['SUMMARY'] + '') } else { '(no title)' }
            $loc     = if ($current.ContainsKey('LOCATION')) { Unescape-IcsText ($current['LOCATION'] + '') } else { $null }
            $rrule   = if ($current.ContainsKey('RRULE'))    { $current['RRULE']    } else { $null }
            $status  = if ($current.ContainsKey('STATUS'))   { ($current['STATUS'] + '').ToUpperInvariant() } else { $null }
            $exArr   = @()
            if ($current.ContainsKey('EXDATE')) { $exArr = $current['EXDATE'].Split(',') }

            # Master DTSTART for either single event or base of recurrence
            $parsed = Parse-IcsDateTime -RawValue $current['DTSTART'] -Params $dtParams

            # A per-instance edit or cancel will have RECURRENCE-ID
            $recIdRaw = if ($current.ContainsKey('RECURRENCE-ID')) { $current['RECURRENCE-ID'] } else { $null }

            # Cancelled ordinary VEVENTs must never flow into the display list.
            if (-not $recIdRaw -and $status -eq 'CANCELLED') {
                Write-Log "Skip CANCELLED single event"
                $inEvent = $false
                $current = @{}
                continue
            }

            if ($recIdRaw) {
    # This VEVENT modifies or cancels exactly one instance in a series
    $rid = Parse-IcsDateTime -RawValue $recIdRaw -Params $ridParams
    $origKeys = @()
    $ridLocalForDay = $null

    if ($uid) {
        if ($rid.Ok) {
            $ridLocalForDay = $rid.Start
            # Strict key (preferred)
            $origKeys += ($uid + '|' + $rid.Start.ToUniversalTime().ToString('o'))
        }

        # Fallback key: derive from raw yyyymmddThhmm[ss] if strict parse fails
        if (-not $rid.Ok) {
            if ($recIdRaw -match '^(?<d>\d{8})(T(?<t>\d{4}|\d{6}))?') {
                $d = $matches['d']; $t = $matches['t']
                $yyyy = [int]$d.Substring(0,4); $MM = [int]$d.Substring(4,2); $dd = [int]$d.Substring(6,2)
                if ($t) {
                    $HH = [int]$t.Substring(0,2); $mm = [int]$t.Substring(2,2)
                    $ss = if ($t.Length -ge 6) { [int]$t.Substring(4,2) } else { 0 }
                } else {
                    $HH = 0; $mm = 0; $ss = 0
                }
                $ridLocalForDay = Get-Date -Year $yyyy -Month $MM -Day $dd -Hour $HH -Minute $mm -Second $ss
                $origKeys += ($uid + '|' + $ridLocalForDay.ToUniversalTime().ToString('o'))
            }
        }
    }

    if ($uid -and $ridLocalForDay) {
        # Day-level cancel key to tolerate VALUE=DATE or TZ mismatches
        $dayKey = $uid + '|day|' + $ridLocalForDay.Date.ToString('yyyy-MM-dd')
    } else {
        $dayKey = $null
    }

    # Read RANGE param (e.g., THISANDFUTURE)
$range = $null
if ($ridParams -and $ridParams.ContainsKey('RANGE')) { $range = ($ridParams['RANGE'] + '').ToUpper() }

if ($origKeys.Count -gt 0 -or $dayKey) {
    if ($status -eq 'CANCELLED') {
        foreach ($k in $origKeys) { $cancels[$k] = $true }
        if ($dayKey) { $cancels[$dayKey] = $true }
        if ($summary -and $ridLocalForDay) {
            $sdKey = ($summary.Trim().ToLower() + '|' + $ridLocalForDay.Date.ToString('yyyy-MM-dd'))
            $cancelsBySummaryDay[$sdKey] = $true
        }
        # Cascade cancel — everything from RECURRENCE-ID forward
        if ($range -eq 'THISANDFUTURE' -and $ridLocalForDay) {
            $boundaryUtc = $ridLocalForDay.ToUniversalTime()
            if ($uid) {
                $cancelAfter[$uid] = $boundaryUtc
                Write-Log ("Cascade cancel (THISANDFUTURE): boundary={0}" -f $boundaryUtc.ToString('o'))
            }
            if ($summary) {
                $skey = $summary.Trim().ToLower()
                if (-not $cancelAfterBySummary.ContainsKey($skey) -or $boundaryUtc -lt $cancelAfterBySummary[$skey]) {
                    $cancelAfterBySummary[$skey] = $boundaryUtc
                    Write-Log ("Cascade cancel (summary fallback): boundary={0}" -f $boundaryUtc.ToString('o'))
                }
            }
        }

        } else {
            # Override to new DTSTART/SUMMARY/LOCATION
            if ($parsed.Ok -and $parsed.Start) {
    $title = $summary
    if ($IncludeLocation -and -not [string]::IsNullOrWhiteSpace($loc)) { $title += " — " + $loc.Trim() }
                foreach ($k in $origKeys) {
                    $overrides[$k] = @{
                        Start         = $parsed.Start
                        OriginalStart = $ridLocalForDay
                        IsAllDay      = $parsed.IsAllDay
                        Summary       = $title
                        SeriesSummary = $summary
                        UID           = $uid
                        Source        = $currentSource
                    }
                }
                # We intentionally do NOT make a day-level override to avoid ambiguity
            }
        }
    } else {
        $parseErrors++
    }

    # Done with this VEVENT; continue to next
    $inEvent = $false
    $current = @{}
    continue
}

            # Normal (non-override) event or recurring master
            if ($parsed.Ok -and $parsed.Start) {
                $title = $summary
                if ($IncludeLocation -and -not [string]::IsNullOrWhiteSpace($loc)) { $title += " — " + $loc.Trim() }

                # Expand recurrence if present; otherwise single DTSTART
                $occurs = @()
                if (-not [string]::IsNullOrWhiteSpace($rrule)) {
                    $recurrenceWindowStart = if ($parsed.IsAllDay) { $Now.Date } else { $Now }
                    $occurs = Expand-RRule -StartLocal $parsed.Start -RRule $rrule -WindowStart $recurrenceWindowStart -WindowEnd $WindowEnd -ExDates $exArr -IsAllDay $parsed.IsAllDay
                } else {
                    $occurs = @($parsed.Start)
                }

                foreach ($o in $occurs) {

    # Summary-level cascade boundary (handles UID changes after "This and future" deletes)
    if ($summary) {
        $skey = $summary.Trim().ToLower()
        if ($cancelAfterBySummary.ContainsKey($skey)) {
            if ($o.ToUniversalTime() -ge $cancelAfterBySummary[$skey]) { continue }
        }
    }

    # Local hide list: "Summary|YYYY-MM-DD" (case-insensitive)
    $sumDayKeyCheck = $summary
    if ($sumDayKeyCheck) {
        $sumDayKeyCheck = $sumDayKeyCheck.Trim().ToLower() + '|' + $o.Date.ToString('yyyy-MM-dd')
        if ($HideSet.Contains($sumDayKeyCheck)) { continue }
    }

$outStart = $o
$outTitle = $title
$outAll   = $parsed.IsAllDay

# Local "hide from" rule (suppresses all future instances after the boundary)
if ($summary) {
    $skey = $summary.Trim().ToLower()
    if ($HideFromBoundary.ContainsKey($skey)) {
        if ($o.ToUniversalTime() -ge $HideFromBoundary[$skey]) { continue }
    }
}

                    # Apply RECURRENCE-ID cancel/override if a UID exists
                    if ($uid) {
    # Skip everything at/after cascade boundary (RECURRENCE-ID;RANGE=THISANDFUTURE)
    if ($cancelAfter.ContainsKey($uid)) {
        if ($o.ToUniversalTime() -ge $cancelAfter[$uid]) { continue }
    }

    $key    = $uid + '|' + $o.ToUniversalTime().ToString('o')
    $dayKey = $uid + '|day|' + $o.Date.ToString('yyyy-MM-dd')

    # Summary+day tolerant cancel
    $sumDayKey = $summary
    if ($sumDayKey) { $sumDayKey = $sumDayKey.Trim().ToLower() + '|' + $o.Date.ToString('yyyy-MM-dd') }

    if ( $cancels.ContainsKey($key) -or $cancels.ContainsKey($dayKey) -or ( $sumDayKey -and $cancelsBySummaryDay.ContainsKey($sumDayKey) ) ) { continue }

    if ($overrides.ContainsKey($key)) {
        $ov       = $overrides[$key]
        $outStart = $ov['Start']
        $outAll   = [bool]$ov['IsAllDay']
        $outTitle = $ov['Summary']
    }
}

                    $null = $events.Add([pscustomobject]@{
                        Start         = $outStart
                        OriginalStart = $o
                        IsAllDay      = $outAll
                        Summary       = $outTitle
                        SeriesSummary = $summary
                        UID           = $uid
                        SourceIndex   = $currentSource
                    })
                }
            } else {
                $parseErrors++
            }
        }
        $inEvent = $false
        $current = @{}
        continue
    }

    if (-not $inEvent) { continue }

    # Parse "NAME;P=V;...:VALUE" (robust)
    if ($line -match '^([^:]+):(.*)$') {
        $left      = $matches[1]
        $propValue = $matches[2]

        # Split left into name + (optional) params
        $segments = $left.Split(';')
        $propName = $segments[0].ToUpper()
        $paramStr = $null
        if ($segments.Count -gt 1) {
            $paramStr = ($segments[1..($segments.Count-1)] -join ';')
        }

        # Accumulate EXDATE (don’t overwrite)
        if ($propName -eq 'EXDATE') {
            if ($current.ContainsKey('EXDATE')) { $current['EXDATE'] = $current['EXDATE'] + ',' + $propValue } else { $current['EXDATE'] = $propValue }
        } else {
            $current[$propName] = $propValue
        }

        # Params into a map (e.g., TZID, VALUE=DATE)
        if ($paramStr) {
            $pmap = @{}
            foreach ($p in $paramStr.Split(';')) {
                $kv = $p.Split('=',2)
                if ($kv.Length -eq 2) { $pmap[$kv[0].ToUpper()] = $kv[1] }
            }
            if ($pmap.Count -gt 0) { $current["${propName}_PARAMS"] = $pmap }
        }
    }
}

# Reconcile recurrence exceptions after every VEVENT has been read. Google does
# not guarantee that RECURRENCE-ID records precede their recurring master.
$reconciledEvents = New-Object System.Collections.ArrayList
$matchedOverrideKeys = @{}

foreach ($event in @($events)) {
    $originalStart = if ($event.PSObject.Properties['OriginalStart']) { $event.OriginalStart } else { $event.Start }
    $eventUid = if ($event.PSObject.Properties['UID']) { $event.UID } else { $null }
    $seriesSummary = if ($event.PSObject.Properties['SeriesSummary']) { $event.SeriesSummary } else { $event.Summary }

    if ($seriesSummary) {
        $summaryKey = $seriesSummary.Trim().ToLowerInvariant()
        if ($cancelAfterBySummary.ContainsKey($summaryKey) -and $originalStart.ToUniversalTime() -ge $cancelAfterBySummary[$summaryKey]) { continue }
    }

    $eventStart = $event.Start
    $eventAllDay = [bool]$event.IsAllDay
    $eventSummary = $event.Summary
    $eventSource = $event.SourceIndex

    if ($eventUid) {
        if ($cancelAfter.ContainsKey($eventUid) -and $originalStart.ToUniversalTime() -ge $cancelAfter[$eventUid]) { continue }

        $strictKey = $eventUid + '|' + $originalStart.ToUniversalTime().ToString('o')
        $dayKey = $eventUid + '|day|' + $originalStart.Date.ToString('yyyy-MM-dd')
        $summaryDayKey = if ($seriesSummary) { $seriesSummary.Trim().ToLowerInvariant() + '|' + $originalStart.Date.ToString('yyyy-MM-dd') } else { $null }

        if ($cancels.ContainsKey($strictKey) -or $cancels.ContainsKey($dayKey) -or ($summaryDayKey -and $cancelsBySummaryDay.ContainsKey($summaryDayKey))) { continue }

        if ($overrides.ContainsKey($strictKey)) {
            $override = $overrides[$strictKey]
            $matchedOverrideKeys[$strictKey] = $true
            $eventStart = $override['Start']
            $eventAllDay = [bool]$override['IsAllDay']
            $eventSummary = $override['Summary']
            $eventSource = $override['Source']
        }
    }

    [void]$reconciledEvents.Add([pscustomobject]@{
        Start         = $eventStart
        OriginalStart = $originalStart
        IsAllDay      = $eventAllDay
        Summary       = $eventSummary
        SeriesSummary = $seriesSummary
        UID           = $eventUid
        SourceIndex   = $eventSource
    })
}

# Preserve a moved exception even when its recurring master was not included in
# the feed's exported range. Final window filtering below decides visibility.
foreach ($overrideKey in @($overrides.Keys)) {
    if ($matchedOverrideKeys.ContainsKey($overrideKey)) { continue }
    $override = $overrides[$overrideKey]
    $originalStart = $override['OriginalStart']
    $seriesSummary = $override['SeriesSummary']
    $eventUid = $override['UID']

    if (-not $originalStart -or -not $override['Start']) { continue }
    if ($cancelAfter.ContainsKey($eventUid) -and $originalStart.ToUniversalTime() -ge $cancelAfter[$eventUid]) { continue }

    $summaryKey = if ($seriesSummary) { $seriesSummary.Trim().ToLowerInvariant() } else { $null }
    if ($summaryKey -and $cancelAfterBySummary.ContainsKey($summaryKey) -and $originalStart.ToUniversalTime() -ge $cancelAfterBySummary[$summaryKey]) { continue }
    if ($summaryKey -and $HideFromBoundary.ContainsKey($summaryKey) -and $originalStart.ToUniversalTime() -ge $HideFromBoundary[$summaryKey]) { continue }

    $summaryDayKey = if ($summaryKey) { $summaryKey + '|' + $originalStart.Date.ToString('yyyy-MM-dd') } else { $null }
    if ($summaryDayKey -and ($HideSet.Contains($summaryDayKey) -or $cancelsBySummaryDay.ContainsKey($summaryDayKey))) { continue }

    [void]$reconciledEvents.Add([pscustomobject]@{
        Start         = $override['Start']
        OriginalStart = $originalStart
        IsAllDay      = [bool]$override['IsAllDay']
        Summary       = $override['Summary']
        SeriesSummary = $seriesSummary
        UID           = $eventUid
        SourceIndex   = $override['Source']
    })
}

$events = $reconciledEvents
Write-Log ("VEVENTs found: seen={0} ok={1} errors={2}" -f $totalEvents, $events.Count, $parseErrors)

# Filter and sort events
$selected = New-Object System.Collections.ArrayList
$inCount = 0
$outCount = 0

# ---- STRICT 7-DAY SELECTION + DEDUPE ---------------------------------

$windowStart = Get-Date
$windowEnd   = $windowStart.AddDays($DaysAhead)

# Sort & keep valid
$sorted = @($events | Where-Object { $_ -ne $null -and $_.Start -ne $null } | Sort-Object Start)

# De-dupe across feeds by (StartUTC + Summary)
$seen    = @{}
$deduped = New-Object System.Collections.ArrayList
foreach ($e in $sorted) {
  $k = ($e.Start.ToUniversalTime().ToString('o') + '|' + $e.Summary.Trim()).ToLower()
  if (-not $seen.ContainsKey($k)) { $seen[$k] = $true; [void]$deduped.Add($e) }
}

# Timed items are upcoming; all-day items remain visible for their whole local day.
$within = @($deduped | Where-Object {
  if ($_.IsAllDay) {
    $_.Start.Date -ge $windowStart.Date -and $_.Start.Date -lt $windowEnd.Date
  } else {
    $_.Start -ge $windowStart -and $_.Start -lt $windowEnd
  }
})

# "UP NEXT" is a single chronological queue across every configured calendar.
# Source order must never displace an event that starts sooner.
$withinSorted = @($within | Sort-Object Start, SourceIndex, Summary)

$selected = @($withinSorted | Select-Object -First $MaxItems)

$selPrimary   = @($selected | Where-Object { $_.SourceIndex -eq 1 }).Count
$selSecondary = $selected.Count - $selPrimary
$withinBySource = @($within | Group-Object SourceIndex | Sort-Object Name | ForEach-Object { "{0}:{1}" -f $_.Name, $_.Count }) -join ','
$selectedBySource = @($selected | Group-Object SourceIndex | Sort-Object Name | ForEach-Object { "{0}:{1}" -f $_.Name, $_.Count }) -join ','
Write-Log ("Window summary (chronological): within={0} selected={1} [primary={2} secondary={3}] [within-by-source={4}] [selected-by-source={5}]" -f $within.Count, $selected.Count, $selPrimary, $selSecondary, $withinBySource, $selectedBySource)

# ---- FORMAT & WRITE EVENTS --------------------------------------------

# Build output text with up to 4 events
$outputLines = New-Object System.Collections.ArrayList

if ($selected.Count -eq 0) {
    Write-Log "No events to display"
} else {
    # Format each selected event (up to MaxItems)
    $displayCount = [Math]::Min($selected.Count, $MaxItems)

    for ($i = 0; $i -lt $displayCount; $i++) {
        $evt = $selected[$i]

        if ($evt.IsAllDay) {
            # All-day format: ddd M/d • All day • SUMMARY
            $line = $evt.Start.ToString('ddd M/d') + ' • All day • ' + $evt.Summary
        } else {
            # Timed format: ddd M/d • h:mm tt • SUMMARY
            $line = $evt.Start.ToString('ddd M/d • h:mm tt') + ' • ' + $evt.Summary
        }

        [void]$outputLines.Add($line)
    }

    Write-Log "Formatted $($outputLines.Count) event(s) for display"
}

# Join lines with CRLF and write as UTF-16LE (Unicode)
$outputText = $outputLines -join "`r`n"

try {
    Write-CalendarOutput $outputText
    Write-Log "Atomically wrote $($outputLines.Count) line(s)"
} catch {
    $hadError = $true
    Write-Log ("Failed to write output file: {0}" -f (Get-SafeErrorMessage $_.Exception))
}

$exit = if ($hadError) { 1 } else { 0 }
Write-Log ("[DONE] Exit code: {0}" -f $exit)
exit $exit
