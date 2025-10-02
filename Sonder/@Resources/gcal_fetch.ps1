# gcal_fetch.ps1 (PowerShell 5.1 Compatible)
# Fetches Google Calendar ICS feed and formats up to 4 events for Rainmeter display

# ===================== CONFIG =====================
$IcsUrl = 'INSERT HERE'  # <-- Replace with your private Google Calendar ICS URL

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
    Write-Log "Unknown IANA timezone: $iana"
    return $null
}

function Parse-IcsDateTime {
    param(
        [string]$Raw,
        [string]$Tzid,
        [bool]$IsDateParam
    )
    
    $result = @{ Ok=$false; Start=$null; IsAllDay=$false; Reason='empty'; OriginalString=$Raw }
    
    if ([string]::IsNullOrWhiteSpace($Raw)) { 
        return [pscustomobject]$result 
    }
    
    $r = $Raw.Trim()
    
    # All-day event (YYYYMMDD or VALUE=DATE)
    if ($IsDateParam -or ($r -match '^\d{8}$')) {
        if ($r -match '^(\d{4})(\d{2})(\d{2})$') {
            $y = [int]$matches[1]
            $m = [int]$matches[2]
            $d = [int]$matches[3]
            try {
                # Create as local date at midnight
                $dt = New-Object System.DateTime @($y, $m, $d, 0, 0, 0, [System.DateTimeKind]::Local)
                return [pscustomobject]@{ 
                    Ok=$true
                    Start=$dt
                    IsAllDay=$true
                    Reason='all-day'
                    OriginalString=$Raw
                }
            } catch {
                Write-Log "Failed to create all-day date from $Raw : $_"
                $result['Reason'] = "all-day-error: $_"
                return [pscustomobject]$result
            }
        }
    }
    
    # UTC with Z suffix (20251003T163000Z)
    if ($r -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?Z$') {
        $y = [int]$matches[1]
        $m = [int]$matches[2] 
        $d = [int]$matches[3]
        $hh = [int]$matches[4]
        $mm = [int]$matches[5]
        $ss = if ($matches[6]) { [int]$matches[6] } else { 0 }
        
        try {
            # Create UTC DateTime
            $utc = New-Object System.DateTime @($y, $m, $d, $hh, $mm, $ss, [System.DateTimeKind]::Utc)
            
            # Convert to local time
            $localTz = [System.TimeZoneInfo]::Local
            $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($utc, $localTz)
            
            Write-Log "UTC conversion: $Raw -> UTC: $($utc.ToString('yyyy-MM-dd HH:mm')) -> Local: $($local.ToString('yyyy-MM-dd HH:mm'))"
            
            return [pscustomobject]@{ 
                Ok=$true
                Start=$local
                IsAllDay=$false
                Reason='UTC Z'
                OriginalString=$Raw
            }
        } catch {
            Write-Log "Failed to parse UTC time $Raw : $_"
            $result['Reason'] = "UTC-error: $_"
            return [pscustomobject]$result
        }
    }
    
    # Local/floating time or time with TZID (20251003T163000)
    if ($r -match '^(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2})?$') {
        $y = [int]$matches[1]
        $m = [int]$matches[2]
        $d = [int]$matches[3]
        $hh = [int]$matches[4]
        $mm = [int]$matches[5]
        $ss = if ($matches[6]) { [int]$matches[6] } else { 0 }
        
        try {
            if (-not [string]::IsNullOrWhiteSpace($Tzid)) {
                # Has TZID - convert from that timezone to local
                $winTz = Get-WindowsTimeZone $Tzid
                if ($winTz) {
                    try {
                        $srcTz = [System.TimeZoneInfo]::FindSystemTimeZoneById($winTz)
                        # Create as unspecified
                        $unspec = New-Object System.DateTime @($y, $m, $d, $hh, $mm, $ss, [System.DateTimeKind]::Unspecified)
                        # Convert from source timezone to local
                        $local = [System.TimeZoneInfo]::ConvertTimeToUtc($unspec, $srcTz)
                        $local = [System.TimeZoneInfo]::ConvertTimeFromUtc($local, [System.TimeZoneInfo]::Local)
                        
                        Write-Log "TZID conversion: $Raw TZID=$Tzid -> $($local.ToString('yyyy-MM-dd HH:mm')) local"
                        
                        return [pscustomobject]@{ 
                            Ok=$true
                            Start=$local
                            IsAllDay=$false
                            Reason="TZID:$Tzid"
                            OriginalString=$Raw
                        }
                    } catch {
                        Write-Log "Timezone conversion failed for $Tzid, treating as local: $_"
                        # Fall back to local time
                        $local = New-Object System.DateTime @($y, $m, $d, $hh, $mm, $ss, [System.DateTimeKind]::Local)
                        return [pscustomobject]@{ 
                            Ok=$true
                            Start=$local
                            IsAllDay=$false
                            Reason="TZID-fallback:$Tzid"
                            OriginalString=$Raw
                        }
                    }
                } else {
                    # Unknown TZID - treat as local time
                    Write-Log "Unknown TZID $Tzid, treating as local"
                    $local = New-Object System.DateTime @($y, $m, $d, $hh, $mm, $ss, [System.DateTimeKind]::Local)
                    return [pscustomobject]@{ 
                        Ok=$true
                        Start=$local
                        IsAllDay=$false
                        Reason="TZID-unknown:$Tzid"
                        OriginalString=$Raw
                    }
                }
            } else {
                # No TZID - floating local time
                $local = New-Object System.DateTime @($y, $m, $d, $hh, $mm, $ss, [System.DateTimeKind]::Local)
                return [pscustomobject]@{ 
                    Ok=$true
                    Start=$local
                    IsAllDay=$false
                    Reason='floating/local'
                    OriginalString=$Raw
                }
            }
        } catch {
            Write-Log "Failed to parse local/TZID time $Raw : $_"
            $result['Reason'] = "local-error: $_"
            return [pscustomobject]$result
        }
    }
    
    # Unhandled format
    Write-Log "Unhandled DTSTART format: $Raw (TZID=$Tzid)"
    $result['Reason'] = "unhandled-format"
    return [pscustomobject]$result
}

# =================== MAIN SCRIPT ===================
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

# Fetch ICS
$icsText = $null
try {
    $response = Invoke-WebRequest -Uri $IcsUrl -UseBasicParsing -ErrorAction Stop
    $icsText = $response.Content
    if ([string]::IsNullOrWhiteSpace($icsText)) { 
        throw "Empty ICS response" 
    }
    Write-Log "Fetched ICS successfully (length: $($icsText.Length) chars)"
} catch {
    Write-Log "Fetch failed: $($_.Exception.Message)"
    "No events in the next 7 days" | Set-Content -Path $OutFile -Encoding Unicode
    Write-Log "[DONE]"
    exit
}

# Unfold RFC 5545 continuation lines
$rawLines = $icsText -split "`r?`n"
$lines = New-Object System.Collections.ArrayList
foreach ($line in $rawLines) {
    if ($line -match '^[ \t]' -and $lines.Count -gt 0) {
        # Continuation line - append to previous
        $lines[$lines.Count - 1] += ($line -replace '^[ \t]+', '')
    } else {
        [void]$lines.Add($line)
    }
}

Write-Log "Unfolded $($rawLines.Count) raw lines to $($lines.Count) logical lines"

# Parse events
$events = New-Object System.Collections.ArrayList
$inEvent = $false
$current = @{}
$eventCount = 0
$parseErrors = 0

foreach ($line in $lines) {
    if ($line -eq 'BEGIN:VEVENT') {
        $inEvent = $true
        $current = @{}
        $eventCount++
        continue
    }
    
    if ($line -eq 'END:VEVENT') {
        $summary = Unescape-IcsText $current['SUMMARY']
        $location = Unescape-IcsText $current['LOCATION']
        $dtstart = $current['DTSTART']
        $dtparams = $current['DTSTART_PARAMS']
        
        if ([string]::IsNullOrWhiteSpace($summary)) {
            $summary = '(No title)'
        }
        
        $tzid = $null
        $isDateParam = $false
        
        if ($dtparams) {
            if ($dtparams -match 'TZID=([^;:]+)') { 
                $tzid = $matches[1] 
            }
            if ($dtparams -match 'VALUE=DATE') { 
                $isDateParam = $true 
            }
        }
        
        if (-not [string]::IsNullOrWhiteSpace($dtstart)) {
            $parsed = Parse-IcsDateTime -Raw $dtstart -Tzid $tzid -IsDateParam $isDateParam
            
            if ($parsed.Ok -and $parsed.Start) {
                $eventObj = [pscustomobject]@{
                    Start = $parsed.Start
                    IsAllDay = $parsed.IsAllDay
                    Summary = $summary
                    Location = $location
                    Reason = $parsed.Reason
                }
                [void]$events.Add($eventObj)
                
                $eventType = if ($parsed.IsAllDay) { 'all-day' } else { 'timed' }
                Write-Log "Parsed event: $($parsed.Start.ToString('yyyy-MM-dd HH:mm')) | $eventType | $($parsed.Reason) | $summary"
            } else {
                $parseErrors++
                Write-Log "Failed to parse DTSTART: raw=$dtstart tzid=$tzid reason=$($parsed.Reason) summary=$summary"
            }
        } else {
            $parseErrors++
            Write-Log "Event missing DTSTART: summary=$summary"
        }
        
        $inEvent = $false
        $current = @{}
        continue
    }
    
    if ($inEvent) {
        if ($line -match '^([A-Z][A-Z0-9-]*)(;[^:]*)?:(.*)$') {
            $propName = $matches[1]
            $propParams = $matches[2]
            $propValue = $matches[3]
            
            $current[$propName] = $propValue
            if ($propParams) {
                $current["${propName}_PARAMS"] = $propParams
            }
        }
    }
}

Write-Log "VEVENTs found: $eventCount | Successfully parsed: $($events.Count) | Parse errors: $parseErrors"

# Filter and sort events
$selected = New-Object System.Collections.ArrayList
$inCount = 0
$outCount = 0

# ---- BEGIN HYBRID SELECTION (PS 5.1) ---------------------------------

# Inputs assumed to exist:
#   $events (objects with .Start [DateTime], .IsAllDay [bool], .Summary [string])
#   $Now [DateTime], $DaysAhead [int], $MaxItems [int]
#   Write-Log function is available

$windowStart = $Now
$windowEnd   = $Now.AddDays($DaysAhead)

# Sort events ascending and keep only those with a valid Start
$sorted = @($events | Where-Object { $_ -ne $null -and $_.Start -ne $null } | Sort-Object Start)

Write-Log ("Window(min): {0} -> {1} (Local TZ: {2})" -f `
    $windowStart.ToString('yyyy-MM-dd HH:mm'), `
    $windowEnd.ToString('yyyy-MM-dd HH:mm'), `
    ([System.TimeZoneInfo]::Local).StandardName)

# First pass: classic rule — events within the next $DaysAhead days
$future   = @($sorted | Where-Object { $_.Start -ge $windowStart })
$within   = @($future | Where-Object { $_.Start -lt $windowEnd })
$selected = @()

foreach ($evt in $within) {
    if ($selected.Count -ge $MaxItems) { break }
    $kind = if ($evt.IsAllDay) { 'all-day' } else { 'timed' }
    Write-Log ("IN  -> {0} | {1} | {2}" -f $evt.Start.ToString('yyyy-MM-dd HH:mm'), $kind, $evt.Summary)
    $selected += ,$evt
}

# Hybrid extension: if we still have fewer than $MaxItems, take the next future events
$extended = $false
$effectiveEnd = $windowEnd
if ($selected.Count -lt $MaxItems) {
    $extended = $true
    $need = $MaxItems - $selected.Count
    $extras = @($future | Where-Object { $_.Start -ge $windowEnd } | Select-Object -First $need)
    foreach ($evt in $extras) {
        $kind = if ($evt.IsAllDay) { 'all-day' } else { 'timed' }
        Write-Log ("IN  (extended) -> {0} | {1} | {2}" -f $evt.Start.ToString('yyyy-MM-dd HH:mm'), $kind, $evt.Summary)
        $selected += ,$evt
    }
    if ($selected.Count -gt 0) {
        $effectiveEnd = $selected[$selected.Count-1].Start
    }
}

Write-Log ("Hybrid summary: future={0} within={1} selected={2} extended={3} effectiveEnd={4}" -f `
    $future.Count, $within.Count, $selected.Count, $extended, $effectiveEnd.ToString('yyyy-MM-dd HH:mm'))

# ---- FORMAT & WRITE EVENTS --------------------------------------------

# Build output text with up to 4 events
$outputLines = New-Object System.Collections.ArrayList

if ($selected.Count -eq 0) {
    [void]$outputLines.Add("No events in the next 7 days")
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
    # Write as UTF-16LE (required for bullet character)
    [System.IO.File]::WriteAllText($OutFile, $outputText, [System.Text.Encoding]::Unicode)
    Write-Log "Wrote $($outputLines.Count) line(s) to $OutFile"
} catch {
    Write-Log "Failed to write output file: $_"
    # Fallback write attempt
    $outputText | Set-Content -Path $OutFile -Encoding Unicode
}

Write-Log "[DONE]"
