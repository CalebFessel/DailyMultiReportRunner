<#
.SYNOPSIS
    Follow-up Traumasoft ThirdParty API probe. Standalone, read-only.

.DESCRIPTION
    The first probe answered whether the API can back the reports. This one
    answers the two questions that decide how to build them:

      A. Is /Schedule/Shifts a rolling window anchored to now, or can it be
         steered? The first probe returned 2026-08-17..2026-08-20 for a
         requested date of 2026-08-17 while "today" was the 18th, which looks
         like today-1..today+2 regardless of input. If that holds, staffing and
         the UHU denominator can only ever describe the present -- no backfill.

      B. Do trips backfill? OTP and the UHU numerator are worthless if
         GetTrips is windowed the same way. Pulls 7, 30 and 90 days back.

    It also measures whether the 33% of trip legs missing shift_name actually
    matter for on-time performance (they may be cancellations OTP already
    excludes), and whether vehicle_id offers a second route to cost center.

.EXAMPLE
    .\probe_traumasoft_followup.ps1 -Date 2026-08-17

.NOTES
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>
[CmdletBinding()]
param(
    [string]$BaseUrl,
    [string]$ApiKey,
    [string]$ApiSecret,
    [string]$Date = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd'),
    [string]$OutDir = 'api_probe'
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if (-not $BaseUrl)   { $BaseUrl   = $env:TS_API_BASE_URL }
if (-not $ApiKey)    { $ApiKey    = $env:TS_API_KEY }
if (-not $ApiSecret) { $ApiSecret = $env:TS_API_SECRET }
if (-not $BaseUrl)   { $BaseUrl   = Read-Host 'Traumasoft base URL' }
if (-not $ApiKey)    { $ApiKey    = Read-Host 'API key' }
if (-not $ApiSecret) { $ApiSecret = Read-Host 'API secret' }
if (-not $BaseUrl -or -not $ApiKey -or -not $ApiSecret) {
    Write-Host 'Base URL, key and secret are required.' -ForegroundColor Red; exit 2
}
$BaseUrl = $BaseUrl.TrimEnd('/')
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$script:LastRequestAt = [DateTime]::MinValue
$script:MinIntervalMs = 650

function New-TsHeaders {
    param([string]$Body = '')
    $timestamp = [string]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
    $nonce = [string](Get-Random -Minimum 100000000 -Maximum 999999999)
    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($ApiSecret)
    $bytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body + $timestamp + $nonce))
    $hmac.Dispose()
    return @{
        'X-TS-APIKEY'        = $ApiKey
        'X-TS-TIMESTAMP'     = $timestamp
        'X-TS-ID'            = $nonce
        'X-TS-AUTHORIZATION' = (-join ($bytes | ForEach-Object { $_.ToString('x2') }))
        'Accept'             = 'application/json'
    }
}

function Invoke-Ts {
    param([string]$Path, [hashtable]$Query, [int]$MaxRetries = 3)
    $qs = ''
    if ($Query -and $Query.Count) {
        $pairs = foreach ($k in $Query.Keys) {
            if ($null -eq $Query[$k]) { continue }
            '{0}={1}' -f [Uri]::EscapeDataString([string]$k), [Uri]::EscapeDataString([string]$Query[$k])
        }
        $qs = '?' + ($pairs -join '&')
    }
    $url = '{0}/api/{1}{2}' -f $BaseUrl, $Path.TrimStart('/'), $qs

    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        $since = ([DateTime]::UtcNow - $script:LastRequestAt).TotalMilliseconds
        if ($since -lt $script:MinIntervalMs) { Start-Sleep -Milliseconds ([int]($script:MinIntervalMs - $since)) }
        $script:LastRequestAt = [DateTime]::UtcNow
        try {
            $r = Invoke-WebRequest -Uri $url -Headers (New-TsHeaders) -Method GET -UseBasicParsing -TimeoutSec 120
            $data = $null
            if ($r.Content) { try { $data = $r.Content | ConvertFrom-Json } catch { } }
            return [pscustomobject]@{ Ok = $true; Status = [int]$r.StatusCode; Data = $data; Error = $null }
        }
        catch {
            $status = $null
            if ($_.Exception.Response) { try { $status = [int]$_.Exception.Response.StatusCode } catch { } }
            if ($status -eq 429 -and $attempt -lt $MaxRetries) { Start-Sleep -Seconds 60; continue }
            if ($status -ge 500 -and $attempt -lt $MaxRetries) { Start-Sleep -Seconds ([Math]::Pow(2, $attempt + 1)); continue }
            return [pscustomobject]@{ Ok = $false; Status = $status; Data = $null; Error = $_.Exception.Message }
        }
    }
    return [pscustomobject]@{ Ok = $false; Status = $null; Data = $null; Error = 'retries exhausted' }
}

function Get-TsRows {
    param($Payload)
    if ($null -eq $Payload) { return @() }
    if ($Payload -is [Array]) { return $Payload }
    if ($Payload -isnot [psobject]) { return @() }
    $names = @($Payload.PSObject.Properties.Name)
    foreach ($k in @('rows', 'users', 'custom_statuses', 'pay_types', 'employee_levels',
                     'fee_schedules', 'schedules', 'payor_categories', 'attachment_types')) {
        if ($names -contains $k -and $Payload.$k -is [Array]) { return $Payload.$k }
    }
    $arrays = @()
    foreach ($n in $names) { if ($Payload.$n -is [Array]) { $arrays += , $Payload.$n } }
    if ($arrays.Count -eq 1) { return $arrays[0] }
    return @()
}

function Get-ShiftWindow {
    param($Shifts)
    $starts = @($Shifts | Where-Object { $_.start_time } | ForEach-Object { [string]$_.start_time } | Sort-Object)
    $dates = @($starts | ForEach-Object { $_.Substring(0, [Math]::Min(10, $_.Length)) } | Sort-Object -Unique)
    $first = $null; $last = $null
    if ($starts.Count -gt 0) { $first = $starts[0]; $last = $starts[-1] }
    return [pscustomobject]@{
        Count = @($Shifts).Count; First = $first; Last = $last; Dates = $dates
    }
}

$findings = [ordered]@{ probe_date = $Date; base_url = $BaseUrl; run_at = (Get-Date).ToString('s') }
$md = New-Object System.Collections.Generic.List[string]
$md.Add('# Traumasoft ThirdParty API - follow-up probe')
$md.Add('')
$md.Add("Tenant: ``$BaseUrl``")
$md.Add("Reference date: **$Date**   Run at: **$((Get-Date).ToString('s'))**")
$md.Add('')

Write-Host ''
Write-Host 'Follow-up probe' -ForegroundColor Cyan
Write-Host ''

# === A. Is the Shifts window anchored to now? ================================
Write-Host '[1/5] Testing whether /Schedule/Shifts can be steered at all ...'
$windowTests = @()
$probeDates = @(
    (Get-Date).AddDays(-30).ToString('yyyy-MM-dd'),
    (Get-Date).AddDays(-1).ToString('yyyy-MM-dd'),
    (Get-Date).AddDays(30).ToString('yyyy-MM-dd')
)
$baseline = $null
foreach ($d in @('(none)') + $probeDates) {
    $q = @{ include_deleted = 'false' }
    if ($d -ne '(none)') { $q['start_date'] = $d; $q['date'] = $d }
    $r = Invoke-Ts -Path 'ThirdParty/Data/Schedule/Shifts' -Query $q
    if (-not $r.Ok) {
        Write-Host "    date=$d -> HTTP $($r.Status)"
        $windowTests += @{ requested = $d; error = $r.Status }
        continue
    }
    $w = Get-ShiftWindow (Get-TsRows $r.Data)
    if ($null -eq $baseline) { $baseline = $w }
    $identical = ($w.Count -eq $baseline.Count -and $w.First -eq $baseline.First -and $w.Last -eq $baseline.Last)
    Write-Host "    date=$d -> $($w.Count) rows, $($w.First) .. $($w.Last)"
    $windowTests += @{
        requested = $d; row_count = $w.Count; first = $w.First; last = $w.Last
        distinct_dates = $w.Dates.Count; identical_to_baseline = $identical
    }
}
# These are hashtables, so membership is ContainsKey, not PSObject.Properties.
$divergent = @($windowTests | Where-Object { $_.ContainsKey('identical_to_baseline') -and -not $_.identical_to_baseline })
$allIdentical = ($divergent.Count -eq 0)
$findings['shift_window'] = @{ tests = $windowTests; window_is_fixed = $allIdentical }

Write-Host "    -> window is $(if ($allIdentical) { 'IDENTICAL regardless of input (anchored to now)' } else { 'STEERABLE' })" -ForegroundColor $(if ($allIdentical) { 'Yellow' } else { 'Green' })

$md.Add('## A. Can /Schedule/Shifts be steered?')
$md.Add('')
$md.Add("**$(if ($allIdentical) { 'No - the window is anchored to now.' } else { 'Yes - some parameter changed the result.' })**")
$md.Add('')
$md.Add('| Requested date | Rows | First start | Last start | Same as baseline |')
$md.Add('|---|---|---|---|---|')
foreach ($t in $windowTests) {
    if ($t.ContainsKey('error')) {
        $md.Add("| $($t.requested) | HTTP $($t.error) | | | |")
    } else {
        $md.Add("| $($t.requested) | $($t.row_count) | ``$($t.first)`` | ``$($t.last)`` | $($t.identical_to_baseline) |")
    }
}
$md.Add('')

# === B. Do trips backfill? ===================================================
Write-Host '[2/5] Testing historical trip retrieval (7 / 30 / 90 days back) ...'
$backfill = @()
foreach ($daysBack in @(1, 7, 30, 90)) {
    $d = (Get-Date).AddDays(-$daysBack).ToString('yyyy-MM-dd')
    $r = Invoke-Ts -Path 'ThirdParty/Data/Cad/Trip' -Query @{
        rtype = 'GetTrips'; trip_date = $d; range_days = '1'
    }
    if (-not $r.Ok) {
        Write-Host "    $d ($daysBack d back) -> HTTP $($r.Status)"
        $backfill += @{ date = $d; days_back = $daysBack; error = $r.Status }
        continue
    }
    $legs = Get-TsRows $r.Data
    $withTs = @($legs | Where-Object { $_.timestamps }).Count
    Write-Host "    $d ($daysBack d back) -> $($legs.Count) legs, $withTs with timestamps"
    $backfill += @{ date = $d; days_back = $daysBack; leg_count = $legs.Count; with_timestamps = $withTs }
}
$oldest = @($backfill | Where-Object { $_.ContainsKey('leg_count') -and $_.leg_count -gt 0 } |
            Sort-Object { $_.days_back } -Descending | Select-Object -First 1)
$findings['trip_backfill'] = @{ tests = $backfill }

$md.Add('## B. Do trips backfill?')
$md.Add('')
$md.Add('| Date | Days back | Trip legs | With timestamps |')
$md.Add('|---|---|---|---|')
foreach ($b in $backfill) {
    if ($b.ContainsKey('error')) { $md.Add("| $($b.date) | $($b.days_back) | HTTP $($b.error) | |") }
    else { $md.Add("| $($b.date) | $($b.days_back) | $($b.leg_count) | $($b.with_timestamps) |") }
}
$md.Add('')
if ($oldest) {
    $md.Add("Oldest date returning data: **$($oldest.date)** ($($oldest.days_back) days back).")
    $md.Add('')
}

# === C. Do the shift_name-less legs matter for OTP? =========================
Write-Host "[3/5] Analysing trip legs without shift_name for $Date ..."
$r = Invoke-Ts -Path 'ThirdParty/Data/Cad/Trip' -Query @{
    rtype = 'GetTrips'; trip_date = $Date; range_days = '1'
}
$legs = @()
if ($r.Ok) { $legs = Get-TsRows $r.Data }

if ($legs.Count -eq 0) {
    Write-Host '    no legs returned; skipping' -ForegroundColor Yellow
    $findings['shift_name_gap'] = @{ ok = $false; reason = 'no legs returned' }
}
else {
    $byStatus = @{}
    foreach ($leg in $legs) {
        $status = 'unknown'
        if ($leg.trip_status) { $status = [string]$leg.trip_status }
        if (-not $byStatus.ContainsKey($status)) {
            $byStatus[$status] = [pscustomobject]@{
                Total = 0; WithShift = 0; WithVehicle = 0; WithPickup = 0; WithAtScene = 0
            }
        }
        $bucket = $byStatus[$status]
        $bucket.Total++
        if ($leg.shift_name) { $bucket.WithShift++ }
        if ($leg.vehicle_id) { $bucket.WithVehicle++ }
        if ($leg.pickup_time) { $bucket.WithPickup++ }
        # an at_scene stamp is what OTP scores against
        $hasAtScene = $false
        foreach ($entry in @($leg.timestamps)) {
            if ($entry -is [psobject]) {
                foreach ($n in $entry.PSObject.Properties.Name) {
                    if ($n -like 'at_scene*') { $hasAtScene = $true }
                }
            }
        }
        if ($hasAtScene) { $bucket.WithAtScene++ }
    }

    $statusRows = @()
    foreach ($status in ($byStatus.Keys | Sort-Object)) {
        $b = $byStatus[$status]
        $statusRows += @{
            trip_status = $status; total = $b.Total; with_shift_name = $b.WithShift
            with_vehicle_id = $b.WithVehicle; with_pickup_time = $b.WithPickup
            with_at_scene = $b.WithAtScene
        }
        Write-Host ("    {0,-28} total {1,5}  shift_name {2,5}  vehicle {3,5}  at_scene {4,5}" -f `
            $status, $b.Total, $b.WithShift, $b.WithVehicle, $b.WithAtScene)
    }

    # The OTP-relevant population is legs that actually ran: they have both a
    # scheduled pickup and an at_scene stamp to score against.
    $otpPopulation = @($legs | Where-Object {
        $_.pickup_time -and (@($_.timestamps) | Where-Object {
            $_ -is [psobject] -and (@($_.PSObject.Properties.Name) -like 'at_scene*').Count -gt 0
        }).Count -gt 0
    })
    $otpWithShift = @($otpPopulation | Where-Object { $_.shift_name }).Count
    $otpCoverage = 0.0
    if ($otpPopulation.Count -gt 0) { $otpCoverage = [Math]::Round($otpWithShift / $otpPopulation.Count, 3) }

    Write-Host ''
    Write-Host "    OTP-scorable legs (pickup_time + at_scene): $($otpPopulation.Count)"
    Write-Host "    of those, carrying shift_name: $otpWithShift ($([Math]::Round($otpCoverage * 100, 1))%)" -ForegroundColor $(if ($otpCoverage -ge 0.95) { 'Green' } else { 'Yellow' })

    $findings['shift_name_gap'] = @{
        ok = $true
        by_status = $statusRows
        otp_scorable_legs = $otpPopulation.Count
        otp_scorable_with_shift_name = $otpWithShift
        otp_cost_center_coverage = $otpCoverage
    }

    $md.Add('## C. Do the legs without shift_name matter for OTP?')
    $md.Add('')
    $md.Add('| Trip status | Legs | With shift_name | With vehicle_id | With at_scene |')
    $md.Add('|---|---|---|---|---|')
    foreach ($s in $statusRows) {
        $md.Add("| $($s.trip_status) | $($s.total) | $($s.with_shift_name) | $($s.with_vehicle_id) | $($s.with_at_scene) |")
    }
    $md.Add('')
    $md.Add("OTP-scorable legs (have both ``pickup_time`` and an ``at_scene`` stamp): **$($otpPopulation.Count)**")
    $md.Add("Of those, carrying ``shift_name``: **$otpWithShift** (**$([Math]::Round($otpCoverage * 100, 1))%**)")
    $md.Add('')
    $md.Add('This is the real cost-center coverage for the OTP report. The headline 66.7%')
    $md.Add('from the first probe counted every leg, including ones OTP never scores.')
    $md.Add('')
}

# === D. Vehicle status breakdown ============================================
Write-Host '[4/5] Counting vehicles by status ...'
$vehicles = @()
for ($page = 1; $page -le 20; $page++) {
    $r = Invoke-Ts -Path 'ThirdParty/Data/Fleet/Vehicles' -Query @{
        include_deleted = 'false'; include_disabled = 'false'; rows = '100'; page = $page
    }
    if (-not $r.Ok) { break }
    $rows = Get-TsRows $r.Data
    if ($rows.Count -eq 0) { break }
    $vehicles += $rows
    $total = $null
    if ($r.Data -is [psobject] -and $r.Data -isnot [Array]) {
        $names = @($r.Data.PSObject.Properties.Name)
        if ($names -contains 'total') { $total = [int]$r.Data.total }
        elseif ($names -contains 'total_pages') { $total = [int]$r.Data.total_pages }
    }
    if ($null -eq $total -or $page -ge $total) { break }
}
$statusCounts = @{}
foreach ($v in $vehicles) {
    $s = 'unknown'
    if ($v.vehicle_status) { $s = [string]$v.vehicle_status }
    if (-not $statusCounts.ContainsKey($s)) { $statusCounts[$s] = 0 }
    $statusCounts[$s]++
}
$statusList = @()
foreach ($s in ($statusCounts.Keys | Sort-Object)) {
    $statusList += @{ status = $s; count = $statusCounts[$s] }
    Write-Host ("    {0,-34} {1,4}" -f $s, $statusCounts[$s])
}
$onShift = @($vehicles | Where-Object { $_.shift_name }).Count
Write-Host "    currently on a shift (shift_name set): $onShift"
$findings['vehicle_status'] = @{ total = $vehicles.Count; by_status = $statusList; on_shift_now = $onShift }

$md.Add('## D. Vehicle status breakdown')
$md.Add('')
$md.Add('| Status | Vehicles |')
$md.Add('|---|---|')
foreach ($s in $statusList) { $md.Add("| $($s.status) | $($s.count) |") }
$md.Add("| **Total** | **$($vehicles.Count)** |")
$md.Add('')
$md.Add("Vehicles currently carrying ``shift_name``: **$onShift**")
$md.Add('')
$md.Add('The status strings already separate Retired / Waiting for Inspection /')
$md.Add('New - Waiting for Delivery, which is what the current SQL approximates with')
$md.Add('`status_reason NOT LIKE` patterns. That exclusion logic moves to status names.')
$md.Add('')

# === E. Ambiguous shift profiles, weighted by volume ========================
Write-Host '[5/5] Weighting ambiguous shift profiles by trip volume ...'
$shiftsResult = Invoke-Ts -Path 'ThirdParty/Data/Schedule/Shifts' -Query @{ include_deleted = 'false' }
$shifts = @()
if ($shiftsResult.Ok) { $shifts = Get-TsRows $shiftsResult.Data }

$employees = @()
for ($page = 1; $page -le 30; $page++) {
    $r = Invoke-Ts -Path 'ThirdParty/Data/User/Employees' -Query @{
        include_disabled = 'false'; rows = '100'; page = $page
    }
    if (-not $r.Ok) { break }
    $rows = Get-TsRows $r.Data
    if ($rows.Count -eq 0) { break }
    $employees += $rows
    $total = $null
    if ($r.Data -is [psobject] -and $r.Data -isnot [Array]) {
        $names = @($r.Data.PSObject.Properties.Name)
        if ($names -contains 'total') { $total = [int]$r.Data.total }
    }
    if ($null -eq $total -or $page -ge $total) { break }
}

if ($shifts.Count -eq 0 -or $employees.Count -eq 0 -or $legs.Count -eq 0) {
    Write-Host '    skipped - prerequisite pull empty' -ForegroundColor Yellow
    $findings['ambiguity'] = @{ ok = $false }
}
else {
    $empCc = @{}
    foreach ($e in $employees) {
        if ($e.user_id -and $e.cost_center_name) { $empCc[[string]$e.user_id] = [string]$e.cost_center_name }
    }
    $shiftCc = @{}
    foreach ($s in $shifts) {
        if (-not $s.shift_name -or -not $s.user_id) { continue }
        $cc = $empCc[[string]$s.user_id]
        if (-not $cc) { continue }
        $n = [string]$s.shift_name
        if (-not $shiftCc.ContainsKey($n)) { $shiftCc[$n] = New-Object System.Collections.Generic.HashSet[string] }
        [void]$shiftCc[$n].Add($cc)
    }
    $ambiguousNames = @($shiftCc.Keys | Where-Object { $shiftCc[$_].Count -gt 1 })
    $legsOnAmbiguous = @($legs | Where-Object { $_.shift_name -and ($ambiguousNames -contains [string]$_.shift_name) }).Count
    $legsOnMapped = @($legs | Where-Object { $_.shift_name -and $shiftCc.ContainsKey([string]$_.shift_name) }).Count

    Write-Host "    profiles mapped: $($shiftCc.Count), ambiguous: $($ambiguousNames.Count)"
    Write-Host "    legs landing on an ambiguous profile: $legsOnAmbiguous of $legsOnMapped mapped"

    $examples = @()
    foreach ($n in ($ambiguousNames | Select-Object -First 8)) {
        $examples += @{ shift_name = $n; cost_centers = @($shiftCc[$n]) }
    }
    $findings['ambiguity'] = @{
        ok = $true; profiles_mapped = $shiftCc.Count; ambiguous_profiles = $ambiguousNames.Count
        legs_on_ambiguous = $legsOnAmbiguous; legs_on_mapped = $legsOnMapped; examples = $examples
    }

    $md.Add('## E. Ambiguous shift profiles, weighted by volume')
    $md.Add('')
    $md.Add("- Profiles mapped to a cost center: **$($shiftCc.Count)**")
    $md.Add("- Profiles mapping to more than one: **$($ambiguousNames.Count)**")
    $md.Add("- Trip legs landing on an ambiguous profile: **$legsOnAmbiguous** of $legsOnMapped mapped")
    $md.Add('')
    if ($examples.Count -gt 0) {
        $md.Add('| Shift profile | Cost centers seen |')
        $md.Add('|---|---|')
        foreach ($e in $examples) { $md.Add("| $($e.shift_name) | $($e.cost_centers -join ', ') |") }
        $md.Add('')
    }
}

# --- write out ---------------------------------------------------------------
($findings | ConvertTo-Json -Depth 12) | Set-Content (Join-Path $OutDir 'followup_findings.json') -Encoding UTF8
$mdPath = Join-Path $OutDir 'FOLLOWUP_FINDINGS.md'
($md -join [Environment]::NewLine) | Set-Content -Path $mdPath -Encoding UTF8

Write-Host ''
Write-Host "Findings written to $mdPath" -ForegroundColor Green
Write-Host ''
