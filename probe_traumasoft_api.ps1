<#
.SYNOPSIS
    Standalone Traumasoft ThirdParty API probe. No Python, no git, no installs.

.DESCRIPTION
    Does the same job as probe_traumasoft_api.py but in pure PowerShell, so it
    runs on any Windows machine with nothing checked out and nothing installed.
    Everything it needs (HMAC-SHA256, HTTPS, JSON) is built into .NET.

    Read-only: it issues GET requests only.

    It answers the questions the OpenAPI spec leaves open:
      1. Do the credentials work?
      2. Does /Schedule/Shifts support any date filtering? The spec documents
         only include_deleted, which the staffing report and the UHU
         denominator both depend on.
      3. Which trip timestamp names does this tenant emit, now that ePCR is
         out of scope for this API?
      4. Is the vehicle field allowlist really closed?
      5. Can a trip be resolved to a cost center, and how often?

.EXAMPLE
    .\probe_traumasoft_api.ps1
    (prompts for anything it needs)

.EXAMPLE
    .\probe_traumasoft_api.ps1 -BaseUrl https://lynxems.traumasoft.com `
        -ApiKey xxxx -ApiSecret yyyy -Date 2026-08-17

.NOTES
    If PowerShell refuses to run the file:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

    The API's HMAC timestamp is valid for 300 seconds, so a machine clock more
    than ~5 minutes off will produce 401s that look like bad credentials.
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

# Older PowerShell defaults to TLS 1.0, which Traumasoft will refuse.
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---------------------------------------------------------------- credentials
if (-not $BaseUrl)   { $BaseUrl   = $env:TS_API_BASE_URL }
if (-not $ApiKey)    { $ApiKey    = $env:TS_API_KEY }
if (-not $ApiSecret) { $ApiSecret = $env:TS_API_SECRET }

if (-not $BaseUrl) { $BaseUrl = Read-Host 'Traumasoft base URL (e.g. https://lynxems.traumasoft.com)' }
if (-not $ApiKey)  { $ApiKey  = Read-Host 'API key (X-TS-APIKEY)' }
if (-not $ApiSecret) {
    Write-Host 'API secret (press Enter if the key screen only issued a key) ' -NoNewline
    $ApiSecret = Read-Host
}

if (-not $BaseUrl -or -not $ApiKey) {
    Write-Host 'Base URL and API key are required.' -ForegroundColor Red
    exit 2
}

# The spec's formulas take a secret paired with the key, but the key-creation
# screen may issue only one value. With no secret, the key doubles as the
# secret and the scheme is settled by probing below.
if (-not $ApiSecret) { $ApiSecret = $ApiKey }
$BaseUrl = $BaseUrl.TrimEnd('/')

# Signing scheme in force. Resolved by Resolve-TsAuthMode before any real call.
$script:AuthMode   = 'default'
$script:AuthSecret = $ApiSecret

if ($Date -notmatch '^\d{4}-\d{2}-\d{2}$') {
    Write-Host "Date must be YYYY-MM-DD, got '$Date'." -ForegroundColor Red
    exit 2
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

# ---------------------------------------------------------------- HTTP + HMAC
$script:LastRequestAt = [DateTime]::MinValue
$script:MinIntervalMs = 650   # stay under the documented 100 reads/minute

function New-TsHeaders {
    <#
      Builds the four auth headers under one signing scheme.

        default: hmac_sha256(body + timestamp + nonce, secret)
        legacy:  hmac_sha256(api_key, timestamp + secret + nonce)
                 -- documented for GPS Geofence / the older Postman collection

      Body is an empty string on GET.
    #>
    param(
        [string]$Body = '',
        [string]$Mode,
        [string]$Secret
    )

    if (-not $Mode)   { $Mode   = $script:AuthMode }
    if (-not $Secret) { $Secret = $script:AuthSecret }

    # Not Get-Date -UFormat %s: that is culture- and version-sensitive and has
    # bitten people with local-vs-UTC epochs. This is unambiguous.
    $timestamp = [string]([DateTimeOffset]::UtcNow.ToUnixTimeSeconds())
    $nonce = [string](Get-Random -Minimum 100000000 -Maximum 999999999)

    if ($Mode -eq 'legacy') {
        $hmacKey = $timestamp + $Secret + $nonce
        $message = $ApiKey
    }
    else {
        $hmacKey = $Secret
        $message = $Body + $timestamp + $nonce
    }

    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = [Text.Encoding]::UTF8.GetBytes($hmacKey)
    $hashBytes = $hmac.ComputeHash([Text.Encoding]::UTF8.GetBytes($message))
    $hmac.Dispose()
    $signature = -join ($hashBytes | ForEach-Object { $_.ToString('x2') })

    return @{
        'X-TS-APIKEY'        = $ApiKey
        'X-TS-TIMESTAMP'     = $timestamp
        'X-TS-ID'            = $nonce
        'X-TS-AUTHORIZATION' = $signature
        'Accept'             = 'application/json'
    }
}

function Resolve-TsAuthMode {
    <#
      Establishes which (formula, secret) pair this tenant accepts by trying
      each against a cheap read endpoint. Sets $script:AuthMode and
      $script:AuthSecret on success; returns $null when all are rejected.
    #>
    $candidates = @()
    foreach ($mode in @('default', 'legacy')) {
        foreach ($pair in @(
            @{ Label = 'secret'; Value = $ApiSecret },
            @{ Label = 'api_key'; Value = $ApiKey }
        )) {
            $already = $candidates | Where-Object { $_.Mode -eq $mode -and $_.Secret -eq $pair.Value }
            if ($already) { continue }
            $candidates += [pscustomobject]@{
                Mode = $mode; Secret = $pair.Value; Label = $pair.Label
            }
        }
    }

    $url = '{0}/api/ThirdParty/Data/Organization' -f $BaseUrl
    foreach ($c in $candidates) {
        try {
            $headers = New-TsHeaders -Mode $c.Mode -Secret $c.Secret
            $null = Invoke-WebRequest -Uri $url -Headers $headers -Method GET `
                -UseBasicParsing -TimeoutSec 60
            Write-Host "    accepted: formula=$($c.Mode), secret=$($c.Label)" -ForegroundColor Green
            $script:AuthMode = $c.Mode
            $script:AuthSecret = $c.Secret
            return $c
        }
        catch {
            $status = ''
            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { }
            }
            Write-Host "    formula=$($c.Mode), secret=$($c.Label) -> HTTP $status"
        }
        Start-Sleep -Milliseconds $script:MinIntervalMs
    }
    return $null
}

function ConvertTo-TsQueryString {
    param([hashtable]$Query)
    if (-not $Query -or $Query.Count -eq 0) { return '' }
    $pairs = foreach ($key in $Query.Keys) {
        $value = $Query[$key]
        if ($null -eq $value) { continue }
        '{0}={1}' -f [Uri]::EscapeDataString([string]$key), [Uri]::EscapeDataString([string]$value)
    }
    return '?' + ($pairs -join '&')
}

function Invoke-Ts {
    <#
      One signed GET. Returns a result object rather than throwing, so probes
      can record a 401/422 as a finding instead of aborting the run.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$Query,
        [int]$MaxRetries = 3
    )

    $url = '{0}/api/{1}{2}' -f $BaseUrl, $Path.TrimStart('/'), (ConvertTo-TsQueryString $Query)

    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        # Throttle so a long run never trips the rate limit in the first place.
        $sinceLast = ([DateTime]::UtcNow - $script:LastRequestAt).TotalMilliseconds
        if ($sinceLast -lt $script:MinIntervalMs) {
            Start-Sleep -Milliseconds ([int]($script:MinIntervalMs - $sinceLast))
        }
        $script:LastRequestAt = [DateTime]::UtcNow

        try {
            $response = Invoke-WebRequest -Uri $url -Headers (New-TsHeaders) `
                -Method GET -UseBasicParsing -TimeoutSec 120
            $data = $null
            if ($response.Content) {
                try { $data = $response.Content | ConvertFrom-Json } catch { $data = $null }
            }
            return [pscustomobject]@{
                Ok     = $true
                Status = [int]$response.StatusCode
                Data   = $data
                Error  = $null
            }
        }
        catch {
            $status = $null
            $body = $_.Exception.Message
            $retryAfter = 60

            if ($_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { }
                try {
                    $hdr = $_.Exception.Response.Headers['Retry-After']
                    if ($hdr) { $retryAfter = [int]$hdr }
                } catch { }
                try {
                    $stream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object IO.StreamReader($stream)
                    $body = $reader.ReadToEnd()
                    $reader.Dispose()
                } catch { }
            }

            if ($status -eq 429 -and $attempt -lt $MaxRetries) {
                $wait = [Math]::Min($retryAfter, 120)
                Write-Host "    rate limited; waiting $wait s" -ForegroundColor DarkYellow
                Start-Sleep -Seconds $wait
                continue
            }
            if ($status -ge 500 -and $attempt -lt $MaxRetries) {
                Start-Sleep -Seconds ([Math]::Pow(2, $attempt + 1))
                continue
            }

            return [pscustomobject]@{
                Ok     = $false
                Status = $status
                Data   = $null
                Error  = ($body -replace '\s+', ' ').Trim()
            }
        }
    }

    return [pscustomobject]@{ Ok = $false; Status = $null; Data = $null; Error = 'retries exhausted' }
}

function Get-TsRows {
    <#
      The API uses several envelopes: bare arrays (Shifts, Organization,
      GetTrips), paginated objects keyed on 'rows' or 'users', and small
      reference lists keyed on their own plural name.
    #>
    param($Payload)

    if ($null -eq $Payload) { return @() }
    if ($Payload -is [Array]) { return $Payload }
    if ($Payload -isnot [psobject]) { return @() }

    $names = @($Payload.PSObject.Properties.Name)
    foreach ($key in @('rows', 'users', 'custom_statuses', 'pay_types', 'employee_levels',
                       'fee_schedules', 'schedules', 'payor_categories', 'attachment_types')) {
        if ($names -contains $key -and $Payload.$key -is [Array]) { return $Payload.$key }
    }

    $arrays = @()
    foreach ($n in $names) { if ($Payload.$n -is [Array]) { $arrays += , $Payload.$n } }
    if ($arrays.Count -eq 1) { return $arrays[0] }
    return @()
}

function Invoke-TsPaged {
    <#
      Walks both paginator dialects: current_page/total_pages, and jqGrid's
      page/total (Vehicles, Employees, Payors, Facilities).
    #>
    param([string]$Path, [hashtable]$Query, [int]$PageSize = 100, [int]$MaxPages = 200)

    $all = @()
    $q = @{}
    if ($Query) { foreach ($k in $Query.Keys) { $q[$k] = $Query[$k] } }
    $q['rows'] = $PageSize

    for ($page = 1; $page -le $MaxPages; $page++) {
        $q['page'] = $page
        $result = Invoke-Ts -Path $Path -Query $q
        if (-not $result.Ok) { return [pscustomobject]@{ Ok = $false; Rows = $all; Error = $result.Error; Status = $result.Status } }

        $rows = Get-TsRows $result.Data
        if ($rows.Count -gt 0) { $all += $rows }

        $totalPages = $null
        if ($result.Data -is [psobject] -and $result.Data -isnot [Array]) {
            $names = @($result.Data.PSObject.Properties.Name)
            if ($names -contains 'total_pages') { $totalPages = $result.Data.total_pages }
            elseif ($names -contains 'total')    { $totalPages = $result.Data.total }
        }
        if ($null -eq $totalPages) { break }          # unpaginated endpoint
        if ($page -ge [int]$totalPages) { break }
        if ($rows.Count -eq 0) { break }
    }

    return [pscustomobject]@{ Ok = $true; Rows = $all; Error = $null; Status = 200 }
}

function Get-KeyUnion {
    param($Rows)
    $keys = New-Object System.Collections.Generic.HashSet[string]
    foreach ($row in $Rows) {
        if ($row -is [psobject]) {
            foreach ($n in $row.PSObject.Properties.Name) { [void]$keys.Add($n) }
        }
    }
    return @($keys) | Sort-Object
}

function Save-Sample {
    param($Rows, [string]$Name, [int]$Count = 5)
    $sample = @($Rows) | Select-Object -First $Count
    $path = Join-Path $OutDir $Name
    ($sample | ConvertTo-Json -Depth 10) | Set-Content -Path $path -Encoding UTF8
}

# ---------------------------------------------------------------------- probe
$findings = [ordered]@{
    probe_date = $Date
    base_url   = $BaseUrl
}

Write-Host ''
Write-Host "Traumasoft ThirdParty API probe -> $BaseUrl (date $Date)" -ForegroundColor Cyan
Write-Host ''

# --- 1. credentials -----------------------------------------------------------
Write-Host '[1/6] Detecting the HMAC scheme against /Data/Organization ...'
$scheme = Resolve-TsAuthMode
if (-not $scheme) {
    Write-Host ''
    Write-Host '    Every signing scheme was rejected.' -ForegroundColor Red
    Write-Host '    Check, in order:' -ForegroundColor Yellow
    Write-Host '      - whether the key screen offers a secret you have not captured'
    Write-Host '        (some builds show it once, at creation time only)'
    Write-Host '      - this machine''s clock: the timestamp is valid for only 300 seconds'
    Write-Host '        (w32tm /query /status)'
    Write-Host '      - that the key has read scopes enabled for Organization'
    $findings['auth'] = @{ ok = $false; error = 'no HMAC scheme accepted' }
    ($findings | ConvertTo-Json -Depth 10) | Set-Content (Join-Path $OutDir 'findings.json') -Encoding UTF8
    exit 1
}

$orgResult = Invoke-Ts -Path 'ThirdParty/Data/Organization'
if (-not $orgResult.Ok) {
    Write-Host "    FAILED (HTTP $($orgResult.Status)): $($orgResult.Error)" -ForegroundColor Red
    $findings['auth'] = @{
        ok = $false; status = $orgResult.Status; error = $orgResult.Error; auth_mode = $scheme.Mode
    }
    ($findings | ConvertTo-Json -Depth 10) | Set-Content (Join-Path $OutDir 'findings.json') -Encoding UTF8
    exit 1
}
$org = Get-TsRows $orgResult.Data
Write-Host "    OK - $($org.Count) organization rows" -ForegroundColor Green
$findings['auth'] = @{
    ok                = $true
    auth_mode         = $scheme.Mode
    secret_source     = $scheme.Label
    organization_rows = $org.Count
}
Save-Sample $org 'organization_sample.json'

# --- 2. shifts ----------------------------------------------------------------
Write-Host '[2/6] Pulling /Data/Schedule/Shifts (no documented date filter) ...'
$shifts = @()
$shiftsResult = Invoke-Ts -Path 'ThirdParty/Data/Schedule/Shifts' -Query @{ include_deleted = 'false' }
if (-not $shiftsResult.Ok) {
    Write-Host "    FAILED (HTTP $($shiftsResult.Status)): $($shiftsResult.Error)" -ForegroundColor Red
    $findings['shifts'] = @{ ok = $false; status = $shiftsResult.Status; error = $shiftsResult.Error }
}
else {
    $shifts = Get-TsRows $shiftsResult.Data
    $shiftKeys = Get-KeyUnion $shifts

    # Fields the current staffing SQL filters on that the documented schema omits.
    $wanted = @('published', 'timeoff_type', 'schedule_type', 'cost_center_name',
                'cost_center_id', 'group', 'group_id', 'unit_id', 'schedule_id')
    $missing = @($wanted | Where-Object { $shiftKeys -notcontains $_ })

    $starts = @($shifts | Where-Object { $_.start_time } |
                ForEach-Object { [string]$_.start_time } | Sort-Object)
    $dates = @($starts | ForEach-Object { $_.Substring(0, [Math]::Min(10, $_.Length)) } |
               Sort-Object -Unique)

    Write-Host "    $($shifts.Count) rows spanning $($dates.Count) distinct dates"
    if ($starts.Count -gt 0) {
        Write-Host "    span: $($starts[0]) .. $($starts[-1])"
    }
    if ($missing.Count -gt 0) {
        Write-Host "    Shift rows lack: $($missing -join ', ')" -ForegroundColor Yellow
    }

    # Undocumented filters: one "works" if it changes the row count.
    Write-Host '    Testing undocumented filter shapes ...'
    $candidates = @(
        @{ start_date = $Date },
        @{ date = $Date },
        @{ shift_date = $Date },
        @{ from = $Date; to = $Date },
        @{ start_time = $Date },
        @{ begin_date = $Date; end_date = $Date },
        @{ page = '1'; rows = '10' }
    )
    $filterResults = @()
    foreach ($candidate in $candidates) {
        $q = @{ include_deleted = 'false' }
        foreach ($k in $candidate.Keys) { $q[$k] = $candidate[$k] }
        $r = Invoke-Ts -Path 'ThirdParty/Data/Schedule/Shifts' -Query $q
        $label = ($candidate.Keys | ForEach-Object { "$_=$($candidate[$_])" }) -join '&'
        if ($r.Ok) {
            $n = (Get-TsRows $r.Data).Count
            $changed = ($n -ne $shifts.Count)
            $filterResults += @{ params = $label; row_count = $n; changed_result = $changed }
            $flag = ''
            if ($changed) { $flag = '   <-- FILTER APPLIED' }
            Write-Host "      $label -> $n rows$flag"
        }
        else {
            $filterResults += @{ params = $label; status = $r.Status; error = $r.Error }
            Write-Host "      $label -> HTTP $($r.Status)"
        }
    }

    $earliest = $null
    $latest = $null
    if ($starts.Count -gt 0) { $earliest = $starts[0]; $latest = $starts[-1] }
    $filtersThatWorked = @($filterResults | Where-Object { $_.changed_result })

    $findings['shifts'] = @{
        ok                                   = $true
        row_count                            = $shifts.Count
        observed_keys                        = $shiftKeys
        distinct_start_dates                 = $dates.Count
        earliest_start_time                  = $earliest
        latest_start_time                    = $latest
        covers_requested_date                = ($dates -contains $Date)
        missing_fields_needed_by_current_sql = $missing
        undocumented_filters                 = $filterResults
        server_side_date_filtering           = ($filtersThatWorked.Count -gt 0)
    }
    Save-Sample $shifts 'shifts_sample.json'
}

# --- 3. trips -----------------------------------------------------------------
Write-Host "[3/6] Pulling /Data/Cad/Trip?rtype=GetTrips for $Date ..."
$trips = @()
$tripResult = Invoke-Ts -Path 'ThirdParty/Data/Cad/Trip' -Query @{
    rtype = 'GetTrips'; trip_date = $Date; range_days = '1'
}
if (-not $tripResult.Ok) {
    Write-Host "    FAILED (HTTP $($tripResult.Status)): $($tripResult.Error)" -ForegroundColor Red
    $findings['trips'] = @{ ok = $false; status = $tripResult.Status; error = $tripResult.Error }
}
else {
    $trips = Get-TsRows $tripResult.Data
    $tripKeys = Get-KeyUnion $trips

    # timestamps is an array of "status name -> ISO time" maps. Which names this
    # tenant emits decides what replaces the ePCR arrival time in the OTP report.
    $tsNames = New-Object System.Collections.Generic.HashSet[string]
    $withTs = 0
    foreach ($trip in $trips) {
        if ($trip.timestamps) {
            $withTs++
            foreach ($entry in @($trip.timestamps)) {
                if ($entry -is [psobject]) {
                    foreach ($n in $entry.PSObject.Properties.Name) { [void]$tsNames.Add($n) }
                }
            }
        }
    }
    $tsNameList = @($tsNames) | Sort-Object
    $costFields = @($tripKeys | Where-Object { $_ -match 'cost' })

    Write-Host "    $($trips.Count) trip legs; $withTs carry timestamps"
    Write-Host "    timestamp names: $(if ($tsNameList.Count) { $tsNameList -join ', ' } else { 'none' })"
    if ($costFields.Count -eq 0) {
        Write-Host '    no cost-center field on trip legs (expected) - resolved indirectly below' -ForegroundColor Yellow
    }

    $findings['trips'] = @{
        ok                        = $true
        row_count                 = $trips.Count
        observed_keys             = $tripKeys
        trips_with_timestamps     = $withTs
        observed_timestamp_names  = $tsNameList
        with_pickup_time          = @($trips | Where-Object { $_.pickup_time }).Count
        with_shift_name           = @($trips | Where-Object { $_.shift_name }).Count
        with_late_reasons         = @($trips | Where-Object { $_.late_reasons }).Count
        distinct_call_types       = @($trips | Where-Object { $_.call_type } |
                                      ForEach-Object { [string]$_.call_type } | Sort-Object -Unique)
        any_cost_center_field     = $costFields
    }
    Save-Sample $trips 'trips_sample.json'
}

# --- 4. vehicles --------------------------------------------------------------
Write-Host '[4/6] Pulling /Data/Fleet/Vehicles and confirming the field allowlist ...'
$vehiclesPaged = Invoke-TsPaged -Path 'ThirdParty/Data/Fleet/Vehicles' -Query @{
    include_deleted = 'false'; include_disabled = 'false'
}
if (-not $vehiclesPaged.Ok) {
    Write-Host "    FAILED (HTTP $($vehiclesPaged.Status)): $($vehiclesPaged.Error)" -ForegroundColor Red
    $findings['vehicles'] = @{ ok = $false; status = $vehiclesPaged.Status; error = $vehiclesPaged.Error }
}
else {
    $vehicles = $vehiclesPaged.Rows
    $vehicleKeys = Get-KeyUnion $vehicles
    $statuses = @($vehicles | Where-Object { $_.vehicle_status } |
                  ForEach-Object { [string]$_.vehicle_status } | Sort-Object -Unique)

    # Ask for the columns the vehicle report needs. Unknown names are silently
    # ignored by the API, so anything that comes back is genuinely available.
    $wantedVehicleFields = @('status_reason', 'cost_center_id', 'cost_center_name',
                             'division_id', 'district_id', 'group_id', 'oos_since', 'odometer_type')
    $fieldList = 'id,name,' + ($wantedVehicleFields -join ',')
    $probeResult = Invoke-Ts -Path 'ThirdParty/Data/Fleet/Vehicles' -Query @{
        'fields[vehicles]' = $fieldList; rows = '5'; page = '1'
    }
    $bonus = @()
    if ($probeResult.Ok) {
        $returned = Get-KeyUnion (Get-TsRows $probeResult.Data)
        $bonus = @($wantedVehicleFields | Where-Object { $returned -contains $_ })
    }

    Write-Host "    $($vehicles.Count) vehicles; statuses: $($statuses -join ', ')"
    if ($bonus.Count -gt 0) {
        Write-Host "    bonus fields available: $($bonus -join ', ')" -ForegroundColor Green
    }
    else {
        Write-Host '    allowlist confirmed closed - no status_reason / cost center / OOS history' -ForegroundColor Yellow
    }

    $findings['vehicles'] = @{
        ok                            = $true
        row_count                     = $vehicles.Count
        observed_keys                 = $vehicleKeys
        distinct_vehicle_status       = $statuses
        with_shift_name               = @($vehicles | Where-Object { $_.shift_name }).Count
        unlisted_fields_that_returned = $bonus
    }
    Save-Sample $vehicles 'vehicles_sample.json'
}

# --- 5. employees -------------------------------------------------------------
Write-Host '[5/6] Pulling /Data/User/Employees for cost-center coverage ...'
$employees = @()
$empPaged = Invoke-TsPaged -Path 'ThirdParty/Data/User/Employees' -Query @{ include_disabled = 'false' }
if (-not $empPaged.Ok) {
    Write-Host "    FAILED (HTTP $($empPaged.Status)): $($empPaged.Error)" -ForegroundColor Red
    $findings['employees'] = @{ ok = $false; status = $empPaged.Status; error = $empPaged.Error }
}
else {
    $employees = $empPaged.Rows
    $withCc = @($employees | Where-Object { $_.cost_center_name })
    $distinctCc = @($withCc | ForEach-Object { [string]$_.cost_center_name } | Sort-Object -Unique)

    Write-Host "    $($employees.Count) employees; $($withCc.Count) carry cost_center_name ($($distinctCc.Count) distinct)"

    $findings['employees'] = @{
        ok                     = $true
        row_count              = $employees.Count
        observed_keys          = Get-KeyUnion $employees
        with_cost_center_name  = $withCc.Count
        distinct_cost_centers  = $distinctCc
        with_license_level     = @($employees | Where-Object { $_.license_level }).Count
    }
    Save-Sample $employees 'employees_sample.json'
}

# --- 6. cost-center resolution ------------------------------------------------
Write-Host '[6/6] Testing indirect cost-center resolution for trips ...'
if ($shifts.Count -eq 0 -or $trips.Count -eq 0 -or $employees.Count -eq 0) {
    Write-Host '    skipped - a prerequisite pull returned nothing' -ForegroundColor Yellow
    $findings['cost_center_resolution'] = @{ ok = $false; reason = 'prerequisite pull empty or failed' }
}
else {
    # trip.shift_name -> shift.user_id -> employee.cost_center_name
    $empCc = @{}
    foreach ($e in $employees) {
        if ($e.user_id -and $e.cost_center_name) { $empCc[[string]$e.user_id] = [string]$e.cost_center_name }
    }
    $shiftCc = @{}
    foreach ($s in $shifts) {
        if (-not $s.shift_name -or -not $s.user_id) { continue }
        $cc = $empCc[[string]$s.user_id]
        if (-not $cc) { continue }
        $name = [string]$s.shift_name
        if (-not $shiftCc.ContainsKey($name)) {
            $shiftCc[$name] = New-Object System.Collections.Generic.HashSet[string]
        }
        [void]$shiftCc[$name].Add($cc)
    }

    $ambiguous = @($shiftCc.Keys | Where-Object { $shiftCc[$_].Count -gt 1 })
    $resolvable = @($trips | Where-Object { $_.shift_name -and $shiftCc.ContainsKey([string]$_.shift_name) }).Count
    $rate = 0.0
    if ($trips.Count -gt 0) { $rate = [Math]::Round($resolvable / $trips.Count, 3) }

    Write-Host "    $resolvable/$($trips.Count) trips resolve to a cost center via shift_name ($($ambiguous.Count) ambiguous profiles)"

    $findings['cost_center_resolution'] = @{
        ok                             = $true
        shift_profiles_mapped          = $shiftCc.Count
        ambiguous_shift_profiles       = $ambiguous.Count
        ambiguous_examples             = @($ambiguous | Select-Object -First 5)
        trips_resolvable_via_shift_name = $resolvable
        trips_total                    = $trips.Count
        resolution_rate                = $rate
    }
}

# ------------------------------------------------------------------- reporting
($findings | ConvertTo-Json -Depth 10) | Set-Content (Join-Path $OutDir 'findings.json') -Encoding UTF8

$md = New-Object System.Collections.Generic.List[string]
$md.Add('# Traumasoft ThirdParty API probe')
$md.Add('')
$md.Add("Tenant: ``$BaseUrl``")
$md.Add("Probe date (trip/shift target): **$Date**")
$md.Add('')

$md.Add('## Credentials')
$md.Add('')
$md.Add("- OK - $($findings.auth.organization_rows) organization rows")
$md.Add("- HMAC formula accepted: ``$($findings.auth.auth_mode)``")
$md.Add("- Secret used for signing: ``$($findings.auth.secret_source)`` (``api_key`` means the key doubles as its own secret)")
$md.Add('')

$md.Add('## Schedule/Shifts')
$md.Add('')
if ($findings.shifts.ok) {
    $md.Add("- Rows returned: **$($findings.shifts.row_count)**")
    $md.Add("- Distinct start dates in one pull: **$($findings.shifts.distinct_start_dates)**")
    $md.Add("- Span: ``$($findings.shifts.earliest_start_time)`` .. ``$($findings.shifts.latest_start_time)``")
    $md.Add("- Includes the requested date: **$($findings.shifts.covers_requested_date)**")
    $md.Add("- Server-side date filtering available: **$($findings.shifts.server_side_date_filtering)**")
    $missingText = 'none'
    if ($findings.shifts.missing_fields_needed_by_current_sql.Count -gt 0) {
        $missingText = $findings.shifts.missing_fields_needed_by_current_sql -join ', '
    }
    $md.Add("- Fields the current SQL needs but Shifts omits: ``$missingText``")
}
else { $md.Add("- FAILED (HTTP $($findings.shifts.status)): $($findings.shifts.error)") }
$md.Add('')

$md.Add('## Cad/Trip (GetTrips)')
$md.Add('')
if ($findings.trips.ok) {
    $tsText = 'none'
    if ($findings.trips.observed_timestamp_names.Count -gt 0) {
        $tsText = $findings.trips.observed_timestamp_names -join ', '
    }
    $md.Add("- Trip legs for the day: **$($findings.trips.row_count)**")
    $md.Add("- Legs carrying timestamps: **$($findings.trips.trips_with_timestamps)**")
    $md.Add("- Timestamp names observed: ``$tsText``")
    $md.Add("- Legs with ``late_reasons``: **$($findings.trips.with_late_reasons)**")
    $md.Add("- Legs with ``shift_name``: **$($findings.trips.with_shift_name)**")
}
else { $md.Add("- FAILED (HTTP $($findings.trips.status)): $($findings.trips.error)") }
$md.Add('')

$md.Add('## Fleet/Vehicles')
$md.Add('')
if ($findings.vehicles.ok) {
    $bonusText = 'none'
    if ($findings.vehicles.unlisted_fields_that_returned.Count -gt 0) {
        $bonusText = $findings.vehicles.unlisted_fields_that_returned -join ', '
    }
    $md.Add("- Vehicles: **$($findings.vehicles.row_count)**")
    $md.Add("- Distinct statuses: ``$($findings.vehicles.distinct_vehicle_status -join ', ')``")
    $md.Add("- Unlisted fields that actually returned: ``$bonusText``")
}
else { $md.Add("- FAILED (HTTP $($findings.vehicles.status)): $($findings.vehicles.error)") }
$md.Add('')

$md.Add('## User/Employees')
$md.Add('')
if ($findings.employees.ok) {
    $md.Add("- Employees: **$($findings.employees.row_count)**")
    $md.Add("- With ``cost_center_name``: **$($findings.employees.with_cost_center_name)** ($($findings.employees.distinct_cost_centers.Count) distinct)")
}
else { $md.Add("- FAILED (HTTP $($findings.employees.status)): $($findings.employees.error)") }
$md.Add('')

$md.Add('## Cost-center resolution for trips')
$md.Add('')
if ($findings.cost_center_resolution.ok) {
    $pct = [Math]::Round($findings.cost_center_resolution.resolution_rate * 100, 1)
    $md.Add("- Trips resolvable via ``shift_name``: **$($findings.cost_center_resolution.trips_resolvable_via_shift_name)/$($findings.cost_center_resolution.trips_total)** ($pct%)")
    $md.Add("- Shift profiles mapped: **$($findings.cost_center_resolution.shift_profiles_mapped)**, ambiguous: **$($findings.cost_center_resolution.ambiguous_shift_profiles)**")
}
else { $md.Add("- Not evaluated: $($findings.cost_center_resolution.reason)") }
$md.Add('')

$mdPath = Join-Path $OutDir 'PROBE_FINDINGS.md'
($md -join [Environment]::NewLine) | Set-Content -Path $mdPath -Encoding UTF8

Write-Host ''
Write-Host "Findings written to $mdPath" -ForegroundColor Green
Write-Host 'Raw samples are alongside it as *_sample.json.'
Write-Host ''
