<#
.SYNOPSIS
  Determines whether PATCH /addresses/{id} is a true partial update.
  Self-restoring: changes one radius, reads it back, then puts it exactly back.

.DESCRIPTION
  Samsara does not document whether omitted fields survive a PATCH, or whether
  tagIds replaces rather than appends. That single unknown is why the 138
  curated addresses were skipped during the import, and it is what blocks a
  bulk geofence resize.

  This answers it empirically on ONE address from our own import (tagged
  "TraumaSoft Facilities"), never on a curated one. It snapshots the full
  record, changes only geofence.circle.radiusMeters, re-reads, diffs every
  field, then restores the original radius.

  Token needs Read/Write Addresses.
#>

[CmdletBinding()]
param(
    [string]$BaseUrl   = "https://api.samsara.com",
    [string]$AddressId,
    [string]$ImportTag = "TraumaSoft Facilities",
    [int]$TestRadiusM  = 137,
    [string]$OutDir    = "C:\Users\CalebFessel\Desktop\check-metrics"
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
if ([string]::IsNullOrEmpty($env:SAMSARA_API_TOKEN)) { throw "SAMSARA_API_TOKEN is not set." }
$H = @{ Authorization = "Bearer $env:SAMSARA_API_TOKEN"; "Content-Type" = "application/json" }

function Call {
    param([string]$Method="GET", [string]$Url, $Body)
    try {
        if ($null -ne $Body) {
            return Invoke-RestMethod -Method $Method -Uri $Url -Headers $H -ErrorAction Stop -Body ($Body | ConvertTo-Json -Depth 10 -Compress)
        }
        return Invoke-RestMethod -Method $Method -Uri $Url -Headers $H -ErrorAction Stop
    } catch {
        $c = $null; if ($_.Exception.Response) { $c = [int]$_.Exception.Response.StatusCode }
        throw ("HTTP {0} on {1} {2} :: {3}" -f $c, $Method, $Url, $_.ErrorDetails.Message)
    }
}

# ---- pick a target from our own import ----
if (-not $AddressId) {
    Write-Host "Finding an address from the '$ImportTag' import..."
    $all = New-Object System.Collections.Generic.List[object]
    $after = $null
    do {
        $u = "$BaseUrl/addresses"; if ($after) { $u += "?after=" + [uri]::EscapeDataString($after) }
        $r = Call -Url $u
        foreach ($d in @($r.data)) { [void]$all.Add($d) }
        $after = $null
        if ($r.pagination -and $r.pagination.hasNextPage) { $after = $r.pagination.endCursor }
    } while ($after)

    $cand = $all | Where-Object {
        $_.externalIds -and $_.externalIds.'traumasoftFacilityId' -and
        ($_.tags | ForEach-Object { $_.name }) -contains $ImportTag
    } | Select-Object -First 1

    if (-not $cand) { throw "No address found carrying the '$ImportTag' tag and a traumasoftFacilityId. Pass -AddressId explicitly." }
    $AddressId = $cand.id
}

$before = (Call -Url "$BaseUrl/addresses/$AddressId").data
Write-Host ("target: {0}  (id {1})" -f $before.name, $AddressId)

$origRadius = $null
if ($before.geofence -and $before.geofence.circle) { $origRadius = [int]$before.geofence.circle.radiusMeters }
if (-not $origRadius) { throw "Target has no circle geofence - pick a different address." }
Write-Host ("current radius: {0} m   test radius: {1} m" -f $origRadius, $TestRadiusM)

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$snap = Join-Path $OutDir "patch_test_before.json"
$before | ConvertTo-Json -Depth 12 | Out-File $snap -Encoding utf8
Write-Host "snapshot saved: $snap"

# ---- minimal PATCH: geofence only ----
$body = @{ geofence = @{ circle = @{
    latitude     = [double]$before.geofence.circle.latitude
    longitude    = [double]$before.geofence.circle.longitude
    radiusMeters = [int]$TestRadiusM } } }

Write-Host "`nPATCHing with geofence ONLY (no name, no tags, no externalIds)..."
$null = Call -Method PATCH -Url "$BaseUrl/addresses/$AddressId" -Body $body
Start-Sleep -Seconds 1
$after2 = (Call -Url "$BaseUrl/addresses/$AddressId").data

# ---- diff ----
function TagNames($a) { if ($a.tags) { return (@($a.tags | ForEach-Object { $_.name }) | Sort-Object) -join "|" } return "" }
function ExtIds($a)   { if ($a.externalIds) { return ($a.externalIds | ConvertTo-Json -Compress) } return "" }
function AddrTypes($a){ if ($a.addressTypes) { return (@($a.addressTypes) | Sort-Object) -join "|" } return "" }

$checks = @(
  @{ n="name";             b=$before.name;                a=$after2.name }
  @{ n="formattedAddress"; b=$before.formattedAddress;     a=$after2.formattedAddress }
  @{ n="tags";             b=(TagNames $before);           a=(TagNames $after2) }
  @{ n="externalIds";      b=(ExtIds $before);             a=(ExtIds $after2) }
  @{ n="addressTypes";     b=(AddrTypes $before);          a=(AddrTypes $after2) }
  @{ n="notes";            b=$before.notes;                a=$after2.notes }
  @{ n="latitude";         b=$before.latitude;             a=$after2.latitude }
  @{ n="longitude";        b=$before.longitude;            a=$after2.longitude }
)

Write-Host "`n=== FIELD SURVIVAL AFTER PARTIAL PATCH ===" -ForegroundColor Cyan
$lost = 0
foreach ($c in $checks) {
    $same = ("" + $c.b) -eq ("" + $c.a)
    if (-not $same) { $lost++ }
    "{0,-18} {1}   before='{2}'  after='{3}'" -f $c.n, $(if($same){"KEPT   "}else{"CHANGED"}), ("" + $c.b), ("" + $c.a)
}
$newRadius = $after2.geofence.circle.radiusMeters
"{0,-18} {1}   {2} -> {3}" -f "radiusMeters", $(if ([int]$newRadius -eq $TestRadiusM) { "APPLIED" } else { "FAILED " }), $origRadius, $newRadius

# ---- restore ----
Write-Host "`nRestoring original radius ($origRadius m)..."
$restore = @{ geofence = @{ circle = @{
    latitude     = [double]$before.geofence.circle.latitude
    longitude    = [double]$before.geofence.circle.longitude
    radiusMeters = [int]$origRadius } } }
$null = Call -Method PATCH -Url "$BaseUrl/addresses/$AddressId" -Body $restore
$final = (Call -Url "$BaseUrl/addresses/$AddressId").data
"radius now: {0} m  {1}" -f $final.geofence.circle.radiusMeters, $(if ([int]$final.geofence.circle.radiusMeters -eq $origRadius) { "(restored)" } else { "(RESTORE FAILED - fix by hand)" })

Write-Host ""
if ($lost -eq 0) {
    Write-Host "VERDICT: PATCH is a true partial update. Bulk resize is safe," -ForegroundColor Green
    Write-Host "         and the 138 curated addresses could be resized too." -ForegroundColor Green
} else {
    Write-Host ("VERDICT: PATCH CLOBBERS {0} field(s) above. A bulk resize must re-send" -f $lost) -ForegroundColor Red
    Write-Host "         every field, read from the current record first." -ForegroundColor Red
}
