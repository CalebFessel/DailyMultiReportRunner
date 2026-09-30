<#
.SYNOPSIS
  Samsara Connected Workflows (Forms) truck-check duration metrics.

.DESCRIPTION
  Connected Workflows are exposed through the Forms API. This script lists form
  templates, streams form submissions over a date range, and reports duration
  metrics using Samsara's native durationMs field (client-side elapsed time from
  form-open to submit), falling back to submittedAtTime - createdAtTime where
  durationMs is absent.

  Token needs "Read Form Submissions" AND "Read Form Templates" under Forms.

.EXAMPLE
  # Step 1 - discover which template is the truck check
  .\Get-WorkflowCheckMetrics.ps1 -ListTemplatesOnly

.EXAMPLE
  # Step 2 - full metrics, optionally narrowed to the truck-check template
  .\Get-WorkflowCheckMetrics.ps1 -StartTime 2026-05-01T00:00:00Z -FormTemplateIds "abc-123" -OutDir "$HOME\Desktop\check-metrics"
#>

[CmdletBinding()]
param(
    [string]$StartTime = "2026-05-01T00:00:00Z",
    [string]$EndTime   = "2026-08-25T00:00:00Z",
    [string]$BaseUrl   = "https://api.samsara.com",
    [string[]]$FormTemplateIds,
    [int]$MaxDurationMinutes = 120,
    [int]$ShortTailSeconds   = 60,
    [switch]$ListTemplatesOnly,
    [string]$OutDir = "."
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ([string]::IsNullOrEmpty($env:SAMSARA_API_TOKEN)) {
    throw "SAMSARA_API_TOKEN is not set in this session."
}
$headers = @{ Authorization = "Bearer $env:SAMSARA_API_TOKEN" }

function ConvertTo-Utc {
    param([string]$Text)
    $styles = [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal
    return [datetime]::Parse($Text, [Globalization.CultureInfo]::InvariantCulture, $styles)
}

function Get-Percentile {
    param([double[]]$Values, [double]$P)
    if (-not $Values -or $Values.Count -eq 0) { return $null }
    $s = @($Values | Sort-Object)
    if ($s.Count -eq 1) { return $s[0] }
    $rank = ($P / 100.0) * ($s.Count - 1)
    $lo = [int][math]::Floor($rank)
    $hi = [int][math]::Ceiling($rank)
    if ($lo -eq $hi) { return $s[$lo] }
    $frac = $rank - $lo
    return ($s[$lo] * (1 - $frac)) + ($s[$hi] * $frac)
}

function Invoke-Samsara {
    param([string]$Url)
    for ($attempt = 1; $attempt -le 6; $attempt++) {
        try {
            return Invoke-RestMethod -Uri $Url -Headers $headers -ErrorAction Stop
        } catch {
            $code = $null
            if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
            if ($code -eq 403) {
                throw "403 Forbidden. Token is missing Forms scope. Enable 'Read Form Submissions' and 'Read Form Templates' on the API token."
            }
            if (($code -eq 429 -or $code -ge 500) -and $attempt -lt 6) {
                $wait = [int][math]::Pow(2, $attempt)
                Write-Warning "HTTP $code; retry $attempt in $wait s"
                Start-Sleep -Seconds $wait
                continue
            }
            throw
        }
    }
}

function Get-Stats {
    param([string]$Label, [object[]]$Rows, [int]$ShortTail)
    $secs = [double[]]@($Rows | ForEach-Object { $_.DurationSeconds })
    $short = @($secs | Where-Object { $_ -lt $ShortTail }).Count
    $native = @($Rows | Where-Object { $_.DurationSource -eq "durationMs" }).Count
    $pct = 0.0
    $cov = 0.0
    if ($secs.Count -gt 0) {
        $pct = 100.0 * $short / $secs.Count
        $cov = 100.0 * $native / $secs.Count
    }
    return [PSCustomObject]@{
        Group          = $Label
        Count          = $secs.Count
        MedianMin      = [math]::Round((Get-Percentile -Values $secs -P 50) / 60.0, 2)
        MeanMin        = [math]::Round((($secs | Measure-Object -Average).Average) / 60.0, 2)
        P90Min         = [math]::Round((Get-Percentile -Values $secs -P 90) / 60.0, 2)
        MinMin         = [math]::Round((($secs | Measure-Object -Minimum).Minimum) / 60.0, 2)
        MaxMin         = [math]::Round((($secs | Measure-Object -Maximum).Maximum) / 60.0, 2)
        ShortTailCount = $short
        ShortTailPct   = [math]::Round($pct, 1)
        NativeDurPct   = [math]::Round($cov, 1)
    }
}

# ---------------- templates ----------------
Write-Host "Fetching form templates..."
$templates = @{}
$after = $null
do {
    $url = "$BaseUrl/form-templates"
    if ($after) { $url = $url + "?after=" + [uri]::EscapeDataString($after) }
    $resp = Invoke-Samsara -Url $url
    foreach ($t in @($resp.data)) {
        $templates[[string]$t.id] = [PSCustomObject]@{
            Id       = $t.id
            Title    = $t.title
            Category = $t.formCategory
        }
    }
    $after = $null
    if ($resp.pagination -and $resp.pagination.hasNextPage) { $after = $resp.pagination.endCursor }
} while ($after)

Write-Host ("Found {0} form template(s):" -f $templates.Count)
$templates.Values | Sort-Object Category, Title | Format-Table Id, Category, Title -AutoSize

if ($ListTemplatesOnly) {
    Write-Host "Pick the truck-check template Id above, then re-run with -FormTemplateIds '<id>'."
    return
}

# ---------------- submissions ----------------
$all   = New-Object System.Collections.Generic.List[object]
$after = $null
$page  = 0

$baseQuery = "$BaseUrl/form-submissions/stream?startTime=$StartTime&endTime=$EndTime"
if ($FormTemplateIds -and $FormTemplateIds.Count -gt 0) {
    $baseQuery = $baseQuery + "&formTemplateIds=" + [uri]::EscapeDataString(($FormTemplateIds -join ","))
}

do {
    $url = $baseQuery
    if ($after) { $url = $url + "&after=" + [uri]::EscapeDataString($after) }
    $resp = Invoke-Samsara -Url $url

    $batch = @($resp.data)
    foreach ($s in $batch) { [void]$all.Add($s) }
    $page = $page + 1
    Write-Host ("page {0}: +{1}  (total {2})" -f $page, $batch.Count, $all.Count)

    $after = $null
    if ($resp.pagination -and $resp.pagination.hasNextPage) { $after = $resp.pagination.endCursor }
} while ($after)

Write-Host ""
Write-Host ("Fetched {0} submission(s) across {1} page(s)." -f $all.Count, $page)

if ($all.Count -eq 0) {
    Write-Warning "No submissions in this window. Widen -StartTime or drop -FormTemplateIds."
    return
}

Write-Host "Status breakdown:"
$all | Group-Object status | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize

# ---------------- transform ----------------
$rows = New-Object System.Collections.Generic.List[object]
$exNoDuration = 0
$exNegative   = 0
$exTooLong    = 0

foreach ($s in $all) {
    $sec = $null
    $src = $null

    if ($s.PSObject.Properties.Name -contains "durationMs" -and $s.durationMs) {
        $sec = [double]$s.durationMs / 1000.0
        $src = "durationMs"
    } elseif ($s.createdAtTime -and $s.submittedAtTime) {
        $c = ConvertTo-Utc $s.createdAtTime
        $b = ConvertTo-Utc $s.submittedAtTime
        $sec = ($b - $c).TotalSeconds
        $src = "timestampDelta"
    }

    if ($null -eq $sec) { $exNoDuration = $exNoDuration + 1; continue }
    if ($sec -lt 0)     { $exNegative   = $exNegative + 1;   continue }
    if ($sec -gt ($MaxDurationMinutes * 60)) { $exTooLong = $exTooLong + 1; continue }

    $tplId = $null
    $tplTitle = $null
    if ($s.formTemplate) {
        $tplId = [string]$s.formTemplate.id
        $tplTitle = $s.formTemplate.title
    }
    if (-not $tplTitle -and $tplId -and $templates.ContainsKey($tplId)) { $tplTitle = $templates[$tplId].Title }
    if (-not $tplTitle) { $tplTitle = "(unknown template)" }

    $who = $null
    if ($s.submittedBy) {
        $who = $s.submittedBy.name
        if (-not $who) { $who = $s.submittedBy.id }
    }
    if (-not $who) { $who = "(unknown)" }

    $assetName = $null
    if ($s.asset) { $assetName = $s.asset.name }

    [void]$rows.Add([PSCustomObject]@{
        SubmissionId    = $s.id
        TemplateId      = $tplId
        TemplateTitle   = $tplTitle
        SubmittedBy     = $who
        Asset           = $assetName
        Status          = $s.status
        Score           = $s.score
        CreatedAtUtc    = $s.createdAtTime
        SubmittedAtUtc  = $s.submittedAtTime
        DurationSource  = $src
        DurationSeconds = [math]::Round($sec, 1)
        DurationMinutes = [math]::Round($sec / 60.0, 2)
    })
}

Write-Host ("Usable: {0}   Excluded -> no duration: {1}, negative: {2}, over {3} min: {4}" -f $rows.Count, $exNoDuration, $exNegative, $MaxDurationMinutes, $exTooLong)

if ($rows.Count -eq 0) {
    Write-Warning "No usable rows. Nothing to summarize."
    return
}

# ---------------- summarize ----------------
$overall = Get-Stats -Label "ALL" -Rows $rows -ShortTail $ShortTailSeconds

$byTemplate = @($rows | Group-Object TemplateTitle | Sort-Object Count -Descending | ForEach-Object {
    Get-Stats -Label $_.Name -Rows $_.Group -ShortTail $ShortTailSeconds
})

$byPerson = @($rows | Group-Object SubmittedBy | ForEach-Object {
    Get-Stats -Label $_.Name -Rows $_.Group -ShortTail $ShortTailSeconds
} | Sort-Object MedianMin -Descending)

Write-Host ""
Write-Host "=== OVERALL ===" -ForegroundColor Cyan
$overall | Format-List

Write-Host "=== BY TEMPLATE ===" -ForegroundColor Cyan
$byTemplate | Format-Table -AutoSize

Write-Host "=== BY PERSON (top 25 by median) ===" -ForegroundColor Cyan
$byPerson | Select-Object -First 25 | Format-Table -AutoSize

# ---------------- export ----------------
if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$rows        | Export-Csv (Join-Path $OutDir "check_durations_raw.csv")        -NoTypeInformation -Encoding UTF8
$byTemplate  | Export-Csv (Join-Path $OutDir "check_summary_by_template.csv")  -NoTypeInformation -Encoding UTF8
$byPerson    | Export-Csv (Join-Path $OutDir "check_summary_by_person.csv")    -NoTypeInformation -Encoding UTF8
@($overall)  | Export-Csv (Join-Path $OutDir "check_summary_overall.csv")      -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("Wrote 4 CSVs to {0}" -f (Resolve-Path $OutDir))
