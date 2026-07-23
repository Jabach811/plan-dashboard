# Prints one plan's email for reading during a narrative refresh.
#
#   .\plan-emails.ps1 Bonnier              # everything unreviewed, oldest first
#   .\plan-emails.ps1 Bonnier -All         # the whole thread
#   .\plan-emails.ps1 Bonnier -Tail 40     # the last 40 regardless of watermark
#
# Auto-replies and empty bodies are dropped -- they carry no narrative and are a
# large share of the volume on a busy plan.

param(
  [Parameter(Mandatory)][string]$Plan,
  [switch]$All,
  [int]$Tail = 0,
  [int]$Chars = 1500
)

$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSCommandPath
$dataDir = Join-Path $root 'data'
$rows = @(Import-Csv (Join-Path $dataDir 'corpus.csv') | Where-Object { $_.SubFolder -eq $Plan })
if (-not $rows) { throw "No plan folder named '$Plan' in the corpus." }

$rows = @($rows | Sort-Object SentDate, SentTime, @{Expression={[int]$_.Seq}})
$total = $rows.Count

$rows = @($rows | Where-Object {
    $_.Subject -notmatch '^\s*(Automatic reply|Undeliverable)\b' -and ($_.Body -replace '\s','').Length -gt 40
})

if ($Tail -gt 0) {
    $rows = @($rows | Select-Object -Last $Tail)
} elseif (-not $All) {
    $revPath = Join-Path $dataDir 'reviewed.json'
    $n = 0
    if (Test-Path $revPath) {
        $j = Get-Content $revPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($j.PSObject.Properties.Name -contains $Plan) { $n = [int]$j.$Plan.n }
    }
    $unreviewed = $total - $n
    if ($unreviewed -gt 0) { $rows = @($rows | Select-Object -Last $unreviewed) }
    else { $rows = @() }
}

Write-Host "### $Plan -- showing $($rows.Count) of $total emails (auto-replies and empty bodies dropped)`n"
foreach ($r in $rows) {
    $b = $r.Body -replace '\s*~\s*', "`n"
    $b = $b -replace '<(mailto:|https?://)[^>]*>', '' -replace '\[cid:[^\]]*\]', ''
    $b = [regex]::Replace($b, '(\r?\n){3,}', "`n`n").Trim()
    if ($b.Length -gt $Chars) { $b = $b.Substring(0, $Chars).TrimEnd() + '...' }
    Write-Host ("-" * 78)
    Write-Host "$($r.SentDate) $($r.SentTime)  |  $($r.SenderName)"
    Write-Host "SUBJ: $($r.Subject)"
    Write-Host ""
    Write-Host $b
}
