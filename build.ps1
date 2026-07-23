# Rebuilds the dashboard from every extract in .\extracts
#
#   Drop a new export into .\extracts and run:  .\build.ps1
#
# Emails already seen are discarded, so extracts may overlap freely. Plans
# absent from a newer extract keep everything already known about them --
# which is what makes active-only exports safe.
#
# Folder map:
#   extracts\    raw Outlook exports (input, Joel)
#   data\        corpus + every state file, authored and machine-written alike
#   templates\   the skin. -Template names one; retired\ holds older skins
#   logs\        ingest-log.txt, overwritten each run
#   studies\     design explorations, not part of the build
#   nbi-tracker\ the standalone NBI tracker skill and its worked example
#   .backups\    pre-change copies, never read by anything

param(
  [string]$Template = 'template2.html',
  [string]$Out      = 'plan-dashboard-v2.html'
)

$ErrorActionPreference = 'Stop'
$root       = Split-Path -Parent $PSCommandPath
$dataDir    = Join-Path $root 'data'
$extractDir = Join-Path $root 'extracts'
$corpusPath = Join-Path $dataDir 'corpus.csv'
$narrPath   = Join-Path $dataDir 'narrative.json'
$loadPath   = Join-Path $dataDir 'loadstatus.json'
$revPath    = Join-Path $dataDir 'reviewed.json'
$chgPath    = Join-Path $dataDir 'loadchanges.json'
$nbiPath    = Join-Path $dataDir 'nbi-notes.json'
$tplPath    = Join-Path (Join-Path $root 'templates') $Template
$logPath    = Join-Path (Join-Path $root 'logs') 'ingest-log.txt'
$outPath    = if ([System.IO.Path]::IsPathRooted($Out)) { $Out } else { Join-Path $root $Out }

# ---------------------------------------------------------------- ingest
# An email's identity is a hash of the five fields below -- never a supplied id
# column. Exports vary in shape and a file missing the id column would import
# as zero new rows while appearing to succeed.
#
# $required must stay minimal: these five are the only fields the dashboard
# genuinely needs. Anything in $optional is kept when present and left blank
# when not, so columns can be dropped from the export without breaking a build.
$required = @('SubFolder','SentDate','SenderName','Subject','Body')
$optional = @('SentTime','ToRecips','CcRecips','Thread','AttachNames')
$cols     = @('Seq') + $required + $optional

function Get-RowKey($r) {
    $b = ($r.Body -replace '\s+', ' ').Trim()
    if ($b.Length -gt 300) { $b = $b.Substring(0, 300) }
    $s = ($r.Subject -replace '\s+', ' ').Trim()
    return ('{0}|{1}|{2}|{3}|{4}' -f $r.SubFolder, $r.SentDate, $r.SenderName, $s, $b).ToLower()
}
function Normalize($r, $seq) {
    $o = [ordered]@{}
    $have = $r.PSObject.Properties.Name
    foreach ($c in $cols) {
        if ($c -eq 'Seq')            { $o[$c] = $seq }
        elseif ($have -contains $c)  { $o[$c] = $r.$c }
        else                         { $o[$c] = '' }
    }
    return [PSCustomObject]$o
}

# Exports are ragged: a column that is empty gets dropped from the row rather
# than left blank, so rows vary in width while the header (when there is one)
# still lists every column. A header therefore cannot be trusted -- Import-Csv
# reads a short row against the full header and shifts every field after the
# gap, which corrupts silently. Raggedness, not header presence, picks the
# parser.
#
# Ragged rows are read by position instead. Every layout seen so far shares the
# same tail -- ToRecips, CcRecips, Subject, Thread, HasAttach, AttachNames,
# Importance, (BodyLength), Body -- so anchoring on HasAttach ('Y'/'N' followed
# by an Importance value) locates everything without a header. Which of the
# optional columns got dropped falls out of where ToRecips is expected to start.
#
# Dates in some exports are Excel serials (46210 / 0.5378) -- converted here so
# they match the corpus format and dedupe against emails already ingested.
# Importance is a word in some exports and Outlook's raw 0/1/2 in others.
$IMPORTANCE = @('Normal','High','Low')
function Is-Imp($v) { ($IMPORTANCE -contains $v) -or ($v -match '^[0-2]$') }

function From-Serial($v, $fmt) {
    $d = 0.0
    if ([double]::TryParse($v, [ref]$d)) { return [datetime]::FromOADate($d).ToString($fmt) }
    if ($v -match '^(\d{1,2}:\d{2}):\d{2}$') { return $matches[1] }
    return $v
}
function At($f, $i) { if ($i -ge 0 -and $i -lt $f.Count) { $f[$i] } else { '' } }
function Col($head, $names) {
    foreach ($n in $names) { $i = [array]::IndexOf($head, $n); if ($i -ge 0) { return $i } }
    return -1
}

function Read-Positional($lines, $start, $head) {
    # Headerless files are the 21-column Active Plans layout; those indexes are
    # the fallback when there is no header to read them off.
    $iPlan = if ($head) { Col $head @('PlanFolder') }   else { 0 }
    $iSub  = if ($head) { Col $head @('SubFolder') }    else { 1 }
    $iPath = if ($head) { Col $head @('FolderPath') }   else { 2 }
    $iDate = if ($head) { Col $head @('MessageSentDate','SentDate') } else { 4 }
    $iTime = if ($head) { Col $head @('MessageSentTime','SentTime') } else { 5 }
    $iFrom = if ($head) { Col $head @('SenderName') }   else { 10 }
    $iTo   = if ($head) { Col $head @('ToRecips') }     else { 12 }
    $hasLen = if ($head) { ($head -contains 'BodyLength') } else { $true }

    $out = New-Object System.Collections.ArrayList
    for ($n = $start; $n -lt $lines.Count; $n++) {
        $line = $lines[$n]
        if (-not $line.Trim()) { continue }
        $f = $line -split '\|'
        $h = -1
        for ($i = $f.Count - 1; $i -ge 3; $i--) {
            if ($f[$i] -ne 'Y' -and $f[$i] -ne 'N') { continue }
            if ((Is-Imp (At $f ($i+1))) -or (Is-Imp (At $f ($i+2)))) { $h = $i; break }
        }
        if ($h -lt 3) { continue }
        # AttachNames sits between HasAttach and Importance only when it survived.
        $imp = if (Is-Imp (At $f ($h+1))) { $h + 1 } else { $h + 2 }
        # ToRecips is never empty, so a row four wide before HasAttach kept its
        # Cc and a row three wide dropped it.
        $cc = if ($h - 4 -eq $iTo) { At $f ($h-3) } else { '' }
        [void]$out.Add([PSCustomObject]@{
            PlanFolder  = (At $f $iPlan)
            FolderPath  = (At $f $iPath)
            SubFolder   = (At $f $iSub)
            SentDate    = (From-Serial (At $f $iDate) 'yyyy-MM-dd')
            SentTime    = (From-Serial (At $f $iTime) 'HH:mm')
            SenderName  = (At $f $iFrom)
            ToRecips    = (At $f $iTo)
            CcRecips    = $cc
            Subject     = (At $f ($h-2))
            Thread      = (At $f ($h-1))
            AttachNames = if ($imp -gt $h + 1) { $f[$h+1] } else { '' }
            Importance  = (At $f $imp)
            # An empty body is dropped too, which leaves BodyLength as the last
            # field -- a bare number is never a body.
            Body        = if ($f.Count - 1 -gt ($imp + [int]$hasLen)) { $f[$f.Count - 1] } else { '' }
        })
    }
    return $out
}

function Read-Extract($path) {
    $lines = [System.IO.File]::ReadAllLines($path)
    if (-not $lines) { return @() }
    $lines[0] = $lines[0] -replace "^\xEF\xBB\xBF|^ï»¿", ''
    $head = @($lines[0] -split '\|')
    $hasHeader = ($head -contains 'SubFolder') -and ($head -contains 'Subject')
    if (-not $hasHeader) { return @(Read-Positional $lines 0 $null) }

    $ragged = $false
    for ($i = 1; $i -lt $lines.Count; $i++) {
        if (-not $lines[$i].Trim()) { continue }
        if (($lines[$i] -split '\|').Count -ne $head.Count) { $ragged = $true; break }
    }
    $canonical = -not @($required | Where-Object { $head -notcontains $_ })
    if ($canonical -and -not $ragged) { return @(Import-Csv $path -Delimiter '|') }
    return @(Read-Positional $lines 1 $head)
}

$existing = @()
if (Test-Path $corpusPath) { $existing = @(Import-Csv $corpusPath) }
$seen = [System.Collections.Generic.HashSet[string]]::new()
foreach ($r in $existing) { [void]$seen.Add((Get-RowKey $r)) }

# Ingest order is the tie-breaker whenever SentTime is absent, so "the most
# recent status email" stays deterministic across rebuilds.
$nextSeq = 0
foreach ($r in $existing) { if ([int]$r.Seq -gt $nextSeq) { $nextSeq = [int]$r.Seq } }

$files = @(Get-ChildItem $extractDir -Filter *.txt | Sort-Object Name)
if (-not $files) { throw "No .txt extracts found in $extractDir" }

$fresh = New-Object System.Collections.ArrayList
$perFile = @()
$warnings = @()
# An export taken from the "Active Plans" Outlook folder doubles as the roster
# of what is currently being worked. Only the newest such export counts, so a
# plan that has since been filed away drops off rather than staying active
# forever -- unlike the corpus, which only ever accumulates.
#
# FolderPath is the marker rather than PlanFolder: exports vary in which of the
# two they carry, but a path always spells out the folder the email was filed in.
$rosterCands = @()
foreach ($f in $files) {
    $rows = @(Read-Extract $f.FullName)
    if (-not $rows) { $warnings += "$($f.Name): file is empty or not pipe-delimited -- SKIPPED"; continue }
    $have = $rows[0].PSObject.Properties.Name
    $act = @($rows | Where-Object {
        $_.SubFolder -and (
            (($have -contains 'FolderPath') -and $_.FolderPath -match '\\Active Plans\\') -or
            (($have -contains 'PlanFolder') -and $_.PlanFolder -match 'Active Plans'))
    } | ForEach-Object { $_.SubFolder } | Sort-Object -Unique)
    if ($act) { $rosterCands += [PSCustomObject]@{ File = $f; Names = $act } }
    $have = $rows[0].PSObject.Properties.Name
    $missing = @($required | Where-Object { $have -notcontains $_ })
    if ($missing) {
        $warnings += "$($f.Name): missing required column(s) $($missing -join ', ') -- SKIPPED"
        $perFile += [PSCustomObject]@{ File = $f.Name; Rows = $rows.Count; New = 0 }
        continue
    }
    $added = 0
    foreach ($r in $rows) {
        if (-not $r.SubFolder) { continue }
        if ($seen.Add((Get-RowKey $r))) { $nextSeq++; [void]$fresh.Add((Normalize $r $nextSeq)); $added++ }
    }
    $perFile += [PSCustomObject]@{ File = $f.Name; Rows = $rows.Count; New = $added }
}

$roster = $null; $rosterFile = ''
if ($rosterCands) {
    $pick = @($rosterCands | Sort-Object { $_.File.LastWriteTime })[-1]
    $roster = [System.Collections.Generic.HashSet[string]]::new(
        [string[]]$pick.Names, [System.StringComparer]::OrdinalIgnoreCase)
    $rosterFile = $pick.File.Name
}

$all = @($existing) + @($fresh)
$newByPlan = @{}
foreach ($r in $fresh) {
    if (-not $newByPlan.ContainsKey($r.SubFolder)) { $newByPlan[$r.SubFolder] = 0 }
    $newByPlan[$r.SubFolder]++
}

$all | Select-Object $cols | Export-Csv $corpusPath -NoTypeInformation -Encoding UTF8

# --------------------------------------------------------- source emails
# Narrative entries may carry "src" pointers: {d, from, subj, q?} -- resolved
# here against the corpus so the dashboard can show the actual email. A self
# email (from Joel, subject starting NOTE) is flagged and, when present,
# suppresses every other source on the same entry -- it is the override.
$srcWarn = @()

function Clean-Body($raw) {
    $b = $raw -replace '\s*~\s*', "`n"
    $b = $b -replace '<(mailto:|https?://)[^>]*>', ''
    $b = $b -replace '\[cid:[^\]]*\]', ''
    $b = $b -replace '[ \t]{2,}', ' '
    $b = [regex]::Replace($b, '(\r?\n){3,}', "`n`n").Trim()
    if ($b.Length -gt 2600) { $b = $b.Substring(0, 2600).TrimEnd() + [char]0x2026 }
    return $b
}
function Resolve-Src($rows, $ptr, $key, $what) {
    $q = $null
    if ($ptr.PSObject.Properties.Name -contains 'q') { $q = $ptr.q }
    $hit = @($rows | Where-Object {
        $_.SentDate -eq $ptr.d -and
        $_.SenderName -match [regex]::Escape($ptr.from) -and
        $_.Subject -like ('*' + [System.Management.Automation.WildcardPattern]::Escape($ptr.subj) + '*') -and
        (-not $q -or $_.Body -like ('*' + [System.Management.Automation.WildcardPattern]::Escape($q) + '*'))
    } | Sort-Object @{Expression={[int]$_.Seq}})
    # A substantial body wins whenever one is available -- that keeps a loose
    # pointer off the empty auto-reply rows. But a one-line "Yes, that's
    # correct." is still a real answer, so a pointer specific enough to land
    # only on short rows resolves to them rather than to nothing.
    $rich = @($hit | Where-Object { ($_.Body -replace '\s', '').Length -gt 40 })
    if ($rich) { $hit = $rich }
    if (-not $hit) {
        $script:srcWarn += "{0}: '{1}' -- no email matches {2} / {3} / {4}" -f $key, $what, $ptr.d, $ptr.from, $ptr.subj
        return $null
    }
    $r = $hit[0]
    return [ordered]@{ from = $r.SenderName; d = $r.SentDate; time = $r.SentTime
                       subj = $r.Subject; body = (Clean-Body $r.Body); manual = (Is-Note $r) }
}
# A manual note is how a call, a meeting or a Teams thread gets onto a plan
# without touching narrative.json: mail it to yourself and file it in the plan's
# Outlook folder. Nothing has to be formatted -- the folder says which plan, the
# subject becomes the title, and the body is read as written, so provenance can
# just be stated in it ("this came from a phone call with Tricia").
#
# What identifies one is structural, not a convention that has to be remembered:
# sender is Joel and the only recipient is Joel. A real work email always goes
# to somebody else, so nothing else collides with that. An auto-reply bounces
# back to the sender alone too, hence the one exclusion.
$noteWarn = @()

function Is-Note($r) {
    if ($r.SenderName -notmatch 'Abacherli') { return $false }
    if ($r.CcRecips)                         { return $false }
    if ($r.Subject -match '^\s*Automatic reply\b') { return $false }
    # Batch upload reports land in the same shape and say so themselves.
    if ($r.Body -match 'This is a system generated message') { return $false }
    return (($r.ToRecips -replace '<[^>]*>', '' -replace '\s', '') -match '^Abacherli,Joel;?$')
}

# Joel's own signature block is on every note he writes and carries nothing.
function Strip-Sig($body) {
    $i = [regex]::Match($body, '(?m)^\s*Joel Abacherli\s*$')
    if ($i.Success) { return $body.Substring(0, $i.Index).TrimEnd() }
    return $body
}

function Read-Notes($rows, $key) {
    $out = @()
    foreach ($r in @($rows | Sort-Object SentDate, SentTime, @{Expression={[int]$_.Seq}})) {
        if (-not (Is-Note $r)) { continue }
        $body = Strip-Sig (Clean-Body $r.Body)
        $flat = ($body -replace '\s+', ' ').Trim()
        if (-not $flat) { continue }
        if ($flat.Length -gt 220) { $flat = $flat.Substring(0, 220).TrimEnd() + [char]0x2026 }
        $out += [PSCustomObject]@{
            title = $r.Subject; note = $flat
            d     = ([datetime]$r.SentDate).ToString('MMM d')
            src   = [ordered]@{ from = $r.SenderName; d = $r.SentDate; time = $r.SentTime
                                subj = $r.Subject; body = $body; manual = $true }
        }
    }
    return $out
}

function Set-Prop($o, $name, $value) {
    if ($o.PSObject.Properties.Name -contains $name) { $o.$name = $value }
    else { $o | Add-Member -NotePropertyName $name -NotePropertyValue $value }
}

# You closing things out by hand. This runs last, after the narrative has been
# resolved, so a later narrative refresh cannot quietly un-close an item -- the
# override file is the standing record of what you actioned, not the narrative.
$ovrPath = Join-Path $dataDir 'overrides.json'
$ovr = @{}
if (Test-Path $ovrPath) {
    $oj = Get-Content $ovrPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $oj.PSObject.Properties.Name) { $ovr[$k] = $oj.$k }
}
$ovrLog = @()

function Apply-Override($n, $key) {
    if (-not $ovr.ContainsKey($key)) { return }
    $o = $ovr[$key]
    $has = { param($f) $o.PSObject.Properties.Name -contains $f }

    if (& $has 'phase') { $n.phase = $o.phase }

    $clearAll = (& $has 'clear') -and ($o.clear -is [string]) -and ($o.clear -eq '*')
    $clear    = if ($clearAll -or -not (& $has 'clear')) { @() } else { @($o.clear) }
    $keep     = if (& $has 'keep')  { @($o.keep) }  else { @() }
    $watch    = if (& $has 'watch') { @($o.watch) } else { @() }
    $drop     = if (& $has 'drop')  { @($o.drop) }  else { @() }
    $why      = if (& $has 'resolved') { $o.resolved } else { $null }
    $on       = if (& $has 'on') { $o.on } else { (Get-Date -Format 'MMM dd') }

    # a watched item is still live but is nobody's action right now -- it shows on
    # the plan and is deliberately not counted, so it cannot hold a stage open
    $watching = New-Object System.Collections.ArrayList
    $kept = New-Object System.Collections.ArrayList
    $dropped = @()
    foreach ($it in @($n.open)) {
        $t = $it.task
        # dropped, not closed -- it was never a DC task, so it gets no milestone
        if (@($drop | Where-Object { $t -like "$_*" }).Count -gt 0) {
            $dropped += $t
            $script:ovrLog += ("  override: {0} -- dropped ""{1}""" -f $key, $t)
            continue
        }
        if (@($watch | Where-Object { $t -like "$_*" }).Count -gt 0) {
            [void]$watching.Add($it)
            $script:ovrLog += ("  override: {0} -- watching ""{1}""" -f $key, $t)
            continue
        }
        $isKeep  = @($keep  | Where-Object { $t -like "$_*" }).Count -gt 0
        $isClear = $clearAll -or @($clear | Where-Object { $t -like "$_*" }).Count -gt 0
        if ($isKeep -or -not $isClear) { [void]$kept.Add($it); continue }

        $note = if ($why -and ($why.PSObject.Properties.Name | Where-Object { $t -like "$_*" })) {
            $why.($why.PSObject.Properties.Name | Where-Object { $t -like "$_*" } | Select-Object -First 1)
        } else { 'Closed out.' }
        $n.history = @($n.history) + @([pscustomobject]@{ d = $on; t = $t; n = $note; src = @($it.src) })
        $script:ovrLog += ("  override: {0} -- closed ""{1}""" -f $key, $t)
    }
    $n.open = @($kept)
    Set-Prop $n 'watch' @($watching)

    # a clear or watch that matched nothing is a typo, not a no-op
    foreach ($c in $clear) {
        if (-not (@($n.history) | Where-Object { $_.t -like "$c*" })) {
            $script:ovrLog += ("  override: {0} -- NO MATCH for clear ""{1}""" -f $key, $c)
        }
    }
    foreach ($c in $watch) {
        if (-not (@($watching) | Where-Object { $_.task -like "$c*" })) {
            $script:ovrLog += ("  override: {0} -- NO MATCH for watch ""{1}""" -f $key, $c)
        }
    }
    foreach ($c in $drop) {
        if (-not (@($dropped) | Where-Object { $_ -like "$c*" })) {
            $script:ovrLog += ("  override: {0} -- NO MATCH for drop ""{1}""" -f $key, $c)
        }
    }
}

# Everything a narrative entry already points at, so a note that has been worked
# into the summary by hand is not also appended raw.
function Claimed-Notes($n) {
    $set = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($sec in @('open','history','decisions')) {
        if ($n.PSObject.Properties.Name -notcontains $sec) { continue }
        foreach ($it in @($n.$sec)) {
            if (-not $it -or ($it.PSObject.Properties.Name -notcontains 'src')) { continue }
            foreach ($s in @($it.src)) { if ($s.manual) { [void]$set.Add("$($s.d)|$($s.subj)") } }
        }
    }
    return ,$set
}

# An unclaimed note lands in history, dated. There is no marker in the email
# saying it is an open item or a decision, and guessing would be worse than
# recording it -- it gets promoted by hand at the next narrative refresh.
function Apply-Notes($n, $notes, $key) {
    $claimed = Claimed-Notes $n
    foreach ($note in $notes) {
        if ($claimed.Contains("$($note.src.d)|$($note.title)")) { continue }
        if ($n.PSObject.Properties.Name -notcontains 'history') {
            $n | Add-Member -NotePropertyName 'history' -NotePropertyValue @()
        }
        $new = [ordered]@{ d = $note.d; t = $note.title; n = $note.note; src = @($note.src) }
        $n.history = @($n.history) + @([PSCustomObject]$new)
        $script:noteWarn += "{0}: your note '{1}' ({2}) added to history -- promote it if it belongs in open or decisions" -f $key, $note.title, $note.d
    }
}

function Resolve-Entries($rows, $entries, $key) {
    foreach ($it in @($entries)) {
        if (-not $it) { continue }
        if (($it.PSObject.Properties.Name -notcontains 'src') -or -not $it.src) { continue }
        $what = if ($it.PSObject.Properties.Name -contains 'task') { $it.task } else { $it.t }
        $res = @(foreach ($ptr in @($it.src)) {
            $r = Resolve-Src $rows $ptr $key $what
            if ($r) { $r }
        })
        # Your own account of what happened outranks the email thread it came from.
        $mine = @($res | Where-Object { $_.manual })
        if ($mine) { $res = $mine }
        $it.src = $res
    }
}

# ------------------------------------------------- parse DC status tables
$fields = @('Account Balances','Shares Reregistering','Loans','Fortfeitures','Forfeitures','PCRA/SDA',
            'Hardship Basis',"Roth First Contrib Date","RMD's/Recurring Payments",'Investment Allocations',
            'Census','Deferrals','Auto Enroll','Outstanding Checks','Eligibility','Yrs of Service',
            'Vesting Overrides','YTD Hours','Beneficiary Data','EE exists on another plan with diff DOB',
            'Suspended ppts due to Hardship','YTD Contribs','YTD Comp','Alternate payee statuses and blocked accounts')
$order  = $fields | Where-Object { $_ -ne 'Fortfeitures' }
$rkSet  = $order[0..8]
$alt    = ($fields | ForEach-Object { [regex]::Escape($_) }) -join '|'

$loadStatus = @{}
foreach ($g in ($all | Group-Object SubFolder)) {
    $c = @($g.Group | Where-Object { $_.Body -match 'Account Balances' -and $_.Body -match 'Census' -and $_.SenderName -match 'Abacherli' })
    if (-not $c) { $c = @($g.Group | Where-Object { $_.Body -match 'Account Balances' -and $_.Body -match 'Census' }) }
    if (-not $c) { continue }
    $e = @($c | Sort-Object SentDate, SentTime, @{Expression={[int]$_.Seq}})[-1]
    $b = ($e.Body -replace '\s*~\s*', ' ') -replace '\s{2,}', ' '
    $items = [ordered]@{}
    foreach ($m in [regex]::Matches($b, "(?<lab>$alt)(?<rest>.*?)(?=($alt)|Conversion:|For additional information|$)")) {
        $lab  = $m.Groups['lab'].Value -replace 'Fortfeitures', 'Forfeitures'
        $rest = $m.Groups['rest'].Value.Trim()
        $date = ''
        if ($rest -match '^(\d{1,2}/\d{1,2}/\d{4})') { $date = $matches[1]; $rest = $rest.Substring($matches[1].Length).Trim() }
        elseif ($rest -match '^N/?A\b')              { $date = 'N/A';       $rest = ($rest -replace '^N/?A\b','').Trim() }
        $note = ($rest -replace '^[:\-\s]+','').Trim()
        $note = $note -replace '^(Start-Up No RecordKeeper|[A-Za-z0-9 .,&''\-]{2,45}?)(,?\s*[Aa]ll locations)?\s*:\s*',''
        $note = ($note -replace '\s{2,}',' ').Trim()
        # the note is the whole cell -- keep it verbatim so it can be copied out.
        # the cap only exists to catch a runaway last field that swallowed the
        # rest of the body when no closing label followed it.
        if ($note.Length -gt 2000) { $note = $note.Substring(0,2000).TrimEnd() + [char]0x2026 }
        if (-not $items.Contains($lab)) { $items[$lab] = @{ date = $date; note = $note } }
    }
    $loadStatus[$g.Name] = @{ asOf = $e.SentDate; items = $items
        src = [ordered]@{ from = $e.SenderName; d = $e.SentDate; time = $e.SentTime
                          subj = $e.Subject; body = (Clean-Body $e.Body); self = $false } }
}

# What the latest status table moved. The marks are only recomputed for a plan
# whose table actually changed this run -- build twice with no new email and the
# last refresh's marks still stand, the same way the narrative watermark works.
# A plan seen for the first time has no baseline, so nothing is marked.
$chgLog   = @()
$prevLoad = @{}
if (Test-Path $loadPath) {
    $pj = Get-Content $loadPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $pj.PSObject.Properties.Name) { $prevLoad[$k] = $pj.$k }
}
$changes = @{}
if (Test-Path $chgPath) {
    $cj = Get-Content $chgPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $cj.PSObject.Properties.Name) {
        $h = @{}; foreach ($f in $cj.$k.PSObject.Properties.Name) { $h[$f] = $cj.$k.$f }
        $changes[$k] = $h
    }
}
foreach ($k in @($loadStatus.Keys)) {
    if (-not $prevLoad.ContainsKey($k)) { continue }
    $cur = $loadStatus[$k].items
    $old = $prevLoad[$k].items
    $d = @{}
    foreach ($f in @($cur.Keys)) {
        $o = $old.PSObject.Properties[$f]
        if (-not $o)                                                            { $d[$f] = 'added' }
        elseif ($o.Value.date -ne $cur[$f].date -or $o.Value.note -ne $cur[$f].note) { $d[$f] = 'updated' }
    }
    if ($d.Count) { $changes[$k] = $d; $chgLog += ("  status refresh: {0} -- {1} line(s) moved" -f $k, $d.Count) }
}
$changes | ConvertTo-Json -Depth 4 | Set-Content $chgPath -Encoding UTF8
$loadStatus | ConvertTo-Json -Depth 6 | Set-Content $loadPath -Encoding UTF8

# Authored analysis, plan -> status-line label -> bullets. The status cell is raw
# source, not a note -- this is where the read of it lives. Anything unauthored
# falls back to the raw cell so nothing is silently dropped.
$nbiNotes = @{}
if (Test-Path $nbiPath) {
    $nj = Get-Content $nbiPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $nj.PSObject.Properties.Name) {
        $h = @{}
        foreach ($f in $nj.$k.PSObject.Properties.Name) { $h[$f] = @($nj.$k.$f) }
        $nbiNotes[$k] = $h
    }
}

# ------------------------------------------------------------- narrative
$narr = @{}
if (Test-Path $narrPath) {
    $j = Get-Content $narrPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $j.PSObject.Properties.Name) { $narr[$k] = $j.$k }
}

# How many emails a plan's narrative was last written against. "New in this run"
# is useless as a refresh signal -- build twice and it drops to zero while the
# narrative is still stale -- so the watermark lives in a file and only moves
# when a plan is actually re-summarized. Everything past it is unreviewed.
#
# A plan seen for the first time is seeded at its count *before* this run's
# import, so the email that just arrived is what marks it stale.
$reviewed = @{}
if (Test-Path $revPath) {
    $j = Get-Content $revPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($k in $j.PSObject.Properties.Name) { $reviewed[$k] = [int]$j.$k.n }
}

$stubbed = @()
$plans = New-Object System.Collections.ArrayList
foreach ($g in ($all | Group-Object SubFolder | Sort-Object Name)) {
    $key = $g.Name
    $s   = @($g.Group | Sort-Object SentDate, SentTime, @{Expression={[int]$_.Seq}})
    $top = (@($g.Group | Group-Object SenderName | Sort-Object Count -Descending | Select-Object -First 2) |
            ForEach-Object { ($_.Name -split ',')[0].Trim() }) -join ', '

    $n = $narr[$key]
    $isStub = $false
    if (-not $n) {
        $isStub = $true; $stubbed += $key
        $n = [PSCustomObject]@{
            name = $key; case = '--'; phase = 'kickoff'; kicker = 'not yet summarized'
            prior = '--'; payroll = '--'; com = '--'; client = '--'; golive = '--'
            headline = "New plan folder, $($g.Count) emails, not yet reviewed."
            open = @(); history = @(); decisions = @()
        }
    }

    Resolve-Entries $g.Group $n.open      $key
    Resolve-Entries $g.Group $n.history   $key
    Resolve-Entries $g.Group $n.decisions $key

    # after resolution -- a note carries its own already-built source chip, and
    # an overriding note must replace chips that Resolve-Entries just attached
    Apply-Notes $n (Read-Notes $g.Group $key) $key
    Apply-Override $n $key

    $loads = New-Object System.Collections.ArrayList
    $asOf  = ''
    $loadsSrc = $null
    if ($loadStatus.ContainsKey($key)) {
        $L = $loadStatus[$key]; $asOf = $L.asOf; $loadsSrc = $L.src
        foreach ($f in $order) {
            if (-not $L.items.Contains($f)) { continue }
            $grp = if ($rkSet -contains $f) { 'Recordkeeper data' } else { 'Conversion data' }
            $chg = if ($changes.ContainsKey($key) -and $changes[$key].ContainsKey($f)) { $changes[$key][$f] } else { '' }
            $bul = if ($nbiNotes.ContainsKey($key) -and $nbiNotes[$key].ContainsKey($f)) { $nbiNotes[$key][$f] } else { @() }
            [void]$loads.Add([ordered]@{ what = $f; group = $grp; date = $L.items[$f].date; note = $L.items[$f].note; chg = $chg; bul = @($bul) })
        }
    }

    if (-not $reviewed.ContainsKey($key)) { $reviewed[$key] = $g.Count - [int]$newByPlan[$key] }
    $pending = $g.Count - $reviewed[$key]
    if ($pending -lt 0) { $pending = 0 }

    [void]$plans.Add([ordered]@{
        id = $key; name = $n.name; case = $n.case; phase = $n.phase; kicker = $n.kicker
        prior = $n.prior; payroll = $n.payroll; com = $n.com; client = $n.client; golive = $n.golive
        headline = $n.headline; open = $n.open; history = $n.history; decisions = $n.decisions
        watch = $(if ($n.PSObject.Properties.Name -contains 'watch') { @($n.watch) } else { @() })
        owed = $n.owed; cap = $n.cap; cardNote = $n.cardNote; cardText = $n.cardText
        active = $(if ($roster) { $roster.Contains($key) } else { $n.active })
        loads = $loads; asOf = $asOf; loadsSrc = $loadsSrc; stub = $isStub
        emails = $g.Count; first = $s[0].SentDate; last = $s[-1].SentDate; top = $top
        pending = $pending
    })
}
$revOut = [ordered]@{}
foreach ($k in ($reviewed.Keys | Sort-Object)) { $revOut[$k] = @{ n = $reviewed[$k] } }
$revOut | ConvertTo-Json -Depth 3 | Set-Content $revPath -Encoding UTF8

# ----------------------------------------------------------------- build
$span  = @($all | Sort-Object SentDate)
$build = @{
    today    = (Get-Date -Format 'yyyy-MM-dd')
    emails   = $all.Count
    extracts = $files.Count
    first    = $span[0].SentDate
    last     = $span[-1].SentDate
    roster   = $rosterFile
}
$html = Get-Content $tplPath -Raw -Encoding UTF8
$html = $html.Replace('/*DATA*/',  ($plans | ConvertTo-Json -Depth 8 -Compress))
$html = $html.Replace('/*BUILD*/', ($build | ConvertTo-Json -Compress))
[System.IO.File]::WriteAllText($outPath, $html, (New-Object System.Text.UTF8Encoding($false)))

# ------------------------------------------------------------------- log
$log = New-Object System.Collections.ArrayList
[void]$log.Add("=== build $(Get-Date -Format 'yyyy-MM-dd HH:mm') ===")
foreach ($p in $perFile) { [void]$log.Add(("  {0,-46} {1,6} rows {2,6} new" -f $p.File, $p.Rows, $p.New)) }
if ($warnings.Count) {
    [void]$log.Add("  !! PROBLEM FILES -- nothing was imported from these:")
    foreach ($w in $warnings) { [void]$log.Add("     $w") }
}
[void]$log.Add("  corpus now $($all.Count) emails across $($plans.Count) plans")
if ($roster) {
    [void]$log.Add("  -- active roster ($($roster.Count) plans, from $rosterFile) --")
    foreach ($k in ($roster | Sort-Object)) { [void]$log.Add("     $k") }
    $orphan = @($roster | Where-Object { $_ -notin ($plans | ForEach-Object { $_.id }) })
    if ($orphan) {
        [void]$log.Add("  !! ACTIVE PLANS WITH NO EMAIL IN THE CORPUS -- they will not appear at all:")
        foreach ($k in $orphan) { [void]$log.Add("     $k") }
    }
} else {
    [void]$log.Add("  no Active Plans export present -- active view falls back to phase")
}
if ($newByPlan.Count) {
    [void]$log.Add("  -- new email in this run --")
    foreach ($k in ($newByPlan.Keys | Sort-Object { -$newByPlan[$_] })) {
        [void]$log.Add(("     {0,-34} +{1}" -f $k, $newByPlan[$k]))
    }
} else { [void]$log.Add("  no new email in this run") }

# The refresh list: unreviewed email, not new-this-run email. A plan stays on it
# across rebuilds until its narrative is actually rewritten.
foreach ($l in $chgLog) { [void]$log.Add($l) }
foreach ($l in $ovrLog) { [void]$log.Add($l) }

$hasText = 0; $authored = 0
foreach ($k in $loadStatus.Keys) {
    foreach ($f in $loadStatus[$k].items.Keys) {
        if (-not $loadStatus[$k].items[$f].note) { continue }
        $hasText++
        if ($nbiNotes.ContainsKey($k) -and $nbiNotes[$k].ContainsKey($f)) { $authored++ }
    }
}
[void]$log.Add(("  nbi notes: {0} of {1} status lines with text have an authored read" -f $authored, $hasText))
$due = @($plans | Where-Object { $_.pending -gt 0 } | Sort-Object { -$_.pending })
if ($due) {
    [void]$log.Add("  -- NARRATIVE REFRESH DUE ($($due.Count) plans) --")
    foreach ($p in $due) {
        $tag = if ($roster -and -not $p.active) { '  (not active)' } else { '' }
        [void]$log.Add(("     {0,-34} {1} unreviewed of {2}{3}" -f $p.id, $p.pending, $p.emails, $tag))
    }
} else { [void]$log.Add("  every narrative is current -- nothing to refresh") }
if ($srcWarn.Count) {
    [void]$log.Add("  !! SOURCE POINTERS THAT DID NOT RESOLVE (chip will be missing):")
    foreach ($w in $srcWarn) { [void]$log.Add("     $w") }
}
if ($noteWarn.Count) {
    [void]$log.Add("  -- NOTE emails --")
    foreach ($w in $noteWarn) { [void]$log.Add("     $w") }
}
if ($stubbed.Count) {
    [void]$log.Add("  -- NEW PLAN FOLDERS with no narrative yet --")
    foreach ($k in $stubbed) { [void]$log.Add("     $k") }
}
[void]$log.Add("")
$text = ($log -join "`r`n")
Add-Content $logPath $text -Encoding UTF8
Write-Host $text
Write-Host "wrote $outPath"

