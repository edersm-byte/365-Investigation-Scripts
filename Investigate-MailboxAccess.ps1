<#
.SYNOPSIS
    Investigates mailbox access and outbound mail from a specific source IP for a single
    user, and produces a Word-compatible report (.doc) plus a CSV data appendix.

.DESCRIPTION
    For a given mailbox and a source IP under review, this tool:
      * Pulls MailItemsAccessed audit events and identifies folders accessed and the
        individual messages opened from that IP.
      * Resolves opened-message subjects via message trace (last 90 days); older
        messages are listed by MessageId for eDiscovery lookup.
      * Pulls outbound message-trace records sent from that IP (deliveries + distinct).
      * Writes a formatted Word-compatible report and a CSV appendix to a folder.

    PREREQUISITES
      * Exchange Online PowerShell module 3.7.0 or later (Get-MessageTraceV2).
      * Roles: View-Only Audit Logs (or Audit Logs) to read the unified audit log.
      * MailItemsAccessed depends on mailbox auditing being enabled for the user; if it
        is not, the "opened messages" sections will be empty even if access occurred.

    NOTES
      * Message trace returns a maximum of 10 days per query; this script pages
        automatically, but trace only reaches back ~90 days total.
      * Mailbox auditing records that a message was ACCESSED, not that it was read.
      * The .doc opens in Word; use Save As .docx or Print > Save as PDF to distribute.

.PARAMETER User
    UPN of the mailbox to investigate, e.g. user@contoso.com

.PARAMETER IP
    Source IP address to scope the investigation to.

.PARAMETER Days
    Look-back window in days for access and sent mail. Default 10 (a single message-trace
    query returns up to 10 days; larger values are paged automatically up to ~90 days).

.PARAMETER OutputFolder
    Folder to write the report and CSV into. Default: current directory.

.PARAMETER Classification
    Classification banner text shown on the report. Default: CONFIDENTIAL - RESTRICTED DISTRIBUTION.

.EXAMPLE
    ./Investigate-MailboxAccess.ps1 -User jsmith@contoso.com -IP 203.0.113.5

.EXAMPLE
    ./Investigate-MailboxAccess.ps1 -User jsmith@contoso.com -IP 203.0.113.5 -Days 7 -OutputFolder ~/cases
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$User,
    [Parameter(Mandatory)][string]$IP,
    [int]$Days = 10,
    [int]$TraceDays = 90,
    [string]$OutputFolder = "$PWD",
    [string]$Classification = "CONFIDENTIAL - RESTRICTED DISTRIBUTION"
)

# ── Prerequisites ──────────────────────────────────────────────────────
$mod = Get-Module ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
if (-not $mod -or $mod.Version -lt [version]"3.7.0") {
    Write-Error "Requires ExchangeOnlineManagement 3.7.0+ (for Get-MessageTraceV2). Run: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force"
    return
}
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Not connected to Exchange Online - launching sign-in..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

$stamp     = Get-Date -Format "yyyyMMdd-HHmmss"
$userAlias = ($User -split '@')[0]
$aStart    = (Get-Date).AddDays(-$Days); $aEnd = Get-Date
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null }

function Get-RecIP ($d) { "$($d.ClientIP)$($d.ClientIPAddress)$($d.ActorIpAddress)" }
function Get-Prop ($d,$n) { ($d.OperationProperties | Where-Object { $_.Name -eq $n }).Value }
function Norm ($id) { if ($id) { $id.Trim().Trim('<','>') } }
function HtmlEnc ($s) { if ($null -eq $s) { "" } else { $s -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' } }

# ── Pull access events ─────────────────────────────────────────────────
Write-Host "`nPulling MailItemsAccessed for $User ..." -ForegroundColor Cyan
$mail   = Search-UnifiedAuditLog -StartDate $aStart -EndDate $aEnd -UserIds $User -Operations MailItemsAccessed -ResultSize 5000 -SessionCommand ReturnLargeSet | Sort-Object Identity -Unique
$ipHits = $mail | Where-Object { (Get-RecIP ($_.AuditData | ConvertFrom-Json)) -like "*$IP*" }
$throttled = [bool]($ipHits | Where-Object { (Get-Prop ($_.AuditData|ConvertFrom-Json) "IsThrottled") -eq "True" })

if (-not $ipHits) {
    Write-Host "No MailItemsAccessed records from $IP in the last $Days days." -ForegroundColor Yellow
    Write-Host "(If access is expected, confirm mailbox auditing is enabled: Get-Mailbox $User | fl AuditEnabled,AuditOwner)" -ForegroundColor DarkYellow
}

# Folder summary
$folderRows = foreach ($r in $ipHits) {
    $d=$r.AuditData|ConvertFrom-Json; $type=Get-Prop $d "MailAccessType"
    foreach ($f in $d.Folders) { [PSCustomObject]@{ Folder=$f.Path; AccessType=$type; Items=@($f.FolderItems).Count } }
}
$folderSummary = $folderRows | Group-Object Folder | ForEach-Object {
    [PSCustomObject]@{ Folder=$_.Name
        AccessTypes=(($_.Group.AccessType|Sort-Object -Unique) -join ", ")
        MsgsOpened=($_.Group|?{$_.AccessType -eq "Bind"}|Measure-Object Items -Sum).Sum
        SyncedWhole=if($_.Group.AccessType -contains "Sync"){"YES"}else{""} }
} | Sort-Object Folder
$syncedAny = [bool]($folderSummary.SyncedWhole -contains "YES")

# Opened messages
$bound = foreach ($r in $ipHits) {
    $d=$r.AuditData|ConvertFrom-Json
    if ((Get-Prop $d "MailAccessType") -eq "Bind") {
        foreach ($f in $d.Folders) { foreach ($i in $f.FolderItems) {
            [PSCustomObject]@{ Folder=$f.Path; MessageId=$i.InternetMessageId; Accessed=$r.CreationDate } } }
    }
}
$bound = $bound | Sort-Object MessageId -Unique

# Message-trace lookup for subject resolution (paged 10-day windows)
Write-Host "Building $TraceDays-day message-trace lookup..." -ForegroundColor Cyan
$lookup=@{}; $te=Get-Date
for ($i=0; $i -lt [math]::Ceiling($TraceDays/10); $i++) {
    $ws=$te.AddDays(-10)
    foreach ($p in @(@{RecipientAddress=$User},@{SenderAddress=$User})) {
        Get-MessageTraceV2 @p -StartDate $ws -EndDate $te -ResultSize 5000 -ErrorAction SilentlyContinue |
          ForEach-Object { $k=Norm $_.MessageId; if ($k -and -not $lookup.ContainsKey($k)){ $lookup[$k]=$_ } }
    }
    $te=$ws
}

$report = foreach ($b in $bound) {
    $k=Norm $b.MessageId
    if ($k -and $lookup.ContainsKey($k)) {
        $t=$lookup[$k]
        [PSCustomObject]@{ Folder=$b.Folder; AccessedUTC=$b.Accessed; Status="Resolved"
            ReceivedUTC=$t.Received; Sender=$t.SenderAddress; Subject=$t.Subject; MessageId=$b.MessageId }
    } else {
        [PSCustomObject]@{ Folder=$b.Folder; AccessedUTC=$b.Accessed; Status="Needs eDiscovery"
            ReceivedUTC=""; Sender=""; Subject=""; MessageId=$b.MessageId }
    }
}
$report   = $report | Sort-Object Folder, AccessedUTC
$resolved = @($report | Where-Object Status -eq "Resolved")
$unres    = @($report | Where-Object Status -eq "Needs eDiscovery")
$nTotal=@($report).Count; $nRes=$resolved.Count; $nUnres=$unres.Count

# ── Outbound mail from the IP (paged 10-day windows) ───────────────────
$sentRows=@(); $te=$aEnd
while ($te -gt $aStart) {
    $ws=$te.AddDays(-10); if ($ws -lt $aStart) { $ws=$aStart }
    $sentRows += Get-MessageTraceV2 -SenderAddress $User -StartDate $ws -EndDate $te -ResultSize 5000 -ErrorAction SilentlyContinue |
                 Where-Object { $_.FromIP -like "*$IP*" }
    $te=$ws
}
$sent = @($sentRows | Select-Object @{N='SentUTC';E={$_.Received}}, RecipientAddress, Subject, Status | Sort-Object SentUTC)
$nSent = $sent.Count
$nSentMsgs = @($sentRows | Select-Object -ExpandProperty MessageTraceId -Unique).Count

# ── Terminal view ──────────────────────────────────────────────────────
Write-Host "`n=== FOLDERS accessed from $IP ===" -ForegroundColor Yellow
$folderSummary | Format-Table -AutoSize
Write-Host "=== $nRes messages opened (subjects resolved) ===" -ForegroundColor Yellow
$resolved | Format-Table Folder,AccessedUTC,Sender,Subject -AutoSize -Wrap
Write-Host "=== $nUnres UNRESOLVED - look up by MessageId in eDiscovery ===" -ForegroundColor Green
$unres | Format-Table Folder,AccessedUTC,MessageId -AutoSize -Wrap
Write-Host "=== $nSent deliveries SENT from $IP ($nSentMsgs distinct messages) ===" -ForegroundColor Green
$sent | Format-Table SentUTC,RecipientAddress,Subject,Status -AutoSize -Wrap

# ── CSV appendix ───────────────────────────────────────────────────────
$csv = Join-Path $OutputFolder "$userAlias-access-review-$stamp.csv"
$report | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 -Delimiter ","

# ── Word (.doc) report ─────────────────────────────────────────────────
$foldTbl = "<table><tr><th>Folder</th><th>Access type</th><th>Msgs opened</th><th>Folder fully synced</th></tr>"
foreach ($f in $folderSummary) {
    $sw = if ($f.SyncedWhole -eq "YES") { "<b class='warn'>YES - assume all items taken</b>" } else { "-" }
    $foldTbl += "<tr><td>$(HtmlEnc $f.Folder)</td><td>$($f.AccessTypes)</td><td>$($f.MsgsOpened)</td><td>$sw</td></tr>"
}
$foldTbl += "</table>"

$msgTbl = "<table><tr><th>Folder</th><th>Accessed (UTC)</th><th>Sender</th><th>Subject</th></tr>"
foreach ($m in $resolved) { $msgTbl += "<tr><td>$(HtmlEnc $m.Folder)</td><td>$($m.AccessedUTC)</td><td>$(HtmlEnc $m.Sender)</td><td>$(HtmlEnc $m.Subject)</td></tr>" }
$msgTbl += "</table>"

$edTbl = "<table><tr><th>Folder</th><th>Accessed (UTC)</th><th>MessageId</th></tr>"
foreach ($m in $unres) { $edTbl += "<tr><td>$(HtmlEnc $m.Folder)</td><td>$($m.AccessedUTC)</td><td>$(HtmlEnc $m.MessageId)</td></tr>" }
$edTbl += "</table>"

$sentTbl = "<table><tr><th>Sent (UTC)</th><th>Recipient</th><th>Subject</th><th>Status</th></tr>"
foreach ($m in $sent) { $sentTbl += "<tr><td>$($m.SentUTC)</td><td>$(HtmlEnc $m.RecipientAddress)</td><td>$(HtmlEnc $m.Subject)</td><td>$($m.Status)</td></tr>" }
$sentTbl += "</table>"

$syncLine = if ($syncedAny) {
    "<b class='warn'>One or more folders were downloaded in full (sync); every item in those folders should be treated as exfiltrated.</b>"
} else { "No whole-folder (sync) download was recorded." }
$throttleLine = if ($throttled) {
    "<b class='warn'>Audit throttling was detected; the list is incomplete and full-mailbox exposure should be assumed.</b>"
} else { "No audit throttling was detected; the enumerated list is complete for the reviewed window." }

$doc = @"
<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'><title>$userAlias access review</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->
<style>
 body{font-family:'Calibri',Arial,sans-serif;font-size:11pt;color:#000;}
 .cls{background:#7a0000;color:#fff;font-weight:bold;text-align:center;padding:6px;letter-spacing:1px;}
 h1{font-size:18pt;margin:14pt 0 2pt 0;}
 h2{font-size:13pt;color:#1f3864;border-bottom:1px solid #1f3864;margin-top:18pt;}
 table{border-collapse:collapse;width:100%;font-size:9.5pt;margin-top:6pt;}
 th{background:#1f3864;color:#fff;text-align:left;padding:4px 6px;}
 td{border:0.5pt solid #bbb;padding:3px 6px;vertical-align:top;}
 .warn{color:#7a0000;}
 .meta td{border:none;padding:1px 8px 1px 0;font-size:10.5pt;}
 .foot{font-size:8.5pt;color:#555;margin-top:18pt;}
</style></head>
<body>
<div class='cls'>$Classification</div>
<h1>Mailbox Incident Investigation Report</h1>
<table class='meta'>
 <tr><td><b>Affected mailbox</b></td><td>$User</td></tr>
 <tr><td><b>Source IP</b></td><td>$IP</td></tr>
 <tr><td><b>Report generated</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm') (local)</td></tr>
</table>

<h2>Executive summary</h2>
<p>The mailbox <b>$User</b> was accessed from source IP <b>$IP</b>. A total of <b>$nTotal</b> individual messages were opened from that address across <b>$(@($folderSummary).Count)</b> folders; $nRes are resolved to subject lines below, and $nUnres are listed at the end for eDiscovery lookup. In addition, <b>$nSent</b> outbound deliveries ($nSentMsgs distinct messages) were sent from this address.</p>
<p>$syncLine $throttleLine</p>
<p><i>Note: mailbox auditing records that a message was accessed; it does not independently confirm the content was read.</i></p>

<h2>Folders accessed</h2>
<p style='font-size:9.5pt;color:#555;'><b>Access type:</b> <i>Bind</i> = individual messages opened one at a time (each is enumerated). <i>Sync</i> = the entire folder was downloaded at once, so all of its contents should be treated as taken.</p>
$foldTbl

<h2>Messages opened</h2>
$msgTbl

<h2>Unresolved - look up by MessageId in eDiscovery</h2>
$edTbl

<h2>Messages sent from this IP</h2>
<p style='font-size:9.5pt;color:#555;'>From message trace, covering $($aStart.ToString('yyyy-MM-dd')) to $($aEnd.ToString('yyyy-MM-dd')) UTC (last $Days days). Each row is one delivery per recipient; one message to multiple recipients appears on multiple rows. Messages sent before this window are not captured.</p>
$sentTbl

<div class='foot'>$Classification &middot; Generated $(Get-Date -Format 'yyyy-MM-dd HH:mm'). Contains sensitive investigative detail; handle and distribute accordingly.</div>
</body></html>
"@
$docFile = Join-Path $OutputFolder "$userAlias-access-review-$stamp.doc"
$doc | Out-File -FilePath $docFile -Encoding UTF8

Write-Host "`nFiles written to $OutputFolder :" -ForegroundColor Cyan
Write-Host "  $(Split-Path $docFile -Leaf)   <- open in Word; Save As .docx or Print > Save as PDF"
Write-Host "  $(Split-Path $csv -Leaf)   <- data appendix"
