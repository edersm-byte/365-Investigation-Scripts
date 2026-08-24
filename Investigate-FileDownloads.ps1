<#
.SYNOPSIS
    Finds SharePoint/OneDrive files downloaded, shared, or viewed by a user,
    optionally scoped to a source IP, over a look-back window.

.DESCRIPTION
    Queries the unified audit log for file download and access operations. File auditing
    is standard (not license-gated like MailItemsAccessed), so it works on any tenant with
    auditing enabled. Separates actual downloads (exfiltration signal) from view-only
    access (weaker evidence). Writes results to the terminal and a CSV.

    PREREQUISITES
      * Exchange Online PowerShell module (Search-UnifiedAuditLog).
      * Role: View-Only Audit Logs (or Audit Logs).

.PARAMETER User
    UPN of the account to investigate, e.g. user@contoso.com

.PARAMETER IP
    Optional. Source IP to scope to. If omitted, shows all file activity by the user.

.PARAMETER Days
    Look-back window in days (default 10; unified audit log retains 180 days).

.PARAMETER OutputFolder
    Folder for the CSV output. Default: current directory.

.EXAMPLE
    ./Investigate-FileDownloads.ps1 -User jsmith@contoso.com -IP 203.0.113.5

.EXAMPLE
    ./Investigate-FileDownloads.ps1 -User jsmith@contoso.com -Days 30
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$User,
    [string]$IP = "",
    [int]$Days = 10,
    [string]$OutputFolder = "$PWD"
)

# ── Prerequisites ──────────────────────────────────────────────────────
Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
if (-not (Get-Command Search-UnifiedAuditLog -ErrorAction SilentlyContinue)) {
    Write-Error "Search-UnifiedAuditLog not available. Install: Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force"
    return
}
if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
    Write-Host "Not connected to Exchange Online - launching sign-in..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false
}

$stamp     = Get-Date -Format "yyyyMMdd-HHmmss"
$userAlias = ($User -split '@')[0]
$aStart    = (Get-Date).AddDays(-$Days); $aEnd = Get-Date
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null }

$dlOps    = "FileDownloaded","FileSyncDownloadedFull"
$shareOps = "AnonymousLinkCreated","SharingSet"
$viewOps  = "FileAccessed","FileAccessedExtended","FilePreviewed"

$scope = if ($IP) { "from $IP" } else { "(all IPs)" }
Write-Host "`nSearching file activity for $User $scope over last $Days days..." -ForegroundColor Cyan

$raw = Search-UnifiedAuditLog -StartDate $aStart -EndDate $aEnd -UserIds $User `
        -Operations ($dlOps + $shareOps + $viewOps) -ResultSize 5000 -SessionCommand ReturnLargeSet |
       Sort-Object Identity -Unique

$rows = foreach ($r in $raw) {
    $d = $r.AuditData | ConvertFrom-Json
    $recIP = "$($d.ClientIP)$($d.ClientIPAddress)$($d.ActorIpAddress)"
    if ($IP -and ($recIP -notlike "*$IP*")) { continue }
    [PSCustomObject]@{
        Time       = $r.CreationDate
        Operation  = $r.Operations
        Kind       = if ($r.Operations -in $dlOps) { "DOWNLOAD" } elseif ($r.Operations -in $shareOps) { "SHARE" } else { "view" }
        File       = $d.SourceFileName
        Site       = $d.SiteUrl
        SharedWith = $d.TargetUserOrGroupName
        Path       = $d.ObjectId
        IP         = $recIP
        UserAgent  = $d.UserAgent
    }
}
$rows = @($rows | Sort-Object Time)

$downloads = @($rows | Where-Object Kind -eq "DOWNLOAD")
$shares    = @($rows | Where-Object Kind -eq "SHARE")
$views     = @($rows | Where-Object Kind -eq "view")

if (-not $rows) {
    Write-Host "`nNo file download or access records found for the given scope." -ForegroundColor Yellow
    Write-Host "(A clean result here is genuine - file auditing is standard and not license-dependent.)" -ForegroundColor DarkYellow
    return
}

# ── Terminal output ────────────────────────────────────────────────────
Write-Host "`n=== $($downloads.Count) FILES DOWNLOADED (copy left the tenant) ===" -ForegroundColor Red
$downloads | Format-Table Time,Operation,File,Site,UserAgent -AutoSize -Wrap

Write-Host "`n=== $($shares.Count) SHARING LINKS / PERMISSIONS created (persistent external exposure) ===" -ForegroundColor Red
$shares | Format-Table Time,Operation,File,SharedWith,Site -AutoSize -Wrap

Write-Host "`n=== $($views.Count) files VIEWED / PREVIEWED (opened, no confirmed download) ===" -ForegroundColor Yellow
$views | Format-Table Time,Operation,File,Site -AutoSize -Wrap

# ── CSV ────────────────────────────────────────────────────────────────
$csv = Join-Path $OutputFolder "$userAlias-file-activity-$stamp.csv"
$rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8 -Delimiter ","
Write-Host "`nCSV written to: $csv" -ForegroundColor Cyan
