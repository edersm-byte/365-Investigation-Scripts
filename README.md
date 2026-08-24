# M365 Investigation Scripts

PowerShell tooling for Microsoft 365 incident response. These scripts query the unified audit log and message trace to answer two questions that come up in every suspected account compromise: what mail was accessed from a given IP, and what files were downloaded or shared.

Written and tested on macOS with `pwsh`.

## Scripts

| Script | Purpose | Output |
| --- | --- | --- |
| `Investigate-MailboxAccess.ps1` | Mailbox access and outbound mail from a specific source IP for a single user. | Word-compatible `.doc` report + CSV appendix |
| `Investigate-FileDownloads.ps1` | SharePoint/OneDrive files downloaded, shared, or viewed by a user, optionally scoped to a source IP. | CSV |

## Requirements

- PowerShell 7 or later
- `ExchangeOnlineManagement` **3.7.0 or later** — `Investigate-MailboxAccess.ps1` uses `Get-MessageTraceV2` and will refuse to run on older versions. `Investigate-FileDownloads.ps1` only needs `Search-UnifiedAuditLog` and is less version-sensitive.
- Role: **View-Only Audit Logs** or **Audit Logs**
- Unified audit logging enabled in the tenant

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
```

Both scripts check for an existing Exchange Online session and launch sign-in themselves if there isn't one, so connecting first is optional:

```powershell
Connect-ExchangeOnline -UserPrincipalName you@yourdomain.com
```

## Investigate-MailboxAccess.ps1

Pulls `MailItemsAccessed` events, identifies which folders were touched and which individual messages were opened from the IP under review, then resolves those messages to subject lines via message trace. Messages older than the trace window are listed by `MessageId` for eDiscovery lookup. Also pulls outbound deliveries sent from the same IP.

| Parameter | Required | Default | Description |
| --- | --- | --- | --- |
| `-User` | Yes | — | UPN of the mailbox to investigate |
| `-IP` | Yes | — | Source IP to scope the investigation to |
| `-Days` | No | `10` | Look-back window for access and sent mail |
| `-TraceDays` | No | `90` | How far back to build the message-trace subject lookup |
| `-OutputFolder` | No | current directory | Where the report and CSV are written |
| `-Classification` | No | `CONFIDENTIAL - RESTRICTED DISTRIBUTION` | Banner text on the report |

```powershell
./Investigate-MailboxAccess.ps1 -User jsmith@contoso.com -IP 203.0.113.5

./Investigate-MailboxAccess.ps1 -User jsmith@contoso.com -IP 203.0.113.5 -Days 7 -OutputFolder ~/cases
```

Writes `<alias>-access-review-<timestamp>.doc` and `<alias>-access-review-<timestamp>.csv`. The `.doc` opens in Word — use Save As `.docx` or Print → Save as PDF before distributing.

### Reading the report

The folder table distinguishes two access types, and the difference matters:

- **Bind** — individual messages opened one at a time. Each one is enumerated in the report.
- **Sync** — the entire folder was downloaded at once. Nothing is enumerated because everything in that folder should be treated as taken.

The report flags both whole-folder syncs and audit throttling in the executive summary. If throttling was detected, the enumerated list is incomplete by definition and full-mailbox exposure should be assumed.

## Investigate-FileDownloads.ps1

Queries the unified audit log for file download, sharing, and view operations, then separates them by weight of evidence: downloads and sharing-link creation are exfiltration signals, views are weaker.

| Parameter | Required | Default | Description |
| --- | --- | --- | --- |
| `-User` | Yes | — | UPN of the account to investigate |
| `-IP` | No | all IPs | Source IP to scope to |
| `-Days` | No | `10` | Look-back window |
| `-OutputFolder` | No | current directory | Where the CSV is written |

```powershell
./Investigate-FileDownloads.ps1 -User jsmith@contoso.com -IP 203.0.113.5

./Investigate-FileDownloads.ps1 -User jsmith@contoso.com -Days 30
```

Writes `<alias>-file-activity-<timestamp>.csv`. Terminal output is grouped into downloads, sharing links, and views.

## Notes and limitations

- **Mailbox auditing must be enabled** for the user or `MailItemsAccessed` returns nothing, even where access genuinely occurred. Check with `Get-Mailbox <user> | fl AuditEnabled,AuditOwner` before treating an empty result as clean. Availability also depends on your tenant's audit plan — verify what you retain rather than assuming.
- **Access is not the same as reading.** Mailbox auditing records that an item was accessed. It does not confirm the content was read.
- **Throttling suppresses records.** Once a mailbox exceeds roughly 1,000 bind events in 24 hours, further events stop being logged for the rest of the period. The mailbox script detects and flags this.
- **Message trace reaches back about 90 days** and returns a maximum of 10 days per query. The script pages automatically.
- **The unified audit log retains 180 days** by default, so file activity can be queried further back than mail subjects can be resolved.
- **File auditing is not license-gated** the way mailbox auditing is, so a clean result from `Investigate-FileDownloads.ps1` is more trustworthy than a clean result from the mailbox script.
- Both scripts are read-only. Neither modifies mailboxes, accounts, or tenant configuration.


## Disclaimer

Provided as-is, without warranty. Review any script before running it against a production tenant, and confirm you're authorized to access the audit data you're querying.

## License

MIT — see [LICENSE](LICENSE).
