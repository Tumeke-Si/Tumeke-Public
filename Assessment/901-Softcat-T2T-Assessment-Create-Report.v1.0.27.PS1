#Requires -Modules ImportExcel

<#
.SYNOPSIS
    Combines every CSV file collected by the Softcat T2T assessment
    scripts into a single Excel workbook, built from template.xlsx.

.DESCRIPTION
    RUN ORDER: run this last, after all the 2-Softcat-T2T-Assessment-*
    scripts - ideally after 9-Softcat-T2T-Assessment-Check-Output.v2.0.0.
    PS1 confirms nothing's missing. Combines every CSV file collected by
    this project's scripts into a single Excel workbook, built from
    template.xlsx.

    On start, prompts the consultant to enter their name (re-prompting if
    left blank) and saves it to $ConsultantName, for use later in the
    Step 6 report metadata.

    Steps:
      1. Confirms template.xlsx exists
      2. Confirms the CSV Files folder exists and contains every expected
         CSV file (same known-file list as
         9-Softcat-T2T-Assessment-Check-Output.v2.0.0.PS1 - keep both in
         sync if the data-collection scripts' outputs ever change)
      3. Confirms the ImportExcel PowerShell module is installed - if
         it's not, prints install instructions for the consultant and
         stops, rather than failing later with a less obvious error
      4. Imports each CSV one at a time into its own named variable, in
         alphabetical order by variable name (e.g. $AcceptedDomains,
         $AccountMailboxes, $AllMGDomains, ...), and prints each variable
         name with its record count - not the data itself
      5. Reads the tenant's DisplayName from the $Org variable, turns it
         into a safe file name, and copies template.xlsx to a NEW
         <SafeDisplayName>.Softcat-T2T-Assessment-Report-<yyyyMMdd-HHmmss>.xlsx
         every run - there's no override parameter, so a report name is
         always freshly generated and never reused from a prior run
      6. Resolves the tenant ID and SharePoint admin URL - the actual
         write of org name / tenant ID / SharePoint admin URL /
         consultant name / date / report file name into cells C3:C8, plus
         a live SUMPRODUCT formula counting enabled Member-type users
         into D12 (reading from the AllUsers table built in Step 12), on
         the report's existing "Data" worksheet (from template.xlsx), is
         deliberately deferred until after Steps 7-17, so "Data" ends up
         the last worksheet written (see "Step 6 (continued)" near the
         end of the script) and the AllUsers table already exists for
         D12's formula to reference
      7. Builds a combined mailbox + OneDrive + license view
         ($AccountSummary - Type, DisplayName, UserPrincipalName,
         PrimarySmtpAddress, Items, SizeMB, DeletedSizeMB, ArchiveSizeMB,
         OneDriveMB, Licenses, LitigationHold) by matching
         $AccountMailboxes to $OneDriveSites and $LicenseData per user,
         then writes it into the report file as an Excel Table on its
         own "AccountSummary" worksheet
      8. Builds a combined group + mailbox + SharePoint + Teams view
         ($GroupSummary - Type, DisplayName, PrimarySmtpAddress,
         SPOStorageMB, Items, SizeMB, AccessType, Members, Guests,
         SharePointSiteURL, Visibility, Archived, TeamChannels) by
         matching $UnifiedGroups (GroupID) to $GroupMailboxes (Guid, for
         Items/SizeMB), $SharePointSites (GroupId) and $Teams (GroupId)
         per group - Type is "Teams Channel" if the matched SharePoint
         site's Template starts with "TEAMCHANNEL", else "Team" if
         there's a matching Teams row, else "M365 Group" - then writes it
         into the report file as an Excel Table on its own "GroupSummary"
         worksheet
      9. Builds a SharePoint sites view ($SharePointSiteSummary - Title,
         Type, Template, StorageMB, QuotaMB, LastModifiedDate, Status,
         Sharing, Owner, Url, GroupId) from $SharePointSites, with a Type
         column of "Teams Channel" (Template starts with "TEAMCHANNEL"),
         "Dedicated SharePoint Site" (all-zero GroupId), "Team" (matching
         Teams.GroupId row), or "M365 Group" (anything else) - then
         writes it into the report file as an Excel Table on its own
         "SharePointSites" worksheet
     10. Builds a deduplicated device inventory view ($DeviceSummary -
         DisplayName, OperatingSystem, OSVersion, TrustType, SyncEnabled,
         ProfileType, RegisteredOwners, AccountEnabled,
         ApproximateLastSignInDateTime, ActiveLast30Days) from $Devices -
         DisplayName is made unique by grouping rows that share one and
         keeping only the row with the latest ApproximateLastSignInDateTime
         per group - with ActiveLast30Days true/false based on whether
         that date falls within 30 days of $reportCreationDate (or blank
         if the device has no parseable sign-in date), and TrustType
         mapped from Graph's raw join-type values to friendlier Entra
         terms (Workplace -> Entra Registered, ServerAd -> Entra Hybrid
         Joined, AzureAd -> Entra Joined; anything else passes through
         unchanged) - then writes it into the report file as an Excel
         Table on its own "Devices" worksheet
     11. Builds a OneDrive sites view ($OneDriveSiteSummary - User, Owner,
         StorageMB, QuotaMB, LastModifiedDate, Status, Sharing, Url) from
         $OneDriveSites - then writes it into the report file as an Excel
         Table on its own "OneDrives" worksheet
     12. Builds an all-users view ($AllUsersSummary - ID, DisplayName,
         UserPrincipalName, Mail, UserType, UsageLocation,
         AccountEnabled, IsDirSynced, Licenses) from $AllMGUsers, matched
         to $LicenseData by UserPrincipalName - replaces AllMGUsers.csv's
         LicenseStatus column (which just carries "Active"/"Active;Active"
         noise) with the actual license names - then writes it into the
         report file as an Excel Table on its own "AllUsers" worksheet
     13. Builds a contacts view ($ContactSummary - Name, DisplayName,
         Alias, PrimarySmtpAddress, RecipientTypeDetails, Identity) from
         $Contacts - a straight passthrough, every column kept - then
         writes it into the report file as an Excel Table on its own
         "Contacts" worksheet
     14. Builds a mail-enabled security groups view
         ($MailEnabledSecurityGroupSummary - Name, DisplayName, Alias,
         PrimarySmtpAddress, ManagedBy, OrganizationalUnit,
         RequireSenderAuthentication, HiddenFromAddressLists,
         MemberCount, RecipientTypeDetails, Identity) from
         $MailEnabledSecurityGroups - a straight passthrough, every
         column kept - then writes it into the report file as an Excel
         Table on its own "MailEnabledSecurityGroups" worksheet
     15. Builds a dynamic distribution groups view
         ($DynamicDistributionGroupSummary - Name, DisplayName,
         RecipientFilterType, RecipientTypeDetails, PrimarySmtpAddress,
         ManagedBy) from $DynamicDistributionGroups - a straight
         passthrough, every column kept - then writes it into the report
         file as an Excel Table on its own "DynamicDistributionGroups"
         worksheet
     16. Builds a distribution groups view ($DistributionGroupSummary -
         Name, DisplayName, PrimarySmtpAddress, GroupType, ManagedBy,
         MemberCount, HiddenFromAddressLists, IsDirSynced, Description,
         WhenCreated, WhenChanged) from $DistributionGroups - a curated
         subset, NOT every column: DistributionGroups.csv carries 38 raw
         Exchange properties (moderation settings, routing restrictions,
         message size limits, etc.) that mostly aren't useful in an
         assessment report - then writes it into the report file as an
         Excel Table on its own "DistributionGroups" worksheet
     17. Builds a security groups view ($SecurityGroupSummary -
         DisplayName, Description, Mail, MailEnabled, SecurityEnabled,
         GroupTypes, CreatedDateTime, MemberCount, Owners, Id) from
         $SecurityGroups - a straight passthrough, every column kept -
         then writes it into the report file as an Excel Table on its
         own "SecurityGroups" worksheet
     18. Builds an account mailboxes view ($AccountMailboxSummary - Type,
         DisplayName, UserPrincipalName, PrimarySmtpAddress, ItemCount,
         TotalItemSizeMB, DeletedItemCount, TotalDeletedItemSizeMB,
         Archive, ArchiveItemCount, ArchiveTotalItemSizeMB,
         LitigationHoldEnabled, RetentionPolicy, AccountSKU,
         WhenCreatedUTC, WhenChangedUTC, LastLogonTime) from
         $AccountMailboxes - a curated subset, NOT every column:
         AccountMailboxes.csv carries 28 raw Exchange properties, several
         of which (Alias, EmailAddresses, Guid, litigation/retention hold
         date fields) mostly aren't useful in an assessment report - then
         writes it into the report file as an Excel Table on its own
         "AccountMailboxes" worksheet
     19. Builds a group mailboxes view ($GroupMailboxSummary -
         DisplayName, Alias, PrimarySmtpAddress, RecipientTypeDetails,
         ItemCount, TotalItemSizeMB, DeletedItemCount,
         TotalDeletedItemSizeMB, WhenCreatedUTC, WhenChangedUTC,
         LastLogonTime, EmailAddresses, SPOSiteURL, Guid) from
         $GroupMailboxes - a straight passthrough, every column kept -
         then writes it into the report file as an Excel Table on its
         own "GroupMailboxes" worksheet
     20. Builds a combined domains view ($DomainSummary - DomainName,
         DomainType, IsDefault, IsCustomDomain, MatchSubDomains,
         AuthenticationType, DKIMEnabled, DKIMSelector1CNAME, HasDMARC,
         DMARCPolicy, DMARCRecord, MXRecordCount, MXRecords,
         SPFRecordCount, MultipleSPFRecords, SPFRecords) from
         $AcceptedDomains (driver), matched to $CustomDomains (for
         IsCustomDomain), $DKIMRecords, $DMARCRecords, $MXRecords and
         $SPFRecords, all by domain name - MX records are concatenated
         in preference order since more than one is normal, but more
         than one SPF record is flagged via MultipleSPFRecords since
         that's an actual RFC 7208 misconfiguration, not just noise -
         then writes it into the report file as an Excel Table on its
         own "Domains" worksheet

    At the end, asks whether to open the finished report.

    FRAMEWORK ONLY FOR NOW - the other CSV-imported variables from Step 4
    aren't yet written into the workbook as their own worksheets; only
    the Step 6 report data, $AccountSummary (Step 7), $GroupSummary
    (Step 8), $SharePointSiteSummary (Step 9), $DeviceSummary (Step 10),
    $OneDriveSiteSummary (Step 11), $AllUsersSummary (Step 12),
    $ContactSummary (Step 13), $MailEnabledSecurityGroupSummary
    (Step 14), $DynamicDistributionGroupSummary (Step 15),
    $DistributionGroupSummary (Step 16), $SecurityGroupSummary
    (Step 17), $AccountMailboxSummary (Step 18), $GroupMailboxSummary
    (Step 19) and $DomainSummary (Step 20) are written so far. That's
    deliberately deferred to a later revision.

    WARNING - Step 6 assumes Org.csv exports the tenant ID under a column
    named "Id" (the standard property Get-MgOrganization returns), and
    that the SharePoint Admin URL is available as a plain text file named
    AdminURL.txt in the CSV Files folder (exported by 2-Softcat-T2T-
    Assessment-EXO-1.v2.0.0.PS1, outside the CSV import list in Step 4).
    Neither has been confirmed against the real files - if missing or
    named differently, those two fields will come through blank rather
    than erroring, with a warning in the log.

    WARNING - Step 6 also assumes template.xlsx already has a worksheet
    literally named "Data", with row 3 laid out for these six values in
    column C. Unlike the AccountSummary/GroupSummary worksheets, this
    doesn't create the worksheet itself - if template.xlsx doesn't have
    one, this throws rather than writing silently to the wrong place.

    Step 7's assumptions that OneDriveSites.csv has an 'Owner' column
    containing the user's UPN, and that LicenseData.csv has
    UserPrincipalName and Licenses columns, were confirmed against
    sample exports from 2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1 and
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1, so no WARNING
    caveat there.

    Step 8's join columns (UnifiedGroups.GroupID, GroupMailboxes.Guid,
    SharePointSites.GroupId, Teams.GroupId) and the SharePointSites.
    Template "TEAMCHANNEL*" prefix used to derive Type were confirmed
    directly against sample exports from 2-Softcat-T2T-Assessment-EXO-1.
    v2.0.0.PS1, 2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1 and 2-Softcat-
    T2T-Assessment-Teams-1.v2.0.0.PS1, so no WARNING caveat there.

    Step 8's GroupMailboxes match is deliberately named $groupMailbox,
    not $mailbox - Step 7's $AccountSummary loop above also uses
    $mailbox as its foreach variable, and PowerShell foreach loops don't
    get their own scope, so a $mailbox reference inside Step 8 would
    silently read Step 7's last-iterated value instead of erroring.

    Step 9's Type logic was confirmed directly against a sample
    SharePointSites.csv export: sites for M365 Groups/Teams carry the
    group's real GroupId and a "GROUP#0" Template; Teams Channel sites
    carry a "TEAMCHANNEL#1" Template but an all-zero GroupId; and
    non-group sites (e.g. "SPSMSITEHOST#0", "SitePagePublishing#0",
    "SRCHCEN#0") also carry the all-zero GroupId - so the all-zero check
    only correctly means "Dedicated SharePoint Site" once the
    TEAMCHANNEL check has already been ruled out first, which is why
    that check comes first in the logic. So no WARNING caveat there.

    WARNING - Steps 5 and 6 assume Org.csv (exported by 1-Softcat-T2T-
    Assessment-Initial-Script.v2.0.0.PS1) has a "DisplayName" column,
    since that's the standard property Get-MgOrganization returns for a
    tenant's name. If the actual column is named differently, this will
    need adjusting.

    Step 10's dd/MM/yyyy HH:mm:ss parsing of
    Devices.ApproximateLastSignInDateTime was confirmed directly against
    a sample Devices.csv export. A device with a blank or unparseable
    value gets a blank ActiveLast30Days rather than being treated as
    active or inactive, with a warning logged for unparseable values.
    So no WARNING caveat there.

    Step 10 also deduplicates on DisplayName, keeping only the row with
    the latest ApproximateLastSignInDateTime per DisplayName (an undated
    row always loses to a dated one sharing its DisplayName). $Devices
    with a blank DisplayName would all collapse into a single group
    together under this logic - hasn't come up in the sample data, but
    worth knowing if Devices.csv ever has more than one blank-DisplayName
    row.

    ActiveLast30Days compares against $reportCreationDate (captured once
    near the top of the script), not a fresh Get-Date call at Step 10 -
    the same $reportCreationDate value is also written into the Data
    worksheet's C7 cell in Step 6 (continued), so both stay consistent
    with each other regardless of how long the run takes.

    WARNING - Step 10's TrustType mapping (Workplace/ServerAd/AzureAd ->
    Entra Registered/Hybrid Joined/Joined) was given directly, but only
    "AzureAd" has actually been seen in a sample Devices.csv export so
    far - "Workplace" and "ServerAd" are Graph's documented values for
    the other two join types, not yet confirmed against real data. Any
    TrustType value that doesn't match one of the three passes through
    unchanged rather than being blanked out, so nothing gets lost if the
    real value differs.

    Step 12's LicenseData.csv join (UserPrincipalName) is the same one
    confirmed for Step 7 above, so no WARNING caveat there either.

    Purely local - doesn't connect to Graph/Exchange/Teams/SharePoint, so
    no AppId/TenantId/certificate needed.

    Each worksheet write (AccountSummary, GroupSummary, SharePointSites,
    Devices, OneDrive, AllUsers, Contacts, MailEnabledSecurityGroups,
    DynamicDistributionGroups, DistributionGroups, SecurityGroups,
    AccountMailboxes, GroupMailboxes, Domains, and the deferred Data
    metadata write) is preceded by a short Start-Rest countdown (shown
    via Write-Progress) as a brief pause before touching the report
    file.

.PARAMETER TemplatePath
    Path to the Excel template. Defaults to template.xlsx next to
    this script.

.PARAMETER CsvFolderPath
    Folder containing the collected CSV files. Defaults to a
    "CSV Files" subfolder next to this script.

.PARAMETER OutputPath
    Base folder for outputs (the generated report and the Logs
    subfolder). Defaults to the same folder as this script.

.EXAMPLE
    .\901-Softcat-T2T-Assessment-Create-Report.v1.0.27.PS1

    Runs interactively, prompting for the consultant's name, building
    the report from template.xlsx and the CSV Files subfolder next to
    this script, and offering to open it when done.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    Prerequisites:
      Install-Module ImportExcel -Scope CurrentUser
      (this script also checks for the module itself and prints the
      same instruction if it's missing - see Step 3 below)
#>
[CmdletBinding()]
param(
    [string]$TemplatePath,

    [string]$CsvFolderPath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $scriptDir
}
if ([string]::IsNullOrWhiteSpace($TemplatePath)) {
    $TemplatePath = Join-Path $scriptDir "template.xlsx"
}
if ([string]::IsNullOrWhiteSpace($CsvFolderPath)) {
    $CsvFolderPath = Join-Path $scriptDir "CSV Files"
}

# Captured once, right at the start of the run, so every "how recent
# is this relative to the report" calculation across the whole script
# (Devices' ActiveLast30Days, the Data worksheet's C7 timestamp, etc.)
# uses the exact same moment rather than drifting between Get-Date
# calls made at different points as the script runs.
$reportCreationDate = Get-Date

#region Helper functions
Function Start-Rest {
    param(
        [int]$Seconds = 5,
        [string]$NextItem = "next step"
    )

    # Countdown
    for ($i = $Seconds; $i -ge 0; $i--) {
        $PercentComplete = (($Seconds - $i) / $Seconds) * 100
        Write-Progress -Activity "Having a little rest before writing $NextItem..." -Status "$i seconds remaining..." -PercentComplete $PercentComplete
        Start-Sleep -Seconds 1
    }
}

Function Show-Report {
    param([string]$OutputFile)

    $continue = Read-Host "Do you want to open the report? (Y/N)"
    if ($continue -notmatch '^[Yy]$') {
        Write-Host "Script execution has ended." -ForegroundColor Red
        return
    }
    Export-Excel -Path $OutputFile -Show
}
#endregion

#region Ensure Logs folder exists
$logsFolderPath = Join-Path $OutputPath "Logs"
if (-not (Test-Path $logsFolderPath)) {
    New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
}
#endregion

#region Logging
$logFilePath = Join-Path $logsFolderPath "901-Softcat-T2T-Assessment-Create-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("Warning", "Error")]
        [string]$Level = "Warning"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$Level] $Message" | Out-File -FilePath $logFilePath -Append -Encoding UTF8
    Write-Warning $Message
}
#endregion

#region Get consultant name
# Asked for up front and saved to $ConsultantName - used later when
# building the report metadata.
$ConsultantName = Read-Host -Prompt "Enter the consultant's name"
while ([string]::IsNullOrWhiteSpace($ConsultantName)) {
    Write-Host "Consultant name cannot be blank." -ForegroundColor Yellow
    $ConsultantName = Read-Host -Prompt "Enter the consultant's name"
}
$ConsultantName = $ConsultantName.Trim()
Write-Host "Consultant: $ConsultantName" -ForegroundColor Green
#endregion

#region Known expected CSV files, by source script
# Same list as 9-Softcat-T2T-Assessment-Check-Output.v2.0.0.PS1.
$ExpectedFiles = @(
    # 1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1
    [PSCustomObject]@{ FileName = "Org.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "AllMGDomains.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "AllMGUsers.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "LicenseData.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "LicenseOnePerLine.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "DetailedLicenseSummary.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "MXRecords.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "SPFRecords.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "DMARCRecords.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "TeamsUserActivity_Last180Days.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "SecurityGroups.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "RegisteredApplications.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "BitlockerKeys.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "Devices.csv"; SourceScript = "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1" }

    # 2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1
    [PSCustomObject]@{ FileName = "AcceptedDomains.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "CustomDomains.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "AccountMailboxes.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "UnifiedGroups.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "GroupMailboxes.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1" }

    # 2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1
    [PSCustomObject]@{ FileName = "DKIMRecords.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "InboundConnectors.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "OutboundConnectors.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "TransportRules.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "DistributionGroups.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "DynamicDistributionGroups.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "PublicFolders.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "MailEnabledSecurityGroups.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "Contacts.csv"; SourceScript = "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1" }

    # 2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1
    [PSCustomObject]@{ FileName = "Teams.csv"; SourceScript = "2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "TeamChannels.csv"; SourceScript = "2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "TeamsVoiceRoutingPolicies.csv"; SourceScript = "2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "TeamsPstnUsage.csv"; SourceScript = "2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "TeamsDialPlans.csv"; SourceScript = "2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1" }

    # 2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1
    [PSCustomObject]@{ FileName = "SharePointSites.csv"; SourceScript = "2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1" }
    [PSCustomObject]@{ FileName = "OneDriveSites.csv"; SourceScript = "2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1" }
)
#endregion

#region Step 1: confirm the template exists
Write-Host "Step 1 of 20: checking for template..." -ForegroundColor Cyan
Write-Host "Template path: $TemplatePath"

if (-not (Test-Path $TemplatePath)) {
    Write-Log -Level Error -Message "Template file not found at '$TemplatePath'. Cannot continue without it - use -TemplatePath to point at the correct location if it's not next to this script."
    throw "Template file not found: $TemplatePath"
}
Write-Host "Template found." -ForegroundColor Green
#endregion

#region Step 2: confirm the CSV folder exists and contains every expected file
Write-Host "`nStep 2 of 20: checking CSV Files folder..." -ForegroundColor Cyan
Write-Host "CSV folder path: $CsvFolderPath"

if (-not (Test-Path $CsvFolderPath)) {
    Write-Log -Level Error -Message "CSV Files folder not found at '$CsvFolderPath'. Run the 2-Softcat-T2T-Assessment-* scripts first (and 9-Softcat-T2T-Assessment-Check-Output.v2.0.0.PS1 to confirm) before running this script."
    throw "CSV Files folder not found: $CsvFolderPath"
}

$missingFiles = @($ExpectedFiles | Where-Object { -not (Test-Path (Join-Path $CsvFolderPath $_.FileName)) })

if ($missingFiles.Count -gt 0) {
    foreach ($missing in $missingFiles) {
        Write-Log -Message "Missing expected file '$($missing.FileName)' (produced by $($missing.SourceScript))."
    }
    Write-Log -Level Error -Message "$($missingFiles.Count) of $($ExpectedFiles.Count) expected CSV files are missing from '$CsvFolderPath'. Run 9-Softcat-T2T-Assessment-Check-Output.v2.0.0.PS1 for the full picture, then re-run whichever scripts produced the missing files, before running this script again."
    throw "$($missingFiles.Count) expected CSV file(s) missing - see the log for details."
}
Write-Host "All $($ExpectedFiles.Count) expected CSV files are present." -ForegroundColor Green
#endregion

#region Step 3: confirm the ImportExcel module is installed
Write-Host "`nStep 3 of 20: checking for the ImportExcel PowerShell module..." -ForegroundColor Cyan
$importExcelModule = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1

if (-not $importExcelModule) {
    Write-Host "`n============================================================" -ForegroundColor Red
    Write-Host " ImportExcel module not found" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "This script needs the ImportExcel PowerShell module to build the report, and it isn't installed for this user on this machine."
    Write-Host "`nTo install it, open PowerShell and run:"
    Write-Host "`n    Install-Module -Name ImportExcel -Scope CurrentUser`n" -ForegroundColor Yellow
    Write-Host "(Answer Y if prompted to trust the PSGallery repository.) Then re-run this script."
    Write-Log -Level Error -Message "ImportExcel module not found. Consultant needs to run: Install-Module -Name ImportExcel -Scope CurrentUser"
    throw "ImportExcel module not found - see console output above for install instructions."
}
Write-Host "ImportExcel module found (version $($importExcelModule.Version))." -ForegroundColor Green
Import-Module ImportExcel -MinimumVersion $importExcelModule.Version
#endregion

#region Step 4: import each CSV into its own named variable
Write-Host "`nStep 4 of 20: importing CSV data into variables..." -ForegroundColor Cyan

# Each import is explicit and one at a time (e.g. $Org = Import-Csv
# (Join-Path $CsvFolderPath "Org.csv")), matched to the file name as
# per the "Known expected CSV files" list above. Wrapped individually
# so one bad CSV logs a warning and leaves that variable an empty
# array, rather than stopping the whole run.

try { $AcceptedDomains = @(Import-Csv -Path (Join-Path $CsvFolderPath "AcceptedDomains.csv")) }
catch { Write-Log -Message "Failed to import 'AcceptedDomains.csv' into `$AcceptedDomains: $($_.Exception.Message)"; $AcceptedDomains = @() }

try { $AccountMailboxes = @(Import-Csv -Path (Join-Path $CsvFolderPath "AccountMailboxes.csv")) }
catch { Write-Log -Message "Failed to import 'AccountMailboxes.csv' into `$AccountMailboxes: $($_.Exception.Message)"; $AccountMailboxes = @() }

try { $AllMGDomains = @(Import-Csv -Path (Join-Path $CsvFolderPath "AllMGDomains.csv")) }
catch { Write-Log -Message "Failed to import 'AllMGDomains.csv' into `$AllMGDomains: $($_.Exception.Message)"; $AllMGDomains = @() }

try { $AllMGUsers = @(Import-Csv -Path (Join-Path $CsvFolderPath "AllMGUsers.csv")) }
catch { Write-Log -Message "Failed to import 'AllMGUsers.csv' into `$AllMGUsers: $($_.Exception.Message)"; $AllMGUsers = @() }

try { $BitlockerKeys = @(Import-Csv -Path (Join-Path $CsvFolderPath "BitlockerKeys.csv")) }
catch { Write-Log -Message "Failed to import 'BitlockerKeys.csv' into `$BitlockerKeys: $($_.Exception.Message)"; $BitlockerKeys = @() }

try { $Contacts = @(Import-Csv -Path (Join-Path $CsvFolderPath "Contacts.csv")) }
catch { Write-Log -Message "Failed to import 'Contacts.csv' into `$Contacts: $($_.Exception.Message)"; $Contacts = @() }

try { $CustomDomains = @(Import-Csv -Path (Join-Path $CsvFolderPath "CustomDomains.csv")) }
catch { Write-Log -Message "Failed to import 'CustomDomains.csv' into `$CustomDomains: $($_.Exception.Message)"; $CustomDomains = @() }

try { $DetailedLicenseSummary = @(Import-Csv -Path (Join-Path $CsvFolderPath "DetailedLicenseSummary.csv")) }
catch { Write-Log -Message "Failed to import 'DetailedLicenseSummary.csv' into `$DetailedLicenseSummary: $($_.Exception.Message)"; $DetailedLicenseSummary = @() }

try { $Devices = @(Import-Csv -Path (Join-Path $CsvFolderPath "Devices.csv")) }
catch { Write-Log -Message "Failed to import 'Devices.csv' into `$Devices: $($_.Exception.Message)"; $Devices = @() }

try { $DistributionGroups = @(Import-Csv -Path (Join-Path $CsvFolderPath "DistributionGroups.csv")) }
catch { Write-Log -Message "Failed to import 'DistributionGroups.csv' into `$DistributionGroups: $($_.Exception.Message)"; $DistributionGroups = @() }

try { $DKIMRecords = @(Import-Csv -Path (Join-Path $CsvFolderPath "DKIMRecords.csv")) }
catch { Write-Log -Message "Failed to import 'DKIMRecords.csv' into `$DKIMRecords: $($_.Exception.Message)"; $DKIMRecords = @() }

try { $DMARCRecords = @(Import-Csv -Path (Join-Path $CsvFolderPath "DMARCRecords.csv")) }
catch { Write-Log -Message "Failed to import 'DMARCRecords.csv' into `$DMARCRecords: $($_.Exception.Message)"; $DMARCRecords = @() }

try { $DynamicDistributionGroups = @(Import-Csv -Path (Join-Path $CsvFolderPath "DynamicDistributionGroups.csv")) }
catch { Write-Log -Message "Failed to import 'DynamicDistributionGroups.csv' into `$DynamicDistributionGroups: $($_.Exception.Message)"; $DynamicDistributionGroups = @() }

try { $GroupMailboxes = @(Import-Csv -Path (Join-Path $CsvFolderPath "GroupMailboxes.csv")) }
catch { Write-Log -Message "Failed to import 'GroupMailboxes.csv' into `$GroupMailboxes: $($_.Exception.Message)"; $GroupMailboxes = @() }

try { $InboundConnectors = @(Import-Csv -Path (Join-Path $CsvFolderPath "InboundConnectors.csv")) }
catch { Write-Log -Message "Failed to import 'InboundConnectors.csv' into `$InboundConnectors: $($_.Exception.Message)"; $InboundConnectors = @() }

try { $LicenseData = @(Import-Csv -Path (Join-Path $CsvFolderPath "LicenseData.csv")) }
catch { Write-Log -Message "Failed to import 'LicenseData.csv' into `$LicenseData: $($_.Exception.Message)"; $LicenseData = @() }

try { $LicenseOnePerLine = @(Import-Csv -Path (Join-Path $CsvFolderPath "LicenseOnePerLine.csv")) }
catch { Write-Log -Message "Failed to import 'LicenseOnePerLine.csv' into `$LicenseOnePerLine: $($_.Exception.Message)"; $LicenseOnePerLine = @() }

try { $MailEnabledSecurityGroups = @(Import-Csv -Path (Join-Path $CsvFolderPath "MailEnabledSecurityGroups.csv")) }
catch { Write-Log -Message "Failed to import 'MailEnabledSecurityGroups.csv' into `$MailEnabledSecurityGroups: $($_.Exception.Message)"; $MailEnabledSecurityGroups = @() }

try { $MXRecords = @(Import-Csv -Path (Join-Path $CsvFolderPath "MXRecords.csv")) }
catch { Write-Log -Message "Failed to import 'MXRecords.csv' into `$MXRecords: $($_.Exception.Message)"; $MXRecords = @() }

try { $OneDriveSites = @(Import-Csv -Path (Join-Path $CsvFolderPath "OneDriveSites.csv")) }
catch { Write-Log -Message "Failed to import 'OneDriveSites.csv' into `$OneDriveSites: $($_.Exception.Message)"; $OneDriveSites = @() }

try { $Org = @(Import-Csv -Path (Join-Path $CsvFolderPath "Org.csv")) }
catch { Write-Log -Message "Failed to import 'Org.csv' into `$Org: $($_.Exception.Message)"; $Org = @() }

try { $OutboundConnectors = @(Import-Csv -Path (Join-Path $CsvFolderPath "OutboundConnectors.csv")) }
catch { Write-Log -Message "Failed to import 'OutboundConnectors.csv' into `$OutboundConnectors: $($_.Exception.Message)"; $OutboundConnectors = @() }

try { $PublicFolders = @(Import-Csv -Path (Join-Path $CsvFolderPath "PublicFolders.csv")) }
catch { Write-Log -Message "Failed to import 'PublicFolders.csv' into `$PublicFolders: $($_.Exception.Message)"; $PublicFolders = @() }

try { $RegisteredApplications = @(Import-Csv -Path (Join-Path $CsvFolderPath "RegisteredApplications.csv")) }
catch { Write-Log -Message "Failed to import 'RegisteredApplications.csv' into `$RegisteredApplications: $($_.Exception.Message)"; $RegisteredApplications = @() }

try { $SecurityGroups = @(Import-Csv -Path (Join-Path $CsvFolderPath "SecurityGroups.csv")) }
catch { Write-Log -Message "Failed to import 'SecurityGroups.csv' into `$SecurityGroups: $($_.Exception.Message)"; $SecurityGroups = @() }

try { $SharePointSites = @(Import-Csv -Path (Join-Path $CsvFolderPath "SharePointSites.csv")) }
catch { Write-Log -Message "Failed to import 'SharePointSites.csv' into `$SharePointSites: $($_.Exception.Message)"; $SharePointSites = @() }

try { $SPFRecords = @(Import-Csv -Path (Join-Path $CsvFolderPath "SPFRecords.csv")) }
catch { Write-Log -Message "Failed to import 'SPFRecords.csv' into `$SPFRecords: $($_.Exception.Message)"; $SPFRecords = @() }

try { $TeamChannels = @(Import-Csv -Path (Join-Path $CsvFolderPath "TeamChannels.csv")) }
catch { Write-Log -Message "Failed to import 'TeamChannels.csv' into `$TeamChannels: $($_.Exception.Message)"; $TeamChannels = @() }

try { $Teams = @(Import-Csv -Path (Join-Path $CsvFolderPath "Teams.csv")) }
catch { Write-Log -Message "Failed to import 'Teams.csv' into `$Teams: $($_.Exception.Message)"; $Teams = @() }

try { $TeamsDialPlans = @(Import-Csv -Path (Join-Path $CsvFolderPath "TeamsDialPlans.csv")) }
catch { Write-Log -Message "Failed to import 'TeamsDialPlans.csv' into `$TeamsDialPlans: $($_.Exception.Message)"; $TeamsDialPlans = @() }

try { $TeamsPstnUsage = @(Import-Csv -Path (Join-Path $CsvFolderPath "TeamsPstnUsage.csv")) }
catch { Write-Log -Message "Failed to import 'TeamsPstnUsage.csv' into `$TeamsPstnUsage: $($_.Exception.Message)"; $TeamsPstnUsage = @() }

try { $TeamsUserActivity_Last180Days = @(Import-Csv -Path (Join-Path $CsvFolderPath "TeamsUserActivity_Last180Days.csv")) }
catch { Write-Log -Message "Failed to import 'TeamsUserActivity_Last180Days.csv' into `$TeamsUserActivity_Last180Days: $($_.Exception.Message)"; $TeamsUserActivity_Last180Days = @() }

try { $TeamsVoiceRoutingPolicies = @(Import-Csv -Path (Join-Path $CsvFolderPath "TeamsVoiceRoutingPolicies.csv")) }
catch { Write-Log -Message "Failed to import 'TeamsVoiceRoutingPolicies.csv' into `$TeamsVoiceRoutingPolicies: $($_.Exception.Message)"; $TeamsVoiceRoutingPolicies = @() }

try { $TransportRules = @(Import-Csv -Path (Join-Path $CsvFolderPath "TransportRules.csv")) }
catch { Write-Log -Message "Failed to import 'TransportRules.csv' into `$TransportRules: $($_.Exception.Message)"; $TransportRules = @() }

try { $UnifiedGroups = @(Import-Csv -Path (Join-Path $CsvFolderPath "UnifiedGroups.csv")) }
catch { Write-Log -Message "Failed to import 'UnifiedGroups.csv' into `$UnifiedGroups: $($_.Exception.Message)"; $UnifiedGroups = @() }

# Record counts only - deliberately not displaying any of the imported
# data itself. Built by reading back the variable named after each
# expected file, so this stays in sync automatically with the explicit
# imports above as long as the variable names match the file names.
$importSummary = foreach ($expected in $ExpectedFiles) {
    $varName = [System.IO.Path]::GetFileNameWithoutExtension($expected.FileName)
    $varValue = Get-Variable -Name $varName -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    [PSCustomObject]@{
        VariableName = "`$$varName"
        RecordCount  = if ($null -eq $varValue) { "(not imported)" } else { @($varValue).Count }
    }
}
$importSummary | Format-Table -AutoSize
Write-Host "Imported $($ExpectedFiles.Count) CSV files into variables." -ForegroundColor Green
#endregion

#region Step 5: name and copy the report from the template
Write-Host "`nStep 5 of 20: preparing the report file..." -ForegroundColor Cyan

# A new report file name is always generated from the tenant's
# DisplayName + today's date - there's no override parameter, so this
# can't accidentally reuse or overwrite a previous run's name.
if (-not $Org -or [string]::IsNullOrWhiteSpace($Org[0].DisplayName)) {
    Write-Log -Level Error -Message "Couldn't read a DisplayName value from the imported `$Org variable (from Org.csv). Cannot build the report file name without it."
    throw "DisplayName not found in `$Org."
}
$orgDisplayName = $Org[0].DisplayName

# Strip characters that aren't safe in a Windows file name, plus
# control characters, then collapse whitespace and trim trailing
# dots/spaces (also invalid at the end of a Windows file name) - so
# an org name like 'Contoso, Ltd. / EMEA' doesn't break Copy-Item.
$safeDisplayName = $orgDisplayName -replace '[\\/:\*\?"<>\|]', ''
$safeDisplayName = $safeDisplayName -replace '[\x00-\x1F]', ''
$safeDisplayName = ($safeDisplayName -replace '\s+', ' ').Trim()
$safeDisplayName = $safeDisplayName.TrimEnd('.', ' ')

if ([string]::IsNullOrWhiteSpace($safeDisplayName)) {
    Write-Log -Message "DisplayName '$orgDisplayName' from `$Org became empty after removing unsafe characters - falling back to 'UnknownOrg' for the file name."
    $safeDisplayName = "UnknownOrg"
}

$ReportFileName = "$safeDisplayName.Softcat-T2T-Assessment-Report-$(Get-Date -Format 'yyyyMMdd-HHmmss').xlsx"

$reportPath = Join-Path $OutputPath $ReportFileName
Write-Host "Copying template to: $reportPath"
Copy-Item -Path $TemplatePath -Destination $reportPath -Force
Write-Host "Report file created from template." -ForegroundColor Green

# CSV-to-worksheet import (using the remaining $-prefixed variables
# from Step 4) comes here in a later revision. $ReportData (Step 6) and
# $AccountSummary (Step 7) are now written to this report as their own
# worksheets.
#endregion

#region Step 6: resolve report metadata (written to the report at the end, see below)
Write-Host "`nStep 6 of 20: compiling and writing report data..." -ForegroundColor Cyan

# WARNING: assuming Org.csv exports the tenant ID under a column named
# 'Id' - that's the standard property Get-MgOrganization returns for a
# tenant's ID, but hasn't been confirmed against the real CSV.
$tenantId = if ($Org -and $Org[0].Id) { $Org[0].Id } else { "" }
if ([string]::IsNullOrWhiteSpace($tenantId)) {
    Write-Log -Message "Couldn't read an Id value (tenant ID) from the imported `$Org variable (from Org.csv). TenantId will be blank in the report data."
}

# SharePoint Admin URL isn't one of the CSVs imported in Step 4 - it's
# exported separately as a plain text file (AdminURL.txt) by
# 2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1.
$adminUrlPath = Join-Path $CsvFolderPath "AdminURL.txt"
$sharePointAdminUrl = if (Test-Path $adminUrlPath) {
    (Get-Content -Path $adminUrlPath -Raw).Trim()
}
else {
    Write-Log -Message "AdminURL.txt not found in '$CsvFolderPath'. SharePoint Admin URL will be blank in the report data."
    ""
}
# The actual write to the report's "Data" worksheet is deliberately
# deferred to the end of the script (after Steps 7-17), so it's the
# last worksheet written - see "Step 6 (continued)" below.
#endregion

#region Step 7: build a combined mailbox + OneDrive + license view and write it to the report
Write-Host "`nStep 7 of 20: compiling mailbox, OneDrive and license data..." -ForegroundColor Cyan

$AccountSummary = foreach ($mailbox in $AccountMailboxes) {

    # Match the OneDrive site for this user - OneDriveSites.csv's
    # 'Owner' column containing the UPN was confirmed against a sample
    # export, so no WARNING caveat here.
    $oneDrive = $OneDriveSites | Where-Object { $_.Owner -eq $mailbox.UserPrincipalName }

    # Match all licenses for this user - LicenseData.csv's
    # UserPrincipalName and Licenses columns were confirmed against a
    # sample export (Licenses already comes pre-flattened to a single
    # ";"-separated string per user), so no WARNING caveat here either.
    $userLicenses = $LicenseData  | Where-Object { $_.UserPrincipalName -eq $mailbox.UserPrincipalName }

    [PSCustomObject]@{
        Type                = $mailbox.Type
        DisplayName         = $mailbox.DisplayName
        UserPrincipalName   = $mailbox.UserPrincipalName
        PrimarySmtpAddress  = $mailbox.PrimarySmtpAddress
        Items               = $mailbox.ItemCount
        SizeMB              = $mailbox.TotalItemSizeMB
        DeletedSizeMB       = $mailbox.TotalDeletedItemSizeMB
        ArchiveSizeMB       = $mailbox.ArchiveTotalItemSizeMB
        OneDriveMB          = $oneDrive.StorageUsageCurrentMB
        Licenses            = $userLicenses.Licenses
        LitigationHold      = $mailbox.LitigationHoldEnabled
    }
}

#$AccountSummary | Format-Table -AutoSize
Write-Host "Compiled $($AccountSummary.Count) user record(s) into `$AccountSummary." -ForegroundColor Green

# Write $AccountSummary into the report workbook created in Step 5, as
# its own worksheet, formatted as an Excel Table (not just raw cells) -
# -AutoSize/-FreezeTopRow/-BoldTopRow are redundant with -TableStyle
# but kept for consistent header behavior if the table is later removed.
# Export-Excel adds/replaces just this worksheet when -Path already
# exists, leaving the template's other sheets alone.
try {
    Start-Rest -NextItem "the AccountSummary worksheet"
    $AccountSummary | Export-Excel -Path $reportPath -WorksheetName "AccountSummary" -TableName "AccountSummary" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($AccountSummary.Count) row(s) to the 'AccountSummary' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$AccountSummary to the 'AccountSummary' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the AccountSummary worksheet - see the log for details."
}
#endregion

#region Step 8: build a combined group + mailbox + Teams view and write it to the report
Write-Host "`nStep 8 of 20: compiling group, mailbox and Teams data..." -ForegroundColor Cyan

$GroupSummary = foreach ($group in $UnifiedGroups) {

    # Match this group's mailbox stats from GroupMailboxes.csv, by
    # Guid - named $groupMailbox (not $mailbox) so it can't collide
    # with the $mailbox loop variable from Step 7's $AccountSummary
    # foreach above, which would otherwise still be sitting in scope
    # (PowerShell foreach loops don't get their own scope) and leak
    # its last value from Step 7 into every row here.
    $groupMailbox = $GroupMailboxes | Where-Object { $_.SPOSiteURL -eq $group.SharePointSiteURL }

    # Match the Teams record for this group, if it's also a Team - not
    # every M365 Group is a Team, so this can come back empty.
    $team = $Teams | Where-Object { $_.GroupId -eq $group.GroupID }

    # Match the group's SharePoint site for its storage usage and
    # template - a Template starting with "TEAMCHANNEL" means the site
    # actually belongs to a private/shared Teams channel rather than
    # the group's main Team.
    $site = $SharePointSites | Where-Object { $_.GroupId -eq $group.GroupID }

    if ($site.Template -like 'TEAMCHANNEL*') {
        $groupType = 'Teams Channel'
    }
    elseif ($team) {
        $groupType = 'Team'
    }
    else {
        $groupType = 'M365 Group'
    }

    [PSCustomObject]@{
        Type                    = $groupType
        DisplayName             = $group.DisplayName
        PrimarySmtpAddress      = $group.PrimarySmtpAddress
        SPOStorageMB            = $site.StorageUsageCurrentMB
        Items                   = $groupMailbox.ItemCount
        SizeMB                  = $groupMailbox.TotalItemSizeMB
        AccessType              = $group.AccessType
        Members                 = $group.Members
        Guests                  = $group.Guests
        SharePointSiteURL       = $group.SharePointSiteURL
        Visibility              = $team.Visibility
        Archived                = $team.Archived
        TeamChannels            = $team.Channels
    }
}

#$GroupSummary | Format-Table -AutoSize
Write-Host "Compiled $($GroupSummary.Count) group record(s) into `$GroupSummary." -ForegroundColor Green

# Write $GroupSummary into the report workbook, as its own worksheet,
# formatted as an Excel Table - same pattern as AccountSummary above.
# Export-Excel adds/replaces just this worksheet when -Path already
# exists, leaving the template's other sheets alone.
try {
    Start-Rest -NextItem "the GroupSummary worksheet"
    $GroupSummary | Export-Excel -Path $reportPath -WorksheetName "GroupSummary" -TableName "GroupSummary" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($GroupSummary.Count) row(s) to the 'GroupSummary' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$GroupSummary to the 'GroupSummary' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the GroupSummary worksheet - see the log for details."
}
#endregion

#region Step 9: build a SharePoint sites view with a Type column and write it to the report
Write-Host "`nStep 9 of 20: compiling SharePoint sites data..." -ForegroundColor Cyan

$AllZeroGroupId = '00000000-0000-0000-0000-000000000000'

$SharePointSiteSummary = foreach ($site in $SharePointSites) {

    # Match the Teams record for this site's group, if any - only
    # meaningful for sites with a real GroupId (Teams Channel and
    # dedicated sites carry the all-zero GroupId, so this comes back
    # empty for those).
    $team = $Teams | Where-Object { $_.GroupId -eq $site.GroupId }

    if ($site.Template -like 'TEAMCHANNEL*') {
        $siteType = 'Teams Channel'
    }
    elseif ($site.GroupId -eq $AllZeroGroupId) {
        $siteType = 'SharePoint Site'
    }
    elseif ($team) {
        $siteType = 'Team'
    }
    else {
        $siteType = 'M365 Group'
    }

    [PSCustomObject]@{

        Type                    = $siteType
        Title                   = $site.Title
        Template                = $site.Template
        StorageMB               = $site.StorageUsageCurrentMB
        QuotaMB                 = $site.StorageQuotaMB
        LastModifiedDate        = $site.LastContentModifiedDate
        Status                  = $site.Status
        Sharing                 = $site.SharingCapability
        Owner                   = $site.Owner
        Url                     = $site.Url
        GroupId                 = $site.GroupId
    }
}

#$SharePointSiteSummary | Format-Table -AutoSize
Write-Host "Compiled $($SharePointSiteSummary.Count) SharePoint site record(s) into `$SharePointSiteSummary." -ForegroundColor Green

# Write $SharePointSiteSummary into the report workbook, as its own
# worksheet, formatted as an Excel Table - same pattern as
# AccountSummary/GroupSummary above.
try {
    Start-Rest -NextItem "the SharePointSites worksheet"
    $SharePointSiteSummary | Export-Excel -Path $reportPath -WorksheetName "SharePointSites" -TableName "SharePointSites" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($SharePointSiteSummary.Count) row(s) to the 'SharePointSites' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$SharePointSiteSummary to the 'SharePointSites' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the SharePointSites worksheet - see the log for details."
}
#endregion

#region Step 10: build a deduplicated device inventory view with an ActiveLast30Days column and write it to the report
Write-Host "`nStep 10 of 20: compiling device data..." -ForegroundColor Cyan

# ApproximateLastSignInDateTime is exported as dd/MM/yyyy HH:mm:ss (UK
# format) - parsed explicitly with that format rather than relying on
# default date parsing, which can misread day/month depending on the
# machine's locale.
#
# Devices.csv can carry multiple rows for the same DisplayName (e.g. a
# device re-registered, or seen via more than one join type) - the
# report needs DisplayName to be unique, so each device is parsed
# first, then grouped by DisplayName and only the row with the latest
# ApproximateLastSignInDateTime is kept per group. A device with no
# parseable sign-in date is only kept if every row sharing that
# DisplayName is also undated - it always loses to a dated row from
# the same DisplayName.
$parsedDevices = foreach ($device in $Devices) {

    $lastSignIn = $null
    if (-not [string]::IsNullOrWhiteSpace($device.ApproximateLastSignInDateTime)) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParseExact($device.ApproximateLastSignInDateTime, 'dd/MM/yyyy HH:mm:ss', $null, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            $lastSignIn = $parsed
        }
        else {
            Write-Log -Message "Couldn't parse ApproximateLastSignInDateTime '$($device.ApproximateLastSignInDateTime)' for device '$($device.DisplayName)' as dd/MM/yyyy HH:mm:ss - ActiveLast30Days will be blank for this device."
        }
    }

    [PSCustomObject]@{
        Device     = $device
        LastSignIn = $lastSignIn
    }
}

$duplicateDisplayNameCount = ($parsedDevices | Group-Object { $_.Device.DisplayName } | Where-Object Count -gt 1).Count
if ($duplicateDisplayNameCount -gt 0) {
    Write-Log -Message "$duplicateDisplayNameCount DisplayName(s) in Devices.csv had more than one row - kept only the row with the latest ApproximateLastSignInDateTime for each."
}

$DeviceSummary = foreach ($group in ($parsedDevices | Group-Object { $_.Device.DisplayName })) {

    # Sort undated rows ([datetime]::MinValue stand-in) below any dated
    # row in the same group, then take the most recent.
    $latest = $group.Group | Sort-Object -Property @{ Expression = { if ($null -eq $_.LastSignIn) { [datetime]::MinValue } else { $_.LastSignIn } }; Descending = $true } | Select-Object -First 1
    $device = $latest.Device
    $lastSignIn = $latest.LastSignIn

    # TrustType comes through as Graph's raw join-type values - mapped
    # here to the friendlier Entra terms; switch is case-insensitive by
    # default, and anything unrecognised (including blank) passes
    # through unchanged rather than being blanked out.
    $trustType = switch ($device.TrustType) {
        'Workplace' { 'Entra Registered' }
        'ServerAd'  { 'Entra Hybrid Joined' }
        'AzureAd'   { 'Entra Joined' }
        default     { $device.TrustType }
    }

    [PSCustomObject]@{
        DisplayName                   = $device.DisplayName
        OperatingSystem               = $device.OperatingSystem
        OSVersion                     = $device.OperatingSystemVersion
        TrustType                     = $trustType
        SyncEnabled                   = $device.OnPremisesSyncEnabled
        ProfileType                   = $device.ProfileType
        RegisteredOwners              = $device.RegisteredOwners
        AccountEnabled                = $device.AccountEnabled
        ApproximateLastSignInDateTime = $device.ApproximateLastSignInDateTime
        ActiveLast30Days              = if ($null -eq $lastSignIn) { $null } else { $lastSignIn -ge $reportCreationDate.AddDays(-30) }
    }
}

#$DeviceSummary | Format-Table -AutoSize
Write-Host "Compiled $($DeviceSummary.Count) unique device record(s) into `$DeviceSummary (from $($Devices.Count) row(s) in Devices.csv)." -ForegroundColor Green

# Write $DeviceSummary into the report workbook, as its own worksheet,
# formatted as an Excel Table - same pattern as the summaries above.
try {
    Start-Rest -NextItem "the Devices worksheet"
    $DeviceSummary | Export-Excel -Path $reportPath -WorksheetName "Devices" -TableName "Devices" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($DeviceSummary.Count) row(s) to the 'Devices' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$DeviceSummary to the 'Devices' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the Devices worksheet - see the log for details."
}
#endregion

#region Step 11: build a OneDrive sites view and write it to the report
Write-Host "`nStep 11 of 20: compiling OneDrive sites data..." -ForegroundColor Cyan

# Every row in OneDriveSites.csv is a personal site (Template
# "SPSPERS#12") with an all-zero GroupId - confirmed against a sample
# export - so unlike SharePointSites there's no Team/M365 Group/Teams
# Channel distinction to derive here.
$OneDriveSiteSummary = foreach ($site in $OneDriveSites) {
    [PSCustomObject]@{
        User             = $site.Title
        Owner            = $site.Owner
        StorageMB        = $site.StorageUsageCurrentMB
        QuotaMB          = $site.StorageQuotaMB
        LastModifiedDate = $site.LastContentModifiedDate
        Status           = $site.Status
        Sharing          = $site.SharingCapability
        Url              = $site.Url
    }
}

#$OneDriveSiteSummary | Format-Table -AutoSize
Write-Host "Compiled $($OneDriveSiteSummary.Count) OneDrive site record(s) into `$OneDriveSiteSummary." -ForegroundColor Green

# Write $OneDriveSiteSummary into the report workbook, as its own
# worksheet, formatted as an Excel Table - same pattern as the
# summaries above.
try {
    Start-Rest -NextItem "the OneDrive worksheet"
    $OneDriveSiteSummary | Export-Excel -Path $reportPath -WorksheetName "OneDrives" -TableName "OneDrives" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($OneDriveSiteSummary.Count) row(s) to the 'OneDrive' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$OneDriveSiteSummary to the 'OneDrive' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the OneDrive worksheet - see the log for details."
}
#endregion

#region Step 12: build an all-users view with a looked-up Licenses column and write it to the report
Write-Host "`nStep 12 of 20: compiling all-users data..." -ForegroundColor Cyan

$AllUsersSummary = foreach ($user in $AllMGUsers) {

    # Look up this user's licenses from LicenseData.csv, by
    # UserPrincipalName - replaces the LicenseStatus column (which
    # just carries "Active"/"Active;Active" noise) with the actual
    # license names.
    $userLicenses = $LicenseData | Where-Object { $_.UserPrincipalName -eq $user.UserPrincipalName }

    [PSCustomObject]@{
        ID                = $user.ID
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        Mail              = $user.Mail
        UserType          = $user.UserType
        UsageLocation     = $user.UsageLocation
        AccountEnabled    = $user.AccountEnabled
        IsDirSynced       = $user.IsDirSynced
        Licenses          = $userLicenses.Licenses
    }
}

#$AllUsersSummary | Format-Table -AutoSize
Write-Host "Compiled $($AllUsersSummary.Count) user record(s) into `$AllUsersSummary." -ForegroundColor Green

# Write $AllUsersSummary into the report workbook, as its own
# worksheet, formatted as an Excel Table - same pattern as the
# summaries above.
try {
    Start-Rest -NextItem "the AllUsers worksheet"
    $AllUsersSummary | Export-Excel -Path $reportPath -WorksheetName "AllUsers" -TableName "AllUsers" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($AllUsersSummary.Count) row(s) to the 'AllUsers' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$AllUsersSummary to the 'AllUsers' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the AllUsers worksheet - see the log for details."
}
#endregion

#region Step 13: build a contacts view and write it to the report
Write-Host "`nStep 13 of 20: compiling contacts data..." -ForegroundColor Cyan

$ContactSummary = foreach ($contact in $Contacts) {
    [PSCustomObject]@{
        Name                 = $contact.Name
        DisplayName          = $contact.DisplayName
        Alias                = $contact.Alias
        PrimarySmtpAddress   = $contact.PrimarySmtpAddress
        RecipientTypeDetails = $contact.RecipientTypeDetails
        Identity             = $contact.Identity
    }
}

#$ContactSummary | Format-Table -AutoSize
Write-Host "Compiled $($ContactSummary.Count) contact record(s) into `$ContactSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the Contacts worksheet"
    $ContactSummary | Export-Excel -Path $reportPath -WorksheetName "Contacts" -TableName "Contacts" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($ContactSummary.Count) row(s) to the 'Contacts' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$ContactSummary to the 'Contacts' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the Contacts worksheet - see the log for details."
}
#endregion

#region Step 14: build a mail-enabled security groups view and write it to the report
Write-Host "`nStep 14 of 20: compiling mail-enabled security groups data..." -ForegroundColor Cyan

$MailEnabledSecurityGroupSummary = foreach ($mesg in $MailEnabledSecurityGroups) {
    [PSCustomObject]@{
        Name                          = $mesg.Name
        DisplayName                   = $mesg.DisplayName
        Alias                         = $mesg.Alias
        PrimarySmtpAddress            = $mesg.PrimarySmtpAddress
        ManagedBy                     = $mesg.ManagedBy
        OrganizationalUnit            = $mesg.OrganizationalUnit
        RequireSenderAuthentication   = $mesg.RequireSenderAuthentication
        HiddenFromAddressLists        = $mesg.HiddenFromAddressListsEnabled
        MemberCount                   = $mesg.MemberCount
        RecipientTypeDetails          = $mesg.RecipientTypeDetails
        Identity                      = $mesg.Identity
    }
}

#$MailEnabledSecurityGroupSummary | Format-Table -AutoSize
Write-Host "Compiled $($MailEnabledSecurityGroupSummary.Count) mail-enabled security group record(s) into `$MailEnabledSecurityGroupSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the MailEnabledSecurityGroups worksheet"
    $MailEnabledSecurityGroupSummary | Export-Excel -Path $reportPath -WorksheetName "MailEnabledSecurityGroups" -TableName "MailEnabledSecurityGroups" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($MailEnabledSecurityGroupSummary.Count) row(s) to the 'MailEnabledSecurityGroups' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$MailEnabledSecurityGroupSummary to the 'MailEnabledSecurityGroups' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the MailEnabledSecurityGroups worksheet - see the log for details."
}
#endregion

#region Step 15: build a dynamic distribution groups view and write it to the report
Write-Host "`nStep 15 of 20: compiling dynamic distribution groups data..." -ForegroundColor Cyan

$DynamicDistributionGroupSummary = foreach ($ddg in $DynamicDistributionGroups) {
    [PSCustomObject]@{
        Name                 = $ddg.Name
        DisplayName          = $ddg.DisplayName
        RecipientFilterType  = $ddg.RecipientFilterType
        RecipientTypeDetails = $ddg.RecipientTypeDetails
        PrimarySmtpAddress   = $ddg.PrimarySmtpAddress
        ManagedBy            = $ddg.ManagedBy
    }
}

#$DynamicDistributionGroupSummary | Format-Table -AutoSize
Write-Host "Compiled $($DynamicDistributionGroupSummary.Count) dynamic distribution group record(s) into `$DynamicDistributionGroupSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the DynamicDistributionGroups worksheet"
    $DynamicDistributionGroupSummary | Export-Excel -Path $reportPath -WorksheetName "DynamicDistributionGroups" -TableName "DynamicDistributionGroups" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($DynamicDistributionGroupSummary.Count) row(s) to the 'DynamicDistributionGroups' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$DynamicDistributionGroupSummary to the 'DynamicDistributionGroups' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the DynamicDistributionGroups worksheet - see the log for details."
}
#endregion

#region Step 16: build a distribution groups view and write it to the report
Write-Host "`nStep 16 of 20: compiling distribution groups data..." -ForegroundColor Cyan

# DistributionGroups.csv carries 38 raw Exchange properties (moderation
# settings, routing restrictions, etc.) - most aren't useful in an
# assessment report, so this keeps a curated subset rather than every
# column. Ask if any of the dropped columns (e.g. ModerationEnabled,
# MaxSendSize/MaxReceiveSize) turn out to matter for the report.
$DistributionGroupSummary = foreach ($dg in $DistributionGroups) {
    [PSCustomObject]@{
        Name                    = $dg.Name
        DisplayName             = $dg.DisplayName
        PrimarySmtpAddress      = $dg.PrimarySmtpAddress
        GroupType               = $dg.GroupType
        ManagedBy               = $dg.ManagedBy
        MemberCount             = $dg.MemberCount
        HiddenFromAddressLists  = $dg.HiddenFromAddressListsEnabled
        IsDirSynced             = $dg.IsDirSynced
        Description             = $dg.Description
        WhenCreated             = $dg.WhenCreated
        WhenChanged             = $dg.WhenChanged
    }
}

#$DistributionGroupSummary | Format-Table -AutoSize
Write-Host "Compiled $($DistributionGroupSummary.Count) distribution group record(s) into `$DistributionGroupSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the DistributionGroups worksheet"
    $DistributionGroupSummary | Export-Excel -Path $reportPath -WorksheetName "DistributionGroups" -TableName "DistributionGroups" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($DistributionGroupSummary.Count) row(s) to the 'DistributionGroups' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$DistributionGroupSummary to the 'DistributionGroups' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the DistributionGroups worksheet - see the log for details."
}
#endregion

#region Step 17: build a security groups view and write it to the report
Write-Host "`nStep 17 of 20: compiling security groups data..." -ForegroundColor Cyan

$SecurityGroupSummary = foreach ($sg in $SecurityGroups) {
    [PSCustomObject]@{
        DisplayName     = $sg.DisplayName
        Description     = $sg.Description
        Mail            = $sg.Mail
        MailEnabled     = $sg.MailEnabled
        SecurityEnabled = $sg.SecurityEnabled
        GroupTypes      = $sg.GroupTypes
        CreatedDateTime = $sg.CreatedDateTime
        MemberCount     = $sg.MemberCount
        Owners          = $sg.Owners
        Id              = $sg.Id
    }
}

#$SecurityGroupSummary | Format-Table -AutoSize
Write-Host "Compiled $($SecurityGroupSummary.Count) security group record(s) into `$SecurityGroupSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the SecurityGroups worksheet"
    $SecurityGroupSummary | Export-Excel -Path $reportPath -WorksheetName "SecurityGroups" -TableName "SecurityGroups" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($SecurityGroupSummary.Count) row(s) to the 'SecurityGroups' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$SecurityGroupSummary to the 'SecurityGroups' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the SecurityGroups worksheet - see the log for details."
}
#endregion

#region Step 18: build an account mailboxes view and write it to the report
Write-Host "`nStep 18 of 20: compiling account mailboxes data..." -ForegroundColor Cyan

# AccountMailboxes.csv carries 28 raw Exchange properties - most of the
# archive sub-fields are blank unless Archive is enabled, and several
# (Alias, EmailAddresses, Guid, LitigationHoldDuration, InPlaceHolds,
# RetentionHoldEnabled, StartDateForRetentionHold,
# EndDateForRetentionHold) aren't useful in an assessment report, so
# this keeps a curated subset rather than every column. Ask if any of
# the dropped columns turn out to matter for the report.
$AccountMailboxSummary = foreach ($am in $AccountMailboxes) {
    [PSCustomObject]@{
        Type                    = $am.Type
        DisplayName             = $am.DisplayName
        UserPrincipalName       = $am.UserPrincipalName
        PrimarySmtpAddress      = $am.PrimarySmtpAddress
        ItemCount               = $am.ItemCount
        TotalItemSizeMB         = $am.TotalItemSizeMB
        DeletedItemCount        = $am.DeletedItemCount
        TotalDeletedItemSizeMB  = $am.TotalDeletedItemSizeMB
        Archive                 = $am.Archive
        ArchiveItemCount        = $am.ArchiveItemCount
        ArchiveTotalItemSizeMB  = $am.ArchiveTotalItemSizeMB
        LitigationHoldEnabled   = $am.LitigationHoldEnabled
        RetentionPolicy         = $am.RetentionPolicy
        AccountSKU              = $am.AccountSKU
        WhenCreatedUTC          = $am.WhenCreatedUTC
        WhenChangedUTC          = $am.WhenChangedUTC
        LastLogonTime           = $am.LastLogonTime
    }
}

#$AccountMailboxSummary | Format-Table -AutoSize
Write-Host "Compiled $($AccountMailboxSummary.Count) account mailbox record(s) into `$AccountMailboxSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the AccountMailboxes worksheet"
    $AccountMailboxSummary | Export-Excel -Path $reportPath -WorksheetName "AccountMailboxes" -TableName "AccountMailboxes" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($AccountMailboxSummary.Count) row(s) to the 'AccountMailboxes' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$AccountMailboxSummary to the 'AccountMailboxes' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the AccountMailboxes worksheet - see the log for details."
}
#endregion

#region Step 19: build a group mailboxes view and write it to the report
Write-Host "`nStep 19 of 20: compiling group mailboxes data..." -ForegroundColor Cyan

$GroupMailboxSummary = foreach ($gm in $GroupMailboxes) {
    [PSCustomObject]@{
        DisplayName            = $gm.DisplayName
        Alias                  = $gm.Alias
        PrimarySmtpAddress     = $gm.PrimarySmtpAddress
        RecipientTypeDetails   = $gm.RecipientTypeDetails
        ItemCount              = $gm.ItemCount
        TotalItemSizeMB        = $gm.TotalItemSizeMB
        DeletedItemCount       = $gm.DeletedItemCount
        TotalDeletedItemSizeMB = $gm.TotalDeletedItemSizeMB
        WhenCreatedUTC         = $gm.WhenCreatedUTC
        WhenChangedUTC         = $gm.WhenChangedUTC
        LastLogonTime          = $gm.LastLogonTime
        EmailAddresses         = $gm.EmailAddresses
        SPOSiteURL             = $gm.SPOSiteURL
        Guid                   = $gm.Guid
    }
}

#$GroupMailboxSummary | Format-Table -AutoSize
Write-Host "Compiled $($GroupMailboxSummary.Count) group mailbox record(s) into `$GroupMailboxSummary." -ForegroundColor Green

try {
    Start-Rest -NextItem "the GroupMailboxes worksheet"
    $GroupMailboxSummary | Export-Excel -Path $reportPath -WorksheetName "GroupMailboxes" -TableName "GroupMailboxes" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($GroupMailboxSummary.Count) row(s) to the 'GroupMailboxes' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$GroupMailboxSummary to the 'GroupMailboxes' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the GroupMailboxes worksheet - see the log for details."
}
#endregion

#region Step 20: build a combined domains view and write it to the report
Write-Host "`nStep 20 of 20: compiling domain data..." -ForegroundColor Cyan

# Pulls the policy value (p=...) out of a DMARC TXT record string,
# e.g. "v=DMARC1; p=reject; rua=mailto:..." -> "reject".
$dmarcPolicyRegex = '(?i)(?:^|;)\s*p=([^;]+)'

$DomainSummary = foreach ($domain in $AcceptedDomains) {

    $domainName = $domain.DomainName

    # A domain is "custom" if it also shows up in CustomDomains.csv -
    # AcceptedDomains.csv includes the *.onmicrosoft.com domains too,
    # which CustomDomains.csv excludes. Confirmed against a sample
    # export: CustomDomains.csv's rows are a subset of
    # AcceptedDomains.csv's.
    $isCustomDomain = [bool]($CustomDomains | Where-Object { $_.DomainName -eq $domainName })

    # DKIMRecords.csv - one row per domain (or none, if DKIM was never
    # set up for that domain).
    $dkim = $DKIMRecords | Where-Object { $_.Domain -eq $domainName }

    # DMARCRecords.csv - one row per domain, but String comes back
    # blank for domains with no DMARC record published - confirmed
    # against a sample export.
    $dmarc = $DMARCRecords | Where-Object { $_.Domain -eq $domainName }
    $hasDmarc = [bool]($dmarc -and -not [string]::IsNullOrWhiteSpace($dmarc.String))
    $dmarcPolicy = $null
    if ($hasDmarc) {
        $match = [regex]::Match($dmarc.String, $dmarcPolicyRegex)
        if ($match.Success) { $dmarcPolicy = $match.Groups[1].Value.Trim() }
    }

    # A domain can legitimately have more than one MX record (multiple
    # mail exchangers at different preference values) - that's normal,
    # so they're just concatenated in preference order for reference,
    # not flagged as an issue.
    $mxRecords = @($MXRecords | Where-Object { $_.Name -eq $domainName } | Sort-Object { [int]$_.Preference })
    $mxSummary = ($mxRecords | ForEach-Object { "$($_.NameExchange) (pref $($_.Preference))" }) -join '; '

    # Unlike MX, a domain having more than one SPF TXT record IS a
    # misconfiguration (RFC 7208 - a domain should publish exactly one
    # SPF record; more than one causes lookups to permerror), so this
    # is flagged explicitly (MultipleSPFRecords) rather than silently
    # concatenated away like MX is. Seen in a sample export.
    $spfRecords = @($SPFRecords | Where-Object { $_.Name -eq $domainName })
    $spfSummary = ($spfRecords.String) -join '; '

    [PSCustomObject]@{
        DomainName         = $domainName
        DomainType         = $domain.DomainType
        IsDefault          = $domain.Default
        IsCustomDomain     = $isCustomDomain
        MatchSubDomains    = $domain.MatchSubDomains
        AuthenticationType = $domain.AuthenticationType
        DKIMEnabled        = if ($dkim) { $dkim.Enabled } else { $null }
        DKIMSelector1CNAME = $dkim.Selector1CNAME
        HasDMARC           = $hasDmarc
        DMARCPolicy        = $dmarcPolicy
        DMARCRecord        = $dmarc.String
        MXRecordCount      = $mxRecords.Count
        MXRecords          = $mxSummary
        SPFRecordCount     = $spfRecords.Count
        MultipleSPFRecords = $spfRecords.Count -gt 1
        SPFRecords         = $spfSummary
    }
}

#$DomainSummary | Format-Table -AutoSize
Write-Host "Compiled $($DomainSummary.Count) domain record(s) into `$DomainSummary." -ForegroundColor Green

# Write $DomainSummary into the report workbook, as its own worksheet,
# formatted as an Excel Table - same pattern as the summaries above.
try {
    Start-Rest -NextItem "the Domains worksheet"
    $DomainSummary | Export-Excel -Path $reportPath -WorksheetName "Domains" -TableName "Domains" -TableStyle Medium2 -AutoSize -FreezeTopRow -BoldTopRow
    Write-Host "Wrote $($DomainSummary.Count) row(s) to the 'Domains' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write `$DomainSummary to the 'Domains' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write the Domains worksheet - see the log for details."
}
#endregion

#region Step 6 (continued): write report metadata to the Data worksheet
Write-Host "`nStep 6 (continued): writing report data - deferred until now so it's the last worksheet written..." -ForegroundColor Cyan

# WARNING: writes directly into cells on a worksheet named "Data" that
# must already exist in template.xlsx (with row 3 formatted for its
# labels) - unlike the AccountSummary worksheet below, this doesn't
# create the worksheet itself. If template.xlsx doesn't have a "Data"
# worksheet, this throws rather than writing silently to the wrong
# place.
Start-Rest -NextItem "report data"
try {
    $excel = Open-ExcelPackage -Path $reportPath
    $worksheet = $excel.Workbook.Worksheets['Data']
    if (-not $worksheet) {
        throw "No worksheet named 'Data' was found in '$reportPath' - check template.xlsx has one."
    }

    $worksheet.Cells['C3'].Value = $orgDisplayName
    $worksheet.Cells['C4'].Value = $tenantId
    $worksheet.Cells['C5'].Value = $sharePointAdminUrl
    $worksheet.Cells['C6'].Value = $ConsultantName
    $worksheet.Cells['C7'].Value = $reportCreationDate.ToString('yyyy-MM-dd HH:mm:ss')
    $worksheet.Cells['C8'].Value = $ReportFileName

    # Count of enabled Member-type users, via a formula (not a static
    # value) against the AllUsers table written in Step 12, so it stays
    # live if that table is edited later.
    $worksheet.Cells['D12'].Formula = 'SUMPRODUCT((AllUsers[UserType]="Member")*(AllUsers[AccountEnabled]="True"))'
    $worksheet.Cells['E12'].Formula = 'COUNTIFS(AllUsers[UserType],"Member",AllUsers[Licenses],"<>")'
    $worksheet.Cells['F12'].Formula = 'SUMPRODUCT((AllUsers[UserType]="Member")*(AllUsers[IsDirSynced]="True"))'
    $worksheet.Cells['D13'].Formula = 'SUMPRODUCT((AllUsers[UserType]="Guest")*(AllUsers[AccountEnabled]="True"))'
    $worksheet.Cells['E13'].Formula = 'COUNTIFS(AllUsers[UserType],"Guest",AllUsers[Licenses],"<>")'
    $worksheet.Cells['F13'].Formula = 'SUMPRODUCT((AllUsers[UserType]="Guest")*(AllUsers[IsDirSynced]="True"))'
    $worksheet.Cells['F28'].Formula = 'SUMPRODUCT((AccountSummary[Type]="UserMailbox")*(AccountSummary[LitigationHold]="True"))'
    $worksheet.Cells['F29'].Formula = 'SUMPRODUCT((AccountSummary[Type]="SharedMailbox")*(AccountSummary[LitigationHold]="True"))'
    $worksheet.Cells['F30'].Formula = 'SUMPRODUCT((AccountSummary[Type]="EquipmentMailbox")*(AccountSummary[LitigationHold]="True"))'
    $worksheet.Cells['F31'].Formula = 'SUMPRODUCT((AccountSummary[Type]="RoomMailbox")*(AccountSummary[LitigationHold]="True"))'
    $worksheet.Cells['C40'].Formula = 'SUMPRODUCT(--(LEN(TRIM(AccountSummary[ArchiveSizeMB]&""))>0))'
    $worksheet.Cells['L28'].Formula = 'SUMPRODUCT((Devices[OperatingSystem]="Windows")*(Devices[TrustType]="Entra Registered"))'
    $worksheet.Cells['M28'].Formula = 'SUMPRODUCT((Devices[OperatingSystem]="Windows")*(Devices[TrustType]="Entra Joined"))'
    $worksheet.Cells['N28'].Formula = 'SUMPRODUCT((Devices[OperatingSystem]="Windows")*(Devices[TrustType]="Entra Hybrid Joined"))'

    # -Calculate forces EPPlus to evaluate formulas before saving, so
    # D12 shows its value immediately rather than needing Excel to
    # recalculate on first open.
    Close-ExcelPackage $excel -Calculate
    Write-Host "Wrote report metadata to the 'Data' worksheet in: $reportPath" -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to write report metadata to the 'Data' worksheet in '$reportPath': $($_.Exception.Message)"
    throw "Failed to write report metadata - see the log for details."
}
#endregion

#region Summary
Write-Host "`n============================================================" -ForegroundColor Yellow
Write-Host " Framework run complete" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "Report file (template + Data + AccountSummary + GroupSummary + SharePointSites + Devices + OneDrives + AllUsers + Contacts + MailEnabledSecurityGroups + DynamicDistributionGroups + DistributionGroups + SecurityGroups + AccountMailboxes + GroupMailboxes + Domains worksheets; other CSV data not yet written): $reportPath" -ForegroundColor Green
#endregion

# The warning/error log for this run is at: $logFilePath

Show-Report -OutputFile $reportPath
