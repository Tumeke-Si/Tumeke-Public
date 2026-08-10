#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Connects to Exchange Online for the Softcat T2T assessment and
    exports mailbox, group and domain data.

.DESCRIPTION
    Connects to Exchange Online using the certificate-secured Azure AD
    app created by 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1, then
    retrieves and exports (to a "CSV Files" subfolder):
      - Accepted domains                        -> AcceptedDomains.csv
      - SharePoint Admin URL                     -> AdminURL.txt
      - Custom (non-onmicrosoft.com) domains     -> CustomDomains.csv
      - Mailbox stats, holds and retention data  -> AccountMailboxes.csv
      - Microsoft 365 Groups                     -> UnifiedGroups.csv
      - Microsoft 365 Group mailbox statistics   -> GroupMailboxes.csv

    One of the "separate scripts" referenced by
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 - only connects
    to Exchange Online (not Microsoft Graph), since that's all this
    needs.

    Resolves AppId/TenantId/CertificateThumbprint from (in order of
    preference): explicit parameters, a ConnectionDetails.json file, or
    manual entry if neither is available.

    Connect-ExchangeOnline's -Organization parameter does NOT accept
    the TenantId GUID (confirmed: "Organization cannot be a Guid,
    please enter the name of the tenant instead") - it needs the
    tenant's .onmicrosoft.com domain. Rather than connecting to Graph
    just to look this up, it's resolved from the
    *-OrganizationDetails.json file already exported by
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 (in its JSON
    subfolder), with manual entry as the fallback.

    Assumes the certificate has already been imported into this
    machine's certificate store (e.g. via
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1, or manually
    with Import-PfxCertificate). If Connect-ExchangeOnline below fails
    with a certificate-not-found error, import it first.

    Warnings and errors - including per-mailbox and per-group failures,
    which are logged and skipped rather than stopping the whole run -
    are written to a timestamped log file in a "Logs" subfolder.

.PARAMETER AppId
    The Azure AD app registration's ClientId. If not supplied, resolved
    from a ConnectionDetails.json file or prompted for manually.

.PARAMETER TenantId
    The Azure AD tenant Id. If not supplied, resolved from a
    ConnectionDetails.json file or prompted for manually.

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate (in the current user's certificate
    store) used to authenticate the app registration. If not supplied,
    resolved from a ConnectionDetails.json file or prompted for
    manually.

.PARAMETER ConnectionDetailsPath
    Path to a *-ConnectionDetails.json file exported by
    0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1. If not supplied,
    the script looks for one automatically in its own folder.

.PARAMETER Organization
    Value for Connect-ExchangeOnline's -Organization parameter (the
    tenant's .onmicrosoft.com domain - Connect-ExchangeOnline does not
    accept the TenantId GUID here). If not supplied, resolved from a
    *-OrganizationDetails.json file (see -OrganizationDetailsPath), or
    asked for manually if that's not available.

.PARAMETER OrganizationDetailsPath
    Path to a *-OrganizationDetails.json file exported by
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1. If not
    supplied, the script looks for one automatically in a JSON
    subfolder alongside this script.

.PARAMETER OutputPath
    Base folder for outputs. "CSV Files" and "Logs" subfolders are
    created under this (if they don't already exist). Defaults to the
    same folder as this script.

.EXAMPLE
    .\2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1

    Runs interactively, discovering ConnectionDetails.json and
    OrganizationDetails.json files automatically and writing output
    under the script's folder.

.EXAMPLE
    .\2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1 -AppId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -Organization contoso.onmicrosoft.com -OutputPath C:\Assessments\Contoso

    Runs non-interactively with explicit connection details, writing
    output under C:\Assessments\Contoso.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    RUN ORDER: this script starts with "2-Softcat-T2T-Assessment-", so
    it can be run in any order relative to the other 2-Softcat-T2T-
    Assessment-* scripts, once 0-Softcat-T2T-Assessment-Create-App and
    1-Softcat-T2T-Assessment-Initial-Script have both been run.

    Prerequisites:
      Install-Module ExchangeOnlineManagement -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [string]$AppId,

    [string]$TenantId,

    [string]$CertificateThumbprint,

    [string]$ConnectionDetailsPath,

    [string]$Organization,

    [string]$OrganizationDetailsPath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $scriptDir
}

#region Ensure output folders exist
$csvFolderPath = Join-Path $OutputPath "CSV Files"
if (-not (Test-Path $csvFolderPath)) {
    New-Item -Path $csvFolderPath -ItemType Directory | Out-Null
}

$logsFolderPath = Join-Path $OutputPath "Logs"
if (-not (Test-Path $logsFolderPath)) {
    New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
}
#endregion

#region Logging
# Every warning/error from this point on - including the WARNING
# connection-resolution notices below - is written to this file as
# well as the console, so a consultant can review what happened after
# the run without having to scroll back through the console output.
$logFilePath = Join-Path $logsFolderPath "Get-MailboxStatistics-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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

#region Resolve AppId / TenantId / CertificateThumbprint
# Precedence: explicit parameters > a ConnectionDetails.json file > manual entry.
if ([string]::IsNullOrWhiteSpace($AppId) -or [string]::IsNullOrWhiteSpace($TenantId) -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {

    if ([string]::IsNullOrWhiteSpace($ConnectionDetailsPath)) {
        $discovered = Get-ChildItem -Path $scriptDir -Filter "*-ConnectionDetails.json" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($discovered) {
            $useJsonAnswer = Read-Host "Found connection details file '$($discovered.Name)' - use it? (Y/N)"
            if ($useJsonAnswer -match '^[Yy]') {
                $ConnectionDetailsPath = $discovered.FullName
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ConnectionDetailsPath)) {
        if (-not (Test-Path $ConnectionDetailsPath)) {
            Write-Log -Message "WARNING: couldn't find a connection details file at '$ConnectionDetailsPath' - falling back to manual entry for anything still missing."
        }
        else {
            $details = Get-Content -Path $ConnectionDetailsPath -Raw | ConvertFrom-Json
            if ([string]::IsNullOrWhiteSpace($AppId)) { $AppId = $details.AppId }
            if ([string]::IsNullOrWhiteSpace($TenantId)) { $TenantId = $details.TenantId }
            if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $CertificateThumbprint = $details.CertificateThumbprint }
            Write-Host "Loaded connection details from: $ConnectionDetailsPath" -ForegroundColor Green
        }
    }

    if ([string]::IsNullOrWhiteSpace($AppId)) { $AppId = Read-Host "Enter the AppId (ClientId)" }
    if ([string]::IsNullOrWhiteSpace($TenantId)) { $TenantId = Read-Host "Enter the TenantId" }
    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) { $CertificateThumbprint = Read-Host "Enter the certificate thumbprint" }
}
#endregion

#region Resolve the Exchange Online -Organization value
# Precedence: explicit parameter > an OrganizationDetails.json file > manual entry.
if ([string]::IsNullOrWhiteSpace($Organization)) {

    if ([string]::IsNullOrWhiteSpace($OrganizationDetailsPath)) {
        $jsonFolderPath = Join-Path $scriptDir "JSON"
        $discoveredOrgDetails = Get-ChildItem -Path $jsonFolderPath -Filter "*-OrganizationDetails.json" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($discoveredOrgDetails) {
            $useOrgAnswer = Read-Host "Found organisation details file '$($discoveredOrgDetails.Name)' - use it to resolve the Exchange Online organisation domain? (Y/N)"
            if ($useOrgAnswer -match '^[Yy]') {
                $OrganizationDetailsPath = $discoveredOrgDetails.FullName
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($OrganizationDetailsPath)) {
        if (-not (Test-Path $OrganizationDetailsPath)) {
            Write-Log -Message "WARNING: couldn't find an organisation details file at '$OrganizationDetailsPath' - falling back to manual entry."
        }
        else {
            $orgDetails = Get-Content -Path $OrganizationDetailsPath -Raw | ConvertFrom-Json
            $resolvedDomain = ($orgDetails.VerifiedDomains | Where-Object { $_.IsInitial }).Name
            if ($resolvedDomain) {
                $Organization = $resolvedDomain
                Write-Host "Resolved Exchange Online organisation domain from file: $Organization" -ForegroundColor Green
            }
            else {
                Write-Log -Message "WARNING: couldn't find an initial verified domain in '$OrganizationDetailsPath' - falling back to manual entry."
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($Organization)) {
        $Organization = Read-Host "Enter the tenant's primary domain for Exchange Online (e.g. contoso.onmicrosoft.com)"
    }
}
#endregion

#region Connect to Exchange Online
try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to connect to Exchange Online: $($_.Exception.Message)"
    throw
}
#endregion

try {
    #region Accepted domains, SharePoint Admin URL and custom domains
    Write-Host "Getting accepted domains..." -ForegroundColor Cyan
    $AcceptedDomains = Get-AcceptedDomain | Sort-Object DomainName
    $AcceptedDomains | Export-Csv (Join-Path $csvFolderPath "AcceptedDomains.csv") -NoTypeInformation
    Write-Host "Accepted domains exported." -ForegroundColor Green

    Write-Host "Exporting SharePoint Admin URL for later scripts..." -ForegroundColor Cyan
    $OnMicrosoftDomain = $AcceptedDomains | Where-Object { $_.DomainName -like "*.onmicrosoft.com" } | Select-Object -ExpandProperty DomainName -First 1
    $OnMicrosoftPrefix = $OnMicrosoftDomain.Split('.')[0]
    $AdminURL = "https://" + $OnMicrosoftPrefix + "-admin.sharepoint.com"
    $AdminURL | Out-File (Join-Path $csvFolderPath "AdminURL.txt")
    Write-Host "SharePoint Admin URL exported." -ForegroundColor Green

    Write-Host "Exporting custom domains..." -ForegroundColor Cyan
    $CustomDomains = $AcceptedDomains | Where-Object { $_.DomainName -notlike "*.onmicrosoft.com" }
    $CustomDomains | Export-Csv (Join-Path $csvFolderPath "CustomDomains.csv") -NoTypeInformation
    Write-Host "Custom domains exported." -ForegroundColor Green
    #endregion

    #region Mailbox statistics, holds and retention data
    Write-Host "Getting account mailboxes..." -ForegroundColor Cyan
    $AllMailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.PrimarySMTPAddress -notlike "DiscoverySearchMailbox*" } | Sort-Object UserPrincipalName
    $mailboxCount = $AllMailboxes.Count
    $mailboxIndex = 0
    $AllMailboxesArray = @()

    Foreach ($Mailbox in $AllMailboxes) {
        $mailboxIndex++
        Write-Progress -Id 1 -Activity "Retrieving mailbox statistics" -Status "$mailboxIndex of $mailboxCount - $($Mailbox.UserPrincipalName)" -PercentComplete (($mailboxIndex / $mailboxCount) * 100)

        try {
            $MailboxStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }

            If ($Mailbox.ArchiveStatus -eq "Active") {
                $ArchiveStats = Get-MailboxStatistics -Identity $Mailbox.UserPrincipalName -Archive | Select-Object DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }
            }
            Else {
                $ArchiveStats = ""
            }

            $AllMailboxesArray += [PSCustomObject]@{
                Type                          = $Mailbox.RecipientTypeDetails
                UserPrincipalName             = $Mailbox.UserPrincipalName
                DisplayName                   = $Mailbox.DisplayName
                Alias                         = $Mailbox.Identity
                PrimarySmtpAddress            = $Mailbox.PrimarySmtpAddress
                ItemCount                     = $MailboxStats.ItemCount
                TotalItemSizeMB               = $MailboxStats.TotalItemSizeMB
                DeletedItemCount              = $MailboxStats.DeletedItemCount
                TotalDeletedItemSizeMB        = $MailboxStats.TotalDeletedItemSizeMB
                Archive                       = $Mailbox.ArchiveStatus
                ArchiveDisplayName            = $ArchiveStats.DisplayName
                ArchiveItemCount              = $ArchiveStats.ItemCount
                ArchiveTotalItemSizeMB        = $ArchiveStats.TotalItemSizeMB
                ArchiveDeletedItemCount       = $ArchiveStats.DeletedItemCount
                ArchiveTotalDeletedItemSizeMB = $ArchiveStats.TotalDeletedItemSizeMB
                WhenCreatedUTC                = $Mailbox.WhenCreatedUTC
                WhenChangedUTC                = $Mailbox.WhenChangedUTC
                LastLogonTime                 = $MailboxStats.LastLogonTime
                EmailAddresses                = $Mailbox.EmailAddresses -join ','
                LitigationHoldEnabled         = $Mailbox.LitigationHoldEnabled
                LitigationHoldDuration        = $Mailbox.LitigationHoldDuration
                InPlaceHolds                  = $Mailbox.InPlaceHolds -join ','
                RetentionPolicy               = $Mailbox.RetentionPolicy
                RetentionHoldEnabled          = $Mailbox.RetentionHoldEnabled
                StartDateForRetentionHold     = $Mailbox.StartDateForRetentionHold
                EndDateForRetentionHold       = $Mailbox.EndDateForRetentionHold
                AccountSKU                    = $licenseString -join ','
                Guid                          = $Mailbox.Guid
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to retrieve statistics for mailbox '$($Mailbox.UserPrincipalName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 1 -Activity "Retrieving mailbox statistics" -Completed

    $AllMailboxesArray | Export-Csv (Join-Path $csvFolderPath "AccountMailboxes.csv") -NoTypeInformation
    Write-Host "Mailbox data exported." -ForegroundColor Green
    #endregion

    #region Microsoft 365 Groups
    Write-Host "Getting Microsoft 365 Groups..." -ForegroundColor Cyan
    $UnifiedGroups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName
    $groupCount = $UnifiedGroups.Count
    $groupIndex = 0
    $UnifiedGroupArray = @()

    foreach ($UnifiedGroup in $UnifiedGroups) {
        $groupIndex++
        Write-Progress -Id 2 -Activity "Processing Microsoft 365 Groups" -Status "$groupIndex of $groupCount - $($UnifiedGroup.DisplayName)" -PercentComplete (($groupIndex / $groupCount) * 100)

        try {
            $Members = Get-UnifiedGroup -Identity $UnifiedGroup.ExternalDirectoryObjectId | Get-UnifiedGroupLinks -LinkType Member
            $MemberCount = $Members.Count

            $UnifiedGroupArray += [PSCustomObject]@{
                DisplayName        = $UnifiedGroup.DisplayName
                AccessType         = $UnifiedGroup.AccessType
                PrimarySmtpAddress = $UnifiedGroup.PrimarySmtpAddress
                GroupID            = $UnifiedGroup.ExternalDirectoryObjectId
                Members            = $MemberCount
                Guests             = $UnifiedGroup.GroupExternalMemberCount
                WhenCreatedUTC     = $UnifiedGroup.WhenCreatedUTC
                WhenChangedUTC     = $UnifiedGroup.WhenChangedUTC
                SharePointSiteURL  = $UnifiedGroup.SharePointSiteURL
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process Microsoft 365 Group '$($UnifiedGroup.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 2 -Activity "Processing Microsoft 365 Groups" -Completed

    Write-Host "Resting....." -ForegroundColor Yellow
    Start-Sleep 5

    $UnifiedGroupArray | Export-Csv (Join-Path $csvFolderPath "UnifiedGroups.csv") -NoTypeInformation
    Write-Host "Unified Group data exported." -ForegroundColor Green
    #endregion

    #region Microsoft 365 Group mailbox data
    Write-Host "Getting Microsoft 365 Group mailbox data..." -ForegroundColor Cyan
    $groupMailboxIndex = 0
    $UnifiedGroupMailboxesArray = @()

    foreach ($UnifiedGroup in $UnifiedGroups) {
        $groupMailboxIndex++
        Write-Progress -Id 3 -Activity "Building group mailbox data" -Status "$groupMailboxIndex of $groupCount - $($UnifiedGroup.DisplayName)" -PercentComplete (($groupMailboxIndex / $groupCount) * 100)

        try {
            $UnifiedMailboxStats = Get-MailboxStatistics -Identity $UnifiedGroup.PrimarySmtpAddress | Select-Object LastLogonTime, DisplayName, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }

            $UnifiedGroupMailboxesArray += [PSCustomObject]@{
                DisplayName            = $UnifiedGroup.DisplayName
                Alias                  = $UnifiedGroup.Identity
                PrimarySmtpAddress     = $UnifiedGroup.PrimarySmtpAddress
                RecipientTypeDetails   = $UnifiedGroup.RecipientTypeDetails
                ItemCount              = $UnifiedMailboxStats.ItemCount
                TotalItemSizeMB        = $UnifiedMailboxStats.TotalItemSizeMB
                DeletedItemCount       = $UnifiedMailboxStats.DeletedItemCount
                TotalDeletedItemSizeMB = $UnifiedMailboxStats.TotalDeletedItemSizeMB
                WhenCreatedUTC         = $UnifiedGroup.WhenCreatedUTC
                WhenChangedUTC         = $UnifiedGroup.WhenChangedUTC
                LastLogonTime          = $UnifiedMailboxStats.LastLogonTime
                EmailAddresses         = $UnifiedGroup.EmailAddresses -join ','
                SPOSiteURL             = $UnifiedGroup.SharePointSiteURL
                Guid                   = $UnifiedGroup.Guid
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to retrieve mailbox statistics for group '$($UnifiedGroup.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 3 -Activity "Building group mailbox data" -Completed

    $UnifiedGroupMailboxesArray | Export-Csv (Join-Path $csvFolderPath "GroupMailboxes.csv") -NoTypeInformation
    Write-Host "Unified Group mailbox data exported." -ForegroundColor Green
    #endregion
}
catch {
    Write-Log -Level Error -Message "2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1 stopped due to an unexpected error: $($_.Exception.Message)"
    throw
}

# All CSV/txt outputs are in: $csvFolderPath
# The warning/error log for this run is at: $logFilePath
#
# The connection is left open intentionally so further EXO cmdlets can
# be run against it. Disconnect when finished:
#   Disconnect-ExchangeOnline -Confirm:$false
