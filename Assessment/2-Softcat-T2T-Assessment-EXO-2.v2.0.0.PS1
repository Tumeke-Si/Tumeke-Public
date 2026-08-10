#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Connects to Exchange Online for the Softcat T2T assessment and
    exports connector, transport, group and contact data.

.DESCRIPTION
    Connects to Exchange Online using the certificate-secured Azure AD
    app created by 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1, then
    retrieves and exports (to a "CSV Files" subfolder):
      - DKIM signing configuration                  -> DKIMRecords.csv
      - Inbound connectors                          -> InboundConnectors.csv
      - Outbound connectors                         -> OutboundConnectors.csv
      - Transport rules                              -> TransportRules.csv
      - Distribution groups (with membership)        -> DistributionGroups.csv
      - Dynamic distribution groups                  -> DynamicDistributionGroups.csv
      - Public folders, recursive, with statistics   -> PublicFolders.csv
      - Mail-enabled security groups (with members)  -> MailEnabledSecurityGroups.csv
      - Mail contacts                                -> Contacts.csv

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

    Warnings and errors - including per-item failures, which are logged
    and skipped rather than stopping the whole run - are written to a
    timestamped log file in a "Logs" subfolder.

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
    .\2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1

    Runs interactively, discovering ConnectionDetails.json and
    OrganizationDetails.json files automatically and writing output
    under the script's folder.

.EXAMPLE
    .\2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1 -AppId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -Organization contoso.onmicrosoft.com -OutputPath C:\Assessments\Contoso

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
$logFilePath = Join-Path $logsFolderPath "Get-EXOData-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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
    #region DKIM signing configuration
    Write-Host "Getting DKIM signing configuration..." -ForegroundColor Cyan
    $DKIMConfig = Get-DkimSigningConfig
    $DKIMConfigArray = @()
    $dkimCount = $DKIMConfig.Count
    $dkimIndex = 0

    Foreach ($DKIM in $DKIMConfig) {
        $dkimIndex++
        Write-Progress -Id 1 -Activity "Processing DKIM signing configuration" -Status "$dkimIndex of $dkimCount - $($DKIM.Domain)" -PercentComplete (($dkimIndex / $dkimCount) * 100)
        try {
            $DKIMConfigArray += [PSCustomObject]@{
                Domain         = $DKIM.Domain
                Enabled        = $DKIM.Enabled
                Selector1CNAME = $DKIM.Selector1CNAME
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process DKIM config for '$($DKIM.Domain)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 1 -Activity "Processing DKIM signing configuration" -Completed

    $DKIMConfigArray | Export-Csv (Join-Path $csvFolderPath "DKIMRecords.csv") -NoTypeInformation
    Write-Host "DKIM data exported." -ForegroundColor Green
    #endregion

    #region Inbound connectors
    Write-Host "Getting inbound connectors..." -ForegroundColor Cyan
    $InboundConnectors = Get-InboundConnector | Sort-Object Name
    $InboundConnectorsArray = @()
    $inboundCount = $InboundConnectors.Count
    $inboundIndex = 0

    Foreach ($InboundConnector in $InboundConnectors) {
        $inboundIndex++
        Write-Progress -Id 2 -Activity "Processing inbound connectors" -Status "$inboundIndex of $inboundCount - $($InboundConnector.Name)" -PercentComplete (($inboundIndex / $inboundCount) * 100)
        try {
            $InboundConnectorsArray += [PSCustomObject]@{
                Name              = $InboundConnector.Name
                Enabled           = $InboundConnector.Enabled
                ConnectorType     = $InboundConnector.ConnectorType
                SenderIPAddresses = $InboundConnector.SenderIPAddresses -join ","
                SenderDomains     = $InboundConnector.SenderDomains -join ","
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process inbound connector '$($InboundConnector.Name)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 2 -Activity "Processing inbound connectors" -Completed

    $InboundConnectorsArray | Export-Csv (Join-Path $csvFolderPath "InboundConnectors.csv") -NoTypeInformation
    Write-Host "Inbound connector data exported." -ForegroundColor Green
    #endregion

    #region Outbound connectors
    Write-Host "Getting outbound connectors..." -ForegroundColor Cyan
    $OutboundConnectors = Get-OutboundConnector | Sort-Object Name
    $OutboundConnectorsArray = @()
    $outboundCount = $OutboundConnectors.Count
    $outboundIndex = 0

    Foreach ($OutboundConnector in $OutboundConnectors) {
        $outboundIndex++
        Write-Progress -Id 3 -Activity "Processing outbound connectors" -Status "$outboundIndex of $outboundCount - $($OutboundConnector.Name)" -PercentComplete (($outboundIndex / $outboundCount) * 100)
        try {
            $OutboundConnectorsArray += [PSCustomObject]@{
                Name             = $OutboundConnector.Name
                Enabled          = $OutboundConnector.Enabled
                ConnectorType    = $OutboundConnector.ConnectorType
                UseMXRecord      = $OutboundConnector.UseMXRecord
                IsValidated      = $OutboundConnector.IsValidated
                TlsSettings      = $OutboundConnector.TlsSettings
                SmartHosts       = $OutboundConnector.SmartHosts -join ","
                RecipientDomains = $OutboundConnector.RecipientDomains -join ","
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process outbound connector '$($OutboundConnector.Name)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 3 -Activity "Processing outbound connectors" -Completed

    $OutboundConnectorsArray | Export-Csv (Join-Path $csvFolderPath "OutboundConnectors.csv") -NoTypeInformation
    Write-Host "Outbound connector data exported." -ForegroundColor Green
    #endregion

    #region Transport rules
    Write-Host "Getting transport rules..." -ForegroundColor Cyan
    $TransportRules = Get-TransportRule | Sort-Object Priority
    $TransportRulesArray = @()
    $trCount = $TransportRules.Count
    $trIndex = 0

    ForEach ($TransportRule in $TransportRules) {
        $trIndex++
        Write-Progress -Id 4 -Activity "Processing transport rules" -Status "$trIndex of $trCount - $($TransportRule.Name)" -PercentComplete (($trIndex / $trCount) * 100)
        try {
            $TransportRulesArray += [PSCustomObject]@{
                Name     = $TransportRule.Name
                State    = $TransportRule.State
                Mode     = $TransportRule.Mode
                Priority = $TransportRule.Priority
                Comments = $TransportRule.Comments
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process transport rule '$($TransportRule.Name)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 4 -Activity "Processing transport rules" -Completed

    $TransportRulesArray | Export-Csv (Join-Path $csvFolderPath "TransportRules.csv") -NoTypeInformation
    Write-Host "Transport rules data exported." -ForegroundColor Green
    #endregion

    #region Distribution groups
    Write-Host "Getting distribution groups..." -ForegroundColor Cyan
    $DistributionGroups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
    $DistributionGroupsArray = @()
    $dgCount = $DistributionGroups.Count
    $dgIndex = 0

    Foreach ($DistributionGroup in $DistributionGroups) {
        $dgIndex++
        Write-Progress -Id 5 -Activity "Processing distribution groups" -Status "$dgIndex of $dgCount - $($DistributionGroup.DisplayName)" -PercentComplete (($dgIndex / $dgCount) * 100)
        try {
            $DistributionGroupMembers = Get-DistributionGroupMember -ResultSize Unlimited -Identity $DistributionGroup.Name

            $DistributionGroupsArray += [PSCustomObject]@{
                Name                                                   = $DistributionGroup.Name
                DisplayName                                            = $DistributionGroup.DisplayName
                ManagedBy                                              = $DistributionGroupMembers.ManagedBy -join ";"
                GroupType                                              = $DistributionGroup.GroupType
                RecipientTypeDetails                                   = $DistributionGroup.RecipientTypeDetails
                Id                                                     = $DistributionGroup.Id
                PrimarySmtpAddress                                     = $DistributionGroup.PrimarySmtpAddress -join ";"
                WindowsEmailAddress                                    = $DistributionGroup.WindowsEmailAddress -join ";"
                EmailAddresses                                         = $DistributionGroup.EmailAddresses -join ";"
                HiddenFromAddressListsEnabled                          = $DistributionGroup.HiddenFromAddressListsEnabled
                MaxSendSize                                            = $DistributionGroup.MaxSendSize
                MaxReceiveSize                                         = $DistributionGroup.MaxReceiveSize
                ModeratedBy                                            = $DistributionGroup.ModeratedBy -join ";"
                IsDirSynced                                            = $DistributionGroup.IsDirSynced
                GrantSendOnBehalfTo                                    = $DistributionGroup.GrantSendOnBehalfTo -join ";"
                AcceptMessagesOnlyFromDLMembersWithDisplayNames        = $DistributionGroup.AcceptMessagesOnlyFromDLMembersWithDisplayNames -join ";"
                AcceptMessagesOnlyFromSendersOrMembersWithDisplayNames = $DistributionGroup.AcceptMessagesOnlyFromSendersOrMembersWithDisplayNames -join ";"
                AcceptMessagesOnlyFromWithDisplayNames                 = $DistributionGroup.AcceptMessagesOnlyFromWithDisplayNames -join ";"
                BccBlocked                                             = $DistributionGroup.BccBlocked
                BypassModerationFromSendersOrMembersWithDisplayNames   = $DistributionGroup.BypassModerationFromSendersOrMembersWithDisplayNames
                Description                                            = $DistributionGroup.Description
                GrantSendOnBehalfToWithDisplayNames                    = $DistributionGroup.GrantSendOnBehalfToWithDisplayNames -join ";"
                HiddenGroupMembershipEnabled                           = $DistributionGroup.HiddenGroupMembershipEnabled
                MemberDepartRestriction                                = $DistributionGroup.MemberDepartRestriction
                MemberJoinRestriction                                  = $DistributionGroup.MemberJoinRestriction
                MigrationToUnifiedGroupInProgress                      = $DistributionGroup.MigrationToUnifiedGroupInProgress
                ModeratedByWithDisplayNames                            = $DistributionGroup.ModeratedByWithDisplayNames -join ";"
                ModerationEnabled                                      = $DistributionGroup.ModerationEnabled
                RejectMessagesFrom                                     = $DistributionGroup.RejectMessagesFrom -join ";"
                RejectMessagesFromDLMembers                            = $DistributionGroup.RejectMessagesFromDLMembers -join ";"
                RejectMessagesFromSendersOrMembers                     = $DistributionGroup.RejectMessagesFromSendersOrMembers -join ";"
                RejectMessagesFromSendersOrMembersWithDisplayNames     = $DistributionGroup.RejectMessagesFromSendersOrMembersWithDisplayNames -join ";"
                ReportToManagerEnabled                                 = $DistributionGroup.ReportToManagerEnabled
                ReportToOriginatorEnabled                              = $DistributionGroup.ReportToOriginatorEnabled
                RequireSenderAuthenticationEnabled                     = $DistributionGroup.RequireSenderAuthenticationEnabled
                SendOofMessageToOriginatorEnabled                      = $DistributionGroup.SendOofMessageToOriginatorEnabled
                WhenChanged                                            = $DistributionGroup.WhenChanged
                WhenCreated                                            = $DistributionGroup.WhenCreated
                MemberCount                                            = $DistributionGroupMembers.Count
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process distribution group '$($DistributionGroup.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 5 -Activity "Processing distribution groups" -Completed

    $DistributionGroupsArray | Export-Csv (Join-Path $csvFolderPath "DistributionGroups.csv") -NoTypeInformation
    Write-Host "Distribution group data exported." -ForegroundColor Green
    #endregion

    #region Dynamic distribution groups
    Write-Host "Getting dynamic distribution groups..." -ForegroundColor Cyan
    $DynamicDistributionGroups = Get-DynamicDistributionGroup -ResultSize Unlimited | Sort-Object DisplayName
    $DynamicDistributionGroupsArray = @()
    $ddgCount = $DynamicDistributionGroups.Count
    $ddgIndex = 0

    Foreach ($DynamicDistributionGroup in $DynamicDistributionGroups) {
        $ddgIndex++
        Write-Progress -Id 6 -Activity "Processing dynamic distribution groups" -Status "$ddgIndex of $ddgCount - $($DynamicDistributionGroup.DisplayName)" -PercentComplete (($ddgIndex / $ddgCount) * 100)
        try {
            $DynamicDistributionGroupsArray += [PSCustomObject]@{
                Name                 = $DynamicDistributionGroup.Name
                DisplayName          = $DynamicDistributionGroup.DisplayName
                RecipientFilterType  = $DynamicDistributionGroup.RecipientFilterType
                RecipientTypeDetails = $DynamicDistributionGroup.RecipientTypeDetails
                PrimarySmtpAddress   = $DynamicDistributionGroup.PrimarySmtpAddress
                ManagedBy            = $DynamicDistributionGroup.ManagedBy -join ','
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process dynamic distribution group '$($DynamicDistributionGroup.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 6 -Activity "Processing dynamic distribution groups" -Completed

    $DynamicDistributionGroupsArray | Export-Csv (Join-Path $csvFolderPath "DynamicDistributionGroups.csv") -NoTypeInformation
    Write-Host "Dynamic distribution group data exported." -ForegroundColor Green
    #endregion

    #region Public folders
    Write-Host "Getting public folders..." -ForegroundColor Cyan
    # Public folder trees can throw non-terminating errors on individual
    # subfolders (e.g. permissions) - briefly relaxed here so the
    # recursive walk doesn't abort, then restored to Stop afterward so
    # the rest of the script keeps its normal error handling.
    $ErrorActionPreference = "SilentlyContinue"
    $PublicFolders = Get-PublicFolder -Recurse -ResultSize Unlimited | Sort-Object ParentPath
    $ErrorActionPreference = "Stop"

    $PublicFolderArray = @()
    $pfCount = $PublicFolders.Count
    $pfIndex = 0

    Foreach ($PublicFolder in $PublicFolders) {
        $pfIndex++
        Write-Progress -Id 7 -Activity "Processing public folders" -Status "$pfIndex of $pfCount - $($PublicFolder.Name)" -PercentComplete (($pfIndex / $pfCount) * 100)
        try {
            $PublicFolderStats = Get-PublicFolderStatistics -Identity $PublicFolder.Identity | Select-Object Name, @{Name = "TotalItemSizeMB"; Expression = { [math]::Round(($_.TotalItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }, ItemCount, DeletedItemCount, @{Name = "TotalDeletedItemSizeMB"; Expression = { [math]::Round(($_.TotalDeletedItemSize.ToString().Split("(")[1].Split(" ")[0].Replace(",", "") / 1MB), 0) } }

            $PublicFolderArray += [PSCustomObject]@{
                Name            = $PublicFolder.Name
                Identity        = $PublicFolder.Identity
                ParentPath      = $PublicFolder.ParentPath
                ItemCount       = $PublicFolderStats.ItemCount
                TotalItemSizeMB = $PublicFolderStats.TotalItemSizeMB
                FolderClass     = $PublicFolder.FolderClass
                MailEnabled     = $PublicFolder.MailEnabled
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to retrieve statistics for public folder '$($PublicFolder.Identity)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 7 -Activity "Processing public folders" -Completed

    $PublicFolderArray | Export-Csv (Join-Path $csvFolderPath "PublicFolders.csv") -NoTypeInformation
    Write-Host "Public folder data exported." -ForegroundColor Green
    #endregion

    #region Mail-enabled security groups
    Write-Host "Getting mail-enabled security groups..." -ForegroundColor Cyan
    $MailEnabledSecurityGroups = Get-DistributionGroup -RecipientTypeDetails MailUniversalSecurityGroup -ResultSize Unlimited | Sort-Object Name
    $MailEnabledSecurityGroupsArray = @()
    $mesgCount = $MailEnabledSecurityGroups.Count
    $mesgIndex = 0

    Foreach ($MailEnabledSecurityGroup in $MailEnabledSecurityGroups) {
        $mesgIndex++
        Write-Progress -Id 8 -Activity "Processing mail-enabled security groups" -Status "$mesgIndex of $mesgCount - $($MailEnabledSecurityGroup.DisplayName)" -PercentComplete (($mesgIndex / $mesgCount) * 100)
        try {
            $MemberCount = (Get-DistributionGroupMember -Identity $MailEnabledSecurityGroup.Identity -ResultSize Unlimited).Count

            $MailEnabledSecurityGroupsArray += [PSCustomObject]@{
                Name                           = $MailEnabledSecurityGroup.Name
                DisplayName                    = $MailEnabledSecurityGroup.DisplayName
                Alias                          = $MailEnabledSecurityGroup.Alias
                PrimarySmtpAddress             = $MailEnabledSecurityGroup.PrimarySmtpAddress
                ManagedBy                      = ($MailEnabledSecurityGroup.ManagedBy -join ",")
                OrganizationalUnit             = $MailEnabledSecurityGroup.OrganizationalUnit
                RequireSenderAuthentication    = $MailEnabledSecurityGroup.RequireSenderAuthenticationEnabled
                HiddenFromAddressListsEnabled  = $MailEnabledSecurityGroup.HiddenFromAddressListsEnabled
                MemberCount                    = $MemberCount
                RecipientTypeDetails           = $MailEnabledSecurityGroup.RecipientTypeDetails
                Identity                       = $MailEnabledSecurityGroup.Identity
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process mail-enabled security group '$($MailEnabledSecurityGroup.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 8 -Activity "Processing mail-enabled security groups" -Completed

    $MailEnabledSecurityGroupsArray | Export-Csv (Join-Path $csvFolderPath "MailEnabledSecurityGroups.csv") -NoTypeInformation -Encoding UTF8
    Write-Host "Mail-enabled security group data exported." -ForegroundColor Green
    #endregion

    #region Mail contacts
    Write-Host "Getting mail contacts..." -ForegroundColor Cyan
    $Contacts = Get-MailContact -ResultSize Unlimited
    $ContactArray = @()
    $contactCount = $Contacts.Count
    $contactIndex = 0

    Foreach ($Contact in $Contacts) {
        $contactIndex++
        Write-Progress -Id 9 -Activity "Processing mail contacts" -Status "$contactIndex of $contactCount - $($Contact.DisplayName)" -PercentComplete (($contactIndex / $contactCount) * 100)
        try {
            $ContactArray += [PSCustomObject]@{
                Name                 = $Contact.Name
                DisplayName          = $Contact.DisplayName
                Alias                = $Contact.Alias
                PrimarySmtpAddress   = $Contact.PrimarySmtpAddress
                RecipientTypeDetails = $Contact.RecipientTypeDetails
                Identity             = $Contact.Identity
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process mail contact '$($Contact.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 9 -Activity "Processing mail contacts" -Completed

    $ContactArray | Export-Csv (Join-Path $csvFolderPath "Contacts.csv") -NoTypeInformation -Encoding UTF8
    Write-Host "Mail contact data exported." -ForegroundColor Green
    #endregion
}
catch {
    Write-Log -Level Error -Message "2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1 stopped due to an unexpected error: $($_.Exception.Message)"
    throw
}

# All CSV outputs are in: $csvFolderPath
# The warning/error log for this run is at: $logFilePath
#
# The connection is left open intentionally so further EXO cmdlets can
# be run against it. Disconnect when finished:
#   Disconnect-ExchangeOnline -Confirm:$false
