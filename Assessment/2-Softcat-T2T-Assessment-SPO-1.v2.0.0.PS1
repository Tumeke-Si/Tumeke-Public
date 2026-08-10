#Requires -Modules PnP.PowerShell

<#
.SYNOPSIS
    Connects to SharePoint Online for the Softcat T2T assessment and
    exports SharePoint/OneDrive site data.

.DESCRIPTION
    Connects to SharePoint Online using PnP PowerShell, with the same
    certificate-secured Azure AD app created by
    0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1, then retrieves and
    exports (to a "CSV Files" subfolder):
      - All SharePoint sites -> SharePointSites.csv
      - All OneDrive sites   -> OneDriveSites.csv

    Both come from a single Get-PnPTenantSite -IncludeOneDriveSites
    call, split by Template (OneDrive/personal sites use "SPSPERS")
    rather than querying the tenant twice.

    One of the "separate scripts" referenced by
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 - only connects
    to SharePoint Online (not Graph/Exchange Online/Teams), since
    that's all this needs. Still deliberately minimal beyond that -
    more SharePoint data (permissions, external sharing, storage
    detail, etc.) can be added the same way later.

    Requires the app to have been created with
    0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1, which grants the
    SharePoint Online Sites.Read.All permission this needs. That's a
    SEPARATE permission from Microsoft Graph's own Sites.Read.All - PnP
    PowerShell talks to SharePoint's REST/CSOM API directly, which
    doesn't accept Graph-issued permissions.

    Resolves AppId/TenantId/CertificateThumbprint from (in order of
    preference): explicit parameters, a ConnectionDetails.json file, or
    manual entry if neither is available.

    Connect-PnPOnline needs a SharePoint admin center URL
    (https://<tenant>-admin.sharepoint.com), not just a tenant ID. This
    is resolved from the *-OrganizationDetails.json file already
    exported by 1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 (in
    its JSON subfolder) - the tenant's initial .onmicrosoft.com domain
    gives us the prefix needed to build the admin URL - with manual
    entry as the fallback.

    Assumes the certificate has already been imported into this
    machine's certificate store (e.g. via
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1, or manually
    with Import-PfxCertificate).

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

.PARAMETER AdminUrl
    SharePoint admin center URL (https://<tenant>-admin.sharepoint.com).
    If not supplied, resolved from a *-OrganizationDetails.json file
    (see -OrganizationDetailsPath), or asked for manually if that's
    not available.

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
    .\2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1

    Runs interactively, discovering ConnectionDetails.json and
    OrganizationDetails.json files automatically and writing output
    under the script's folder.

.EXAMPLE
    .\2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1 -AppId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -AdminUrl https://contoso-admin.sharepoint.com -OutputPath C:\Assessments\Contoso

    Runs non-interactively with explicit connection details, writing
    output under C:\Assessments\Contoso.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    RUN ORDER: this script starts with "2-Softcat-T2T-Assessment-", so
    it can be run in any order relative to the other 2-Softcat-T2T-
    Assessment-* scripts, once 0-Softcat-T2T-Assessment-Create-App and
    1-Softcat-T2T-Assessment-Initial-Script have both been run.

    Prerequisites:
      Install-Module -Name PnP.PowerShell -Scope CurrentUser

    Troubleshooting - "Current site is not a tenant administration site":
      The AdminUrl auto-resolved from *-OrganizationDetails.json is built
      from that tenant's IsInitial verified domain (<prefix>.onmicrosoft.com
      -> https://<prefix>-admin.sharepoint.com). If the tenant's SharePoint
      domain was changed (e.g. via a Microsoft-assisted tenant rename), that
      prefix can be stale even though the domain still resolves in DNS, and
      Connect-PnPOnline will succeed but Get-PnPTenantSite will fail with
      this error. If you hit it, check the SharePoint admin center for the
      tenant's current admin URL, then re-run this script, answer "N" when
      asked whether to use the discovered OrganizationDetails.json file, and
      enter the correct AdminUrl manually when prompted.
#>
[CmdletBinding()]
param(
    [string]$AppId,

    [string]$TenantId,

    [string]$CertificateThumbprint,

    [string]$ConnectionDetailsPath,

    [string]$AdminUrl,

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
$logFilePath = Join-Path $logsFolderPath "Get-SharePointData-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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

#region Resolve the SharePoint admin center URL
# Precedence: explicit parameter > an OrganizationDetails.json file > manual entry.
if ([string]::IsNullOrWhiteSpace($AdminUrl)) {

    if ([string]::IsNullOrWhiteSpace($OrganizationDetailsPath)) {
        $jsonFolderPath = Join-Path $scriptDir "JSON"
        $discoveredOrgDetails = Get-ChildItem -Path $jsonFolderPath -Filter "*-OrganizationDetails.json" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($discoveredOrgDetails) {
            Write-Host "Note: the URL resolved from this file can be stale if the tenant's SharePoint domain was changed (e.g. a tenant rename) - if the script later fails with 'Current site is not a tenant administration site', answer N here, check the SharePoint admin center for the current admin URL, and enter it manually." -ForegroundColor Yellow
            Write-Host "Note: if you suspect the URL is stale, answer N here now, check the SharePoint admin center for the current admin URL, and enter it manually." -ForegroundColor Yellow
            $useOrgAnswer = Read-Host "Found organisation details file '$($discoveredOrgDetails.Name)' - use it to resolve the SharePoint admin URL? (Y/N)"
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
            $initialDomain = ($orgDetails.VerifiedDomains | Where-Object { $_.IsInitial }).Name
            if ($initialDomain) {
                $prefix = $initialDomain.Split('.')[0]
                $AdminUrl = "https://$prefix-admin.sharepoint.com"
                Write-Host "Resolved SharePoint admin URL from file: $AdminUrl" -ForegroundColor Green
            }
            else {
                Write-Log -Message "WARNING: couldn't find an initial verified domain in '$OrganizationDetailsPath' - falling back to manual entry."
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($AdminUrl)) {
        $AdminUrl = Read-Host "Enter the SharePoint admin centre URL (e.g. https://contoso-admin.sharepoint.com)"
    }
}
#endregion

#region Connect to SharePoint Online
try {
    Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $AdminUrl -ClientId $AppId -Tenant $TenantId -Thumbprint $CertificateThumbprint
    Write-Host "Connected to SharePoint Online." -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to connect to SharePoint Online: $($_.Exception.Message). If this is an access-denied error rather than a connection failure, confirm the app was created with 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1, which grants the SharePoint-specific Sites.Read.All permission this needs - see header notes."
    throw
}
#endregion

try {
    #region SharePoint sites and OneDrive sites
    # Get-PnPTenantSite excludes OneDrive personal sites by default -
    # -IncludeOneDriveSites brings them in too, in the same call, and
    # they're split out into their own CSV below by Template
    # (OneDrive/personal sites use the "SPSPERS" template) rather than
    # making a second full tenant enumeration call just for those.
    Write-Host "Getting SharePoint and OneDrive sites..." -ForegroundColor Cyan
    $Sites = Get-PnPTenantSite -IncludeOneDriveSites
    $SitesArray = @()
    $OneDriveSitesArray = @()
    $siteCount = $Sites.Count
    $siteIndex = 0

    Foreach ($Site in $Sites) {
        $siteIndex++
        Write-Progress -Id 1 -Activity "Processing sites" -Status "$siteIndex of $siteCount - $($Site.Url)" -PercentComplete (($siteIndex / $siteCount) * 100)
        try {
            $siteRow = [PSCustomObject]@{
                Url                     = $Site.Url
                Title                   = $Site.Title
                Owner                   = $Site.Owner
                Template                = $Site.Template
                StorageUsageCurrentMB   = $Site.StorageUsageCurrent
                StorageQuotaMB          = $Site.StorageQuota
                LastContentModifiedDate = $Site.LastContentModifiedDate
                Status                  = $Site.Status
                SharingCapability       = $Site.SharingCapability
                GroupId                 = $Site.GroupId
            }

            if ($Site.Template -like "SPSPERS*") {
                $OneDriveSitesArray += $siteRow
            }
            else {
                $SitesArray += $siteRow
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process site '$($Site.Url)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 1 -Activity "Processing sites" -Completed

    $SitesArray | Export-Csv (Join-Path $csvFolderPath "SharePointSites.csv") -NoTypeInformation
    Write-Host "SharePoint site data exported." -ForegroundColor Green

    $OneDriveSitesArray | Export-Csv (Join-Path $csvFolderPath "OneDriveSites.csv") -NoTypeInformation
    Write-Host "OneDrive site data exported." -ForegroundColor Green
    #endregion
}
catch {
    Write-Log -Level Error -Message "2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1 stopped due to an unexpected error: $($_.Exception.Message)"
    throw
}

# All CSV outputs are in: $csvFolderPath
# The warning/error log for this run is at: $logFilePath
#
# The connection is left open intentionally so further PnP cmdlets can
# be run against it. Disconnect when finished:
#   Disconnect-PnPOnline
