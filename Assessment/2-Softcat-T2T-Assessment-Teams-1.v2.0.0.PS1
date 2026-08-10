#Requires -Modules MicrosoftTeams

<#
.SYNOPSIS
    Connects to Microsoft Teams for the Softcat T2T assessment and
    exports team, channel and voice configuration data.

.DESCRIPTION
    Connects to Microsoft Teams using the same certificate-secured
    Azure AD app created by 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1,
    then retrieves and exports (to a "CSV Files" subfolder):
      - Microsoft Teams (with channel counts) -> Teams.csv
      - Team channels                          -> TeamChannels.csv
      - Voice routing policies                 -> TeamsVoiceRoutingPolicies.csv
      - PSTN usage                              -> TeamsPstnUsage.csv
      - Tenant dial plans                       -> TeamsDialPlans.csv

    One of the "separate scripts" referenced by
    1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 - only connects
    to Microsoft Teams (not Graph/Exchange Online), since that's all
    this needs.

    Resolves AppId/TenantId/CertificateThumbprint from (in order of
    preference): explicit parameters, a ConnectionDetails.json file, or
    manual entry if neither is available. Like Power BI/Power Apps (and
    unlike Exchange Online), Connect-MicrosoftTeams accepts the
    TenantId GUID directly, so no separate domain resolution is needed
    here.

    WARNING - Teams PowerShell's application-based (certificate) auth
    has been rolled out gradually, cmdlet by cmdlet, and Microsoft's
    own docs note coverage is still expanding. Get-Team and
    Get-TeamChannel are confirmed supported. The voice configuration
    section below uses several *-Cs cmdlets
    (Get-CsOnlineVoiceRoutingPolicy, Get-CsOnlinePstnUsage,
    Get-CsTenantDialPlan) - these are the category most likely to need
    the "Teams Administrator" directory role assigned to the app's
    service principal under app-only auth. If any of them fail with an
    access-denied error, that role assignment is the likely fix (same
    pattern as the Exchange Administrator role fix for Exchange
    Online).

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

.PARAMETER OutputPath
    Base folder for outputs. "CSV Files" and "Logs" subfolders are
    created under this (if they don't already exist). Defaults to the
    same folder as this script.

.EXAMPLE
    .\2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1

    Runs interactively, discovering a *-ConnectionDetails.json file
    automatically and writing output under the script's folder.

.EXAMPLE
    .\2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1 -AppId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -OutputPath C:\Assessments\Contoso

    Runs non-interactively with explicit connection details, writing
    output under C:\Assessments\Contoso.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    RUN ORDER: this script starts with "2-Softcat-T2T-Assessment-", so
    it can be run in any order relative to the other 2-Softcat-T2T-
    Assessment-* scripts, once 0-Softcat-T2T-Assessment-Create-App and
    1-Softcat-T2T-Assessment-Initial-Script have both been run.

    Prerequisites:
      Install-Module -Name MicrosoftTeams -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [string]$AppId,

    [string]$TenantId,

    [string]$CertificateThumbprint,

    [string]$ConnectionDetailsPath,

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
$logFilePath = Join-Path $logsFolderPath "Get-TeamsData-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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

#region Connect to Microsoft Teams
try {
    Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
    Connect-MicrosoftTeams -ApplicationId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint | Out-Null
    Write-Host "Connected to Microsoft Teams." -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to connect to Microsoft Teams: $($_.Exception.Message). If this is an access-denied error rather than a connection failure, the app's service principal may need the 'Teams Administrator' directory role assigned - see header notes."
    throw
}
#endregion

try {
    #region Teams and channels
    Write-Host "Getting Microsoft Teams..." -ForegroundColor Cyan
    $Teams = Get-Team | Sort-Object DisplayName
    $TeamsArray = @()
    $TeamChannelArray = @()
    $teamCount = $Teams.Count
    $teamIndex = 0

    Foreach ($Team in $Teams) {
        $teamIndex++
        Write-Progress -Id 1 -Activity "Processing Microsoft Teams and channels" -Status "$teamIndex of $teamCount - $($Team.DisplayName)" -PercentComplete (($teamIndex / $teamCount) * 100)
        try {
            $Channels = Get-TeamChannel -GroupId $Team.GroupId

            $TeamsArray += [PSCustomObject]@{
                GroupId      = $Team.GroupId
                DisplayName  = $Team.DisplayName
                Description  = $Team.Description
                Visibility   = $Team.Visibility
                MailNickName = $Team.MailNickName
                Archived     = $Team.Archived
                Channels     = $Channels.Count
            }

            Foreach ($Channel in $Channels) {
                $TeamChannelArray += [PSCustomObject]@{
                    GroupId            = $Team.GroupId
                    TeamDisplayName    = $Team.DisplayName
                    ChannelID          = $Channel.Id
                    ChannelDisplayName = $Channel.DisplayName
                    Description        = $Channel.Description
                    MembershipType     = $Channel.MembershipType
                }
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process team '$($Team.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 1 -Activity "Processing Microsoft Teams and channels" -Completed

    $TeamsArray | Export-Csv (Join-Path $csvFolderPath "Teams.csv") -NoTypeInformation
    Write-Host "Teams data exported." -ForegroundColor Green

    $TeamChannelArray | Export-Csv (Join-Path $csvFolderPath "TeamChannels.csv") -NoTypeInformation
    Write-Host "Channel data exported." -ForegroundColor Green
    #endregion

    Write-Host "Getting voice routing policies..." -ForegroundColor Cyan
    try {
        $voiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy
        $voiceRoutingPolicies | Export-Csv (Join-Path $csvFolderPath "TeamsVoiceRoutingPolicies.csv") -NoTypeInformation
        Write-Host "Voice routing policy data exported." -ForegroundColor Green
    }
    catch {
        Write-Log -Level Error -Message "Failed to retrieve voice routing policies: $($_.Exception.Message)"
    }

    Write-Host "Getting PSTN usage..." -ForegroundColor Cyan
    try {
        $pstnUsages = Get-CsOnlinePstnUsage
        $pstnUsages | Export-Csv (Join-Path $csvFolderPath "TeamsPstnUsage.csv") -NoTypeInformation
        Write-Host "PSTN usage data exported." -ForegroundColor Green
    }
    catch {
        Write-Log -Level Error -Message "Failed to retrieve PSTN usage: $($_.Exception.Message)"
    }

    Write-Host "Getting tenant dial plans..." -ForegroundColor Cyan
    try {
        $dialPlans = Get-CsTenantDialPlan
        $dialPlans | Export-Csv (Join-Path $csvFolderPath "TeamsDialPlans.csv") -NoTypeInformation
        Write-Host "Dial plan data exported." -ForegroundColor Green
    }
    catch {
        Write-Log -Level Error -Message "Failed to retrieve tenant dial plans: $($_.Exception.Message)"
    }
    #endregion
}
catch {
    Write-Log -Level Error -Message "2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1 stopped due to an unexpected error: $($_.Exception.Message)"
    throw
}

# All CSV outputs are in: $csvFolderPath
# The warning/error log for this run is at: $logFilePath
#
# The connection is left open intentionally so further Teams cmdlets
# can be run against it. Disconnect when finished:
#   Disconnect-MicrosoftTeams
