#Requires -Modules Microsoft.Graph

<#
.SYNOPSIS
    Removes the Entra ID app registration created for the Softcat T2T
    assessment. Optional and out-of-band - not part of the 0/1/2/9
    sequence.

.DESCRIPTION
    RUN ORDER: optional and out-of-band - not part of the
    0 -> 1 -> 2 -> 9 sequence. Run this only when you want to decommission
    the app created by 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1,
    typically at the end of an engagement once all the assessment scripts
    have been run and their CSV output has been collected.

    Removes the Azure AD (Entra ID) App Registration and Service
    Principal created by 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1
    (fixed display name "Softcat-T2TMigration-Summary-Report"):
      - Revokes every Microsoft Graph / Exchange Online / SharePoint
        Online application permission granted to it
      - Removes the Exchange Administrator, Teams Administrator and
        SharePoint Administrator directory role assignments held by its
        service principal - these can't be removed individually first,
        since Entra role assignments are tied to the principal; they go
        away automatically once the service principal itself is deleted
      - Optionally removes the local self-signed certificate from this
        machine's Cert:\CurrentUser\My store
      - Optionally deletes the exported .cer/.pfx/*-ConnectionDetails.json
        files left behind by the create-app script

    WARNING - deleting the app is a soft delete: Entra ID keeps it in
    "Deleted applications" for 30 days by default, recoverable by a
    Global Administrator in that window (Entra admin center > Identity >
    Applications > Deleted applications, or
    Restore-MgDirectoryDeletedItem). Pass -PurgeImmediately to instead
    permanently purge it from the recycle bin straight away - only do
    this if you're certain the app won't be needed again, since it can't
    be undone.

    Resolves which app to remove from (in order of preference): an
    explicit -AppId parameter, a *-ConnectionDetails.json file, or the
    fixed display name the create-app script always uses.

    Requires the same admin role as
    0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1: Global Administrator,
    OR Application Administrator + Privileged Role Administrator (the
    latter is needed to see/remove the Exchange Administrator, Teams
    Administrator and SharePoint Administrator role assignments).

.PARAMETER AppId
    AppId (ClientId) of the app to remove. If not supplied, resolved
    from -ConnectionDetailsPath, or falls back to searching by the
    fixed display name "Softcat-T2TMigration-Summary-Report".

.PARAMETER ConnectionDetailsPath
    Path to a *-ConnectionDetails.json file exported by
    0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1. If not supplied,
    the script looks for one automatically in its own folder.

.PARAMETER CertExportPath
    Folder to look in for exported .cer/.pfx/ConnectionDetails.json
    files (only used if file removal is confirmed). Defaults to the
    same folder as this script.

.PARAMETER RemoveCertificate
    Also removes the local self-signed certificate from this
    machine's Cert:\CurrentUser\My store. If not given, the script
    asks interactively.

.PARAMETER RemoveExportedFiles
    Also deletes the exported .cer/.pfx/*-ConnectionDetails.json
    files from $CertExportPath. If not given, the script asks
    interactively.

.PARAMETER PurgeImmediately
    Permanently purges the app from Entra ID's "Deleted applications"
    recycle bin immediately, instead of leaving it there for the
    default 30-day recovery window. Not asked interactively - this
    is opt-in only, since it can't be undone.

.PARAMETER Force
    Skips the final "are you sure" confirmation before removing the
    app registration and service principal.

.EXAMPLE
    .\0z-Softcat-T2T-Assessment-Remove-App.v2.0.0.PS1

    Runs interactively, resolving the app from a ConnectionDetails.json
    file or the fixed display name, and asking whether to also remove
    the certificate and exported files.

.EXAMPLE
    .\0z-Softcat-T2T-Assessment-Remove-App.v2.0.0.PS1 -AppId $appId -RemoveCertificate -RemoveExportedFiles -Force

    Removes the app, its certificate and exported files non-interactively.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    Prerequisites:
      Install-Module Microsoft.Graph -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [string]$AppId,

    [string]$ConnectionDetailsPath,

    [string]$CertExportPath,

    [switch]$RemoveCertificate,

    [switch]$RemoveExportedFiles,

    [switch]$PurgeImmediately,

    [switch]$Force
)

$ErrorActionPreference = "Stop"

# Fixed naming - must match 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1
# exactly, since it's used as the fallback lookup if no AppId is
# resolved another way.
$AppDisplayName = "Softcat-T2TMigration-Summary-Report"
$CertSubject = "CN=$AppDisplayName"

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($CertExportPath)) {
    $CertExportPath = $scriptDir
}

#region Logging
$logsFolderPath = Join-Path $scriptDir "Logs"
if (-not (Test-Path $logsFolderPath)) {
    New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
}
$logFilePath = Join-Path $logsFolderPath "0z-Softcat-T2T-Assessment-Remove-App-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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

#region Resolve AppId from a ConnectionDetails.json file, if not given explicitly
if ([string]::IsNullOrWhiteSpace($AppId)) {
    if ([string]::IsNullOrWhiteSpace($ConnectionDetailsPath)) {
        $discovered = Get-ChildItem -Path $scriptDir -Filter "*-ConnectionDetails.json" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($discovered) {
            $useJsonAnswer = Read-Host "Found connection details file '$($discovered.Name)' - use its AppId? (Y/N)"
            if ($useJsonAnswer -match '^[Yy]') {
                $ConnectionDetailsPath = $discovered.FullName
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ConnectionDetailsPath) -and (Test-Path $ConnectionDetailsPath)) {
        $details = Get-Content -Path $ConnectionDetailsPath -Raw | ConvertFrom-Json
        $AppId = $details.AppId
        Write-Host "Loaded AppId from: $ConnectionDetailsPath" -ForegroundColor Green
    }
}
# If still blank, the app lookup below falls back to searching by the
# fixed display name instead.
#endregion

#region Connect to Microsoft Graph
$graphScopes = @(
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $graphScopes | Out-Null
#endregion

try {
    #region Find the app registration
    $app = $null
    if (-not [string]::IsNullOrWhiteSpace($AppId)) {
        $app = Get-MgApplication -Filter "appId eq '$AppId'" | Select-Object -First 1
        if (-not $app) {
            Write-Host "AppId '$AppId' didn't match any app registration - falling back to searching by the fixed display name '$AppDisplayName' instead." -ForegroundColor Yellow
        }
    }

    if (-not $app) {
        $appMatches = @(Get-MgApplication -Filter "displayName eq '$AppDisplayName'")
        if ($appMatches.Count -eq 0) {
            Write-Log -Level Error -Message "No app registration found with AppId '$AppId' or display name '$AppDisplayName'. Nothing to remove."
            throw "App registration not found."
        }
        elseif ($appMatches.Count -eq 1) {
            $app = $appMatches[0]
        }
        else {
            Write-Host "Multiple app registrations found named '$AppDisplayName':" -ForegroundColor Yellow
            $appMatches | Select-Object DisplayName, AppId, Id | Format-Table -AutoSize
            $chosenAppId = Read-Host "Enter the AppId of the one to remove"
            $app = $appMatches | Where-Object { $_.AppId -eq $chosenAppId }
            if (-not $app) {
                throw "AppId '$chosenAppId' didn't match any of the app registrations listed above."
            }
        }
    }

    $AppId = $app.AppId
    $appObjectId = $app.Id
    Write-Host "Found app registration: $($app.DisplayName) (AppId: $AppId)" -ForegroundColor Green
    #endregion

    #region Find the service principal and its directory role assignments
    $sp = Get-MgServicePrincipal -Filter "appId eq '$AppId'" | Select-Object -First 1
    $roleNames = @()
    if ($sp) {
        $memberOf = Get-MgServicePrincipalMemberOf -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
        $roleNames = $memberOf |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.directoryRole' } |
            ForEach-Object { $_.AdditionalProperties['displayName'] }
    }
    #endregion

    #region Show what will be removed, and confirm
    Write-Host "`n============================================================" -ForegroundColor Yellow
    Write-Host " About to remove" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "App display name  : $($app.DisplayName)"
    Write-Host "AppId              : $AppId"
    Write-Host "App object ID      : $appObjectId"
    if ($sp) {
        Write-Host "Service principal  : $($sp.Id)"
        Write-Host "Directory roles    : $(if ($roleNames) { $roleNames -join ', ' } else { '(none found)' })"
    }
    else {
        Write-Host "Service principal  : (not found - may already have been removed)"
    }

    if (-not $Force) {
        $confirm = Read-Host "`nThis will permanently remove the app registration and service principal above (recoverable from Entra ID's recycle bin for 30 days unless -PurgeImmediately was used). Continue? (Y/N)"
        if ($confirm -notmatch '^[Yy]') {
            Write-Host "Cancelled - nothing was removed." -ForegroundColor Yellow
            return
        }
    }
    #endregion

    #region Remove the service principal, then the app registration
    if ($sp) {
        Write-Host "Removing service principal..." -ForegroundColor Cyan
        Remove-MgServicePrincipal -ServicePrincipalId $sp.Id
        Write-Host "Service principal removed." -ForegroundColor Green
    }

    Write-Host "Removing app registration..." -ForegroundColor Cyan
    Remove-MgApplication -ApplicationId $appObjectId
    Write-Host "App registration removed." -ForegroundColor Green
    #endregion

    #region Optionally purge from the recycle bin
    if ($PurgeImmediately) {
        Write-Host "Purging app permanently from the Deleted applications recycle bin..." -ForegroundColor Cyan
        try {
            Remove-MgDirectoryDeletedItem -DirectoryObjectId $appObjectId
            Write-Host "App permanently purged - it can no longer be restored." -ForegroundColor Green
        }
        catch {
            Write-Log -Message "Failed to purge the app from the recycle bin: $($_.Exception.Message). It's still soft-deleted and recoverable for 30 days; you can purge it manually later from the Entra admin center."
        }
    }
    else {
        Write-Host "The app is soft-deleted and recoverable from Entra ID's 'Deleted applications' for 30 days. Re-run with -PurgeImmediately to purge it now instead." -ForegroundColor Yellow
    }
    #endregion
}
catch {
    Write-Log -Level Error -Message "0z-Softcat-T2T-Assessment-Remove-App.v2.0.0.PS1 stopped due to an unexpected error: $($_.Exception.Message)"
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

#region Optionally remove the local certificate
if ($RemoveCertificate) {
    $removeCert = $true
}
else {
    $answer = Read-Host "`nAlso remove the local self-signed certificate ('$CertSubject') from Cert:\CurrentUser\My on this machine? (Y/N)"
    $removeCert = $answer -match '^[Yy]'
}

if ($removeCert) {
    $certsToRemove = Get-ChildItem -Path "Cert:\CurrentUser\My" -ErrorAction SilentlyContinue | Where-Object { $_.Subject -eq $CertSubject }
    if (-not $certsToRemove) {
        Write-Log -Message "No certificate found in Cert:\CurrentUser\My with subject '$CertSubject' - nothing to remove locally. (It may live on a different machine, e.g. if the assessment scripts were run by an off-site consultant.)"
    }
    else {
        foreach ($certToRemove in $certsToRemove) {
            try {
                Remove-Item -Path $certToRemove.PSPath -Force
                Write-Host "Removed certificate: $($certToRemove.Thumbprint)" -ForegroundColor Green
            }
            catch {
                Write-Log -Message "Failed to remove certificate '$($certToRemove.Thumbprint)': $($_.Exception.Message)"
            }
        }
    }
}
#endregion

#region Optionally remove exported files
if ($RemoveExportedFiles) {
    $removeFiles = $true
}
else {
    $answer = Read-Host "Also delete the exported .cer/.pfx/*-ConnectionDetails.json files in '$CertExportPath'? (Y/N)"
    $removeFiles = $answer -match '^[Yy]'
}

if ($removeFiles) {
    $filesToRemove = Get-ChildItem -Path $CertExportPath -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -eq "$AppDisplayName.cer" -or
        $_.Name -eq "$AppDisplayName.pfx" -or
        $_.Name -eq "$AppDisplayName-ConnectionDetails.json"
    }
    if (-not $filesToRemove) {
        Write-Log -Message "No exported .cer/.pfx/ConnectionDetails.json files found in '$CertExportPath' to remove."
    }
    else {
        foreach ($fileToRemove in $filesToRemove) {
            try {
                Remove-Item -Path $fileToRemove.FullName -Force
                Write-Host "Deleted file: $($fileToRemove.Name)" -ForegroundColor Green
            }
            catch {
                Write-Log -Message "Failed to delete file '$($fileToRemove.Name)': $($_.Exception.Message)"
            }
        }
    }
}
#endregion

#region Summary
Write-Host "`n============================================================" -ForegroundColor Yellow
Write-Host " Removal complete" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "Removed app registration and service principal for: $AppDisplayName (AppId: $AppId)"
#endregion

# The warning/error log for this run is at: $logFilePath
