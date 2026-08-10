#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.DirectoryManagement

<#
.SYNOPSIS
    Creates the Entra ID app registration used by the M365 Migration Summary report.

.DESCRIPTION
    Creates an Azure AD (Entra ID) App Registration secured by a self-signed
    certificate and grants it the application permissions needed by
    New-M365MigrationSummaryREPORT-APPv1.1.ps1:

    Microsoft Graph (application permissions):
      - Application.Read.All
      - AuditLog.Read.All
      - BitLockerKey.ReadBasic.All
      - ChannelMessage.Read.All
      - Device.Read.All
      - Directory.Read.All
      - Files.Read.All
      - Group.Read.All
      - OrgContact.Read.All
      - Organization.Read.All
      - Policy.Read.All
      - Reports.Read.All
      - Sites.Read.All
      - User.Read.All

    Exchange Online (application permission):
      - Exchange.ManageAsApp
      Paired with the "Exchange Administrator" Azure AD directory role,
      which is also assigned to the service principal automatically -
      app-only connections to Exchange Online fail without it, even with
      Exchange.ManageAsApp granted.

    Microsoft Teams:
      No separate API permissions needed - Connect-MicrosoftTeams and
      non-Cs cmdlets (Get-Team, etc.) work with the Graph permissions
      above. The "Teams Administrator" Azure AD directory role is
      assigned to the service principal automatically though, since the
      *-Cs cmdlets (Get-CsOnlineUser, Get-CsPhoneNumberAssignment,
      Get-CsOnlineVoiceRoutingPolicy, Get-CsOnlinePstnUsage,
      Get-CsTenantDialPlan, etc.) fail with "Access Denied" without it.

    SharePoint Online (application permissions):
      - Sites.FullControl.All
      - Sites.Read.All
      - User.ReadWrite.All
      This is a SEPARATE resource from Microsoft Graph - Graph's own
      Sites.Read.All/User.Read.All (already granted above) do not cover
      PnP PowerShell or other SharePoint CSOM/REST-based tools, which
      need these SharePoint-specific permissions of the same names
      instead. Sites.FullControl.All and User.ReadWrite.All
      are write-capable - broader than the read-only philosophy every
      other permission on this app follows - included because PnP can 
      fail without them.
      Paired with the "SharePoint Administrator" Azure AD directory
      role, which is also assigned to the service principal
      automatically - tenant-admin scoped PnP cmdlets like
      Get-PnPTenantSite fail with "Attempted to perform an unauthorized
      operation" without it, even with the permissions above granted.

    Outputs the AppId (ClientId), TenantId and certificate thumbprint
    required to connect from the reporting script, and writes them to a
    ConnectionDetails.json file alongside the certificate - this app is
    created here by the tenant owner, but New-M365MigrationSummaryREPORT-
    APPv1.1.ps1 may be run separately, off-site, by a consultant on a
    different machine, so this script hands over details rather than
    running the reporting script itself.

.PARAMETER CertValidityYears
    Number of years the self-signed certificate remains valid for. Defaults to 2.

.PARAMETER CertExportPath
    Folder to export the certificate (and .pfx, if one is exported) to.
    Defaults to the script's own folder.

.PARAMETER NonExportableKey
    Makes the private key non-exportable (locked to this machine's
    certificate store - no .pfx can ever be created from it). If
    the reporting script is going to be run by someone else on a different
    machine (e.g. an off-site consultant), the private key MUST be
    exportable so it can be handed to them as a .pfx - a non-exportable key
    means only this machine can ever authenticate as this app. Only choose
    this if the tenant owner will also be the one running the reporting
    script, on this same machine. If this switch isn't given, the script
    asks interactively instead of silently defaulting to exportable.

.PARAMETER ExportPfx
    Forces a password-protected .pfx (private key) to be exported without
    asking. If neither this nor -NonExportableKey is given, the script asks
    interactively whether to export a .pfx.

.EXAMPLE
    .\New-M365MigrationSummary-CREATEAPPv1.8.ps1
    Runs interactively, asking whether the private key should be
    non-exportable and whether to export a .pfx.

.EXAMPLE
    .\New-M365MigrationSummary-CREATEAPPv1.7.ps1 -ExportPfx -CertExportPath C:\Handover
    Exports a password-protected .pfx (prompting for the password) alongside
    the .cer and ConnectionDetails.json to C:\Handover, ready to hand over
    to an off-site consultant.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    Requires: an account that is Global Administrator, OR Application
    Administrator + Privileged Role Administrator (the latter is needed to
    assign the Exchange Administrator, Teams Administrator and SharePoint
    Administrator roles).

    Prerequisites:
      Install-Module Microsoft.Graph -Scope CurrentUser
      Must be run on Windows PowerShell / PowerShell 7 on Windows (uses the
      PKI module's New-SelfSignedCertificate, which is not available on
      macOS/Linux PowerShell).

    Disclaimer:
      Provided as-is, with no warranty of any kind. This script creates a
      highly privileged app registration - it grants tenant-wide read
      access (plus write access to SharePoint sites and user profiles) and
      assigns the Exchange Administrator, Teams Administrator and
      SharePoint Administrator directory roles to its service principal.
      Review the permissions and role assignments above and get the
      customer's sign-off before running this against a production tenant.
      Keep the exported certificate/.pfx secure - anyone holding it can
      authenticate as this app.
#>

[CmdletBinding()]
param(
    [int]$CertValidityYears = 2,

    [string]$CertExportPath,

    [switch]$NonExportableKey,

    [switch]$ExportPfx
)

$ErrorActionPreference = "Stop"

#region Disclaimer
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host " DISCLAIMER" -ForegroundColor Yellow
Write-Host " Provided as-is, with no warranty of any kind. This script" -ForegroundColor Yellow
Write-Host " creates a highly privileged app registration - it grants" -ForegroundColor Yellow
Write-Host " tenant-wide read access (plus write access to SharePoint" -ForegroundColor Yellow
Write-Host " sites and user profiles) and assigns the Exchange" -ForegroundColor Yellow
Write-Host " Administrator, Teams Administrator and SharePoint" -ForegroundColor Yellow
Write-Host " Administrator directory roles. Get sign-off before running" -ForegroundColor Yellow
Write-Host " this against a production tenant." -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow

$answer = Read-Host "Continue? (C = continue, Q = quit)"
if ($answer -notmatch '^[Cc]') {
    Write-Host "Quit - no changes were made." -ForegroundColor Yellow
    exit
}
#endregion

# Fixed naming - not exposed as parameters, so these can't be overridden
# from the command line.
$AppDisplayName = "Softcat-T2TMigration-Summary-Report"

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($CertExportPath)) {
    $CertExportPath = $scriptDir
}

#region Connect to Microsoft Graph
$graphScopes = @(
    "Application.ReadWrite.All",
    "AppRoleAssignment.ReadWrite.All",
    "Directory.ReadWrite.All",
    "RoleManagement.ReadWrite.Directory"
)

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes $graphScopes | Out-Null
$tenantId = (Get-MgContext).TenantId
#endregion

#region Ensure the app display name is unique in this tenant
# If an app registration called "Softcat-T2TMigration-Summary-
# Report" already exists (e.g. this script was run before against this
# tenant), creating a second one with the identical name would leave
# two indistinguishable apps in Entra ID - a real risk of the wrong one
# being used later. Instead, keep appending a number until the name is
# free, and flag it clearly so the consultant knows why.
$baseAppDisplayName = $AppDisplayName
$suffix = 1
while (Get-MgApplication -Filter "displayName eq '$AppDisplayName'") {
    $suffix++
    $AppDisplayName = "$baseAppDisplayName-$suffix"
}
if ($AppDisplayName -ne $baseAppDisplayName) {
    Write-Warning "An app registration named '$baseAppDisplayName' already exists in this tenant, so this one will be created as '$AppDisplayName' instead to avoid a naming collision. Consultant: use '$AppDisplayName' (not the default name) when locating the app in Entra ID and in the certificate/connection files handed to you."
}

$CertSubject = "CN=$AppDisplayName"
#endregion

try {
    #region Create self-signed certificate
    Write-Host "Creating self-signed certificate..." -ForegroundColor Cyan

    # Decide the key export policy. -NonExportableKey skips the question;
    # otherwise ask interactively rather than silently defaulting.
    if ($NonExportableKey) {
        $isNonExportable = $true
    }
    else {
        $answer = Read-Host "Make the private key non-exportable? Only choose this if YOU will run the reporting script on this same machine - it rules out ever handing the key to a consultant elsewhere. (Y/N)"
        $isNonExportable = $answer -match '^[Yy]'
    }
    $keyExportPolicy = if ($isNonExportable) { "NonExportable" } else { "Exportable" }

    if ($isNonExportable -and $ExportPfx) {
        Write-Warning "-ExportPfx was specified, but the private key was set to non-exportable - the private key can't be exported once it's non-exportable, so no .pfx will be created."
    }

    $cert = New-SelfSignedCertificate `
        -Subject $CertSubject `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy $keyExportPolicy `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($CertValidityYears)

    $certThumbprint = $cert.Thumbprint
    $cerPath = Join-Path $CertExportPath "$AppDisplayName.cer"
    Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
    Write-Host "Certificate created. Thumbprint: $certThumbprint" -ForegroundColor Green

    # Decide whether to export a .pfx. A non-exportable key rules it out
    # entirely; -ExportPfx forces it without asking; otherwise ask.
    $shouldExportPfx = $false
    if ($isNonExportable) {
        $shouldExportPfx = $false
    }
    elseif ($ExportPfx) {
        $shouldExportPfx = $true
    }
    else {
        $answer = Read-Host "Export a password-protected .pfx (private key) to hand to the consultant who'll run the reporting script? (Y/N)"
        $shouldExportPfx = $answer -match '^[Yy]'
    }

    if ($shouldExportPfx) {
        $pfxPath = Join-Path $CertExportPath "$AppDisplayName.pfx"
        $pfxPassword = Read-Host -AsSecureString "Enter a password to protect the exported .pfx"
        Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $pfxPassword | Out-Null
        Write-Host "Private key exported to: $pfxPath (keep this secure)" -ForegroundColor Yellow
    }
    #endregion

    #region Create the App Registration and its Service Principal
    Write-Host "Creating App Registration '$AppDisplayName'..." -ForegroundColor Cyan

    $certBytes = [System.IO.File]::ReadAllBytes($cerPath)
    $keyCredential = @{
        Type  = "AsymmetricX509Cert"
        Usage = "Verify"
        Key   = $certBytes
    }

    $app = New-MgApplication -DisplayName $AppDisplayName -SignInAudience "AzureADMyOrg" -KeyCredentials @($keyCredential)
    $appId = $app.AppId
    $appObjectId = $app.Id
    Write-Host "App created. AppId: $appId" -ForegroundColor Green

    $sp = New-MgServicePrincipal -AppId $appId
    $spObjectId = $sp.Id
    #endregion

    #region Resolve and assign Microsoft Graph application permissions
    # Permission IDs are looked up dynamically by name against the Graph
    # service principal's AppRoles, rather than hard-coded, so this
    # works regardless of exact GUID values and is resilient to the
    # casing of the names below.
    $graphResourceAppId = "00000003-0000-0000-c000-000000000000"
    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphResourceAppId'"

    $graphPermissionNames = @(
        "Application.Read.All",
        "AuditLog.Read.All",
        "BitLockerKey.ReadBasic.All",
        "ChannelMessage.Read.All",
        "Device.Read.All",
        "Directory.Read.All",
        "Files.Read.All",
        "Group.Read.All",
        "OrgContact.Read.All",
        "Organization.Read.All",
        "Policy.Read.All",
        "Reports.Read.All",
        "Sites.Read.All",
        "User.Read.All"
    )

    $graphResourceAccess = @()
    foreach ($permName in $graphPermissionNames) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $permName -and $_.AllowedMemberTypes -contains "Application" }
        if (-not $role) {
            Write-Warning "Couldn't find a Graph application permission named '$permName' - skipping it. Double-check the name in Entra ID > App registrations > API permissions."
            continue
        }
        $graphResourceAccess += @{ Id = $role.Id; Type = "Role" }
    }

    Update-MgApplication -ApplicationId $appObjectId -RequiredResourceAccess @(
        @{ ResourceAppId = $graphResourceAppId; ResourceAccess = $graphResourceAccess }
    )

    foreach ($permName in $graphPermissionNames) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $permName -and $_.AllowedMemberTypes -contains "Application" }
        if ($role) {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spObjectId -PrincipalId $spObjectId -ResourceId $graphSp.Id -AppRoleId $role.Id | Out-Null
            Write-Host "Granted Graph permission: $permName" -ForegroundColor Green
        }
    }
    #endregion

    #region Resolve and assign the Exchange Online application permission
    $exoResourceAppId = "00000002-0000-0ff1-ce00-000000000000"   # Office 365 Exchange Online
    $exoSp = Get-MgServicePrincipal -Filter "appId eq '$exoResourceAppId'"

    if (-not $exoSp) {
        Write-Warning "The'Office 365 Exchange Online' service principal wasn't found in this tenant, so Exchange.ManageAsApp couldn't be granted. This usually resolves itself after Connect-ExchangeOnline has been run interactively once in the tenant - try that, then re-run this script."
    }
    else {
        $exoPermName = "Exchange.ManageAsApp"
        $exoRole = $exoSp.AppRoles | Where-Object { $_.Value -eq $exoPermName -and $_.AllowedMemberTypes -contains "Application" }

        if (-not $exoRole) {
            Write-Warning "Couldn't find the '$exoPermName' application permission on the Exchange Online service principal - skipping it."
        }
        else {
            $currentAccess = (Get-MgApplication -ApplicationId $appObjectId).RequiredResourceAccess
            $currentAccess += @{ ResourceAppId = $exoResourceAppId; ResourceAccess = @(@{ Id = $exoRole.Id; Type = "Role" }) }
            Update-MgApplication -ApplicationId $appObjectId -RequiredResourceAccess $currentAccess

            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spObjectId -PrincipalId $spObjectId -ResourceId $exoSp.Id -AppRoleId $exoRole.Id | Out-Null
            Write-Host "Granted Exchange Online permission: $exoPermName" -ForegroundColor Green
        }
    }
    #endregion

    #region Resolve and assign the SharePoint Online application permissions
    # SharePoint's own API is a separate resource from Microsoft Graph -
    # Graph's Sites.Read.All/User.Read.All (already granted above) do
    # NOT cover PnP PowerShell or other SharePoint CSOM/REST-based
    # tools, which need these SharePoint-specific permissions of the
    # same names instead.

    $spoResourceAppId = "00000003-0000-0ff1-ce00-000000000000"   # Office 365 SharePoint Online
    $spoSp = Get-MgServicePrincipal -Filter "appId eq '$spoResourceAppId'"

    if (-not $spoSp) {
        Write-Warning "The 'Office 365 SharePoint Online' service principal wasn't found in this tenant, so its permissions couldn't be granted. Try re-running this script, or add them manually in Entra ID > App registrations > API permissions."
    }
    else {
        $spoPermissionNames = @(
            "Sites.FullControl.All",
            "Sites.Read.All",
            "User.ReadWrite.All"
        )

        $spoResourceAccess = @()
        foreach ($permName in $spoPermissionNames) {
            $role = $spoSp.AppRoles | Where-Object { $_.Value -eq $permName -and $_.AllowedMemberTypes -contains "Application" }
            if (-not $role) {
                Write-Warning "Couldn't find a SharePoint application permission named '$permName' - skipping it."
                continue
            }
            $spoResourceAccess += @{ Id = $role.Id; Type = "Role" }
        }

        $currentAccess = (Get-MgApplication -ApplicationId $appObjectId).RequiredResourceAccess
        $currentAccess += @{ ResourceAppId = $spoResourceAppId; ResourceAccess = $spoResourceAccess }
        Update-MgApplication -ApplicationId $appObjectId -RequiredResourceAccess $currentAccess

        foreach ($permName in $spoPermissionNames) {
            $role = $spoSp.AppRoles | Where-Object { $_.Value -eq $permName -and $_.AllowedMemberTypes -contains "Application" }
            if ($role) {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spObjectId -PrincipalId $spObjectId -ResourceId $spoSp.Id -AppRoleId $role.Id | Out-Null
                Write-Host "Granted SharePoint Online permission: $permName" -ForegroundColor Green
            }
        }
    }
    #endregion

    #region Assign the Exchange Administrator directory role
    # Confirmed required: app-only connections to Exchange Online check
    # that the service principal holds a supported Azure AD directory
    # role - Exchange.ManageAsApp on its own isn't enough. Without this,
    # Connect-ExchangeOnline fails with "The role assigned to application
    # ... isn't supported in this scenario."
    $roleName = "Exchange Administrator"
    $directoryRole = Get-MgDirectoryRole -Filter "displayName eq '$roleName'"
    if (-not $directoryRole) {
        $template = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $roleName }
        $directoryRole = New-MgDirectoryRole -DisplayName $template.DisplayName -RoleTemplateId $template.Id
    }
    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $directoryRole.Id -BodyParameter @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$spObjectId"
    }
    Write-Host "Assigned '$roleName' directory role to the service principal." -ForegroundColor Green
    Write-Host "Note: role assignments can take a few minutes to propagate - if Connect-ExchangeOnline fails immediately after this, wait 5-10 minutes and retry." -ForegroundColor Yellow
    #endregion

    #region Assign the Teams Administrator directory role
    # Confirmed required: app-only connections to Microsoft Teams check
    # that the service principal holds a supported Azure AD directory
    # role for the *-Cs cmdlets (Get-CsOnlineUser,
    # Get-CsPhoneNumberAssignment, Get-CsOnlineVoiceRoutingPolicy,
    # Get-CsOnlinePstnUsage, Get-CsTenantDialPlan, etc.) - the Graph API
    # permissions granted above aren't enough on their own. Without
    # this, those cmdlets fail with "Access Denied" even though
    # Connect-MicrosoftTeams itself succeeds and non-Cs cmdlets like
    # Get-Team work fine.
    $roleName = "Teams Administrator"
    $directoryRole = Get-MgDirectoryRole -Filter "displayName eq '$roleName'"
    if (-not $directoryRole) {
        $template = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $roleName }
        $directoryRole = New-MgDirectoryRole -DisplayName $template.DisplayName -RoleTemplateId $template.Id
    }
    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $directoryRole.Id -BodyParameter @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$spObjectId"
    }
    Write-Host "Assigned '$roleName' directory role to the service principal." -ForegroundColor Green
    Write-Host "Note: role assignments can take a few minutes to propagate - if *-Cs Teams cmdlets fail immediately after this, wait 5-10 minutes and retry." -ForegroundColor Yellow
    #endregion

    #region Assign the SharePoint Administrator directory role
    # Confirmed required: Get-PnPTenantSite and other tenant-admin
    # scoped PnP cmdlets call the SharePoint Admin API, which is a step
    # up from ordinary site content access - the SharePoint Online
    # Sites.Read.All permission granted above covers reading site
    # content, but tenant-wide admin operations also need the service
    # principal to hold this role. Without it, those cmdlets fail with
    # "Attempted to perform an unauthorized operation" even though
    # Connect-PnPOnline itself succeeds.
    $roleName = "SharePoint Administrator"
    $directoryRole = Get-MgDirectoryRole -Filter "displayName eq '$roleName'"
    if (-not $directoryRole) {
        $template = Get-MgDirectoryRoleTemplate | Where-Object { $_.DisplayName -eq $roleName }
        $directoryRole = New-MgDirectoryRole -DisplayName $template.DisplayName -RoleTemplateId $template.Id
    }
    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $directoryRole.Id -BodyParameter @{
        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$spObjectId"
    }
    Write-Host "Assigned '$roleName' directory role to the service principal." -ForegroundColor Green
    Write-Host "Note: role assignments can take a few minutes to propagate - if Get-PnPTenantSite fails immediately after this, wait 5-10 minutes and retry." -ForegroundColor Yellow
    #endregion

    #region Summary
    Write-Host "`n============================================================" -ForegroundColor Yellow
    Write-Host " App registration complete - store these details securely" -ForegroundColor Yellow
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "AppId (ClientId) : $appId"
    Write-Host "TenantId         : $tenantId"
    Write-Host "CertThumbprint   : $certThumbprint"
    Write-Host "Cert (.cer) path : $cerPath"
    Write-Host "`nUse these values in New-M365MigrationSummaryREPORT-APPv1.1.ps1"
    Write-Host "The private key currently exists only in Cert:\CurrentUser\My on this machine."
    Write-Host "Use -ExportPfx if the reporting script needs to run elsewhere."
    #endregion

    #region Export connection details for handover to the consultant
    # This app is created here by the tenant owner, but
    # New-M365MigrationSummaryREPORT-APPv1.1.ps1 may be run separately,
    # off-site, by a consultant on a different machine - so connection
    # details are written to a file to hand over, rather than this
    # script running the reporting script itself.
    $connectionDetails = [ordered]@{
        AppId                 = $appId
        TenantId              = $tenantId
        CertificateThumbprint = $certThumbprint
        Notes                 = "Give this file, plus the .pfx (if one was exported), to the consultant. They import the .pfx into their own certificate store, then run New-M365MigrationSummaryREPORT-APPv1.1.ps1 -AppId <AppId> -TenantId <TenantId> -CertificateThumbprint <CertificateThumbprint>."
    }
    $connectionDetailsPath = Join-Path $CertExportPath "$AppDisplayName-ConnectionDetails.json"
    $connectionDetails | ConvertTo-Json | Out-File -FilePath $connectionDetailsPath -Encoding UTF8
    Write-Host "Connection details for the consultant written to: $connectionDetailsPath" -ForegroundColor Green

    if (-not $shouldExportPfx) {
        Write-Host "Note: no .pfx was exported, so the consultant has no way to authenticate as this app from their own machine yet. Re-run with -ExportPfx (or answer Y to the earlier prompt) if the reporting script will run off-site." -ForegroundColor Yellow
    }
    #endregion
}
finally {
    Disconnect-MgGraph | Out-Null
}
