
<#
.SYNOPSIS
    Connects to Microsoft Graph and Exchange Online for the Softcat T2T
    assessment, then retrieves and exports tenant/domain context plus
    most of the Graph data needed for the migration summary report.

.DESCRIPTION
    Connects to Microsoft Graph and Exchange Online using the
    certificate-secured Azure AD app created by
    0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1, then retrieves and
    exports:

      To a JSON subfolder (for other scripts to import):
        - Tenant organisation details -> JSON\*-OrganizationDetails.json
        - All accepted domains        -> JSON\*-AcceptedDomains.json

      To a CSV Files subfolder:
        - Organisation details                        -> Org.csv
        - Entra ID domains                             -> AllMGDomains.csv
        - All Entra ID users                           -> AllMGUsers.csv
        - Per-user licences (friendly names)           -> LicenseData.csv
        - Per-user licences, one row per licence       -> LicenseOnePerLine.csv
        - Tenant-wide licence purchase/consumption     -> DetailedLicenseSummary.csv
        - MX records for each accepted domain          -> MXRecords.csv
        - SPF records for each accepted domain         -> SPFRecords.csv
        - DMARC records for each accepted domain       -> DMARCRecords.csv
        - Teams user activity (last 180 days)          -> TeamsUserActivity_Last180Days.csv
        - Security groups (with members/owners)        -> SecurityGroups.csv
        - Registered applications (with secrets/certs) -> RegisteredApplications.csv
        - BitLocker recovery keys                      -> BitlockerKeys.csv
        - Entra ID devices                             -> Devices.csv

    The *-ConnectionDetails.json file (from the 0-Softcat-T2T-Assessment-
    Create-App script) is the one exception - it stays alongside the
    certificate, not in the JSON subfolder, since it's read from there
    by design.

    This script has grown beyond just connecting and pulling tenant/
    domain context - it now also collects most of the Graph data for the
    migration summary report directly. Mailbox statistics and Exchange
    org-level data (connectors, transport rules, distribution groups,
    etc.) are still handled by the separate Get-MailboxStatistics and
    Get-EXOData scripts, which take the same AppId/TenantId/
    CertificateThumbprint connection details as their own parameters.

    Connects using an Azure App Registration (Client Credentials /
    certificate). Run 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1
    first if the app and certificate don't already exist.

    Designed to be run by a consultant who received, from the tenant
    owner, the certificate (.pfx) and a *-ConnectionDetails.json file:
      - Optionally imports the .pfx into this machine's certificate store
      - Resolves AppId/TenantId/CertificateThumbprint from (in order of
        preference): explicit parameters, a ConnectionDetails.json file,
        or manual entry if neither is available

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

.PARAMETER PfxPath
    Path to the .pfx certificate file. Only used if you answer Y to
    importing the certificate below.

.PARAMETER OutputPath
    Base folder for outputs. "JSON", "CSV Files" and "Logs"
    subfolders are created under this (if they don't already exist).
    Defaults to the same folder as this script.

.EXAMPLE
    .\1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1

    Runs interactively, discovering a *-ConnectionDetails.json file in
    the script's own folder (or prompting for connection details) and
    writing output under the script's folder.

.EXAMPLE
    .\1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 -AppId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -OutputPath C:\Assessments\Contoso

    Runs non-interactively with explicit connection details, writing
    output under C:\Assessments\Contoso.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    RUN ORDER: 0-Softcat-T2T-Assessment-Create-App.v2.0.0.PS1 must be
    run first, to create the app. Then run this script. After that,
    all scripts named 2-Softcat-T2T-Assessment-* can be run in any
    order.

    Prerequisites:
      Install-Module Microsoft.Graph -Scope CurrentUser
      Install-Module ExchangeOnlineManagement -Scope CurrentUser

    Certificate import and the MX/SPF/DMARC lookups require Windows
    (Import-PfxCertificate and Resolve-DnsName are not available on
    macOS/Linux PowerShell).
#>
[CmdletBinding()]
param(
    [string]$AppId,

    [string]$TenantId,

    [string]$CertificateThumbprint,

    [string]$ConnectionDetailsPath,

    [string]$PfxPath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $scriptDir
}

#region Ensure output folders exist
$jsonFolderPath = Join-Path $OutputPath "JSON"
if (-not (Test-Path $jsonFolderPath)) {
    New-Item -Path $jsonFolderPath -ItemType Directory | Out-Null
}

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
$logFilePath = Join-Path $logsFolderPath "New-M365MigrationSummaryREPORT-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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

#region Optionally import the certificate handed over by the tenant owner
$importAnswer = Read-Host "Import the certificate (.pfx) provided by the tenant owner into your certificate store now? (Y/N)"
if ($importAnswer -match '^[Yy]') {
    if ([string]::IsNullOrWhiteSpace($PfxPath)) {
        $PfxPath = Read-Host "Path to the .pfx file"
    }
    if (-not (Test-Path $PfxPath)) {
        throw "Couldn't find a .pfx file at '$PfxPath'."
    }

    $pfxPassword = Read-Host -AsSecureString "Password for the .pfx"
    $importedCert = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation "Cert:\CurrentUser\My" -Password $pfxPassword
    Write-Host "Certificate imported. Thumbprint: $($importedCert.Thumbprint)" -ForegroundColor Green

    if ([string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        $CertificateThumbprint = $importedCert.Thumbprint
    }
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

#region Connect to Microsoft Graph
try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    throw
}
#endregion

#region Retrieve tenant organisation details
# Get-MgOrganization also gives us the tenant's initial (.onmicrosoft.com)
# domain, needed for the Connect-ExchangeOnline -Organization value below,
# so this is fetched once and reused later for the Org.csv export too,
# rather than calling Get-MgOrganization a second time.
Write-Host "Retrieving organisation details..." -ForegroundColor Cyan
$organizationDetails = Get-MgOrganization
$initialDomain = ($organizationDetails.VerifiedDomains | Where-Object { $_.IsInitial }).Name

if (-not $initialDomain) {
    Write-Log -Message "WARNING: couldn't find a verified initial domain via Get-MgOrganization - falling back to TenantId for -Organization. If Connect-ExchangeOnline fails below, supply the tenant's .onmicrosoft.com domain manually instead."
    $initialDomain = $TenantId
}
Write-Host "Retrieved organisation details into `$organizationDetails." -ForegroundColor Green

$organizationDetailsPath = Join-Path $jsonFolderPath "$initialDomain-OrganizationDetails.json"
$organizationDetails | ConvertTo-Json -Depth 10 | Out-File -FilePath $organizationDetailsPath -Encoding UTF8
Write-Host "Organisation details exported to: $organizationDetailsPath" -ForegroundColor Green
#endregion

#region Connect to Exchange Online
try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertificateThumbprint -Organization $initialDomain -ShowBanner:$false
    Write-Host "Connected to Exchange Online." -ForegroundColor Green
}
catch {
    Write-Log -Level Error -Message "Failed to connect to Exchange Online: $($_.Exception.Message)"
    throw
}
#endregion

#region Retrieve accepted domains
Write-Host "Retrieving accepted domains..." -ForegroundColor Cyan
$acceptedDomains = Get-AcceptedDomain
Write-Host "Retrieved $($acceptedDomains.Count) accepted domains into `$acceptedDomains." -ForegroundColor Green

$acceptedDomainsPath = Join-Path $jsonFolderPath "$initialDomain-AcceptedDomains.json"
$acceptedDomains | ConvertTo-Json -Depth 10 | Out-File -FilePath $acceptedDomainsPath -Encoding UTF8
Write-Host "Accepted domains exported to: $acceptedDomainsPath" -ForegroundColor Green
#endregion

try {
    #region Organisation details (CSV)
    Write-Host "Exporting organisation details to CSV..." -ForegroundColor Cyan
    Write-Host "Connected to tenant: $($organizationDetails.DisplayName)" -ForegroundColor Green
    $organizationDetails | Export-Csv (Join-Path $csvFolderPath "Org.csv") -NoTypeInformation
    Write-Host "Organisation details (CSV) exported." -ForegroundColor Green
    #endregion

    #region Entra ID domains
    Write-Host "Getting Entra ID domains..." -ForegroundColor Cyan
    $AllMGDomains = Get-MgDomain -All
    $AllMGDomainsArray = @()
    $domainCount = $AllMGDomains.Count
    $domainIndex = 0

    Foreach ($MGDomain in $AllMGDomains) {
        $domainIndex++
        Write-Progress -Id 1 -Activity "Processing Entra ID domains" -Status "$domainIndex of $domainCount - $($MGDomain.Id)" -PercentComplete (($domainIndex / $domainCount) * 100)
        try {
            $AllMGDomainsArray += [PSCustomObject]@{
                Id                 = $MGDomain.Id
                AuthenticationType = $MGDomain.AuthenticationType
                IsVerified         = $MGDomain.IsVerified
                IsDefault          = $MGDomain.IsDefault
                IsInitial          = $MGDomain.IsInitial
                SupportedServices  = $MGDomain.SupportedServices -join "; "
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process domain '$($MGDomain.Id)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 1 -Activity "Processing Entra ID domains" -Completed

    $AllMGDomainsArray | Export-Csv (Join-Path $csvFolderPath "AllMGDomains.csv") -NoTypeInformation
    Write-Host "Domain data exported." -ForegroundColor Green
    #endregion

    #region Entra ID users
    Write-Host "Getting Entra ID users..." -ForegroundColor Cyan
    $Properties = @(
        'ID', 'UserPrincipalName', 'DisplayName', 'Mail',
        'LicenseAssignmentStates', 'UsageLocation', 'UserType',
        'AccountEnabled', 'OnPremisesSyncEnabled'
    )
    $MGUsers = Get-MgUser -All -Property $Properties | Sort-Object DisplayName
    Start-Sleep 10

    $MGUsersArray = @()
    $userCount = $MGUsers.Count
    $userIndex = 0

    Foreach ($MGUser in $MGUsers) {
        $userIndex++
        Write-Progress -Id 2 -Activity "Processing Entra ID users" -Status "$userIndex of $userCount - $($MGUser.DisplayName)" -PercentComplete (($userIndex / $userCount) * 100)
        try {
            $MGUsersArray += [PSCustomObject]@{
                ID                = $MGUser.ID
                DisplayName       = $MGUser.DisplayName
                UserPrincipalName = $MGUser.UserPrincipalName
                Mail              = $MGUser.Mail
                UserType          = $MGUser.UserType
                UsageLocation     = $MGUser.UsageLocation
                AccountEnabled    = $MGUser.AccountEnabled
                IsDirSynced       = $MGUser.OnPremisesSyncEnabled
                LicenseStatus     = $MGUser.LicenseAssignmentStates.State -join ';'
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process user '$($MGUser.UserPrincipalName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 2 -Activity "Processing Entra ID users" -Completed

    Start-Sleep 5
    $MGUsersArray | Export-Csv (Join-Path $csvFolderPath "AllMGUsers.csv") -NoTypeInformation
    Write-Host "User data exported." -ForegroundColor Green
    #endregion

    #region Licensing
    Write-Host "Getting friendly licence names from Microsoft Learn..." -ForegroundColor Cyan
    $ProductSkuCSV = @()
    try {
        $MSLicenseInfo = Invoke-WebRequest -UseBasicParsing "https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference"
        $SkuCSVLink = ($MSLicenseInfo.links | Where-Object { $_.href -match ".csv" }).href
        $WebClient = [System.Net.WebClient]::New()
        $InMemoryCSV = $WebClient.DownloadString($SkuCSVLink)
        $ProductSkuCSV = ConvertFrom-Csv -Delimiter ',' -InputObject $InMemoryCSV
        $WebClient.Dispose()
        Write-Host "Friendly licence names retrieved." -ForegroundColor Green
    }
    catch {
        Write-Log -Message "WARNING: couldn't retrieve friendly licence names from Microsoft Learn ($($_.Exception.Message)) - licence exports will fall back to raw SKU/part numbers instead of friendly names."
    }

    Write-Host "Getting tenant SKUs..." -ForegroundColor Cyan
    [Array]$Skus = Get-MgSubscribedSku | Sort-Object SkuPartNumber

    Write-Host "Getting per-user licence details..." -ForegroundColor Cyan
    $licenseDetailCount = $MGUsers.Count
    $licenseDetailIndex = 0

    foreach ($user in $MGUsers) {
        $licenseDetailIndex++
        Write-Progress -Id 3 -Activity "Processing per-user licence details" -Status "$licenseDetailIndex of $licenseDetailCount - $($user.UserPrincipalName)" -PercentComplete (($licenseDetailIndex / $licenseDetailCount) * 100)
        try {
            $user.LicenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
        }
        catch {
            Write-Log -Level Error -Message "Failed to retrieve licence details for '$($user.UserPrincipalName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 3 -Activity "Processing per-user licence details" -Completed

    $UserLicenseDetails = $MGUsers |
        Where-Object { $_.LicenseDetails } |
        Select-Object UserPrincipalName, @{
            Name       = 'Licenses'
            Expression = {
                ($_ | ForEach-Object {
                    $_.LicenseDetails | ForEach-Object {
                        $license = $_
                        $FriendlyLicense = ($ProductSkuCSV | Where-Object {
                                $_.String_ID -eq $license.SkuPartNumber
                            } | Select-Object -First 1).Product_Display_Name

                        if ($FriendlyLicense) {
                            $FriendlyLicense
                        }
                        else {
                            $license.SkuPartID
                        }
                    }
                }) -join ';'
            }
        }

    Start-Sleep 5
    $UserLicenseDetails | Export-Csv (Join-Path $csvFolderPath "LicenseData.csv") -NoTypeInformation
    Write-Host "User licence data exported." -ForegroundColor Green

    $LicenseOnePerLine = @()
    foreach ($user in $UserLicenseDetails) {
        $licenses = $user.Licenses -split ';'
        foreach ($license in $licenses) {
            $LicenseOnePerLine += [PSCustomObject]@{
                UserPrincipalName = $user.UserPrincipalName
                License           = $license.Trim()
            }
        }
    }
    $LicenseOnePerLine | Export-Csv (Join-Path $csvFolderPath "LicenseOnePerLine.csv") -NoTypeInformation
    Write-Host "Flattened licence data exported." -ForegroundColor Green

    Write-Host "Building tenant licence summary..." -ForegroundColor Cyan
    $DetailedLicenseSummary = @()
    foreach ($sku in $Skus) {
        $FriendlyName = ($ProductSkuCSV | Where-Object {
                $_.String_ID -eq $sku.SkuPartNumber
            } | Select-Object -First 1).Product_Display_Name

        if (-not $FriendlyName) {
            $FriendlyName = $sku.SkuPartNumber
        }

        $Purchased = $sku.PrepaidUnits.Enabled
        $Consumed = $sku.ConsumedUnits
        $Available = $Purchased - $Consumed

        $DetailedLicenseSummary += [PSCustomObject]@{
            PartNumber        = $sku.SkuPartNumber
            LicenseName       = $FriendlyName
            PurchasedLicenses = $Purchased
            ConsumedUnits     = $Consumed
            AvailableLicenses = $Available
        }
    }
    $DetailedLicenseSummary | Export-Csv (Join-Path $csvFolderPath "DetailedLicenseSummary.csv") -NoTypeInformation
    Write-Host "Tenant licence summary exported." -ForegroundColor Green
    #endregion

    #region MX records
    Write-Host "Resolving MX records for accepted domains..." -ForegroundColor Cyan
    $AllMXRecordsArray = @()
    $mxCount = $acceptedDomains.Count
    $mxIndex = 0
    # DNS lookups against external domains commonly fail for individual
    # domains (no record, timeout, etc.) - relaxed here so one bad
    # lookup doesn't abort the rest, then restored to Stop afterward so
    # the rest of the script keeps its normal error handling.
    $ErrorActionPreference = "SilentlyContinue"

    Foreach ($AcceptedDomain in $acceptedDomains) {
        $mxIndex++
        Write-Progress -Id 4 -Activity "Resolving MX records" -Status "$mxIndex of $mxCount - $($AcceptedDomain.DomainName)" -PercentComplete (($mxIndex / $mxCount) * 100)
        try {
            $MXRecords = $AcceptedDomain.DomainName | Resolve-DnsName -Type MX -Server 8.8.8.8 | Where-Object { $_.QueryType -eq "MX" } | Select-Object Name, NameExchange, Preference, TTL | Sort-Object Preference
            Foreach ($MXrecord in $MXRecords) {
                $AllMXRecordsArray += [PSCustomObject]@{
                    Name         = $MXRecord.Name
                    NameExchange = $MXRecord.NameExchange
                    Preference   = $MXRecord.Preference
                    TTL          = $MXRecord.TTL
                }
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to resolve MX records for '$($AcceptedDomain.DomainName)': $($_.Exception.Message)"
            continue
        }
    }
    $ErrorActionPreference = "Stop"
    Write-Progress -Id 4 -Activity "Resolving MX records" -Completed

    $AllMXRecordsArray | Export-Csv (Join-Path $csvFolderPath "MXRecords.csv") -NoTypeInformation
    Write-Host "MX record data exported." -ForegroundColor Green
    #endregion

    #region SPF records
    Write-Host "Resolving SPF records for accepted domains..." -ForegroundColor Cyan
    $AllSPFRecordsArray = @()
    $spfCount = $acceptedDomains.Count
    $spfIndex = 0
    $ErrorActionPreference = "SilentlyContinue"

    Foreach ($AcceptedDomain in $acceptedDomains) {
        $spfIndex++
        Write-Progress -Id 5 -Activity "Resolving SPF records" -Status "$spfIndex of $spfCount - $($AcceptedDomain.DomainName)" -PercentComplete (($spfIndex / $spfCount) * 100)
        try {
            $SPFRecords = $AcceptedDomain | Resolve-DnsName -Type TXT -Server 8.8.8.8
            Foreach ($SPFRecord in $SPFRecords) {
                If ($SPFRecord.Strings -like "V=SPF*") {
                    $AllSPFRecordsArray += [PSCustomObject]@{
                        Name   = $SPFRecord.Name
                        String = $SPFRecord.Strings -join ","
                        TTL    = $SPFRecord.TTL
                    }
                }
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to resolve SPF records for '$($AcceptedDomain.DomainName)': $($_.Exception.Message)"
            continue
        }
    }
    $ErrorActionPreference = "Stop"
    Write-Progress -Id 5 -Activity "Resolving SPF records" -Completed

    $AllSPFRecordsArray | Export-Csv (Join-Path $csvFolderPath "SPFRecords.csv") -NoTypeInformation
    Write-Host "SPF record data exported." -ForegroundColor Green
    #endregion

    #region DMARC records
    Write-Host "Resolving DMARC records for accepted domains..." -ForegroundColor Cyan
    $AllDMARCRecordsArray = @()
    $dmarcCount = $acceptedDomains.Count
    $dmarcIndex = 0
    $ErrorActionPreference = "SilentlyContinue"

    Foreach ($AcceptedDomain in $acceptedDomains) {
        $dmarcIndex++
        Write-Progress -Id 6 -Activity "Resolving DMARC records" -Status "$dmarcIndex of $dmarcCount - $($AcceptedDomain.DomainName)" -PercentComplete (($dmarcIndex / $dmarcCount) * 100)
        try {
            $DomainName = $AcceptedDomain.DomainName
            $dnsTxt = Resolve-DnsName -Name "_dmarc.$DomainName" -Type TXT | Where-Object Type -eq TXT
            $AllDMARCRecordsArray += [PSCustomObject]@{
                Domain  = $DomainName
                Name    = $dnsTxt.Name
                Section = $dnsTxt.Section
                String  = $dnsTxt.Strings -join ","
                TTL     = $dnsTxt.TTL
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to resolve DMARC record for '$($AcceptedDomain.DomainName)': $($_.Exception.Message)"
            continue
        }
    }
    $ErrorActionPreference = "Stop"
    Write-Progress -Id 6 -Activity "Resolving DMARC records" -Completed

    $AllDMARCRecordsArray | Export-Csv (Join-Path $csvFolderPath "DMARCRecords.csv") -NoTypeInformation
    Write-Host "DMARC record data exported." -ForegroundColor Green
    #endregion

    #region Teams user activity
    Write-Host "Getting Teams user activity (last 180 days)..." -ForegroundColor Cyan
    try {
        Get-MgReportTeamUserActivityUserDetail -Period "D180" -OutFile (Join-Path $csvFolderPath "TeamsUserActivity_Last180Days.csv")
        Write-Host "Teams user activity data exported." -ForegroundColor Green
    }
    catch {
        Write-Log -Level Error -Message "Failed to retrieve Teams user activity report: $($_.Exception.Message)"
    }
    #endregion

    #region Security groups
    Write-Host "Getting security groups..." -ForegroundColor Cyan
    $SecurityGroups = Get-MgGroup -All -Filter "securityEnabled eq true and mailEnabled eq false" -Property "id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,createdDateTime" | Sort-Object DisplayName
    $SecurityGroupsArray = @()
    $sgCount = $SecurityGroups.Count
    $sgIndex = 0

    Foreach ($SecurityGroup in $SecurityGroups) {
        $sgIndex++
        Write-Progress -Id 7 -Activity "Processing security groups" -Status "$sgIndex of $sgCount - $($SecurityGroup.DisplayName)" -PercentComplete (($sgIndex / $sgCount) * 100)
        try {
            $MemberCount = $null
            try {
                $MemberCount = (Get-MgGroupMember -GroupId $SecurityGroup.Id -All -Property Id -ErrorAction Stop).Count
            }
            catch {
                Write-Log -Message "Couldn't get member count for security group '$($SecurityGroup.DisplayName)': $($_.Exception.Message)"
            }

            $OwnerUPNs = $null
            try {
                $Owners = Get-MgGroupOwner -GroupId $SecurityGroup.Id -All -ErrorAction Stop
                $OwnerUPNs = ($Owners | ForEach-Object { $_.AdditionalProperties.userPrincipalName } | Where-Object { $_ }) -join ","
            }
            catch {
                Write-Log -Message "Couldn't get owners for security group '$($SecurityGroup.DisplayName)': $($_.Exception.Message)"
            }

            $SecurityGroupsArray += [PSCustomObject]@{
                DisplayName     = $SecurityGroup.DisplayName
                Description     = $SecurityGroup.Description
                Mail            = $SecurityGroup.Mail
                MailEnabled     = $SecurityGroup.MailEnabled
                SecurityEnabled = $SecurityGroup.SecurityEnabled
                GroupTypes      = ($SecurityGroup.GroupTypes -join ",")
                CreatedDateTime = $SecurityGroup.CreatedDateTime
                Id              = $SecurityGroup.Id
                MemberCount     = $MemberCount
                Owners          = $OwnerUPNs
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process security group '$($SecurityGroup.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 7 -Activity "Processing security groups" -Completed

    $SecurityGroupsArray | Export-Csv (Join-Path $csvFolderPath "SecurityGroups.csv") -NoTypeInformation -Encoding UTF8
    Write-Host "Security group data exported." -ForegroundColor Green
    #endregion

    #region Registered applications
    Write-Host "Getting registered applications..." -ForegroundColor Cyan
    $Applications = Get-MgApplication -All
    $AppAuthMethodsArray = @()
    $regAppCount = $Applications.Count
    $regAppIndex = 0

    Foreach ($Application in $Applications) {
        $regAppIndex++
        Write-Progress -Id 8 -Activity "Processing registered applications" -Status "$regAppIndex of $regAppCount - $($Application.DisplayName)" -PercentComplete (($regAppIndex / $regAppCount) * 100)
        try {
            $Owners = (Get-MgApplicationOwner -ApplicationId $Application.Id | ForEach-Object { (Get-MgUser -UserId $_.Id -ErrorAction SilentlyContinue).UserPrincipalName }) -join '; '

            $AppAuthMethodsArray += [PSCustomObject]@{
                DisplayName        = $Application.DisplayName
                AppId              = $Application.AppId
                PublisherDomain    = $Application.PublisherDomain
                CreatedDateTime    = $Application.CreatedDateTime
                SecretCount        = $Application.PasswordCredentials.Count
                CertificateCount   = $Application.KeyCredentials.Count
                LatestSecretExpiry = ($Application.PasswordCredentials | Sort-Object EndDateTime -Descending | Select-Object -First 1).EndDateTime
                LatestCertExpiry   = ($Application.KeyCredentials | Sort-Object EndDateTime -Descending | Select-Object -First 1).EndDateTime
                Owners             = $Owners
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process registered application '$($Application.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 8 -Activity "Processing registered applications" -Completed

    $AppAuthMethodsArray | Export-Csv (Join-Path $csvFolderPath "RegisteredApplications.csv") -NoTypeInformation -Encoding UTF8
    Write-Host "Registered application data exported." -ForegroundColor Green
    #endregion

    #region BitLocker recovery keys
    Write-Host "Getting BitLocker recovery keys..." -ForegroundColor Cyan
    $BitlockerKeys = Get-MgInformationProtectionBitlockerRecoveryKey
    $BitlockerKeysArray = @()
    $bitlockerCount = $BitlockerKeys.Count
    $bitlockerIndex = 0

    Foreach ($BitlockerKey in $BitlockerKeys) {
        $bitlockerIndex++
        Write-Progress -Id 9 -Activity "Processing BitLocker recovery keys" -Status "$bitlockerIndex of $bitlockerCount - $($BitlockerKey.DeviceId)" -PercentComplete (($bitlockerIndex / $bitlockerCount) * 100)
        try {
            $BitlockerKeysArray += [PSCustomObject]@{
                CreatedDateTime = $BitlockerKey.CreatedDateTime
                DeviceId        = $BitlockerKey.DeviceId
                ID              = $BitlockerKey.ID
                VolumeType      = $BitlockerKey.VolumeType
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process BitLocker key for device '$($BitlockerKey.DeviceId)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 9 -Activity "Processing BitLocker recovery keys" -Completed

    $BitlockerKeysArray | Export-Csv (Join-Path $csvFolderPath "BitlockerKeys.csv") -NoTypeInformation
    Write-Host "BitLocker key data exported." -ForegroundColor Green
    #endregion

    #region Devices
    Write-Host "Getting device information..." -ForegroundColor Cyan
    $MGDevices = Get-MgDevice -Property "AccountEnabled,ApproximateLastSignInDateTime,DeviceCategory,DeviceId,DeviceOwnership,DisplayName,EnrollmentType,Id,IsCompliant,IsManaged,ManagementType,Manufacturer,Model,OperatingSystem,OperatingSystemVersion,OnPremisesSyncEnabled,ProfileType,RegisteredOwners,TrustType" -All | Sort-Object DisplayName
    $MGDeviceArray = @()
    $deviceCount = $MGDevices.Count
    $deviceIndex = 0

    Foreach ($MGDevice in $MGDevices) {
        $deviceIndex++
        Write-Progress -Id 10 -Activity "Processing devices" -Status "$deviceIndex of $deviceCount - $($MGDevice.DisplayName)" -PercentComplete (($deviceIndex / $deviceCount) * 100)
        try {
            $MGDeviceArray += [PSCustomObject]@{
                AccountEnabled                = $MGDevice.AccountEnabled
                ApproximateLastSignInDateTime = $MGDevice.ApproximateLastSignInDateTime
                DeviceCategory                = $MGDevice.DeviceCategory
                DeviceId                      = $MGDevice.DeviceId
                DeviceOwnership               = $MGDevice.DeviceOwnership
                DisplayName                   = $MGDevice.DisplayName
                EnrollmentType                = $MGDevice.EnrollmentType
                Id                            = $MGDevice.Id
                IsCompliant                   = $MGDevice.IsCompliant
                IsManaged                     = $MGDevice.IsManaged
                ManagementType                = $MGDevice.ManagementType
                Manufacturer                  = $MGDevice.Manufacturer
                Model                         = $MGDevice.Model
                OperatingSystem               = $MGDevice.OperatingSystem
                OperatingSystemVersion        = $MGDevice.OperatingSystemVersion
                OnPremisesSyncEnabled         = $MGDevice.OnPremisesSyncEnabled
                ProfileType                   = $MGDevice.ProfileType
                RegisteredOwners              = ($MGDevice.RegisteredOwners | ForEach-Object { $_.AdditionalProperties.userPrincipalName } | Where-Object { $_ }) -join ","
                TrustType                     = $MGDevice.TrustType
            }
        }
        catch {
            Write-Log -Level Error -Message "Failed to process device '$($MGDevice.DisplayName)': $($_.Exception.Message)"
            continue
        }
    }
    Write-Progress -Id 10 -Activity "Processing devices" -Completed

    $MGDeviceArray | Export-Csv (Join-Path $csvFolderPath "Devices.csv") -NoTypeInformation
    Write-Host "Device data exported." -ForegroundColor Green
    #endregion
}
catch {
    Write-Log -Level Error -Message "1-Softcat-T2T-Assessment-Initial-Script.v2.0.0.PS1 stopped due to an unexpected error: $($_.Exception.Message)"
    throw
}

# $organizationDetails and $acceptedDomains are available in this
# session, and also written to disk under the JSON subfolder so other
# scripts can import them, e.g.:
#   $organizationDetails = Get-Content "<path>\JSON\<domain>-OrganizationDetails.json" -Raw | ConvertFrom-Json
#   $acceptedDomains     = Get-Content "<path>\JSON\<domain>-AcceptedDomains.json" -Raw | ConvertFrom-Json
#
# All CSV outputs are in: $csvFolderPath
# The warning/error log for this run is at: $logFilePath
#
# AppId, TenantId and CertificateThumbprint also remain in scope for
# use by Get-MailboxStatistics and Get-EXOData, which take the same
# three connection values as their own parameters.
#
# Connections are left open intentionally so further Graph/EXO cmdlets
# can be run against them in this session too. Disconnect when finished:
#   Disconnect-MgGraph
#   Disconnect-ExchangeOnline -Confirm:$false
