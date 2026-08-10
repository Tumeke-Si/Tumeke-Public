<#
.SYNOPSIS
    Verifies that the Softcat T2T assessment scripts' expected CSV
    outputs are present.

.DESCRIPTION
    RUN ORDER: run this last, after some or all of the
    2-Softcat-T2T-Assessment-* scripts, to verify their CSV outputs are
    present. Safe to run at any point, though - it only reads the CSV
    Files folder, so it won't complain about anything more than "not
    produced yet" if you run it early.

    Checks a "CSV Files" folder against the full list of CSV files this
    project's scripts are expected to produce (1-Softcat-T2T-Assessment-
    Initial-Script.v2.0.0.PS1, 2-Softcat-T2T-Assessment-EXO-1.v2.0.0.PS1,
    2-Softcat-T2T-Assessment-EXO-2.v2.0.0.PS1,
    2-Softcat-T2T-Assessment-Teams-1.v2.0.0.PS1 and
    2-Softcat-T2T-Assessment-SPO-1.v2.0.0.PS1), and reports:
      - Which expected files are present (with row count and last
        modified time, as a quick sanity check that they're not empty
        or stale)
      - Which expected files are missing (that script hasn't been run
        yet, or failed before exporting)
      - Any extra .csv files in the folder that aren't part of the
        known list (renamed files, leftovers, or output from a script
        this inventory doesn't know about yet)

    Purely local - doesn't connect to Graph/Exchange/Teams/SharePoint,
    so no AppId/TenantId/certificate needed.

    Exports its own findings to CSVFileInventory.csv in the same
    folder, and writes warnings (missing/unexpected files) to a
    timestamped log file in a "Logs" subfolder, same as the other
    scripts in this suite.

.PARAMETER CsvFolderPath
    Folder to check for CSV files. Defaults to a "CSV Files" subfolder
    next to this script.

.PARAMETER OutputPath
    Base folder for the Logs subfolder. Defaults to the same folder
    as this script.

.EXAMPLE
    .\9-Softcat-T2T-Check-Output.v2.0.0.PS1

    Checks the "CSV Files" subfolder next to this script and reports
    the inventory.

.NOTES
    Authors: simonan@softcat.com, si@tumeke.cloud

    Prerequisites:
      None - built-in cmdlets only.
#>
[CmdletBinding()]
param(
    [string]$CsvFolderPath,

    [string]$OutputPath
)

$ErrorActionPreference = "Stop"
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $scriptDir
}

if ([string]::IsNullOrWhiteSpace($CsvFolderPath)) {
    $CsvFolderPath = Join-Path $OutputPath "CSV Files"
}

#region Ensure output folders exist
if (-not (Test-Path $CsvFolderPath)) {
    New-Item -Path $CsvFolderPath -ItemType Directory | Out-Null
}

$logsFolderPath = Join-Path $OutputPath "Logs"
if (-not (Test-Path $logsFolderPath)) {
    New-Item -Path $logsFolderPath -ItemType Directory | Out-Null
}
#endregion

#region Logging
$logFilePath = Join-Path $logsFolderPath "9-Softcat-T2T-Check-Output-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

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

#region Known expected CSV files, by source script
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

#region Check expected files
Write-Host "Checking $($ExpectedFiles.Count) expected CSV files in: $CsvFolderPath" -ForegroundColor Cyan
$InventoryArray = @()
$fileCount = $ExpectedFiles.Count
$fileIndex = 0

Foreach ($Expected in $ExpectedFiles) {
    $fileIndex++
    Write-Progress -Id 1 -Activity "Checking expected CSV files" -Status "$fileIndex of $fileCount - $($Expected.FileName)" -PercentComplete (($fileIndex / $fileCount) * 100)

    $filePath = Join-Path $CsvFolderPath $Expected.FileName

    if (Test-Path $filePath) {
        $rowCount = $null
        try {
            $rowCount = @(Import-Csv -Path $filePath).Count
        }
        catch {
            Write-Log -Message "Found '$($Expected.FileName)' but couldn't read it as CSV: $($_.Exception.Message)"
        }

        $lastModified = (Get-Item -Path $filePath).LastWriteTime

        $InventoryArray += [PSCustomObject]@{
            FileName     = $Expected.FileName
            SourceScript = $Expected.SourceScript
            Status       = "Present"
            RowCount     = $rowCount
            LastModified = $lastModified
        }
    }
    else {
        Write-Log -Message "Missing expected file '$($Expected.FileName)' (produced by $($Expected.SourceScript)) - that script may not have been run yet, or may have failed before exporting it."

        $InventoryArray += [PSCustomObject]@{
            FileName     = $Expected.FileName
            SourceScript = $Expected.SourceScript
            Status       = "Missing"
            RowCount     = $null
            LastModified = $null
        }
    }
}
Write-Progress -Id 1 -Activity "Checking expected CSV files" -Completed
#endregion

#region Check for unexpected extra files
$actualCsvFiles = Get-ChildItem -Path $CsvFolderPath -Filter "*.csv" -ErrorAction SilentlyContinue
$expectedNames = $ExpectedFiles.FileName
$unexpectedFiles = $actualCsvFiles | Where-Object { $_.Name -notin $expectedNames -and $_.Name -ne "CSVFileInventory.csv" }

Foreach ($UnexpectedFile in $unexpectedFiles) {
    Write-Log -Message "Unexpected file '$($UnexpectedFile.Name)' found in the CSV Files folder - not part of the known list for this project. Could be a renamed file, a leftover, or output from a script this inventory doesn't know about yet."

    $InventoryArray += [PSCustomObject]@{
        FileName     = $UnexpectedFile.Name
        SourceScript = "(unknown)"
        Status       = "Unexpected"
        RowCount     = $null
        LastModified = $UnexpectedFile.LastWriteTime
    }
}
#endregion

#region Summary
$presentCount = ($InventoryArray | Where-Object { $_.Status -eq "Present" }).Count
$missingCount = ($InventoryArray | Where-Object { $_.Status -eq "Missing" }).Count
$unexpectedCount = ($InventoryArray | Where-Object { $_.Status -eq "Unexpected" }).Count

Write-Host "`n============================================================" -ForegroundColor Yellow
Write-Host " CSV file inventory summary" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "Present:    $presentCount of $($ExpectedFiles.Count) expected files" -ForegroundColor Green
if ($missingCount -gt 0) {
    Write-Host "Missing:    $missingCount expected files" -ForegroundColor Yellow
}
if ($unexpectedCount -gt 0) {
    Write-Host "Unexpected: $unexpectedCount extra files not in the known list" -ForegroundColor Yellow
}

$InventoryArray | Sort-Object Status, SourceScript, FileName | Format-Table FileName, SourceScript, Status, RowCount, LastModified -AutoSize

$inventoryPath = Join-Path $CsvFolderPath "CSVFileInventory.csv"
$InventoryArray | Export-Csv $inventoryPath -NoTypeInformation
Write-Host "Inventory exported to: $inventoryPath" -ForegroundColor Green
#endregion

# The warning/error log for this run is at: $logFilePath
