#Requires -Version 5.1
#Requires -Modules PowerShellGet

<#
.SYNOPSIS
    Checks installed PowerShell modules for available updates and interactively
    prompts before installing each one.

.DESCRIPTION
    Enumerates modules installed via PowerShellGet, queries their source repository
    (e.g. PSGallery) for the latest available version, and for every module that is
    out of date prompts the user Y/N/A before running Update-Module. Supports -WhatIf
    and -Confirm via ShouldProcess in addition to the interactive Y/N/A prompt.

.PARAMETER ExcludeModule
    Module name(s) to skip. Wildcards allowed.

.PARAMETER Scope
    Update scope: CurrentUser or AllUsers. AllUsers requires an elevated session.
    Defaults to CurrentUser.

.PARAMETER Force
    Automatically confirm every update (no per-module prompting). Use with caution.

.EXAMPLE
    .\Update-InstalledModulesInteractiveV1.0.ps1

.EXAMPLE
    .\Update-InstalledModulesInteractiveV1.0.ps1 -ExcludeModule Az.* -Scope AllUsers

.EXAMPLE
    .\Update-InstalledModulesInteractiveV1.0.ps1 -WhatIf

.NOTES
    Name:    Update-InstalledModules.v2.0.0.ps1
    Version: 2.0.0
    Author: simonan@softcat.com
    Author: si@tumeke.cloud
    Prerequisites:
      Install-Module PowerShellGet -Scope CurrentUser -Force
      Internet access to reach the PowerShell Gallery
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
[OutputType([PSCustomObject])]
param(
    [SupportsWildcards()]
    [string[]]$ExcludeModule = @(),

    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]$Scope = 'CurrentUser',

    [switch]$Force
)

Set-StrictMode -Version Latest

function Get-LatestModuleVersion {
    [CmdletBinding()]
    [OutputType('Microsoft.PowerShell.Commands.PSRepositoryItemInfo')]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    try {
        Find-Module -Name $Name -ErrorAction Stop |
            Sort-Object -Property Version -Descending |
            Select-Object -First 1
    }
    catch {
        Write-Verbose "Could not find '$Name' in the gallery: $($_.Exception.Message)"
        $null
    }
}

function Test-ModuleExcluded {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [string[]]$Pattern
    )

    foreach ($p in $Pattern) {
        if ($Name -like $p) { return $true }
    }
    return $false
}

if ($Scope -eq 'AllUsers') {
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Scope 'AllUsers' requires an elevated session. Re-run this script as Administrator."
    }
}

Write-Host "Gathering installed modules..." -ForegroundColor Cyan

$installedModules = @(Get-InstalledModule -ErrorAction SilentlyContinue | Sort-Object -Property Name)

if ($installedModules.Count -eq 0) {
    Write-Warning "No modules found via Get-InstalledModule. Are you running this in a session with PowerShellGet available?"
    exit 0
}

$totalChecked = 0
$updatesFound = 0
$updatesInstalled = 0
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$yesToAll = $Force.IsPresent
$moduleIndex = 0

foreach ($module in $installedModules) {

    $latest = $null
    $moduleIndex++
    Write-Progress -Activity 'Checking installed modules' `
        -Status "($moduleIndex of $($installedModules.Count)) $($module.Name)" `
        -PercentComplete (($moduleIndex / $installedModules.Count) * 100)

    if (Test-ModuleExcluded -Name $module.Name -Pattern $ExcludeModule) {
        continue
    }

    $totalChecked++
    Write-Verbose "Checking $($module.Name) (installed: $($module.Version))"

    $latest = Get-LatestModuleVersion -Name $module.Name
    if (-not $latest) { continue }

    if ([version]$latest.Version -le [version]$module.Version) { continue }

    $updatesFound++

    Write-Host ""
    Write-Host "Update available for '$($module.Name)':" -ForegroundColor Yellow
    Write-Host "  Installed: $($module.Version)"
    Write-Host "  Latest:    $($latest.Version)"

    $doUpdate = $yesToAll

    if (-not $yesToAll) {
        do {
            $response = Read-Host "Install this update? (Y/N/A=Yes to all)"
        } until ($response -match '^[YyNnAa]$')

        if ($response -match '^[Aa]$') {
            $yesToAll = $true
            $doUpdate = $true
        }
        else {
            $doUpdate = $response -match '^[Yy]$'
        }
    }

    if (-not $doUpdate) {
        Write-Host "Skipped $($module.Name)" -ForegroundColor DarkGray
        $results.Add([PSCustomObject]@{
            Module    = $module.Name
            Installed = $module.Version
            Latest    = $latest.Version
            Action    = 'Skipped'
        })
        continue
    }

    if ($PSCmdlet.ShouldProcess($module.Name, "Update-Module to $($latest.Version)")) {
        try {
            Write-Host "Installing update for $($module.Name)..." -ForegroundColor Cyan
            Update-Module -Name $module.Name -RequiredVersion $latest.Version -Scope $Scope -Force -ErrorAction Stop
            Write-Host "Updated $($module.Name) to $($latest.Version)" -ForegroundColor Green
            $updatesInstalled++
            $results.Add([PSCustomObject]@{
                Module    = $module.Name
                Installed = $module.Version
                Latest    = $latest.Version
                Action    = 'Updated'
            })
        }
        catch {
            Write-Warning "Failed to update $($module.Name): $($_.Exception.Message)"
            $results.Add([PSCustomObject]@{
                Module    = $module.Name
                Installed = $module.Version
                Latest    = $latest.Version
                Action    = 'Failed'
            })
        }
    }
    else {
        $results.Add([PSCustomObject]@{
            Module    = $module.Name
            Installed = $module.Version
            Latest    = $latest.Version
            Action    = 'Skipped (WhatIf)'
        })
    }
}

Write-Progress -Activity 'Checking installed modules' -Completed

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Summary" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "Modules checked   : $totalChecked"
Write-Host "Updates found     : $updatesFound"
Write-Host "Updates installed : $updatesInstalled"

if ($results.Count -gt 0) {
    Write-Host ""
    $results | Format-Table -AutoSize
}
else {
    Write-Host "All modules are up to date."
}
