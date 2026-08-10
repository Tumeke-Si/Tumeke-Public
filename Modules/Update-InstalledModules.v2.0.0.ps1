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

    Some modules (e.g. Microsoft.Graph) are "meta" packages whose RequiredModules
    pin a whole family of submodules (Microsoft.Graph.*) to the same version. Updating
    the parent therefore silently updates those submodules too. After updating a
    module listed in $CascadeParentModules, the script re-checks the installed
    version of every not-yet-processed module in that family so it doesn't prompt to
    update something that was already brought current as a side effect.

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
    Version: 2.1.0
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

    [switch]$Force,

    # Modules whose update cascades to a family of submodules (e.g. updating
    # Microsoft.Graph also updates Microsoft.Graph.Authentication, Microsoft.Graph.Users, etc.
    # because they're pinned RequiredModules). After updating one of these, the
    # remaining not-yet-processed modules matching "<Name>.*" are re-checked
    # against their real installed version instead of the stale pre-update snapshot.
    [string[]]$CascadeParentModules = @('Microsoft.Graph')
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

function Sync-CascadedModuleVersions {
    <#
        After a cascade-parent module (e.g. Microsoft.Graph) is updated, its
        RequiredModules submodules get updated as a side effect. Refresh the
        recorded Version for the remaining not-yet-processed modules in $Modules
        so later loop iterations don't see the stale pre-update version and
        re-prompt to "update" something that's already current.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[PSCustomObject]]$Modules,

        [Parameter(Mandatory)]
        [int]$FromIndex,

        [Parameter(Mandatory)]
        [string]$ParentName
    )

    $current = @(Get-InstalledModule -Name "$ParentName*" -ErrorAction SilentlyContinue)
    if ($current.Count -eq 0) { return }

    $lookup = @{}
    foreach ($c in $current) { $lookup[$c.Name] = $c.Version }

    for ($j = $FromIndex; $j -lt $Modules.Count; $j++) {
        $entry = $Modules[$j]
        if ($lookup.ContainsKey($entry.Name) -and [version]$lookup[$entry.Name] -ne [version]$entry.Version) {
            Write-Verbose "$($entry.Name) was already updated as a dependency of $ParentName ($($entry.Version) -> $($lookup[$entry.Name]))"
            $entry.Version = $lookup[$entry.Name]
        }
    }
}

if ($Scope -eq 'AllUsers') {
    $currentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Scope 'AllUsers' requires an elevated session. Re-run this script as Administrator."
    }
}

Write-Host "Gathering installed modules..." -ForegroundColor Cyan

$installedModules = [System.Collections.Generic.List[PSCustomObject]]::new()
@(Get-InstalledModule -ErrorAction SilentlyContinue | Sort-Object -Property Name) | ForEach-Object {
    $installedModules.Add([PSCustomObject]@{
        Name    = $_.Name
        Version = $_.Version
    })
}

if ($installedModules.Count -eq 0) {
    Write-Warning "No modules found via Get-InstalledModule. Are you running this in a session with PowerShellGet available?"
    exit 0
}

$totalChecked = 0
$updatesFound = 0
$updatesInstalled = 0
$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$yesToAll = $Force.IsPresent

for ($moduleIndex = 0; $moduleIndex -lt $installedModules.Count; $moduleIndex++) {

    $module = $installedModules[$moduleIndex]
    $latest = $null
    Write-Progress -Activity 'Checking installed modules' `
        -Status "($($moduleIndex + 1) of $($installedModules.Count)) $($module.Name)" `
        -PercentComplete ((($moduleIndex + 1) / $installedModules.Count) * 100)

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

            if ($module.Name -in $CascadeParentModules) {
                Write-Verbose "$($module.Name) is a cascade parent - re-checking installed versions of its submodules..."
                Sync-CascadedModuleVersions -Modules $installedModules -FromIndex ($moduleIndex + 1) -ParentName $module.Name
            }
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
