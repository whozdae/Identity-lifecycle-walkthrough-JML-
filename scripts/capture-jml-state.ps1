<#
.SYNOPSIS
    Snapshots the access state of the JML test population, so before/after can be
    compared with evidence rather than memory.

.DESCRIPTION
    Records, for every test user: account enabled state, department, manager,
    lifecycle dates, license count, and every group they belong to. Writes a
    labelled CSV.

    Run it with -Label before, run the workflow, then run it with -Label after.
    The second run automatically diffs against the matching before-snapshot and
    prints exactly what changed. That diff is the line that goes in the README.

    This script only reads. It changes nothing.

    Delegated Graph scopes requested:
      User.Read.All                     read user state
      Group.Read.All                    read group membership
      User-LifeCycleInfo.ReadWrite.All  read employeeHireDate / employeeLeaveDateTime
      Organization.Read.All             resolve license SKU names

.PARAMETER Label
    Snapshot name, typically "before" or "after". Files are written as
    jml-state-<Label>.csv in the script folder.

.PARAMETER Force
    Overwrite a snapshot file that already exists. Without it, an existing snapshot is
    left alone, because a before-state cannot be recaptured once the workflow has run.

.PARAMETER Scenario
    Optional tag for which demo this snapshot belongs to (joiner, mover, leaver).
    Keeps snapshots from overwriting each other across the three demos.

.EXAMPLE
    .\capture-jml-state.ps1 -Label before -Scenario leaver

.EXAMPLE
    .\capture-jml-state.ps1 -Label after -Scenario leaver
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$Label,
    [ValidateSet("joiner", "mover", "leaver", "all")] [string]$Scenario = "all",
    [string[]]$TestDepartments = @("Security", "IT Operations", "Finance", "Human Resources", "Governance", "Vendor"),
    [switch]$Force,
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = "Stop"
$Graph = "https://graph.microsoft.com/v1.0"

function Invoke-Graph ([string]$Uri) {
    try { Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject }
    catch {
        $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw "Graph GET $Uri failed: $detail"
    }
}

function Get-GraphAll ([string]$Uri) {
    $items = @(); $next = $Uri
    while ($next) {
        $page = Invoke-Graph $next
        if ($page.value) { $items += $page.value }
        $next = $page.'@odata.nextLink'
    }
    return $items
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " JML state snapshot: '$Label' ($Scenario)" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

Import-Module Microsoft.Graph.Authentication
$connect = @{
    Scopes = @("User.Read.All", "Group.Read.All", "User-LifeCycleInfo.ReadWrite.All", "Organization.Read.All")
    NoWelcome = $true
}
if ($TenantId)      { $connect.TenantId = $TenantId }
if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
Connect-MgGraph @connect

$select = "id,displayName,department,jobTitle,accountEnabled,employeeHireDate,employeeLeaveDateTime,assignedLicenses"
$users  = @(Get-GraphAll "$Graph/users?`$select=$select&`$top=999") |
          Where-Object { $TestDepartments -contains $_.department } |
          Sort-Object displayName

$snapshot = foreach ($u in $users) {
    $groups = @(Get-GraphAll "$Graph/users/$($u.id)/memberOf?`$select=displayName") |
              Where-Object { $_.displayName } | Select-Object -ExpandProperty displayName | Sort-Object
    $mgr = try { (Invoke-Graph "$Graph/users/$($u.id)/manager?`$select=displayName").displayName } catch { $null }

    [PSCustomObject]@{
        User        = $u.displayName
        Department  = $u.department
        Enabled     = $u.accountEnabled
        Manager     = $mgr
        HireDate    = if ($u.employeeHireDate)      { ([datetime]$u.employeeHireDate).ToString("yyyy-MM-dd") } else { "" }
        LastDay     = if ($u.employeeLeaveDateTime) { ([datetime]$u.employeeLeaveDateTime).ToString("yyyy-MM-dd") } else { "" }
        Licenses    = @($u.assignedLicenses).Count
        GroupCount  = $groups.Count
        Groups      = ($groups -join " | ")
    }
}

$snapshot | Format-Table User, Department, Enabled, Licenses, GroupCount -AutoSize

$path = Join-Path $PSScriptRoot "jml-state-$Scenario-$Label.csv"

# A before-snapshot is unrecoverable once the workflow runs, so never clobber one by
# accident. -Force overwrites deliberately.
if ((Test-Path $path) -and -not $Force) {
    $existing = Import-Csv $path
    Write-Host "`n[STOP] $path already exists and was NOT overwritten." -ForegroundColor Red
    Write-Host "       It holds $($existing.Count) users captured earlier." -ForegroundColor Yellow
    Write-Host "       If that is your real 'before' state, keep it: run with -Label after instead." -ForegroundColor Yellow
    Write-Host "       To replace it on purpose, rerun with -Force.`n" -ForegroundColor Yellow
    return
}

$snapshot | Export-Csv -Path $path -NoTypeInformation -Force
Write-Host "[OK] Snapshot written to $path" -ForegroundColor Green

# ---------------------------------------------------------------- diff
if ($Label -ne "before") {
    $beforePath = Join-Path $PSScriptRoot "jml-state-$Scenario-before.csv"
    if (-not (Test-Path $beforePath)) {
        Write-Host "[INFO] No before-snapshot at $beforePath, so nothing to compare." -ForegroundColor Cyan
        return
    }

    $before = Import-Csv $beforePath
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host " WHAT CHANGED" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    $changes = @()
    foreach ($now in $snapshot) {
        $was = $before | Where-Object { $_.User -eq $now.User } | Select-Object -First 1
        if (-not $was) { continue }

        foreach ($field in @("Department", "Enabled", "Manager", "Licenses", "GroupCount")) {
            if ("$($was.$field)" -ne "$($now.$field)") {
                $changes += [PSCustomObject]@{ User = $now.User; Field = $field; Before = $was.$field; After = $now.$field }
            }
        }

        $wasGroups = @($was.Groups -split " \| " | Where-Object { $_ })
        $nowGroups = @($now.Groups -split " \| " | Where-Object { $_ })
        foreach ($g in ($wasGroups | Where-Object { $nowGroups -notcontains $_ })) {
            $changes += [PSCustomObject]@{ User = $now.User; Field = "Group removed"; Before = $g; After = "" }
        }
        foreach ($g in ($nowGroups | Where-Object { $wasGroups -notcontains $_ })) {
            $changes += [PSCustomObject]@{ User = $now.User; Field = "Group added"; Before = ""; After = $g }
        }
    }

    if ($changes.Count -eq 0) {
        Write-Host "No differences. Either the workflow hasn't run yet, or it changed nothing." -ForegroundColor Yellow
    } else {
        $changes | Format-Table -AutoSize
        $diffPath = Join-Path $PSScriptRoot "jml-diff-$Scenario.csv"
        $changes | Export-Csv -Path $diffPath -NoTypeInformation -Force
        Write-Host "[OK] $($changes.Count) changes written to $diffPath" -ForegroundColor Green
    }
}
