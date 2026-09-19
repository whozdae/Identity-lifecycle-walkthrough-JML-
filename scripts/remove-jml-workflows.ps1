<#
.SYNOPSIS
    Rolls back the JML lab: deletes the lifecycle workflows and, optionally, the lab
    groups and the lifecycle dates seeded onto the test users.

.DESCRIPTION
    Deleted workflows go to the Lifecycle Workflows deleted-items list, where Entra
    keeps them for 30 days, so this is recoverable rather than destructive.

    -IncludeGroups also deletes the SG-All-Employees and SG-Dept-* groups and clears
    employeeHireDate / employeeLeaveDateTime on the test users, returning the tenant
    to its pre-lab state.

    Delegated Graph scopes requested:
      LifecycleWorkflows-Workflow.ReadWrite.All  delete the workflows
      Group.ReadWrite.All                        delete the lab groups (-IncludeGroups)
      User-LifeCycleInfo.ReadWrite.All           clear lifecycle dates (-IncludeGroups)
      User.Read.All                              find the test users

.EXAMPLE
    .\remove-jml-workflows.ps1 -WhatIf

.EXAMPLE
    .\remove-jml-workflows.ps1 -IncludeGroups -Confirm:$false
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
param(
    [string[]]$WorkflowNames = @("Onboard pre-hire employee", "Department transfer - access update", "Offboard on last day"),
    [string]$ConfigPath,
    [switch]$IncludeGroups,
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = "Stop"
$Graph = "https://graph.microsoft.com/v1.0"
$LCW   = "$Graph/identityGovernance/lifecycleWorkflows"

function Invoke-Graph {
    param(
        [ValidateSet("GET","POST","PATCH","DELETE")] [string]$Method = "GET",
        [Parameter(Mandatory)] [string]$Uri,
        $Body
    )
    $p = @{ Method = $Method; Uri = $Uri; OutputType = "PSObject" }
    if ($Body) { $p.Body = ($Body | ConvertTo-Json -Depth 10); $p.ContentType = "application/json" }
    try { Invoke-MgGraphRequest @p }
    catch {
        $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw "Graph $Method $Uri failed: $detail"
    }
}

try {
    Import-Module Microsoft.Graph.Authentication
    $scopes = @("LifecycleWorkflows-Workflow.ReadWrite.All", "User.Read.All")
    if ($IncludeGroups) { $scopes += @("Group.ReadWrite.All", "User-LifeCycleInfo.ReadWrite.All") }

    $connect = @{ Scopes = $scopes; NoWelcome = $true }
    if ($TenantId)      { $connect.TenantId = $TenantId }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect

    # ---------------------------------------------------------- workflows
    $workflows = @((Invoke-Graph -Uri "$LCW/workflows").value)
    foreach ($name in $WorkflowNames) {
        $wf = $workflows | Where-Object { $_.displayName -eq $name } | Select-Object -First 1
        if (-not $wf) { Write-Host "[SKIP]    '$name' not found" -ForegroundColor Gray; continue }
        if ($PSCmdlet.ShouldProcess($name, "Delete lifecycle workflow")) {
            Invoke-Graph -Method DELETE -Uri "$LCW/workflows/$($wf.id)" | Out-Null
            Write-Host "[DELETED] Workflow '$name' (recoverable for 30 days)" -ForegroundColor Green
        }
    }

    # ---------------------------------------------------------- groups + seeded dates
    if ($IncludeGroups) {
        if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot "jml-config.json" }
        if (-not (Test-Path $ConfigPath)) { throw "Config not found at $ConfigPath, so there's nothing to reverse." }
        $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

        foreach ($prop in $config.groups.PSObject.Properties) {
            if ($PSCmdlet.ShouldProcess($prop.Name, "Delete group")) {
                try {
                    Invoke-Graph -Method DELETE -Uri "$Graph/groups/$($prop.Value)" | Out-Null
                    Write-Host "[DELETED] Group '$($prop.Name)'" -ForegroundColor Green
                } catch {
                    Write-Host "[WARN]    Couldn't delete '$($prop.Name)': $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        $depts = $config.testDepartments
        $users = @((Invoke-Graph -Uri "$Graph/users?`$select=id,displayName,department&`$top=999").value |
                   Where-Object { $depts -contains $_.department })
        foreach ($u in $users) {
            if ($PSCmdlet.ShouldProcess($u.displayName, "Clear lifecycle dates")) {
                Invoke-Graph -Method PATCH -Uri "$Graph/users/$($u.id)" `
                             -Body @{ employeeHireDate = $null; employeeLeaveDateTime = $null } | Out-Null
                Write-Host "[CLEARED] Lifecycle dates on $($u.displayName)" -ForegroundColor Green
            }
        }
        Write-Host "`nNote: accounts disabled or emptied by a leaver run are NOT restored here." -ForegroundColor Yellow
        Write-Host "Re-enable them and rerun create-10-test-users.ps1 / seed-jml-test-data.ps1 to reset." -ForegroundColor Yellow
    }

    Write-Host "`nRollback finished. Verify: entra.microsoft.com > ID Governance > Lifecycle workflows." -ForegroundColor Cyan
}
catch {
    Write-Host "`n[FAILED] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
