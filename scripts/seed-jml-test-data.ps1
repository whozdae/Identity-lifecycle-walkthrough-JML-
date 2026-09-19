<#
.SYNOPSIS
    Seeds the lifecycle attributes and groups the JML lab needs, then writes a config
    file the workflow deployment script reads.

.DESCRIPTION
    Lifecycle Workflows key off attributes that none of the test users currently have.
    This script sets them up:

      1. Managers. Every task that emails "the manager" silently does nothing without
         one, so each test user gets a department lead, and the leads report to the
         Governance lead.
      2. employeeHireDate. Existing staff get a backdated hire date. One designated
         joiner gets a hire date 7 days out, so the pre-hire workflow has something
         real to fire against.
      3. employeeLeaveDateTime. One designated leaver gets a last day of today, so the
         offboarding workflow triggers on its next scheduled pass.
      4. Groups: SG-All-Employees (granted at onboarding) and a department group per
         test department (swapped during a department transfer).
      5. jml-config.json, holding the group and user IDs the deployment script needs,
         so no IDs are hardcoded in either script.

    Delegated Graph scopes requested:
      User.ReadWrite.All                directory-wide: assign managers
      User-LifeCycleInfo.ReadWrite.All  employeeHireDate and employeeLeaveDateTime are
                                        protected by their own permission
      Group.ReadWrite.All               create the lab's groups
      Directory.Read.All                read department and manager relationships

.PARAMETER JoinerName
    Display name of the user treated as the incoming hire. Default: "Sarah Jenkins".

.PARAMETER LeaverName
    Display name of the user treated as the departing worker. Default: "Viktor Novak".

.PARAMETER MoverName
    Display name of the user whose department transfer drives the mover workflow.
    Default: "Carlos Mendoza" (Finance, transferring to Security).

.EXAMPLE
    .\seed-jml-test-data.ps1 -WhatIf

.EXAMPLE
    .\seed-jml-test-data.ps1

.NOTES
    ROLLBACK: .\remove-jml-workflows.ps1 -IncludeGroups clears the groups and the
    lifecycle dates this script sets.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$JoinerName = "Sarah Jenkins",
    [string]$LeaverName = "Viktor Novak",
    [string]$MoverName  = "Carlos Mendoza",
    [int]$JoinerStartsInDays = 7,
    [string[]]$TestDepartments = @("Security", "IT Operations", "Finance", "Human Resources", "Governance", "Vendor"),
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = "Stop"
$Graph = "https://graph.microsoft.com/v1.0"

# Department -> the test user who leads it. Everyone in a department reports to its lead.
$DepartmentLeads = @{
    "Security"        = "Marcus Vance"
    "IT Operations"   = "David Chen"
    "Finance"         = "Ananya Sharma"
    "Human Resources" = "Rachel Green"
    "Governance"      = "Tariq Al-Mansoor"
    "Vendor"          = "Tariq Al-Mansoor"
}
$TopManager = "Marcus Vance"   # the leads' own manager, so no one is left without one

$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -WhatIf:$false | Out-Null }
Start-Transcript -Path (Join-Path $logDir "seed-$(Get-Date -Format yyyyMMdd-HHmmss).log") -WhatIf:$false | Out-Null

function Write-Step ($Text) { Write-Host "`n[+] $Text" -ForegroundColor Yellow }
function Write-Ok   ($Text) { Write-Host "    [OK]     $Text" -ForegroundColor Green }
function Write-Skip ($Text) { Write-Host "    [EXISTS] $Text" -ForegroundColor Gray }
function Write-Info ($Text) { Write-Host "    [INFO]   $Text" -ForegroundColor Cyan }

function Invoke-Graph {
    param(
        [ValidateSet("GET","POST","PATCH","PUT","DELETE")] [string]$Method = "GET",
        [Parameter(Mandatory)] [string]$Uri,
        $Body
    )
    $p = @{ Method = $Method; Uri = $Uri; OutputType = "PSObject" }
    if ($Body) { $p.Body = ($Body | ConvertTo-Json -Depth 20); $p.ContentType = "application/json" }
    try { Invoke-MgGraphRequest @p }
    catch {
        $detail = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
        throw "Graph $Method $Uri failed: $detail"
    }
}

function Get-GraphAll ([string]$Uri) {
    $items = @(); $next = $Uri
    while ($next) {
        $page = Invoke-Graph -Uri $next
        if ($page.value) { $items += $page.value }
        $next = $page.'@odata.nextLink'
    }
    return $items
}

try {
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host " JML Lab - Seed lifecycle attributes and groups" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    Import-Module Microsoft.Graph.Authentication
    $connect = @{
        Scopes = @("User.ReadWrite.All", "User-LifeCycleInfo.ReadWrite.All", "Group.ReadWrite.All", "Directory.Read.All")
        NoWelcome = $true
    }
    if ($TenantId)      { $connect.TenantId = $TenantId }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect
    $ctx = Get-MgContext
    Write-Ok "Signed in as $($ctx.Account -replace '^(.{3}).*@.*$', '$1***@***')"

    # ------------------------------------------------------------ users
    Write-Step "Loading test users"
    $select = "id,displayName,userPrincipalName,department,accountEnabled,employeeHireDate,employeeLeaveDateTime"
    $all  = @(Get-GraphAll "$Graph/users?`$select=$select&`$top=999")
    $test = @($all | Where-Object { $TestDepartments -contains $_.department })
    if ($test.Count -eq 0) { throw "No users found in the test departments. Run create-10-test-users.ps1 first." }
    Write-Ok "Test users: $($test.Count)"

    function Find-TestUser ($Name) {
        $u = $test | Where-Object { $_.displayName -eq $Name } | Select-Object -First 1
        if (-not $u) { throw "Test user '$Name' not found." }
        return $u
    }
    $joiner = Find-TestUser $JoinerName
    $leaver = Find-TestUser $LeaverName
    $mover  = Find-TestUser $MoverName
    Write-Info "Joiner: $($joiner.displayName) | Mover: $($mover.displayName) | Leaver: $($leaver.displayName)"

    # ------------------------------------------------------------ managers
    Write-Step "Assigning managers"
    foreach ($u in $test) {
        $leadName = if ($DepartmentLeads.ContainsKey($u.department)) { $DepartmentLeads[$u.department] } else { $TopManager }
        if ($u.displayName -eq $leadName) { $leadName = $TopManager }      # leads report upward
        if ($u.displayName -eq $TopManager) { Write-Skip "$($u.displayName) is the top of the chain"; continue }

        $lead = $test | Where-Object { $_.displayName -eq $leadName } | Select-Object -First 1
        if (-not $lead) { Write-Info "No lead found for $($u.displayName), skipping"; continue }

        $current = $null
        try { $current = (Invoke-Graph -Uri "$Graph/users/$($u.id)/manager?`$select=id,displayName").displayName } catch { }
        if ($current -eq $lead.displayName) {
            Write-Skip "$($u.displayName) already reports to $current"
        } elseif ($PSCmdlet.ShouldProcess($u.displayName, "Set manager to $($lead.displayName)")) {
            Invoke-Graph -Method PUT -Uri "$Graph/users/$($u.id)/manager/`$ref" `
                         -Body @{ "@odata.id" = "$Graph/users/$($lead.id)" } | Out-Null
            Write-Ok "$($u.displayName) -> $($lead.displayName)"
        }
    }

    # ------------------------------------------------------------ lifecycle dates
    Write-Step "Setting lifecycle dates"
    $joinerHire = (Get-Date).Date.AddDays($JoinerStartsInDays)
    $leaverLast = (Get-Date).Date
    $backdated  = (Get-Date).Date.AddYears(-1)

    foreach ($u in $test) {
        $hire = if ($u.id -eq $joiner.id) { $joinerHire } else { $backdated }
        $body = @{ employeeHireDate = $hire.ToString("yyyy-MM-ddTHH:mm:ssZ") }

        if ($u.id -eq $leaver.id) {
            $body.employeeLeaveDateTime = $leaverLast.ToString("yyyy-MM-ddTHH:mm:ssZ")
        }

        if ($PSCmdlet.ShouldProcess($u.displayName, "Set lifecycle dates")) {
            Invoke-Graph -Method PATCH -Uri "$Graph/users/$($u.id)" -Body $body | Out-Null
            $note = if ($u.id -eq $joiner.id) { "hire $($hire.ToString('yyyy-MM-dd')) (incoming hire)" }
                    elseif ($u.id -eq $leaver.id) { "hire $($hire.ToString('yyyy-MM-dd')), last day $($leaverLast.ToString('yyyy-MM-dd'))" }
                    else { "hire $($hire.ToString('yyyy-MM-dd'))" }
            Write-Ok "$($u.displayName): $note"
        }
    }

    # ------------------------------------------------------------ groups
    Write-Step "Creating lab groups"
    $groupNames = @("SG-All-Employees") + ($TestDepartments | ForEach-Object { "SG-Dept-$($_ -replace '\s','')" })
    $groups = @{}

    foreach ($name in $groupNames) {
        $existing = @(Get-GraphAll "$Graph/groups?`$filter=displayName eq '$name'&`$select=id,displayName") | Select-Object -First 1
        if ($existing) {
            $groups[$name] = $existing.id
            Write-Skip "$name"
        } elseif ($PSCmdlet.ShouldProcess($name, "Create security group")) {
            $g = Invoke-Graph -Method POST -Uri "$Graph/groups" -Body @{
                displayName     = $name
                description     = "JML lab group. Membership is granted and removed by lifecycle workflows."
                mailEnabled     = $false
                mailNickname    = ($name -replace "[^A-Za-z0-9]", "").ToLower()
                securityEnabled = $true
            }
            $groups[$name] = $g.id
            Write-Ok "Created $name"
        }
    }

    # ------------------------------------------------------------ config for the deploy script
    Write-Step "Writing jml-config.json"
    $config = [ordered]@{
        generatedUtc   = (Get-Date).ToUniversalTime().ToString("o")
        testDepartments = $TestDepartments
        groups         = $groups
        joiner         = @{ id = $joiner.id; displayName = $joiner.displayName; hireDate = $joinerHire.ToString("yyyy-MM-dd") }
        mover          = @{ id = $mover.id;  displayName = $mover.displayName;  fromDepartment = $mover.department; toDepartment = "Security" }
        leaver         = @{ id = $leaver.id; displayName = $leaver.displayName; lastDay = $leaverLast.ToString("yyyy-MM-dd") }
    }
    $configPath = Join-Path $PSScriptRoot "jml-config.json"
    if ($PSCmdlet.ShouldProcess($configPath, "Write config")) {
        $config | ConvertTo-Json -Depth 10 | Set-Content -Path $configPath -Encoding UTF8
        Write-Ok "Wrote $configPath"
    }

    # ------------------------------------------------------------ verify
    Write-Step "Verification"
    $verify = foreach ($u in @($joiner, $mover, $leaver)) {
        $fresh = Invoke-Graph -Uri "$Graph/users/$($u.id)?`$select=$select"
        $mgr = $null
        try { $mgr = (Invoke-Graph -Uri "$Graph/users/$($u.id)/manager?`$select=displayName").displayName } catch { }
        [PSCustomObject]@{
            Role      = if ($u.id -eq $joiner.id) { "Joiner" } elseif ($u.id -eq $mover.id) { "Mover" } else { "Leaver" }
            User      = $fresh.displayName
            Dept      = $fresh.department
            HireDate  = if ($fresh.employeeHireDate)      { ([datetime]$fresh.employeeHireDate).ToString("yyyy-MM-dd") } else { "-- missing --" }
            LastDay   = if ($fresh.employeeLeaveDateTime) { ([datetime]$fresh.employeeLeaveDateTime).ToString("yyyy-MM-dd") } else { "n/a" }
            Manager   = if ($mgr) { $mgr } else { "-- missing --" }
        }
    }
    $verify | Format-Table -AutoSize

    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host " Seeding complete. Next: .\deploy-jml-workflows.ps1" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan
}
catch {
    Write-Host "`n[FAILED] $($_.Exception.Message)" -ForegroundColor Red
    $exitCode = 1
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
}
if ($exitCode) { exit $exitCode }
