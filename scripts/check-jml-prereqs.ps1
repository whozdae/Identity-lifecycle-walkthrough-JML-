<#
.SYNOPSIS
    Read-only prerequisite check for the Identity Lifecycle (JML) lab.

.DESCRIPTION
    Answers four questions before any build work starts:

      1. Can this tenant use Lifecycle Workflows? Lifecycle Workflows requires a
         Microsoft Entra ID Governance or Entra Suite license. Entra ID P2 alone is
         NOT enough, even though access reviews and entitlement management are.
         The check probes the Graph endpoint directly, which is more reliable than
         guessing from SKU names.
      2. What licenses does the tenant actually hold?
      3. Do the test users have the attributes a JML process needs
         (employeeHireDate, employeeLeaveDateTime, department, manager)?
      4. Which build path should the lab take: native Lifecycle Workflows, or a
         scripted JML pipeline that works on any tenant?

    This script only reads. It creates, modifies and deletes nothing.

    Delegated Graph scopes requested:
      LifecycleWorkflows-Workflow.ReadWrite.All  probe Lifecycle Workflows availability
                                                 (falls back to the legacy scope name)
      Organization.Read.All                      read subscribedSkus to list licenses
      User.Read.All                              read test user properties
      User-LifeCycleInfo.ReadWrite.All           read/write employeeHireDate and
                                                 employeeLeaveDateTime, which are
                                                 protected by their own permission

.PARAMETER TestUserDepartments
    Departments whose users are treated as the JML test population.

.EXAMPLE
    .\check-jml-prereqs.ps1

.EXAMPLE
    .\check-jml-prereqs.ps1 -TestUserDepartments "Vendor","Finance"
#>

[CmdletBinding()]
param(
    [string[]]$TestUserDepartments = @("Security", "IT Operations", "Finance", "Human Resources", "Governance", "Vendor"),
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = "Stop"
$Graph = "https://graph.microsoft.com/v1.0"

function Write-Step ($Text) { Write-Host "`n[+] $Text" -ForegroundColor Yellow }
function Write-Ok   ($Text) { Write-Host "    [OK]     $Text" -ForegroundColor Green }
function Write-No   ($Text) { Write-Host "    [NO]     $Text" -ForegroundColor Red }
function Write-Info ($Text) { Write-Host "    [INFO]   $Text" -ForegroundColor Cyan }

function Get-GraphAll ([string]$Uri, [hashtable]$Headers) {
    $items = @()
    $next  = $Uri
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -Headers $Headers
        if ($page.value) { $items += $page.value }
        $next = $page.'@odata.nextLink'
    }
    return $items
}

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Identity Lifecycle (JML) - Prerequisite Check" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

Import-Module Microsoft.Graph.Authentication

# ---------------------------------------------------------------- 1. connect
# The granular Lifecycle Workflows scopes replaced the older single scope. Try the
# current names first and fall back, so an older tenant registration still works.
Write-Step "Connecting to Microsoft Graph"
$baseScopes = @("Organization.Read.All", "User.Read.All", "User-LifeCycleInfo.ReadWrite.All")
$scopeSets  = @(
    @($baseScopes + "LifecycleWorkflows-Workflow.ReadWrite.All"),
    @($baseScopes + "LifecycleWorkflows.ReadWrite.All"),
    @($baseScopes)
)

$connected = $false
foreach ($set in $scopeSets) {
    try {
        $connect = @{ Scopes = $set; NoWelcome = $true }
        if ($TenantId)      { $connect.TenantId = $TenantId }
        if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
        Connect-MgGraph @connect
        $connected = $true
        Write-Ok "Consented scopes: $((Get-MgContext).Scopes -join ', ')"
        break
    } catch {
        Write-Info "Scope set rejected, trying a narrower one. ($($_.Exception.Message -split "`n" | Select-Object -First 1))"
    }
}
if (-not $connected) { throw "Could not sign in with any scope set." }

$context = Get-MgContext
$maskedAccount = $context.Account -replace '^(.{3}).*@.*$', '$1***@***'
Write-Ok "Signed in as $maskedAccount (tenant $($context.TenantId.ToString().Substring(0,8))-****)"

# ---------------------------------------------------------------- 2. probe Lifecycle Workflows
Write-Step "Checking whether Lifecycle Workflows is usable in this tenant"
$lcwUsable = $false
$lcwReason = ""
# No query parameters here on purpose. A malformed query string returns 400, which
# would look like a licensing failure when it is really a bad request.
try {
    $wf = Invoke-MgGraphRequest -Method GET -OutputType PSObject `
          -Uri "$Graph/identityGovernance/lifecycleWorkflows/workflows"
    $lcwUsable = $true
    Write-Ok "Lifecycle Workflows responded. Existing workflows: $(@($wf.value).Count)"
} catch {
    $status    = $null
    $lcwReason = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
    try { $status = [int]$_.Exception.Response.StatusCode } catch { }

    # Distinguish the three failures that look alike: not licensed, not consented,
    # and a request this script got wrong.
    $lcwDiagnosis = switch ($status) {
        403     { if ($lcwReason -match "licen|subscription|sku") { "NOT LICENSED" } else { "PERMISSION NOT CONSENTED" } }
        401     { "NOT AUTHENTICATED" }
        400     { "BAD REQUEST (a bug in this script, not a tenant problem)" }
        default { "UNEXPECTED ($status)" }
    }
    Write-No "Lifecycle Workflows probe failed: $lcwDiagnosis"
    Write-Info "HTTP status: $status"
    Write-Info "Response: $($lcwReason.Substring(0, [Math]::Min(600, $lcwReason.Length)))"
}

# ---------------------------------------------------------------- 3. licenses
Write-Step "Tenant licenses"
$governanceHit = $false
try {
    $skus = @(Get-GraphAll "$Graph/subscribedSkus")
    foreach ($sku in $skus) {
        $enabled  = $sku.prepaidUnits.enabled
        $consumed = $sku.consumedUnits
        Write-Info "$($sku.skuPartNumber)  ($consumed of $enabled assigned)"
        if ($sku.servicePlans | Where-Object { $_.servicePlanName -match "GOVERNANCE|IDENTITY_GOVERNANCE|ENTRA_SUITE" }) {
            $governanceHit = $true
            Write-Ok "  ^ includes a governance service plan"
        }
    }
    if (-not $governanceHit) {
        Write-No "No Entra ID Governance / Entra Suite service plan found in the tenant's SKUs"
    }
} catch {
    Write-No "Could not read subscribedSkus: $($_.Exception.Message)"
}

# ---------------------------------------------------------------- 4. test user attributes
Write-Step "Checking JML attributes on the test population"
$select = "id,displayName,userPrincipalName,department,jobTitle,accountEnabled,employeeHireDate,employeeLeaveDateTime"
$users  = @(Get-GraphAll "$Graph/users?`$select=$select&`$top=999")
$test   = @($users | Where-Object { $TestUserDepartments -contains $_.department })

if ($test.Count -eq 0) {
    Write-No "No users found in departments: $($TestUserDepartments -join ', ')"
} else {
    Write-Ok "Test users found: $($test.Count)"
    $report = foreach ($u in $test) {
        $mgr = $null
        try {
            $mgr = (Invoke-MgGraphRequest -Method GET -OutputType PSObject `
                    -Uri "$Graph/users/$($u.id)/manager?`$select=displayName").displayName
        } catch { $mgr = $null }

        [PSCustomObject]@{
            DisplayName = $u.displayName
            Department  = $u.department
            Enabled     = $u.accountEnabled
            HireDate    = if ($u.employeeHireDate)      { ([datetime]$u.employeeHireDate).ToString("yyyy-MM-dd") } else { "-- missing --" }
            LeaveDate   = if ($u.employeeLeaveDateTime) { ([datetime]$u.employeeLeaveDateTime).ToString("yyyy-MM-dd") } else { "-- missing --" }
            Manager     = if ($mgr) { $mgr } else { "-- missing --" }
        }
    }
    $report | Format-Table -AutoSize

    $noHire  = @($report | Where-Object { $_.HireDate  -eq "-- missing --" }).Count
    $noLeave = @($report | Where-Object { $_.LeaveDate -eq "-- missing --" }).Count
    $noMgr   = @($report | Where-Object { $_.Manager   -eq "-- missing --" }).Count
    Write-Info "Missing employeeHireDate: $noHire | employeeLeaveDateTime: $noLeave | manager: $noMgr"
    Write-Info "These are the attributes a JML process keys on. The build script will set them."

    $exportPath = Join-Path $PSScriptRoot "jml-prereq-report.csv"
    $report | Export-Csv -Path $exportPath -NoTypeInformation -Force
    Write-Ok "Report exported to $exportPath"
}

# ---------------------------------------------------------------- 5. verdict
Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host " VERDICT" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

if (-not $lcwUsable -and $governanceHit) {
    Write-Host @"
 INCONCLUSIVE: the tenant HAS a governance license, but the probe failed.

 That combination means the blocker is not licensing. Check the HTTP status above:
   403 -> consent to LifecycleWorkflows-Workflow.ReadWrite.All, or the signed-in
          account needs the Lifecycle Workflows Administrator role
   400 -> a malformed request from this script; report it
 Send this output to Alfred before building either path.
"@ -ForegroundColor Yellow
} elseif ($lcwUsable) {
    Write-Host @"
 Build path A: NATIVE LIFECYCLE WORKFLOWS

 This tenant can run Lifecycle Workflows, which is the version worth building:
 it's the product SC-300 names, and recruiters recognize it.
 Next: joiner, mover and leaver workflows plus scheduled-run evidence.
"@ -ForegroundColor Green
} else {
    Write-Host @"
 Build path B: SCRIPTED JML PIPELINE (plus a Governance trial, if you want path A)

 Lifecycle Workflows needs an Entra ID Governance or Entra Suite license.
 Two options:
   1. Start the free Microsoft Entra ID Governance trial (Entra admin center >
      Identity Governance), then rerun this script. Path A is the stronger lab.
   2. Build the scripted JML pipeline instead. It proves the same lifecycle
      logic through Graph and runs on any tenant, so it's a legitimate portfolio
      project rather than a consolation prize.
"@ -ForegroundColor Yellow
}

Write-Host " Send this output to Alfred and the build script follows.`n" -ForegroundColor Cyan
