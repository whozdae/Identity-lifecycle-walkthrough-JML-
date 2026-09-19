<#
.SYNOPSIS
    Shows what each lifecycle workflow run actually did, task by task, with failure reasons.

.DESCRIPTION
    A workflow run reports only a count of failed tasks. This pulls the per-task detail
    behind that number: which task failed, when, and why. It's both the debugging tool
    and the evidence artifact for the lab write-up.

    For each workflow it lists every user processed, then every task attempted for that
    user, with status and failure reason. Results are exported to CSV.

    This script only reads. It changes nothing.

    Delegated Graph scopes requested:
      LifecycleWorkflows-Reports.Read.All        read run and task reports
      LifecycleWorkflows-Workflow.Read.All       list the workflows themselves
      User.Read.All                              resolve subject IDs to names

.PARAMETER WorkflowName
    Limit output to one workflow. Default: all of the lab's workflows.

.PARAMETER FailuresOnly
    Show only tasks that did not complete.

.EXAMPLE
    .\show-jml-run-results.ps1

.EXAMPLE
    .\show-jml-run-results.ps1 -WorkflowName "Onboard pre-hire employee" -FailuresOnly
#>

[CmdletBinding()]
param(
    [string]$WorkflowName,
    [switch]$FailuresOnly,
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = "Stop"
$Graph = "https://graph.microsoft.com/v1.0"
$LCW   = "$Graph/identityGovernance/lifecycleWorkflows"

function Write-Step ($Text) { Write-Host "`n[+] $Text" -ForegroundColor Yellow }
function Write-Ok   ($Text) { Write-Host "    [OK]     $Text" -ForegroundColor Green }
function Write-No   ($Text) { Write-Host "    [FAIL]   $Text" -ForegroundColor Red }
function Write-Info ($Text) { Write-Host "    [INFO]   $Text" -ForegroundColor Cyan }

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

try {
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host " JML Lab - Workflow run results" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    Import-Module Microsoft.Graph.Authentication
    $connect = @{
        Scopes = @("LifecycleWorkflows-Reports.Read.All", "LifecycleWorkflows-Workflow.Read.All", "User.Read.All")
        NoWelcome = $true
    }
    if ($TenantId)      { $connect.TenantId = $TenantId }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect
    Write-Ok "Signed in as $((Get-MgContext).Account -replace '^(.{3}).*@.*$', '$1***@***')"

    $workflows = @(Get-GraphAll "$LCW/workflows")
    if ($WorkflowName) { $workflows = @($workflows | Where-Object { $_.displayName -eq $WorkflowName }) }
    if ($workflows.Count -eq 0) { throw "No workflows found." }

    $userCache = @{}
    function Resolve-Name ($UserId) {
        if (-not $UserId) { return "(unknown)" }
        if ($userCache.ContainsKey($UserId)) { return $userCache[$UserId] }
        $name = try { (Invoke-Graph "$Graph/users/$UserId`?`$select=displayName").displayName } catch { $UserId }
        $userCache[$UserId] = $name
        return $name
    }

    $rows = @()

    foreach ($wf in $workflows) {
        Write-Step "$($wf.category.ToUpper()): $($wf.displayName)"

        $userResults = @(Get-GraphAll "$LCW/workflows/$($wf.id)/userProcessingResults")
        if ($userResults.Count -eq 0) {
            Write-Info "No runs yet. Use -RunNow on the deploy script, or -EnableScheduling and wait."
            continue
        }

        foreach ($ur in $userResults) {
            $subject = Resolve-Name $ur.subject.id
            $status  = $ur.processingStatus
            $line    = "$subject : $status ($($ur.failedTasksCount) of $($ur.totalTasksCount) tasks failed)"
            if ($status -eq "completed" -and $ur.failedTasksCount -eq 0) { Write-Ok $line } else { Write-No $line }

            $tasks = @(Get-GraphAll "$LCW/workflows/$($wf.id)/userProcessingResults/$($ur.id)/taskProcessingResults")
            foreach ($t in $tasks) {
                $taskStatus = $t.processingStatus
                if ($FailuresOnly -and $taskStatus -eq "completed") { continue }

                $marker = if ($taskStatus -eq "completed") { "[ok]  " } else { "[FAIL]" }
                $colour = if ($taskStatus -eq "completed") { "Green" } else { "Red" }
                Write-Host "        $marker $($t.task.displayName) - $taskStatus" -ForegroundColor $colour
                if ($t.failureReason) {
                    Write-Host "               reason: $($t.failureReason)" -ForegroundColor Yellow
                }

                $rows += [PSCustomObject]@{
                    Workflow      = $wf.displayName
                    Category      = $wf.category
                    Subject       = $subject
                    UserStatus    = $status
                    Task          = $t.task.displayName
                    TaskStatus    = $taskStatus
                    FailureReason = $t.failureReason
                    CompletedUtc  = $t.completedDateTime
                }
            }
        }
    }

    if ($rows.Count -gt 0) {
        $exportPath = Join-Path $PSScriptRoot "jml-run-results.csv"
        try {
            $rows | Export-Csv -Path $exportPath -NoTypeInformation -Force
            Write-Ok "`nExported $($rows.Count) task results to $exportPath"
        } catch {
            $exportPath = Join-Path $PSScriptRoot "jml-run-results-$(Get-Date -Format yyyyMMdd-HHmmss).csv"
            $rows | Export-Csv -Path $exportPath -NoTypeInformation -Force
            Write-Info "Original CSV was locked; wrote $exportPath"
        }

        $failed = @($rows | Where-Object { $_.TaskStatus -ne "completed" })
        Write-Host "`n----------------------------------------------------------" -ForegroundColor Cyan
        Write-Host " $($rows.Count) tasks run, $($failed.Count) not completed" -ForegroundColor Cyan
        Write-Host "----------------------------------------------------------" -ForegroundColor Cyan
        if ($failed.Count -gt 0) {
            $failed | Select-Object Task, TaskStatus, FailureReason -Unique | Format-List
        }
    }
}
catch {
    Write-Host "`n[FAILED] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
