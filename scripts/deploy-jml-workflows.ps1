<#
.SYNOPSIS
    Creates joiner, mover and leaver Lifecycle Workflows in Microsoft Entra ID.

.DESCRIPTION
    Builds the three workflows that make up a joiner-mover-leaver process:

      JOINER  "Onboard pre-hire employee"
              Fires 7 days before employeeHireDate.
              Enables the account, generates a Temporary Access Pass and emails it to
              the manager, adds the user to SG-All-Employees, sends a welcome email.

      MOVER   "Department transfer - access update"
              Fires when the department attribute changes.
              Emails the manager, revokes refresh tokens so the old session can't keep
              stale access, removes the old department group, adds the new one.

      LEAVER  "Offboard on last day"
              Fires on employeeLeaveDateTime.
              Disables the account, revokes refresh tokens, removes every group
              membership, removes all license assignments.

    Task definition IDs are NOT hardcoded. The script reads the tenant's live task
    catalog and resolves each task by name, so a renamed or re-issued GUID surfaces as
    a clear error instead of a failed workflow run.

    Requires the Entra ID Governance or Entra Suite license (check-jml-prereqs.ps1
    confirms this) and the Lifecycle Workflows Administrator role.

    Delegated Graph scopes requested:
      LifecycleWorkflows-Workflow.ReadWrite.All  create and read workflows
      LifecycleWorkflows-Workflow.Activate       run a workflow on demand (-RunNow)
      Group.Read.All                             resolve the lab's group IDs
      User.Read.All                              resolve the demo subject

.PARAMETER ConfigPath
    Path to jml-config.json, written by seed-jml-test-data.ps1.

.PARAMETER Update
    Push task changes into workflows that already exist, via createNewVersion. The
    version number increments and prior run history is preserved.

.PARAMETER EnableScheduling
    Let the workflows run on Entra's own schedule (every 3 hours). Off by default so
    nothing fires until you've captured your evidence.

.PARAMETER RunNow
    After deployment, run the joiner workflow on demand against the joiner from the
    config, so there's a run to screenshot without waiting for the scheduler.

.EXAMPLE
    .\deploy-jml-workflows.ps1 -WhatIf

.EXAMPLE
    .\deploy-jml-workflows.ps1 -RunNow

.EXAMPLE
    .\deploy-jml-workflows.ps1 -EnableScheduling

.NOTES
    ROLLBACK: .\remove-jml-workflows.ps1
    Verify in Entra: ID Governance > Lifecycle workflows.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ConfigPath,
    [switch]$EnableScheduling,
    [switch]$Update,
    [switch]$RunNow,
    [string]$TenantId,
    [switch]$UseDeviceCode
)

$ErrorActionPreference = "Stop"
$Graph = "https://graph.microsoft.com/v1.0"
$LCW   = "$Graph/identityGovernance/lifecycleWorkflows"

$logDir = Join-Path $PSScriptRoot "logs"
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -WhatIf:$false | Out-Null }
Start-Transcript -Path (Join-Path $logDir "deploy-workflows-$(Get-Date -Format yyyyMMdd-HHmmss).log") -WhatIf:$false | Out-Null

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
    Write-Host " JML Lab - Deploy Lifecycle Workflows" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    # ------------------------------------------------------------ config
    if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot "jml-config.json" }
    if (-not (Test-Path $ConfigPath)) { throw "Config not found at $ConfigPath. Run seed-jml-test-data.ps1 first." }
    $config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    Write-Ok "Loaded config from $ConfigPath"

    Import-Module Microsoft.Graph.Authentication
    $connect = @{
        Scopes = @(
            "LifecycleWorkflows-Workflow.ReadWrite.All",
            "LifecycleWorkflows-Workflow.Activate",
            "Group.Read.All",
            "User.Read.All"
        )
        NoWelcome = $true
    }
    if ($TenantId)      { $connect.TenantId = $TenantId }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect
    $ctx = Get-MgContext
    Write-Ok "Signed in as $($ctx.Account -replace '^(.{3}).*@.*$', '$1***@***')"

    # ------------------------------------------------------------ task catalog
    # Resolving by name against the live catalog beats hardcoding GUIDs: if Microsoft
    # renames or re-issues one, this fails loudly here instead of at run time.
    Write-Step "Reading the tenant's lifecycle task catalog"
    $taskDefs = @(Get-GraphAll "$LCW/taskDefinitions")
    Write-Ok "Task definitions available: $($taskDefs.Count)"

    # Patterns are the exact display names this tenant reports. Microsoft's published
    # task lists disagree with the live catalog (it is "Generate TAP and send email",
    # not "Temporary Access Pass"), so the catalog is the only source of truth.
    function Get-TaskId ([string]$Pattern) {
        $hits = @($taskDefs | Where-Object { $_.displayName -like $Pattern })
        if ($hits.Count -eq 0) {
            throw "No lifecycle task matches '$Pattern'. Available: $(($taskDefs.displayName | Sort-Object) -join '; ')"
        }
        if ($hits.Count -gt 1) {
            throw "Pattern '$Pattern' matched $($hits.Count) tasks: $(($hits.displayName) -join '; '). Narrow it."
        }
        Write-Info "$($hits[0].displayName) -> $($hits[0].id)"
        return $hits[0].id
    }

    function New-Task ([string]$Pattern, [string]$DisplayName, [string]$Description, [hashtable]$Arguments, [bool]$ContinueOnError = $false) {
        # Named taskArgs because $args is an automatic variable inside a function.
        $taskArgs = @()
        if ($Arguments) { foreach ($k in $Arguments.Keys) { $taskArgs += @{ name = $k; value = [string]$Arguments[$k] } } }
        return @{
            continueOnError  = $ContinueOnError
            description      = $Description
            displayName      = $DisplayName
            isEnabled        = $true
            taskDefinitionId = (Get-TaskId $Pattern)
            arguments        = $taskArgs
        }
    }

    # ------------------------------------------------------------ group IDs
    $allEmployees = $config.groups."SG-All-Employees"
    $fromGroup    = $config.groups."SG-Dept-$($config.mover.fromDepartment -replace '\s','')"
    $toGroup      = $config.groups."SG-Dept-$($config.mover.toDepartment   -replace '\s','')"
    if (-not $allEmployees) { throw "SG-All-Employees is missing from the config. Rerun the seed script." }
    Write-Info "Move: $($config.mover.fromDepartment) -> $($config.mover.toDepartment)"

    # Scope every workflow to the lab's departments only, so nothing touches a real account.
    $scopeRule = ($config.testDepartments | ForEach-Object { "(department eq '$_')" }) -join " or "

    # ------------------------------------------------------------ workflow definitions
    Write-Step "Building workflow definitions"

    $joiner = @{
        category    = "joiner"
        displayName = "Onboard pre-hire employee"
        description = "Seven days before the hire date: enable the account, issue a Temporary Access Pass to the manager, grant baseline access, and welcome the new hire."
        isEnabled   = $true
        isSchedulingEnabled = [bool]$EnableScheduling
        executionConditions = @{
            "@odata.type" = "#microsoft.graph.identityGovernance.triggerAndScopeBasedConditions"
            scope   = @{ "@odata.type" = "#microsoft.graph.identityGovernance.ruleBasedSubjectSet"; rule = $scopeRule }
            trigger = @{
                "@odata.type"      = "#microsoft.graph.identityGovernance.timeBasedAttributeTrigger"
                timeBasedAttribute = "employeeHireDate"
                offsetInDays       = -7
            }
        }
        # Order matters: access first, notifications last. A mail failure must not stop
        # provisioning, so both email-sending tasks continue on error.
        tasks = @(
            (New-Task "Enable user account" "Enable the account" "Turn on the account before day one."),
            (New-Task "Add user to groups" "Grant baseline access" "Group-based access instead of direct assignment." @{ groupID = $allEmployees }),
            (New-Task "Generate TAP and send email" "Issue a Temporary Access Pass" "Passwordless first sign-in; the TAP goes to the manager, not the new hire. Needs a mailbox on the manager." @{
                tapLifetimeMinutes = 480
                tapIsUsableOnce    = "false"
            } $true),
            (New-Task "Send welcome email" "Send the welcome email" "Confirms onboarding ran. Needs a mailbox." $null $true)
        )
    }

    $mover = @{
        category    = "mover"
        displayName = "Department transfer - access update"
        description = "When department changes: notify the manager, revoke existing sessions, and swap department group membership."
        isEnabled   = $true
        isSchedulingEnabled = [bool]$EnableScheduling
        executionConditions = @{
            "@odata.type" = "#microsoft.graph.identityGovernance.triggerAndScopeBasedConditions"
            scope   = @{ "@odata.type" = "#microsoft.graph.identityGovernance.ruleBasedSubjectSet"; rule = $scopeRule }
            trigger = @{
                "@odata.type"     = "#microsoft.graph.identityGovernance.attributeChangeTrigger"
                triggerAttributes = @(@{ name = "department" })
            }
        }
        tasks = @(
            (New-Task "Send email to notify manager of user move" "Notify the manager" "The receiving manager should know access changed. Needs a mailbox." $null $true),
            (New-Task "Revoke all refresh tokens for user" "Revoke active sessions" "Without this, an existing session keeps the old access until the token expires."),
            (New-Task "Remove user from selected groups" "Remove the previous department's access" "Leaving old access in place is how privilege creep starts." @{ groupID = $fromGroup } $true),
            (New-Task "Add user to groups" "Grant the new department's access" "" @{ groupID = $toGroup })
        )
    }

    $leaver = @{
        category    = "leaver"
        displayName = "Offboard on last day"
        description = "On the last day: disable the account, revoke sessions, strip all group memberships and licenses."
        isEnabled   = $true
        isSchedulingEnabled = [bool]$EnableScheduling
        executionConditions = @{
            "@odata.type" = "#microsoft.graph.identityGovernance.triggerAndScopeBasedConditions"
            scope   = @{ "@odata.type" = "#microsoft.graph.identityGovernance.ruleBasedSubjectSet"; rule = $scopeRule }
            trigger = @{
                "@odata.type"      = "#microsoft.graph.identityGovernance.timeBasedAttributeTrigger"
                timeBasedAttribute = "employeeLeaveDateTime"
                offsetInDays       = 0
            }
        }
        tasks = @(
            (New-Task "Disable user account" "Disable the account" "Sign-in stops immediately; the account is kept for any investigation."),
            (New-Task "Revoke all refresh tokens for user" "Revoke active sessions" "A disabled account with a live token can still reach resources until it expires."),
            (New-Task "Remove user from all groups" "Remove all group memberships" "" $null $true),
            (New-Task "Remove all licenses for user" "Remove all licenses" "Stops paying for a departed worker." $null $true)
        )
    }

    # ------------------------------------------------------------ deploy
    $existing = @(Get-GraphAll "$LCW/workflows")
    $deployed = @()

    foreach ($wf in @($joiner, $mover, $leaver)) {
        Write-Step "$($wf.category.ToUpper()): '$($wf.displayName)'"
        $match = $existing | Where-Object { $_.displayName -eq $wf.displayName } | Select-Object -First 1
        if ($match -and $Update) {
            if ($PSCmdlet.ShouldProcess($wf.displayName, "Update tasks on the existing $($wf.category) workflow")) {
                # PATCH on a workflow only accepts displayName, description, isEnabled and
                # isSchedulingEnabled. Changing tasks requires createNewVersion, which bumps
                # the version and keeps prior run history (delete-and-recreate would lose it).
                $new = Invoke-Graph -Method POST -Uri "$LCW/workflows/$($match.id)/createNewVersion" `
                                    -Body @{ workflow = $wf }
                Write-Ok "New version $($new.version) created ($($wf.tasks.Count) tasks, id $($match.id))"
                $deployed += [PSCustomObject]@{ Category = $wf.category; Name = $wf.displayName; Id = $match.id; Tasks = $wf.tasks.Count; State = "updated" }
            }
        } elseif ($match) {
            Write-Skip "Already exists (id $($match.id)). Use -Update to push task changes."
            $deployed += [PSCustomObject]@{ Category = $wf.category; Name = $wf.displayName; Id = $match.id; Tasks = $wf.tasks.Count; State = "existing" }
        } elseif ($PSCmdlet.ShouldProcess($wf.displayName, "Create $($wf.category) workflow")) {
            $new = Invoke-Graph -Method POST -Uri "$LCW/workflows" -Body $wf
            Write-Ok "Created with $($wf.tasks.Count) tasks (id $($new.id))"
            $deployed += [PSCustomObject]@{ Category = $wf.category; Name = $wf.displayName; Id = $new.id; Tasks = $wf.tasks.Count; State = "created" }
        }
    }

    # ------------------------------------------------------------ optional on-demand run
    if ($RunNow) {
        Write-Step "Running the joiner workflow on demand against $($config.joiner.displayName)"
        $joinerWf = $deployed | Where-Object { $_.Category -eq "joiner" } | Select-Object -First 1
        if ($joinerWf -and $PSCmdlet.ShouldProcess($config.joiner.displayName, "Run '$($joinerWf.Name)' now")) {
            Invoke-Graph -Method POST -Uri "$LCW/workflows/$($joinerWf.Id)/activate" `
                         -Body @{ subjects = @(@{ id = $config.joiner.id }) } | Out-Null
            Write-Ok "Activated. Results take a minute or two to appear."
            Start-Sleep -Seconds 45
            try {
                $results = @(Get-GraphAll "$LCW/workflows/$($joinerWf.Id)/userProcessingResults")
                if ($results.Count -gt 0) {
                    $results | Select-Object -First 5 processingStatus, completedDateTime, totalTasksCount, failedTasksCount | Format-Table -AutoSize
                } else {
                    Write-Info "No results yet. Check Entra: ID Governance > Lifecycle workflows > the workflow > Workflow history."
                }
            } catch {
                Write-Info "Couldn't read results yet: $($_.Exception.Message)"
            }
        }
    }

    # ------------------------------------------------------------ verify
    Write-Step "Deployed workflows"
    $deployed | Format-Table -AutoSize
    $reportPath = Join-Path $PSScriptRoot "jml-workflows-report.csv"
    try {
        $deployed | Export-Csv -Path $reportPath -NoTypeInformation -Force
        Write-Ok "Report written to $reportPath"
    } catch {
        Write-Info "Couldn't write the report (file open?): $($_.Exception.Message)"
    }

    if (-not $EnableScheduling) {
        Write-Info "Scheduling is OFF. The workflows exist but won't fire on their own."
        Write-Info "Turn it on with -EnableScheduling once you've captured the 'before' evidence."
    } else {
        Write-Info "Scheduling is ON. Entra evaluates workflows roughly every 3 hours."
    }

    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host " Verify: entra.microsoft.com > ID Governance > Lifecycle workflows" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan
}
catch {
    Write-Host "`n[FAILED] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Anything created before this point is kept. Fix and rerun; existing workflows are reused." -ForegroundColor Red
    $exitCode = 1
}
finally {
    try { Stop-Transcript | Out-Null } catch { }
}
if ($exitCode) { exit $exitCode }
