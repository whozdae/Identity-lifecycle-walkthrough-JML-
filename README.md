# Automated Joiner-Mover-Leaver Lifecycle in Microsoft Entra ID

I automated the three moments where access is created, changed and revoked — and proved the offboarding half: a departing vendor's **3 group memberships went to 0 and the account was disabled, with no administrator touching it**.

`Microsoft Entra ID Governance` `Lifecycle Workflows` `Microsoft Graph API` `PowerShell 7` `Joiner-Mover-Leaver` `Temporary Access Pass` `NIST 800-53 AC-2`

---

## Problem Statement

Access is granted at hire, changed at every internal move, and is supposed to disappear on the last day. In practice only the first one is reliable: someone is motivated to get a new hire working. Nobody is motivated to strip the access of a person who already left, and nobody is motivated to remove the *old* access of someone who transferred, so entitlements accumulate. That accumulation is what an attacker inherits when a dormant account is compromised, and it's what gets written up under account management in an audit.

I built the joiner, mover and leaver halves as automation instead of as a checklist someone remembers.

## Objective

1. Drive each lifecycle stage from an attribute rather than a ticket: `employeeHireDate`, a department change, `employeeLeaveDateTime`.
2. Make offboarding remove everything — sign-in, sessions, groups, licenses — without human action.
3. Build it through the Graph API so it's reproducible and reviewable, not clicked together.
4. Prove each stage with before/after evidence rather than screenshots of configuration.

## Tools & Environment

| Component | Detail |
|---|---|
| Platform | Microsoft Entra ID Governance (Lifecycle Workflows) + Entra ID P2 |
| Automation | PowerShell 7, `Invoke-MgGraphRequest` against Microsoft Graph v1.0 |
| Module | `Microsoft.Graph.Authentication` only |
| Graph scopes | `LifecycleWorkflows-Workflow.ReadWrite.All`, `LifecycleWorkflows-Workflow.Activate`, `Group.Read.All`, `User.Read.All` |
| Test population | 10 identities across 6 departments, including 2 vendors |

**Licensing note:** Lifecycle Workflows requires the **Entra ID Governance** or Entra Suite license. Entra ID P2 alone is not enough, even though access reviews and entitlement management are included in P2. My prerequisite script checks this before anything is built.

## Step-by-Step Breakdown

### 1. Checked what the tenant could actually do
I wrote a read-only prerequisite script that probes the Lifecycle Workflows endpoint directly, lists the tenant's SKUs, and reports which lifecycle attributes the test users are missing. It distinguishes three failures that look identical from the outside: not licensed, not consented, and a malformed request.

### 2. Seeded the attributes the automation keys on
All 10 users were missing `employeeHireDate`, `employeeLeaveDateTime` and a manager. Every task that emails "the manager" silently does nothing without one, so the seed script assigns a department lead to each user, backdates hire dates for existing staff, and sets up three cases: an incoming hire starting in 7 days, a transfer, and a vendor whose last day is today.

### 3. Built the three workflows

| Stage | Trigger | Tasks in order |
|---|---|---|
| **Joiner** | 7 days before `employeeHireDate` | Enable account → add to `SG-All-Employees` → issue a Temporary Access Pass to the manager → welcome email |
| **Mover** | `department` attribute changes | Notify manager → revoke refresh tokens → remove old department group → add new department group |
| **Leaver** | On `employeeLeaveDateTime` | Disable account → revoke refresh tokens → remove all group memberships → remove all licenses |

**Revoking refresh tokens appears in both mover and leaver deliberately.** Disabling an account does not kill a live session; the existing token keeps working until it expires. Disabling is deactivating the badge, revoking is walking the person out of the building.

### 4. Resolved task IDs from the tenant, never hardcoded
Each task is identified by a GUID. Published task lists disagree with what tenants actually return — the Temporary Access Pass task is named **"Generate TAP and send email"**, not "Temporary Access Pass". So the deployment script reads the live task catalog (32 definitions) and resolves each task by name, failing at deploy time with the full catalog printed rather than producing a workflow that breaks mid-run.

### 5. Proved each stage with state snapshots
A snapshot script records every test user's enabled state, department, manager, license count and full group list to CSV. Run it before, run the workflow, run it after, and it diffs the two. It refuses to overwrite an existing snapshot without `-Force`, because a before-state can't be recaptured once the workflow has run.

## Visual Evidence

| Screenshot | What it shows |
|---|---|
| `01-workflows-deployed.png` | Three workflows, one per lifecycle stage, each at version 2 |
| `02-joiner-run-history.png` | v1 `failed` beside v2 `completedWithErrors` — the cancel-cascade fix |
| `03-leaver-task-results.png` | Every leaver task completed, applied by the service |
| `04-leaver-diff.png` | `3 groups → 0` for the departing vendor |
| `05-mover-diff.png` | Department `Finance → Security`, new department group granted |

## Final Results

| Measure | Outcome |
|---|---|
| Lifecycle stages automated | 3 (joiner, mover, leaver), 12 tasks total |
| Departing vendor's group memberships | **3 → 0**, automatically |
| Account disable, session revocation | Completed by the workflow, no admin action |
| Joiner provisioning | Account enabled and baseline access granted on an on-demand run |
| Mover | Department change granted the new department's access |
| Administrator actions at execution time | **0** — every run was triggered by attribute or on demand, then ran unattended |
| Graph permissions | 4 delegated scopes, no application permissions, no stored secrets |

### What I could not prove in this tenant, and why

Honest limits, because a portfolio that only shows green checkmarks isn't credible:

- **Every email task fails.** The tenant holds Entra ID Governance and P2 but no Exchange licenses, so no user has a mailbox or a `mail` attribute. Tasks fail with "the mail attribute was missing for all of the provided email recipients". The tasks are configured correctly; the environment cannot deliver mail. The leaver workflow is the only one with no email dependency, which is exactly why it's the one that runs completely clean.
- **The mover's removal half is untested.** The transferring user wasn't a member of his old department group, so there was nothing to remove. The grant path is proven; the revoke path is configured but unexercised.
- **The TAP was never redeemed.** Issuing it depends on emailing the manager, which is blocked by the above.

## Lessons Learned & What I'd Improve Next

**A failed notification must never block provisioning.** My first joiner ran TAP generation before granting access, with the default "stop on error". TAP failed on the missing mailbox, and the two tasks after it were *canceled* — leaving an account that was enabled but had no access, which is the worst of both states. I reordered so access is granted first, and set the email tasks to continue on error. A half-provisioned account that silently stopped mid-workflow is harder to detect than a provisioned account with a visible failed task.

**Lifecycle workflows are versioned, not editable.** `PATCH` accepts only `displayName`, `description`, `isEnabled` and `isSchedulingEnabled`. Changing tasks requires `createNewVersion`. That's a deliberate audit-trail decision: if someone could silently edit what an offboarding workflow does, historical run records would be meaningless. Every past run points at the exact task list that ran.

**Verify identifiers against the tenant, not the docs.** Between a wrong Graph permission name in an earlier lab and a wrong task name here, the same lesson landed twice: published names drift, and the live API is the only source of truth. Resolving names at runtime turned a confusing mid-run failure into a clear deploy-time error.

**Build the guardrail before it's needed.** My snapshot script happily overwrote a before-snapshot, destroying evidence I couldn't recapture. I added a refuse-to-overwrite check. Same category as error handling that fails quietly: code that permits a silent, unrecoverable loss.

**What I'd do next:** move the vendor accounts to B2B guests, which is what they'd be in production (and where the per-guest governance billing applies); add a custom task extension calling a Logic App so offboarding also revokes access in a non-Microsoft SaaS app; alert on `failedTasksCount > 0` so failures surface without someone opening the portal; and enable scheduling to capture unattended runs rather than on-demand ones.

## Skills Demonstrated

| Skill | Where | Maps to |
|---|---|---|
| Joiner-mover-leaver automation | Three attribute-triggered workflows | SC-300: Plan and implement identity lifecycle management |
| Identity governance | Lifecycle Workflows, licensing prerequisites, run reporting | SC-300: Plan and implement identity governance |
| Graph API automation | Live task-catalog resolution, versioning, on-demand activation, run reports | SC-300: Implement an identity management solution |
| Passwordless onboarding | Temporary Access Pass issued to the manager, not the new hire | SC-300: Implement authentication methods |
| Session revocation | Refresh-token revocation in both mover and leaver | MITRE ATT&CK T1078 Valid Accounts |
| Automated deprovisioning | Disable, strip groups, remove licenses on the last day | NIST 800-53 AC-2, AC-2(3) |
| Production scripting practice | `-WhatIf`, idempotency, transcripts, verification, rollback, evidence snapshots | — |

## Repository Structure

```
entra-identity-lifecycle-jml/
├── README.md
├── scripts/
│   ├── check-jml-prereqs.ps1             # read-only licensing + attribute probe
│   ├── seed-jml-test-data.ps1            # managers, lifecycle dates, groups, config
│   ├── deploy-jml-workflows.ps1          # the three workflows (-Update, -RunNow)
│   ├── show-jml-run-results.ps1          # per-task results and failure reasons
│   ├── capture-jml-state.ps1             # before/after snapshots with auto-diff
│   └── remove-jml-workflows.ps1          # rollback
├── evidence/
│   ├── jml-diff-leaver.csv               # 3 group memberships -> 0
│   ├── jml-diff-mover.csv                # department change and access grant
│   └── jml-run-results.csv               # every task, status and failure reason
└── screenshots/
    ├── 01-workflows-deployed.png
    ├── 02-joiner-run-history.png
    ├── 03-leaver-task-results.png
    ├── 04-leaver-diff.png
    └── 05-mover-diff.png
```

> Performed in a Microsoft Entra ID tenant I own. All identities are fictional test accounts. Tenant identifiers, sign-in names and object IDs are redacted from screenshots and script output; the scripts mask the signed-in account and tenant ID by default.
