# Visual evidence

| File | What it proves |
|---|---|
| `01-workflows-deployed.png` | All three lifecycle workflows deployed at version 2, each task GUID resolved from the live tenant catalog rather than hardcoded. |
| `02-joiner-run-history.png` | Both joiner runs side by side — v1 `failed` with two tasks *canceled* behind a failed notification, v2 `completedWithErrors` after access tasks were moved ahead of email tasks. |
| `03-leaver-audit-trail.png` | Entra audit log for the departing vendor: account disabled, refresh tokens revoked, three group memberships removed — every event initiated by **Lifecycle Workflows**, not by an administrator. |
| `04-leaver-diff.png` | Before/after snapshot diff: `GroupCount 3 → 0`, each removed group named. |
| `05-mover-diff.png` | Before/after snapshot diff: `Department Finance → Security` with `SG-Dept-Security` granted. |
| `06-joiner-audit-trail.png` | Entra audit log for the new hire: workflow execution, group membership, and Temporary Access Pass registration, attributed to the service. |

Portal screenshots are redacted: the signed-in account, tenant name, and user principal names are blacked out.
