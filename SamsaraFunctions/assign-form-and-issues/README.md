# Assign Form & Issues by Permission Profile + Tag

A [Samsara Function](https://developers.samsara.com/docs/functions) that:

1. Finds all dashboard **users** whose permission profile (role) matches
   `RoleName` **scoped to the tag** `TagName`.
2. Creates a **form submission** from `FormTemplateId` (latest revision) and
   assigns one to each matched user, due in `DueInHours`.
3. Assigns **issues** to those users (round-robin when more than one user
   matches): either the explicit `IssueIds` you pass, or every *open,
   unassigned* issue updated in the last `IssueLookbackDays` days.

## Files

| File | Purpose |
|------|---------|
| `function.py` | Handler (`function.main`) and all logic — stdlib only |
| `samsarafnsecrets.py` | Standard Samsara Functions secrets helper (STS + SSM) |

## Setup

### 1. API token

Create an API token (Settings → API Tokens) with these scopes:

- **Read Users** (Setup & Administration)
- **Read Form Submissions**, **Write Form Submissions**, **Read Issues**,
  **Write Issues** (Forms category)

### 2. Create the Function

Zip the two `.py` files (no folder inside the zip):

```bash
zip assign-form-and-issues.zip function.py samsarafnsecrets.py
```

In the Samsara dashboard (Settings → Functions → Create Function):

- **Handler:** `function.main`
- **Secret:** `SamsaraApiToken` = the API token from step 1
- **Parameters:**

| Parameter | Required | Example | Notes |
|-----------|----------|---------|-------|
| `RoleName` | yes | `Field Supervisor` | Permission profile name, case-insensitive |
| `TagName` | yes | `North Region` | Tag the role must be scoped to; `*` = org-wide role |
| `FormTemplateId` | yes | `9e118726-41e2-…` | From the template's URL or `GET /form-templates` |
| `DueInHours` | no | `24` | Due time for the form and issues |
| `AssignIssues` | no | `true` | Set `false` to only assign the form |
| `IssueIds` | no | `id1,id2` | Explicit issues; omit to auto-pick open unassigned ones |
| `IssueLookbackDays` | no | `7` | Window for the auto-pick |
| `DryRun` | no | `false` | `true` logs actions without writing |

### 3. Trigger it

Run it manually from the dashboard, on a schedule, from an alert workflow, or
via `POST /functions/{id}/runs` ([Start a Function run](https://developers.samsara.com/reference/startfunctionrun)).
Per-run parameter overrides let you pass specific `IssueIds` from a workflow.

## Behavior notes

- A user matches when **any** of their role assignments has
  `role.name == RoleName` **and** `tag.name == TagName`. A role assignment
  with no tag is org-wide and only matches `TagName = *`.
- Issue assignment sets `assignedTo` (type `user`), status `inProgress`, and
  the same due date as the form. Edit `assign_issue()` if you want issues
  left as `open`.
- Multiple matched users each get their own form submission; issues are
  distributed round-robin.
- Test first with `DryRun = true` and check the run logs.

## Local testing

With the [`samsara-fn` CLI](https://pypi.org/project/samsara-fn/):

```bash
pip install samsara-fn
samsara-fn bundle function.py samsarafnsecrets.py
samsara-fn run --param RoleName="Field Supervisor" --param TagName="North Region" \
  --param FormTemplateId=<uuid> --param DryRun=true
```
