# Route Forms & Issues to Operations Managers by Cost Center

A [Samsara Function](https://developers.samsara.com/docs/functions) that, for
each submitted form:

1. Works out the **cost center**: the tags on the vehicle/asset the form was
   submitted for, falling back to the tags of the worker who submitted it
   (driver tags, or a dashboard user's tag-scoped roles).
2. Finds the user whose permission profile matches `RoleName` (default
   **Operations Manager**) scoped to one of those tags. Vehicle tags take
   priority over submitter tags; ties are broken by name so the pick is
   deterministic.
3. Assigns the submission's open, unassigned **issues** to that manager
   (status `inProgress`, due in `DueInHours`).
4. Optionally assigns the manager a **follow-up form** (`FormTemplateId`),
   with a duplicate guard so overlapping scheduled runs don't double-assign.

## Files

| File | Purpose |
|------|---------|
| `function.py` | Handler (`function.main`) and all logic — stdlib only |
| `samsarafnsecrets.py` | Standard Samsara Functions secrets helper (STS + SSM) |

## Setup

### 1. API token

Create an API token (Settings → API Tokens) with these scopes:

- **Read Users**, **Read Tags** (Setup & Administration)
- **Read Drivers** (Drivers)
- **Read Form Submissions**, **Write Form Submissions**, **Read Issues**,
  **Write Issues** (Forms category)

### 2. Create the Function

Zip the two `.py` files (no folder inside the zip):

```bash
zip route-forms-issues.zip function.py samsarafnsecrets.py
```

In the Samsara dashboard (Settings → Functions → Create Function):

- **Handler:** `function.main`
- **Secret:** `SamsaraApiToken` = the API token from step 1
- **Parameters:**

| Parameter | Required | Example | Notes |
|-----------|----------|---------|-------|
| `RoleName` | no | `Operations Managers` | Permission profile(s), comma-separated. Defaults to `Operations Managers, Operations Manager Geofence and Asset Movement` |
| `FormSubmissionId` | no | `9e11…` | Process exactly this submission (workflow/API runs) |
| `LookbackMinutes` | no | `60` | Poll mode: process forms submitted in the last N minutes |
| `TriggerTemplates` | no | `Truck Check, Wheelchair Van Check` | Poll mode: template **names** (or UUIDs) that trigger routing. Defaults to the six vehicle-check forms: Secure Car Truck Check - Ohio, Secure Car Vehicle Check, Truck Check, Truck Check - Ohio, Wheelchair Van Check, Wheelchair Vehicle Check - Ohio |
| `FormTemplateId` | no | `Corrective Action` | Follow-up form for the manager — UUID **or exact template name**; omit to only assign issues |
| `DueInHours` | no | `24` | Due time for issues and the follow-up form |
| `ExcludeTags` | no | `ALL, BLS` | Tags never treated as cost centers. Default `ALL` |
| `DryRun` | no | `false` | `true` logs actions without writing |

### 3. Trigger it

Two ways to run it:

- **Scheduled (poll mode):** schedule the Function every N minutes with
  `LookbackMinutes = N`. It finds forms submitted in that window and routes
  each one.
- **Event-driven:** invoke it from an alert workflow or
  [`POST /functions/{id}/runs`](https://developers.samsara.com/reference/startfunctionrun)
  with a per-run `FormSubmissionId` parameter override.

## Where to find a form template ID

- **Dashboard:** open the template under **Forms → Templates** (or your
  Connected Workflows forms page) and copy the UUID from the browser URL.
- **API:** `GET https://api.samsara.com/form-templates` lists every template
  with its `id`, `revisionId`, and `name`.
- **Or skip it:** `FormTemplateId` and `TriggerTemplates` both accept the
  template's **exact name** — e.g. `FormTemplateId = Corrective Action` —
  and resolve the UUID at runtime. A typo'd name fails the run with a log
  line listing every template name in the org.

## Behavior notes

- A manager matches when any of their role assignments has
  `role.name == RoleName` **and** that assignment's tag is one of the
  candidate cost-center tags. Org-wide role assignments (no tag) do not
  match — scope the role to the cost-center tag in the dashboard.
- Vehicle/asset tag membership comes from `GET /tags` (each tag lists its
  member vehicles and assets), so untracked/manually-entered assets on a
  form have no tags and fall through to the submitter's tags.
- Issues are matched to the triggering form via their `source` reference
  and only touched when still open and unassigned.
- Test first with `DryRun = true` and check the run logs.

## Local testing

With the [`samsara-fn` CLI](https://pypi.org/project/samsara-fn/):

```bash
pip install samsara-fn
samsara-fn bundle function.py samsarafnsecrets.py
samsara-fn run --param DryRun=true --param LookbackMinutes=1440
```
