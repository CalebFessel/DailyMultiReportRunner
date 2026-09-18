# Traumasoft → Paycor timecard integration

Reference for the punch push: which endpoints on both sides, what is verified
against a published spec versus assumed, and the rules the code will not let
you configure away.

Code: `paycor_api.py` (client), `push_paycor_timecards.py` (the push),
`probe_punch_quality.py` (read-only measurement of what would be sent).
Specs: `docs/paycor-public-api-v1.json`, `docs/traumasoft-thirdparty-openapi.yaml`.

---

## Why this exists

Crews clock in and out in two systems. Only Paycor pays them, so nobody has a
reason to close a Traumasoft punch — and every unit-hour figure in the report
bundle rests on that unclosed record. Making the Traumasoft punch the one that
pays fixes the data at its source rather than correcting for it downstream.

That is also the risk. A bad Samsara route wastes a driver's morning; a bad
punch changes what somebody is paid. Every default in this integration is set
accordingly.

---

## Data flow

```
Traumasoft  GET /api/ThirdParty/Data/Schedule/Shifts   → shift rows + punches
            GET /api/ThirdParty/Data/User/Employees    → employee_num, user_id
            GET /api/ThirdParty/Data/Cad/Trip          → tenant UTC offset only

                         ↓  resolve employee, split each punch into two events

Paycor      GET  /v1/legalentities/{id}/employees      → employeeId, departmentId
            GET  /v1/legalentities/{id}/departments
            GET  /v1/legalentities/{id}/activitytypes  → activityTypeId
            GET  /v1/legalEntities/{id}/punches        → reconcile (capital E)
            POST /v1/legalentities/{id}/CreatePunches  → 202 + tracking id
            GET  /v1/legalentities/{id}/punchErrorLog/{trackingId}
```

**The capital `E` in the punch read is not a typo.** Paycor's own spec spells
that one path `legalEntities` and every other path `legalentities`. URL paths
are case-sensitive, so the client reproduces the inconsistency verbatim.

---

## Authentication

### Traumasoft
HMAC-SHA256 over four headers — `X-TS-APIKEY`, `X-TS-TIMESTAMP`, `X-TS-ID`,
`X-TS-AUTHORIZATION`. The client implements both documented formulas and
auto-detects which (formula, secret) pair the tenant accepts, since the key
screen may issue no separate secret. Timestamps are valid for 300 seconds, so a
machine with a wrong clock fails auth rather than failing mysteriously.

### Paycor
Two headers on **every** call, both required:

```
Authorization: Bearer <access token>
Ocp-Apim-Subscription-Key: <APIm subscription key>
```

The bearer comes from the authorization-code grant, which needs a human in a
browser **once**. The refresh token that falls out of it is what belongs in
`.env`; the client exchanges it at `POST /sts/v1/common/token` and caches the
access token in memory for the run.

One oddity worth stating so it does not look like a bug: **on the token call
the subscription key goes in the query string**, and on every other call it
goes in a header.

---

## The three things a guess got wrong

This client was first written from search results, because
`developers.paycor.com` is unreachable from the build environment. All of it
was wrong. It was rebuilt against Paycor's own OpenAPI document, and these are
the corrections worth knowing:

**1. Punches are events, not intervals.** One Traumasoft punch row — a start
and an end — becomes **two** Paycor punches, an In and an Out. A client that
sends one object per shift is sending a shape Paycor has no field for.

**2. Three GUIDs are required per punch, and Paycor defaults none of them.**

| Field | Where it comes from |
|---|---|
| `employeeId` | Paycor employee roster, matched on `employee_num` |
| `departmentId` | the employee's own department, or `PAYCOR_DEFAULT_DEPARTMENT_ID` |
| `activityTypeId` | `PAYCOR_ACTIVITY_TYPE` (default `Work`), resolved by name |

**3. A 202 is not success.** `CreatePunches` validates asynchronously: it
accepts the batch and returns a tracking id. Whether the punches *landed* is
visible only by reading `punchErrorLog/{trackingId}` afterwards. Treating the
202 as "sent" reports success for punches Paycor threw away, so `create_punches`
returns the tracking id and `punch_errors` is the other half of the operation.
**Never call one without the other.**

### The punch object

```json
{
  "employeeId":      "<guid>",
  "departmentId":    "<guid>",
  "punchDateTime":   "2026-09-15T07:00:00",
  "punchStatusType": "In",
  "activityTypeId":  "<guid>",
  "isTransfer":      false,
  "correlationId":   "<uuid5, deterministic>",
  "note":            "optional, 1..300 chars"
}
```

`isTransfer` is always false here: a transfer punch means moving between
departments or activities mid-shift, which nothing in the Traumasoft feed
describes. Timestamps carry no microseconds and no timezone suffix.

---

## Idempotency

**By correlation id, not by timestamp.** Each punch carries a `correlationId`
derived deterministically from the Traumasoft punch id:

```
uuid5(namespace, f"ts-punch:{traumasoft_punch_id}:{half}")
```

Paycor returns it on read, so *"have I sent this already?"* is an exact lookup
rather than a timestamp comparison with a tolerance. The answer survives a
re-run, a restart, and clocks that disagree by a minute.

---

## Employee mapping

Resolved in order; the first that answers wins:

1. **`state/paycor_employee_overrides.json`** — a hand-written map, keyed on
   `employee_num` or `user_id:<id>`. An override always wins, being a decision
   rather than an observation.
2. **`employee_num` against the Paycor roster**, case-insensitive.

**An ambiguous number is refused, never guessed.** Two people sharing a payroll
number is exactly where a guess pays the wrong person. An employee with no
`employee_num` at all is refused for the same reason.

---

## Safety rails

None of these are configurable away.

- **An open punch is never sent.** A punch with no clock-out is not a payable
  record. The shift-end fallback that makes it usable for *reporting* would
  here mean inventing the end of somebody's paid day.
- **An unmappable employee is refused**, with the reason printed.
- **A punch Paycor already holds, by correlation id, is skipped.**
- **The first failed batch stops the run.** A partial push is easier to reason
  about than a push that kept going past an error.
- **Sandbox is the default.** `PAYCOR_ENVIRONMENT` must read exactly
  `production`; anything else — including a typo — resolves to the sandbox.
- **Production needs two flags**: `PAYCOR_ENVIRONMENT=production` *and*
  `--i-understand-this-is-payroll` on the command line.
- **Batches cap at 100 punches** so one error log stays readable.

---

## The three modes, in order

```powershell
python push_paycor_timecards.py                     # 1. dry run (default)
python push_paycor_timecards.py --reconcile         # 2. compare, still read-only
python push_paycor_timecards.py --reconcile --publish --limit 2
```

**1. Dry run.** Builds the plan and prints it, including the exact JSON that
would be sent and everything refused with reasons. Reads only.

**2. Reconcile.** Reads what Paycor already holds for the same window and
compares. **This is the test that matters before any write** — it proves the
employee mapping resolves, the clocks agree, and both systems describe the same
shifts. Still read-only on both sides. `PAYCOR_MATCH_TOLERANCE_MINUTES`
(default 2) sets how far apart two records can be and still be the same punch.

**3. Publish.** Sends, then reads the error log back. Start with `--limit 2`.

`probe_punch_quality.py` is the step before all three: it measures Traumasoft
punch quality — how many are open, how many close at a plausible time — without
touching Paycor at all. It needs only Traumasoft credentials.

---

## Configuration

| Variable | Required | Notes |
|---|---|---|
| `PAYCOR_ENVIRONMENT` | — | `sandbox` (default) or `production` |
| `PAYCOR_SUBSCRIPTION_KEY` | **yes** | APIm key, sent on every call |
| `PAYCOR_REFRESH_TOKEN` | **yes**\* | from the one-time browser grant |
| `PAYCOR_ACCESS_TOKEN` | \* | substitutes for the refresh token; expires, so it is for a one-off test, not a scheduled run |
| `PAYCOR_CLIENT_ID` / `_SECRET` | if your grant needs them | sent on the token exchange only when set |
| `PAYCOR_LEGAL_ENTITY_ID` | **yes** | integer, in every path |
| `PAYCOR_ACTIVITY_TYPE` | — | default `Work`, resolved by name |
| `PAYCOR_ACTIVITY_TYPE_ID` | — | set when the name is ambiguous |
| `PAYCOR_DEFAULT_DEPARTMENT_ID` | — | fallback when an employee has none |
| `PAYCOR_MATCH_TOLERANCE_MINUTES` | — | default 2, reconcile only |
| `PAYCOR_MAX_PUNCH_HOURS` | — | default 24; longer is refused as bad data |
| `PAYCOR_MAX_PUNCH_BATCH` | — | default 100 |

\* One of `PAYCOR_REFRESH_TOKEN` or `PAYCOR_ACCESS_TOKEN` must be set; the
client refuses to start without either. Prefer the refresh token — it is what
survives a scheduled run.

---

## The window limitation

`/Schedule/Shifts` returns **`today-1..today+2`** and ignores every date filter
— verified byte-identical for requests at −30, +0 and +30 days. Punches live
only on those rows.

So this push covers a narrow window and **must run daily to cover a pay
period**. A day missed is a day whose punches the API will not hand back, and
no later run can recover it. That is the same constraint that caps the UHU
denominator and the staffing review; it is a property of the API, not of this
code.

---

## Known gaps

- **Not yet run against live Paycor.** Everything above is built from Paycor's
  published spec and tested against stubbed transports; no call has been made
  to a real tenant. The reconcile step exists to be the first thing that meets
  one.
- **Traumasoft punch quality is the real unknown.** The integration is sound
  only if the punches are. Crews have had no reason to close them, which is the
  problem this fixes — but it means the first real dry run may show a lot of
  open punches, and those are refused rather than repaired.
- **No transfer punches.** Moving between departments or activities mid-shift
  is not modelled, because the Traumasoft feed does not describe it.
- **Reading Paycor timecards for reporting is not built.** Paycor holds real
  historical timecards and could answer the hours questions Traumasoft cannot
  backfill. That is a separate read path, not this write path.
