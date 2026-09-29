# Assigned vs not assigned, by day, for one cost center

Asked for **Boardman, July, per day**, with the not-assigned list counting
Boardman's vehicles only.

```
python vehicle_assignment_by_day.py --cost-center Boardman --month 2026-07 \
    --xlsx Boardman_July.xlsx
```

## Run it now

Trips backfill **roughly 90 days** through `GetTrips`. That figure is an
observation recorded in `traumasoft_reports.py` and `validate_against_history.py`,
not a documented guarantee.

On **29 September 2026, 1 July is exactly 90 days back.** The front of the
month is at the edge of the window and another day falls off it every day that
passes. Section 1 of the report states the coverage it actually got, so you
will see immediately how much of July is still reachable — but nothing can get
back a day that has already aged out. Pull it and keep the export.

This is the one piece of good news in the recent run of these questions: it
needs no archive file and no database. Trip legs come straight from the API.

## A day with no legs is not a quiet day

It is a day outside the backfill window, and reading it as "nothing was
assigned" would invent an idle fleet out of an API limit.

The two are told apart by looking at **every** cost center's legs, not just
this one's:

| Observed | Means |
| --- | --- |
| The API returned legs for that day, none for this station | A real zero |
| The API returned no legs at all for that day | Outside the window — `no data` |

`no data` days are never counted as a vehicle sitting idle.

## Three states per vehicle-day, not two

| State | Mark | Means |
| --- | --- | --- |
| `assigned` | `A` | A leg for this cost center |
| `assigned elsewhere` | `E` | A leg, but for somewhere else |
| `not assigned` | `.` | No leg at all |
| `no data` | `?` | The backfill never reached that day |

The middle one matters. A Boardman truck loaned to Cincinnati on the 12th was
not idle, and a two-state report would file it as not assigned and have
somebody asking why Boardman had a truck doing nothing.

A day where a vehicle ran for this station *and* elsewhere counts as
`assigned` — it did this station's work that day, and running somewhere else
too does not take the day away.

## Whose vehicles count as Boardman's

This is the hard part, and it bounds the whole answer. **Cost center is on
neither vehicles nor trips** — only on employees — so the fleet has to be
assembled, and the not-assigned list is only ever as complete as that
assembly.

The trap: a truck that sat idle for the whole of July ran no legs in July, so
defining the fleet from July's own assignments would make **exactly the
vehicle being asked about** invisible.

Sources, best first, each labelled per row in `in_fleet_because`:

1. **Ran a leg for this cost center during the period.** Strongest.
2. **Ran one in a lookback window before it** (`--lookback-days`, default 60).
3. **Its live `shift_name` maps to this cost center.** Today's shift, not July's.
4. **Named in `--roster`** — a JSON list of unit names. The only source that
   cannot miss a truck which has not moved in months.

Section 2 breaks the fleet down by which source placed each vehicle. If no
roster file is given, the report says so and explains what it may therefore be
missing.

## The cost center name

`Boardman` is matched against the names the period's legs actually resolve to,
in three passes, and **the matched names are printed**:

1. Exact, ignoring case.
2. Ignoring the legal-entity wrapper — `Lynx EMS LLC dba Lynx Boardman` and
   `Boardman` are one station. Same normalisation as `probe_samsara_tags`.
3. Substring, **flagged loudly**. `Columbus` matches Columbus Ohio and
   Columbus Indiana alike, and folding them together silently is how a
   station's numbers get somebody else's trucks.

`--list-cost-centers` prints every name the period's legs resolve to.

## Legs nobody can attribute

A leg whose shift profile is not in `state/shift_cost_center_map.json` belongs
to no cost center. It is counted in `legs_unattributed`, **kept separate from
`legs_for_others`** — it may well have been this station's, and filing it under
another station would undercount the days this station had a truck out, which
is the direction that makes a station look idler than it was.

Section 5 names any vehicle whose *only* activity is unattributable, because
those are the verdicts that would flip once the profiles are mapped. Fix them
in `state/shift_cost_center_overrides.json`.

## Output

Three sheets in the `--xlsx` export:

- **By Day** — fleet size, assigned, elsewhere, not assigned, no data, and the
  names of the not-assigned vehicles for that day.
- **By Vehicle** — day counts per state, leg counts split three ways, and how
  the vehicle got into the fleet.
- **Day Grid** — one row per vehicle, one column per day, `A` / `E` / `.` / `?`.
  This is the one to put in front of somebody.

Read-only, GET only. No patient-identifying value is read or printed.
