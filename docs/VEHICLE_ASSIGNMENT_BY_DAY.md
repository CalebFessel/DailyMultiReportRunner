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

This is the hard part, and it bounds the whole answer.

**It is not dispatch subzone.** Nothing in this report reads `dispatch_subzone`
or `response_subzone`. Attribution runs through the shift profile:

```
trip leg -> shift_name -> the crew who staffed that profile -> employee.cost_center_name
```

accumulated into `state/shift_cost_center_map.json` by the daily run, resolved
by the repository's own `CostCenterMap`.

**But the concern behind the question is right, for a different reason.** That
chain answers *where did this truck work*. It is being read as *where does this
truck belong*, and the two come apart exactly when it matters — a unit covering
a neighbouring station for a week, a truck that ran nothing at all.

And there is no better field to switch to inside Traumasoft. Cost center is not
on a vehicle, and **not on a shift profile either**:
`Lists/Schedule/ShiftProfiles` returns a generic list item — `id` and `name`,
nothing more. Cost center appears on employees, and in the
division/district/group tree from `Data/Organization`, where a district can
span several cost centers. That is the whole surface.

### Samsara tags: ownership instead of inference

Samsara returns tags on its vehicle list. **If** those tags carry the station,
they are a statement of ownership rather than an inference from behaviour, and
they cover a vehicle that never moved.

Whether they do is not assumed. Run the probe first:

```
python probe_samsara_tags.py
```

Section 2 is the tag vocabulary, section 4 the Traumasoft×Samsara join that
bounds everything else, section 5 the tags per unit, section 6 which tags look
like cost center names. If the answer is yes:

```
python build_vehicle_cost_centers.py --compare-days 30
python vehicle_assignment_by_day.py --cost-center Boardman --month 2026-07 \
    --vehicle-cost-centers
```

The generator writes `state/vehicle_cost_centers.json` and, before it does,
**compares the tag answer against the behaviour-derived one**. That comparison
is the point: agreement is reassurance, and each disagreement is either a truck
filed under the wrong station in every report built so far, or a tag nobody
kept up to date. Both are worth settling before the file is used.

Two things it refuses to do:

- A vehicle whose tags name **two** stations is left unset, not assigned to one.
  A roster that looks authoritative and is partly invented is worse than no
  roster.
- Tags are matched by exact name, and by name ignoring the legal-entity wrapper
  (`Parma Heights` finds `Lynx EMS LLC dba Lynx Parma Heights`). **No substring
  pass.** A tag is a short label, and substring-matching short labels is how
  `Columbus` claims Columbus Ohio and Columbus Indiana at once. Tags that say
  the same thing in different words go in `state/tag_cost_center_map.json`.

**The tags are current, not historical.** Using them for July assumes no vehicle
changed station since. That is a far safer assumption for ownership than for
behaviour, but it is still one, and the generated file records its build date so
a reader can judge it.

### What the tags actually say on this fleet (30 September 2026)

`probe_samsara_tags.py` has been run. **The tags carry the station and the
join holds.**

| | |
| --- | --- |
| Fleet-sheet units joined to Samsara | **123 of 125 (98.4%)** |
| Not joined | `A-108`, `Fleet Supplies` |
| Resolved to exactly one station | 110 |
| Tags name two stations | 11 |
| Matched but no station tag | 12 |

The raw figure the probe also prints, 68.3%, is over all 186 Traumasoft
records — mostly iPads and retired trucks. The 98.4% is the one that matters.
The only real gap is `A-108`; `Fleet Supplies` is not a vehicle.

### Boardman is one vehicle, and that is now settled

`M-121` — In Service, tagged `Secure Car`. Nothing else carries the
`Boardman` tag.

The no-station-tag list is the only place Boardman's other trucks could
hide, and it is 12 long. Four of those resolve once the map rules below are
in (`A-301`–`A-304` are Ellicott City, `A-149` is Parma), and `A-500` is
Colorado, which has no cost center. That leaves six Ohio-tagged units with no
station at all — `A-116`, `A-118`, `A-211`, `M-109`, `M-123`, `WC-107`.

**So Boardman is between 1 and 7 vehicles, and 1 is the only number the data
positively supports.** Worth confirming with operations, and worth asking
whether those six Ohio units belong to Boardman before anyone reads a
Boardman report.

### The rules this fleet needs

Three go in `state/tag_cost_center_map.json` (shipped in the `.example`):

| Tag | Cost center |
| --- | --- |
| `Ellicott City` | `Ellicott` |
| `Newburg` | `Newburgh` |
| `Parma` | `Parma Heights` |

A fourth, `"Indiana": "Indianapolis"`, **must not be added.** The probe
offered it because one name contains the other, and it was wrong to: `Indiana`
is the top-level tag over Indianapolis, Salem, Sellersburg *and* Newburg, so
that line would hand Salem's and Sellersburg's trucks to Indianapolis. The
probe now uses Samsara's tag hierarchy to refuse a suggestion for any
top-level tag, and says why. (The resolver was never at risk — it requires an
exact or wrapper-normalised match, and `Indiana` never matched
`Indianapolis`. Only the suggestion was wrong.)

`Admin - Non Providers`, `CPR` and `McDowell` have no tag and are unlikely to
be vehicle-owning cost centers. `Johnstown` is a tag with no cost center.

### Eleven vehicles carry two station tags

Left unset by design, and they fall back to behaviour-derived attribution:

| Vehicles | Tagged |
| --- | --- |
| `A-204`, `A-206`, `M-200`, `M-201`, `WC-201` | Salem / Sellersburg |
| `A-410` | Charles Town / Charleston |
| `A-411` | Charleston / Morgantown |
| `A-415` | Beckley / Clarksburg |
| `WC-406` | Berkeley Springs / Charles Town |
| `WC-409` | Beckley / Parkersburg |
| `M-111` | Massillon / Toledo |

Five of the eleven are Salem/Sellersburg, which looks systematic rather than
accidental — those two may share a pool, or the tags were applied loosely.
None of the eleven touches Boardman, so none of it affects that report.

### Two by-products

- Samsara's `Unit Type` attribute and the `BLS` / `Secure Car` / `Wheelchair`
  tags are a **real** class field, better than the name-prefix convention
  `docs/VEHICLE_SERVICE_DAYS.md` leans on. It also settles the MH question
  with data: `M-121` is tagged `Secure Car`, so on this tenant `M-` is secure
  car, not Medicar.
- The `notes` field, on 4% of vehicles, carries lines like
  `Primary Base: Newburgh`. About five trucks. Too thin to build on, worth
  knowing about for any of the six unplaced units it happens to cover.

### The source ladder

Sources, best first, each labelled per row in `in_fleet_because`:

| | Source | Says |
| --- | --- | --- |
| 0 | `--vehicle-cost-centers` (Samsara tags) | where it **belongs** |
| 1 | Ran a leg for this cost center during the period | where it worked |
| 2 | Ran one in the lookback (`--lookback-days`, default 60) | where it worked |
| 3 | Its live `shift_name` maps here | where it works today |
| 4 | Named in `--roster` | a decision |

The map wins outright where it speaks, and is silent where it has no entry —
those vehicles fall back to the ladder below it. **A vehicle the map places at
another station is excluded even if it ran legs here**, which is the point: a
Cincinnati truck that covered Boardman for a week is Cincinnati's, and letting
its behaviour override the map would put us back where we started.

Section 2 of the report breaks the fleet down by which source placed each
vehicle, and says so when it is leaning on inference.

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
