# Vehicle productivity: what can be answered, and what cannot

Written against the definition upper management supplied, before building on
it, because two of its six inputs are not reachable today and one of them
decides the entire question.

## The definition as given

```
Productive      Engine On AND Moved >1mile in X time window AND Has assigned
                crew AND Has scheduled pickup AND Has Completed ePCR AND Do
                not have Out of service label.
Non-Productive  No crew assigned OR Known out of service
Indetermined    Engine on, Moving, crewed, Run assigned but ePCR not complete
```

## Where each input comes from

| Input | System | Endpoint | Status |
| --- | --- | --- | --- |
| Engine on | Samsara | `fleet/vehicles/stats/history` `engineStates` | Built, **never run against live data** |
| Moved > 1 mile | Samsara | same, odometer or `gps` | Built, **never run against live data** |
| Assigned crew | Traumasoft | `Schedule/Shifts`, joined on `vehicle_name` | Working |
| Scheduled pickup | Traumasoft | CAD trip legs, joined on `vehicle_id` | Working |
| **Completed ePCR** | Traumasoft | — | **No published route. See below.** |
| Out of service | Traumasoft | `Fleet/Vehicles` `vehicle_status` | Working, with a caveat below |

`vehicle_productivity.py` implements all six. `probe_samsara_movement.py`
sections 1, 3 and 5 establish the two Samsara ones against the live API.

## The ePCR problem

ePCR completion appears in `Productive` (must be complete) and in
`Indetermined` (must be incomplete). It is the only condition that separates
the two positive categories from each other. **Without it, nothing can be
called Productive and nothing can be called Indetermined.**

The published Traumasoft ThirdParty API does not expose it. The spec's own
preamble lists `ThirdParty/Data/Epcr/Huly` and `Trip?rtype=HulyUpdateTrip`
under *"Not included in this spec — private or non-partner integrations"*,
directing partners to dedicated credentials and internal documentation. The
only place an ePCR identifier appears anywhere in the published schema is
`epcr_run_id` / `epcr_run_number` as a **write** target on attachment upload
— you can attach a file to an ePCR whose id you already know, and you cannot
ask whether one was completed.

**This is not yet settled, and settling it is the first thing to do.** An
earlier bare GET of `Epcr/Huly` answered `501`, but one unparameterised call
is weak evidence: these endpoints dispatch on an `rtype` query parameter, and
a surface that answers 501 to a bare call can still answer 200 to a named
action — which is exactly what `Cad/Trip` does. `probe_epcr_huly.py`
enumerates named read actions and records what each one answers. **Run it
before asking anyone for anything.** If every action refuses, its findings
file is the evidence for the ticket.

**If it does refuse, the ask is: credentials and documentation for the
`Epcr/Huly` surface.**
One line changes in `vehicle_productivity.py` when they arrive — the
`epcr_complete` key is already carried through every row as an explicit
`None`.

Until then the report reads as follows, which is the honest shape of the
answer rather than a degraded version of it:

- **`non-productive`** is fully answerable. Both of its terms — no crew, or
  out of service — come from Traumasoft, so it does not depend on Samsara
  either. This is the number that can go to management today.
- **`undetermined`** with the reason *"productive or indetermined — every
  other condition holds"* is the set that would be one or the other. Its size
  is the value of getting ePCR access.
- **`productive`** and **`indetermined`** are always zero. If they are ever
  not, something is reading ePCR from a source this document does not know
  about and it should be checked.

## The gap in the definition, and how management closed it

As first written, the three categories did not cover every vehicle. A truck
that is crewed, in service, engine on, moving, and has **no call assigned**
matched none of them — it is not Non-Productive as defined, because it has
crew and is in service, and it is not Productive or Indetermined, because
both require a scheduled pickup. The same was true of a crewed, in-service
truck that never started.

**Management settled both as Non-Productive.** The rule implemented is
therefore *any Productive condition known to fail*, which covers the original
two terms and both gaps:

```
Non-Productive  out of service, OR no crew assigned, OR no call assigned,
                OR the engine never ran, OR it did not move far enough
```

Note what this also buys: a condition **known** to fail settles the verdict
without waiting on inputs that could not be read. A truck with no call
assigned cannot become Productive on the strength of telematics nobody has,
so it is called Non-Productive even when Samsara is unreachable. Without that,
a dead telematics feed would turn the whole fleet into undetermined rows.

### One verdict, several different problems

A vehicle usually fails more than one condition at once — an unstaffed truck
gets no calls and never starts. Reporting all of them would fragment the
grouping into one bucket per combination and bury the thing to act on, so the
**earliest break in the chain is the headline reason** and every failing
condition is carried in the `failed_conditions` column:

```
available (in service) → staffed (crew) → given work (a call) → actually went
```

You cannot fix "never started" when the real problem is that nobody
dispatched it. Section 3 of the report groups by that headline reason, and
calls out two cases separately because they are different problems:

- **Crewed, moved more than a mile, no call assigned.** Crew hours and fuel
  spent with nothing dispatched — posting, repositioning, a maintenance run.
  This is the blind spot `docs/VEHICLE_USAGE_DATA_SOURCES.md` identifies in
  the current Daily Vehicle Overview, which reads these as simply unused.
- **A call was assigned and the engine never ran.** Dispatch committed a unit
  that did not go.

## Two more decisions the definition leaves open

**What is "X time window"?** The report uses one tenant-local calendar day,
which is the unit the existing fleet sheets use and the one a daily report
needs. `--window-hours` narrows it. If "X" was meant as a rolling window
inside the day — *moved a mile within any 2 hours* — that is a different
measurement and the code would need to change.

**Does a cancelled call count as a scheduled pickup?** It did have one. The
report counts it by default and carries `legs_cancelled` as its own column so
the effect is visible; `--exclude-cancelled` reads it the other way.

## Caveats that will not show up as errors

**Out-of-service status is today's, not the analysed day's.** Traumasoft's
ThirdParty API returns one current `vehicle_status` per vehicle and offers no
status history. For any day but today this is an anachronism: a truck that
went out of service this morning reads as out of service for last Tuesday
too. `--oos-history` reads the observations `traumasoft_reports` accumulates
in `state/vehicle_oos_history.json`, which is the only per-day answer
available and only reaches back to when that file started. Every row records
which of the two it used in `out_of_service_source`.

**The Samsara join has never been measured on this data.** Vehicles join
Traumasoft to Samsara on VIN, falling back to unit name. A vehicle that
matches neither has no engine state and no distance, so it can never be
Productive — it will sit in `undetermined` forever while looking like a data
problem rather than a join problem. The report counts these explicitly, names
them, and separates VIN matches from name matches. `probe_samsara_tags.py`
section 8 measures the join properly. **The Paycor employee matching in this
same repository failed in exactly this way**, so this is measured before it
is trusted.

**A silent gateway is not a parked truck.** A vehicle that reports no engine
states reads as unknown, never as engine-off. One odometer reading in a day
reads as unknown distance, never as zero miles. Zero would be a claim the
data does not support, and it would mark working trucks unproductive.

**Idling counts as engine-on.** An EMS unit idles for climate control and
equipment power. Only Samsara's `off` state counts as not in use; any value
outside the known vocabulary counts as on and is reported by name rather than
silently bucketed.

**GPS track length is a lower bound.** It is a polyline through the samples,
so a gap between fixes cuts the corner off every turn taken inside it. The
odometer delta does not have this problem and is the default. Section 5 of
`probe_samsara_movement.py` reports which series this fleet actually
populates and where the two disagree about a specific truck.

## Running it

```
python probe_epcr_huly.py                             # is ePCR reachable at all?
python probe_samsara_tags.py                          # does the vehicle join hold?
python probe_samsara_movement.py --day 2026-09-22     # engine state and distance
python vehicle_productivity.py --day 2026-09-22 --csv productivity.csv
```

The three probes establish the inputs; run them before trusting the fourth.

Both are strictly read-only, GET only. Neither reads or prints any
patient-identifying value; vehicle names, crew counts and call counts are
operational and the output is safe to paste.

`SAMSARA_TENANT_TIMEZONE` must be set — a calendar-day window cannot be drawn
without it, and getting it wrong shifts every boundary by the offset and
silently reassigns a night shift's work to the wrong date.
