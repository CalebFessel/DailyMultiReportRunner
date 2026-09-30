# Where vehicle usage numbers come from

Written because the answer is not what the sheet's name suggests, and the
difference changes how the numbers should be read.

## The short version

The Daily Vehicle Overview says which trucks **dispatch assigned calls to**.
It says nothing about which trucks **moved**. Samsara is not consulted.

## What the Daily Vehicle Overview is built from

Traumasoft CAD trip legs, and one function:

```python
def used_vehicle_ids(legs):
    """Vehicles that actually ran a leg on the day."""
    return {str(leg.get("vehicle_id")) for leg in legs if leg.get("vehicle_id")}
```

`traumasoft_reports.py`. Every sheet on the tab derives from that set:

| Sheet | Built by | Means |
| --- | --- | --- |
| Summary | `build_vehicle_summary` | Counts over the same set |
| In Use | `build_vehicles_in_use` | A trip leg was assigned to it |
| Unused In Service | `build_vehicles_unused_in_service` | In service, no leg assigned |
| All In Service | `build_vehicles_all_in_service` | Fleet roster, no usage element |
| Out Of Service | `build_vehicles_out_of_service` | Status plus OOS history |

The fleet list, `vehicle_status` and the odometer column are all Traumasoft's.
The odometer is whatever Traumasoft last recorded -- usually a work order --
not a live reading.

### The gap this leaves

A truck that was driven but never assigned a leg reads as unused. Posting
moves, repositioning, a staffed unit that had a quiet shift, a maintenance
run: all invisible. **The Unused In Service list is therefore an upper bound
on genuinely idle trucks, not a measurement of them.**

## What unit hour utilization is built from

Also Traumasoft, also dispatch timestamps. The numerator is time on task,
enroute to clear (`UHU_SPAN`, default `task`); the denominator is unit hours
the truck was staffed (`UHU_DENOMINATOR`, default `worked`).

A documented problem on this tenant, recorded in `traumasoft_reports.py`:
Clear is being pressed when the next call is assigned rather than when the
last one ended, which tiles the shift and pushes UHU toward 100%. The closer
a unit sits to 100%, the more likely that is the artifact rather than a
saturated truck. `UHU_SPAN=transport` measures enroute to at-destination
instead, cutting off the unreliable tail.

## What Samsara is used for

Route publishing, and nothing else. The entire surface of `samsara_api.py`:

| Call | Endpoint | Direction |
| --- | --- | --- |
| `list_vehicles` | GET /fleet/vehicles | read |
| `list_addresses` | GET /addresses | read |
| `list_routes` | GET /fleet/routes | read |
| `create_route` | POST /fleet/routes | write |
| `delete_route` | DELETE /fleet/routes/{id} | write |

No telematics is read anywhere in this repository. A search for idle, engine
state, fuel, mileage and distance finds nothing outside Traumasoft's own
odometer fields.

`request()` takes an explicit `write=True` from the caller rather than
inferring it from the HTTP verb, so a read-only client cannot be talked into
a write by a helper that happens to POST. Keep new reads `write=False`.

## Closing the gap

Two probes and one report now exist for this. None of the three has been run
against live data yet, so everything below is built and unverified.

**`probe_samsara_movement.py`** asks the four questions that cannot be
answered from the documentation: whether the stats history endpoint returns
engine states in the shape assumed, how far back retention actually reaches,
whether an idle threshold separates the fleet at all, and which series -- OBD
odometer, GPS odometer or GPS track -- carries distance on these trucks. It
also counts, for one day, how far the movement-based answer differs from the
dispatch-based one this document describes.

**`probe_samsara_tags.py`** section 8 measures the VIN join against the name
join it would replace, and names the units each one loses. Run it before
either becomes the join a number depends on -- the Paycor employee matching
in this same repository failed in exactly this way.

**`vehicle_productivity.py`** classifies each vehicle-day against the
Productive / Non-Productive / Indetermined definition upper management
supplied, using Samsara for engine state and distance and Traumasoft for
crew, calls and status. Two of the six inputs it needs are not settled: ePCR
completion has no published route, and the Samsara telematics calls have
never been run here. `docs/VEHICLE_PRODUCTIVITY.md` sets out what that leaves
answerable, what it does not, and the decisions the definition leaves open.

