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

## If the gap is to be closed

Samsara does carry engine state, idle time and distance; this integration
simply never asks for them. Splitting "unused" into genuinely parked versus
moved-but-never-dispatched is a real operational distinction and probably
what a reader of that sheet wants.

Two things to settle before building it, both of which cost real time if
skipped:

**The endpoints must be checked, not recalled.** Samsara's vehicle stats API
is roughly `/fleet/vehicles/stats` with a `types` parameter, plus a history
variant for a window -- but that is a starting point for reading the live
documentation, not a specification. `probe_samsara_readiness.py` is the
read-only probe to extend.

**The join key is unknown.** Samsara vehicles carry `id`, `name` and tags;
Traumasoft vehicles carry their own `id` and `name`. Whether those names
match across both systems has never been checked against live data. A
near-miss join silently drops trucks and still produces a plausible-looking
number. The Paycor employee matching in this same repository failed exactly
this way -- 55 of roughly 85 crew did not match on what looked like a shared
identifier -- so build the diagnostic that measures the overlap before
building the feature that depends on it.
