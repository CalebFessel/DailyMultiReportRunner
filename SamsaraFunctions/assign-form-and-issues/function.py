"""Samsara Function: route a submitted form and its issues to the right
Operations Manager by cost-center tag.

For each triggering form submission, the cost center is taken from the
tags of EITHER the worker who submitted the form OR the vehicle/asset the
form was submitted for. The function then finds the user whose permission
profile (role) matches RoleName scoped to one of those tags, assigns the
submission's issues to them, and (optionally) assigns them a follow-up
form.

Handler: function.main

Event parameters (all arrive as strings in the event dict):
    RoleName           (optional) Permission profile name to route to.
                                  Default: "Operations Manager".
    FormSubmissionId   (optional) Process exactly this submission (use when
                                  invoked from a workflow or the API).
    LookbackMinutes    (optional) Poll mode, used when FormSubmissionId is
                                  absent: process every form submitted in
                                  the last N minutes. Default "60". Run the
                                  Function on a matching schedule.
    TriggerTemplates   (optional) Comma-separated form template names or
                                  UUIDs that trigger routing in poll mode.
                                  Defaults to the six vehicle-check forms
                                  in DEFAULT_TRIGGER_TEMPLATES below.
                                  (TriggerTemplateIds works as an alias.)
    FormTemplateId     (optional) Follow-up form to assign to the matched
                                  manager. Accepts a template UUID **or its
                                  exact name**. Omit to only assign issues.
    DueInHours         (optional) Due time for the follow-up form and
                                  issues. Default "24".
    ExcludeTags        (optional) Comma-separated tag names that are never
                                  treated as cost centers. Default "ALL".
    DryRun             (optional) "true" logs what would happen without
                                  writing anything.

Secrets (configured on the Function in the Samsara dashboard):
    SamsaraApiToken    API token with scopes: Read Users, Read Tags, Read
                       Drivers, Read Form Submissions, Write Form
                       Submissions, Read Issues, Write Issues.
"""

import json
import re
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

from samsarafnsecrets import get_secrets

API_BASE = "https://api.samsara.com"
UUID_RE = re.compile(r"[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}")

# Vehicle-check forms that trigger routing when TriggerTemplates is not set.
DEFAULT_TRIGGER_TEMPLATES = (
    "Secure Car Truck Check - Ohio",
    "Secure Car Vehicle Check",
    "Truck Check",
    "Truck Check - Ohio",
    "Wheelchair Van Check",
    "Wheelchair Van Check - Ohio",
)


# ---------------------------------------------------------------------------
# Thin Samsara REST helpers (stdlib only, so the bundle needs no extra deps)
# ---------------------------------------------------------------------------

def _request(token, method, path, params=None, body=None):
    url = API_BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read() or b"{}")
    except urllib.error.HTTPError as e:
        detail = e.read().decode(errors="replace")
        raise RuntimeError(f"{method} {path} -> HTTP {e.code}: {detail}") from e


def _paginate(token, path, params=None, data_key="data"):
    """Yield items across all pages of a cursor-paginated endpoint."""
    params = dict(params or {})
    while True:
        page = _request(token, "GET", path, params)
        yield from page.get(data_key, [])
        pagination = page.get("pagination") or {}
        if not pagination.get("hasNextPage"):
            return
        params["after"] = pagination.get("endCursor")


def _rfc3339(dt_obj):
    return dt_obj.strftime("%Y-%m-%dT%H:%M:%SZ")


def _parse_time(value):
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


# ---------------------------------------------------------------------------
# Org data lookups
# ---------------------------------------------------------------------------

def load_users(token):
    return list(_paginate(token, "/users"))


def build_asset_tag_index(token):
    """Read the org's tags once.

    Returns (asset index, rank map): asset index maps vehicle/asset ID ->
    set of tag names; rank map orders tag names by specificity — deeper in
    the tag hierarchy first (a city under a state), then smaller membership
    first — so a true cost center beats broad parents like a state or "ALL".
    """
    tags = list(_paginate(token, "/tags"))
    parent_of = {str(t["id"]): str(t.get("parentTagId") or "") for t in tags}

    def depth(tag):
        d, cursor = 0, parent_of.get(str(tag["id"]), "")
        while cursor and cursor in parent_of and d < 20:
            d += 1
            cursor = parent_of.get(cursor, "")
        return d

    index, ranks = {}, {}
    for tag in tags:
        name = (tag.get("name") or "").strip()
        if not name:
            continue
        size = sum(
            len(tag.get(k) or [])
            for k in ("vehicles", "assets", "drivers", "machines", "sensors", "addresses")
        )
        ranks[name.lower()] = (-depth(tag), size)
        for member_key in ("vehicles", "assets"):
            for member in tag.get(member_key) or []:
                index.setdefault(str(member.get("id")), set()).add(name)
    return index, ranks


def submitter_tag_names(token, users, submitted_by):
    """Cost-center tags of the worker who submitted the form.

    Drivers carry tags directly; dashboard users get the tags their
    (tag-scoped) roles are attached to.
    """
    if not submitted_by:
        return set()
    sid, stype = str(submitted_by.get("id")), submitted_by.get("type")
    if stype == "driver":
        driver = _request(token, "GET", f"/fleet/drivers/{sid}").get("data") or {}
        return {t["name"].strip() for t in driver.get("tags") or [] if (t.get("name") or "").strip()}
    if stype == "user":
        for user in users:
            if str(user.get("id")) == sid:
                return {
                    ((a.get("tag") or {}).get("name") or "").strip()
                    for a in user.get("roles") or []
                    if ((a.get("tag") or {}).get("name") or "").strip()
                }
    return set()


def managers_for_tags(users, role_name, tag_names_in_priority):
    """Users holding `role_name` scoped to one of the candidate tags.

    Candidate tags are tried in priority order; within a tag, users are
    ordered by name so the pick is deterministic.
    """
    role_name = role_name.strip().lower()
    matched = []
    seen = set()
    for tag_name in tag_names_in_priority:
        wanted = tag_name.strip().lower()
        hits = []
        for user in users:
            for assignment in user.get("roles") or []:
                role_ok = ((assignment.get("role") or {}).get("name") or "").strip().lower() == role_name
                tag_ok = ((assignment.get("tag") or {}).get("name") or "").strip().lower() == wanted
                if role_ok and tag_ok and str(user["id"]) not in seen:
                    hits.append(user)
                    seen.add(str(user["id"]))
                    break
        matched.extend(sorted(hits, key=lambda u: u.get("name", "")))
    return matched


def load_form_templates(token):
    return list(_paginate(token, "/form-templates"))


def _template_title(tpl):
    return (tpl.get("title") or tpl.get("name") or "").strip()


def resolve_form_template(templates, id_or_name):
    """Return (templateId, revisionId, title) from a UUID or an exact title."""
    wanted = id_or_name.strip()
    if UUID_RE.fullmatch(wanted):
        for tpl in templates:
            if tpl["id"] == wanted:
                return tpl["id"], tpl["revisionId"], _template_title(tpl)
        raise RuntimeError(f"Form template {wanted} not found")
    for tpl in templates:
        if _template_title(tpl).lower() == wanted.lower():
            return tpl["id"], tpl["revisionId"], _template_title(tpl)
    known = ", ".join(sorted(_template_title(t) or "?" for t in templates))
    raise RuntimeError(f"No form template named '{wanted}'. Templates in this org: {known}")


# ---------------------------------------------------------------------------
# Trigger submissions and their issues
# ---------------------------------------------------------------------------

def get_submission(token, submission_id):
    res = _request(token, "GET", "/form-submissions", {"ids": submission_id})
    subs = res.get("data") or []
    if not subs:
        raise RuntimeError(f"Form submission {submission_id} not found")
    return subs[0]


def recent_submitted_forms(token, lookback_minutes, trigger_template_ids):
    start = datetime.now(timezone.utc) - timedelta(minutes=lookback_minutes)
    params = {"startTime": _rfc3339(start)}
    if trigger_template_ids:
        params["formTemplateIds"] = ",".join(trigger_template_ids)
    subs = _paginate(token, "/form-submissions/stream", params)
    return [s for s in subs if s.get("submittedAtTime")]


def issues_for_submission(token, submission):
    """Open, unassigned issues created from this form submission."""
    anchor = submission.get("submittedAtTime") or submission.get("createdAtTime")
    start = _parse_time(anchor) - timedelta(hours=1)
    issues = _paginate(
        token,
        "/issues/stream",
        {"startTime": _rfc3339(start), "status": "open"},
    )
    return [
        i
        for i in issues
        if (i.get("source") or {}).get("type") == "form"
        and str((i.get("source") or {}).get("id")) == str(submission["id"])
        and not i.get("assignedTo")
    ]


# ---------------------------------------------------------------------------
# Assignment actions
# ---------------------------------------------------------------------------

def assign_follow_up_form(token, manager, template, due_at, trigger_id, dry_run):
    template_id, revision_id, _ = template
    title = f"Follow-up for submission {trigger_id}"
    body = {
        "formTemplate": {"id": template_id, "revisionId": revision_id},
        "status": "notStarted",
        "assignedTo": {"id": str(manager["id"]), "type": "user"},
        "dueAtTime": due_at,
        "title": title,
    }
    if dry_run:
        print(f"[DryRun] would POST /form-submissions: {json.dumps(body)}")
        return None
    res = _request(token, "POST", "/form-submissions", body=body)
    submission_id = (res.get("data") or {}).get("id")
    print(f"Assigned follow-up form {submission_id} to {manager.get('name')} ({manager['id']})")
    return submission_id


def follow_up_already_exists(token, template_id, trigger_id, lookback_minutes):
    """Guard against double-assigning on overlapping scheduled runs."""
    start = datetime.now(timezone.utc) - timedelta(minutes=lookback_minutes * 2)
    subs = _paginate(
        token,
        "/form-submissions/stream",
        {"startTime": _rfc3339(start), "formTemplateIds": template_id},
    )
    marker = f"Follow-up for submission {trigger_id}"
    return any((s.get("title") or "") == marker for s in subs)


def assign_issue(token, issue_id, manager, due_at, dry_run):
    body = {
        "assignedTo": {"id": str(manager["id"]), "type": "user"},
        "status": "inProgress",
        "dueDate": due_at,
    }
    if dry_run:
        print(f"[DryRun] would PATCH /issues/{issue_id}: {json.dumps(body)}")
        return
    _request(token, "PATCH", f"/issues/{issue_id}", body=body)
    print(f"Assigned issue {issue_id} to {manager.get('name')} ({manager['id']})")


# ---------------------------------------------------------------------------
# Per-submission routing
# ---------------------------------------------------------------------------

def route_submission(token, submission, users, asset_tags, tag_ranks, cfg):
    sub_id = submission["id"]
    asset_id = str((submission.get("asset") or {}).get("id") or "")

    # Most specific tag first: deepest in the hierarchy, then smallest
    # membership, so a city cost center beats a state tag; name breaks
    # remaining ties deterministically. Excluded tags (e.g. "ALL") are
    # never cost-center candidates.
    def by_specificity(names):
        kept = [n for n in names if n.strip().lower() not in cfg["exclude_tags"]]
        return sorted(kept, key=lambda n: (*tag_ranks.get(n.lower(), (0, 0)), n))

    vehicle_tags = by_specificity(asset_tags.get(asset_id, set()))
    worker_tags = by_specificity(submitter_tag_names(token, users, submission.get("submittedBy")))
    print(f"Submission {sub_id}: vehicle tags {vehicle_tags}, submitter tags {worker_tags}")

    # Vehicle cost center wins when both match a manager; submitter is the
    # fallback. Dedupe while keeping priority order.
    candidates = list(dict.fromkeys(vehicle_tags + worker_tags))
    if not candidates:
        print(f"Submission {sub_id}: no tags found on vehicle or submitter; skipping")
        return {"submission": sub_id, "skipped": "no cost-center tags"}

    managers = managers_for_tags(users, cfg["role_name"], candidates)
    if not managers:
        print(
            f"Submission {sub_id}: no '{cfg['role_name']}' scoped to any of {candidates}; skipping"
        )
        return {"submission": sub_id, "skipped": "no matching manager"}
    manager = managers[0]
    if len(managers) > 1:
        others = ", ".join(m.get("name", str(m["id"])) for m in managers[1:])
        print(f"Submission {sub_id}: multiple managers matched; using {manager.get('name')} (others: {others})")

    issues = issues_for_submission(token, submission)
    for issue in issues:
        assign_issue(token, issue["id"], manager, cfg["due_at"], cfg["dry_run"])

    form_created = None
    if cfg["template"]:
        if not cfg["dry_run"] and follow_up_already_exists(
            token, cfg["template"][0], sub_id, cfg["lookback_minutes"]
        ):
            print(f"Submission {sub_id}: follow-up form already exists; not duplicating")
        else:
            form_created = assign_follow_up_form(
                token, manager, cfg["template"], cfg["due_at"], sub_id, cfg["dry_run"]
            )

    return {
        "submission": sub_id,
        "manager": manager.get("name", str(manager["id"])),
        "issuesAssigned": len(issues),
        "followUpForm": form_created,
    }


# ---------------------------------------------------------------------------
# Handler
# ---------------------------------------------------------------------------

def main(event, _context):
    print({
        "trigger": event.get("SamsaraFunctionTriggerSource"),
        "correlationId": event.get("SamsaraFunctionCorrelationId"),
    })

    role_name = event.get("RoleName", "Operations Manager").strip() or "Operations Manager"
    submission_id = event.get("FormSubmissionId", "").strip()
    lookback_minutes = int(event.get("LookbackMinutes", "60"))
    trigger_param = event.get("TriggerTemplates", "") or event.get("TriggerTemplateIds", "")
    trigger_templates = [s.strip() for s in trigger_param.split(",") if s.strip()] or list(
        DEFAULT_TRIGGER_TEMPLATES
    )
    follow_up_template = event.get("FormTemplateId", "").strip()
    due_in_hours = float(event.get("DueInHours", "24"))
    dry_run = event.get("DryRun", "false").lower() == "true"
    exclude_tags = {
        s.strip().lower() for s in event.get("ExcludeTags", "ALL").split(",") if s.strip()
    }

    secrets = get_secrets()
    if "SamsaraApiToken" not in secrets:
        raise RuntimeError(
            "Secret 'SamsaraApiToken' is not configured on this Function. "
            f"Configured secrets: {sorted(secrets)}"
        )
    # Strip whitespace/newlines that sneak in when the token is pasted.
    token = str(secrets["SamsaraApiToken"]).strip()
    templates = load_form_templates(token)

    cfg = {
        "role_name": role_name,
        "due_at": _rfc3339(datetime.now(timezone.utc) + timedelta(hours=due_in_hours)),
        "dry_run": dry_run,
        "lookback_minutes": lookback_minutes,
        "exclude_tags": exclude_tags,
        "template": resolve_form_template(templates, follow_up_template) if follow_up_template else None,
    }
    if cfg["template"]:
        print(f"Follow-up template '{cfg['template'][2]}' revision {cfg['template'][1]}")

    if submission_id:
        submissions = [get_submission(token, submission_id)]
    else:
        resolved = [resolve_form_template(templates, t) for t in trigger_templates]
        trigger_ids = [tpl_id for tpl_id, _, _ in resolved]
        print("Trigger templates: " + ", ".join(f"'{name}'" for _, _, name in resolved))
        submissions = recent_submitted_forms(token, lookback_minutes, trigger_ids)
        print(f"Found {len(submissions)} submitted form(s) in last {lookback_minutes}m")

    users = load_users(token)
    asset_tags, tag_ranks = build_asset_tag_index(token)

    results = [route_submission(token, s, users, asset_tags, tag_ranks, cfg) for s in submissions]

    summary = {"dryRun": dry_run, "processed": len(results), "results": results}
    print(json.dumps(summary))
    return summary
