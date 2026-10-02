"""Samsara Function: assign a form and open issues to users matched by
permission profile (role) and tag.

Handler: function.main

Event parameters (all arrive as strings in the event dict):
    RoleName          (required)  Name of the permission profile / role,
                                  e.g. "Field Supervisor".
    TagName           (required)  Name of the tag the role must be scoped to,
                                  e.g. "North Region". Pass "*" to match the
                                  role at the organization level instead.
    FormTemplateId    (required)  UUID of the form template to assign.
    DueInHours        (optional)  Hours from now until the form/issues are
                                  due. Default: "24".
    AssignIssues      (optional)  "true"/"false". Default "true".
    IssueIds          (optional)  Comma-separated issue IDs to assign. When
                                  omitted, all *open, unassigned* issues
                                  updated in the last IssueLookbackDays are
                                  assigned instead.
    IssueLookbackDays (optional)  Default "7".
    DryRun            (optional)  "true" logs what would happen without
                                  writing anything.

Secrets (configured on the Function in the Samsara dashboard):
    SamsaraApiToken   API token with scopes: Read Users, Read Form
                      Submissions, Write Form Submissions, Read Issues,
                      Write Issues.
"""

import json
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

from samsarafnsecrets import get_secrets

API_BASE = "https://api.samsara.com"


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


# ---------------------------------------------------------------------------
# Matching logic
# ---------------------------------------------------------------------------

def find_users_by_role_and_tag(token, role_name, tag_name):
    """Return users holding `role_name`, scoped to tag `tag_name`.

    A user's `roles` is a list of assignments: {role: {id, name},
    tag: {id, name} | absent}. No tag on the assignment means the role
    applies org-wide. tag_name "*" matches the org-wide assignment.
    """
    role_name = role_name.strip().lower()
    tag_name = tag_name.strip().lower()
    matched = []
    for user in _paginate(token, "/users"):
        for assignment in user.get("roles") or []:
            name = ((assignment.get("role") or {}).get("name") or "").lower()
            if name != role_name:
                continue
            tag = assignment.get("tag")
            if tag_name == "*":
                if tag is None:
                    matched.append(user)
                    break
            elif tag and (tag.get("name") or "").lower() == tag_name:
                matched.append(user)
                break
    return matched


def find_unassigned_open_issues(token, lookback_days):
    start = datetime.now(timezone.utc) - timedelta(days=lookback_days)
    issues = _paginate(
        token,
        "/issues/stream",
        {
            "startTime": start.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "status": "open",
        },
    )
    return [i for i in issues if not i.get("assignedTo")]


def get_form_template_revision(token, template_id):
    res = _request(token, "GET", "/form-templates", {"ids": template_id})
    templates = res.get("data") or []
    if not templates:
        raise RuntimeError(f"Form template {template_id} not found")
    tpl = templates[0]
    return tpl["id"], tpl["revisionId"], tpl.get("name", "")


# ---------------------------------------------------------------------------
# Assignment actions
# ---------------------------------------------------------------------------

def assign_form(token, user, template_id, revision_id, due_at, dry_run):
    body = {
        "formTemplate": {"id": template_id, "revisionId": revision_id},
        "status": "notStarted",
        "assignedTo": {"id": str(user["id"]), "type": "user"},
        "dueAtTime": due_at,
    }
    if dry_run:
        print(f"[DryRun] would POST /form-submissions: {json.dumps(body)}")
        return None
    res = _request(token, "POST", "/form-submissions", body=body)
    submission_id = (res.get("data") or {}).get("id")
    print(f"Assigned form submission {submission_id} to {user.get('name')} ({user['id']})")
    return submission_id


def assign_issue(token, issue_id, user, due_at, dry_run):
    body = {
        "assignedTo": {"id": str(user["id"]), "type": "user"},
        "status": "inProgress",
        "dueDate": due_at,
    }
    if dry_run:
        print(f"[DryRun] would PATCH /issues/{issue_id}: {json.dumps(body)}")
        return
    _request(token, "PATCH", f"/issues/{issue_id}", body=body)
    print(f"Assigned issue {issue_id} to {user.get('name')} ({user['id']})")


# ---------------------------------------------------------------------------
# Handler
# ---------------------------------------------------------------------------

def main(event, _context):
    print({
        "trigger": event.get("SamsaraFunctionTriggerSource"),
        "correlationId": event.get("SamsaraFunctionCorrelationId"),
    })

    role_name = event.get("RoleName", "")
    tag_name = event.get("TagName", "")
    template_id = event.get("FormTemplateId", "")
    if not role_name or not tag_name or not template_id:
        raise ValueError("RoleName, TagName and FormTemplateId parameters are required")

    due_in_hours = float(event.get("DueInHours", "24"))
    assign_issues = event.get("AssignIssues", "true").lower() == "true"
    lookback_days = int(event.get("IssueLookbackDays", "7"))
    dry_run = event.get("DryRun", "false").lower() == "true"
    issue_ids_param = [s.strip() for s in event.get("IssueIds", "").split(",") if s.strip()]

    token = get_secrets()["SamsaraApiToken"]

    users = find_users_by_role_and_tag(token, role_name, tag_name)
    if not users:
        raise RuntimeError(
            f"No users found with role '{role_name}' and tag '{tag_name}'"
        )
    print(f"Matched {len(users)} user(s): " + ", ".join(u.get("name", str(u["id"])) for u in users))

    due_at = (datetime.now(timezone.utc) + timedelta(hours=due_in_hours)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )

    template_id, revision_id, template_name = get_form_template_revision(token, template_id)
    print(f"Form template '{template_name}' revision {revision_id}")

    submissions = [
        assign_form(token, user, template_id, revision_id, due_at, dry_run)
        for user in users
    ]

    issues_assigned = 0
    if assign_issues:
        if issue_ids_param:
            issue_ids = issue_ids_param
        else:
            issue_ids = [i["id"] for i in find_unassigned_open_issues(token, lookback_days)]
            print(f"Found {len(issue_ids)} open unassigned issue(s) in last {lookback_days}d")
        # Round-robin issues across the matched users
        for n, issue_id in enumerate(issue_ids):
            assign_issue(token, issue_id, users[n % len(users)], due_at, dry_run)
            issues_assigned += 1

    summary = {
        "dryRun": dry_run,
        "usersMatched": len(users),
        "formSubmissionsCreated": len([s for s in submissions if s]),
        "issuesAssigned": issues_assigned,
    }
    print(json.dumps(summary))
    return summary
