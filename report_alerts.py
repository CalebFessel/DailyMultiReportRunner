"""
Alerting for the unattended runs -- loud when it matters, silent when it does not.

Two different failures need two different mechanisms, and only one of them can
be solved from inside the job:

  A run that BREAKS can shout. It is running, it knows something went wrong,
  and it can send mail, post to a channel and text a phone.

  A run that NEVER HAPPENS cannot. A machine that is switched off, a scheduled
  task someone disabled, a laptop that went back to IT -- none of them raise
  an error, and silence looks exactly like success. Only something OUTSIDE the
  job notices that, which is what the heartbeat is for: every good run pings a
  URL, and the service behind it alerts when a ping does not arrive on time.
  Without that, the worst failure is the quietest one.

Channels are each optional and configured by environment variable, so a host
that has only mail still works:

    ALERT_EMAIL_TO       comma-separated; uses the same SMTP as the reports
    ALERT_SMS_TO         comma-separated carrier gateway addresses, e.g.
                         5551234567@vtext.com -- CRITICAL only, short body
    ALERT_WEBHOOK_URL    a Teams or Slack incoming webhook
    ALERT_HEARTBEAT_URL  pinged on success; the dead man's switch
    ALERT_SOURCE         a name for this host, so two machines are told apart

Severity decides who is woken:

    CRITICAL   the job failed, or payroll data may be wrong. Everything fires,
               including SMS. This is the tier that is allowed to wake someone.
    WARNING    the job worked but something needs a person today. Mail and
               webhook, no SMS.
    INFO       normal completion. Mail only, and only if asked for.

Nothing here raises. An alerting failure must never be the reason a payroll
job dies, and a channel that fails is reported on the others rather than
swallowed -- the one thing worse than a missed alert is a missed alert nobody
knows was missed.
"""

import os
import json
import logging
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

log = logging.getLogger(__name__)

CRITICAL = "CRITICAL"
WARNING = "WARNING"
INFO = "INFO"

SEVERITIES = (CRITICAL, WARNING, INFO)

# A subject prefix an inbox rule can match on without parsing the body.
SUBJECT_PREFIX = {
    CRITICAL: "[PAYROLL FAILURE]",
    WARNING: "[PAYROLL WARNING]",
    INFO: "[payroll]",
}

# Teams renders this colour down the side of the card.
WEBHOOK_COLOUR = {
    CRITICAL: "D93F3F",
    WARNING: "E8A33D",
    INFO: "5B8DEF",
}

TIMEOUT = int(os.getenv("ALERT_TIMEOUT", "20"))


def _recipients(name):
    raw = os.getenv(name, "")
    return [part.strip() for part in raw.split(",") if part.strip()]


def source_name():
    """Which host this came from, so two machines are not confused."""
    return os.getenv("ALERT_SOURCE") or os.getenv("COMPUTERNAME") or "unknown host"


def _send_email(severity, headline, detail, recipients):
    # Imported here so a host with no reporting dependencies can still alert
    # through the other channels.
    from report_output import send_email

    subject = f"{SUBJECT_PREFIX[severity]} {headline}"
    body = (
        f"{headline}\n\n"
        f"{detail}\n\n"
        f"---\n"
        f"host: {source_name()}\n"
        f"time: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
    )
    send_email(subject, body, recipients)


def _send_sms(headline, recipients):
    from report_output import send_email

    # Carrier gateways truncate hard and drop subjects inconsistently, so the
    # whole message goes in the body and stays inside one segment.
    body = f"PAYROLL FAILED: {headline}"[:300]
    send_email("", body, recipients)


def _send_webhook(severity, headline, detail, url):
    payload = {
        "@type": "MessageCard",
        "@context": "https://schema.org/extensions",
        "themeColor": WEBHOOK_COLOUR[severity],
        "summary": f"{SUBJECT_PREFIX[severity]} {headline}",
        "title": f"{SUBJECT_PREFIX[severity]} {headline}",
        "text": (detail or "").replace("\n", "\n\n"),
        "sections": [{
            "facts": [
                {"name": "Host", "value": source_name()},
                {"name": "Time", "value": f"{datetime.now():%Y-%m-%d %H:%M:%S}"},
            ]
        }],
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
        response.read()


def alert(severity, headline, detail=""):
    """
    Raise an alert on every configured channel. Returns what actually went.

    Each channel is attempted independently: one that fails must not stop the
    others, because the failing one may be the reason an alert was needed.
    """
    if severity not in SEVERITIES:
        raise ValueError(f"severity must be one of {SEVERITIES}, got {severity!r}")

    sent, failed = [], []

    email_to = _recipients("ALERT_EMAIL_TO")
    if email_to:
        try:
            _send_email(severity, headline, detail, email_to)
            sent.append("email")
        except Exception as exc:
            failed.append(f"email ({exc})")

    webhook = os.getenv("ALERT_WEBHOOK_URL", "").strip()
    if webhook and severity in (CRITICAL, WARNING):
        try:
            _send_webhook(severity, headline, detail, webhook)
            sent.append("webhook")
        except Exception as exc:
            failed.append(f"webhook ({exc})")

    sms_to = _recipients("ALERT_SMS_TO")
    if sms_to and severity == CRITICAL:
        try:
            _send_sms(headline, sms_to)
            sent.append("sms")
        except Exception as exc:
            failed.append(f"sms ({exc})")

    if sent:
        log.info("Alert (%s) sent via %s: %s", severity, ", ".join(sent), headline)
    if failed:
        # Loud, because a channel that failed here is a channel that will fail
        # next time, and the log is all that is left.
        log.error("Alert (%s) could NOT be delivered via %s: %s",
                  severity, "; ".join(failed), headline)
    if not sent and not failed:
        log.warning("No alert channel is configured, so this was not sent "
                    "anywhere: [%s] %s", severity, headline)

    return {"sent": sent, "failed": failed}


def heartbeat(ok=True, detail=""):
    """
    Tell the outside world this run finished.

    The point of failure this covers is the one the job cannot report itself:
    not running at all. Ping only on genuine success -- a heartbeat sent from a
    run that did nothing is worse than none, because it silences the only
    alarm that would have noticed.
    """
    url = os.getenv("ALERT_HEARTBEAT_URL", "").strip()
    if not url:
        return False
    if not ok:
        url = url.rstrip("/") + "/fail"

    try:
        request = urllib.request.Request(
            url, data=(detail or "").encode("utf-8")[:10000], method="POST"
        )
        with urllib.request.urlopen(request, timeout=TIMEOUT) as response:
            response.read()
        log.info("Heartbeat sent (%s).", "ok" if ok else "fail")
        return True
    except Exception as exc:
        # Worth an alert of its own: a heartbeat that stops arriving is
        # indistinguishable from a host that died, so the watcher will page
        # someone over a run that actually succeeded.
        log.error("Heartbeat could not be sent to the watcher: %s", exc)
        return False


class guarded:
    """
    Wrap an unattended run so nothing fails quietly.

    An unhandled exception alerts CRITICAL and re-raises; a clean exit pings
    the heartbeat. Use it as the outermost thing in a scheduled entry point:

        with alerts.guarded("Paycor timeclock push"):
            return main()
    """

    def __init__(self, name, heartbeat_on_success=True):
        self.name = name
        self.heartbeat_on_success = heartbeat_on_success

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback_obj):
        if exc_type is None:
            if self.heartbeat_on_success:
                heartbeat(ok=True, detail=f"{self.name} completed")
            return False

        import traceback as tb
        detail = "".join(tb.format_exception(exc_type, exc, traceback_obj))
        alert(CRITICAL, f"{self.name} failed: {exc_type.__name__}: {exc}", detail)
        heartbeat(ok=False, detail=detail[:2000])
        return False  # never swallow: the exit code still has to be wrong
