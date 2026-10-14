#!/usr/bin/env bash
# One-time setup on the sync VM (Ubuntu).  Run as the admin user:
#   bash setup-vm.sh https://github.com/CalebFessel/DailyMultiReportRunner.git claude/intelligent-brown-3xhkap
# Installs Python + git, clones the repo, creates a venv, writes an .env template, and installs an
# hourly cron job.  Re-running is safe.
set -euo pipefail
REPO_URL="${1:?repo url}"
BRANCH="${2:-main}"
APP_DIR="${APP_DIR:-$HOME/wo-sync}"

sudo apt-get update -qq
sudo apt-get install -y -qq python3 python3-venv python3-pip git

if [ ! -d "$APP_DIR/.git" ]; then
  git clone --branch "$BRANCH" "$REPO_URL" "$APP_DIR"
else
  git -C "$APP_DIR" fetch origin "$BRANCH" && git -C "$APP_DIR" checkout "$BRANCH" && git -C "$APP_DIR" pull --ff-only
fi
cd "$APP_DIR"
python3 -m venv .venv
.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet pandas openpyxl python-dotenv requests pymysql   # no pyodbc needed with TS_DB_HOST

if [ ! -f .env ]; then
  cat > .env <<'ENV'
# ---- TraumaSoft database (read-only user) ----
TS_DB_HOST=
TS_DB_PORT=3306
TS_DB_NAME=
TS_DB_USER=
TS_DB_PASSWORD=
TS_DB_SSL=true
# ---- Samsara ----
SAMSARA_API_TOKEN=
# ---- sync ----
TS_WO_SOURCE=db
OUTPUT_DIR=Reports
SAMSARA_WO_LEDGER_SQLITE=Reports/State/samsara_workorder_map.db
TS_TIMEZONE=America/New_York
TEST_MODE=true
TEST_MODE_RECIPIENT=
STATUS_EMAIL_RECIPIENT=
SMTP_SERVER=
SMTP_PORT=587
SMTP_USER=
SMTP_PASS=
SMTP_FROM=
ENV
  chmod 600 .env
  echo ">> wrote $APP_DIR/.env template - fill it in, then run: .venv/bin/python TraumaSoft_Samsara_WorkOrder_Sync.py --discover"
fi

mkdir -p Reports/logs
CRON_LINE="17 * * * * cd $APP_DIR && .venv/bin/python TraumaSoft_Samsara_WorkOrder_Sync.py >> Reports/logs/cron.log 2>&1"
( crontab -l 2>/dev/null | grep -v "TraumaSoft_Samsara_WorkOrder_Sync.py" ; echo "$CRON_LINE" ) | crontab -
echo ">> cron installed: $CRON_LINE"
echo ">> smoke test (no DB needed): .venv/bin/python -m unittest discover -s tests"
