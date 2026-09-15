#!/usr/bin/env bash
set -Eeuo pipefail

readonly APP_DIR="/opt/ural-vape-regions-bot/app"
readonly APP_USER="uralbo"
readonly APP_HOME="/opt/ural-vape-regions-bot"
readonly SERVICE="ural-vape-regions-bot.service"
readonly REVISION="${1:-}"

if [[ ! "$REVISION" =~ ^[0-9a-f]{40}$ ]]; then
    echo "Expected a full 40-character Git revision." >&2
    exit 2
fi

exec 9>"/run/lock/ural-vape-regions-bot-deploy.lock"
if ! flock -n 9; then
    echo "Another deployment is already running." >&2
    exit 3
fi

run_as_app() {
    runuser -u "$APP_USER" -- env HOME="$APP_HOME" "$@"
}

if [[ ! -d "$APP_DIR/.git" ]]; then
    echo "Application repository not found: $APP_DIR" >&2
    exit 4
fi

# The SSH deploy account has a private home directory. Move into the application
# directory before invoking commands as APP_USER so pytest can resolve its start path.
cd "$APP_DIR"

if [[ -n "$(run_as_app git -C "$APP_DIR" status --porcelain --untracked-files=no)" ]]; then
    echo "Tracked server files contain manual changes; deployment stopped." >&2
    exit 5
fi

run_as_app git -C "$APP_DIR" fetch --prune origin main
remote_revision="$(run_as_app git -C "$APP_DIR" rev-parse origin/main)"
if [[ "$remote_revision" != "$REVISION" ]]; then
    echo "origin/main is $remote_revision, but workflow requested $REVISION." >&2
    exit 6
fi

run_as_app git -C "$APP_DIR" checkout main
run_as_app git -C "$APP_DIR" merge --ff-only "$REVISION"
run_as_app "$APP_DIR/.venv/bin/python" -m pip install -r "$APP_DIR/requirements.txt"
run_as_app "$APP_DIR/.venv/bin/python" -m py_compile \
    "$APP_DIR/bot.py" \
    "$APP_DIR/moysklad_bonus.py" \
    "$APP_DIR/bonus_report_v2.py" \
    "$APP_DIR/inventory_inline.py"
run_as_app "$APP_DIR/.venv/bin/python" -m pytest -q "$APP_DIR/tests"

systemctl restart "$SERVICE"
sleep 3
systemctl is-active --quiet "$SERVICE"
systemctl --no-pager --full status "$SERVICE"
echo "Successfully deployed $REVISION"
