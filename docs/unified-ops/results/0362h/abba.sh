#!/usr/bin/env bash
# Run a protected, calibrated before/after qualification for the File probe.
# The Python driver owns process snapshots and child-tree exemptions; this
# wrapper deliberately contains no timing or build logic.
set -Eeuo pipefail
exec python3 -B "$(dirname "${BASH_SOURCE[0]}")/run.py" "$@"
