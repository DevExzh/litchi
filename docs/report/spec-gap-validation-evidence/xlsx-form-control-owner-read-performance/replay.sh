#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
repo_root=${1:-$(git -C "$script_dir" rev-parse --show-toplevel)}
repo_root=$(CDPATH= cd -- "$repo_root" && pwd)
evidence_dir="$script_dir"
commit=$(python3 -c 'import json,sys; print(json.load(open(sys.argv[1]))["commit"])' "$evidence_dir/source-root-gate.json")
restore_root=$(mktemp -d "${TMPDIR:-/var/tmp}/litchi-form-owner-read-source.XXXXXX")
cleanup() {
    git -C "$repo_root" worktree remove --force "$restore_root" >/dev/null 2>&1 || true
}
trap cleanup EXIT

git -C "$repo_root" rev-parse --verify "$commit^{commit}" >/dev/null
git -C "$repo_root" worktree add --detach "$restore_root" "$commit" >/dev/null
cp "$evidence_dir/source-Cargo.lock" "$restore_root/Cargo.lock"

TMPDIR=/var/tmp \
LITCHI_FORM_OWNER_READ_SOURCE_ROOT="$restore_root" \
LITCHI_FORM_OWNER_READ_FIXTURE_ROOT="$restore_root/crates/litchi-xlsx/tests/fixtures/form_control_properties" \
"$evidence_dir/harness/run.sh"
