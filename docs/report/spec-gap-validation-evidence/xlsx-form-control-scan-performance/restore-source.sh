#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
evidence_dir=$script_dir
base_repo=${1:-${LITCHI_FORM_OWNER_BASE_REPO:-}}
destination=${2:-${LITCHI_FORM_OWNER_RESTORE_DIR:-}}
mode=${3:-baseline}

if [[ -z "$base_repo" ]]; then
    printf 'usage: %s BASE_REPOSITORY DESTINATION [baseline|candidate]\n' "$0" >&2
    exit 2
fi
if [[ "$mode" != baseline && "$mode" != candidate ]]; then
    printf 'mode must be baseline or candidate\n' >&2
    exit 2
fi
base_repo=$(CDPATH= cd -- "$base_repo" && pwd)
base_commit=$(cat "$evidence_dir/base-commit.txt")
git -C "$base_repo" rev-parse --verify "$base_commit^{commit}" >/dev/null

if [[ -z "$destination" ]]; then
    destination=$(mktemp -d "${TMPDIR:-/tmp}/litchi-form-owner-restored.XXXXXX")
else
    mkdir -p "$destination"
    destination=$(CDPATH= cd -- "$destination" && pwd)
    if find "$destination" -mindepth 1 -print -quit | grep -q .; then
        printf 'refusing to overwrite non-empty restore directory: %s\n' "$destination" >&2
        exit 2
    fi
fi

python3 - "$base_repo" "$base_commit" "$destination" <<'PY'
import subprocess
import sys
import tarfile

repo, commit, destination = sys.argv[1:]
files = subprocess.check_output(
    ["git", "-C", repo, "ls-tree", "-r", "--name-only", commit], text=True
).splitlines()
selected = []
for path in files:
    if path in {
        "Cargo.toml",
        ".cargo/config.toml",
        "clippy.toml",
        "rust-toolchain.toml",
        "rustfmt.toml",
    }:
        selected.append(path)
    elif path.startswith("crates/") and (
        path.endswith("/Cargo.toml")
        or path.endswith("/build.rs")
        or "/src/" in path
        or "/build/" in path
        or "/proto/" in path
        or "/schema/" in path
        or "/schemas/" in path
        or "/resources/" in path
        or "/include/" in path
    ):
        selected.append(path)

process = subprocess.Popen(
    ["git", "-C", repo, "archive", "--format=tar", commit, "--", *selected],
    stdout=subprocess.PIPE,
)
assert process.stdout is not None
with tarfile.open(fileobj=process.stdout, mode="r|") as archive:
    for member in archive:
        archive.extract(member, destination)
if process.wait() != 0:
    raise SystemExit("git archive failed")
PY

cp "$evidence_dir/source-Cargo.lock" "$destination/Cargo.lock"
if [[ "$mode" == candidate ]]; then
    git -C "$destination" apply --binary "$evidence_dir/source-delta.patch"
fi

python3 - "$destination" "$mode" <<'PY'
import hashlib
import pathlib
import subprocess
import sys

source = pathlib.Path(sys.argv[1])
mode = sys.argv[2]
owner = source / "crates/litchi-xlsx/src/form_control/owner.rs"
lock = source / "Cargo.lock"
expected_owner = {
    "baseline": "5ee687c00133761fddc6ee3cd4be3cc9ebea8e67",
    "candidate": "541c1a3ca4d4b742ac6a134aae8dc942f61404ed",
}[mode]
expected_lock = "aa945c79965460e74a64063e1c45396eaae7a68e71072730e2bc426cded22f02"

def sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()

actual_owner = subprocess.check_output(["git", "hash-object", str(owner)], text=True).strip()
if actual_owner != expected_owner:
    raise SystemExit(f"{mode} owner blob mismatch: {actual_owner} != {expected_owner}")
actual_lock = sha256(lock)
if actual_lock != expected_lock:
    raise SystemExit(f"source Cargo.lock mismatch: {actual_lock} != {expected_lock}")
print(f"{mode} owner git blob: {actual_owner}")
print(f"{mode} owner SHA-256: {sha256(owner)}")
print(f"source Cargo.lock SHA-256: {actual_lock}")
PY

printf '%s\n' "$destination"
