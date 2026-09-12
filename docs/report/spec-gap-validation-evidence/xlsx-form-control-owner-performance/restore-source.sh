#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
evidence_dir=$script_dir
base_repo=${1:-${LITCHI_FORM_OWNER_BASE_REPO:-}}
destination=${2:-${LITCHI_FORM_OWNER_RESTORE_DIR:-}}
if [[ -z "$base_repo" ]]; then
    printf 'usage: %s BASE_REPOSITORY [DESTINATION]\n' "$0" >&2
    exit 2
fi
base_repo=$(CDPATH= cd -- "$base_repo" && pwd)
base_commit=$(python3 -c 'import json,sys; print(json.load(open(sys.argv[1]))["base"])' \
    "$evidence_dir/source-root-gate.json")
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
import pathlib
import subprocess
import sys
import tarfile

repo, commit, destination = sys.argv[1:]
files = subprocess.check_output(
    ["git", "-C", repo, "ls-tree", "-r", "--name-only", commit], text=True
).splitlines()
selected = []
for path in files:
    if path in {"Cargo.toml", ".cargo/config.toml", "clippy.toml", "rust-toolchain.toml"}:
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
        or path.startswith("crates/litchi-xlsx/tests/fixtures/form_control_properties/")
        and path.endswith(".xlsx")
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
git -C "$destination" apply --binary "$evidence_dir/source-delta.patch"
cp "$evidence_dir/source-Cargo.lock" "$destination/Cargo.lock"

python3 - "$destination" "$evidence_dir" <<'PY'
import hashlib
import json
import pathlib
import sys
import tarfile

destination = pathlib.Path(sys.argv[1])
evidence = pathlib.Path(sys.argv[2])
gate = json.loads((evidence / "source-root-gate.json").read_text())

def digest(path):
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()

for relative, expected in gate["sha256"].items():
    path = destination / relative
    actual = digest(path)
    if actual != expected:
        raise SystemExit(f"restored source hash mismatch: {relative}: {actual} != {expected}")

expected_lock = (evidence / "source-Cargo.lock.sha256").read_text().split()[0]
actual_lock = digest(destination / "Cargo.lock")
if actual_lock != expected_lock:
    raise SystemExit(f"restored Cargo.lock hash mismatch: {actual_lock} != {expected_lock}")

manifest_path = evidence / "source-manifest.sha256"
if manifest_path.is_file():
    for line in manifest_path.read_text().splitlines():
        expected, relative = line.split("  ", 1)
        path = destination / relative
        actual = digest(path)
        if actual != expected:
            raise SystemExit(f"restored source manifest mismatch: {relative}: {actual} != {expected}")

with tarfile.open(evidence / "source-bundle.tar", "r") as bundle:
    for member in bundle.getmembers():
        if not member.isfile():
            raise SystemExit(f"unexpected non-file source bundle member: {member.name}")
        bundled = bundle.extractfile(member)
        if bundled is None:
            raise SystemExit(f"cannot read source bundle member: {member.name}")
        restored = (destination / member.name).read_bytes()
        if restored != bundled.read():
            raise SystemExit(f"source bundle differs from restored tree: {member.name}")
PY

printf '%s\n' "$destination"
