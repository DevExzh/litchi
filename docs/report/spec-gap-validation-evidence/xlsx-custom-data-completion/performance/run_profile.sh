#!/usr/bin/env bash
set -euo pipefail

script_dir=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
harness_dir="$script_dir/harness"
source_link="$script_dir/source"
smoke=0
if [[ "${1:-}" == "--smoke" ]]; then
    smoke=1
    shift
fi

if (( smoke )); then
    source_root=${1:-}
    [[ -n "$source_root" ]] || {
        printf '%s\n' 'smoke usage: run_profile.sh --smoke /absolute/source/root' >&2
        exit 2
    }
    samples=1
    warmups=1
    capture_kind=smoke
    levels=(small)
    output_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-xlsx-custom-smoke.XXXXXX")
else
    source_root=${LITCHI_XLSX_CUSTOM_DATA_SOURCE_ROOT:-}
    freeze_receipt=${LITCHI_XLSX_CUSTOM_DATA_FREEZE_RECEIPT:-}
    capture_token=${LITCHI_XLSX_CUSTOM_DATA_CAPTURE_TOKEN:-}
    capture_kind=${LITCHI_XLSX_CUSTOM_DATA_CAPTURE_KIND:-candidate}
    samples=${LITCHI_XLSX_CUSTOM_DATA_SAMPLES:-15}
    warmups=${LITCHI_XLSX_CUSTOM_DATA_WARMUPS:-3}
    output_dir=${LITCHI_XLSX_CUSTOM_DATA_RESULT_DIR:-$script_dir/results/final}
    [[ "$capture_kind" == "baseline" || "$capture_kind" == "candidate" ]] || {
        printf '%s\n' 'capture kind must be baseline or candidate' >&2
        exit 2
    }
    expected_token="final-source-frozen"
    if [[ "$capture_kind" == "baseline" ]]; then
        expected_token="baseline-source-frozen"
    fi
    [[ "$capture_token" == "$expected_token" ]] || {
        printf 'refusing %s capture: set LITCHI_XLSX_CUSTOM_DATA_CAPTURE_TOKEN=%s\n' "$capture_kind" "$expected_token" >&2
        exit 2
    }
    [[ -n "$freeze_receipt" && -f "$freeze_receipt" ]] || {
        printf '%s\n' 'refusing final capture: missing LITCHI_XLSX_CUSTOM_DATA_FREEZE_RECEIPT' >&2
        exit 2
    }
    levels=(small medium large)
fi

source_root=$(CDPATH= cd -- "$source_root" && pwd)
[[ "$samples" =~ ^[1-9][0-9]*$ && "$warmups" =~ ^[0-9]+$ ]] || {
    printf '%s\n' 'samples must be positive and warmups must be non-negative integers' >&2
    exit 2
}
[[ -f "$source_root/Cargo.toml" && -f "$source_root/Cargo.lock" ]] || {
    printf 'source root must contain Cargo.toml and Cargo.lock: %s\n' "$source_root" >&2
    exit 2
}
[[ ! -e "$source_link" ]] || {
    printf 'refusing to replace existing source link: %s\n' "$source_link" >&2
    exit 2
}

if (( ! smoke )); then
    python3 - "$freeze_receipt" "$source_root" "$script_dir" <<'PY'
import hashlib
import json
import pathlib
import sys

receipt = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding="utf-8"))
root = pathlib.Path(sys.argv[2]).resolve()
profile_root = pathlib.Path(sys.argv[3]).resolve()
if pathlib.Path(receipt.get("source_root", root)).resolve() != root:
    raise SystemExit("freeze receipt source_root does not match the selected source")
selected = receipt.get("selected_files")
if not isinstance(selected, dict) or not selected:
    raise SystemExit("freeze receipt selected_files is empty")
for relative, expected in selected.items():
    candidate = root / relative
    if not candidate.is_file():
        raise SystemExit(f"freeze receipt file is missing: {relative}")
    actual = hashlib.sha256(candidate.read_bytes()).hexdigest()
    if actual != expected:
        raise SystemExit(f"freeze receipt hash mismatch for {relative}")
if "Cargo.lock" not in selected:
    raise SystemExit("freeze receipt must hash Cargo.lock")
profile_files = receipt.get("profile_files", {})
if profile_files:
    if pathlib.Path(receipt.get("profile_root", profile_root)).resolve() != profile_root:
        raise SystemExit("freeze receipt profile_root does not match the selected profile")
    for relative, expected in profile_files.items():
        candidate = profile_root / relative
        if not candidate.is_file():
            raise SystemExit(f"freeze receipt profile file is missing: {relative}")
        actual = hashlib.sha256(candidate.read_bytes()).hexdigest()
        if actual != expected:
            raise SystemExit(f"freeze receipt profile hash mismatch for {relative}")
PY
fi

mkdir -p "$output_dir"
output_dir=$(CDPATH= cd -- "$output_dir" && pwd)
if find "$output_dir" -mindepth 1 -print -quit | grep -q .; then
    printf 'refusing to overwrite non-empty result directory: %s\n' "$output_dir" >&2
    exit 2
fi

target_dir=$(mktemp -d "${TMPDIR:-/tmp}/litchi-xlsx-custom-target.XXXXXX")
cleanup() {
    python3 - "$source_link" "$target_dir" "$output_dir" "$smoke" <<'PY'
import pathlib
import shutil
import sys

source_link = pathlib.Path(sys.argv[1])
target = pathlib.Path(sys.argv[2])
output = pathlib.Path(sys.argv[3])
smoke = sys.argv[4] == "1"
if source_link.is_symlink():
    source_link.unlink()
elif source_link.exists():
    raise SystemExit(f"refusing cleanup of non-symlink source path: {source_link}")
if target.parent.resolve() != pathlib.Path("/tmp").resolve() or not target.name.startswith("litchi-xlsx-custom-target."):
    raise SystemExit(f"refusing cleanup of unexpected target: {target}")
if target.is_dir():
    shutil.rmtree(target)
if smoke and output.parent.resolve() == pathlib.Path("/tmp").resolve() and output.name.startswith("litchi-xlsx-custom-smoke.") and output.is_dir():
    shutil.rmtree(output)
PY
}
trap cleanup EXIT

ln -s "$source_root" "$source_link"
export CARGO_INCREMENTAL=0
export LITCHI_PROFILE_COMMIT=$(git -C "$source_root" rev-parse HEAD)

python3 - "$source_root" "$output_dir/source-manifest-before.sha256" <<'PY'
import hashlib
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
output = pathlib.Path(sys.argv[2])
roots = ("crates/litchi-core", "crates/litchi-opc", "crates/litchi-ooxml-common", "crates/litchi-xlsx", "crates/soapberry-zip")
files = [root / "Cargo.toml", root / "Cargo.lock"]
for relative_root in roots:
    base = root / relative_root
    files.extend(path for path in base.rglob("*") if path.is_file())
with output.open("w", encoding="utf-8") as stream:
    for path in sorted(set(files), key=lambda value: value.relative_to(root).as_posix()):
        stream.write(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}\n")
PY

python3 - "$script_dir" "$output_dir/profile-manifest.sha256" <<'PY'
import hashlib
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
output = pathlib.Path(sys.argv[2])
relative_files = (
    "harness/Cargo.toml",
    "harness/Cargo.lock",
    "harness/main.rs",
    "harness/support.rs",
    "run_profile.sh",
    "summarize.py",
    "PLAN.md",
    "README.md",
)
with output.open("w", encoding="utf-8") as stream:
    for relative in relative_files:
        path = root / relative
        if not path.is_file():
            raise SystemExit(f"profile file is missing: {relative}")
        stream.write(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {relative}\n")
PY

CARGO_TARGET_DIR="$target_dir" cargo build --locked --release --manifest-path "$harness_dir/Cargo.toml"
binary="$target_dir/release/xlsx-custom-data-performance"
[[ -x "$binary" ]] || { printf 'missing profile binary: %s\n' "$binary" >&2; exit 1; }

sha256sum "$binary" > "$output_dir/binary.sha256"
rustc -Vv > "$output_dir/rustc-vv.txt"
cargo -V > "$output_dir/cargo-version.txt"
uname -a > "$output_dir/host.txt"
python3 - "$source_root" "$harness_dir" "$output_dir" "$capture_kind" "$samples" "$warmups" "${levels[*]}" <<'PY'
import hashlib
import json
import os
import pathlib
import subprocess
import sys

source_root = pathlib.Path(sys.argv[1])
harness_dir = pathlib.Path(sys.argv[2])
output = pathlib.Path(sys.argv[3])
capture_kind = sys.argv[4]
samples = int(sys.argv[5])
warmups = int(sys.argv[6])
levels = sys.argv[7].split()

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

payload = {
    "schema": "xlsx-custom-data-performance-provenance-v1",
    "capture_kind": capture_kind,
    "source_root": str(source_root),
    "source_commit": subprocess.check_output(
        ["git", "-C", str(source_root), "rev-parse", "HEAD"], text=True
    ).strip(),
    "samples": samples,
    "warmups": warmups,
    "levels": levels,
    "lanes": ["read", "noop-commit-save", "payload-replacement", "rename-binding-rewrite", "remove-inverse"],
    "cargo_incremental": "0",
    "rustflags": os.environ.get("RUSTFLAGS", ""),
    "rustdocflags": os.environ.get("RUSTDOCFLAGS", ""),
    "source_cargo_lock_sha256": digest(source_root / "Cargo.lock"),
    "harness_cargo_lock_sha256": digest(harness_dir / "Cargo.lock"),
    "profile_manifest": "profile-manifest.sha256",
    "source_manifest_before": "source-manifest-before.sha256",
}
(output / "provenance.json").write_text(
    json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8"
)
PY

for level in "${levels[@]}"; do
    for lane in read noop-commit-save payload-replacement rename-binding-rewrite remove-inverse; do
        raw="$output_dir/${level}-${lane}.jsonl"
        : > "$raw"
        for sample in $(seq 1 "$samples"); do
            json_file="$output_dir/${level}-${lane}-sample${sample}.json"
            stderr_file="$output_dir/${level}-${lane}-sample${sample}.stderr.log"
            time_file="$output_dir/${level}-${lane}-sample${sample}.time.txt"
            if ! /usr/bin/time -v -o "$time_file" "$binary" \
                --lane "$lane" --level "$level" --sample "$sample" --warmups "$warmups" \
                > "$json_file" 2> "$stderr_file"; then
                cat "$stderr_file" >&2
                exit 1
            fi
            cat "$json_file" >> "$raw"
        done
    done
done

python3 - "$source_root" "$output_dir/source-manifest-after.sha256" "$output_dir/source-manifest-before.sha256" "$output_dir/source-manifest-stability.json" <<'PY'
import hashlib
import json
import pathlib
import sys

root = pathlib.Path(sys.argv[1])
after = pathlib.Path(sys.argv[2])
before = pathlib.Path(sys.argv[3])
stability = pathlib.Path(sys.argv[4])
roots = ("crates/litchi-core", "crates/litchi-opc", "crates/litchi-ooxml-common", "crates/litchi-xlsx", "crates/soapberry-zip")
files = [root / "Cargo.toml", root / "Cargo.lock"]
for relative_root in roots:
    base = root / relative_root
    files.extend(path for path in base.rglob("*") if path.is_file())
with after.open("w", encoding="utf-8") as stream:
    for path in sorted(set(files), key=lambda value: value.relative_to(root).as_posix()):
        stream.write(f"{hashlib.sha256(path.read_bytes()).hexdigest()}  {path.relative_to(root).as_posix()}\n")
equal = before.read_bytes() == after.read_bytes()
stability.write_text(
    json.dumps(
        {
            "schema": "xlsx-custom-data-source-stability-v1",
            "equal": equal,
            "before": before.name,
            "after": after.name,
        },
        indent=2,
    )
    + "\n",
    encoding="utf-8",
)
if not equal:
    raise SystemExit("source manifest changed during profile capture")
PY

if (( smoke )); then
    printf 'smoke passed; results were removed from %s\n' "$output_dir"
else
    printf '%s profile written to %s\n' "$capture_kind" "$output_dir"
fi
