#!/usr/bin/env bash
set -euo pipefail

if [[ -n "${RUSTFLAGS:-}" || -n "${CARGO_ENCODED_RUSTFLAGS:-}" || -n "${RUSTC_BOOTSTRAP:-}" ]]; then
    echo "profile runner refuses compiler override flags" >&2
    exit 2
fi

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
root_dir="$(cd -- "$script_dir/../../../.." && pwd)"
profile_dir="$script_dir"
source_commit="16102fe751d7c5492042330f1bd1f49c304495f0"
result_dir="$(realpath -m -- "${XLSB_CUSTOM_DATA_RESULT_DIR:-$profile_dir/receipts/run-$(date -u +%Y%m%dT%H%M%SZ)-$$}")"
source_root="${XLSB_CUSTOM_DATA_SOURCE_ROOT:-}"
target_dir=""

if [[ -e "$result_dir" ]] && [[ -n "$(find "$result_dir" -mindepth 1 -maxdepth 1 -print -quit)" ]]; then
    echo "result directory is not fresh: $result_dir" >&2
    exit 2
fi
mkdir -p "$result_dir/raw"

cleanup() {
    status=$?
    if [[ -n "$target_dir" && -d "$target_dir" ]]; then
        find "$target_dir" -depth -delete
    fi
    exit "$status"
}
trap cleanup EXIT

if [[ -z "$source_root" ]]; then
    source_root="$(mktemp -d /var/tmp/litchi-xlsb-custom-data-source.XXXXXX)"
    git --no-replace-objects -C "$root_dir" worktree add --detach --no-checkout "$source_root" "$source_commit"
    git --no-replace-objects -C "$source_root" sparse-checkout init --no-cone
    git --no-replace-objects -C "$source_root" sparse-checkout set --no-cone \
        /Cargo.toml /rust-toolchain.toml /.cargo/ \
        /crates/litchi-xlsb/ /crates/litchi-opc/ /crates/litchi-core/ \
        /crates/litchi-xldm/ /crates/litchi-sheet/ /crates/litchi-ooxml-common/ \
        /crates/litchi-drawingml/ /crates/litchi-spreadsheet-drawing/ \
        /crates/soapberry-zip/ /crates/xml-minifier/ /crates/xml-minifier-macros/ \
        /docs/report/spec-gap-validation-evidence/xlsb-custom-data-design.md \
        /docs/report/spec-gap-validation-evidence/xlsb-custom-data-owner-final/
    git --no-replace-objects -C "$source_root" read-tree -mu HEAD
fi
actual_source="$(git --no-replace-objects -C "$source_root" rev-parse HEAD)"
if [[ "$actual_source" != "$source_commit" ]]; then
    echo "source checkout is $actual_source, expected $source_commit" >&2
    exit 2
fi

clean_harness="$source_root/docs/report/spec-gap-validation-evidence/xlsb-custom-data-performance/harness"
mkdir -p "$clean_harness"
for file in Cargo.toml Cargo.lock .gitignore main.rs support.rs; do
    cp "$profile_dir/harness/$file" "$clean_harness/$file"
done

target_dir="$(mktemp -d /var/tmp/litchi-xlsb-custom-data-target.XXXXXX)"
export CARGO_INCREMENTAL=0
export LC_ALL=C
unset RUSTFLAGS CARGO_ENCODED_RUSTFLAGS RUSTC_BOOTSTRAP RUSTDOCFLAGS

{
    echo "source_root=$source_root"
    echo "source_commit=$source_commit"
    rustc -Vv
    cargo -V
    uname -a
    echo "cargo_incremental=$CARGO_INCREMENTAL"
    echo "compiler_flags=none"
    echo "allocator=CountingAllocator (process-local GlobalAlloc observer)"
    echo "copy_scope=explicit validation copies only; internal production copies are uninstrumented"
    echo "build_profile=release"
} >"$result_dir/toolchain.txt"
printf '%s\n' "$source_root" >"$result_dir/source-root.txt"

source_extra_args=()
for file in \
    "$profile_dir/PLAN.md" \
    "$profile_dir/README.md" \
    "$profile_dir/source_manifest.py" \
    "$profile_dir/verify.py" \
    "$profile_dir/summarize.py" \
    "$profile_dir/run_profile.sh" \
    "$profile_dir/replay.sh" \
    "$profile_dir/harness/Cargo.toml" \
    "$profile_dir/harness/Cargo.lock" \
    "$profile_dir/harness/.gitignore" \
    "$profile_dir/harness/main.rs" \
    "$profile_dir/harness/support.rs"; do
    if [[ -f "$file" ]]; then
        source_extra_args+=(--extra "$file")
    fi
done
python3 "$profile_dir/source_manifest.py" \
    --root "$source_root" \
    --commit "$source_commit" \
    --output "$result_dir/source-manifest.txt" \
    "${source_extra_args[@]}" \
    >"$result_dir/source-manifest.log"

cargo metadata --format-version=1 --locked --offline \
    --manifest-path "$clean_harness/Cargo.toml" \
    >"$result_dir/metadata.json"

CARGO_TARGET_DIR="$target_dir" cargo build --release --locked --offline \
    --manifest-path "$clean_harness/Cargo.toml" \
    >"$result_dir/build.log" 2>&1
clean_binary="$target_dir/release/xlsb-custom-data-performance"
if [[ ! -x "$clean_binary" ]]; then
    echo "clean profile binary was not produced" >&2
    exit 1
fi
cp "$clean_binary" "$result_dir/binary"
chmod 755 "$result_dir/binary"
sha256sum "$result_dir/binary" >"$result_dir/binary.sha256"

printf '%s\n' \
    "source=$source_commit" \
    "build=cargo build --release --locked --offline --manifest-path harness/Cargo.toml" \
    "binary=$result_dir/binary" >"$result_dir/commands.txt"
for level in small medium large; do
    for lane in read snapshotclone noop editpayload rename-many-to-one mixedinsertremove patchinverse publicapply; do
        stem="${lane}_${level}"
        command_line="$result_dir/binary --lane $lane --level $level --samples 5 --warmups 1"
        printf '%s\n' "$command_line" >>"$result_dir/commands.txt"
        /usr/bin/time -v -o "$result_dir/raw/$stem.time.txt" \
            "$result_dir/binary" --lane "$lane" --level "$level" --samples 5 --warmups 1 \
            >"$result_dir/raw/$stem.jsonl" \
            2>"$result_dir/raw/$stem.stderr.log"
    done
done

python3 - "$result_dir/raw" "$result_dir/fixture-sha256.txt" <<'PY'
import json
import sys
from pathlib import Path

raw_dir = Path(sys.argv[1])
output = Path(sys.argv[2])
rows = []
for path in sorted(raw_dir.glob("*.jsonl")):
    for line in path.read_text().splitlines():
        if line.strip():
            rows.append(json.loads(line))
seen = {}
for row in rows:
    fixture = row["fixture"]
    key = (row["size_class"], fixture["source_sha256"])
    seen[key] = fixture
lines = []
for (level, source_sha), fixture in sorted(seen.items()):
    lines.append(
        "\t".join(
            [
                level,
                source_sha,
                fixture["source_fnv1a64"],
                fixture["connections_sha256"],
                fixture["opaque_sha256"],
                str(fixture["source_bytes"]),
                str(fixture["connections_bytes"]),
                str(fixture["opaque_bytes"]),
            ]
        )
    )
output.write_text(
    "# level\tsource_sha256\tsource_fnv1a64\tconnections_sha256\topaque_sha256\tsource_bytes\tconnections_bytes\topaque_bytes\n"
    + "\n".join(lines)
    + "\n"
)
PY

python3 "$profile_dir/verify.py" --results "$result_dir"
printf '%s\n' "$result_dir"
