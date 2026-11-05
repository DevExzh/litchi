#!/usr/bin/env bash
# Build the M0a S10 allocation probe against one revision's facade.
#
# usage: mem_probe_build.sh <repo-root> <before|after> <work-dir> <target-dir> [source.rs bin-name]
#
# Without the last two arguments it builds the allocation probe
# (mem_probe.rs as s10-mem-probe).
#
# The probe's [dependencies] are the revision's tools/perf-baseline
# [dependencies], verbatim except for absolute paths, so Cargo unifies the
# same leaf features as the timed harness. The revision's harness Cargo.lock
# seeds the probe's lock, so every third-party version is the harness's.
# Default release profile (no LTO): allocation counts, not timings.
set -euo pipefail
repo=$(cd "$1" && pwd)
leg=$2
work=$3
target=$4
here=$(cd "$(dirname "$0")" && pwd)
source=${5:-mem_probe.rs}
bin=${6:-s10-mem-probe}
mkdir -p "$work"
{
  printf '[package]\nname = "%s"\nversion = "0.1.0"\nedition = "2024"\npublish = false\n\n' "$bin"
  printf '[[bin]]\nname = "%s"\npath = "%s/%s"\n\n' "$bin" "$here" "$source"
  printf '[features]\nbefore = []\nafter = []\n\n'
  # The [dependencies] table of the harness, with ../../ made absolute.
  awk '/^\[dependencies\]/{on=1; print; next} /^\[/{on=0} on' \
    "$repo/tools/perf-baseline/Cargo.toml" | sed "s#\"\.\./\.\./#\"$repo/#g"
  printf '\n[workspace]\n'
} > "$work/Cargo.toml"
cp "$repo/tools/perf-baseline/Cargo.lock" "$work/Cargo.lock"
# The probe lives outside the repository, so pin the repository's toolchain.
toolchain=$(sed -n 's/^channel = "\(.*\)"/\1/p' "$repo/rust-toolchain.toml")
env RUSTUP_TOOLCHAIN="$toolchain" CARGO_TARGET_DIR="$target" CARGO_INCREMENTAL=0 \
  CARGO_NET_OFFLINE=true LC_ALL=C TZ=UTC \
  cargo build --manifest-path "$work/Cargo.toml" --offline --release --features "$leg"
