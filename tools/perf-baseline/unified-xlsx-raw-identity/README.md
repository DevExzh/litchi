# XLSX raw identity lifecycle probe

This probe records the immediate pre-repair pair
`0befb112d96d2409e6f40c9abfeb4813665905ad` and
`5726f03c7db2dd7dd539000b45f109f14ea8c33a`, then carries the matched
lifecycle into the current assembly. The versioned OPC ingress is deliberate:
the historical source uses its one-argument `OpcPackage::from_vec` default,
while the current source uses the required-limit
`OpcPackage::from_vec(bytes, SourceLimits::default())`. `--api-variant` labels
the report; Cargo features select the source-compatible constructor.

The current v2 harness is based on `a3b649ca857673c9c0ce62e04c7d44e7ba78c502`.
After the adapter change its source-file SHA-256 is
`6d41c01c4df68c0921989b6230d20a539d080cda3a33ce559797ba0ea3398730`; the
adapter block (the `open_opc` documentation and both cfg branches) is
`72d56f8c5ec104e96037a208e362c31c6da462db39d2ba8aca22b34a88a6f313`.
The native v1 source files at both historical revisions have SHA-256
`6f5e5606afa342f87a7c1e93538b7302878a6b45113cef7b6514f62a7f1cdb34`.
Do not reuse a 5726-only result as a performance qualification for the later
current assembly.

Each corpus contains a small and a sixteen-part 512 KiB multipart package,
with one-cache and eight-cache variants. For every corpus it measures owned
open, the first selected deferred read, raw slicer insertion, raw timeline
batch insertion, typed slicer replacement, and typed timeline replacement.
The correctness pass outside the measured region checks the serialized output
bytes, native relationship IDs and paths, untouched worksheet bytes, patch
replay, inverse restoration, stale-patch refusal, no-op preservation, and
deferred decode counters.

The harness is in `tools/perf-baseline/src/xlsx_raw_identity.rs`; the normal
perf-baseline binary and the existing allocator binary both dispatch the
`xlsx-raw-identity` selector. The allocator binary reports operation-scoped
allocation samples. The normal binary reports the same elapsed samples with
an explicit unavailable allocation status.

The coordinator must keep captures outside the checkout and reuse existing
target directories. The recorded directories for this pair are
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63`
for the pinned historical source and
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current` for the
current assembly. This change creates no target, evidence directory, or
binary.

The feature selection is part of the reproducibility contract. In the current
manifest, `default = ["current-api", "current-xlsx-raw-opc"]`, so the ordinary
default build and `--all-features` both select the required-limit constructor;
when both current and legacy markers are present, the current cfg wins. A
pinned historical checkout is built with `--no-default-features` and the
explicit `legacy-xlsx-raw-opc` marker, so it selects the one-argument
constructor. The historical manifest must contain that marker before the
command (the source revision predates the marker):

```sh
HIST_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/reader-paths-e63
CURRENT_TARGET=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current

# Run this in the exact 0bef or 5726 checkout after adding the recorded
# legacy marker to its temporary manifest overlay. The overlay is not part of
# the pinned revision and must not be committed.
CHECKOUT=/path/to/checkout-at-0bef-or-5726
python3 - "$CHECKOUT/tools/perf-baseline/Cargo.toml" <<'PY'
from pathlib import Path
import sys

path = Path(sys.argv[1])
text = path.read_text()
if "legacy-xlsx-raw-opc = []" not in text:
    text = text.replace(
        "allocator-metrics = []\n",
        "allocator-metrics = []\n"
        "current-xlsx-raw-opc = []\n"
        "legacy-xlsx-raw-opc = []\n",
        1,
    )
path.write_text(text)
PY

source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
CARGO_NET_OFFLINE=true cargo build --release --locked --no-default-features \
  --features allocator-metrics,legacy-xlsx-raw-opc \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline-alloc \
  --target-dir "$HIST_TARGET"

taskset -c 0 "$HIST_TARGET/release/litchi-perf-baseline-alloc" \
  xlsx-raw-identity --api-variant original \
  --source-revision 0befb112d96d2409e6f40c9abfeb4813665905ad \
  --iterations 5 --warmups 2 --json <coordinator-evidence>/raw-original.json

taskset -c 0 "$HIST_TARGET/release/litchi-perf-baseline-alloc" \
  xlsx-raw-identity --api-variant candidate \
  --source-revision 5726f03c7db2dd7dd539000b45f109f14ea8c33a \
  --iterations 5 --warmups 2 --json <coordinator-evidence>/raw-candidate.json

# Run this in the current assembly. The default command has the same adapter;
# --all-features proves that current wins if the legacy marker is also enabled.
source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline-alloc \
  --target-dir "$CURRENT_TARGET"

CARGO_NET_OFFLINE=true cargo build --release --locked --all-features \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline-alloc \
  --target-dir "$CURRENT_TARGET"

taskset -c 0 "$CURRENT_TARGET/release/litchi-perf-baseline-alloc" \
  xlsx-raw-identity --api-variant current \
  --source-revision <current-assembly-commit> \
  --iterations 5 --warmups 2 --json <coordinator-evidence>/raw-current.json
```

Run the same selected case under Callgrind for instruction counts, using the
same target, CPU pinning, iterations, warmups, and an evidence path outside
the repository. Record each binary SHA-256 from the JSON report and keep raw
Callgrind captures outside the checkout. No performance conclusion is made by
this source-only commit; the owned-open and first-read costs include the new
per-deferred-part metadata already present in the candidate and require a
matched run before qualification.
