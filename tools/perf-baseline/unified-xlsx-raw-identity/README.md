# XLSX raw identity lifecycle probe

This probe compares the same owned lifecycle on the immediate pre-repair
revision `0befb112d96d2409e6f40c9abfeb4813665905ad` and candidate
`5726f03c7db2dd7dd539000b45f109f14ea8c33a`. The harness has one API surface;
`--api-variant` labels the report and does not select a compatibility adapter.

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

The coordinator must provide an existing target directory and keep captures
outside the checkout. At this handoff the existing coordinator target root is
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets`; this commit creates
no target, evidence directory, or binary. Select an assigned existing
subdirectory under that root rather than reusing a target from another
measurement.

After consumer integration releases the build hold, use the same command line
on both revisions, with the harness present in each checkout:

```sh
TARGET_DIR=/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/<assigned-existing-target>

CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline-alloc \
  --target-dir "$TARGET_DIR"

taskset -c 0 "$TARGET_DIR/release/litchi-perf-baseline-alloc" \
  xlsx-raw-identity --api-variant original \
  --source-revision 0befb112d96d2409e6f40c9abfeb4813665905ad \
  --iterations 5 --warmups 2 --json <coordinator-evidence>/raw-original.json

taskset -c 0 "$TARGET_DIR/release/litchi-perf-baseline-alloc" \
  xlsx-raw-identity --api-variant candidate \
  --source-revision 5726f03c7db2dd7dd539000b45f109f14ea8c33a \
  --iterations 5 --warmups 2 --json <coordinator-evidence>/raw-candidate.json
```

Run the same selected case under Callgrind for instruction counts, using the
same target, CPU pinning, iterations, warmups, and an evidence path outside
the repository. Record each binary SHA-256 from the JSON report and keep raw
Callgrind captures outside the checkout. No performance conclusion is made by
this source-only commit; the owned-open and first-read costs include the new
per-deferred-part metadata already present in the candidate and require a
matched run before qualification.
