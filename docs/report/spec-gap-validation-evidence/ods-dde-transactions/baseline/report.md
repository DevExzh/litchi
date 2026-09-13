# ODS inert DDE metadata: read-only baseline

This receipt records the exact pre-transaction baseline at commit
`5347801fc2e8086d67fc6c140a023aeedd260c91` (`perf(ods): binary search sheet metadata cell selectors`).
The source is checked out in a detached sparse worktree for replay. The complete
input source, harness, lockfile, and 15-path source manifest are retained in
this evidence directory; `source-hashes.sha256` and `hash-check.txt` verify the
checkout used for the measurements.

The baseline API has `dde::Snapshot::parse` and read accessors only. It has no
DDE `stage`, `commit`, `noop`, or `edit` transaction phase, and it does not
accept the core `ExecutionContext`. Accordingly, its `work_units` and
`memory_budget_bytes` are recorded as `unavailable`. These results are an
absolute read-only baseline and must not be used as a speedup claim against a
later transaction API.

## Workloads and timing boundaries

The standalone generator in [`../harness/src/main.rs`](../harness/src/main.rs)
creates a synthetic ODF content stream with explicit `office` and `table`
namespaces, `office:version="1.4"`, and the required self-closing
`table:table-column` child in every worksheet and cache table. Each worksheet
row has one empty self-closing cell. `small` has one sheet source, two links, and 8x8 cached cells;
`large-cache` has one sheet source, one link, and a 256x256 cached table;
`many-links` has four sheet sources and 2,048 links with 1x1 cached tables.
All DDE identifiers use inert `file:///never/...` topics; the harness does
not contact a native producer. Cached cells remain self-closing and contain no
character data, matching the baseline reader's bounded DDE-container grammar.
The generated small, large-cache, and many-links streams each passed complete
ODF 1.4 RNG validation; the receipt is
[`generated-shapes-schema-validation.json`](../fixtures/generated-shapes-schema-validation.json).

`parse` starts its timer immediately before `Snapshot::parse` and retains only
cheap lengths afterward. `read` parses before the timed region, then reads all
sheet-source fields, link-source fields, and cached-table lengths. Therefore
`read` measures accessor traversal after parsing; it does not include parsing.
Each lane uses three warmups and 15 measured iterations. The table reports
mean, p50, p95, and p99 per-iteration nanoseconds. `probes` is the cumulative
accessor count across all 15 measured iterations, not a per-iteration count.

Allocation columns come from the harness's process-local counting allocator;
they are requested/released bytes and live-byte deltas for the measured
region, not a general allocator profile. `Maximum resident set size` is the
OS-reported whole-process maximum from `/usr/bin/time -v`; it includes startup
and, for `read`, the snapshot parsed before the timer. It is not the core
Memory budget counter or a claim about peak allocation.

## Results

| lane | source bytes | mean ns | p50 ns | p95 ns | p99 ns | alloc calls p50/max | requested bytes p50/max | peak live delta p50/max | max RSS KiB | probes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| parse-small | 9,845 | 28,163 | 27,950 | 29,961 | 29,961 | 264 / 264 | 25,673 / 25,673 | 12,454 / 12,454 | 2,280 | 0 |
| read-small | 9,845 | 22 | 20 | 30 | 30 | 0 / 0 | 0 / 0 | 0 / 0 | 2,280 | 45 |
| parse-large-cache | 4,132,098 | 8,252,339 | 8,203,595 | 8,961,108 | 8,961,108 | 66,128 / 66,128 | 8,071,563 / 8,071,563 | 4,134,342 / 4,134,342 | 14,388 | 0 |
| read-large-cache | 4,132,098 | 38 | 30 | 180 | 180 | 0 / 0 | 0 / 0 | 0 / 0 | 14,440 | 30 |
| parse-many-links | 863,375 | 5,759,717 | 5,755,525 | 5,929,346 | 5,929,346 | 55,445 / 55,445 | 4,736,697 / 4,736,697 | 1,251,345 / 1,251,345 | 4,584 | 0 |
| read-many-links | 863,375 | 5,124 | 5,120 | 5,290 | 5,290 | 0 / 0 | 0 / 0 | 0 / 0 | 4,556 | 30,780 |

The machine-readable copy is [`raw.csv`](raw.csv); per-lane stdout and
`/usr/bin/time -v` receipts are retained beside it. All six commands exited
with status 0. The build used the pinned toolchain and `cargo build --release
--offline`; the exact commands are in [`commands.txt`](commands.txt).

## Replay and limits

Use a separate worktree so a candidate commit is not removed while replaying
the baseline:

```sh
EVIDENCE=/absolute/path/to/ods-dde-transactions
BASE=/var/tmp/ods-dde-baseline-replay
TARGET=/var/tmp/ods-dde-baseline-replay-target
git worktree add --detach "$BASE" 5347801fc2e8086d67fc6c140a023aeedd260c91
mkdir -p "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/harness/src"
cp "$EVIDENCE/harness/Cargo.toml" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/harness/Cargo.toml"
cp "$EVIDENCE/harness/Cargo.lock" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/harness/Cargo.lock"
cp "$EVIDENCE/harness/src/main.rs" "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/harness/src/main.rs"
git -C "$BASE" status --short
CARGO_TARGET_DIR="$TARGET" cargo build \
  --manifest-path "$BASE/docs/report/spec-gap-validation-evidence/ods-dde-transactions/harness/Cargo.toml" \
  --release --locked --offline
taskset -c 2 /usr/bin/time -v -o "$EVIDENCE/replay.time" \
  "$TARGET/release/ods-dde-profile" --workload parse --case small --warmups 3 --iterations 15
git worktree remove --force "$BASE"
```

Run any lane with the command shape in `commands.txt`, replacing the binary
path. The recorded host was shared with unrelated CPU-heavy processes,
including a concurrent `git fsck`; the binary was pinned to CPU 2 but the
machine was not otherwise isolated. This limits conclusions about wall-clock
variance. The receipt is intentionally scoped to these three generated
shapes and these six read-only lanes; it does not establish general ODS
performance or allocator behavior.
