# ODS `sheet_metadata` bounded profile

This is absolute evidence for the final ODS metadata source after the selector
identity and root sibling fixes. It is a deterministic synthetic XML profile,
not a before/after comparison and not a claim about all ODS documents. The
scope follows [ADR 0005](../../../../adr/0005-io-memory-and-performance.md) and
the owner design's [measurement contract](../../ods-sheet-metadata-design.md).
The
profile was run on 2026-09-13 UTC with three warmups and fifteen measured
iterations per process, pinned to CPU 2. The host and toolchain are recorded in
[final-host.txt](receipts/final-host.txt); the source and replay-harness hashes are in
[final-source-hashes.sha256](receipts/final-source-hashes.sha256), with the verification
receipt in [final-source-hashes.check](receipts/final-source-hashes.check).

The retained replay harness is [harness/Cargo.toml](harness/Cargo.toml) and
[harness/src/main.rs](harness/src/main.rs). It generates one
`office:document-content` source with `office:scripts`,
`office:font-face-decls`, and `office:automatic-styles` siblings before the
required body. Each table has the configured physical row and cell grid. The
`metadata` fixture adds one `table:cell-range-source` and one detective owner
to the first cell of every row. No ZIP member is decompressed or written by
the profile.

## Measurements

The elapsed timer includes parse/index work for `parse`, `metadata`, `noop`, and
`edit-*`. For `lookup` and `stage-*`, parsing completes before the timer. For
`commit-*`, parsing and staging complete before the timer. The `Memory` columns
are current `Resource::Memory` budget bytes at the measurement point, not
allocator bytes or peak RSS. `Δ` is relative to the phase's pre-timer point;
`final` is the retained budget with the relevant snapshot/edit/commit values
held alive. Maximum RSS is the process-level value from `/usr/bin/time -v`.
`p99` is the nearest-rank result from fifteen measured samples; full precision
and means are in the linked raw receipts.

| Phase / workload | Fixture (sheets × rows × columns) | XML bytes | p50 / p95 / p99 (ms) | Memory Δ / final (bytes) | Work Δ / final (units) | Max RSS (KiB) | Probes |
|---|---:|---:|---:|---:|---:|---:|---:|
| Parse, sparse medium | 4 × 512 × 8 | 1,415,831 | 26.926 / 27.577 / 27.577 | 32,616,517 / 32,616,517 | 5,021,255 / 5,021,255 | 30,420 | 0 |
| Parse, metadata medium | 4 × 512 × 8 | 2,209,575 | 138.036 / 138.942 / 138.942 | 46,289,981 / 46,289,981 | 11,504,070 / 11,504,070 | 98,100 | 0 |
| Metadata read, medium | 4 × 512 × 8 | 2,209,575 | 137.117 / 139.674 / 139.674 | 46,289,981 / 46,289,981 | 11,504,070 / 11,504,070 | 98,072 | 75 |
| Lookup, grid 32 | 1 × 32 × 32 | 85,610 | 0.040 / 0.047 / 0.047 | 0 / 1,990,540 | 0 / 302,330 | 4,316 | 15,360 |
| Lookup, grid 64 | 1 × 64 × 64 | 338,634 | 0.278 / 0.284 / 0.284 | 0 / 7,909,580 | 0 / 1,197,402 | 9,312 | 61,440 |
| Stage one, grid 64 | 1 × 64 × 64 | 338,634 | 0.000610 / 0.000690 / 0.000690 | 1,052 / 7,910,632 | 0 / 1,197,402 | 9,324 | 15 |
| Stage batch, grid 32 | 1 × 32 × 32 | 85,610 | 0.416 / 0.429 / 0.429 | 1,077,248 / 3,067,788 | 0 / 302,330 | 4,520 | 15,360 |
| Stage batch, grid 64 | 1 × 64 × 64 | 338,634 | 2.109 / 2.207 / 2.207 | 4,308,992 / 12,218,572 | 0 / 1,197,402 | 10,200 | 61,440 |
| Commit one, grid 64 | 1 × 64 × 64 | 338,634 | 6.456 / 6.565 / 6.565 | 7,911,667 / 15,822,299 | 1,536,439 / 2,733,841 | 16,480 | 15 |
| Commit batch, grid 32 | 1 × 32 × 32 | 85,610 | 4.554 / 4.612 / 4.612 | 4,127,628 / 7,195,416 | 798,597 / 1,100,927 | 8,920 | 15,360 |
| Commit batch, grid 64 | 1 × 64 × 64 | 338,634 | 18.479 / 19.053 / 19.053 | 16,457,932 / 28,676,504 | 3,178,597 / 4,375,999 | 29,360 | 61,440 |
| No-op, grid 64 | 1 × 64 × 64 | 338,634 | 5.985 / 6.316 / 6.316 | 7,909,580 / 7,909,580 | 1,197,402 / 1,197,402 | 9,276 | 0 |
| Edit one, grid 64 | 1 × 64 × 64 | 338,634 | 12.929 / 13.006 / 13.006 | 15,822,299 / 15,822,299 | 2,733,841 / 2,733,841 | 16,464 | 15 |
| Edit batch, grid 32 | 1 × 32 × 32 | 85,610 | 6.652 / 6.687 / 6.687 | 7,195,416 / 7,195,416 | 1,100,927 / 1,100,927 | 8,864 | 15,360 |
| Edit batch, grid 64 | 1 × 64 × 64 | 338,634 | 26.774 / 27.732 / 27.732 | 28,676,504 / 28,676,504 | 4,375,999 / 4,375,999 | 29,352 | 61,440 |

Receipts are grouped by phase: [parse outputs](receipts/final-parse-medium.out),
[metadata outputs](receipts/final-parse-metadata-medium.out), [lookup outputs](receipts/final-lookup-grid-64.out),
[staging outputs](receipts/final-stage-batch-grid-64.out), [commit outputs](receipts/final-commit-batch-grid-64.out),
and [end-to-end outputs](receipts/final-edit-batch-grid-64.out). The matching
`.time` files contain the RSS values in the table. The complete final receipt
set remains in this directory for the 32-cell and 64-cell scaling points.

Representative whole-process counters from `perf stat` are below. They include
source generation, warmup, all measured iterations, and process setup; they are
diagnostic counters for these runs and do not identify a causal hotspot.

| Workload | Cycles | Instructions | IPC | Branches | Branch misses | Cache misses | Minor / major faults |
|---|---:|---:|---:|---:|---:|---:|---:|
| Parse, sparse medium | 2,400,505,863 | 6,419,983,331 | 2.674 | 1,289,459,737 | 2,859,456 | 4,407,187 | 57,139 / 0 |
| Parse, metadata medium | 11,551,476,644 | 26,498,079,642 | 2.294 | 5,429,101,477 | 9,905,409 | 20,510,960 | 29,612 / 0 |
| Lookup, grid 64 | 568,018,161 | 1,600,442,171 | 2.818 | 323,300,980 | 993,582 | 1,490,093 | 7,100 / 0 |
| Stage batch, grid 64 | 744,917,864 | 2,111,833,020 | 2.835 | 427,368,837 | 1,544,356 | 1,705,184 | 21,998 / 0 |
| Commit batch, grid 64 | 2,342,040,195 | 7,141,562,884 | 3.049 | 1,506,011,341 | 4,950,766 | 8,397,311 | 20,179 / 0 |

The raw counter receipts are [parse](receipts/final-parse-medium.perf.stat),
[metadata](receipts/final-parse-metadata-medium.perf.stat),
[lookup](receipts/final-lookup-grid-64.perf.stat), [stage](receipts/final-stage-batch-grid-64.perf.stat),
and [commit](receipts/final-commit-batch-grid-64.perf.stat).

## Findings scoped to this implementation

The sparse medium scan reserves 32.6 MB against 1.416 MB of XML (23.0×), and
the metadata medium scan reserves 46.3 MB against 2.210 MB of XML (20.9×).
Those are budget reservations. The process RSS values are separate
measurements, and this harness does not measure allocator allocation counts or
allocator peak live bytes. The implementation's `Span` stores owned namespace,
local-name, qualified-name, attributes, child indexes, and source ranges for
every indexed XML element ([index.rs](../../../../../crates/litchi-ods/src/sheet_metadata/index.rs#L31));
the scan publishes a span for each start/empty event ([index.rs](../../../../../crates/litchi-ods/src/sheet_metadata/index.rs#L437)).
The observed ratio is therefore a focused memory-shape signal for this
synthetic source, not a production-wide amplification factor.

The all-cell lookup p50 rises from 0.040 ms at 32 × 32 to 0.278 ms at 64 × 64
(7.0× for 4× probes). Stage-batch p50 rises from 0.416 ms to 2.109 ms
(5.1× for 4× staged cells). These timers exclude parse/index work. The
selector path walks table rows and row cells, with a merge fallback over cells
([index.rs](../../../../../crates/litchi-ods/src/sheet_metadata/index.rs#L178)); each
source staging call resolves the target, baseline, and effective value before
recording its operation ([mod.rs](../../../../../crates/litchi-ods/src/sheet_metadata/mod.rs#L1139)).
The lookup and stage samples add **zero** work units because those loops do not
call `context.consume(Resource::Work, ...)` or a context check. This is an
unresolved boundedness and cancellation-accounting gap under the design's
accepted eight-units-per-selector-comparison rule; the profile does not justify
a particular rewrite or speedup claim.

Commit-batch p50 rises from 4.554 ms at 32 × 32 to 18.479 ms at 64 × 64,
while its work and retained-memory deltas rise by about 4×. End-to-end
edit-batch p50 rises from 6.652 ms to 26.774 ms. On the 64 × 64 fixture, the
end-to-end no-op, one-cell edit, and 4,096-cell edit are 5.985 ms, 12.929 ms,
and 26.774 ms respectively. These are operation-specific synthetic values;
there is no reconstructed pre-fix binary, native producer corpus, ZIP
round-trip, confidence interval, or broad latency/regression claim here.

## Replay and validation

From the repository root, rebuild the retained harness and replay any row with
the command shape below (the evidence runs used the release binary directly so
`/usr/bin/time` and `perf` measured only the process):

```sh
cargo build --manifest-path docs/report/spec-gap-validation-evidence/ods-sheet-metadata/harness/Cargo.toml --release
taskset -c 2 docs/report/spec-gap-validation-evidence/ods-sheet-metadata/harness/target/release/ods-meta-profile \
  --workload edit-batch --sheets 1 --rows 64 --columns 64 --warmups 3 --iterations 15
```

The focused transaction gate receipt was captured with this historical command:

```sh
cargo test -p litchi-ods --test ods_sheet_metadata_transactions --quiet -- \
  --skip __dump_native_sheet_metadata_candidate
```

The retained receipt in [final-focused-tests.log](receipts/final-focused-tests.log)
reports 35 passed tests. The skip addressed a temporary dump helper during
that capture; the helper is absent from the retained baseline source and the
current source, so this command is historical and must not be used for a fresh
gate. The temporary environment note about `/tmp` is not a native-producer
profile. A current gate should run the focused target without a skip.
