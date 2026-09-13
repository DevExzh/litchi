# ODS `sheet_metadata` corrected-fixture profile

This is absolute phase evidence for the frozen ODS metadata source. The retained
harness generates a deterministic synthetic ODF content document with valid
explicit ranges on both detective endpoints. It records parse/index, selector
lookup, staged update, commit, and end-to-end timing together with the
implementation's Work and reserved Memory counters. It makes no throughput,
speedup, or all-ODS performance claim.

The frozen source and harness inputs are verified by
[final-corrected/source-hashes.sha256](final-corrected/source-hashes.sha256) and
[final-corrected/source-hashes.check](final-corrected/source-hashes.check).
The release build passed before measurement; its log and exit status are
[build.log](final-corrected/build.log) and [build.exit](final-corrected/build.exit). The
retained binary digest is [binary.sha256](final-corrected/binary.sha256).
The source generator emits `office:scripts`,
`office:font-face-decls`, and `office:automatic-styles` before the body.
The metadata fixture uses
`SheetN.A1:SheetN.A1` for both range endpoints, and the source-backed
fixture uses `source.ods#SheetN.A1`. No ZIP member is read or written.

The profile ran on 2026-09-13 UTC with three warmups and fifteen measured
iterations per process, pinned to CPU 2. The start and end host snapshots are
[host.txt](final-corrected/host.txt) and
[host-end.txt](final-corrected/host-end.txt). CPU 2 was pinned but not
exclusively reserved: both snapshots show concurrent
`check_crate_boundaries.py` processes near 100% CPU on other cores, so the
timings are observations under shared host load.

## Final measurements

The timer includes parse/index work for `parse`, `metadata`, `noop`, and
`edit-*`. Parsing completes before `lookup` and `stage-*`; parsing and
staging complete before `commit-*`. The `Memory` columns are current
`Resource::Memory` budget bytes, not allocator bytes or allocator peak.
`Δ` is relative to the phase's pre-timer point; `final` retains the
snapshot/edit/commit values alive. RSS is the complete process maximum from
`/usr/bin/time -v`. `p99` is the nearest-rank value from fifteen samples.

| Phase / workload | Fixture (sheets × rows × columns) | XML bytes | p50 / p95 / p99 (ms) | Memory Δ / final (bytes) | Work Δ / final (units) | Max RSS (KiB) | Probes |
|---|---:|---:|---:|---:|---:|---:|---:|
| Parse, sparse medium | 4 × 512 × 8 | 1,415,831 | 26.995 / 28.046 / 28.046 | 32,616,517 / 32,616,517 | 5,021,255 / 5,021,255 | 30,408 | 0 |
| Parse, metadata medium | 4 × 512 × 8 | 2,223,911 | 137.404 / 139.181 / 139.181 | 46,505,021 / 46,505,021 | 11,547,078 / 11,547,078 | 98,128 | 0 |
| Metadata read, medium | 4 × 512 × 8 | 2,223,911 | 137.481 / 139.211 / 139.211 | 46,505,021 / 46,505,021 | 11,547,222 / 11,547,222 | 98,164 | 75 |
| Lookup, grid 32 | 1 × 32 × 32 | 85,610 | 0.321 / 0.325 / 0.325 | 0 / 1,990,540 | 278,528 / 580,858 | 4,320 | 15,360 |
| Lookup, grid 64 | 1 × 64 × 64 | 338,634 | 2.316 / 2.323 / 2.323 | 0 / 7,909,580 | 2,162,688 / 3,360,090 | 9,380 | 61,440 |
| Stage one, grid 64 | 1 × 64 × 64 | 338,634 | 0.000580 / 0.000670 / 0.000670 | 1,052 / 7,910,632 | 88 / 1,197,490 | 9,232 | 15 |
| Stage batch, grid 32 | 1 × 32 × 32 | 85,610 | 1.202 / 1.212 / 1.212 | 1,077,248 / 3,067,788 | 851,968 / 1,154,298 | 4,596 | 15,360 |
| Stage batch, grid 64 | 1 × 64 × 64 | 338,634 | 8.177 / 8.184 / 8.184 | 4,308,992 / 12,218,572 | 6,553,600 / 7,751,002 | 10,212 | 61,440 |
| Commit one, grid 64 | 1 × 64 × 64 | 338,634 | 6.459 / 6.492 / 6.492 | 7,912,416 / 15,823,048 | 1,536,721 / 2,734,211 | 16,028 | 15 |
| Commit batch, grid 32 | 1 × 32 × 32 | 85,610 | 6.374 / 6.493 / 6.493 | 4,894,604 / 7,962,392 | 1,849,221 / 3,003,519 | 9,920 | 15,360 |
| Commit batch, grid 64 | 1 × 64 × 64 | 338,634 | 29.558 / 31.291 / 31.291 | 19,525,836 / 31,744,408 | 10,526,821 / 18,277,823 | 42,376 | 61,440 |
| No-op, grid 64 | 1 × 64 × 64 | 338,634 | 6.038 / 6.346 / 6.346 | 7,909,580 / 7,909,580 | 1,197,402 / 1,197,402 | 9,352 | 0 |
| Edit one, grid 64 | 1 × 64 × 64 | 338,634 | 12.693 / 12.814 / 12.814 | 15,823,048 / 15,823,048 | 2,734,211 / 2,734,211 | 15,972 | 15 |
| Edit batch, grid 32 | 1 × 32 × 32 | 85,610 | 9.293 / 9.417 / 9.417 | 7,962,392 / 7,962,392 | 3,003,519 / 3,003,519 | 9,916 | 15,360 |
| Edit batch, grid 64 | 1 × 64 × 64 | 338,634 | 44.364 / 46.871 / 46.871 | 31,744,408 / 31,744,408 | 18,277,823 / 18,277,823 | 42,372 | 61,440 |

Each raw stdout and `.time` receipt is retained under
[final-corrected/](final-corrected/). The complete command ledger is
[commands.txt](final-corrected/commands.txt).

## Scoped observations

The sparse medium parse reserves 32.6 MB against 1.416 MB of XML (23.0×). The
metadata medium parse reserves 46.5 MB against 2.224 MB (20.9×). These are budget
reservations, not allocator measurements. RSS is reported separately. The indexed
`Span` retains owned namespace, local-name, qualified-name, attributes, child
indexes, and ranges for each XML element
([index.rs](../../../../crates/litchi-ods/src/sheet_metadata/index.rs#L82));
the scan publishes spans for start and empty events
([index.rs](../../../../crates/litchi-ods/src/sheet_metadata/index.rs#L519)).
This identifies the memory shape of this fixture only.

Selector lookup and staging report positive Work under the frozen source. Grid
32 lookup adds 278,528 units over 1,024 probes, and grid 64 adds 2,162,688 over
4,096 probes. Stage-batch adds 851,968 units for grid 32 and 6,553,600 for grid
64. Stage-one adds 88 units. The implementation charges eight units and checks
the execution context before each candidate comparison
([index.rs](../../../../crates/litchi-ods/src/sheet_metadata/index.rs#L59)).
Staging shares the selector budget across target, baseline, effective, key, and
clear-state resolution
([mod.rs](../../../../crates/litchi-ods/src/sheet_metadata/mod.rs#L1244)).
The lookup grid begins sparse and measures physical position resolution. The
stage grid also begins sparse and creates a source owner on each selected cell;
metadata owners are present only in the parse/metadata XML workloads.

The selector path still scans physical table rows and row cells, including a
merge fallback. These absolute grid points leave scan scaling as an open
implementation question. They do not establish a complexity bound, allocator
bound, or general latency result.

## Historical comparison boundary

The earlier [pre-accounting profile](pre-accounting/report.md) and
[post-accounting candidate snapshot](post-accounting-candidate/) remain
available for the accounting investigation. Their raw values were captured
with the same historical shorthand detective endpoint and earlier source
snapshots. They are not a corrected-fixture pre/post pair and are not used to
claim a speedup, regression, or final-source latency change. The corrected
fixture and frozen-source values above are the authoritative absolute profile
for this batch.

## Replay and validation

From the repository root, rebuild the harness and run any workload using the
command shape below. The release binary was built with four Cargo jobs.

```sh
CARGO_BUILD_JOBS=4 cargo build --manifest-path docs/report/spec-gap-validation-evidence/ods-sheet-metadata/harness/Cargo.toml --release
taskset -c 2 docs/report/spec-gap-validation-evidence/ods-sheet-metadata/harness/target/release/ods-meta-profile \
  --workload edit-batch --sheets 1 --rows 64 --columns 64 --warmups 3 --iterations 15
```

The final source hash manifest extends the root
[source-hashes-final-freeze-0202.sha256](source-hashes-final-freeze-0202.sha256)
with the harness manifest, lockfile, and source. It covers the frozen 15 ODS
production and test paths, including facade, authoring, model
structure/consolidation/label range/detective, sheet metadata, context, and XML
conformance sources. The hash check passes. The current
focused source contains no `__dump_native_sheet_metadata_candidate` helper;
the historical skip-filtered command remains documented only inside
[pre-accounting/report.md](pre-accounting/report.md).

No native producer corpus, ZIP round trip, allocator tracer, exclusive CPU
reservation, or confidence interval was used. The profile is bounded
phase evidence for the frozen corrected fixture.
