# Change 0469: borrow changed SpreadsheetML compaction events

`performance_claim: none; descriptive latency probe and whole-process allocation evidence`

`claim_authorized: false`

Retain production commit `7cc58fc1b`: the slice-backed compactor consumes each
XML event before the next read, so it no longer converts every event to owned
storage. All normalization, attribute checks, namespace handling, semantic
whitespace and error ordering remain. The output reservation and bounded Store
handoff are unchanged. The complete XLSX suite passes 1,245 tests across 59
targets, including differential valid/malformed cases against the old loop.
Scoped formatting, workspace all-feature checking, warning-denied XLSX Clippy
and rustdoc, and crate boundaries pass. All 30 reviewed ADR hashes are unchanged.

## Retention evidence

The authenticated 0468 control and fresh candidate use the same absolute build
path, Rust 1.98.1, release debug level 1, forced frame pointers and unwind tables.
Clean source inventories contain the same 6,992 paths and two compile-time
fixtures; only `compact.rs` differs. Both roles run on CPU 2 with one worker.
The candidate build occurs after control A1 and before candidate B1, without
capture overlap. Tests, captures and Heaptrack postprocessing are serialized.

Whole-process Heaptrack on dense one-percent commit/save, five samples and one
warmup, records 44,815,468 / 39,892,490 allocation calls: 4,922,978 fewer
(10.985%). Temporary allocations are 14,569,551 / 10,633,497 (27.016% fewer).
The rounded peak heap display is `104.38M` for both. These totals include
fixture generation, expected output, warmups, verification and teardown; they
are not per-commit or per-cell counts. No exact peak-byte, copied-byte, or
instrumented-latency claim follows. The substantial allocation reduction and
unchanged semantic checks justify retaining this small work-elimination step.

The normal six-row ABBA probe uses 100 samples and five warmups. Median
observations in milliseconds are:

| Ordinary commit/save | Shape | A1 | B1 | B2 | A2 |
|---|---|---:|---:|---:|---:|
| One cell | Tiny | 0.232680 | 0.230651 | 0.229361 | 0.234001 |
| One cell | Medium | 2.616890 | 2.569201 | 2.568611 | 2.611311 |
| One cell | Dense-wide | 178.803343 | 176.563782 | 176.451770 | 178.035800 |
| One percent | Tiny | 0.423877 | 0.421267 | 0.417081 | 0.424242 |
| One percent | Medium | 10.507999 | 10.179140 | 10.255704 | 10.530288 |
| One percent | Dense-wide | 357.997130 | 356.288987 | 355.599849 | 358.492225 |

All six medians pass the diagnostic directional/drift checks. Dense one-percent
mean, p95 and p99 are mixed: the second pair changes by +0.112%, +1.367% and
+1.121%, respectively. This probe is below the unchanged 500-sample registered
latency minimum; no new registry entry or qualified speedup claim is made.

Normal maximum RSS is 120,928 / 113,340 / 115,124 / 110,764 KiB in ABBA
order. The paired changes are -6.27% / +3.94%, but the control's own RSS drifts
-8.41%; no peak-memory improvement is inferred. Full-guard RSS is
155,624 / 152,584 KiB. Instrumented RSS is kept separately and excluded.

## Regression review and limits

The 201-row, 15-sample/three-warmup full guard passes all workload oracles and
compares 1,205 metrics. It retains 57 latency policy flags and no non-latency
metric flag. Reviewing every mean/p50/p95/p99 at a uniform five-percent trigger
finds 102 cells across 41 rows. The larger mean/median flags are OPC mutated
save (compressible few-large), CFB read-one (incompressible few-large), and CFB
concurrent reads (incompressible few-large); XLS fresh-write large also has
larger tail flags. See the complete machine-readable lists, not a geometric mean.

A separate predeclared seven-row 100/5 ABBA follows those four scenario families.
All four processes pass their workload oracles. Strict ABBA comparison rejects
the set because CFB concurrent-read `max_in_flight_reads` differs: one control
sample observes two concurrent reads while the other vectors observe one.
The rejection is retained; no row or counter is erased to manufacture acceptance.
The raw follow-up medians do not reproduce the large CFB read-one difference
consistently. OPC compressible mutated-save remains +6.69% in the first pair
and +3.13% in the second; this guard does not establish blanket regression-free
behavior. Those paths have unchanged compiled source, but that fact alone does
not dismiss process/layout/noise effects. No global performance claim follows.

The shared comparator needs the same established 0467 treatment of empty,
optional top-level source vectors as absent. Only comparison copies are adjusted,
both roles must expose exactly the same empty paths, and every removed path is
listed. Raw reports, measured values and the frozen policy remain unchanged.
The initial empty-vector rejection, exporter suffix mismatch and fractional
Heaptrack-count parser test failure are documented in the bundle. The parser
was fixed to reject fractional counts; no successful capture was rerun.

Next work should target the larger remaining eager-parser/snapshot scans and
validated full-pass reuse identified in 0468. Event borrowing is a modest step,
not completion of the broader non-iWork program. No new fuzz campaign, native
Office, physical-cold, remote/range, SIMD or worker-scaling result is claimed.

See the [sealed evidence bundle](../results/change-0469/README.md),
[normal/heap summary](../results/change-0469/summary.json), and
[targeted guard observations and rejection](../results/change-0469/review-summary.json).
