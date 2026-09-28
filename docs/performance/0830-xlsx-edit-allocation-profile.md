# 0830 — XLSX edit allocation attribution

Three dense column-assignment buffers account for **90.45% of requested
allocation bytes** observed inside the real-file XLSX public edit. This changes
the next optimization target: first eliminate the unused column-action owner
map for a cell-only edit, then investigate the two parser maps while retaining
their bounded handling of overlapping column ranges. Production is unchanged
in this diagnostic batch; no latency or memory improvement is claimed.

## Scope and reproducibility

The base is `87eb57182be1622385dc3b28dfc3c7be868dca32`. The fixture is the
8,435-byte LibreOffice `dateAutofilter.xlsx`, SHA-256
`d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4`.
The public operation edits `Munka1!A1` to
`litchi-perf-0638-ordinary-save`, commits, requires a nonempty patch and adopts
the resulting workbook. Replacement drops the old workbook inside the timer.
Each iteration opens a fresh workbook before timing. Serialization, final
owner drop, hashing and semantic readback follow the timer.

Every warmup and measured output equals the separately admitted 0821 reference
bytes: 8,521 bytes, SHA-256
`0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68`.
Public reopen also verifies the target marker, worksheet name, one worksheet,
all eight stored cells and their complete stored-cell projection. Byte equality
covers preservation beyond that projection, including the producer extensions.

The [packet](results/change-0830/README.md) retains source/driver/probe hashes,
the exact dependency lock, host, tool versions, commands, receipts, symbol
disassembly, raw interpreted Heaptrack captures and numerical readers. The
AMD EPYC 9R45 host exposes 32 logical CPUs; captures run serially on CPU 12.
Rust 1.95.0 release builds use opt-level 3, thin LTO, one codegen unit, debug
level 1, unwind panics and two build jobs. The second executable forces frame
pointers. The absolute checkout path is part of this protocol; relocation needs
a new input freeze.

All **23 reports / 559 measured samples** complete: three qualification reports
of three samples, eighteen native reports of thirty samples, and two Heaptrack
reports of five edits. Native reports have three warmups; the other reports
have none. Six native blocks cover every permutation of the three arms.

## Allocation attribution

Both captures independently produce the same totals:

| Population | Allocation calls | Requested bytes |
| --- | ---: | ---: |
| Whole process | 27,539 | 33,446,948 |
| Exact edit owner, five edits | 12,550 | 14,491,270 |
| Exact-owner average per edit | 2,510 | 2,898,254 |
| Three column-assignment traces, five edits | 15 | 13,107,200 |

Attribution requires exactly one `xlsx_edit_profile_0830::edit_region_0830`
frame in the exact frame-pointer executable module. A strict sixteen-digit
Rust symbol hash suffix is normalized. Repeated-owner and wrong-module counts
are zero. The following are disjoint complete allocation traces, not inclusive
caller rows:

| Allocation route | Calls per capture | Requested bytes per capture |
| --- | ---: | ---: |
| Initial `Worksheet::store` → worksheet parser → `start_columns` | 5 | 5,242,880 |
| Required post-write worksheet parse → `start_columns` | 5 | 5,242,880 |
| Value-only rewrite → `validate_column_actions` | 5 | 2,621,440 |

All three call [`Assignments::new`](../../crates/litchi-xlsx/src/column.rs),
which reserves and initializes `2 * COLUMNS` nodes. For the parser's column
state this is 1 MiB; for the validator's source-owner index it is 512 KiB.
The two parser calls are visible at `transaction.rs:1881` and `:2310`; the
validator route enters through the rewrite at `:2180`. Full frames and exact
source locations are retained in the [compact summary](results/change-0830/summary.json)
and [complete compressed analysis](results/change-0830/analysis.json.gz).

The 512 KiB map contributes **18.09%** of the observed requested bytes per edit.
`validate_column_actions` constructs it even when its action map is empty.
An empty-action fast path is therefore the first optimization candidate identified
by this measurement. It must
retain protected-sheet and style/width refusal behavior for actual actions,
and pass correctness and matched performance gates before adoption.

The parser maps provide `O(log COLUMNS)` range assignment against malicious
overlap. Their larger byte share does not authorize replacing them with an
unbounded scan or removing either required parse. A compact representation
needs its own error-order, allocation-failure, overlap and performance evidence.
The earlier rejected [0471 buffer-lifetime experiment](changes/0471-xlsx-rewrite-buffer-lifetime.md)
is not revived: this finding concerns avoidable allocation itself.

Whole-process allocation calls agree with the independent `heaptrack_print`
summary. Its owner-filtered report still emits a global summary, so that line
is not treated as an owner-call cross-check. Independent arithmetic also
recounts every allocation event and requested size, verifies compressed/plain
trace equality, and conserves the size, leaf and full-trace projections.

One whole-process allocation has an unresolved function (73,728 bytes); none
of the owner-attributed traces has an unresolved function. Source-file locations
are absent in two frames per allocation trace, retained explicitly in the
diagnostics. Maximum expanded owner depth is 56 frames. These observations do
not prove complete unwinding. No peak, net-live, retained-byte, physical-copy
or profiler-timing improvement is inferred. Heaptrack interposition and native
Rust allocator instrumentation are distinct methods; their equality is not
assumed.

## Native instrumentation controls

Values below are medians of six process statistics, in milliseconds except
whole-process maximum RSS. Within-process quantiles use nearest rank.

| Arm | p50 ms | p95 ms | p99 ms | Mean ms | RSS KiB |
| --- | ---: | ---: | ---: | ---: | ---: |
| Ordinary / direct | 0.284002 | 0.310327 | 0.373137 | 0.288445 | 8,574 |
| Ordinary / wrapped | 0.282951 | 0.375177 | 0.438072 | 0.292397 | 8,422 |
| Frame pointer / wrapped | 0.281756 | 0.369297 | 0.436862 | 0.288859 | 8,510 |

Paired p50 wrapped/direct is **0.995417 [0.980679–1.015332]**;
frame-pointer/wrapped is **0.998841 [0.981111–1.028213]**. Intervals use
10,000 bootstrap resamples, seed 830830 and sorted endpoints 250/9749.
Both include 1.0. Ten greater-than-5% process-spread flags and all three
aggregate p99/p50 tail flags remain visible in the summary; the frame-pointer
p50 spread is 1.058175. No RSS spread flag fires. Tail variability prevents
any claim of uniformly negligible instrumentation cost. These short controls
are not a before/after optimization comparison.

## Validation and remaining work

All five fresh probe gates pass: formatting, all-target checking, three tests,
warning-denied Clippy and warning-denied rustdoc. Both release builds and their
address-bounded symbol/disassembly checks pass. These are probe gates; a fresh
full production test suite is not claimed for this unchanged production tree.

The first reader preflight fails because its Unicode check assumes the first
sorted leaf is the owner. The corrected check selects the exact owner frame;
that failed attempt and its source remain retained. A later passing preflight
also verifies missing-source-file diagnostics. Analysis and an independent
arithmetic/custody audit pass. Complete analysis JSON is stored with
deterministic gzip to avoid a 96 MB expanded artifact; the initial analysis is
also preserved losslessly with a round-trip hash receipt. No measurements are
discarded or rerun to improve a number.

The owned build target is removed after live-binary validation (3,968 files,
2,861,187,648 bytes), with exact
binary identities retained for offline replay. Unrelated workspace edits are
preserved. The broader non-iWork performance goal remains open; the next batch
should test the empty column-action fast path against this measured baseline.
