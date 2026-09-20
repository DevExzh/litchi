# Independent audit review: 0726 XLS indexed empty-slot setup

This is the independent source and evidence review for the bounded
`0726-xls-empty-slot-setup` pilot. The comparison point is
`3ad29e42dade8cf51c09b498f4085066683a902b`
(`perf(xls): record revised checkpoint rejection`). The review owns
`audit.py` and this document. It does not alter production sources, the pilot
driver, the primary analyzer, or earlier evidence packets, and it never runs
Cargo, native probes, profilers, or instrumented builds.

## Pre-capture source review

The candidate source census must differ from the baseline at exactly one
path:

* `crates/litchi-xls/src/workbook/source.rs`.

The audit binds `candidate-source.json` and the archived source bytes to the
live candidate. It also compares `query_cache.rs` byte-for-byte with the
baseline revision. This makes a query-cache, scanner, fixture, or test-file
change fail the source audit even if it is outside the stated candidate
description.

The candidate keeps the indexed cache's fixed `INDEX_OVERHEAD` at 224 bytes.
0726 adds no retained state and does not introduce a 40-byte charge. The
indexed cache continues to retain locator/checkpoint metadata and occurrence
slots, while replay still decodes selected source frames on demand. No cell
values, errors, strings, or decoded payloads are published in the index.

In `replay_indexed_cell`, worksheet lookup and its `WorksheetNotFound` error
remain before the slot lookup. The candidate binds
`let slots = index.slots_for(row, column)` before constructing workbook path
references, the shared-string resolver, or the replay chain. When the slice
is empty, the branch takes the final execution check and
`owner.ensure_current()` before returning `Ok(None)`. It performs no path
vector, cursor, resolver, or source read setup. When slots are present, the
loop iterates the already-selected slice and retains the existing per-slot
execution checks, decoding, duplicate ordering, and final execution/version
fences.

The static audit checks the existing public integration control
`xls_query_index_cache::an_indexed_missing_target_still_takes_the_trailing_freshness_fence`.
That test mutates the source before the trailing observation, queries an
indexed coordinate with no slots, and requires `SourceChanged` with zero
source reads. The test is baseline-owned and is checked by name and its
actual observation/read assertions; the candidate does not add a duplicate
mock or private replay helper.

The pre-capture audit currently passes these source, archive, build, fixture,
and tooling checks while the candidate is installed. It records the exact
one-file delta, unchanged 224-byte charge, empty-slot ordering and fences,
five baseline and five candidate build receipts, and 17 synthetic negative
controls bound to the current analyzer digest. The terminal freeze and raw
captures remain required before the normal audit can pass.

## Frozen matrix and hard gates

The frozen plan has 12 native cases, 8 repeat cases, four allocator
operations (`q1`, `q2`, `q3`, `q8`), 12 primary budget cases, and four budget
fence cases. The native matrix has 24 groups (six legs and owned/file modes),
the repeat matrix has 16 groups, the allocator matrix has 96 groups, and the
budget fence has 34 baseline/candidate comparison rows. The repeat command
uses its normal 2 MiB default; its positive benefit case is still named
`54016-missing-1048576` so the native and repeat gates refer to the same
missing coordinate while exercising their frozen command settings.

Native timing uses the 5% threshold with the 10 ns absolute exception only
for `q3`, `q8`, and `q3-to-q8-mean`. Every native group must retain exact
semantic outcomes across samples and legs. Repeat q8 uses the separate
`repeat-q8` metric and therefore remains under the strict 5% rule. The
positive benefit gate requires owned `54016-missing-1048576` q8 improvement
of at least 10% at both p50 and mean in both B1/A1 and B2/A2 pairs for both
native and repeat evidence.

All allocator operations are strict in this packet, including q2. The q2
allowance is exactly zero for all six fields: allocation calls, allocated
bytes, deallocation calls, deallocated bytes, peak live delta, and retained
live delta. For a positive-budget missing target, q3 and q8 must equal the
explicit exact delta:

```text
allocation_calls   -1       allocated_bytes    -16
deallocation_calls -1       deallocated_bytes  -16
peak_live_delta   -16       retained_live_delta   0
```

Zero-budget, refusal, value, q1, and q2 routes require exact zero deltas.
Repeated allocator observations must be identical within each phase, and
non-metric report metadata must match between baseline and candidate.

The primary budget matrix requires exact semantic outcomes and an exact
source-metric vector shape. Every budget-fence row requires both semantic
parity and all four counted-source metrics to be unchanged. A changed
source-metric vector sets `route_changed` and fails the hard gate; this is
also covered by the synthetic exact and changed-metric controls. The audit
does not convert a route change into a timing or benefit allowance.

## Independent terminal custody audit

`audit.py` independently revalidates the frozen plan, freeze record, source
start/end maps, baseline and candidate build receipts, immutable probe and
fixture maps, case manifest, capture configuration, binary identities and
hashes, command exit codes, every raw output hash, and every manifest's
start/end bindings. It accepts an owned binary removed during cleanup only
when `cleanup.json` contains an exact path/hash witness matching the frozen
manifest. A live binary with a mismatching hash fails the same check.

It then recomputes:

* 24 native groups, all 100 samples and six legs, semantic parity, timing
  pairs, and the configured absolute exceptions;
* 16 repeat groups, exact repetition/found counts, normalized q8 timing, and
  both benefit pairs;
* 96 allocator groups, repeat stability, metadata parity, exact strict
  fields, zero q2 deltas, and the explicit missing q3/q8 deltas;
* 12 primary budget groups and 34 budget-fence rows, including exact metric
  vectors and zero fence deltas; and
* the 17 recorded negative controls, bound to the SHA-256 of the current
  `analyze.py`.

The optional diagnostic route trace is intentionally omitted from this small
source-only pilot. Its absence is accepted by the audit; no trace-based gate
is invented. Code structure, allocator evidence, counted-source evidence,
semantic outcomes, and the frozen matrix provide the bounded checks required
for this candidate.

The ordinary terminal command is:

```text
python3 docs/performance/results/change-0726/audit.py
```

It is fail-closed. If a captured candidate fails a timing, benefit,
allocator, semantic, or budget-fence gate, run the explicit reporting mode
while the candidate source and binary custody are still available:

```text
python3 docs/performance/results/change-0726/audit.py --allow-rejected
```

That mode still checks all bindings and raw custody, emits the recomputed
rows and hard-gate booleans with a `REJECTED` label, and exits 1. It cannot
turn a failed candidate into a pass or bypass malformed or missing evidence.
Save its stdout as the candidate-bound terminal audit record. After restoring
the baseline source, a normal candidate audit should reject because the
candidate source delta is no longer installed; any restoration check is a
separate custody record and must not be relabeled as retained candidate
evidence.

Before capture, `--draft` runs the source/build/fixture/tooling preflight and
reports that freeze and captures are pending. A terminal PASS requires every
configured hard gate and every binding check to pass. The audit recomputes
captured formulas and does not resample, reinterpret, or soften a frozen
gate.

## Initial terminal audit

The independent terminal command was run with the candidate source and owned
binaries still installed:

```text
python3 docs/performance/results/change-0726/audit.py --allow-rejected
```

It returned exit 1 and wrote `audit-initial.log`. The log is 1,576 bytes with
SHA-256
`aa201f2ca94512428918f06e5a93377f5c85ac7f5593a8ab656feb7bb2f974ca`.
The result is `REJECTED`, preserving the ordinary retention decision. All
custody and non-timing gates passed: the five complete manifests covered 144
native raw files, 864 repeat files, 576 allocator files, 24 primary-budget
files, and 68 budget-fence files; all raw hashes, commands, source/tool/probe/
corpus/case bindings, and binary/build receipts matched. The 17 negative
controls passed and remained bound to analyzer SHA
`6b2f605b4f4fd84b1434eacff4d66656cd769dff8ff5602628f75cea791f0923`.

The independent recomputation agrees with the primary report on every central
statistical and semantic result. All 672 native paired p50/mean percentage
cells and all 64 repeat percentage cells agree to floating-point rounding
(maximum absolute difference below `4e-14`); native/repeat/allocator/budget
gate booleans and every allocator/fence delta had zero row mismatches. The
native timing gate rejects four groups, each at one B/A pair's mean while its
p50 remains within the gate:

* `54016-stored-2097152` owned, `q3`, B2/A2 mean `+8.679%`;
* `54016-stored-2097152` file, `q3`, B1/A1 mean `+5.587%`;
* `synthetic-70000-default` owned, `q3-to-q8-mean`, B2/A2 mean `+5.876%`;
* `45365-late` file, `q8`, B2/A2 mean `+8.399%`.

Thus the final independent hard-gate vector is
`bindings=true`, `native=false`, `repeat=true`, `native_benefit=true`,
`repeat_benefit=true`, `allocator=true`, `budget_primary=true`,
`budget_fence=true`, and `trace=true` because the optional trace is absent.
The native benefit target still passes: owned
`54016-missing-1048576` q8 improves by 25.00%/27.68% in B1/A1 and
25.00%/26.03% in B2/A2 (p50/mean). Repeat improves by
32.43%/32.51% and 33.12%/33.09% in the same order. Benefits do not override
the four failed native timing groups.

All 96 allocator groups pass. The eight positive-budget missing q3/q8 rows
have exactly the planned six-field negative delta, while q2 and every other
strict, zero-budget, and refusal route have exact zero deltas. All 12 primary
budget comparisons retain semantic parity and the 34 budget-fence rows retain
both semantic parity and zero deltas for every counted-source metric.

## Current status

The 0726 draft preflight and terminal custody audit pass, while the candidate
is retained as rejected because four native timing groups fail. The ordinary
terminal command remains fail-closed and would exit 1 for this candidate. The
explicit `--allow-rejected` report records custody and independent math; it
does not change retention policy. Root will handle cleanup and exact baseline
restoration separately.
