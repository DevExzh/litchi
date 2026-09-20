# Independent audit review: 0723 XLS worksheet replay checkpoint

This is the independent source and evidence review for the bounded
`0723-xls-worksheet-replay-checkpoint` pilot. The immutable comparison point is
`45cb480eaa` (`perf(docx): retain writer-local structural scan fusion`). This
review owns `audit.py`; it does not alter production sources, the pilot driver,
the primary analyzer, or earlier evidence packets. It does not run Cargo,
native probes, profilers, or instrumented builds.

## Pre-capture source review

The candidate source census differs from the baseline at exactly these three
paths:

* `crates/litchi-xls/src/workbook/query_cache.rs`;
* `crates/litchi-xls/src/workbook/source.rs`;
* `crates/litchi-xls/tests/xls_query_chain_checkpoint.rs`.

The archived candidate source bytes and `candidate-source.json` agree with the
live candidate for all three paths. The five baseline build receipts contain
491 Rust source entries and the five candidate receipts contain 492; every
receipt is successful and the only Rust-source differences are the two
production files plus the focused test. Package metadata and documentation are
bound separately by the pilot source census.

The retained index fixed charge is 224 bytes at the baseline and 264 bytes in
the candidate, pricing the new state at 40 logical bytes before slot
collection. The published index carries at most one
`Option<(u64, StreamChainCheckpoint)>`; the candidate transfers it on
publication and clears it on both abandonment and drop. The independent static
check rejects `CellValue`, `SourceBackedError`, and `String` in
`query_cache.rs` code, while allowing the pre-existing one-byte hotness table.
The source helper obtains the target checkpoint with fallible metadata-only
CFB cursor construction, has no direct `read_at` call, and returns to ordinary
replay when reservation or cursor construction refuses. Capture is gated on a
successful selected target, so missing/error/refused candidates retain the
worksheet-start checkpoint. Replay selects the target checkpoint only when the
first matching frame is at or after its saved offset and otherwise keeps the
worksheet-start fallback. The focused integration test calls the public
`SourceBackedWorkbook` path, uses a real `ReadAt` fault wrapper, and includes a
cold scan control that arms the same later-sector fault.

The allocator plan prices the successful positive-budget `q2` path at one
64-byte allowance: the 40-byte logical checkpoint plus at most 16 bytes of
temporary path-vector scratch. `q1`, `q3`, `q8`, zero-budget, and refusal paths
remain exact allocator gates. The budget fence includes the old 224-byte and
new 264-byte boundaries and the intermediate refusal boundaries.

The independent plan check binds the native timing exceptions exactly to
`q3`, `q8`, and `q3-to-q8-mean` at 10 ns, with the 5% limit retained for every
other native metric. Repeat q8 uses the separate `repeat-q8` metric name and
therefore remains under the strict 5% timing gate. The positive-benefit gates
require owned `54016-late` q8 improvement of at least 10% at both p50 and mean
in both B1/A1 and B2/A2 pairs for native and repeat evidence.

## Diagnostic route review

An earlier draft raised a concern about the replay-link assertion in
`trace-analyze.py`. The concrete missing-target case does not enter that
assertion: `54016-missing-1048576` has no indexed slot at the queried
coordinate, so its baseline `old_replay_links` vector is empty. The analyzer's
empty-link branch requires no candidate checkpoint-build links. The source
replay loop likewise iterates an empty `slots_for(row, column)` result, so the
candidate does not fall through to a worksheet walk. The exact semantic
outcome and source-I/O checks remain the evidence for this refusal path.

The separate sequence probe deliberately covers a different route: it checks
the earlier query's preserved replay-link prefix and the later origin query's
zero-link checkpoint replay by semantic labels. It also requires the candidate
checkpoint-build route and exact equality of the metadata-only I/O report.
Those checks are compatible with the fixed-target matrix assertion and do not
soften the missing-target refusal gate.

The first diagnostic attempt exposed a sequence-probe compile defect:
`query()` tried to call `.value()` on the `CellValue` returned by the public
worksheet API. The corrected probe must compile and complete both baseline and
candidate traces before those artifacts can be considered evidence.

The primary analyzer must also count one budget-fence row per baseline/candidate
pair. The plan contains 34 budget observations, so a correct independent gate
expects 34 comparison rows, not 68. Build-manifest verification must use the
path recorded in each frozen manifest (`baseline-builds.json` and
`candidate-builds.json` currently live at packet level), and source start/end,
tool start/end, probe, corpus, fixture, and binary hashes must be asserted at
both freeze and capture completion.

## Terminal audit boundary

`audit.py` independently revalidates every capture manifest and raw hash,
requires the frozen candidate source at capture start and end, checks both
baseline and candidate build receipts, and accepts a removed binary only when
an exact `cleanup.json` path/hash witness matches the manifest. It recomputes:

* all 24 native groups, every query outcome across all 100 samples and six
  legs, and both p50/mean B1/A1 and B2/A2 timing gates;
* all 16 repeated-query groups, exact found counts, and both normalized p50/mean
  q8 timing gates;
* all 96 allocator groups, repeat stability, metadata equality, exact strict
  fields, and the bounded positive-budget `q2` allowance;
* all 12 primary counted-I/O groups and all 34 budget-fence rows, with exact
  eight-query semantic parity and source-metric deltas; and
* optional diagnostic route artifacts when present, including restored source
  bytes, timing-free semantic reports, and the ordered sequence evidence.

The normal terminal command is:

```text
python3 docs/performance/results/change-0723/audit.py
```

The default command remains fail-closed. If a captured candidate is rejected
by a timing, allocator, semantic, or route gate, use
`python3 docs/performance/results/change-0723/audit.py --allow-rejected` to
finish the independent custody and gate recomputation. It emits `REJECTED`
with the computed rows and hard-gate booleans, and still exits 1; the option
cannot turn a rejected candidate into a pass or bypass malformed/missing
evidence checks. `--report-rejected` is an equivalent spelling. Save that
stdout as the terminal audit record when the primary analyzer also rejects.
Run this while the candidate source and owned binaries are still available (or
while their exact cleanup witness is present). After the candidate checkout is
restored to baseline, the ordinary source-delta audit must reject by design;
retain this candidate-bound report and treat any post-restore check as a
separate restoration/custody audit.

Before capture, use `--draft` after the diagnostic trace restores the
candidate source. A terminal PASS requires every hard gate and every binding
check to pass. The audit does not resample, reinterpret, or soften a captured
gate.
