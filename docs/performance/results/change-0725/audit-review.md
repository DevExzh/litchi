# Independent audit review: 0725 XLS worksheet replay checkpoint

This is the independent source and evidence review for the bounded
`0725-xls-worksheet-replay-checkpoint` pilot. The immutable comparison point is
`ee5e0b0650` (`perf(xls): isolate checkpoint layout and replay costs`). This
review owns `audit.py`; it does not alter production sources, the pilot driver,
the primary analyzer, or earlier evidence packets. It does not run Cargo,
native probes, profilers, or instrumented builds.

## Pre-capture source review

The candidate source census differs from the baseline at exactly these three
paths:

* `crates/litchi-xls/src/workbook/query_cache.rs`;
* `crates/litchi-xls/src/workbook/source.rs`;
* `crates/litchi-xls/tests/xls_query_chain_checkpoint.rs`.

The audit requires the archived candidate source bytes and
`candidate-source.json` to agree with the live candidate for all three paths.
It also binds every successful five-binary build receipt to its phase source
census; package metadata and documentation files are bound separately by the
pilot source census.
With the prepared candidate currently installed, this source audit passes the
exact three-file delta, archive hashes, 224-to-264 charge, retained-type check,
EOF finished-hook check, empty-slot fences, and focused test custody. The
terminal audit below binds the same candidate source to the completed capture
manifests.

The retained index fixed charge is 224 bytes at the baseline and 264 bytes in
the candidate, pricing the new state at 40 logical bytes before slot
collection. The published index carries at most one
`Option<(u64, StreamChainCheckpoint)>`; the candidate transfers it on
publication and clears it on both abandonment and drop. The independent static
check rejects `CellValue`, `SourceBackedError`, and `String` in
`query_cache.rs` code, while allowing the pre-existing one-byte hotness table.
The source scan records the first raw target-frame offset before value
conversion. After a successful EOF `finish_scan` fence, the target sink uses
the already-borrowed workbook path to attempt one metadata-only CFB cursor
probe; it does not build a second path vector or read source bytes. Any cursor
refusal leaves the worksheet-start checkpoint in force. An empty indexed slot
set keeps the pre-existing worksheet lookup, final execution check, and
`ensure_current` version fence before returning the ordinary missing result;
it does not enter the path-vector replay setup. Replay selects the
target checkpoint only when the first matching frame is at or after its saved
offset. The focused integration test calls the public `SourceBackedWorkbook`
path, uses a real `ReadAt` fault wrapper, and includes a cold scan control that
arms the same later-sector fault. The existing `xls_query_index_cache.rs`
integration test independently exercises the indexed-missing trailing
freshness fence and asserts zero source reads after the injected mutation; the
audit checks that test rather than requiring a duplicate helper in the new
focused test.

The allocator plan prices the successful positive-budget `q2` path at exactly
40 bytes with zero additional allocation/deallocation calls. For a positive
budget missing target, `q3` and `q8` must match the explicit exact delta
`allocation_calls=-1`, `allocated_bytes=-16`, `deallocation_calls=-1`,
`deallocated_bytes=-16`, `peak_live_delta=-16`, and
`retained_live_delta=0`. All other strict routes, zero-budget routes, and
refusal routes require exact zero deltas. The budget fence includes the old
224-byte and new 264-byte boundaries and the intermediate refusal boundaries.

The independent tooling receipt contains 14 synthetic controls, including the
40/41-byte allowance boundary and accepted/rejected missing-warm allocation
mutations, bound to the current analyzer hash. These are verifier evidence,
not 0725 measurements.

The independent plan check binds the native timing exceptions exactly to
`q3`, `q8`, and `q3-to-q8-mean` at 10 ns, with the 5% limit retained for every
other native metric. Repeat q8 uses the separate `repeat-q8` metric name and
therefore remains under the strict 5% timing gate. The positive-benefit gates
require owned `54016-late` q8 improvement of at least 10% at both p50 and mean
in both B1/A1 and B2/A2 pairs for native and repeat evidence.

## Diagnostic route review

The route analyzer must keep the missing-target refusal distinct from selected
target replay. `54016-missing-1048576` has no indexed slot at the queried
coordinate, so its baseline replay-link vector is empty. The empty-slot branch
returns after the worksheet lookup and final fences; it does not fall through
to a worksheet walk. The exact semantic outcome and source-I/O checks remain
the evidence for this refusal path.

The separate sequence probe deliberately covers a different route: it checks
the earlier query's preserved replay-link prefix and the later origin query's
zero-link checkpoint replay by semantic labels. It also requires the candidate
checkpoint-build route and exact equality of the metadata-only I/O report.
Those checks are compatible with the fixed-target matrix assertion and do not
soften the missing-target refusal gate.

The baseline and candidate sequence probes compiled and completed; their
restored-source manifests, raw traces, and semantic reports are included in
the terminal custody review below.

The completed route comparison confirms the intended replay boundaries. On
the late-target cases, the candidate emits one `worksheet-checkpoint-build`
after the cold worksheet scan, then its replay starts at the saved target
offset with zero additional FAT links; the baseline replays from the
worksheet-start checkpoint. The sequence comparison reports
`candidate_checkpoint_build_links=[1044]`, preserves the earlier replay prefix
at 14 links for both phases, reduces the origin-late replay from 1,044 links to
zero, and keeps semantic and I/O reports equal. The missing sequence query has
zero source reads and two version observations in both phases.

The archived source explains why duplicate ordering remains intact. The sink
sets `first_target_offset` with `get_or_insert` before value conversion, while
publication retains every matching `CellSlot` sorted by stream offset. Replay
walks all matching slots in that order and overwrites `found` for each decoded
value, so the checkpoint begins at the first duplicate and the returned value
still follows the existing last-occurrence rule. The existing duplicate-order
integration test remains part of the source custody.

At worksheet EOF, the scan first validates the empty EOF payload and pending
formula state, then runs `finish_scan` for the execution and source-current
fences. The optional `finished` hook uses the still-borrowed path only to
construct a CFB cursor from immutable directory/FAT metadata; it does not read
payload bytes or allocate a second path vector. A cursor refusal leaves the
worksheet-start checkpoint unchanged. Candidate publication checks execution
again before transfer, while indexed replay checks execution before each slot
and after the replay loop, then calls `ensure_current`; the empty-slot return
performs its own final execution and source-current checks.

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
  fields, the 40-byte positive-budget `q2` allowance, and the explicit missing
  warm `q3`/`q8` deltas;
* all 12 primary counted-I/O groups and all 34 budget-fence rows, with exact
  eight-query semantic parity and source-metric deltas; and
* optional diagnostic route artifacts when present, including restored source
  bytes, timing-free semantic reports, and the ordered sequence evidence.

The normal terminal command is:

```text
python3 docs/performance/results/change-0725/audit.py
```

The default command remains fail-closed. If a captured candidate is rejected
by a timing, allocator, semantic, or route gate, use
`python3 docs/performance/results/change-0725/audit.py --allow-rejected` to
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

## Initial terminal audit

The independent terminal command was run while the candidate source and all
owned binaries were still present:

```text
python3 docs/performance/results/change-0725/audit.py --allow-rejected
```

It returned exit 1 and wrote `audit-initial.log` with the `REJECTED` label. The
log is 1,609 bytes with SHA-256
`c25b40ab74f7cf2daf2a256a374d28df1416abdf3a4330def959abfa573b0c45`. The
independent result preserves the native timing rejection: native timing passed
7 of 24 groups and failed 17; all 24 groups retained exact semantic outcomes.
The repeat gate passed all 16 groups. Both owned `54016-late` benefit gates
passed at p50 and mean in both pairs: native improvements were 80.0%/79.85%
and 80.82%/80.51%, while repeat improvements were 83.11%/83.06% and
83.11%/83.01% (B1/A1 p50/mean, then B2/A2 p50/mean). These benefits do not
override the failed native hard gate.

Custody and non-timing gates independently passed. The audit verified the
frozen plan, source start/end maps, archived candidate bytes, probe and corpus
maps, tooling hashes, five baseline and five candidate build receipts, and all
capture binary hashes. The five complete manifests bound 144 native raw files,
864 repeat files, 576 allocator files, 24 primary budget files, and 68
budget-fence files; each raw path was present with its recorded SHA-256 and
each recorded command exited zero. No cleanup witness was needed because the
owned binaries remained live. The 14 negative controls passed and remained
bound to analyzer SHA
`4f013975dfb8ec0300f309fdc11a3d61a6ea43ff06be41304bb57c9800cbcc66`.

The independent recomputation found all 96 allocator groups valid, including
the eight positive-budget missing-warm q3/q8 rows with the exact six-field
negative deltas. The q2 allowance, 12 primary budget groups, 34 budget-fence
rows, and diagnostic trace all passed. The final hard-gate vector was
`bindings=true`, `allocator=true`, `budget_primary=true`, `budget_fence=true`,
`repeat=true`, `native_benefit=true`, `repeat_benefit=true`, `trace=true`, and
`native=false`. The ordinary command therefore remains fail-closed; the
rejected report records custody and math without changing retention policy.
