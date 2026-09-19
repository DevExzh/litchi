# Change 0692 — capture-local PPTX root/name projection

This packet compares frozen baseline `e9bd3360c` and the private capture-only
projection change. `performance_claim: none`: measurements are scoped to this
host, corpus and native public-call phases; no registered claim or coverage
promotion is made. The broader GOAL remains active; iWork is excluded.

The native probe covers five one-edit, five exact semantic no-op, and three
two-slide edit cases. Source bytes are materialized and packages opened outside
timers. Capture, working clone, set text, commit and apply are timed; total
includes interphase timer/check overhead. Save, reopen and preservation checks
are outside timers. Every process validates revisions, slide/shape projection,
sorted part inventory/content types/relationships/non-part members and untouched part
payload hashes. The oracle sorts metadata and does not assert its original ordering. The no-op
requires unchanged before/after serialization and revision, not equality to
the original source archive;
two-slide cases validate both markers. Edited-part unknown markup and physical
ZIP metadata remain covered by library tests rather than this bounded oracle.

`build.py baseline` freezes separate native and counting-allocator binaries;
`measure.py baseline` records two A/A legs before edits. After implementation,
`build.py candidate` freezes the counterparts and `measure.py compare` records
A/B/B/A legs. Each of 78 processes has five warmups and 100 fresh samples.
Odd legs reverse case order. CPU 12 is pinned on a shared Linux host; warm
OS caches are assumed, and host-wide quiescence is not established. Raw legs,
means, nearest-rank tails and descriptive seeded bootstrap intervals are
retained. Within-process intervals do not establish population tail bounds.

`measure-allocations.py baseline|candidate` runs three samples per case in
separate instrumented binaries. Request bytes count the full successful new
reallocation size. `profile.py baseline|candidate` retains native sampling,
hardware counters and whole-child RSS for open-plus-capture prefixes, not the
isolated phase timers. Raw perf data lives only in the owned scratch directory.

The marker control changes 43 exact MCE namespace URIs to an inert spelling.
All 103 ZIP members retain names, timestamps and uncompressed lengths;
compression is regenerated and flags/external attributes change. This alters
semantics and is only a mechanism control, never an original-deck oracle.
`prepare-control.py` and `check-control.py` reproduce and disclose the changes.

The candidate trace (`trace.py --profile candidate --model-trace ...`) is a
separate temporary instrumented build. It binds the frozen candidate source,
archives before/after bytes, and restores them in `finally`. No trace timings
are native evidence. Baseline MCE counts reuse 0691's final trace against
byte-identical production source; the probe's capture prefix is unchanged.
`trace-summary.py` verifies retained receipts after binaries are cleaned;
`--verify-live-binary` additionally checks the executable while it exists.

Reproduction uses locked release builds with warnings denied, two Cargo jobs,
absolute sibling scratch paths and CPU 12; adapt these machine choices
consistently elsewhere. For baseline reproduction, use the recorded revision
and this retained probe before applying the candidate. `run-integration.py`,
`run-evidence.py`, summarizers and `audit.py` retain the final verification.
The final facade test uses standard rustc warnings because narrow existing
feature combinations contain unused test helpers; failed warning-denied
attempts remain under `integration-initial`. Owner check/Clippy/tests and
rustdoc deny warnings.
No physical cold-cache, concurrent, remote-source, cross-platform, native
Office or complete open/edit/save workflow result is claimed.
