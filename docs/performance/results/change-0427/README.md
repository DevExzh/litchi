# 0427: explicit PPTX allocator retention boundaries

The prior lifecycle brackets stop while caller objects are still alive. This
batch adds an explicit allocator-only diagnostic mode to observe input
preparation, opening, planning, publication and ordered drops of the returned
result, plan, document handles, caller source owners and output sink. It reuses
the existing plain and media-rich PPTX lifecycle corpora and their correctness
oracles. Existing timing selectors and production code are unchanged.

The [protocol](protocol.json) defines eight fresh processes, two repeats of each
owned/source-backed and plain/media-rich combination, 30 retained samples and
three warmups. All task builds, tests and captures are serialized by root with
Rust 1.98.1, four build jobs, one worker and CPU 2 for capture. Source-only agent
work can overlap. The shared host's background activity is uncontrolled.

These are absolute process-global callback observations and one whole-cycle
region peak. Input clones and output reservation are inside this cycle, unlike
the established timing brackets. Fixed scalar checkpoint storage avoids report
allocations between boundaries. Exact output bytes are checked before the sink
is dropped; corpus validation and report generation occur outside the cycle.
Explicit lifetimes help explain observed differences but do not turn global
counters into object-owned bytes. Whole-process RSS and command elapsed time,
if retained separately, are context rather than phase-local observations.

No optimization, latency, physical-copy, cache-occupancy, managed-budget or
RSS-release claim is made. Near-limit cache/budget tests, native and cold/range
sources, broader CRUD coverage, streaming and scaling remain open. The full
non-iWork goal remains active.

`check.py` retains immutable command receipts, raw logs and deduplicated source
hash manifests. `capture.py` refuses existing output, verifies the frozen build
and protocol bindings, and runs the declared matrix serially. Report replay and
mutation probes are independent Python programs; no original executable is
needed to validate retained observations after task-temporary cleanup.

## Captured observations

The release capture is bound to harness revision
`392a11a1e7fca51ec930adff024b0f7a7f20d4fd` and executable SHA-256
`e48043229d70195d3394631ccbc796f6155747dc070f4ea29963a6d0991a25a1`.
Eight fresh processes retain 240 samples, plus 24 validated warmups. Exact
output equality succeeds in all 264 iterations. The [table](result-table.md),
[raw-derived summary](summary.json), and [resource review](resource-review.md)
retain every boundary and distinguish point values from the whole-probe peak.

All 240 final sink-drop points equal their own entry live-byte value. Phase
changes are identical within each process and across its two repeats, so no
repeat row crosses the 5% review threshold. The media-rich publication delta
is 160.719714 MiB for owned and 118.309136 MiB for source-backed; these have
different live owner sets and are descriptive API observations. At document
drop the source-backed caller still holds the input sources and sink. Dropping
its caller source Arcs changes callback live bytes by −33,631,741 B, followed
by the sink drop returning the delta to zero. This does not prove RSS release,
cache eviction, absence of leaks, or a matched optimization improvement.

`verify.py --portable` replays reports, source/build custody, summaries,
compression and the complete inventory using only an exported bundle. Each
full replay rejects 160 report mutations (20 per report), plus explicit
portable mutations of the pinned validator and a valid-shaped output digest
in one repeat. The four debug reports separately pass 80 final report probes.
The original executable is not needed for this replay.

## Pre-capture validation

The complete harness command retained 287 passes, one failure and one ignored
opt-in test. Its sole failure was the stale selector-count literal (427 versus
429), reproduced with the exact prior `lib.rs` and corrected in a focused
passing rerun. The resulting composite coverage is 288 passes and one ignored;
the failed full-command receipt remains visible. The six new retention unit
tests and five allocator-binary tests are included in that coverage. No full
suite rerun is inferred from the focused assertion correction.

Four actual debug CLI reports pass the independent verifier. CLI probes reject
missing/duplicate/out-of-bounds arguments and normal-binary dispatch, and an
existing output file remains byte-identical. Final report replay rejects 20
mutations per report, including distinct callback/ownership scope changes.
These debug reports are validation artifacts, separate from release capture.

Warning-denied documentation passes. Strict Clippy retains exactly the prior
29 diagnostic-message/source-file findings, with none in the retention module.
Formatting initially found the new module declaration out of order; pinned
rustfmt corrected it and the final format check passes. These failures and
corrections are retained without warning suppressions or production changes.

## Reproduction

The build receipt and per-process capture receipts retain exact commands.
`prepare-build.py` copies the release executable into a new task directory;
`capture.py` refuses existing outputs and runs the declared matrix.
To replay the committed evidence without rebuilding:

```sh
python3 -B docs/performance/results/change-0427/verify.py --portable
```

The final bundle retains 20 command receipts and 114 inventoried files plus
`SHA256SUMS`. All 36 logs are losslessly compressed with original and stored
hashes. Post-cleanup portable replay passes after removal of the hash-bound
copied executable; the root and harness build directories remain intact.
