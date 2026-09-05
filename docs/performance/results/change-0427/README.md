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
