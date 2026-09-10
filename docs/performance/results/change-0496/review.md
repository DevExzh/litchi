# Independent review — change 0496

Review status at the current freeze boundary: **blocked for acceptance pending
analysis completeness and final capture receipts**.

The frozen Rust harness at
`tools/perf-baseline/src/docx_managed_edit.rs` has SHA-256
`e3732d932356018dff0bd7a17b8eb38e60e177ee7d83992e9a04dbe82fe26c94`. The
diagnostic change is opt-in and the source diff preserves the existing 0495
correctness path. The existing output byte and semantic/media checks, source
version fence, one-materialization and commit-operation checks, commit identity
check, preflight replay/inverse/stale/foreign patch oracles, read/sink/range
conservation, managed budget release checks, and allocator sample remain in
`run_sample`. The only lifecycle restructuring moves the existing commit drop
and budget evidence into the interval-producing branch; it does not remove an
oracle.

The phase instrumentation is semantically bounded. It uses wall-clock
`Instant` intervals for open, edit staging, commit, diagnostics/XML identity,
publication, published snapshot drop, and commit drop. The phase sum is checked
with checked arithmetic and the residual is checked against the outer lifecycle
latency. The residual explicitly covers phase-boundary arithmetic, budget
evidence, result handling, and unassigned package work. The source and driver
use the same scope string, including the statement that these are not CPU-time
measurements. `instrumentation_overhead_ns` is explicitly absent rather than
invented. The single full-lifecycle allocator region surrounds `execute_api`;
the phase clocks do not create nested allocator regions, consistent with the
non-reentrant allocator observer.

Default report compatibility is preserved by an optional, skipped
`phase_diagnostics` row field. The source tests cover both default omission and
enabled conservation. The Python driver requires the exact phase field set,
scope strings, nonnegative durations, phase-sum conservation, lifecycle
conservation, and absent instrumentation overhead. It removes only the explicit
row field in a temporary projection, rejects nested phase fields, and passes
that projection to the unchanged 0495 `validate_report`; retained reports are
not rewritten.

The current static custody review is also consistent. `builds.json` contains
the four expected normal/allocator records for the before and after revisions,
and the before and after source manifests contain identical hashes for all four
shared harness files (`docx_managed_edit.rs`, `lib.rs`, `main.rs`, and the
allocator binary entry point). The driver binds source manifests, build gates,
Cargo argv, selected build environment, binary metadata, machine/provenance
inputs, protocol, terminal receipts, and private cleanup roots. Terminal
validation is exact, requires an exited successful child, rechecks every
retained artifact hash, and rejects failed, timed-out, or incomplete receipts.
The formal matrix is 32 children and 960 measured samples: 12 before
unmanaged, 12 after unmanaged, and 8 after managed, with normal and allocator
roles, two reversed repeats, three warmups, and 30 measured samples per child.
The provider observations are owned, warm file, and short-range; they do not
support cold filesystem, native producer, network, borrowed-lifetime, or
concurrency claims.

The acceptance blocker is in `capture.py`, not the Rust instrumentation.
`_entry_metrics` currently reports latency, one whole-child GNU-time RSS
scalar, and phase vectors. `_paired_comparisons` compares latency, RSS, and
phase vectors only. It never extracts or summarizes the allocator row fields
that the retained reports carry and that the 0495 validator validates:
allocation, deallocation, and reallocation calls; allocated and deallocated
bytes; live bytes; and region peak/increment values. The 0496 plan explicitly
requires full-lifecycle allocation counters in the analysis, and the raw
allocator receipts alone do not satisfy that analysis contract. Add the
allocator vectors/percentiles and relevant repeat or paired visibility before
accepting the formal analysis, or explicitly revise the frozen scope and its
acceptance criteria.

At this review point the final frozen protocol, formal capture inventory,
analysis receipt, and verification receipt were not yet present, so terminal
custody and all 960 report-level validations remain pending root's capture and
verification run. Once those receipts exist, the review can be closed only if
the allocator analysis blocker is resolved and the protocol, source/binary
bindings, actual chronological order, cleanup receipts, phase conservation,
and unchanged 0495 oracle validation all pass.
