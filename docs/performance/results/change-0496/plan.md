# 0496: attribute the DOCX opened-edit lifecycle

0495 retained eight whole-child RSS flags and three latency comparison cells
above the five-percent review threshold. Their causes remain unresolved.
Whole-child profiles include expensive output verification, so they cannot
attribute the timed edit's cost. This follow-up instruments the existing
lifecycle without changing production code or removing any preservation oracle.

## Hypothesis and decision

The phase breakdown can distinguish open, edit staging, commit, diagnostics,
publication, and drop costs. It may reveal a stable production-phase delta or
show that the historical observation does not repeat under this capture. The
latter does not prove the earlier flags were noise or resolve their historical
cause. No flag is silently deleted. Any optimization requires a further
mechanism-specific experiment and correctness proof.

The diagnostic flag is opt-in. Existing default serialized reports retain their
schema. Phase elapsed values are wall time measured with `Instant`;
they include timing/checkpoint overhead and are not CPU-cycle measurements.
There is no nested per-phase allocator region because the allocator observer is
non-reentrant. Full-lifecycle allocator measurements and whole-child RSS remain
separate. Package publication consumes the package, so this harness cannot
invent cache-after-publication or cache-after-drop observations.

## Frozen comparison scope

Use the same diagnostic harness source for:

- before production: `de8ee88b0727ae59e4d2b4c8b8a6c24349724ae8`;
- after production: `44a4710699ef17041d5969240c30984dffbc3319`.

The before checkout receives the 0495 harness overlay plus the diagnostic
change; the after checkout receives only the diagnostic change. Each executable
is bound to its source manifest, patch, lockfile, build command, environment,
and binary hash. Reuse the 0495 corpus and exact output identity, which the
canonical report validator checks. No source or performance file in the sealed
0495 bundles is changed.

Two reversed repeats cover before/after unmanaged owned, file-warm, and short
providers in normal and allocator roles (24 processes), plus after-managed
file-warm and short in both roles (8 processes). Each formal process uses three
warmups and 30 measured samples: 32 processes and 960 formal samples. CPU 2 and
the existing shared measurement lock serialize captures. Record commands,
selected environment, host identity, raw reports, process start/terminal state,
and timeout failures. Reject drift, failed or missing oracles, nonconserving
phase intervals, and inconsistent source or sink counters.

Report individual phase and full-lifecycle p50/p95/p99, paired median bootstrap
intervals, full-lifecycle allocation counters, and whole-child RSS. Flag adverse
changes above five percent without aggregating them away. Warm filesystem and
short-read adapters remain explicit provider observations; this adds no cold,
native-producer, borrowed-lifetime, real-network, concurrent-scaling, or atomic
save claim. The full non-iWork performance goal remains open.

## Validation and retention

Run focused diagnostic parser/schema/oracle tests, the existing harness library
tests, warning-denied Clippy, rustdoc, formatting, boundary checks, and Python
capture-validator tests. Preserve failed developmental attempts separately.
Retain reproducible source/build custody, binaries, raw captures, analysis, and
cleanup evidence; remove the two disposable source checkouts and shared Cargo target
once terminal state and custody are verified. Preserve unrelated primary-tree
changes and never access `~/code/litchi-spec-gaps`.
