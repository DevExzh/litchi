# 0961 XLSX `apply_batch` capacity trial

This is a prospective, screening-only protocol.  It does not authorize a
build or a workload.  Capture starts only after root GO, the candidate source
freeze, the candidate build receipt, and the post-build source/binary binding
receipt have all passed.

The hypothesis is deliberately narrow: `MultiSourceEdit::apply_batch` may use
the iterator's bounded `size_hint` when reserving its temporary pending edit
vector.  `size_hint` is advisory.  The implementation must retain fallible
growth, the 256-edit limit, cancellation checks, duplicate/invalid request
ordering, and rollback.  It must not use an infallible allocation or treat a
hint as permission to bypass the limit.  The production allowlist entry is
`crates/litchi-xlsx/src/cell_values/source.rs`; the only additional allowed
path is the test-only admission coverage at
`crates/litchi-xlsx/src/cell_values/admission_tests.rs`.

## Custody and build

`tools/verify_0961_bindings_v1.py` has passed baseline mode and wrote
`../receipts/baseline-binding-v1.json`.  It proves that the fresh
`source-baseline` has 10,933 files whose 10,932 retained files and modes match
0957, and that the retained 0957 benchmark and allocator executables have
their recorded SHA-256 values.  The additional root `Cargo.lock` is recorded
as an overlay and is not part of the retained runtime closure.  Candidate
builds must use the exact `tools/perf-baseline/Cargo.lock` recorded by that
receipt; using the root lock would change dependency resolution.  The build
is one Cargo process with the inherited `CARGO_BUILD_JOBS=2`; benchmark
children remain strictly serial.

The existing 0957 executables are role A only.  They may be reused after this
gate.  Role B must be built from the frozen 0961 candidate by the root-owned
`build-tools/build_candidate_managed_v1.py`; no executable is valid until its
build receipt binds source, harness, lockfile, command, and SHA-256.

## Minimal screening schedule

Use the same three managed cases and oracle data as 0960:

* `medium-one-percent` (93 edits),
* `dense-sparse-one-percent` (178 edits), and
* `medium-one-edit` (one edit, the ordinary `set` path).

Run one fresh serial process per child, worker count 1, CPU mask 12–15,
fixed C/UTC environment, and the inherited ABBA/BAAB native blocks: three
blocks, 10 warmups and 60 samples per role.  This is 36 native children and
2,160 measured native rows.  Run the inherited allocator lane as three AB/BA
paired repeats per case, two warmups and 10 samples: 18 children and 180
measured allocator rows.  The full screening set is therefore 54 children,
2,340 measured rows, 396 warmups.  Keep the 900 s per-child deadline and stop
on the first terminal, oracle, binding, or resource failure.  Never retry or
silently replace a child; the plan performs no Cargo work while capturing.

The ordinary, fallback, and batch controls are correctness gates, not a new
timed matrix.  Use the candidate's test-only admission coverage for empty,
unknown or underreported hints, exact one-item, exact 8-item and 256-item
batches, 257-item refusal, duplicate or invalid late items, cancellation,
rollback, and error ordering.  The benchmark's existing output, semantic,
inverse, resource, and sink oracles then cover the three managed cases.  A
refused operation must retain its documented error class and leave no partial
staging; no separate unbuilt control harness is promised.

## Decision rule

Every child must have terminal evidence before its report is parsed.  Any
source or executable mismatch, malformed report, timeout, non-zero exit,
oracle mismatch, resource-counter mismatch, inverse/publication mismatch,
allocator conservation failure, or control failure rejects the packet.  There
is no retry.

Screening is eligible for full confirmation only when at least one named
native primary metric, native peak-RSS metric, or allocator region peak live
bytes has candidate/baseline median ratio `<= 0.95`; requested bytes or
allocation calls alone never constitute signal.  A native primary, native
peak-RSS, allocator region peak, or allocator peak-RSS ratio `> 1.05` is an
immediate safety rejection.  A screen with no eligible signal is
`reject_screen`, even if it looks directionally better.

If and only if signal and all safety/oracle gates pass, run a fresh confirmation
plan: six native ABBA/BAAB blocks with 20 warmups and 200 samples per child,
plus a fresh allocator confirmation with three AB/BA repeats, two warmups and
10 samples.  Confirmation requires a named primary ratio `<= 0.95` with the
bootstrap confidence interval upper bound `< 1.0`, no primary safety ratio
above `1.05`, and every control/oracle pass.  Otherwise reject.  A successful
screen is never adoption evidence by itself.

The expected serial wall time is feasible for this scope: budget roughly
45–90 minutes for screening on the inherited host, with the 900 s child
deadline as the hard ceiling; confirmation is conditional and adds roughly
2–4 hours.  These are planning budgets, not measured claims.  Do not start
either lane until root authorizes capture.

Before adoption, exact successful 8-item and 256-item MultiSourceEdit batches
and ordinary/fallback performance controls remain required. These are not
claimed as completed by the screening admission tests. Root separately
authorized the isolated release build after quality; capture still requires
a fresh root GO receipt.
