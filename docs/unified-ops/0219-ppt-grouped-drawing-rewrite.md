# 0219 — Grouped native PPT drawing rewrites

Record 0218 measured near-native facade overhead at 1,000 operations but more
than 120× time growth from 1,000 to 10,000 operations. Native text batches
repeatedly inspected and rebuilt an entire drawing for each selected shape.

Batch text admission now indexes requested native shape containers by persisted
slide. Text/encoding/dependency admission remains in request order, and only
selected textboxes are interpreted. Publication groups accepted replacements by
slide and traverses each affected drawing once, retaining each exact-before
payload check. Source shapes and records remain authoritative; the facade does
not acquire a separate presentation model.

The 256-item public chunk bound, 4,096-operation primary edit bound, source
fences, unrelated-stream validation and complete candidate reopen/readback
remain. The existing named harness routes retain their calls and oracle; new
native/unified eager/positional routes at 256 and 4,096 operations expose those
boundaries.

## Verification

The full native all-feature suite passes 1,411 tests and doctests. The facade
all-feature suite passes 1,603 tests/doctests after the initial grouped rewrite;
after the final native error-order fixes, 32 focused facade coalescing, write and
structure tests pass. The final 14 focused native batch tests also pass.
Coverage includes exact 256-target publication across slides, out-of-order
chunks, repeated targets across chunks, per-chunk duplicate refusal and retry,
no-ops, unrelated stream/shape preservation, exact patch replay/inverse, grouped
versus sequential output bytes, and malformed OfficeArt framing/error order.

The grouped index tracks active target identities while traversing nested
OfficeArt containers. A malformed index takes the former single-target
inspection path in caller order, preserving both framing refusal and the
first typed text error. A later positional resolution failure is likewise
reported only after the successful prefix completes text/encoding admission.

Native/facade all-feature all-target Clippy and warning-denied rustdoc pass.
The native glossary profile, 66-package crate-boundary check, generated
OfficeCLI parity check and 15 parity-tool tests pass. Both repository lockfiles
are unchanged. Validation uses an external root lock with cached cc 1.6.0;
the harness uses an external exact copy of its own lockfile.

The three focused harness tests and harness all-target warning-denied Clippy
pass. All 28 named routes pass exact output/source and expected commit-count
checks. The twenty pre-existing routes retain exactly the source and output
SHA-256 values from record 0218; this optimization changes no publication bytes.

## Paired measurements

Runtime source: `27c3cb1e80`; grouped implementation begins at `758e81d1df`.
[Samples, source/output hashes and build identity](results/0219-ppt-grouped-batch.json)
record one warmup and five samples per leg, native then unified, for eager and
positional owners. No builds or tests overlapped timing. Runtime sources were
committed; only documentation was uncommitted. Release LTO was disabled.
Fixture generation, opening, command-vector construction, serialization and
verification are outside timing; unified string-path parsing and patch
destruction are included. The 10k case explicitly uses the finite 1-GiB snapshot
profile. The other cases use the default 256-MiB profile.

| Operations | Eager native ms | Eager unified ms | Ratio | Positional native ms | Positional unified ms | Ratio |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 0.110 | 0.122 | 1.103 | 0.117 | 0.128 | 1.097 |
| 10 | 0.146 | 0.194 | 1.321 | 0.150 | 0.215 | 1.431 |
| 100 | 0.546 | 0.901 | 1.650 | 0.550 | 1.040 | 1.892 |
| 256 | 1.255 | 2.145 | 1.709 | 1.251 | 2.483 | 1.986 |
| 1,000 | 14.358 | 17.926 | 1.248 | 13.692 | 18.642 | 1.362 |
| 4,096 | 267.629 | 289.798 | 1.083 | 256.373 | 289.768 | 1.130 |
| 10,000 | 4439.748 | 4573.782 | 1.030 | 4251.434 | 4661.923 | 1.097 |

At 1k, native medians improve 13.48× eager / 14.03× positional and unified
medians improve 11.01× / 10.54× versus record 0218. At 10k, native improves
5.45× / 5.55× and unified improves 5.25× / 5.22×. The default-profile 1k
workload remains within the 2× relative budget, with facade overhead of 24.8%
eager / 36.2% positional. The faster leaf makes facade path/admission overhead
more visible: the 100–256-operation ratios are 1.65–1.99. The one-operation
eager ratio is 1.103, narrowly above the 10% review trigger; these local samples
do not support a universal 10% claim.

**Linear scaling still fails.** Native 1k→10k growth is 309.21× eager / 310.51×
positional; unified growth is 255.15× / 250.08×. The ratio grew because the
1k case improved more than 10k, even though all large-case absolute times
improved substantially. Publication remains one per 256-target chunk: four
calls at 1k, sixteen at 4096 and forty at 10k. Native primary commits remain
one through 4096 and three at 10k; the facade edit is atomic across all of them.

These are synthetic same-slide text closure edits on warm prepared owners.
Blocked leg ordering can retain drift. Cold opens, command JSON conversion,
geometry, allocation qualification and broad-corpus acceptance remain outside
this evidence.

The two owned agent worktrees/branches and this turn's temporary targets,
locks, logs, probe binary and raw result copies were removed after evidence
reduction. No compiled artifacts are committed.


## Remaining work

This change does not aggregate publication across native chunks. Every chunk
still opens and publishes a native editor and retains append-only history.
The next native step is one bounded publication per primary edit profile,
including cumulative admission, source fences, exact-before checks and complete
readback. Actual scaling acceptance depends on measurements, not only on the
absence of per-target drawing rewrites.

Blank slide insertion, top-level shape insertion/removal, broader CRUD, delayed
formats, Python/CLI coverage and the full unified-operations acceptance goal
remain required. Record 0218 contains the native insertion plan.
