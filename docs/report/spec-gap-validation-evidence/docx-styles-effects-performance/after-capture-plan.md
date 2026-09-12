# Matched after-capture plan

Status: **planning only**. This document authorizes no build, smoke run, or
timing capture. It defines the evidence gates for a future comparison after
the staged-apply production change.

## Fixed before record

The retained before record is the profile committed with `23ca02f01`. Its
production source pin is
`8702fd4db8723acceb7deb51bcb40ff66604bf10`, its clean profile checkout is
`bf21536b35380005041917e7aee03a39804b2659`, and its raw receipts are in
`results/profile-clean-bf21536b3/`. The approved current-source 52-lane smoke
remains the prerequisite retained under `fa927a8a9de94891858a6bb3d44c21d5fa63697d`.
The before receipts, source bundle, generated fixtures, and retention
manifest are immutable inputs to the comparison.

## Required after identity

After the production change is committed, the after capture must use a new
clean descendant and record full hashes for:

1. the complete production source commit and transitive local Cargo source
   closure;
2. the fresh 52-lane correctness-smoke checkout, retention commit, exact
   receipt file set, and Git blobs;
3. the after profile checkout and its source bundle; and
4. the built binary, Cargo metadata, lockfiles, toolchain, target triple,
   linker, flags, allocator observer, host, and environment.

The production source pin and current-smoke prerequisite must change
transparently in the after provenance. The before profile receipts and source
manifest are never rewritten. The generated native and synthetic fixture
bytes, package/member hashes, seeds, XML event/depth counts, and fixture
manifest must remain byte-identical. A fixture or corpus change blocks a
matched comparison.

The harness diff is allowlisted to source identity labels and current-smoke
prerequisite updates needed to bind the after run. It must contain no
algorithm, fixture-generator, phase-boundary, allocator-accounting, verifier,
or matrix changes. Any other harness diff requires a new scaffold review and
invalidates this before/after comparison.

Cargo metadata contains absolute checkout paths, so raw metadata JSON from
different clean worktrees is not expected to be byte-identical across the two
runs. Compare a normalized representation of package identities, resolved
dependencies, features, targets, profiles, and source revisions across runs;
require identical lockfiles and equivalent resolved dependency closure,
toolchain, target, linker, and flags. Within each capture, the before and
after metadata receipts must remain byte-identical. A within-capture metadata
change or a cross-run normalized closure difference fails closed.

## Correctness gate

Run a fresh after-source 52-lane smoke before building the profile. It must
pass all semantic, opaque-member, inverse, source-readback, cap, refusal,
native-fixture, and no-output checks with the after source identity. The
profile must bind that exact smoke result tree by its committed file set and
blob hashes. A smoke failure, stale source label, changed fixture hash, or
unbound receipt stops the after capture.

## Matched profile matrix

Use the same 31 lane/scale rows, native/64 KiB/1 MiB classes, three fresh
processes per row, two warmups per process, and twenty measured samples per
process. This produces 1,860 measured samples and 186 warmups:

`31 rows × 3 processes × 20 measured samples = 1,860 measured samples`

`31 rows × 3 processes × 2 warmups = 186 warmups`

Use fresh disjoint external results and target paths for the after run. Keep
the same operation-clock setup exclusions, `publish_ns`/`apply_ns` nesting,
`serialize_ns`, reopen, inverse, allocator, and RSS definitions. Do not add
8 MiB, near-limit, or refusal timing rows during this comparison.

## Comparison and decision gates

Compare matched lane medians and process-level spread for elapsed time,
`capture_ns`, `publish_ns`, nested `apply_ns`, `serialize_ns`, reopen and
inverse phases, requested allocation bytes, live/peak allocation values, and
process maximum RSS. Nested `apply_ns` is contained by `publish_ns`; its value
must not be added to publish time. Allocator values and RSS remain separate
observations and do not establish a memory cap.

Record host load and sampled process-census evidence for both captures. Use
observational wording for competing-process checks; sampled censuses do not
prove continuous absence of interference. Differences in host, load, or
toolchain are limitations on interpretation and must be reported beside the
data.

Only when the after smoke, source/blob closure, fixture hashes, normalized
metadata, harness allowlist, phase schema, and receipt verification all pass
may the result be described as a matched before/after observation. Report
lane-specific deltas and uncertainty with the exact source and host limits.
Do not convert the result into a blanket speedup, asymptotic scaling, managed
memory-cap, or native Office acceptance claim.
