# 0911 v2 PPTX opened-snapshot Callgrind protocol review

This is an amendment review for the v2 capture protocol. The original
`protocol-review.md` and the separate scope clarification remain unchanged and
hash-bound. This file changes neither the retained binary nor production or
harness code; it records the correction required after the first v1 child was
shown to contain two positive owner dumps.

## v1 finding and the required scope correction

The v1 owner filter was:

```text
--collect-atstart=no
--toggle-collect=litchi_pptx::package::model::Package::opened_presentation_transaction
--zero-before=litchi_pptx::package::model::Package::opened_presentation_transaction
--dump-after=litchi_pptx::package::model::Package::opened_presentation_transaction
```

That filter correctly selected the transaction method, but the first child
proved that the owner is also called during corpus construction. The benchmark
entry point first calls `build_pptx_source_edit_corpus()`; that calls
`pptx_source_edit_bytes()` and reaches
`package.opened_presentation_transaction()` at
`tools/perf-baseline/src/lib.rs:16978` before the measured benchmark parent is
entered. The independent v1 raw output therefore contains a setup owner dump
and a measured owner dump. The v1 child is a retained failed-attribution
receipt; neither dump may be silently selected or combined as a result.

The source audit also finds an owner call in `pptx_cross_copy_bytes()` at line
16937, but that path is not reached by the opt-in `pptx_owned_benchmark`
example. The dynamic v1 dump is the authoritative evidence for the reachable
corpus-builder call. The post-operation semantic oracle uses `from_bytes` and
OPC parsing and does not call this owner.

## v2 command and reset boundary

Use the same retained 0910 executable, route schedule, CPU 12 affinity,
`warmup=0`, and `samples=1`. Do not rebuild or overwrite v1 artifacts. Keep
`--collect-atstart=no` and the exact transaction owner as the only collection
toggle, but reset and dump at the measured helper boundary:

```text
taskset --cpu-list 12 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --zero-before=litchi_perf_baseline::pptx_owned_benchmark::timed_iteration \
  --dump-after=litchi_perf_baseline::pptx_owned_benchmark::timed_iteration \
  --callgrind-out-file=<fresh-callgrind-path> \
  docs/performance/results/change-0910/binaries/native \
  <route> 0 1 <fresh-artifact-directory>
```

`zero-before=timed_iteration` resets the Ir counters immediately before the
single measured helper call. The setup owner invocation remains real process
execution and may remain visible in Valgrind's process-level stderr accounting,
but its collected counters are discarded by that reset. The owner toggle then
turns collection on when the transaction method is entered and off when it
returns. `dump-after=timed_iteration` writes the scoped result after the
parent returns. The owner is called once inside this parent for the one
measured sample, so the scoped dump must contain one measured owner interval.

This is a reset-scoped collection, not a whole-process instruction total. Do
not assert that a `Collected` number in Valgrind stderr equals the accepted
Callgrind dump summary: stderr may account for setup instructions collected
before the reset, while the accepted dump is the post-reset interval. Do not
add the v1 setup and measured summaries, and do not subtract any semantic
oracle estimate from v2.

## Static dynamic-call audit

The retained binary has one exact defined symbol for both
`litchi_pptx::package::model::Package::opened_presentation_transaction` and
`litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`. With
`run(route, 0, 1, directory)`, the route's loop executes one
`timed_iteration` call. Its body has one
`package.opened_presentation_transaction()` call at
`tools/perf-baseline/src/pptx_owned_benchmark.rs:294`. The corpus builder's
transitive call occurs before the reset boundary. The source/output semantic
oracle and digest checks occur after the parent returns and after owner
collection is toggled off.

The v2 boundary therefore admits the setup call as intentionally excluded
pre-reset work and the measured call as the only scoped owner call. This does
not claim that setup is absent from the process or from Valgrind stderr; it
claims only that the accepted Callgrind dump begins after the reset and is
rooted in the measured parent.

## First-child raw admission

Retain every v2 raw dump, command receipt, stdout, stderr, and semantic report.
Use a fail-closed raw reader rather than `callgrind_annotate` output. The first
v2 child is admissible only if all of these conditions hold:

1. The binary path, size, SHA-256, route, CPU, environment, Valgrind version,
   and exact v2 options match the frozen plan. The plan must record the v1
   failure and this amendment as inputs without rewriting either artifact.
2. There is exactly one positive dump triggered by
   `--dump-after=litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`
   after the reset, and its `Ir` summary and totals are positive and equal.
   A final program-termination dump, if emitted, has zero summary and totals.
3. In the positive scoped dump, the exact transaction owner has exactly one
   positive incoming edge from
   `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`, with one
   dynamic call. There is no positive owner edge from
   `pptx_source_edit_bytes`, corpus construction, output verification, report
   writing, or another caller. Any extra positive owner edge is an admission
   failure, not a row to discard.
4. The positive dump's cost is rooted in the selected owner interval. The
   owner self cost plus its immediate direct-child inclusive costs reconstruct
   the owner edge without adding nested descendants twice. All positive
   function costs in the accepted dump are retained and accounted for.
5. The raw header's command, event, trigger, and dump parts identify the
   reset-scoped v2 capture. Dump numbering is not used as the scope authority;
   the trigger and raw edge structure are.

Zero-cost call-graph context created by the process may be retained, but it is
not a second measured owner interval. The gate concerns positive Ir and exact
owner edges. A dump with no positive measured owner edge, a nonzero terminal
dump, an unknown positive owner caller, or a summary that cannot be
reconstructed is rejected.

## Independent semantic gates and interpretation

The one-result JSON report must independently pass the frozen corpus, route,
source/output identity, semantic output, sink bound, source immutability,
`warmup=0`, and `samples=1` checks. These checks run after the measured parent
and are outside the Callgrind owner interval. Their successful output does not
repair an ambiguous profile.

The selected owner still includes opened-presentation catalog capture, slide
and notes validation, MCE work, package cloning, transaction construction, and
the internal complete-package fingerprint and payload/Part digest work
described by `protocol-clarification.md`. It excludes the parent code before
and after the owner, including ingress, edit/commit, apply, serialization,
sink writes, post-timer output SHA-256, semantic/Part-preservation checks, and
artifact writes. Debug information is disabled, so snapshot versus apply must
be described through the owner and parent function boundaries and raw function
edges, never invented source-line attribution.

If the v2 first-child gates pass, run the remaining three ABBA legs with the
same reset protocol, requiring the same one scoped owner edge and zero terminal
dump in every child. If any leg produces an extra positive owner context,
retain it and stop the profile lane. Even a fully admitted four-leg result is
guest `Ir` evidence for one generated in-memory fixture; it is not native
latency, allocation, RSS, physical I/O, or a production speedup claim.
