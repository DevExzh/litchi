# 0911 v3 PPTX opened-snapshot Callgrind protocol review

This is the bounded v3 protocol after the v1 and v2 first-child admission
failures. The original protocol, scope clarification, and v2 review remain
unchanged and are retained as history. V3 uses the original owner-scoped dump
shape, but admits and classifies both expected dynamic owner contexts instead
of relying on a reset at the measured parent boundary.

## V2 observation and decision

The v2 first child used the exact retained 0910 binary and passed the workload
gates, but its positive dump was not a measured-only interval. Its command used
`--zero-before=litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` and
`--dump-after` at the same parent. The raw dump was triggered after that
parent, had summary and totals **1,127,351,444 Ir**, and contained both
positive owner edges:

| dynamic context | owner Ir |
| --- | ---: |
| `litchi_perf_baseline::build_pptx_source_edit_corpus` | 126,674,905 |
| `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration` | 1,000,676,539 |

The two values conserve the raw summary exactly. The terminal dump is zero and
the semantic report passes, but the positive dump is rejected because it
contains setup and measured owner work together. This is an observation of the
captured raw file, not a claim about the implementation reason for the
reset-option behavior. V3 does not rely on that reset behavior.

The v2 raw file and receipts remain immutable evidence. Its process-level
stderr `Collected` value is also not a scope authority. No v2 total is reused
as a measured result.

## Frozen command and schedule

Use the retained 0910 binary at
`docs/performance/results/change-0910/binaries/native`, SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`, without
rebuilding or changing the source. Keep the four fresh serial ABBA legs,
CPU 12, `warmup=0`, `samples=1`, and new output and Callgrind paths. Preserve
all v1 and v2 artifacts; v3 must not overwrite them.

Use the original owner-scoped options:

```text
taskset --cpu-list 12 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --zero-before=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --dump-after=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --callgrind-out-file=<fresh-callgrind-path> \
  docs/performance/results/change-0910/binaries/native \
  <route> 0 1 <fresh-artifact-directory>
```

Each dynamic owner entry creates an owner-scoped positive dump. The expected
contexts are independent rows, not two samples of one phase. Dump numbering
and file order are not classification authorities.

## Expected dynamic contexts

At `warmup=0` and `samples=1`, exactly two reachable calls to the selected
owner are expected:

1. Setup: `litchi_perf_baseline::build_pptx_source_edit_corpus`, reached
   transitively through `pptx_source_edit_bytes` at `lib.rs:16978` while the
   fixed corpus is generated.
2. Measurement: `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`,
   at `pptx_owned_benchmark.rs:294`, which is the opened snapshot phase in the
   native benchmark timer.

The source also contains an owner call in `pptx_cross_copy_bytes` at
`lib.rs:16937`, but that route is not executed by this opt-in benchmark
example. The output semantic oracle and digest checks use `from_bytes` and OPC
inspection after `timed_iteration`; they are not expected owner callers.

The exact owner and measured-parent symbols must each remain unique in
`nm -C --defined-only`. A raw profile must classify every positive owner edge
by its demangled caller name and dynamic call count. It must not classify by
part number, dump order, largest Ir value, or source line because the binary
has debug information disabled.

## Raw admission gates

Use a fail-closed raw Callgrind reader and retain every profile part, command,
environment, stdout, stderr, and semantic report. Each child is admissible
only when all of the following hold:

1. The binary path, 9,301,456-byte size, binary hash, route, CPU, environment,
   Valgrind version, and four owner options match the frozen v3 plan. The
   source/build/lockfile custody remains bound to 0910.
2. There are exactly two positive owner-scoped dumps, one classified as the
   setup caller and one as the measured caller. Each owner edge has
   `calls=1`, a positive summary equal to its totals, and the exact owner
   function name. A zero program-termination dump must follow; it is not a
   lifecycle row.
3. The setup caller is exactly
   `litchi_perf_baseline::build_pptx_source_edit_corpus`, and the measured
   caller is exactly
   `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`. Every
   positive owner edge is accounted for. A duplicate expected caller, an
   unknown positive caller, a missing context, or an extra positive dump stops
   admission.
4. In each positive dump, the owner self Ir plus its immediate direct-child
   inclusive Ir reconstructs the owner edge. The sum of all function self
   costs reconstructs that dump summary. Nested descendants are retained as
   diagnostics and are never added again to a direct-child partition.
5. The positive dumps' headers identify the exact owner dump trigger and
   `events: Ir`; the terminal header identifies `Program termination` and has
   zero summary and totals. The parser must retain and account for all raw
   positive edges, including setup, rather than deleting setup records.

The v3 measured row is exclusively the dump whose owner caller is
`pptx_owned_benchmark::timed_iteration`. The setup row is a custody and scope
diagnostic. It must not be subtracted from the measured row, added to it,
averaged with it, or used as a route result.

## Workload and interpretation gates

The one-result JSON report for each child must independently pass the frozen
0910 corpus, route, source/output identity, semantic output, sink bound,
source immutability, `warmup=0`, and `samples=1` checks. Those checks occur
outside the owner collection interval and cannot repair ambiguous raw scope.

The measured owner includes opened-presentation catalog capture, slide and
notes validation, MCE work, package cloning, detached transaction construction,
and the internal complete-package fingerprint and payload/Part digest work
described by `protocol-clarification.md`. It excludes the parent code outside
the owner, including ingress, edit/commit, apply, serialization, sink writes,
post-timer output SHA-256, semantic/Part-preservation verification, and
artifact writes. The setup owner row includes corpus-generation transaction
work and is not comparable to the measured snapshot row.

All accepted rows are guest `Ir` instruction counts from one generated
in-memory fixture. They are not native latency, allocation, RSS, physical I/O,
or a causal proof of the 0910 wall-clock route ratio. No source-line snapshot
versus apply attribution may be invented in this debug-info-free binary.

If any child fails these gates, retain its raw artifacts and stop the profile
lane. A fully admitted v3 packet may publish one measured snapshot Ir row per
route, with the setup row retained as an explicitly excluded diagnostic.
