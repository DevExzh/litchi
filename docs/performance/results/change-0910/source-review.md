# 0910 harness source review

## Review binding

This review is a read-only source inspection. It binds the review to the
following current files and parent source:

| Item | SHA-256 |
| --- | --- |
| parent source | `468fe2931d6532c32f4c6608590816b8eaa88e31` |
| `tools/perf-baseline/src/pptx_owned_benchmark.rs` | `121a4b8817da5410222d4bd35252327cea75fc73866c3b7f8ed9f71b15a3aa46` |
| `tools/perf-baseline/examples/pptx_owned_benchmark.rs` | `1dee0ad9fa5c24188912df28477a585beb346864491138515e4bb27d5b2a7675` |
| `tools/perf-baseline/src/lib.rs` | `8542f4e8e580bef2b12b3782b03bcfabc8e0d60cefc80f27c55042c57baff860` |
| `docs/performance/results/change-0910/protocol-review.md` | `d527584ebeeb835e5a8db5c9bbe118f3658059de35c731da42fab129e12cf6aa` |

The source files are the 0910 overlay over the 0909 parent. No production
library path or normal selector was changed by this benchmark module.

## Findings

The route bodies and primary input boundary match the reviewed protocol. The
owning route uses `Package::from_vec(input)`. The materialized route uses
`OwnedSource::new(input)`, `SourceBackedPackage::from_read_at`,
`into_opc_package`, and `Package::from_opc_package`. The same generated corpus,
64 text replacements, transaction, commit, and `Package::to_bytes` operation
are used after ingress. The byte clone is timed for both routes, and the
inner lexical scope drops route-local package state before `total_ns` is read.

The code does not use an instrumented `ReadAt` adapter in the primary route,
so it avoids adding counter or atomic overhead to only the materialized leg.
Semantic output verification and digest stability checks occur after each
timed iteration. Source and output artifacts are created after the measured
loop, and the output artifact reuses the final already-verified sample rather
than adding an extra timed route execution.

## Final source decision

The reviewed source now uses one preallocated sink shared by the warmups and
measured iterations. It is reset outside each timer, receives both routes'
`to_bytes` output in 64 KiB chunks inside `publish_ns`, and remains alive for
post-timer byte-count and write-count gates. The sink capacity and maximum
write bound are recorded in the report. The route package, transaction,
source owner, temporary output, and other route-local state are dropped before
`total_ns` is read.

The current report's `semantic_output_verified` and `stable_output_verified`
fields are proof-state markers reached only because `verify_output` returns
success; the independent reducer should bind the per-sample output digest and
the semantic guard rather than trusting those literal booleans alone.

## Scope decision

The source is approved for the planned native capture. The source review does
not run a build, workload, profiler, or replay and makes no performance claim.
Capture remains limited to the generated in-memory fixture and the explicitly
recorded 0910 build.
