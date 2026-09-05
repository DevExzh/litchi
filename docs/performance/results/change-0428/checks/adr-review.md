# 0428 ADR review

This review covers the additive `try_cache_diagnostics` methods on
`SourceBackedPresentation` and `SourceBackedPresentationEditor` and the
untimed evidence harness in
`tools/perf-baseline/src/pptx_cache_retention.rs`. It records protocol and
ownership fit; it is not a performance result or an approval of the broader
non-iWork goal.

| Accepted ADR | 0428 boundary | Assessment and required guard |
| --- | --- | --- |
| [0001: priorities and API layers](../../../../adr/0001-priorities-and-api-layers.md) | The new methods are explicit source-backed diagnostics returning a typed `Result<SourceCacheDiagnostics, SourceCacheDiagnosticsError>`. The existing infallible compatibility method remains separate. | Fits the explicit low-level/source-backed API layer; no panic or guessed metric is accepted. Ordinary CRUD remains free of cache implementation details. The harness must reject the fallible error and never fall back to the compatibility snapshot. |
| [0002: crate topology](../../../../adr/0002-crate-topology.md) | `litchi-pptx` forwards the already-owned `litchi-opc` diagnostic through its existing downward dependency. The harness stays in the standalone performance tool. | No new production dependency edge, archive ownership, or facade implementation seam is introduced. The harness must use the PPTX owner for the lifecycle rather than opening a second OPC package for diagnostics. |
| [0003: snapshots, edits, patches](../../../../adr/0003-snapshots-edits-and-patches.md) | Source and destination remain immutable source-backed owners; the consuming destination publication and named result/plan/view/sink drops are represented as explicit phase boundaries. | Source identity and output/preservation gates remain unchanged. A consumed destination has an unavailable post-publication diagnostic state; the harness must not fabricate zero or retain a hidden owner solely to sample it. |
| [0005: I/O, memory, measured performance](../../../../adr/0005-io-memory-and-performance.md) | The protocol uses caller-provided positional sources, finite `ReadLimits`, explicit `SourceCacheLimits`, separate source/destination `ExecutionContext` and `Budget` roots, fail-closed cache snapshots, checked counter deltas, and content-free phase records. | This is descriptive resource evidence only. Memory/Object gauges are separated from cumulative InputBytes/Work/OutputBytes; source counters are logical `ReadAt` observations; RSS/VmHWM is process-wide. No latency, physical-copy, allocator, cache-efficiency, leak, or causal optimization claim is allowed. |
| [0006: validation, security, compatibility](../../../../adr/0006-validation-security-and-compatibility.md) | Corpus setup completes the existing semantic, source/output, preservation, refusal, cancellation, and sink gates before samples. Diagnostic calls do not load payloads; near-limit rows require typed refusal and zero output/read deltas where specified. | The fixed synthetic PPTX corpora and bounded limits do not add ambient I/O, networking, or execution. Every row must preserve exact output/source identities and fail closed on incomplete ownership, invalid counters, source changes, or limit violations. Native, cold/range, and adversarial breadth remain unproven. |
| [0024: current topology](../../../../adr/0024-current-topology.md) | The production change remains in the existing `crates/litchi-pptx` presentation owner over `litchi-opc`; `pptx_cache_retention.rs` remains a tool-side observer. | Current workspace ownership and dependency direction are preserved. This note records no package extraction, compatibility alias, or topology exception. |

The protocol's separate source and destination budgets, explicit cache and
read limits, fixed plain/media-rich corpus builders, named lifecycle phases,
and unavailable-owner states are the minimum needed to make the observations
reviewable. The exact-admission, one-byte-under, pinning/eviction,
oversized-bypass, and repeated-publication rows remain evidence gates; they do
not turn this API addition into an optimization.

0428 therefore remains a bounded measurement enabler. Formal capture, replay,
and the global requirements for native/cold/range sources, full CRUD coverage,
streaming/append, scaling, and broader CPU/RSS/memory evidence remain open.
