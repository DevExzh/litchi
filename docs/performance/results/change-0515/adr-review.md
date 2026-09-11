# 0515 ADR scope

All 30 previously read ADR documents, including the index, were checked against
the 0514 hashes before this investigation; none changed. The current source
manifest and base Git objects are checked separately by the evidence verifier.
No production Rust, harness Rust, feature, dependency, public API, cache,
source provider, execution policy or publication behavior is changed.

| Constraint | Evidence and application |
| --- | --- |
| 0001/0004 API layers and semantic ownership | Existing public XLSX commit and sequential-save benchmark cases exercise the same selectors. No new API is introduced. |
| 0002/0010/0011/0024 owner boundaries | Profiling uses external Callgrind, an existing standalone harness and unchanged owners. No archive implementation is imported into a facade. |
| 0003 immutable snapshots, exact no-ops, atomic publication and patches | The existing commit, projection, output validation, semantic readback and bounded Store handoff remain intact. The source review identifies the proof a future output-fusion candidate would require. |
| 0005 evidence, memory and execution | Serial CPU-2 captures use one worker on a recorded shared host. Source, binary, corpus, raw vectors and receipt hashes bind measurements. Simulated instruction profiles, normal timings and whole-child RSS are distinct scopes. Phase allocation and hardware/cache metrics are unavailable in this batch. |
| 0006 preservation, validation and errors | No required parse or validation is removed. In particular, pre-compaction input events cannot automatically prove emitted-byte semantics. Existing malformed-input and error-precedence behavior remains authoritative. |
| 0008 verification state | This is attribution evidence, not a completed format optimization or a new broad correctness certification. The full OLE2/OOXML goal stays active; ODF is deferred and iWork excluded. |
| 0017/0018 producer templates and calculation chain | Deterministic existing corpus construction and ordinary cell updates use unchanged generation and invalidation paths. |

There is no ADR conflict or proposed amendment. A later implementation must
supply its own matching correctness, preservation, performance and memory
evidence; this diagnostic result does not authorize omission of output checks.
