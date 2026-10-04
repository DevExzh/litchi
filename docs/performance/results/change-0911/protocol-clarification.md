# 0911 scope clarification: internal fingerprinting versus output verification

This note supplements the hash-bound 0911 protocol review. It does not change
the binary, Callgrind options, owner, schedule, or admission gates in that
review.

The selected Callgrind owner is
`litchi_pptx::package::model::Package::opened_presentation_transaction`.
The native `snapshot_ns` clock starts immediately before the call at
`tools/perf-baseline/src/pptx_owned_benchmark.rs:294` and stops immediately
after it at line 295. The owner is implemented by
`Package::opened_presentation()?.edit()` (`crates/litchi-pptx/src/package/model.rs:296-298`).
Consequently, all work reached by that expression belongs to the selected
snapshot attribution.

That selected interval includes the opened snapshot's internal complete-graph
revision and payload-digest work. The fresh package created by each benchmark
iteration starts with an empty `part_digests` slot: both `Package::from_vec`
and `Package::from_opc_package` initialize `OnceLock` empty
(`crates/litchi-pptx/src/package/codec.rs:225-239, 264-308`). The first owner
call therefore follows `opened_presentation_with_limits` into
`capture_with_provenance`, `capture_internal`, and the `Revision::Cold` branch
(`crates/litchi-pptx/src/package/model.rs:253-288`;
`crates/litchi-pptx/src/opened/model.rs:657-670, 730-853`). In that branch,
`package_fingerprint_with_memo` sorts the parts, feeds root relationships and
non-part members, and computes or records per-part payload and part digests
before finalizing the complete-package revision
(`crates/litchi-pptx/src/opened/model.rs:880-991`). Snapshot catalog capture,
slide/notes validation, MCE work, package cloning, and construction of the
detached transaction are also within this owner. The transaction constructor
clones the captured package (`crates/litchi-pptx/src/opened/model.rs:547-551`).

This means a raw owner `Ir` total may legitimately contain substantial
`package_fingerprint_with_memo`, payload hashing, and logical revision work.
Those costs are part of the existing `opened_presentation_transaction` phase;
they must be retained in the total and may be shown as direct-child or nested
diagnostic rows without double-counting. The profile must not describe this
owner as “snapshot without digesting” or subtract internal fingerprint costs.

The phrase “digesting outside the owner” refers only to later benchmark
verification. After `timed_iteration` returns, `verify_output` reopens the
serialized output, runs the semantic and Part-preservation oracle, computes
`sha256_hex(output)`, and checks the stable output digest
(`tools/perf-baseline/src/pptx_owned_benchmark.rs:335-351, 410-418`). Source
identity hashing, source immutability checks, report serialization, and artifact
writes are likewise outside the native phase clocks and outside the selected
Callgrind owner. They must not be attributed to the snapshot profile, and they
must not be reconstructed or subtracted from it.

The later edit/commit loop, `apply_opened_presentation_commit`, `Package::to_bytes`,
and bounded sink writes occur after the owner returns
(`pptx_owned_benchmark.rs:297-321`). Thus the narrow profile measures the
implementation and direct call tree of the opened snapshot, including its
internal logical fingerprint and payload hashing, while excluding the
post-timer output digest/oracle and later publication work. The debug-info-zero
limitation remains: report function and edge attribution, never invented source
line attribution for these boundaries.
