# Change 0387: OPC source materialization shared payload

Date: 2026-09-03

Status: implemented

`performance_claim: none`

`claim_authorized: false`

## Scope and mechanism

Unmanaged `SourceBackedPackage::{into_opc_package,to_opc_package}` conversion
used to call `PartData::as_bytes().to_vec()` for every admitted Part and then
pass that new `Vec` to `PartFactory::load`, which allocated an additional
`Arc<Vec<u8>>` owner. The conversion now clones the cache's immutable
`Arc<Vec<u8>>` handle and passes it to `PartFactory::load_shared`. This removes
one full logical payload copy and one new Arc allocation per Part without
changing Part classification, relationship copying, package construction, or
source/signature policy.

Managed packages remain refused before ordinary payload reads. Their
`PartData` handles retain hierarchical reservations and still cannot escape as
bare owning-package allocations. The conversion retains all source-version,
cancellation, execution-context, allocation-failure, relationship, and final
freshness checks. Mutating a materialized Part uses the existing replacement
path and does not alter the source-backed cache.

## Matched evidence

The release allocator target measured revision
`7edfad113e62f45025848273ef04d6a94c433b27` against that same revision plus
the exact `source_backed.rs` patch. The control/candidate source blobs are
`a39b8272e329e6f4cee5eeeddc24d2a97f7ca533` and
`2b1058c42eee28c5c9eaeb7afe9ca81702425676`; the binary patch SHA-256 is
`189717f6eb5ac86d21d13f5ccd3b0c23a1d260db5602af327b449f52e89292aa`.
The later `78e961bc4` revision changed only CI, Python policy tests, and its
change record; the final Rust file still has the measured candidate blob.

Each report used three warmups and 15 retained in-process samples. The
allocation-call and allocated-byte vectors used below were constant across all
15 samples, and the logical Part/read vectors were identical between control
and candidate. Absolute live-byte gauges varied with retained allocator state
and are not used for the decision.

| Deterministic corpus | Parts | Allocation calls, control -> candidate | Allocated bytes, control -> candidate | Exact removed volume |
|---|---:|---:|---:|---:|
| tiny compressible, 1,536 uncompressed payload bytes | 3 | 53 -> 47 | 246,756 -> 245,100 | 6 calls; 1,656 bytes |
| many-small incompressible, 262,144 payload bytes | 256 | 3,609 -> 3,097 | 21,290,113 -> 21,017,729 | 512 calls; 272,384 bytes |
| few-large incompressible, 16,777,216 payload bytes | 4 | 69 -> 61 | 33,879,589 -> 17,102,213 | 8 calls; 16,777,376 bytes |

For all three shapes, the delta is exactly two allocator calls per Part and
the declared uncompressed payload volume plus 40 bytes per Part. That matches
the removed `Vec` payload allocation and the removed `Arc<Vec<u8>>` owner
allocation. This is operation-local allocator/mechanism evidence, not a
direct copied-byte counter.

The checked-in [machine-readable summary](../results/opc-source-materialization-shared-0387-summary.json)
binds the source blobs, patch, binaries, environment, corpora, report hashes,
and exact vectors. Its [raw artifact directory](../results/change-0387/)
contains zstd-compressed normal and allocator-enabled reports plus their
schema-v2 corpus catalogs. All 12 operation-metric envelopes and all 12
report/catalog bindings validated. Normal-binary reports retained allocator
status as explicitly unavailable.

## Correctness and ownership verification

Two focused tests prove allocation identity with `Arc::ptr_eq` for consuming
and borrowed conversion. The owning package remains valid after all
source-backed/source/cache handles are dropped. The borrowed case also mutates
the owning package and proves copy-on-write separation from the still-readable
source cache before dropping it.

Independent validation on explicitly selected stable Rust 1.98.1 passed:

- `litchi-opc` library: 279 tests;
- `litchi-opc` unit plus integration targets: 279 plus 97 tests;
- `litchi-opc` doctests: 5 tests;
- strict all-feature, no-dependency `litchi-opc` library Clippy;
- DOCX source-backed semantic tests: 8 tests;
- PPTX validation tests: 12 tests.

The repository-pinned Rust 1.95 installation lacked Cargo, rustfmt, and Clippy
components in this environment. All-target Clippy reached one pre-existing
`needless_lifetimes` finding in `tests/phys_pkg_borrowed.rs`; four XLSX
row-visibility failures reproduced unchanged on a clean control archive and
do not enter this conversion path.

## Claim boundary and remaining work

The harness labels this selector
`evidence_only_opc_source_materialization`; its latency is deliberately not
comparable. Runs were not balanced ABBA and were not pinned to one core, so no
latency or throughput statistic is accepted. Procfs RSS deltas include probe
overhead and reported peak RSS is process-lifetime, not operation-local, so no
peak-memory/RSS claim is made. Copied, decompressed, recompressed, physical-I/O,
and lock-wait bytes remain unavailable and are not inferred.

This removes a duplicate allocation only when an unmanaged source-backed view
is explicitly converted to the owning package. It does not make ordinary
`OpcPackage` open lazy, avoid the conversion's all-Part materialization, change
managed-package policy, add parallelism, or establish a real-producer/broad
OOXML result. No strict claim-registry entry is added.
