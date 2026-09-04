# Change 0397: OPC owned-open validation/index deduplication

Date: 2026-09-04

Status: accepted scoped production change. The retained timing claim is limited
to the normal owned `OpcPackage::from_vec(owned)` constructor and the fixed
benchmark protocol below.

`performance_claim: scoped`

`claim_authorized: true`

## Production change and boundary

Production commit `f275d4566` removes one redundant eager ZIP
validation/index pass before construction of the real `PhysPkgReader` on the
owning `OpcPackage` path. The authorized timing scope is only
`OpcPackage::from_vec(owned)`. `authorize_owned_source` still performs the
preservation work needed for exact owned-source behavior; this is not a claim
that only one ZIP index exists overall. The public `OwnedPhysPkgReader`
constructors remain eager-validating.

The path and `from_reader` constructors, and the scheduled/session path, are
covered for correctness but are not separately timed here. Read limits,
error ordering, and session input charging remain preserved. Exact mixed-
storage byte preservation is covered by
`owned_constructors_preserve_mixed_storage_and_exact_source`,
`owned_constructors_check_input_limit_before_zip_validation`, and
`owned_constructors_report_malformed_zip_after_bounded_read`; the session
tests `explicit_owned_open_rejects_malformed_zip_before_input_charge` and
`explicit_owned_open_rejects_input_limit_before_zip_and_input_charge` cover
the corresponding refusal and charge order.

## Benchmark protocol

The opt-in `opc_casefold_owned_open` selector raises the selectable registry
from **418** to **419**; the default remains **36 cases / 198 rows**. It uses
fixed stored OPC corpora with exactly 256, 2,047, 2,048, and 16,384 ordinary
Parts, each with a 32-byte payload. The 2,047-Part corpus is the boundary case
immediately below the 2,048 source case-fold threshold.

The candidate is revision
`f20d3f417edc3f3da07bf515676b8e71285ad76f`; the control is
`6e98db9ece29c1e50241cf3e84c9410ce71dd748`. Release measurements used CPU 2,
one worker, balanced A1/B1/B2/A2 ABBA order, five warmups, and 30 samples.
The measured toolchain was rustc/Cargo/Rustdoc 1.98.1; the pinned 1.95
toolchain could not be used because it lacks Cargo.

All latency values below are normal, non-allocator release-binary p50
speedups, where a positive value means the candidate is faster. Values are in
256 / 2,047 / 2,048 / 16,384 ordinary Parts order. The allocator target's
elapsed time is observational only.

| Corpus | A1→B1 normal p50 speedup | A2→B2 normal p50 speedup | pooled normal p50 |
| ---: | ---: | ---: | ---: |
| 256 | `+8.617829%` | `+8.204676%` | `+8.452941%` |
| 2,047 | `+8.298670%` | `+8.719476%` | `+8.356702%` |
| 2,048 | `+8.945417%` | `+8.268274%` | `+8.490980%` |
| 16,384 | `+4.648655%` | `+4.348226%` | `+4.645459%` |

## Allocation evidence

Allocator measurements are exact candidate-minus-control call and byte
vectors, not latency evidence. On each matched ABBA leg, the exact
alloc/dealloc call reductions, in the corpus order above, are
`-1,038 / -8,202 / -8,206 / -65,550`. The corresponding
allocated/deallocated-byte reductions are
`-152,024 / -1,212,620 / -1,212,888 / -9,699,800`. Per-sample net-live
after-before bytes and reallocations are exactly unchanged; raw global
live-before/after baselines are not cross-run metrics. These vectors do not
establish RSS, peak operation memory, total memory, or a system-memory result;
allocator-enabled elapsed time remains observational only.

## Provenance and correctness

The candidate source hashes are:

| File | SHA-256 |
| --- | --- |
| `crates/litchi-opc/src/execution.rs` | `239d969195eb30ed9832d59df54aefd4bfff8e0020dc49fa32d6162fec6be519` |
| `crates/litchi-opc/src/package.rs` | `e2e5794b79e66aa4aef2212b9eeeda7e6012eb33da43aa0a3137cba93539f81a` |
| `crates/litchi-opc/src/phys_pkg.rs` | `228498510d9346ca7c57d7b5d30939e2aa9349aed38346e71d2c36141e7780e8` |

The [0397 evidence bundle](../results/change-0397/) retains the reports,
corpus identities, hashes, allocator vectors, and adjudication. Earlier
invalid captures were rejected and deleted; they are not evidence and do not
contribute to the result.

The accepted claim is only the normal, non-allocator `OpcPackage::from_vec(owned)`
p50 result on these four fixed stored corpora under the stated CPU-2 ABBA
protocol; no p99 claim follows. Allocator elapsed time is observational only.
No claim follows for
path, `from_reader`, or session timing; public `OwnedPhysPkgReader` timing;
validation-constructor timing; RSS or peak operation memory; physical I/O;
cold-cache behavior; throughput; other formats or facades; or generalization
beyond this owning constructor and corpus.
