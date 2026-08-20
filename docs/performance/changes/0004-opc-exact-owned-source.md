# OPC exact owned-source no-op publication

Status: accepted as a narrow preservation stage
Production base: `2665d572b78f0b3efd9ecfc4bd1fda09f8786ae3`

## Mechanism

Successful owned OPC ingress (`open`, `from_reader`, and `from_vec`) retains
the already-validated archive allocation in `Arc<Vec<u8>>`. An unchanged
package can publish that exact archive to a sequential sink in chunks no
larger than 64 KiB or copy it fallibly into the requested output `Vec`.

Entering any mutable `OpcPackage` API revokes the authorization
conservatively, including failed and semantic no-op calls. A revoked package
uses the complete prevalidated `PublicationPlan` rewrite. Clones share the
source allocation but revoke independently. Borrowed `from_bytes` ingress does
not acquire an owner or authorize passthrough.

This preserves the original ZIP bytes—not only OPC semantics—including entry
order, compression streams, timestamps, extras, comments, unknown non-Part
members, and signature framing. DOCX, PPTX, and XLSX public-owner tests append a
nonzero EOCD comment and prove byte-identical no-op output.

## Save latency

Matched release binaries used identical memory-backed sequential sinks. The
sink reserves its checked budget before timing, copies every output byte, caps
individual writes at 64 KiB, and compares the complete output after timing.
Runs were interleaved before/after/after/before. Non-large cells contain 200
samples per replicate; four-4-MiB cells contain 30. Raw reports are the eight
`results/abba-exact-source*.json` files excluding the separately named open
and mutated-save reports.

| No-op save corpus | Before p50 | After p50 | Change | Before/after writes |
|---|---:|---:|---:|---:|
| 256 Parts, compressible | 1.510 ms | 0.001 ms | -99.97% | 1,813 / 1 |
| 256 Parts, incompressible | 5.516 ms | 0.003 ms | -99.95% | 1,813 / 5 |
| 2,048 Parts, compressible | 11.977 ms | 0.005 ms | -99.96% | 14,357 / 7 |
| 2,048 Parts, incompressible | 17.028 ms | 0.006 ms | -99.97% | 14,357 / 7 |
| four 4 MiB Parts, compressible (99 KB ZIP) | 3.185 ms | 0.001 ms | -99.97% | 49 / 2 |
| four 4 MiB Parts, incompressible (16.78 MB ZIP) | 211.531 ms | 3.443 ms | -98.37% | 557 / 257 |

The small hot-memory copy rates are not filesystem throughput claims. The
material result is removal of XML regeneration, ZIP construction, and Deflate
work while a real sink still consumes and verifies every byte.

Heaptrack comparisons used 100 many-small saves and 10 incompressible
four-large saves:

| Workload | Allocation calls | Temporary allocations | Peak heap | Profiler RSS |
|---|---:|---:|---:|---:|
| many-small before | 225,466 | 53,285 | 2.28 MB | 13.53 MB |
| many-small after | 14,270 (-93.7%) | 1,685 (-96.8%) | 2.20 MB | 13.26 MB |
| four-large before | 1,500 | 236 | 72.26 MB | 66.22 MB |
| four-large after | 754 (-49.7%) | 90 (-61.9%) | 88.62 MB (+22.6%) | 83.08 MB (+25.5%) |

The large-package peak increase is expected and material: exact publication
retains the 16.78 MB compressed source in addition to the current eager 16 MiB
Part materialization. It is accepted for exact preservation and save latency,
but it strengthens the priority of source-backed lazy Part materialization.

## Open and changed-save guardrails

The established `OwnedPhysPkgReader` validation/index path is retained; source
bytes are recovered through its existing zero-copy `into_inner` handoff. A
one-worker, single-CPU ABBA run avoids the known global-Rayon variance and
shows four-large owned-open p50 changing from 4.184 to 4.047 ms
(compressible, -3.3%) and 2.293 to 2.225 ms (incompressible, -3.0%). Parallel
owned-open distributions remain scheduler-sensitive and are not claimed as an
improvement.

`opc_mutated_save` changes one byte in the target Part before timing and
verifies the deterministic full rewrite. With a fixed CPU, 256-Part
compressible output improves about 3.6% and incompressible output differs by
about 0.2%; sink byte/write summaries match. Raw reports are the
`results/abba-exact-source-open-*` and
`results/abba-exact-source-mutated-*` files.

## Correctness and remaining boundary

- OPC: 116 tests plus 5 doctests; all-target/all-feature warning-denied Clippy.
- Exact source with a nonzero ZIP comment survives `to_bytes`, sequential
  streaming, and public DOCX/PPTX/XLSX no-op output.
- Clone-local revocation, edited-Part reopen, relationship/options/signature
  invalidation, 64 KiB chunking, and partial accepted-byte accounting are
  covered.
- No unsafe code, dependency, public archive type, executor, or lock was added.

This is not lazy OPC, unchanged-entry copy-through after edits, remote/range
I/O, or ADR 0005 completion. The current `Part::blob() -> &[u8]` contract cannot
represent fallible deferred loading or pin evictable cache data. The staged
follow-up requires a fallible owning `PartData`/metadata-only `PartView`, one
immutable positional ZIP index, source-version checks, a byte-weighted
single-flight cache, and raw-entry provenance.

## 2026-08-20 safety amendment

The exact-source fast path remains the sole byte-identical no-op path. After
any mutation revokes that authorization, an owned source uses targeted
preservation and now returns the typed `OpcError::PreservationUnavailable`
before sequential sink output when prefixes, trailing/junk members, ZIP64 or
other framing, opaque members, or changed topology cannot be proven safe. The
ordinary full writer remains available for new packages and borrowed unsigned
sources; it is not a fallback for a mutated owned source.

Signature tracking now treats origin/signature/certificate relationships,
signature targets and paths, signature content types, orphan infrastructure,
and mutable package/Part relationship seams conservatively. Changed signed
sources return `OpcError::SignedSourceRequiresExplicitPolicy` unless the
explicit sign, resign, or unsign operation authorizes the resulting graph.
Untouched borrowed signed sources have no exact bytes to copy and therefore
return the same typed policy refusal instead of normalizing signed content.
These capability errors map to `litchi_core::Error::Unsupported` through the
OPC core conversion and the XLSX facade adapter, rather than `Other`.
