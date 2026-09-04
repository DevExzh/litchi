# Change 0402: unmanaged OPC overlay validation decoder reuse

Date: 2026-09-04

Status: implemented; descriptive evidence only. The accepted cells below are
the output of the dedicated publication-phase validator; they are not a
general OPC latency claim.

`performance_claim: none`

`claim_authorized: false`

## Decision and implementation boundary

Retain one operation-scoped indexed-read session while an unmanaged
`SourceBackedPackage` validates the selected source Parts of a multi-Part
overlay. The session reuses one Deflate decoder across the sequential
validation reads. Stored members bypass the decoder, and cache hits remain
cache-only. Managed packages intentionally keep the existing one-shot read
path: retaining decoder workspace across a managed load would escape the
managed memory budget. The managed cancellation and reservation policy is
unchanged.

The implementation is candidate commit
`51964019db3f6b0787645e3a56c2ecb83bdca65c`, measured against control commit
`46ef44966d5be16f153b1f3375ac14401b7139ac`. The production seam is
`SourceBackedPackage::write_part_overlays_to_stream`; the change only
reuses decoder state during its selected-source validation. It does not alter
the overlay limit, source lookup, compression/CRC/declared-size checks,
candidate XML validation, source freshness, cancellation, signatures,
partial-sink behavior, cache publication, managed refusal, or the complete
raw-copy/regeneration plan. The measured selector uses a non-empty
equal-payload replacement plan (a semantic no-op), so its timing evidence is
not a claim for changed-payload publication or other overlay modes.

## Fixed matrix and correctness oracle

The opt-in selector is
`opc_source_overlay_multi_part_noop`. It covers the fixed three-shape by
three-count matrix below. Every shape has 32 ordinary entries and 34 archive
members; the target overlay count is 2, 8, or 32. The generated corpus
identities are:

| Shape | Payload and entry size | Archive bytes | Uncompressed payload bytes | Archive SHA-256 |
| --- | --- | ---: | ---: | --- |
| `overlay-small` | compressible, 1 KiB | 7,451 | 32 KiB | `4338dea03f37b0ea2ad63a055fb5cfb7df79a5b0de864365e981e453e1a65509` |
| `overlay-large` | incompressible, 64 KiB | 2,103,195 | 2 MiB | `8356d7467215b04a3d1c3703f50fbd6322f2002ca7c3ead1f24414c5e550ef73` |
| `overlay-media-incompressible` | incompressible, 256 KiB | 8,396,580 | 8 MiB | `bf8c309af5306c6682b9df65b97246f81b022fe5e3b5e02cc2c4dcf3e1e87883` |

The no-op output reopens through the eager OPC reader with the expected
semantic identity, preserves raw member order and untouched ZIP records, and
has the same archive SHA-256 as its source for every retained sample. These
are correctness identities, not timing or physical-I/O evidence.

## Normal publication-phase evidence

The normal release run uses stable Rust/Cargo/Rustdoc 1.98.1, CPU affinity
2, one execution worker, and strict `A1(control) / B1(candidate) /
B2(candidate) / A2(control)` order. Each leg has 20 warmups and 500 retained
**in-process** samples. The configured global cache-state envelope is
`["warm", "cold-requested"]`; it is configuration metadata, not evidence of
a cold run. No fresh-child or process-isolated semantics are claimed for
these normal operation samples.

The dedicated validator verifies the top-level phase identity
`elapsed_ns = preparation_ns + open_ns + planning_ns + publication_ns` for
each sample, but does not summarize top-level elapsed time. All accepted
statistics below are recomputed only from
`source.opc_source_overlay.publication_ns`. Positive percentages mean that
the candidate publication phase is lower. The per-statistic drift ceilings
are 5% for p50/mean, 10% for p95, and 15% for p99.

| Shape / overlay count | Accepted publication statistics | A1 → B1 reduction | A2 → B2 reduction |
| --- | --- | --- | --- |
| `overlay-small / 2` | none | — | — |
| `overlay-small / 8` | p50, mean, p95, p99 | p50 `+15.135453%`; mean `+14.915579%`; p95 `+14.550562%`; p99 `+9.750859%` | p50 `+12.433393%`; mean `+12.919635%`; p95 `+13.126761%`; p99 `+5.805085%` |
| `overlay-small / 32` | p50, mean, p95, p99 | p50 `+22.344144%`; mean `+22.316386%`; p95 `+27.338269%`; p99 `+20.303797%` | p50 `+21.426460%`; mean `+21.438222%`; p95 `+23.738113%`; p99 `+16.127923%` |
| `overlay-large / 2` | none | — | — |
| `overlay-large / 8` | none | — | — |
| `overlay-large / 32` | p50 only | `+1.095149%` | `+2.088386%` |
| `overlay-media-incompressible / 2` | p50, mean, p95, p99 | p50 `+1.510196%`; mean `+1.927179%`; p95 `+3.881384%`; p99 `+5.553444%` | p50 `+0.814482%`; mean `+1.216448%`; p95 `+1.447249%`; p99 `+3.045589%` |
| `overlay-media-incompressible / 8` | none | — | — |
| `overlay-media-incompressible / 32` | none | — | — |

The withheld cells include adverse paired directions and/or drift failures;
they are not silently pooled with accepted cells. This is a partial matrix
of publication-phase observations, not an overall or end-to-end latency
improvement claim.

## Allocator observation

A separate allocator ABBA uses the same four-leg shape, CPU, worker, and warm
configuration with three warmups and 30 retained samples per leg. The
operation-scoped `CountingSystemAllocator(std::alloc::System)` vectors are
constant per shape/count. Candidate-minus-control deltas are:

| Shape / overlay count | Allocation calls | Deallocation calls | Reallocation calls | Failed allocation calls | Allocated bytes | Deallocated bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `overlay-small / 2` | -2 | -2 | 0 | 0 | -80,320 | -80,320 |
| `overlay-small / 8` | -14 | -14 | 0 | 0 | -562,240 | -562,240 |
| `overlay-small / 32` | -62 | -62 | 0 | 0 | -2,489,920 | -2,489,920 |
| `overlay-large / 2` | -2 | -2 | 0 | 0 | -80,320 | -80,320 |
| `overlay-large / 8` | -14 | -14 | 0 | 0 | -562,240 | -562,240 |
| `overlay-large / 32` | -62 | -62 | 0 | 0 | -2,489,920 | -2,489,920 |
| `overlay-media-incompressible / 2` | -2 | -2 | 0 | 0 | -80,320 | -80,320 |
| `overlay-media-incompressible / 8` | -14 | -14 | 0 | 0 | -562,240 | -562,240 |
| `overlay-media-incompressible / 32` | -62 | -62 | 0 | 0 | -2,489,920 | -2,489,920 |

These are exact allocator call/byte observations for the fixed no-op matrix.
Allocator elapsed time, live bytes, high-water values, RSS, and total memory
are not claim metrics.

## Correctness and validation

The `litchi-opc` correctness gate passed 289 library tests and 386 total test
items. Focused source-backed coverage includes mixed Store/Deflate semantic
parity, managed cold-work/budget behavior, raw ZIP preservation, exact no-op
output, and the existing overlay limit, cancellation, freshness, and sink
failure boundaries. The dedicated validators also passed:

- `tools.test_perf_opc_overlay_abba_summary`: 10 tests;
- `tools.test_validate_opc_overlay_allocator_abba`: 22 tests;
- publication-phase summary SHA-256:
  `65e78362e712f15d73383102dd129ce96ba4f07b7073e41b91e1ed92c9cd4085`.

The summary validator checks report provenance, corpus and oracle identity,
sample order, in-process scope, sink/source counter vectors, phase sums, and
the nested publication statistics. The allocator validator checks the four
allocator reports, all nine matrix rows, and their exact per-sample vectors.
The retained [Change 0402 evidence bundle](../results/change-0402/) contains
the compressed raw reports, projections, manifests, and validator bindings.

## Claim boundary

This change deliberately makes no claim for top-level elapsed latency,
general/end-to-end latency, allocator elapsed latency, RSS, peak/live operation memory, physical I/O,
decompression or recompression volume, throughput, scaling, or cache
temperature. The global `["warm", "cold-requested"]` setting does not turn
the in-process normal samples into cold evidence, and no fresh-child claim is
made. No general OPC/OOXML, managed, eager, mutable, changed-payload,
other-overlay-mode, real-producer, or iWork claim follows. The partial
publication matrix remains intentionally absent from `claim-registry-v1.json`:
the v1 extension now supports the closed custom `publication_ns` scope and
partial accepted/adverse-cell adjudication, but this change adds no 0402
registry claim entry. Allocator vectors remain outside the registry boundary
and are validated by `tools/validate_opc_overlay_allocator_abba.py`; the full
self-excluding inventory remains outside it and is checked through the
retained `evidence-manifest.json` hashes and documented bundle
integrity/audit checks.

## Reproduction and artifacts

The canonical raw and projected artifacts are linked from the durable
[Change 0402 evidence bundle](../results/change-0402/). Its identity is bound
by `summary.json` SHA-256
`65e78362e712f15d73383102dd129ce96ba4f07b7073e41b91e1ed92c9cd4085`, the
package manifest `0402-opc-source-overlay-abba-manifest.json` SHA-256
`beb4953f45617c0925c4e7d44b20e7a86c75c0ea95ed20c3829cb42a491da5b7`, and the
self-excluding `evidence-manifest.json` SHA-256
`0e748923f9da1e1173562a90014970173f3b8350de04b0de5144924118cec5e3`.
The normal projection is generated by
[`perf_opc_overlay_abba_summary.py`](../../../tools/perf_opc_overlay_abba_summary.py);
the allocator projection is generated by
[`validate_opc_overlay_allocator_abba.py`](../../../tools/validate_opc_overlay_allocator_abba.py).
The bundle binds the exact control/candidate revisions above, the normal and
allocator binary identities, all eight raw report frames, and the self-
excluding evidence manifest.
