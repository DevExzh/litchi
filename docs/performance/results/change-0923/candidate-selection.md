# 0923 — DOCX fingerprint opportunity selection

## Selection

The next bounded candidate is an owner-level source-fingerprint proof for the
DOCX source-backed existing-section layout route. It may reuse a complete
source digest between the query snapshot and current-snapshot publication only
when package construction retains an explicit immutable-source capability.
This is a selection for a controlled comparison, not an implementation or an
adoption claim.

## Measured evidence

The 0914 baseline reports a 4.810184 ms pooled total p50. Query and publication
are 1.331 ms (27.68%) and 2.996 ms (62.29%), respectively, or 89.98% together.
Each iteration performs 202 logical source reads returning 8,409,488 bytes.

The 0915 Callgrind profile selects two owners in the same forward route:

| Selected owner | Total IR | Fingerprint IR | SHA-256 `compress256` IR |
| --- | ---: | ---: | ---: |
| Query snapshot | 118,144,825 | 111,584,396 | 109,404,406 |
| Publication current-snapshot subcall | 117,957,908 | 111,584,412 | 109,404,406 |

The independent repeat reports 118,166,114 and 117,959,533 selected total IR
for those owners. Thus the route performs two full source-hash passes under the
profile; their guest instruction counts are attribution evidence, not a native
latency estimate. The 0915 resolver-borrow candidate is separate: its clone
site has 1,838 copies, 3,676 allocations, and 3,674 deallocations per selected
owner, but its native paired ratio was only 1.013824842 with interval
[1.013222680, 1.014816515], below the 1.05 admission gate.

## Ownership and freshness boundary

The candidate may apply only to a private or explicitly typed immutable-owner
route, such as a retained `OwnedSource` capability. The proof must remain
bound to the exact owner allocation, `SourceLineage`, captured
`SourceVersion`, source length, and complete artifact digest. The source owner
must continue to enforce its existing freshness and execution checks before,
during, and after any source read, with the same work and memory accounting.

The current package API erases sources to `Arc<dyn ReadAt>`. A generic adapter's
version token and package lineage do not provide a reusable immutable-byte
proof under the existing contract, so the generic route must continue to run
`SourceArtifact::fingerprint()` in full. Equal main XML from a different ZIP
artifact must still fail the complete-artifact check. A cached source digest
must not be keyed only by version, lineage, or main-part bytes.

The output `FingerprintingWriter` hash is excluded from this candidate. It
identifies the emitted archive and reauthorizes inverse publication; it is not
the duplicate source hash. Existing source-version fences, stale/foreign/
fingerprint errors, limits, output accounting, and final publication checks
remain mandatory.

## Prior work and ranking

- 0916 target-snapshot reuse is already adopted and measured a 6.684% pooled
  p50 improvement; it is not a new candidate.
- 0915 resolver borrowing is rejected by its native gate and supplies no
  evidence for digest reuse.
- 0918 stream handoff, 0920 exact output reservation, 0921 source-bound
  recapture, and 0922 borrowed FIB lookup were rejected by their frozen native
  gates. Their allocation or guest-profile findings do not reopen those
  candidates.
- Deferred DOC physical emission remains a second, higher-theoretical-ROI
  option. The 0919 Float32/32 profile attributes 44,759,310 requested bytes to
  staged rendering and 32,756,816 to staged package recapture. A transaction
  that emits once could address repeated work, but current mutators expose
  per-operation validation and publication behavior.

Deferred emission requires a new scoped transaction contract before any
performance claim: preserve clone-first atomicity, per-operation refusal and
budget boundaries, `retained_render_bytes` and publication errors, source
version and complete-fingerprint fences, protection/signature checks, and
finish-time reopen, recapture, and rediscovery. Deferring those checks or
silently changing observable errors is not an admissible optimization.

The fingerprint candidate is therefore the smaller next experiment, provided
the owner proof is designed first. Its expected native benefit is bounded by
removing at most one complete source-hash pass; the retained 0915 guest counts
do not predict the percentage. A matched native operation comparison is
required before adoption.
