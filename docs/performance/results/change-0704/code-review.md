# 0704 production candidate code review

The independent review covers the authoritative main tree after the coder
handoff and append/sort correction. Coordinator completion notes update the
review's test status: all 26 focused tests and seven integration gates pass.
The original coder patch is preserved under `preflight/`; the root
`candidate-production.patch` now records the full final candidate diff. The
reviewer ran no Cargo commands and changed no production code.

The ownership and value-preservation design is sound. The raw owner is obtained
from the completed `PartDigests` memo before an entry is published, so a copied
foreign `blob_arc()` cannot pin or answer for a retained source. Holding that
owner closes the address-reuse/ABA hole. `RetainedMce::lookup` uses
the sorted `(pointer, length)` table with binary search, and the debug hit path
re-runs `process_ooxml` and compares the bytes. The processed `Arc<Vec<u8>>` is
shared without copying, while root, producer name, notes proof, content type,
relationship, and refusal checks still run in their established order. The
`SlideRootProof` raw borrow remains alive through notes validation; only the
processed XML is dropped before a fallback part-name allocation.

## Findings

### F1 — capture-local quadratic scans are resolved on main

The earlier handoff used `used.contains` and pending `.find` scans. Main now
keeps a checked running `pending_charge`, records optional visits in one
capture-local vector, and sorts/deduplicates that vector once during finish.
Parent lookup remains binary-searchable. The resulting capture-local path is
linear admission plus O(n log n) final sorting, while repeated raw visits may
still recompute on a cold capture as documented in the design.

The source slice witness remains in each pending record, and a reservation or
budget refusal still falls back to the ordinary result. This finding is closed;
the stale line references in the first review are retained only in repository
history, not as a current blocker.

### F2 — published logical charge and transient capture peak are now explicit (resolved)

The final charge is now calculated correctly from the actual final entry-vector
capacity, processed-vector capacities, entry metadata, and the outer
`Arc<RetainedMce>` allowance. It does not double-count the raw package payload,
and the stale charge-after-pop defect from the old packet is gone.

During capture, pending, candidate, and final entry vectors can coexist, but
main now documents that their vector metadata is transient workspace. The
running provisional charge conservatively counts every admitted visit,
including duplicates; final `charged_bytes` counts unique entries once with the
actual published capacity. Thus the accessor is explicitly a logical
retained-table budget rather than a peak RSS measurement, and entry metadata is
not silently counted twice in the published charge.

### F3 — prefix admission is safe but should be named (P2)

`RetainedMce::from_candidates` performs one forward prefix admission and stops
at the first candidate that does not fit. A later, smaller output is therefore
discarded even if it would fit independently. The O(n) prefix choice avoids the
old repeated descending allocations and never changes a value or refusal, but
it is a retention-quality policy rather than a best-fit aggregate budget. Keep
it if the intended fallback is prefix retention and document that choice; add a
test or measurement if retaining later small slides matters to the changed
commit target.

### F4 — deterministic final-table reservation fallback is covered

Coordinator completion: the final private reservation seam requests an
impossible entry capacity, observes `CapacityOverflow`, returns no table and
proves both candidate owners are released. The independent accounting review
accepted this narrow test together with disabled/over-ceiling semantic parity.
Pending and projection reservation failures remain source-reviewed; no global
allocator injection or recoverable `Arc::new` OOM claim is made.

## Resolved checks

The following issues found in the stale packet are fixed on authoritative main:

- Capture-local visits are appended only after fallible reservation; the
  quadratic `used` vector is removed.
- Final table charge uses actual `Vec` capacity and trims over-reserved tables
  before publication, so `charged_bytes` cannot describe a rejected capacity.
- Entry-vector and outer `Arc<RetainedMce>` metadata are included once in the
  checked logical charge.
- Published table lookup is sorted/binary rather than a linear scan.
- Cache hits use the existing source slice and do not add `blob` or `blob_arc`
  observations; fresh processing remains the fallback.
- Parent, rebound, and candidate entries are admitted only after the current
  package's digest memo proves ownership. A same-byte new allocation misses by
  identity, while an unrelated current owner may retain a bounded unused
  projection safely because default MCE processing depends only on raw bytes.
- Snapshot clone, local release, changed-commit reuse, no-op commit sharing,
  and commit release preserve semantic values and package bytes.
- `Limits::with_max_retained_mce_bytes` is additive, zero disables only the
  optional retention, the six-argument constructor keeps its old nonzero
  behavior, and patch/cross-copy limit intersections carry the tighter MCE
  ceiling.

F1/F2 are closed on the final tested source. F3 remains an explicit prefix
admission choice, not a correctness blocker. F4 has the deterministic narrow
evidence accepted in `accounting-review.md`. Native measurements and retained
memory costs support only the scoped disposition in the root report.
