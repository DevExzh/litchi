# Log sections for change 0751

## For `HOTSPOTS.md`

## 0751 — owned cross-copy application stops re-hashing proven bytes

[0751](0751-pptx-cross-copy-apply-digest-reuse.md) removes four of the five
SHA-256 passes that 0742 left in the owned cross-copy's commit phase, and one
pass from planning. The two live revisions are answered from the facades'
payload-digest memos and from a SHA-256 that `litchi-opc` now binds to each
retained owned archive. The two candidate captures consult the snapshots'
memos and each other's. The plan's retained archive is shared instead of
copied; its digest is still recomputed at application, as 0646's G5 requires.

On the media-rich pair, with both legs built by the same command:

- lifecycle median p50: 180.8 → 102.3 ms (paired 0.564);
- commit: 88.7 → 25.7 ms;
- planning: 63.9 → 48.4 ms.

What remains in commit is that one digest (59.5%) and the shared reopen
(25.5%). Planning keeps the first hash of each archive and the candidate
serialization digest.

## For `REPORT.md`

## 0751 — digest reuse in the owned PPTX cross-copy

[0751](0751-pptx-cross-copy-apply-digest-reuse.md) adds to `litchi-opc`, all
additive:

- `OpcPackage::exact_source_sha256`, a digest memo created with the retained
  owned archive by one constructor, filled once and shared by clones;
- shared-archive ingress, `from_shared_vec_reusing_payloads` and
  `from_shared_vec_with_limits`.

In `litchi-pptx`:

- the application's live revisions consult the facades' memos;
- the candidate captures consult the snapshots' memos and the earlier capture
  of the same call;
- an unmodified owned source's physical revision is sealed from the bound
  digest;
- a plan's retained archive is shared with the package it publishes;
- `Package::opened_presentation` keeps its capture's memo when the facade holds
  none and every part is built in;
- `apply_slide_removal_plan` now adopts its snapshot's memo.

Every `LPCP0004` byte, recorded revision, published byte and refusal equals
the base's, pinned by a 98-line golden transcript printed on the base.

Median process p50:

- media-rich lifecycle: 180.8 → 102.3 ms;
- media-rich: 161.9 → 86.4 ms;
- plain: −2.3% and −3.0%;
- source-backed control: unchanged.

Allocated bytes fall 12.3% and peak live bytes 5.4%. The tiny semantic control
moves +0.65% at p50, with per-iteration instructions unchanged (0.9999) and
front-end stalls up. `performance_claim: none`.

## For `GOAL_AUDIT.md`

## 0751 — proven-byte digest reuse under the alpha trade-offs

[0751](0751-pptx-cross-copy-apply-digest-reuse.md) applies 0652's trade-offs:

- **Trade-off 3.** It speeds the common benign path: unmodified owned packages
  captured and published through the facade.
- **Trade-off 2.** It declines three hash skips:
  - the retained archive's digest is recomputed at application, under 0646's
    G5;
  - each archive's first hash at planning stays;
  - packages no memo describes are hashed.
- **Bound memos.** Every skipped hash is answered by a memo bound to its bytes
  by allocation identity, or by a single-constructor digest cell on an
  immutable archive, and every read is re-derived in debug builds. Every
  freshness, revision, physical-provenance, budget and refusal check still
  runs.
- **ADR 0005.** The facade's memo is now also filled by a capture. The record
  proves each condition of the 2026-09-16 memo amendment for that route. It
  reads the amendment's publication sentence as a requirement rather than a
  closed list, and leaves any wording change to the coordinator.

The non-iWork goal stays open.
