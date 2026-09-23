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

- commit: 88.9 → 25.8 ms;
- planning: 67.0 → 55.3 ms;
- lifecycle median p50: 184.2 → 109.8 ms (paired 0.61), and 0.565 at equal
  page faults.

What remains in commit is that one digest (59.5%) and the shared reopen
(25.5%). Planning keeps the first hash of each archive and the candidate
serialization digest.

## For `REPORT.md`

## 0751 — digest reuse in the owned PPTX cross-copy

[0751](0751-pptx-cross-copy-apply-digest-reuse.md) adds to `litchi-opc`, all
additive:

- `OpcPackage::exact_source_sha256`, a digest memo created with the retained
  owned archive by one constructor, filled once and shared by clones;
- `OpcPackage::exact_source_len`;
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
  none;
- every facade adoption keeps the snapshot's memo re-projected onto the
  facade's own allocations. The review found that adopting it as it was could
  keep a caller-defined part's cloned payload alive, also on the base's
  cross-copy publications.

Every `LPCP0004` byte, recorded revision, published byte and refusal equals
the base's. A 196-line golden transcript printed on the base pins them,
including caller-defined destinations and size-limit refusals.

Median process p50:

- media-rich lifecycle: 184.2 → 109.8 ms;
- media-rich: 159.6 → 80.6 ms;
- plain: −3.3% and −4.1%;
- source-backed control: unchanged.

Allocated bytes fall 12.3% and peak live bytes 5.4%. The tiny and medium
semantic controls move +1.1% and +0.8% at p50: the re-projection adds 0.17%
instructions per iteration there, and code layout adds cycles.
`performance_claim: none`.

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

  It also accepts a small cost on the semantic controls to meet the ADR 0005
  memo amendment's re-projection clause literally at every facade adoption.
- **Bound memos.** Every skipped hash is answered by a memo bound to its bytes
  by allocation identity, or by a single-constructor digest cell on an
  immutable archive, and every read is re-derived in debug builds. Every
  freshness, revision, physical-provenance, budget and refusal check still
  runs.
- **ADR 0005.** Filling an empty facade slot from a capture follows the
  coordinator's ruling in 0751's review. The ADR text is unchanged, and the
  reading is listed for the owner's confirmation.

The non-iWork goal stays open.
