# 0362rn — DOCX source-backed restore recipient policy

## Scope

The source-backed family replay fast path can restore a complete artifact when
an independently opened recipient matches a patch's recorded `before` source.
The old path copied `after`'s semantic and physical policy into the returned
package. That allowed a broad producer's policy to widen a narrower recipient,
and skipped the family publisher's recipient-limit checks. Hyperlink detachment
uses the same generic path; the main-document owner path had the same bypass.

## Change

- `restored` now retains the recipient package's semantic limits and physical
  read limits, and rejects a restored artifact whose length exceeds the
  recipient input ceiling before installing deferred output.
- A restored owner is rebuilt even when either endpoint is already deferred or
  effective; cloning a producer's deferred owner would leak its execution
  context. First materialization of a managed deferred artifact now reopens
  the moved bytes with the retained recipient execution context, while the
  unmanaged path keeps the existing zero-copy owned-vector admission.
- The generic source-backed replay fast path and the main-document owner fast
  path require equal semantic policy with the recorded `before` package. A
  recipient with narrower semantic policy goes through `Change::publish`,
  which validates the selected family under the recipient's limits. Equal
  policies retain the authenticated streaming artifact shortcut.
- Same-owner replay now checks the source freshness fence first, then shares
  the recorded `after` owner only when the cache identity and semantic policy
  prove the recipient is the same physical, execution, effective-topology,
  and deferred-owner lineage. `with_execution` is the sole cache-identity
  minting boundary; clones and semantic-limit projections retain it. This
  avoids the prior same-owner hash/materialization cost without widening a
  recipient context.
- Story-text and hyperlink-detachment tests cover a broad producer replayed
  into a narrow independently opened recipient. The tests assert the typed
  recipient limit error rather than accepting the producer's policy.
- An internal regression test asserts that same-owner inverse replay leaves
  the deferred artifact unmaterialized. The content-free OPC fallback reason
  also no longer interpolates document-family error text.
- Every deferred publisher now accepts an optional `ArtifactStreamPolicy` and
  forwards it through part overlays, relationship-removal overlays, SVG
  topology publication, and paragraph/tail splice publication. The DOCX
  package implements the OPC artifact-stream adapter, so cross-owner restore
  can authenticate lazy physical, deferred, and effective endpoints without
  wrapping the output sink and charging producer output twice. Ordinary
  opened-source restore retains the existing `SourceArtifact` fast path.
- The DOCX artifact-stream adapter now declares `Error` as its producer error
  and returns it without a lossy OPC conversion. The OPC seam wraps failures
  as `ArtifactStreamError<docx::Error>`: embedded `OpcError` values retain
  cancellation, source-change and limit classification, DOCX semantic
  limits/refusals/patch errors remain inspectable producer causes, and
  publication records the accepted prefix plus any secondary failure. The
  DOCX error envelope boxes this recursive stream error while keeping its
  content-free display text.

## Remaining bounded-streaming follow-up

The generic OPC artifact-stream hooks are now available in the isolated OPC
source tree, but the final changed-plan census is still being completed on the
assembled branch. The DOCX preflight and policy path must prove recipient part
and aggregate-part bytes, part count, content-types bytes/mappings,
relationship part/XML/count/graph/event/depth/attribute ceilings, and ZIP
aggregate limits before accepting a deferred restore. Focused tests use
same-size equal copies with tight `max_parts` and part-byte ceilings; the
relationship, deferred/effective forward and inverse, output-limit,
freshness, cancellation, and exact-byte cases remain queued for the assembled
policy census. This remains an engineering follow-up, not an accepted
source-backed exception.

## Validation

This source-only slice was reviewed with `git diff --check` and rustfmt
1.95.0. Cargo and test commands were intentionally not run while the
coordinator's host hold was active. The focused tests are queued for the next
capped validation wave; no performance claim is made here.

## Assembly policy and conflict resolution

The assembled adapter retains the existing forward-only redaction patch types,
fixed typed conflict classification, and source/effective fingerprint memos.
Its native publication routes use the sealed target and do not materialize
deferred output during freshness checks or patch replay.

Materialization retains the native OPC execution context exactly, including
`None`; the enclosing DOCX owner retains the facade context independently.
Using `facade.or_else(native)` would discard a distinct native context or turn
an unmanaged native source into a managed one. Regression tests cover both
context combinations through an ordinary deferred edit, an opened restore,
and a restore whose target is deferred. These tests await assembled execution.
