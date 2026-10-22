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

## Remaining bounded-streaming follow-up

Cross-owner replay whose `before` or `after` endpoint is already deferred or
effective still needs an OPC streaming artifact publisher. The current
opened-source restore path is bounded, but a deferred/effective endpoint must
preflight its streamed length and digest without materializing a full archive
at apply time, then publish under the recipient context. This remains an
engineering follow-up coordinated with the OPC contract; it is not accepted
as a source-backed exception.

## Validation

This source-only slice was reviewed with `git diff --check` and rustfmt
1.95.0. Cargo and test commands were intentionally not run while the
coordinator's host hold was active. The focused tests are queued for the next
capped validation wave; no performance claim is made here.
