# 0501 design review: PPTX touched-payload digest

This review covers the proposed removal of only the decoded image and chart
payload bytes from the private `digest_touched` input in
`crates/litchi-pptx/src/presentation/source_cross_copy.rs`. It does not approve
removing payload equality, graph proof, source freshness checks, or OPC
publication authorization.

## Evidence and measured scope

The historical 0448 profile reported SHA-256 as 63.2--63.9% of weighted
self-period in the whole lifecycle-frame subset. That is a broad hotness
observation: it includes untimed harness work and other SHA callers. The 0449
caller attribution separates the production touched digest as follows:

| caller | separate sleeps | minimum service |
| --- | ---: | ---: |
| planning `digest_touched` (% lifecycle-frame period) | 15.879% | 15.510% |
| publication `digest_touched` (% lifecycle-frame period) | 16.115% | 15.930% |

The same attribution assigns 49.762%/50.088% of lifecycle-frame SHA period to
the untimed harness output hash. Therefore the 63% result is not a timer-local
production target or an expected speedup for this change. The 0501 matched
matrix must establish the candidate result independently.

## Proof equivalence

`Prepared::matches` compares `images` and `charts` before it compares
`touched_digest`. `PreparedImage` and `PreparedChart` equality compares source
URI, target URI, content type, declared size, and the exact decoded byte slice
through `PreparedPayload::as_slice()`. The same exact-byte condition is also
checked by `reuse_or_clone_payload` before a retained payload is reused.

`verify_candidate` then rereads each selected image and chart, checks declared
size, content type, relationship closure, and exact decoded bytes, and validates
chart XML. Publication retains the existing source lineage/revision checks,
source-monitoring writer, and OPC precompressed authorization. Those checks
remain the source and payload authority.

Consequently, removing only these two digest updates is acceptance-equivalent
under the existing source contract:

```text
digest_bytes(image.bytes.as_slice(), ...)
digest_bytes(chart.bytes.as_slice(), ...)
```

The implementation should retain the declared-size and actual-length updates,
all payload metadata, all XML byte inputs, relationship order, names, URIs,
and the graph digest. In particular, the source layout URI is represented in
the private digest rather than as a separate `Prepared` field, and the
layout/master/theme graph bytes are represented by `graph_digest`; neither
may be removed by this change. The source and destination presentation bytes
should also remain in the digest because they are not each independently
retained as exact `Prepared` fields.

The digest is private and is not a public fingerprint or authorization token.
The plan and rerun both use the new digest definition, so changing its private
domain does not create a compatibility obligation. A payload difference that
leaves the new digest unchanged still fails the independent exact vector
comparison.

## Cancellation and resource accounting

The existing payload hash walks 64 KiB chunks and calls `check_execution` for
each chunk. The proposed lightweight replacement is acceptable if it retains
the same chunk boundaries and check placement without allocating or consuming
`Resource::Work`:

```text
for chunk in payload.chunks(64 * 1024) {
    check_execution(execution_context)?;
}
```

The helper should skip the loop when there is no execution context, since an
unmanaged operation has no cancellation observer. This preserves the managed
cancel observation cadence while removing SHA compression work. The old digest
path did not charge `Resource::Work`; the helper must not introduce a new
charge or counter delta. Source reads, cache admission, memory reservations,
object reservations, and final release behavior are unchanged.

Publication already has checked chunked equality or checked cloning in the
payload-reuse path. A cancellation test should nevertheless cover a large
warm-cache image and chart so the initial planning path does not accidentally
lose its only per-payload observation points.

Replacing exact equality with a memoized SHA-256 would not be safer. A
memoized digest is acceptable only as a cache or diagnostic while exact bytes
remain mandatory; using it as the authorization predicate would violate the
exact-source rule because a digest collision is not byte equality. Carrying a
digest per payload also adds 32 bytes and managed-accounting/invalidation
complexity without improving the proof. Directly removing the redundant hash
input is the smaller change.

## Required adversarial coverage

The focused suite should retain or add these cases:

1. Change one byte of an image without changing its length; `Prepared::matches`
   or publication must refuse the stale plan.
2. Do the same independently for a chart, including a mixed image/chart slide.
3. Change declared size, actual length, source URI, target URI, or content type
   while keeping payload bytes equal; the plan must refuse.
4. Exercise shared media/chart leaves and retained compressed captures; capture
   reuse must remain an optimization only, with exact decoded output unchanged.
5. Retarget the source slide's layout relationship to an equivalent-looking
   layout URI and mutate one layout/master/theme byte. The retained source
   layout URI and graph digest must continue to reject stale plans.
6. Cancel during planning and publication with a large retained media/chart
   payload, and assert the typed cancellation result, no invalid publication,
   and zero remaining managed memory/object reservations after drops.
7. Keep the existing source-version mutation and mid-publication source-change
   tests. A deliberately dishonest adapter that changes bytes without changing
   `SourceVersion` is outside the `ReadAt` contract, but it is useful for
   proving that exact payload comparison remains defensive where the cache
   exposes the changed bytes.
8. Add a private unit case that holds `touched_digest` equal while changing one
   payload byte, proving that digest equality alone cannot authorize a match.
   No real SHA collision is needed or desirable.

The after report should compare the matched plain/media-rich owned and warm-file
lanes, preserve every semantic/package/source/cache/budget oracle, and retain
all latency or RSS flags above five percent. It may report the touched-digest
change as a scoped source-backed PPTX result only; it must not promote the
whole 63% SHA observation into a causal or repository-wide performance claim.
