# Change 0394: OPC case-fold lookup baseline

Date: 2026-09-03

Status: harness-only, opt-in evidence coverage. No production code or
existing result artifact changes are part of this batch.

The four selectors are:

```text
opc_casefold_eager_open
opc_casefold_source_open
opc_casefold_eager_lookup
opc_casefold_source_lookup
```

## Corpus and oracle

The harness generates the same deterministic stored OPC archive at exactly
256, 2,048, and 16,384 ordinary parts. Each part is a stable
`casefold/ordinary-*.bin` member with a 32-byte payload. `[Content_Types].xml`
and `_rels/.rels` are the only structural members. Every result records the
generator ID, archive/member count and SHA-256, payload SHA-256, and sorted
canonical-name oracle SHA-256. All three corpora are below the ordinary OPC
read limits.

The lookup vector is built before any timed or allocator-scoped operation and
is fixed at nine queries in this order:

1. exact first, middle, and last;
2. case-only aliases for first, middle, and last, exercising the bounded
   case-insensitive scan;
3. genuine first, middle, and last misses.

Each vector is repeated 16 times in the same order. The report includes the
query classes, canonical corpus positions (not implementation traversal
positions), expected-found flags, total lookup count,
canonical-name oracle, and a deterministic output digest. Malformed equivalent
part names are generated separately by
`litchi-opc-casefold-equivalent-name-gate-v1` and are an untimed correctness
gate for both implementations; each must return `EquivalentPartNames`.

## Timing boundary and metrics

The eager and source-backed open selectors time only their respective package
constructors. Source-backed elapsed and allocator samples use the immutable
production `litchi_core::OwnedSource` adapter. The repeated-lookup selectors
construct and validate the package once outside timing, then time only the
fixed query loop. Query construction, archive cloning/source-wrapper setup,
output hashing, and correctness checks are outside the timed lookup phase.

Source-backed results report exact operation-scoped `InstrumentedSource`
read-call, read-byte, source-version, ordinary-payload, and maximum-in-flight
counter vectors from an independent untimed replay. The replay must reproduce
the timed operation's semantic digest. This keeps range classification and
counter atomics out of elapsed and allocator samples. Eager results mark
source counters not applicable. The operation-metrics envelope includes
allocator observations only when the allocator-enabled benchmark target
supplies them; unavailable observations are omitted rather than inferred.

`CaseResult` now stores its large tagged `SourceSummary` behind one report-
assembly box. This is serialization-transparent and happens after the timed
and allocator-scoped sample operations. Without the indirection, adding the
new evidence record made the unoptimized `run` frame exceed the normal 8 MiB
main-thread stack, so even `--help` aborted before option parsing. The normal-
stack help path and the complete 12-row one-sample CLI now pass; no larger
stack environment is required.

`performance_claim: none`. This batch does not recommend an index and does
not change the default 36-case / 198-record matrix. The four selectors raise
the selectable registry from 411 to 415.
