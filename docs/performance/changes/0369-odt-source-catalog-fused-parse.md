# Change 0369: ODT source-backed catalog fused parse

## Scope

The source-backed ODT catalog open path previously validated `content.xml` and
then scanned it again to classify text blocks. Change 0369 replaces those
sequential passes with one borrowing XML pass. Each XML event is delivered to
the existing validation state and the private text-block-kind handler while
the reader advances once through the source.

The handler is private to the ODT owner and adds no public API, dependency
edge, archive handle, runtime handle, lock, unsafe storage, or semantic model
change. Styles, media, and semantic payloads remain cold in this open path;
the timed operation only admits the existing package metadata and
`content.xml` work needed for the catalog.

## Correctness and resource boundaries

The fused pass preserves the sequential error contract. XML/tokenizer errors,
validation errors, and validation/tokenizer finish errors retain precedence
over a deferred text-kind scan error, matching the former validate-then-scan
order. Source freshness, ZIP verification, cancellation, and other outer
fences still complete before the catalog is published. The existing 256 MiB
`content.xml` limit, 1,000,000-block execution ceiling, and 4,096 nesting-depth
ceiling remain in force.

Correctness evidence is an exact rustfmt check, `cargo test -p litchi-odt
--lib --tests` with 557 library tests and every integration target passing
(926 tests total), and scoped Clippy with `-D warnings`. An independent code
and resource review accepted the private handler, ownership boundary, error
precedence, and bounded-execution behavior.

## Clean ABBA evidence

The clean CPU-2 ABBA used 30 warmups and 500 samples per leg over the
deterministic large media-rich ODT corpus, whose archive SHA-256 is
`d63726138d0a50c8ff7e150af4a86385df1a34d886bb5f61f985c78ac79b0220`.
Control was revision `bf1cb55c6` with binary `a7991b...`; candidate was
revision `b712aafbf20e` with binary `1a75eb...`. Confidence intervals did not
overlap, and same-side drifts remained below 15%.

| ABBA legs | p50 reduction | mean reduction | p95 reduction | p99 reduction |
| --- | ---: | ---: | ---: | ---: |
| A1/B1 | 53.560% | 53.116% | 51.008% | 49.508% |
| A2/B2 | 56.320% | 56.078% | 54.542% | 54.304% |

The retained machine-readable evidence is [the clean ABBA result](../results/odt-source-catalog-0369-abba.json).

## Claim boundary and decision

The accepted claim is limited to deterministic large media-rich ODT corpus
latency for a fresh `SourceBackedDocumentCatalog::from_read_at` open. The
`catalog()` list projection is not a claim: its observed tens-of-nanoseconds
change was unstable. The selected-block query is also not a claim: its
2.9%-3.4% change is below materiality.

No claim is made for all ODT inputs, RSS, allocation volume, fixed memory,
physical I/O, cold-cache behavior, throughput, or OOM prevention. The change
is accepted only for the stated open operation and corpus boundary.
