# Change 0180: cache repeated source-backed ODT text projections

Date: 2026-08-17

## Decision

Retain a bounded full-text projection once the two-call query threshold has
been reached and a complete `SourceBackedDocument::text()` parse succeeds. The
first invocation keeps the established validating parser path. A later
threshold-reaching call parses again, proves the source is still current, and
fallibly publishes at most one 16 MiB `String`. Failed parses advance the
saturating call counter but remain retryable and never publish. Later cache
hits check source freshness, clone the retained text into the existing owned
return type, and check freshness again.

The optional cache is local to the immutable source-backed document. It adds no
public type, API, executor, lock, unsafe code, global state, or dependency edge.
Oversized results and retained-copy allocation refusal install a terminal
non-retaining state; parse errors remain retryable. Concurrent construction is
safe but intentionally not single-flight: callers racing before publication
may each perform the bounded construction parse. Every successful caller still
receives a distinct owned `String`.

## Deterministic work reduction

The matched workload prepares one 10,000-paragraph source-backed ODT and makes
four complete text projections. The uncached control performs four full
`content.xml` block-model projection phases. The retained path performs two.
Source preparation has already retained validated `content.xml`; every sample
proves zero post-preparation `ReadAt` calls, so this is parser/projection work,
not physical-I/O or decompression evidence.

The cache adds one bounded retained `String` and leaves four returned owned
strings. Source-version observations change from `[2, 2, 2, 2]` to
`[2, 4, 2, 2]`: the second call includes two additional checks around cache
publication. The content parser, namespace/depth/text limits, malformed-input
behavior, source-change precedence, archive topology, media payloads, and
immutable source contract remain unchanged.

## Two clean release A/B/B/A cycles

The control revision is `238184ff5`; the candidate is `3023f9cec`. Distinct
release binaries have SHA-256 `0211cffee8...` and `c198e16a03...`. Every leg
is clean, pinned to CPU 2, exposes one logical CPU, and records 20 warmups plus
500 samples for both matched selectors. The fixed 16,812,034-byte archive has
10,000 paragraphs, 13 members, and eight verified incompressible 2 MiB picture
members. Canonical text, archive, `content.xml`, picture-payload, and projection
digests match in all eight legs.

Positive paired values mean lower candidate latency for four complete matched
text projections; the candidate uses the public method while the control uses
its public-equivalent parser helper:

| Cycle | Metric | A -> B | B -> A | Control drift | Candidate drift | Decision |
|---|---|---:|---:|---:|---:|---|
| 1 | p50 | 50.95% | 47.07% | **5.25%** | 2.23% | accepted after balanced retry |
| 1 | mean | 51.29% | 46.83% | 4.22% | 4.56% | accepted after balanced retry |
| 1 | p95 | 52.15% | 46.71% | 1.12% | **10.13%** | withheld |
| 1 | p99 | 53.60% | 38.09% | 3.82% | **28.34%** | withheld |
| 2 | p50 | 47.01% | 47.71% | 1.24% | 2.54% | accepted |
| 2 | mean | 47.01% | 48.40% | 0.70% | 1.94% | accepted |
| 2 | p95 | 44.03% | 48.90% | 9.57% | 0.05% | retry passes, still withheld |
| 2 | p99 | 43.83% | 50.06% | 9.78% | 2.40% | retry passes, still withheld |

The first-cycle p50 control drift also narrowly triggered the retry. Across all
four paired directions, the repeated-text p50 reduction is 47.01%-50.95% and
mean reduction is 46.83%-51.29%; both directions repeat in the second cycle
inside their 5% stability thresholds. The unchanged uncached guard is noisy in
cycle 1 and near-neutral in cycle 2, consistent with the retry decision. p95
and p99 remain withheld because the original candidate tail gates failed; the
clean retry is retained rather than used to erase that observation.

## Verification and scope

- focused threshold, fresh-clone, concurrent-construction, parse-error retry,
  oversized fallback, stale-publication, and `Send + Sync` regressions;
- complete ODT all-target tests, strict Clippy with warnings and deprecations
  denied, rustdoc, formatting, and diff checks;
- focused harness test plus clean release semantic/source/archive/media gates;
- independent production and evidence-contract reviews.

This adds two opt-in selector names, raising the harness from 320 to 322 while
leaving the historical default at 36 cases / 198 records. It adds no CRUD
closure. No p95/p99, single-call/open, allocation/RSS, physical-I/O,
decompression/recompression/copy-volume, cold-cache, scaling, real-producer,
generic ODF, non-text projection, or broad ODT claim is made. Ordinary OOXML
eager ingress remains a separate design problem because its public infallible
borrowed Part payload contract cannot safely become deferred in place.

Artifacts:

- [summary](../results/odt-text-cache-0180-summary.json)
- [manifest](../results/odt-text-cache-0180-manifest.json)
- raw A1/B1/B2/A2 and retry A3/B3/B4/A4 reports listed in the manifest
