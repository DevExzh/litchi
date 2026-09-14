# 0571: one archive index per detect-and-open

Status: retained. `performance_claim: none` — this record carries deterministic
construction and read counts plus paired microbenchmark medians, not a registry
claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

This is **two changes in one diff**, and they should be read separately. The
prepared source is a performance change with measured evidence. The typed
wrong-format error is an error-model correction that reuses the classification
the first change stops discarding; it makes no performance claim and is not
measured as one.

## What was removed

Change [0567](0567-ooxml-single-index-per-open.md) established that one
library-level OOXML open builds exactly one ZIP archive index, and that repeated
indexing is real only **across** two calls: detection builds a package,
classifies it, throws it away, and the open builds another. Change
[0569](0569-ooxml-detect-then-open-priced.md) priced that at 42.2, 219.7 and 49.3
microseconds for a document, presentation and workbook, and **rejected** change
0567's proposed documentation fix, because the pattern is unavoidable for any
caller that must dispatch across the three facade types and this repository's own
flagship example is exactly such a caller.

This change makes the unavoidable pattern cheap. A new opaque handle carries the
already-built package from detection to the opener, so the index is built once.

## Measured effect

**Archive constructions**, counted through the facade's existing probe counter:

| Fixture | detect then open | detect-and-prepare then adopt |
| --- | ---: | ---: |
| a document | 2 | **1** |
| a presentation | 2 | **1** |
| a workbook | 2 | **1** |

**Reads on the package**, traced:

| Fixture | pattern | opens of the file | reads |
| --- | --- | ---: | ---: |
| document, 4,926 B | two-call | 2 | 47 |
| document | single-call | **1** | **24** |
| presentation, 68,822 B | two-call | 2 | 132 |
| presentation | single-call | **1** | **59** |
| workbook, 9,309 B | two-call | 2 | 44 |
| workbook | single-call | **1** | **25** |

**Wall clock**, paired medians over 3,000 iterations, warm cache, with a second
independent run in parentheses:

| Pattern | two-call | single-call | saving | share |
| --- | ---: | ---: | ---: | ---: |
| document | 73.89 (74.48) us | 37.95 (38.05) us | **35.94 (36.43) us** | 48.6% |
| presentation | 434.80 (432.89) us | 233.06 (231.12) us | **201.74 (201.77) us** | 46.4% |
| workbook | 107.00 (106.49) us | 64.98 (64.71) us | **42.02 (41.78) us** | 39.3% |

The **shares** reproduce change 0569's independently measured 48.6%, 46.3% and
38.6% closely; the absolute microseconds differ because the host and fixtures
differ, so the shares are the comparable quantity.

### The repository's own example

`crates/litchi/examples/to_markdown.rs` was the one place that taught the
two-call pattern, and change 0569 noted it had no correct rewrite under the
rejected advice. It now uses the prepared path, falling back to a plain open for
the formats that have no handle. Measured over the full conversion workload:

| Fixture | before | after | saving | share |
| --- | ---: | ---: | ---: | ---: |
| a presentation | 1,254.09 us | 1,026.74 us | **227.34 us** | 18.1% |
| a document | 887.19 us | 839.26 us | 47.93 us | 5.4% |

The presentation figure is stable across runs. The document figure is noisy
because Markdown conversion dominates a 5 KB input, and should be read as "tens
of microseconds" rather than as a precise number.

### What it costs where there is nothing to prepare

A non-OOXML input gets its format and no handle, at the price of one extra file
open plus a discarded probe: **+4.83 microseconds** for an OpenDocument text file
and **+1.38** for a legacy OLE2 document.

That cost is deliberate. Every path that cannot prepare defers to
`detect_file_format`, so the classification can never diverge from the neutral
detector. The alternative — answering from the container probe directly — would
have saved most of those microseconds on a deferred format family at the price of
that guarantee. `detect_file_format` itself is untouched, so no existing caller
pays anything.

## The second change: the classification stops being discarded

Change 0569 measured the consequence of throwing the classification away: the
workbook opener returned an identical unit "not an Office file" for a document, a
legacy document, an OpenDocument text file, a presentation, an image and a text
file. It could not distinguish "this is a Word document" from "this is not an
Office file", where detection separates them cleanly.

A typed variant now carries the detected format, and the ten sites that read
`let _ = format;` report it. ADR 0010's decision section already assigns the
coordinator the job of mapping content types to the neutral classification, so
surfacing it stays inside the boundary. The format type lives in the same crate
as the error and the enum is non-exhaustive, so this needed no new home and is
not a breaking change.

**This is not a performance fix.** Reaching the error still costs a failed open.

## ADR compliance

ADR 0010's 2026-08-08 amendment fixes the accepted shape as one classified
immutable package snapshot consumed by exactly one adapter, with an existing
prepared-source type as precedent. The handle owns one package and yields it
through a **consuming** accessor, and is not cloneable, so a second adapter
cannot receive the same snapshot. It stores the **content type**, never a format
classification, and the content-type-to-format mapping stays in the facade in one
table read by both the catalog scan and the classifier, so a handle cannot
disagree with the classification that produced it.

The handle differs from its precedent in one respect deliberately: that one
materializes eagerly and has no post-preparation version check, whereas an OOXML
source-backed package stays lazy, so this one retains the revision it was
prepared at and re-checks it on adoption.

ADR 0011 is satisfied because the type lives in the crate that owns the
translation to the ZIP implementation and reuses its existing constructor. No new
parser, no new ZIP code, no new dependency, and the ZIP crate is **unchanged**.
Constructing the handle does no I/O: it validates the content type by walking the
package's already-parsed in-memory catalog. Crate boundaries and dependency
direction pass.

## Correctness evidence

Thirty tests were added and all fail against the pre-change code. Seven cover the
handle itself, including that a source changed between preparation and adoption
is refused, driven through a revision-switching double rather than simulated
through the filesystem. Seven count constructions and pin that a non-OOXML input
is still classified with no handle and agrees with the neutral detector. Sixteen
cover the typed error across all three openers, including one that drives five
wrong formats through the workbook opener and asserts they are **mutually
distinguishable**, which is the exact defect change 0569 measured.

The author did not rely on "these tests do not compile at HEAD" as the
pre-change argument, and instead wrote a throwaway probe asserting the old
behaviour against the new tree, confirming all three old assertions now fail.

Four existing tests that asserted the untyped behaviour were updated, intent
preserved and assertion tightened to name the format.

The facade and package crates pass **4,335 tests with zero failures**, and the
facade suite keeps exactly its four pre-existing unrelated failures.

## A pre-existing gate failure found and fixed

The Clippy gate **fails at HEAD** for `litchi-opc` and the facade, on five sites
unrelated to this work: two unusual byte groupings, an integer comparison, a
needless return flagged across configuration-alternated blocks, and a field
reassignment after default. This was invisible to the standing gate list, which
runs Clippy only on the container and spreadsheet crates. The five are repaired
here so the gate can report a result for this change at all, and they are called
out as separable rather than absorbed silently.

This is the second gate blind spot this program has found in two batches; the
first was a test runner aborting at the first failing target and hiding every
later one.

## Judgement calls recorded for review

An OOXML family whose owner is compiled out previously returned the untyped
error, with a comment stating that was deliberate parity with another detector.
It now reports the detected format. The neutral detector already names that
format in such a build, because the content-type table is not feature-gated, so
this improves parity rather than breaking it. It is easy to revert if the
project disagrees.

One opener that is deliberately single-format still returns the untyped error for
three of its arms; converting it is a wider decision that was left alone. Its
two arms matching this defect were converted.

## Limitations

Warm cache, one host, one fixture per format, release profile, paired medians.
No cold-cache, allocation, peak-RSS, throughput or multi-fixture-corpus result is
claimed. The construction and read counts are exact; the microsecond figures are
microbenchmark medians, not profile attributions. The new entry point is feature-
and platform-gated to builds with an OOXML owner and a positional source; other
targets keep the existing detector.
