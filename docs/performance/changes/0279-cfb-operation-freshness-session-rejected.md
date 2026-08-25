# Change 0279: rejected CFB operation-scoped freshness session

**Date:** 2026-08-25
**Status:** Rejected and reverted
**Performance claim:** none

## Decision

The operation-scoped freshness-session candidate was reverted because its
strict ABBA evidence failed the predeclared, unqualified same-side drift limit.
The candidate produced a large and repeatable direct reduction, but four
same-side p95/p99 comparisons exceeded 5%. Narrowing the gate to central
statistics after collection would be retrospective.

Production commit 0b0e1ed02c2fe4a262826a25a1f64b533bd0e686 was reverted by
e198d6048. The repository therefore retains the prior per-read freshness
behavior.

## Candidate under test

The candidate introduced a low-level, closure-scoped CFB stream session and
used it for XLS global-header preflight and selected-sheet scans. The intended
contract was:

- fence once after cursor construction;
- omit per-read fences after successful reads;
- immediately fence raw read failures;
- final-fence every completed operation;
- final-fence aborts after a successful read;
- permit cancellation before source I/O to abort without a final fence;
- prevent session state or borrowed output from escaping the closure.

The exact global-range read remained a separate before/after-fenced source
operation. Physical reads were not coalesced. The candidate intentionally
changed freshness from per-read publication boundaries to operation-scoped
boundaries, so its public contract and error precedence were reviewed before
measurement.

## Correctness evidence before measurement

The candidate passed:

- 271 litchi-cfb library tests;
- 1,309 listed litchi-xls all-feature library/integration tests;
- the no-default-feature XLS library check;
- strict Clippy for CFB and for XLS library plus source_backed;
- rustdoc with -D warnings;
- formatting;
- the crate-boundary policy gate;
- architecture, freshness-precedence, cancellation, locality, and public-layer
  review with no final P0-P2 finding.

Tests covered start/final staleness, raw read-error precedence, stable and
mutated read-then-abort, cancellation before and after reads, FILEPASS,
unknown-payload no-read behavior, malformed tails, STRING/CONTINUE,
duplicate-last semantics, selected-sheet bounds, limits, and exact version
probe counts.

## Strict ABBA protocol

The corpus was test-data/ole/xls/ConditionalFormattingSamples.xls, 1,402,368
bytes, SHA-256
d1942d857ffbd4d10ebca1745cd5d70c14af9d9f1388c91ed0a0800e31ad5ce7.
The stable cell oracle was worksheet index 1, row 1, column 0: string:4:Date.

The control was 9adbf9ee527b4ccd2e3494c824c761973633ea14; the candidate was
0b0e1ed02c2fe4a262826a25a1f64b533bd0e686. Clean detached release binaries
were pinned by size and SHA-256 in protocol.json.

Collection used CPU 2, one sequential worker, 20 warmups in every fresh child,
500 retained samples per cell per leg, and fixed order A1/B1/B2/A2. The six
cells were file-source and atomic-file crossed with open, list, and one-cell,
producing 12,000 retained samples. All children passed schema, identity,
counter, source-stability, and semantic validation; failures.log is empty.

Statistics use a floored even-sample midpoint for p50 and nearest-rank p95/p99.
The retained JSON preserves the arithmetic mean; tables floor it for
presentation.

## Descriptive result

Positive percentages mean the candidate was faster.

| Mode/operation | A1 to B1 p50/mean/p95/p99 | A2 to B2 p50/mean/p95/p99 |
|---|---:|---:|
| FileSource open | 52.7415% / 52.2447% / 51.5090% / 49.8418% | 52.7303% / 52.1908% / 48.4374% / 48.5467% |
| FileSource list | 52.5410% / 52.2619% / 49.9129% / 50.5892% | 52.6057% / 52.3932% / 49.9067% / 50.4328% |
| FileSource one-cell | 56.3126% / 55.7649% / 52.0528% / 53.0828% | 56.3109% / 56.1157% / 54.8179% / 55.5605% |
| AtomicFile open | 21.5471% / 21.9640% / 19.2726% / 22.9094% | 22.2377% / 21.8261% / 19.1576% / 19.8689% |
| AtomicFile list | 22.4527% / 20.9647% / 18.6103% / 20.9679% | 22.0094% / 22.3190% / 20.1159% / 22.6092% |
| AtomicFile one-cell | 25.5484% / 24.9136% / 20.5016% / 21.0180% | 25.0879% / 23.9751% / 22.6039% / 25.6315% |

Version calls changed exactly from 1,266 to 26 for open/list and from 1,802 to
34 for one-cell. This is a 97.9463% and 98.1132% reduction, respectively.
Logical reads, read bytes, length calls, locality, identities, source stability,
and semantic results were exact-neutral. The atomic mode improves because it
also pays for version probes; it is an attribution guard, not a general
atomic-path claim.

## Failed keep gate

Four same-side tail comparisons exceeded the declared 5% limit:

| Mode/operation | Side/statistic | Drift |
|---|---|---:|
| FileSource open | control p95 | +5.0369% |
| FileSource one-cell | control p99 | -6.2960% |
| AtomicFile open | control p99 | +5.0963% |
| AtomicFile one-cell | candidate p99 | +5.0746% |

All FileSource central comparisons and all direct candidate tails improved.
Nevertheless, the same-side limit was stated without restricting it to p50 or
mean. The strict result is therefore rejected and descriptive only.

This run does not establish a production performance claim, a general
FileSource/XLS claim, an eager/facade claim, or a cross-family claim.

## Retained evidence

Evidence:
docs/performance/results/0279-cfb-operation-freshness-session-rejected-20260825/

The package retains the predeclared protocol, 12,000 normalized samples,
aggregate statistics, comparisons, validation verdict, 24 representative raw
reports, empty failure log, and artifact manifest.

Manifest SHA-256: 21a0842e85c27beb26981779092b488fd79473e32a32a51b985219dd5d055a9e
Manifest bytes: 5635
