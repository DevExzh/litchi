# Change 0182: fuse PPTX validation catalog and graph traversal

Date: 2026-08-18

## Decision

Retain the fused source-backed PPTX validation traversal. The validator used to
walk the immutable OPC relationship catalog once for external, signature, and
macro presence, then walk the same package and Part relationship lists again
to validate internal graph targets. It now collects catalog facts and graph
facts in one ordered pass.

The graph still visits package relationships before Parts, registers each Part
before its relationships, resolves every internal target through the same
`SourceBackedPackage`, and stops at the first graph-node limit. Catalog counting
continues after that stop so later external, signature, and macro facts remain
complete. No target lookup or graph missing/invalid count occurs after the
limit. A regression covers that exact boundary and asserts exact catalog
evidence counts.

## Deterministic work reduction

The relationship-list traversal count changes as follows:

- package relationships: `2 -> 1` passes;
- each Part's relationships: `4 -> 1` passes; and
- for `N` Parts, one complete package pass plus `3N` Part-list passes are
  removed.

Presentation and slide XML parsing, graph target lookups before the limit,
logical source reads/bytes, ordinary payload materialization, source-version
fences, report checks/issues, and the public API remain unchanged. This is a
metadata CPU-work reduction, not a physical-I/O, decompression, allocation, or
memory result.

The control profile attributes 20.22% of the large validation run to
PresentationML XML inspection, with quick-XML namespace and attribute work
dominating the next entries. The small relationship loops are inlined below
the visible profile threshold, so the measured large-corpus result is scoped
to the complete validator rather than described as an isolated catalog
benchmark.

## Verification

- the focused validation integration suite passes 12/12;
- the complete PPTX all-target suite passes, including 520 library tests and
  every integration/example target;
- all-target PPTX Clippy passes with warnings and deprecations denied;
- PPTX rustdoc passes with warnings denied;
- formatting, diff, and the 64-package crate-boundary check pass; and
- independent architecture and performance reviewers found the fused ordering,
  graph-limit behavior, catalog counts, freshness, and work accounting safe.

## Clean release A/B/B/A

The frozen control binary was built from `48122377c`; the candidate was built
from clean production revision `e11f06d2b`. Their SHA-256 values are
`6d34aa9926...` and `cf89cac195...`. Both were executed from the clean candidate
checkout, so the schema-1 harness's runtime Git field is `e11f06d2b` in all
four raw reports; the manifest and summary separately bind the frozen control
source revision and binary digest.

Fresh processes ran `A1 control, B1 candidate, B2 candidate, A2 control`, pinned
to CPU 2 with one affinity-visible logical CPU. Each existing
`pptx_validation_report` shape retained 20 warmups and 500 samples. The
canonical projection excluding only `elapsed_ns` has SHA-256 `5b3d2fdc95...`
in all four legs. Report SHA-256, check IDs/statuses, issues, corpus/source
hashes, logical source counters, zero ordinary-payload reads, and maximum
in-flight reads are exact across the comparison.

Positive values mean lower candidate latency:

| Shape | Metric | A1 -> B1 | B2 -> A2 | Control drift | Candidate drift | Decision |
|---|---|---:|---:|---:|---:|---|
| tiny | p50 | 1.39% | 10.83% | 6.01% | 4.14% | latency withheld |
| medium | p50 | 1.22% | 6.19% | 8.57% | 3.11% | latency withheld |
| large | p50 | 11.50% | 7.08% | 3.65% | 1.17% | accepted for this corpus |
| large | mean | 11.49% | 6.56% | 3.31% | 2.08% | accepted for this corpus |
| large | p95 | 11.18% | 6.45% | 1.77% | 7.19% | accepted for this corpus |
| large | p99 | 11.62% | 4.87% | 3.44% | 3.93% | accepted for this corpus |

Tiny control p50/mean drift exceeds the 5% thresholds. Medium control p50/mean
also fails and its mean/p95 paired directions disagree. Those shapes retain
only the deterministic traversal reduction. The large shape passes the
predeclared 5%/5%/10%/15% p50/mean/p95/p99 stability gates for both
implementations, and every paired distribution direction agrees. Complete
bounded validation p50 is therefore accepted as 7.08%-11.50% lower for the
deterministic large semantic PPTX corpus.

No latency claim extends to tiny/medium, physical or cold I/O, allocation/RSS,
total memory, throughput/scaling, real-producer files, or PPTX mutation and
publication.

## Scope

This changes no public API, selector, corpus, or default. The harness remains
at 322 selectable cases and the historical default remains 36 cases / 198
records. It does not affect DOCX, XLSX, ODF, OLE2, RTF, or iWork.

Artifacts:

- [summary](../results/pptx-validation-fusion-0182-summary.json)
- [manifest](../results/pptx-validation-fusion-0182-manifest.json)
- raw A1/B1/B2/A2 reports listed in the manifest
