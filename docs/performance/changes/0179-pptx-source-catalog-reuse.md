# Change 0179: retain the source-backed PPTX edit catalog

Date: 2026-08-17

## Decision

Retain the presentation root, package and presentation relationship bindings,
and ordered slide graph already validated when
`SourceBackedPresentationEditor` opens. Slide capture and publication now reuse
that immutable catalog instead of reparsing the retained presentation root and
rebuilding all slide bindings.

The editor remains non-cloneable and publication still consumes it. Each slide
capture checks cancellation, resolves and validates the selected Part, parses
its scene, checks MCE branch selection, and reads the current source version.
Patch application still compares exact lineage, version, presentation and
slide closure. The OPC publisher still enforces signatures, framing, limits,
source freshness, complete raw-member planning, and truthful partial-sink
progress. Cross-presentation slide copy retains its independent complete graph
recapture and validation.

## Deterministic work reduction

The existing media-rich corpus has 200 slides. Before this change:

- one shape-text edit built the source catalog at editor open, slide capture,
  and publication: `3 -> 1`, removing two complete graph builds and exactly
  400 `Arc<SourceSlideData>` allocations;
- the eight-shape same-slide batch had the same `3 -> 1` reduction; and
- the eight-slide batch built the catalog at open, once for each selected
  slide, and once at publication: `10 -> 1`, removing nine complete graph
  builds and exactly 1,800 slide-node allocations.

Each removed build also eliminates one presentation-reference parse, one
reserved 200-entry slide vector, and cloned package, presentation, and slide
relationship strings. This is a semantic metadata-work result. Payload
materializations remain exactly 2, 2, and 9; logical source calls/bytes and
sequential sink topology are unchanged. No physical-I/O or total allocation
claim follows from the source-level count.

A test-only thread-local counter proves exactly one `source_catalog` call over
the complete open, edit, and consuming publication lifecycle. Separate
content-free cache assertions prove slide capture no longer hits the retained
presentation root. The counter is absent from production builds.

## Verification

- the final catalog-count regression and 17 source-backed edit integration
  tests pass;
- the complete PPTX all-target suite passes, including source-backed
  cross-copy, signed/no-op, MCE, stale/foreign, cancellation, raw ZIP,
  inverse, and partial-sink coverage;
- the two existing performance-harness semantic tests for one-slide and
  eight-slide source publication pass;
- all-target PPTX Clippy passes with warnings and deprecations denied;
- PPTX rustdoc, formatting, diff, and the workspace crate-boundary checker
  pass; and
- two independent final reviewers confirm catalog ownership, managed
  `PartData`, source-version/lineage, graph, publication, and deterministic
  work accounting remain safe.

Two unrelated PPTX test expressions were migrated from `get(0)` to `first()`
and from a constant `ok_or_else` closure to `ok_or`, allowing the strict
all-target warning/deprecation gate to complete without changing behavior.
The Unix pathname-replacement regression was also made deterministic: it now
replaces the path, then mutates the original inode through a retained hard link
so the source-version assertion does not depend on timestamp granularity.
`FileSource` production behavior is unchanged.

## Clean release A/B/B/A

The raw control legs record revision `a4f0d884c`; their release binary was
built at `1535b141e`, with the intervening range changing only checked-in
performance documentation/results and no production or harness source. The
measured candidate revision is `a43258127`. Its post-measurement follow-up
`c3bff4af1` changes only the Unix source-version regression described above;
it does not alter a release build. The distinct control/candidate binaries
have SHA-256 `af7a105d9f...` and `4cce0decc5...`. Every leg is clean, pinned to
CPU 2, exposes one logical CPU, and records 20 warmups plus 500 samples for
three existing source-backed PPTX selectors. The canonical non-timing
projection SHA-256 is `490f3d10ce...` in all four legs, including corpus,
output, logical source counters, materializations, and sink topology.

Positive paired values mean lower candidate p50 lifecycle latency:

| Workload | A1 -> B1 | B2 -> A2 | Control p50 drift | Candidate p50 drift | Decision |
|---|---:|---:|---:|---:|---|
| one shape text | 25.48% | -34.21% | 3.93% | 73.03% | latency withheld |
| eight shapes, one slide | -21.28% | 8.04% | 39.59% | 5.84% | latency withheld |
| eight slides | -12.66% | 15.40% | 42.38% | 6.92% | latency withheld |

The paired directions disagree for every workload. Required 5%/5%/10%/15%
p50/mean/p95/p99 same-implementation stability gates also fail. The exact
catalog-build and slide-node-allocation reductions are retained, but no
latency, throughput, allocation/RSS, physical-I/O, cold-cache, scaling, or
real-producer claim is accepted.

## Scope

This adds no selector and leaves the harness at 320 cases and the historical
default at 36 cases / 198 records. It does not broaden PPTX CRUD, add topology
or relationship mutation, weaken signed or MCE refusal, change Part
materialization, or affect DOCX, XLSX, OLE2, RTF, ODF, or iWork.

Artifacts:

- [summary](../results/pptx-catalog-reuse-0179-summary.json)
- [manifest](../results/pptx-catalog-reuse-0179-manifest.json)
- raw A1/B1/B2/A2 reports listed in the manifest
