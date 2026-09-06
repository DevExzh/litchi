# 0434: ODS bounded ordinary-text spans

Change 0434 compares the existing ODS streaming selector before and after a
private encoder change. Already-safe borrowed UTF-8 text is written in spans of
at most 256 bytes; cancellation remains checked per scalar, and a span that
cannot fit the row window or Work budget falls back to the scalar path. Work
continues to count encoded XML bytes. No public API, allocation strategy,
dependency, unsafe path, or common-crate ownership boundary changes.

The control build is retained at `91de70866df67bafd7d2f4c8931269ea540e1a1d`;
the candidate is `65dfb1b141190271763367defce4086ef146c056`.

The [result bundle](../results/change-0434/README.md) retains the frozen
protocol, source/build custody, reports, profiles, and derived summary. The
formal matrix contains 24 reports and 720 retained samples over 64, 8,192, and
32,768 rows, normal and allocator modes, two repeats, and four passing
whole-process profiles. The exact archive/XML/output/semantic/sink identities
are required before the descriptive comparison. The derived `claims[]` is
empty.

## Descriptive observations

Normal p50 deltas are after minus before; R1 and R2 are the two repeats:

| Shape | R1 | R2 |
| ---: | ---: | ---: |
| 64 rows | −12.499% | −16.211% |
| 8,192 rows | −13.731% | −13.121% |
| 32,768 rows | −13.562% | −13.492% |

No matched comparison exceeds the 5% regression trigger, including the
whole-process RSS observation. The only repeat flag is a +9.844% p99 drift in
the baseline role's tiny lane. Allocator calls, requested bytes, and regional
peak vectors are identical before and after for every shape and repeat; the
regional peak above entry is 419,347 bytes.

The four profiles are whole-process diagnostics, including setup and the
untimed oracle. Sampled `ExecutionContext::consume` self share changes from
25.01% before to 10.57% after. Ten before and eleven after addr2line warnings
remain retained. Zero L1 readings are not proof of zero misses, and LLC was
not captured. These observations do not authorize a release speedup, broad
10x claim, RSS or total-memory claim, physical-copy claim, or scaling claim.

## Validation and scope

The ODS release receipt reports 454 tests passed. Scoped Clippy, documentation,
format, crate-boundary, and candidate harness-build receipts pass; the
pre-existing common `ArchiveReaderKind` large-enum strict debt remains
explicit. Copied-bundle replay and six mutation probes pass before and after task cleanup.

This batch remains limited to fresh one-sheet scalar ODS creation. Logical
append, package-Part addition, arbitrary repackaging, native compatibility
breadth, cold/remote I/O, and broader CRUD remain open. The next implementation
slice is bounded ODT paragraph creation, followed by a separately scoped ODP
creation path. The overall non-iWork goal remains open.
