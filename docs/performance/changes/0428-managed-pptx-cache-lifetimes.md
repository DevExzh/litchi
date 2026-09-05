# 0428: managed PPTX cache and budget lifetimes

`SourceBackedPresentation` and `SourceBackedPresentationEditor` now expose
`try_cache_diagnostics()`, forwarding OPC's typed poisoned-state and counter
overflow errors. Observation does not load a part or retain another package.
The existing infallible methods remain compatible.

The normal performance binary's `cache-retention` command records explicit
ownership phases for matched plain/media-rich cross-copy workflows and
selected-image resource boundaries. Source and destination use separate caller
budgets. Publication consumes the destination editor; its cache is thereafter
unavailable. After the final source owner drops, caller budget gauges remain
observable without fabricating a cache snapshot.

Sixteen fresh release processes retain 480 samples and 5,460 phase points.
Every final caller budget has zero Memory, Objects, and Depth; non-RSS numeric
phase observations match within and across repeats. All 27 repeat flags are
process RSS points, already different at entry for the image lanes. Comparative
RSS claims are withheld; every raw flag remains in the bundle.

This is a measurement enabler, not an optimization. It removes no production
work and makes no latency, throughput, allocator, physical-copy, general leak,
or causal RSS claim. The allocator evidence from 0427 remains separate.

The [bundle](../results/change-0428/README.md) retains the frozen protocol,
machine/build identities, all development failures, raw samples, independent
report validation, mutation probes, lossless logs, and portable replay.
[Resource review](../results/change-0428/resource-review.md) distinguishes
cache retention, managed reservation gauges, cumulative charges, source-read
windows, and process RSS. [Source review](../results/change-0428/checks/cache-review.md)
and the [ADR matrix](../results/change-0428/checks/adr-review.md) record the
ownership and admission constraints.

The full PPTX suite passes 846 tests with two ignored. Both new forwarding
tests assert exact diagnostics and no source reads; seven OPC fail-closed
cases cover the underlying diagnostic error paths. Five new harness unit
tests, the existing plain/media lifecycle regression control, all eight real
CLI scenarios, and 136 final control mutations pass. PPTX strict Clippy,
minimal-feature checking, warning-denied documentation, formatting, crate
boundaries, CRUD coverage, and registered performance claims pass. Harness
strict Clippy retains the same 29 findings in 17 message/source-file groups
as 0427, with none in the new module; that command is not labelled passing.

Portable replay before and after cleanup passes all 16 formal reports and their 272 mutation probes,
and rejects a changed repeat output even when its receipt digest is updated.
The hash-bound temporary capture executable has been removed; the original
build executable and both existing target directories are preserved.

The [development corrections](../results/change-0428/checks/development-corrections.md)
retain the initial compile error, new lint findings and fixes, validator
mismatch, corrected admission floor, and observer/read-counter corrections.
No failed preflight report is included in formal release counts.

The broader non-iWork goal remains open. The next evidence gap is matched
native-producer and cold/range-source lifecycle coverage, followed by measured
remaining cost attribution. These synthetic memory and output boundaries do
not establish bounded total process memory, complete CRUD coverage, or scaling.
