# Demand-cache complex payload layout plan

This is a source-only plan for the small accounting delta reported by the
common09 comparison. The exact before executable is
`/home/zhuhe/code/litchi-array-target/retained/value-complex-09`, SHA-256
`a331a2ec03c2363face9887014ee0ca7ee0a989cc8f7d44b1588af06e5fa96d5`. The
common09 row `matrix-lazy-aggregate-4096` changed requested and released bytes
from 3,146,824 to 3,146,840 with unchanged allocation calls, work, and peak
live values. The existing report attributes this to the widened private
`DemandCacheValue`/`DemandCacheEntry`; this plan treats that attribution as a
hypothesis until the candidate layout and a replay confirm it.

## Layout hypothesis

`DemandCacheValue` currently embeds `Complex`, whose two `f64` components and
`char` suffix occupy 24 bytes. The proposed private representation uses
suffix-specific variants, `ComplexI([f64; 2])` and `ComplexJ([f64; 2])`, so the
suffix remains in the enum tag and the payload is only the two components.
This should keep the value inline and `Copy`, with no new allocation or
lifetime change. Reconstruction must use the existing checked `Complex::new`
path (or an equivalent checked private helper), preserve lowercase `i`/`j`,
negative zero, and both finite components, and turn an impossible invalid
cached payload into an evaluation error rather than panic. A private layout
test should record `size_of::<DemandCacheEntry>()` and enforce the expected
bounded footprint (at most 32 bytes on the pinned target). A single tagged
`[f64; 2]` plus a suffix byte is a control design: alignment likely leaves it
at 24 bytes and therefore may not shrink the entry as much.

`DemandCacheValue` is used by the sorted `demand_cache` in
`evaluation/value.rs`; it is distinct from `MatrixState`'s direct branch
cache and from `ConditionCacheValue`, which does not store Complex. The
candidate must leave all non-Complex variants, binary-search Work charges,
cache ordering, and reservation/error paths unchanged.

## Workloads and checks

Build the same focused harness against the before ELF's source snapshot and
the candidate source, retaining both binaries. Use two direct projected
Complex cases from the accepted vector, with the same shape and limits:

* `=IMSUM(IF({TRUE();TRUE()};IMSUM([.A1]);0))` with `A1="3+4j"`, expecting
  scalar/array Complex `6+8j`;
* the same case with an `i` suffix and an independent expected component
  oracle.

These cases force a projected sequence result through `DemandCacheValue` and
check that the first cell's resolver read is reused while every output keeps
its suffix and components. Extend this with K distinct cacheable sequence
nodes (K = 1, 2, 4, 8, 16, 32), using parallel/nested projected IF branches
and alternating `i`/`j` values. Record the actual demand-cache-entry count in
the focused diagnostic; if a shape planner routes a branch through
`MatrixState` instead, classify that lane separately rather than calling it a
demand-cache measurement. Include a boolean AND/OR K-entry control to
separate enum-capacity effects from Complex reconstruction cost.

Replay the unchanged common09 corpus, including
`matrix-lazy-aggregate-4096`, to catch unrelated evaluator changes. For every
lane compare checksum, suffix/components, resolver reads, Work, success/error
kind, allocation calls, requested/released bytes, peak-live accounting, and
budget-retained bytes. Capture p50/p95/p99 and process RSS with fresh serial
children pinned to CPU 6. Use three warmups and 31 measured iterations, in
AB/BA/AB order, with the same locked toolchain, limits, fixtures, and harness
bytes. Hash both ELFs and all source/harness inputs before and after each
capture; reject any change during a run.

The expected one-entry accounting reduction is 16 bytes because
`ensure_capacity` grows an empty cache to two entries; larger K values should
show reductions at vector-capacity boundaries rather than a per-evaluation
allocation-count change. Treat this as diagnostic evidence until repeated
rounds agree. Report instrumented requested/released bytes separately from
allocator peak and process RSS, and make no broad performance claim from the
single prior common09 comparison.

## Measured decision

The hypothesis above was disproved by actual private-type probes on the pinned
Rust compiler: candidate09 and the suffix-specific cache10 prototype both have
Complex=24, DemandCacheValue=24, and DemandCacheEntry=32 bytes. The earlier
40-byte entry estimate overlooked Rust's enum niche layout. The common corpus
also reports identical requested/released bytes before and after. Root removed
the suffix-specific production prototype; it adds reconstruction work without
reducing storage. The exact prototype source, gates, probes, and comparison
remain diagnostic evidence. The public cache suffix/read-reuse regression and
32-byte footprint bound are retained independently of the discarded design.
