# 0422: measure allocator high water within each operation

The corrected process high-water counter from 0421 can be dominated by setup.
Repeated lifecycle reports therefore cannot use its before/after snapshots to
identify the largest live allocation total reached inside each operation.

The allocator observer now maintains a separate `region_peak_live_bytes`,
initialized from entry live bytes. A mutex serializes whole counter callbacks
and region begin/finish/drop, giving them one observation order. The global
counters and lifetime maximum never reset. Normal binaries do not install this
allocator wrapper. No document implementation, dependency or unsafe code is
changed.

The peak counts absolute process live request sizes, including bytes already
live at entry and other threads' callbacks during the region. All operation
workers must finish before the caller ends the region to include their entire
work. Callbacks occur after System allocator calls; the metric excludes hidden
realloc copy overlap and does not measure physical RSS or operation-owned
retention. The observer mutex can alter scheduling, so allocator elapsed time
and scaling remain non-claimable.

The report identity is `serialized_region_peak_v3`. The new field is required
for measured V3 reports, must cover both live endpoints, and cannot exceed
lifetime high water at exit. It is aligned with elapsed sample order and omitted
as a numeric value for unavailable/overflow observations. Historical V2 replay
remains supported under its own tool identity; generations cannot be mixed
under one policy. Dedicated historical experiment policies remain historical
until both sides are recaptured.

The [bundle](../results/change-0422/README.md) retains the frozen protocol,
validation, exact source/binary/corpus/output identities and a current baseline
for the existing media-rich/plain PPTX lifecycles. This is a measurement
enabler, not a speedup or memory-reduction claim. Portable replay includes
pinned validators and needs no original build artifacts.

Both 30-sample repeats report mean region peaks of 272,736,303 bytes for
media-rich and 1,360,003 bytes for plain. Individual samples vary slightly;
the retained vectors and resource review include their ranges. The same reports retain process
lifetime peaks of 812,687,524 and 3,559,492 bytes, respectively. This separates
operation peaks from the setup maximum within one observer generation; it is
not a before/after reduction. Validation includes 50 focused Rust tests, 98
comparator tests, warning-denied rustdoc, repository gates and portable replay.
Strict Clippy remains blocked by documented existing harness warnings.

Explicit retention/drop boundaries, near-limit memory tests and matched
source-backed media lifecycles remain next work. The broader non-iWork goal
remains incomplete.
