# Scoped review

The production XLSX chunk changes retain length/divisibility checks before
fixed-array iteration. `as_chunks` is available before the workspace's Rust
1.89 minimum. A fixed array is copied directly into `u64::from_le_bytes`;
the former exact-slice conversion is no longer needed. Namespace resolution
returns the same borrowed bytes. Boxed errors and borrowed UTF-8 conversion
remain test helpers; all typed variant/message assertions remain present.

The first gate found the array-reference conversion and two unnecessary test
qualifications. The second found an additional existing integration-test
temporary UTF-8 allocation. Both logs are retained, followed by the passing
all-target Clippy and 1,238-test run.

Harness review confirmed identical timer/allocator operation boundaries:
source construction and archive clone before `begin`; open/query inside;
allocation `finish` before validation and owner drop. The new test's explicit
unavailable-allocator assertion shares a test-only lock with tests that toggle
the global enable flag. The nine XLS/allocator tests passed together with four
threads. Release behavior is unaffected by that test-only lock.

The source observer linearly folds each classified range vector on every
read. The fixed opaque payload streams alone contribute 32,768 sector ranges.
`new_xls` also enables range-union tracking, while `SourceSummary::record_xls`
does not publish repeated-range overlap. The CPU capture locates most observed
cost in this diagnostic callback. Removing unused observer work or introducing
a matched plain-source observation is a proposed next experiment, not a change
or performance claim in this batch.

The initial replay of the corruption probes against compressed published files
found a probe-environment defect: zstd refused symlink inputs, causing three
unrelated missing-input failures. The original uncompressed corruption probes
had passed. The replay now copies compressed fixtures into private regular
files and requires each mutation's expected diagnostic, so an unrelated failure
cannot count as rejection. The initial and corrected published results are
retained; measurement files were not changed.

Final evidence review matched the published percentile ranges, allocation
counts/bytes, source/version counters, RSS, PMU and CPU attribution against the
retained JSON. Wording distinguishes the tabled invariant allocation values
from changing live-byte snapshots, and attributes the profile to the complete
ReadAt observer rather than claiming a measured split of its internal costs.
