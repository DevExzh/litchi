# 0421: correct allocator high-water accounting

The benchmark allocator passed the previous live-byte total to its peak
counter. `AtomicU64::fetch_update` returns the value before a successful update;
`checked_add` forwarded that value unchanged. A single 64-byte allocation
followed by deallocation therefore left a reported peak of zero. Six overlapping
17-byte allocations reported a peak of 85 instead of 102 bytes.

The helper now derives the post-update total from the successful atomic
operation's previous value and increment. It retains checked arithmetic and
uses that exact update rather than a later atomic load that could observe a
concurrent deallocation. No production document path, allocator ownership,
unsafe block, live-byte update, or cumulative call/byte counter changes.

Corrected allocator reports add
`tool.allocator_counter_revision = "post_update_peak_v2"`. Normal reports omit
this field. The existing comparator checks the complete tool identity against
policy and between reports, so a markerless historical report cannot be mixed
with a corrected report under one policy. Existing allocator policies remain
historical; opt into the new marker only after capturing both corrected sides.
Historical schema replay is not approval to reuse the defective peak fields.

Five regression tests failed on the old implementation, then passed after the
fix. They cover a transient allocation, growing/shrinking reallocations,
failed allocation, a sequence checked against an independent live-byte model,
and barrier-synchronized overlapping allocations. The concurrent test releases
and joins workers before assertions. Compatibility tests cover matching new
reports, old/new mixtures, and either generation with the wrong policy.

The [affected-evidence audit](../results/change-0421/affected-evidence.md)
traces the defect to `4c0224bc44fb9dae881aa1a6bad696f62239373d`. Historical
`peak_live_bytes_before/after` values and derived differences must not support
high-water claims without corrected captures. Raw reports remain unchanged;
correction notices annotate the change records and central reports. The
registered claims do not depend on this counter. Normal timing, cumulative
allocation requests, live-byte snapshots, independent RSS and Heaptrack
measurements are unaffected by this specific defect.

The [corrected current baseline](../results/change-0421/result-table.md) uses
two fresh allocator processes per existing media-rich/plain PPTX lifecycle,
30 samples and three warmups each. It is a current baseline, not a paired
optimization result: no allocator latency comparison, speedup or historical
peak reduction is claimed. The [bundle](../results/change-0421/README.md)
retains the frozen protocol, exact identities, vectors, source/output gates,
failed and passing tests, and replay tools.

Validation includes 271 passing Rust tests with one ignored, followed by final
counter checks, 92 comparator tests, warning-denied rustdoc, strict claim and
boundary checks, and portable replay. Clippy exited successfully with existing
warnings; it was not a warning-free gate. See the detailed
[validation receipt](../results/change-0421/validation.md) for revision boundaries.

Even corrected allocator high-water values are process-lifetime observations
of successful allocation requests. They exclude allocator-internal realloc
copy overlap and do not measure RSS or an operation-local peak. A separate
region-local tracker and lifecycle retention/drop boundaries remain needed,
followed by near-limit memory and matched source-backed workflows. The full
non-iWork performance goal remains open.
