# 0443: Compact source-fragment scanner frames

The sealed 0442 candidate profile attributes 7.908% of weighted sampled periods
to ContentSource under transaction; the grouped staging scan is 14.665% and
all transaction chains 23.576%. These include warmups and incomplete symbols,
not isolated phase timings. One-shot get_attr chains account for only 1.317%,
so compact fragment scanning is the stronger current ownership hypothesis.

The scanner copies every bound event namespace, including unused end-event
results, and copies local names for start/empty events. Open frames need only
a fixed element classification plus their start offset. Classify using the
current resolver, save that value across namespace pops, and preserve all byte
spans, root declarations, styles, pages, opaque extras, BOM, error priority and
limits. Keep the original scanner as a test-only independent reference.

Freeze A1/B1/B2/A2 before builds/edits: 24 reports, 720 samples, CPU 2, one
worker, 64/4096/8192 slides, normal and allocator modes, 30 samples/3 warmups.
Require 5% medium/large benefit in both repeats for normal p50 or calls,
requested bytes or operation peak. Review all matched and repeat >5% flags.
Four whole-process profiles follow. Retain all attempts; serialize CPU jobs
and source switches. Revert production if the practical gate fails.

The accepted ADR tree c950b6c8be822561b498d7bbe87c460873dcbf49 is unchanged
from the prior complete read. Ownership and grammar remain in ODP; exact
source authority, detached editing, limits and publication readback remain.
No public API, unsafe code, dependency, executor or ambient I/O is added.
The full non-iWork goal remains active; no CRUD coverage promotion is implied.
