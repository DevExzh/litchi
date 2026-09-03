# Change 0392: PPTX source-backed picture-query baseline

Date: 2026-09-03

Status: implemented as evidence-only standalone selectors; release measurement pending

`performance_claim: none`

`claim_authorized: false`

The standalone `litchi-perf-baseline` harness now exposes three opt-in,
source-backed PPTX selectors:

* `pptx_source_backed_images_query`
* `pptx_source_backed_image_query`
* `pptx_source_backed_read_image_query`

They reuse the existing deterministic `pptx_cross_copy_media_rich` source
archive and select its picture-heavy third slide, which is required to contain
eight direct pictures in exact scene order. Package/open and slide preparation
are outside the timed interval. Metadata queries verify the complete ordered
descriptor projection and selected identity without media payload reads;
`read_image(3)` verifies the selected relationship, selected compressed source
range locality, and payload SHA-256. Source reads and source-cache counters come
from an independent replay for each retained sample, and the operation metrics
envelope publishes allocator vectors only in the allocator-enabled target.

This change does not alter production code, CRUD coverage, iWork behavior, or
the default 36-case / 198-record matrix. It raises the opt-in selectable
registry from 408 to 411 cases. No release performance measurement is recorded
yet because the current host workload is noisy.
