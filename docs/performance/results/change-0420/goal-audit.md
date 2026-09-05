# Change 0420 goal audit: payload retention result and remaining evidence

`performance_claim: none`

`claim_authorized: false`

`disposition: OPEN`

This audit is anchored to the current [`docs/GOAL.md`](../../../GOAL.md), the
[0419 resource review](../change-0419/resource-review.md), the [0419 source
review](../change-0419/source-review.md), and the 0420
[source review](source-review.md) and [measurement protocol](measurement-protocol.json).
It covers non-iWork work only.

## What changed after 0419

0419 addressed the measured `BoundedVecWriter` request pattern in the private
PPTX archive buffer. Against control `7c7561160`, candidate `0320a6a88` uses
capped geometric growth followed by a fallible exact-size reservation and
copy. In the media-rich allocator lane, cumulative requested bytes fell from
36,811,874,712.7 to 369,979,465.4 (−98.995%), allocation calls fell 3.288%,
and reallocation calls fell 21.763%. The plain lane's requested bytes fell
36.217%, with calls down 0.696% and reallocations down 5.410%.

Those are allocator-request accounting totals. They are not physical-copy,
resident-memory, or RSS measurements. The retained live and high-water
snapshots were unchanged within the recorded pairs, and whole-process RSS was
slightly lower, but both are process-scoped observations. The normal lane has
100 samples and 10 warmups, so its media and plain timing movement is
diagnostic only; 0419 makes no latency claim. Heaptrack identifies relevant
writer stacks, but its retained filtered histogram is still whole-command
data. The final compaction buffer can coexist with the grown buffer, and the
archive output limit is not an aggregate process-memory budget.

The remaining cost is separate from writer growth. An eager OPC reopen retains
the new raw archive for exact-source preservation and can retain newly decoded
payloads while the donor/source and staged graphs remain alive. 0419 therefore
reduced request churn without proving that the raw archive or decoded payload
allocations are gone.

## 0420 status

The candidate `OpcPackage::from_vec_reusing_payloads` path is present in the
OPC owner and has focused correctness coverage. It still performs complete
physical-package validation and decompression. It reuses a donor payload only
after matching URI, content type, visible bytes, and a no-larger donor
capacity; the returned package retains its own archive metadata, preservation
provenance, and source authorization. Donor metadata, save options, and
authorization are not adopted as authority. Mismatched bytes, oversized donor
capacity, inconsistent custom donor storage, malformed input, and input-limit
failures keep the independent decoded payload or the typed refusal. The
focused payload-reuse suite reports 7 passing tests.

The completed 0420 capture has 16/16 runs passing: 800 normal samples and 240
allocator samples across the two selectors and four ABBA legs. Media
`live_bytes_after` falls 17.51845% in both pairs (287,885,879 to
237,452,746), process high-water falls 9.35929%, and whole-process RSS falls
about 9.15%--9.27%. Plain `live_bytes_after` falls 9.88924%; its high-water
change is only −0.17570%, and RSS is effectively flat (normal pairs are
within 0.02%; allocator pairs are −0.08% and +0.15%). Requested allocation
volume and calls are effectively unchanged, as expected because the new
archive is still fully decompressed.

The normal lane remains diagnostic: media p50 changes +0.656% and −1.554%,
and plain p50 changes +1.688% and +1.224%. Plain A2→B2 p99 is +5.21472% and
is flagged for review. All same-revision drifts are below 5%, but no latency
claim is authorized. The source/full-suite gate reports 1,286 tests passed and
3 ignored, with the focused ownership tests passing. The candidate is retained
for this scoped resource benefit following independent review; this is not a
general memory claim.

The source review distinguishes 33.554432 MB of media in one application
candidate from 16.777216 MB retained by the first plan's patch. Together these
account for 50.331648 MB of the observed lifecycle reduction; 101,485 bytes
remain unassigned without per-allocation attribution. Raw source/output ZIP
allocations remain required, and the candidate does not deduplicate equal
source and destination payloads across packages.

## Ordered remaining work

1. **Keep the retained memory decision scoped to its measured boundary.**
   The independent resource review accepts the lower live bytes and media RSS
   with the disclosed plain p99 cost. The ownership review reconciles the
   principal media contribution across both candidate builds. A larger sample
   or local attribution is needed before any latency or broader workload
   claim; the retained diagnostic supports neither. Carry the same source,
   corpus, output and refusal bindings into any follow-up.

2. **Measure the retained artifact directly after the diagnostic.** Add an
   operation-local retained/peak observation and near-output-limit and
   low-memory owned media cases. Separate raw archive bytes, decoded payload
   bytes, and graph metadata lifetime. Verify that donor reuse cannot transfer
   source authority or release bytes still required for exact preservation.
   A positive snapshot result on this one corpus does not establish a bounded
   memory guarantee; a regression must remain visible rather than being
   hidden by cumulative allocation totals. The current evidence still does
   not establish how the raw archive, decoded payloads, and graph metadata
   coexist at the operation peak.

3. **Resume the larger source-backed OPC/XLSX lifecycle baseline.** The
   current default remains 36 cases/198 rows and the representative index
   remains 15 categories/30 selectors; many representative rows still time
   an already-open query or commit and lack operation allocation attribution.
   Measure existing eager/source-backed selected reads and cell edits with one
   lifecycle definition, operation allocations, source-range requests,
   decompressed/recompressed/copied bytes, output readback, and fixed corpus
   identities. The 0420 PPTX payload result cannot generalize to
   `xlsx_source_*`, range-source, CFB, or other format paths.

4. **Then close environment and workload gaps.** Controlled block-backed cold
   and caller-supplied range-source runs, permitted native Office/LibreOffice
   fixtures, managed-cache contention, and explicit 1/2/4/8-worker scaling
   remain open. Existing simulated range and tmpfs observations, warm
   generated sources, and one-worker PMU data do not establish those
   requirements. ZIP64 semantic/failure breadth and the remaining CFB, ODF,
   RTF, and dependency-closure matrices remain correctness or narrow-phase
   evidence until they receive comparable end-to-end resource measurements.

Independent resource review recommends retaining 0420 for its scoped memory
benefit while accepting the disclosed plain timing cost. The non-iWork goal
remains open until the
definition-of-done requirements in `docs/GOAL.md` have reproducible evidence.
