# 0424 Heaptrack profile notes

The control profile is bound to source revision
`6ca9962c7818e173538a40ed45f3a3c32cc1aa6a`; the candidate profile is bound to
`d18bf7db4e01f35b8a936892abcb5751e7763d16`, with the candidate protocol
retaining the control revision as its ancestry anchor. Both are one-sample,
zero-warmup CPU2 Heaptrack whole-command captures. They are diagnostic stack
evidence and carry no cross-role latency or memory-saving claim.

For the media-rich command, the control trace requested 4,480,659,995 bytes
and deallocated 4,480,659,451 over 596,345 allocation events; the candidate
trace requested 4,413,557,279 bytes and deallocated 4,413,556,735 over
596,383 events. These totals include corpus
construction, setup, lifecycle work, verification, and teardown; they are not
operation-local totals.

The control trace contains two concrete `clone_bytes_checked` allocation
ancestries in the measured lifecycle:

- trace `147579`, under `plan_cross_slide_copy`: 8 allocation events,
  16,777,216 requested and deallocated bytes;
- trace `149788`, under `publish_cross_slide_copy_to_stream`: 8 allocation
  events, 16,777,216 requested and deallocated bytes.

The candidate trace contains one large lifecycle clone:

- trace `147588`, under `plan_cross_slide_copy` through
  `reuse_or_clone_payload`: 8 allocation events, 16,777,216 requested and
  deallocated bytes.

The matched control clone family totals 184,610,503 requested bytes across
99 events, with 33,565,546 lifecycle bytes; the candidate totals 117,501,639
across 67 events, with 16,788,330 lifecycle bytes. Each matched clone family
has zero net live delta. That is allocation/deallocation accounting for each
Heaptrack run; it does not establish retention after a source or snapshot is
dropped.

The first analysis pass tried a demangled `source_cross_copy::prepare` name,
while the retained v0 trace exposes the mangled `source_cross_copy7prepare`
symbol. It also relied on the capped `heaptrack_print` report, which omitted
the low-volume `clone_bytes_checked` symbol. The corrected analyzer scans the
retained interpreted trace `s` symbol records, then derives the stage receipt
from the full ancestry. The earlier pass remains as
`runs/{plain,media_rich}/analysis-initial.json`; the earlier failed replay is
retained in `runs/media_rich/analysis-failure.json`.

The retained control and candidate analysis manifests have passed their
trace-aggregate replay checks. The final portable bundle replay also checks
the matched summary, eight R1 report guards, both profile trees, and pinned
tool custody.
