# Log sections for change 0742

## For `HOTSPOTS.md`

## 0742 — owned cross-copy stops re-deflating copied images

[0742](0742-pptx-owned-cross-copy-media-transfer.md) removes the generated-entry
Deflate 0740 located under owned cross-copy planning: eligible copied image
members (untouched, relationship-free, non-XML `image/*`, provable Store or
Deflate layout) are framed from the source member's verified compressed bytes.
On the media-rich pair the lifecycle median p50 moves 405.4 → 185.7 ms
(paired ratio 0.461, both legs built by the same command) and planning
288.6 → 68.7 ms. What remains is SHA-256-bound: commit (about 89 ms) is 80%
SHA-256 in five passes of about 16% each — live semantic and physical
re-fingerprints, the retained-archive digest, and two semantic captures of
the same reopened candidate. Reusing proven digests there is the next
opportunity; it is not implemented.

## For `REPORT.md`

## 0742 — verified compressed transfer for owned cross-copy media

[0742](0742-pptx-owned-cross-copy-media-transfer.md) adds
`OpcPackage::compressed_transfer_eligible`/`authorize_compressed_transfer` and
a transferred payload the targeted writer frames with fresh sized headers, and
records the copied-media encoding in the plan and in the durable patch
(`LPCP0004`; `LPCP0002`/`LPCP0003` refused by name). Media-rich lifecycle
405.4 → 185.7 ms and media-rich plan+commit+publication 384.4 → 165.3 ms by
median process p50; plain lifecycle +1.0% (interval includes 1.0), plain
+2.15% with no stable direction across three matrices, both byte-identical to
the base; the source-backed control unchanged. Allocated bytes −1.8%, peak
live bytes unchanged. The spread of both arms follows first-touch page faults
of 33.6 MB buffers, reported per sample. `performance_claim: none`.

## For `GOAL_AUDIT.md`

## 0742 — owned cross-copy media transfer under the alpha trade-offs

[0742](0742-pptx-owned-cross-copy-media-transfer.md) applies 0652's standing
trade-offs: the durable cross-copy format bumps to `LPCP0004` with genuine
legacy patches refused by name, a transferring copy refuses a destination
modified since planning (clone-and-apply could not reproduce it without
retaining source captures, contrary to ADR 0005's retained-state rule), and
every capture failure after eligibility is typed. Ineligible members —
replaced payloads, relationship-bearing or XML parts, unprovable layouts —
keep the re-deflating route deterministically. The non-iWork goal stays open.
