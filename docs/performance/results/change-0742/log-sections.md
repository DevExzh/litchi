# Log sections for change 0742

## For `HOTSPOTS.md`

## 0742 — owned cross-copy stops re-deflating copied images

[0742](0742-pptx-owned-cross-copy-media-transfer.md) removes the generated-entry
Deflate that 0740 located under owned cross-copy planning. Eligible copied
image members (relationship-free, non-XML `image/*`, with a provable Store or
Deflate layout and a bounded compressed size) are framed from the source
member's verified compressed bytes. The decision depends only on the bytes the
source publishes.

On the media-rich pair, with both legs built by the same command:

- the lifecycle median p50 moves 410.1 → 183.0 ms (paired ratio 0.444);
- planning moves 292.6 → 65.7 ms.

What remains is bound by SHA-256. Commit (about 89 ms) is 80% SHA-256, in five
passes of about 16% each:

- the live semantic re-fingerprints;
- the live physical re-fingerprints;
- the retained-archive digest;
- two semantic captures of the same reopened candidate.

Reusing proven digests there is the next opportunity; it is not implemented.

## For `REPORT.md`

## 0742 — verified compressed transfer for owned cross-copy media

[0742](0742-pptx-owned-cross-copy-media-transfer.md) adds:

- `OpcPackage::compressed_transfer_size`, a header-only eligibility check that
  includes a size guard;
- `OpcPackage::authorize_compressed_transfer`, which returns `Ok(None)` when
  the member's own bytes disprove the capture;
- a transferred payload that the targeted writer frames with fresh sized
  headers;
- the copied-media encoding, recorded in the plan and in the durable patch
  (`LPCP0004`; `LPCP0002`/`LPCP0003` are refused by name).

A package that is not an unmodified owned source is serialized and reopened,
so the decision depends only on its bytes. Redo after undo and byte-identical
destinations therefore publish the first copy's bytes. A copy never transfers
at the price of a refusal or of a caller's part: captures that alone would
cross `max_patch_bytes`, a re-read its package's own read limits refuse, and a
destination holding a caller-defined `Part` each make planning record the
recompressing route.

Results, by median process p50:

- media-rich lifecycle: 410.1 → 183.0 ms;
- media-rich plan+commit+publication: 386.7 → 159.0 ms;
- plain cases: within 1%, byte-identical to the base;
- source-backed control: unchanged.

Allocated bytes fall 1.8%; peak live bytes are unchanged. The spread of both
arms follows first-touch page faults of 33.6 MB buffers, reported per sample.
`performance_claim: none`.

## For `GOAL_AUDIT.md`

## 0742 — owned cross-copy media transfer under the alpha trade-offs

[0742](0742-pptx-owned-cross-copy-media-transfer.md) applies 0652's standing
trade-offs:

- **Format.** The durable cross-copy format bumps to `LPCP0004`, and genuine
  legacy patches are refused by name.
- **Destinations.** A transferring copy into a destination that is not an
  unmodified owned source publishes the reopened candidate, with save
  preferences carried. A destination holding a caller-defined part is planned
  with the recompressing route, which keeps the part; a transferring plan or
  patch that meets one is refused with `SlideCopyRefusal::CallerDefinedPart`.
  Keeping the captures instead would retain source bytes in the caller's
  package, against ADR 0005's retained-state rule.
- **What stays recompressed.** Members whose own bytes disprove the capture or
  fail the size guard are recompressed deterministically, as are
  relationship-bearing and XML parts and unprovable layouts. Resource failures
  stay typed errors.
- **Stored images.** A Stored compressible image stays Stored, larger than
  recompression would make it. The record states that cost.

The non-iWork goal stays open.
