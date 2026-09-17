# Log sections for change 0673

Four paragraphs for the coordinator to merge into `HOTSPOTS.md`,
`GOAL_AUDIT.md`, `REPORT.md`, and `ADR_COMPLIANCE.md`.

## For `HOTSPOTS.md`

**0673 — the ZIP locator scratch is deferred off the common managed path.**
0651 row 18 left one bounded question after 0632: the managed index allocated a
64 KiB locator window before the `len - 22` fast path could decide. The path now
uses a 46-byte stack probe and allocates that same window only when the probe
misses. On the three retained OOXML probe fixtures this removes one allocation
and 65,537 allocated bytes per open; request counts and returned bytes do not
move. The 533-container open oracle is byte-identical, while the focused
comment fixture verifies the deferred search still requests the historical
64 KiB window.

## For `GOAL_AUDIT.md`

**0673 — the common-path resource is deferred without moving a refusal.** The
fixed EOCD probe is already mandatory for the managed index, so putting its
46-byte storage on the stack removes a heap reservation that was made before
the probe. A rejected probe falls into the same backwards search with the same
bounded 64 KiB scratch, and the central-directory prefill, ZIP64 handling,
short-read loops, limits, and typed scan refusals remain unchanged. The focused
tests and 533-container oracle cover the state that can be reached after the
probe; no performance claim is taken from the allocation count.

## For `REPORT.md`

**0673 — the managed ZIP index no longer allocates its locator window on the
exact terminal-EOCD path.** Retained in `soapberry-zip`, with no public API or
claim-registry change. The paired release probe records open allocated bytes
falling 2,256,370 → 2,190,833 on `xlsx-132`, 541,256 → 475,719 on
`pptx-shapes`, and 365,976 → 300,439 on `docx-comment`, one allocation less
each; every request count, byte count, and source-version observation is equal.
The 533-container reports are byte-identical. The fallback remains measured by
read-shape tests rather than a latency claim.

## For `ADR_COMPLIANCE.md`

**ADR 0005 and ADR 0006 — bounded scratch and fail-closed validation remain
intact.** The new stack storage is a fixed 46 bytes in the managed caller, and
the only deferred heap reservation is the existing `RECOMMENDED_BUFFER_SIZE`
window guarded by `try_reserve_exact`. A missed fixed probe still enters the
old bounded backwards search; EOCD, ZIP64, central-record, short-read, limit,
and typed malformed-directory checks are unchanged. The public locator's
default scratch contract remains unchanged, and the 533-container oracle plus
focused refusal tests show no acceptance or output movement.
