# Change 0481: borrowed DOCX scanner names

`performance_claim: none; scoped control/candidate scanner comparison`

`claim_authorized: false`

The source-backed plain-paragraph scanner consumes element names only during
their parser events. The candidate retains each `LocalName` wrapper and borrows
its bytes, removing the separate fallible byte clone for start, empty and end
events. Namespace checks, attributes, scope transitions, event/depth/paragraph
limits and range construction run in the same order. No scanner pass is
removed. `source-review.md` covers the lifetime and semantic boundary.

## Reproduction and scope

Fresh normal and allocator control/candidate binaries use Rust 1.98.1, release
debug information, frame pointers, unwind tables, four build jobs and disabled
incremental compilation. Measurement commands run serially under the batch
lock. Each source manifest covers 7,048 files; the sole source difference is
the retained seven-line change in `paragraph_copy.rs`. Eight embedded DOCX
templates are copied and authenticated at each build.

The frozen A1/B1/B2/A2 order consists of control forward, candidate forward,
candidate reversed and control reversed. Each group crosses normal/allocator,
64/8,192/131,072 paragraphs and total/six-phase modes. There are 48 reports of
30 samples following three warmups, pinned to CPU 2. Eight one-sample pilots
cover all three sizes, totaling 24 excluded operations. Separate normal
large-case PMU processes provide wider context and remain excluded from the
formal measurements.

The byte-exact historical corpus manifest and report schema come from change
0480, whose corpus was established in 0479. Current arm-specific pilots bind
the same corpus to each fresh executable. Original baseline timestamps and
legacy pilot references remain provenance. The verifier requires the historic
freeze before formal runs and each arm's pilots before that arm's first run.

The lifecycle includes source adapter/package construction, snapshot, staging,
commit, sequential publication, digest finalization and owner destruction.
Corpus and independent oracles preexist the measured region. GNU `time` RSS
and PMU counters cover the whole process, including setup, warmups, oracles,
report serialization and teardown. Normal and allocator timing are separate.
Allocator requested bytes include realloc `new_size`; they are not physical
copy traffic. Phase peaks are never summed into an operation peak.

The operation remains one copy at the tail of a materialized plain document.
The complete XML and paragraph index remain live; this optimization does not
establish explicit-window existing-document append, repeated append scaling,
section-property support or the broader non-iWork goal. `window-contract.md`
records the separate capability design.

The final bundle is independently replayable with `python3 -B verify.py` after
copying this directory. Recorded original checkout and executable paths are
provenance, not runtime dependencies. `setup-attempts.json` retains one root
setup failure before Cargo or its gate started: the missing validation
directory was created before retrying, without changing source or scripts.

## Results and validation

Normal total means are 7.84–12.00% lower across the three sizes and both
repeats. Allocation callback removal is exactly `24*N+28`, and requested-byte
removal is exactly `24*N+108`, in every measured allocator sample. At 131,072
paragraphs, calls fall from 3,653,984 to 508,228. Incremental peak heap remains
35,371,102 bytes. Reallocation counts, output, source reads and sink identity
are unchanged. The complete result tables and the wider PMU scope are in the
per-change record; generic cache misses rise 13.52% in the separate PMU pair.

`measurement-review.md` independently agrees with every final summary row,
pair and repeat flag, including all 261 paired and 89 repeat review flags.
There is no positive total-lifecycle latency or GNU process-RSS pair change
above 5%. The review recommends keeping the seven-line candidate, committed
as `90f47ea41`; no general DOCX speedup or bounded-window claim is made.

All 22 required final gates pass. The ledger retains 24 successful receipts,
including both optional profiles; the earlier root setup failure is separately
retained. Fourteen evidence tests and data-only verification pass. Final source,
templates and all four binaries were authenticated before runtime cleanup.
Shared Cargo caches and both user-owned untracked documents remain intact.

Full sealed verification and a fresh copied baseline pass after removal of
the original runtime binaries. Eight independently resealed corruptions are
rejected: source-read accounting, output digest, allocator exit retention,
instrumentation, summary arithmetic, PMU arithmetic, omitted required gate
and source exclusions. `portable.json` retains the results and
`validation-seal.txt` retains the input seal. Replay copies are removed after
verification.
