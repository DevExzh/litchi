# Change 0477: explicit ZIP central-directory scratch storage

`performance_claim: none; descriptive same-source storage-policy comparison`

`claim_authorized: false`

The previous PPTX allocation evidence identified retained ZIP directory and
ZIP/OPC name metadata after compressor reuse removed repeated allocation work.
Commit `90f04691c` adds an explicit central-directory spool to the low-level ZIP
writer and adapters in the ZIP Office wrapper and OPC physical writer. Each
finalized central record is appended to caller-owned replayable scratch;
finalization replays the admitted extent before the ZIP/ZIP64 tail. Spool mode
releases central headers and their retained names instead of accumulating them.

Callers select both the serialized-byte quota and fixed replay window. No
ambient scratch file, dependency or production unsafe code is added. One
active record/name and optional compressor remain bounded working state. The
provider's own storage is separate, and it must preserve exclusive coherent
access to its admitted range, including aliases. Error displays redact provider
text while retaining the source chain. ZIP64 offsets and central extras,
short/interrupted/zero/overreported transfers, prefix preservation, typed
limits, accepted-output counts and failed/dropped-entry poisoning are tested.

The [evidence bundle](../results/change-0477/README.md) contains 48 successful
formal processes and 1,440 samples: 8/256/8,192 ZIP members, Store/Deflate,
normal/allocator builds, two storage policies and two reversed process repeats.
Two pilots contribute eight excluded operations. Every process byte-compares
both output oracles and reopens every generated member; all measured hashes and
lengths match. The control preallocates its header vector. This compares storage
policies in one implementation, not historical/default-constructor latency.

Every allocator sample has zero failed allocations and zero live-byte exit
delta. The following allocation values are identical across all sixty samples
per policy/method/size. Peak means region high-water minus entry live bytes;
requested bytes are cumulative allocation work.

| Method | Members | Peak heap, control → spool (bytes) | Allocation calls, control → spool | Requested bytes, control → spool |
| --- | ---: | ---: | ---: | ---: |
| Store | 8 | 1,168 → 16,574 | 5 → 19 | 1,350 → 17,260 |
| Store | 256 | 37,376 → 16,574 | 10 → 515 | 44,006 → 41,564 |
| Store | 8,192 | 1,196,032 → 16,574 | 15 → 16,387 | 1,408,998 → 819,292 |
| Deflate | 8 | 414,128 → 429,534 | 7 → 21 | 414,310 → 430,220 |
| Deflate | 256 | 450,336 → 429,534 | 12 → 517 | 456,966 → 454,524 |
| Deflate | 8,192 | 1,608,992 → 429,534 | 17 → 16,389 | 1,821,958 → 1,232,252 |

The spool's measured library heap peak is flat across these counts. Serialized
scratch grows to 576 / 18,432 / 589,824 bytes, separately from the 16 KiB replay
window and 64 MiB quota. Scratch is on **tmpfs**, so this moves storage into
caller-selected system memory; it does not eliminate metadata storage or prove
physical-disk performance. The 8-member case pays more heap for the fixed
window. Per-record name/serialization allocations also increase allocation
calls substantially. All unfavorable observations remain retained.

Normal timing stays separate from allocator timing:

| Method | Members | R1 control / spool p50 (ms) | R2 control / spool p50 (ms) |
| --- | ---: | ---: | ---: |
| Store | 8 | 0.003270 / 0.007330 | 0.003180 / 0.007090 |
| Store | 256 | 0.095810 / 0.133810 | 0.095530 / 0.137300 |
| Store | 8,192 | 3.154354 / 4.356869 | 3.073433 / 4.290238 |
| Deflate | 8 | 0.081110 / 0.086370 | 0.081031 / 0.086771 |
| Deflate | 256 | 3.484366 / 3.551966 | 3.485035 / 3.555285 |
| Deflate | 8,192 | 109.175398 / 113.534436 | 110.987458 / 112.270294 |

Spooling adds exactly 8 / 256 / 8,192 observed write syscalls per operation in
this File provider. It reduces output-sink calls by replaying central bytes in
chunks, but the large Store medians rise from approximately 3.1 to 4.3 ms.
The [measurement review](../results/change-0477/measurement-review.md) examines
all 20 flagged policy pairs and six flagged repeat comparisons. The full
mean/p50/p95/p99, process-RSS and allocator fields remain in `summary.json`.
Whole-process RSS includes matched materialized/reopened oracles and setup;
procfs I/O deltas include observer activity. No general latency, physical-I/O,
RSS or default-path improvement is claimed.

Validation includes 524 ZIP tests (33 spool integration tests), 482 OPC tests,
180 selected streaming unit tests, the named DOCX/PPTX/ODF streaming
integrations, 367 harness tests and the five allocator tests in each enabled
binary. Strict ZIP/OPC and harness Clippy, warning-denied documentation, scoped
formatting and crate-boundary checks pass. A small harness lint rewrite is
retained as `harness-clippy-fix.patch`; the four affected diagnostic tests and
corrected lint gate pass against the measured source. Earlier unaffected gates
bind identical production source with only that declared harness-file
projection. Development failures remain in the validation ledger.

The private writer layout diagnostic measures 160 bytes in the inline-spool
draft and 80 bytes after boxing the optional spool on this build. This is a
within-batch layout observation, not a claim that the pre-change ABI or default
layout was identical. Normal and allocator builds bind the same complete
source manifest and copied executable hashes.

All 23 required gates and six Python evidence tests pass. The ledger retains
40 attempts, including seven development failures. Live source/binary checks
pass before cleanup; a sealed fresh-copy replay and seven resealed negative
probes pass after the runtime binaries are removed. Shared Cargo caches and
user-owned files are preserved. The final bundle records exact file inventory
and digests for portable replay.

ZIP Office and OPC name indexes still grow. The public PPTX streaming route
therefore has no constant-total-memory claim from this low-level experiment.
The [next work](../results/change-0477/next-work.md) calls for a checked finite
generated-name capability, semantic scratch integration and repeated public
PPTX operation measurements. Fresh creation remains distinct from logical
append, Part addition and arbitrary editing/repackaging. Native producers,
source variants, feature breadth, bounded-worker scaling and the full
non-iWork goal remain open.
