# Change 0477: explicit central-directory spool

The low-level ZIP writer can now serialize finalized central records into an
explicit caller-owned `Read + Write + Seek` store. Construction accepts a
serialized-byte quota and fixed replay window. Finalization replays only the
admitted appended extent, then emits the ordinary ZIP/ZIP64 tail. The default
in-memory writer remains available. The Office ZIP wrapper and OPC physical
writer expose the explicit storage capability within their existing owners.

Spool mode releases each member's central header and name after publication.
Its working memory includes the fixed replay buffer, one active name and
record bounded by ZIP field limits, and any active compressor. Provider-owned
storage remains separate. The ZIP Office normalized-name set and OPC full-name
and ancestor/descendant indexes remain in memory. This is not a bounded-total-
memory result for public PPTX streaming creation.

The source review and tests cover exact byte parity, Store/Deflate and sized/
streaming routes, attributes, timestamps and extra fields, ZIP64 counts and
offsets, quota boundaries, short and interrupted I/O, zero and overreported
transfers, truncation, seek/flush errors, provider error redaction, accepted
output progress and failed/dropped-entry poisoning. The provider must preserve
the admitted bytes and exclusive coherent access, including backing aliases.
No replay checksum is promised. Production never opens an ambient scratch
path; the caller supplies storage and its cleanup policy.

## Measurement scope

`protocol.json` freezes 48 isolated processes: two storage policies, three
member counts (8/256/8,192), Store/Deflate, separate normal/allocator binaries,
and two external repeats in reverse order. Each process runs three warmups and
thirty measured operations. The control preallocates its header vector; this
is a same-source policy comparison, not a historical/default-constructor
latency experiment. Previous 0475/0476 evidence selected growing metadata as
the next memory owner; this corpus counts ZIP members, not PPTX slides.

Every member has a deterministic 26-byte name and 256-byte payload. Every
process constructs both complete output oracles, compares them byte for byte,
and reopens/decompresses every member outside the operation interval. Timed
output goes to a fixed hashing sink. Timing includes sink construction,
explicit file creation/open, member output and hashing, central replay,
writer flush and file close. Digest finalization, endpoint metric snapshots,
oracle work, and file unlink are excluded. Oracle/source identities and every
sample's output hash/length are retained.

The configured replay window is 16 KiB and serialized scratch quota is 64 MiB.
The explicit scratch path is on `/tmp` **tmpfs** on this machine. Scratch still
consumes system memory proportional to serialized metadata, even when the
library's live allocation peak is flat. This is not physical-disk performance
or elimination of metadata storage. Each central record currently makes one
temporary allocation and one provider write; replay batches sink writes.

Normal timings are separate from instrumented allocator timings. Requested
bytes count cumulative allocation work; incremental peak subtracts each
region's entry live-byte baseline. Whole-process GNU time RSS includes corpus,
both materialized/reopened oracles, allocator history and teardown. Process I/O
deltas include the procfs observer and do not directly attribute physical I/O.
All costs and all five-percent review flags remain in the evidence; no
registered latency, default-route speedup or full-program claim is made.

## Results

All 48 formal captures pass, retaining 1,440 samples. Both pilots pass with
four operations each, excluded from the formal matrix. Every allocator sample
has zero failed allocations and zero live-byte exit change. Peak heap and
allocation totals are identical across both repeats at each policy/shape.

| Method | Members | Peak heap, control → spool (bytes) | Allocation calls, control → spool | Requested bytes, control → spool |
| --- | ---: | ---: | ---: | ---: |
| Store | 8 | 1,168 → 16,574 | 5 → 19 | 1,350 → 17,260 |
| Store | 256 | 37,376 → 16,574 | 10 → 515 | 44,006 → 41,564 |
| Store | 8,192 | 1,196,032 → 16,574 | 15 → 16,387 | 1,408,998 → 819,292 |
| Deflate | 8 | 414,128 → 429,534 | 7 → 21 | 414,310 → 430,220 |
| Deflate | 256 | 450,336 → 429,534 | 12 → 517 | 456,966 → 454,524 |
| Deflate | 8,192 | 1,608,992 → 429,534 | 17 → 16,389 | 1,821,958 → 1,232,252 |

Normal medians remain separate from allocator timing:

| Method | Members | R1 control / spool p50 (ms) | R2 control / spool p50 (ms) |
| --- | ---: | ---: | ---: |
| Store | 8 | 0.003270 / 0.007330 | 0.003180 / 0.007090 |
| Store | 256 | 0.095810 / 0.133810 | 0.095530 / 0.137300 |
| Store | 8,192 | 3.154354 / 4.356869 | 3.073433 / 4.290238 |
| Deflate | 8 | 0.081110 / 0.086370 | 0.081031 / 0.086771 |
| Deflate | 256 | 3.484366 / 3.551966 | 3.485035 / 3.555285 |
| Deflate | 8,192 | 109.175398 / 113.534436 | 110.987458 / 112.270294 |

Twenty policy pairs and six repeat comparisons have at least one five-percent
review flag. See [measurement review](measurement-review.md) for every flag and
its interpretation. The fixed spool window increases tiny-operation heap, and
per-record allocation/File writes increase call counts and Store latency. The
large scratch extent is 589,824 bytes on tmpfs. This is an opt-in memory/storage
tradeoff, not a global speedup or physical-disk claim.

## Final validation and cleanup

All 23 required gates pass. The ledger retains 40 validation attempts: 33
successful and seven unsuccessful development attempts. Required earlier gates
have identical production source; only the declared benchmark lint-fix file is
projected where applicable. Six Python evidence tests pass. Live verification
matches the final 7,040-source manifest and both copied executables before
cleanup.

`cleanup.json` records removal of the seven owned runtime files and their
root; all per-operation spool files were already unlinked. Shared Cargo caches
and both user-owned files remain intact. After executable removal, a sealed
fresh-copy replay passes and seven independently resealed corruptions are
rejected. `portable.json` records the exact validation seal and failures;
`portable.py` reproduces those probes. The final seal also covers that receipt
and this documentation; no capture was removed or replaced.

## Custody and reproduction

`gate.py` retains exact validation commands, environment, logs and full
content-addressed Rust/TOML/lock source manifests before and after every
attempt. Development failures are retained alongside corrected gates.
`build.py` creates normal and allocator executables, authenticates their copied
bytes, and binds each to its successful build receipt and source manifest.
`capture.py` authenticates those executables against `binaries.json` and binds
each process receipt to the frozen protocol. The allocator implementation and
its five tests were extracted unchanged from the existing diagnostic binary;
`allocator-extraction.json` records that identity.

Run `python3 -B verify.py` from this directory for portable verification after
the bundle is complete. It checks custody, cardinality, producer fields,
oracles, arithmetic, summary replay and exact sealed file inventory without
reading the old temporary executables or rebuilding Cargo targets. Evidence
unit tests and the fresh-copy/negative verification record exercise rejection
of resealed corruptions. Reproduction of new observations requires rebuilding
the same source/feature variants and selecting a new explicit scratch path;
the retained capture driver refuses to overwrite existing records.

The [ownership audit](ownership-audit.md),
[generated-name proposal](generated-name-design.md), and
[next work](next-work.md) distinguish the remaining semantic integration from
this substrate. Fresh creation, logical append, Part addition and arbitrary
editing/repackaging remain distinct. Native-producer breadth, source variants,
feature breadth, scaling and the full non-iWork goal remain open.
