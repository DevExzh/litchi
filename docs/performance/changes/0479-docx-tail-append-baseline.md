# Change 0479: existing DOCX plain-paragraph tail append baseline

`performance_claim: none; descriptive baseline and attribution`

`claim_authorized: false`

This diagnostic measures the existing public source-backed DOCX transaction:
copy source paragraph zero before the final insertion slot, commit, publish to
a sequential sink, and drop all operation owners. Source documents have 64,
8,192, or 131,072 distinct plain paragraphs and an opaque 32 KiB package member.
The public transaction permits one copy operation and refuses section properties.
This covers a single logical append to an existing structure. Bulk repeated
appends, documents containing section properties, streaming creation, adding a
Part, and arbitrary editing/repackaging remain separate workloads.

Production code is unchanged. The new standalone harness has two observation
modes and separate normal/system-counting-allocator builds. Total mode measures
one complete open/snapshot/stage/commit/publish/drop region. Phase mode measures
these six regions in separate executions, permitting retained-owner attribution.
Digest finalization and operation-owner destruction are included. Corpus
construction, full semantic oracles and retained source/candidate archive storage
are outside operation regions. Process observations include their own broader
boundary overhead; external RSS and profiles cover the whole process.

The fixed protocol uses two reversed process repeats, three sizes, two builds
and two modes: 24 reports, each with 30 samples after three warmups, on CPU 2.
The four pilot reports contain one excluded sample per size, for 12 pilot
operations. Three separate profiler processes each request 30 samples after
three warmups; they are excluded from the formal matrix. Environment and build
flags are recorded in the [bundle](../results/change-0479/README.md).

The scalar positional source records request counts, requested/returned bytes
and six fixed request-size buckets. Large tail requests are valid; there is no
64 KiB source cap. The sequential hashing sink accepts at most 16,384 bytes per
write and retains no archive. These are logical in-memory I/O observations,
not physical storage or remote-source measurements. Allocation bytes include
the complete newly requested size of every reallocation, even when the system
allocator can resize in place. They do not measure memory-copy traffic.

The untimed oracle checks independently constructed main XML, full semantic
paragraph order/text, exact single-copy effect, physical member order, every
member's decompressed bytes and CRC, untouched compressed payloads, canonical
durable patch replay, source mismatch refusal and inverse restoration. The
whole-archive runtime digest is produced by the implementation under test and
then checked against the independent XML/member/semantic oracles; it is not an
independent whole-archive golden. The measured source is immutable.

## Validation and scope

Final release harness tests pass 389 tests with one ignored. The unchanged
DOCX paragraph-copy and removal suites pass 12 and 6 tests respectively.
Warning-denied Clippy, formatting, warning-denied rustdoc and crate boundaries
pass for the frozen source. The evidence suite passes 12 focused tests; both builds, all pilots, the
complete formal matrix and strict claim-registry validation pass. All receipts
and the final source ledger are retained in the bundle.
The [ADR matrix](../results/change-0479/adr-matrix.md) records the applicable
ownership, preservation and provider constraints.

This baseline does not demonstrate bounded-window append memory, an
optimization speedup, cold/range-source behavior or worker scaling. The full
non-iWork goal remains open. Measured results and the next production decision follow the complete
capture and independent review below.

## Measured baseline

Each timing row is one normal process with 30 samples. Times and the mean
95% t interval are milliseconds; RSS is the contextual whole-process maximum.
Percentiles use nearest rank. No repeats are pooled.

| Paragraphs | Repeat | Mean ms (95% interval) | p50 ms | p95 ms | p99 ms | RSS KiB |
| ---: | ---: | --- | ---: | ---: | ---: | ---: |
| 64 | 1 | 0.169641 (0.167587–0.171696) | 0.167321 | 0.183771 | 0.185671 | 4,548 |
| 64 | 2 | 0.169123 (0.167783–0.170463) | 0.167691 | 0.176081 | 0.180731 | 4,788 |
| 8,192 | 1 | 17.011582 (16.981149–17.042015) | 17.011344 | 17.127055 | 17.141115 | 10,924 |
| 8,192 | 2 | 16.251897 (16.221003–16.282791) | 16.235436 | 16.346136 | 16.607368 | 11,180 |
| 131,072 | 1 | 271.346097 (270.469410–272.222785) | 271.104147 | 274.986102 | 275.231225 | 113,448 |
| 131,072 | 2 | 271.230824 (270.210957–272.250690) | 270.497808 | 278.344301 | 278.702622 | 113,388 |

Allocator counters below are identical across all total samples and both
repeats for each size. Allocation calls include reallocation callbacks. Peak
is the region high-water mark minus entry live bytes; all 180 total allocator
samples return to their starting live-byte count with zero failed allocations.

| Paragraphs | Allocation calls | Reallocations | Requested bytes | Incremental peak bytes | Logical reads | Returned source bytes |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 64 | 1,888 | 67 | 1,036,969 | 509,974 | 42 | 7,969 |
| 8,192 | 213,346 | 16,453 | 1,615,057,345 | 3,071,310 | 42 | 71,716 |
| 131,072 | 3,653,986 | 507,973 | 549,272,923,585 | 41,793,870 | 62 | 1,033,480 |

The large normal phase means are:

| Repeat | Open ms | Snapshot ms | Stage ms | Commit ms | Publish ms | Drop ms |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 1 | 0.072526 | 56.870787 | 56.062721 | 0.000151 | 159.679907 | 0.833785 |
| 2 | 0.072440 | 56.043762 | 56.655528 | 0.000129 | 160.828549 | 0.860420 |

These phases are separate executions; their peaks cannot be added to obtain
the total peak. At 131,072 paragraphs the allocator phase owners retain
3,408 bytes after open, another 8,520,268 after snapshot, 8,519,952 after
stage, zero additional bytes after commit, and 10,613,492 after publication.
Drop releases the resulting 27,657,120 bytes. Publication is the largest
latency phase; committing already-staged data is a very short Arc handoff.

The small normal total process RSS rises from 4,548 to 4,788 KiB between
repeats (+5.277%), crossing the 5% review threshold. No normal total timing
mean crosses that threshold; the medium mean shifts −4.466%. All 38 timing/RSS
repeat flags across seven observation lanes are retained, including noisy
sub-microsecond commit/drop phases and tail percentiles. These are repeat
stability observations, with no candidate-versus-control regression claim.

The [allocation source audit](../results/change-0479/allocation-source-audit.md)
attributes over 99.99% of large snapshot/stage/publish requested bytes to
paragraph-range growth. After the initial 4,096 ranges, the parser requests
one additional slot with `try_reserve_exact(1)` whenever full. The sum of
requested new capacities matches the allocator counters; it does not prove
that the allocator physically copied those bytes. The publication path also
recaptures and scans source/candidate XML and clones target bytes.

## CPU and syscall attribution

The separate large normal `perf stat` process records 45,360,449,734 cycles,
190,361,776,641 instructions (4.19665 IPC), 38,029,370,097 branches,
142,458,513 branch misses (0.37460%), 10,174,103 generic cache misses and
37,201 page faults. Event running fractions are 100%; generic cache misses
are not an exact L1/LLC decomposition. This whole process includes corpus
construction, oracles, warmups, 30 measured lifecycles and report teardown.

The sampled process contains 1,012 target sample blocks and five helper
blocks. Its lifecycle ancestry accounts for 88.652% of target weighted
periods. The source-backed paragraph scanner within that ancestry accounts
for 68.037% of lifecycle periods (60.317% of target process periods), with
additional scanner work outside the lifecycle reported separately. These are
sampled inclusive subtree shares, not isolated phase timings or additive
leaf costs. The profile retains unresolved frames and sampling limitations.

The syscall profile includes 19,025 `write` calls across the wider process,
including report/output activity. The benchmark publication sink is in memory;
those syscalls must not be attributed to document sink writes. Raw counters,
weighted callstacks, no-children leaf rankings, exact recomputation and limits
are retained in the [CPU review](../results/change-0479/cpu-review.md) and
`profile-summary.json`.

## Final validation and next action

The final ledger retains 16 required successful gates, 30 attempts in total,
and all four development failures. Live source/template/binary checks bind
both executables to harness commit `3e8022336` and the same 7,048-file source
manifest. Final cleanup and portable receipts authenticate removal of runtime
copies and independently verify the sealed bundle with eight resealed
corruption probes. No raw formal capture was discarded or replaced.

The [next-work record](../results/change-0479/next-work.md) prioritizes a
measured shared-ownership publication experiment followed by scanner work while keeping the separately required
explicit-window tail API and broader non-iWork scenario coverage open.
