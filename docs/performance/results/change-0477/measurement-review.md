# Change 0477 measurement review

Status: final capture review, 2026-09-08 UTC. The 48 reports and
`summary.json` were independently read after capture; no build, capture, or
analyzer rerun was needed for this review. The summary SHA-256 is
`2e62582716ffc10bcdbe3eb216db95a7582dbc1a31d6b1368dae223b7ba5c4b1`.
The capture set is append-only; an unfavorable observation or review flag is
evidence to preserve, not a reason to delete or replace a capture.

## What the protocol compares

The experiment is a same-source comparison of two low-level
`ZipArchiveWriter` storage policies. The control calls
`with_capacity(member_count).build` and retains the central directory in its
existing in-memory header/name structures. The spool lane passes an explicit
caller-owned read/write/seek `File` and the frozen `64 MiB` maximum spool and
`16 KiB` replay buffer limits to `build_with_spool`.

The frozen matrix has 48 isolated processes: normal and allocator
instrumentation, 8/256/8192 members, Store and Deflate, control and spool,
and two external repeats. Each process uses 30 measured samples and three
warmups, is pinned to CPU 2, and runs one internal repeat. Repeat two reverses
the complete case order. Every process receives an explicit `--spool-dir`,
including control-only processes, so both policy oracles and their validation
are part of the matched process preflight. Only the selected policy enters the
timed operations.

This design supports a descriptive policy comparison at these synthetic
low-level counts. It does not support a historical speedup claim, a default
configuration claim, or an end-to-end DOCX/PPTX/OPC claim. The preceding
0476 evidence is historical context for the enabler and is not a second arm
of this experiment.

## Timing and correctness boundaries

Corpus construction, expected archive generation, and all reopen oracles are
outside `elapsed_ns`. The control and explicit-spool oracle bytes are compared
exactly, reopened with `ArchiveReader`, and checked member by member against
the deterministic source payload. Measured samples must match that oracle in
both byte count and SHA-256 digest; this is a correctness gate, not a timed
operation.

The timed interval starts after the operation's process and allocator
before-snapshots, constructs the fixed `HashingSink`, and owns the complete
writer lifetime. It includes explicit spool `File` creation/open, each local
header and member write, data descriptors, central-directory publication,
sink acceptance and digest updates, writer flush, and spool `File` close.
Digest finalization, metric endpoint snapshots, oracle construction/reopen,
and unlink cleanup are outside the interval. The benchmark does not call
`sync_all`, so the result describes this caller-visible operation and logical
file activity; it makes no durability, physical-device, or filesystem
throughput claim.

The hashing sink does not retain the generated archive. The oracle path does
materialize bytes in a `Vec`, but that materialization is deliberately outside
the timed interval. Therefore timed elapsed values should be read as writer
and sink work under a fixed non-retaining output sink, while process-wide RSS
still includes the setup, both policy oracles, validation, and any output
materialization that occurred earlier in that process.

## How to read the resource fields

Allocator rows use the operation-scoped counting allocator. `allocated_bytes`
and allocation calls are requested allocation activity; they are not physical
memory traffic. `region_peak_live_bytes - live_bytes_before` is the allocator
region high-water increase. The allocator fields do not measure the bytes
occupied by the caller's spool file or tmpfs page cache. The normal binary has
no allocator sample, so allocator and normal timings are never pooled.

`/usr/bin/time -v` maximum RSS is a whole-process high-water value. In this
protocol it includes matched preflight/oracle work and retained process state,
so it is useful as a process resource observation but cannot be interpreted as
the timed writer's heap peak. The operation process delta is bracketed around
the timed call, but `/proc/self/io` and related procfs reads are performed by
the observer itself. In particular, the after snapshot can add `rchar` and
`syscr`; `rchar`, `wchar`, `read_bytes`, `write_bytes`, `syscr`, and `syscw`
are logical process counters rather than physical-storage attribution.

The spool implementation serializes one central-directory record into a fresh
per-record `Vec` and writes it to the caller's file. The fixed replay buffer
bounds the replay window; it does not bound the cumulative spool extent or
the benchmark's corpus/index structures. Per-record allocation and file-write
work are expected tradeoffs of this explicit policy and should remain visible
in allocator, write-call, and process-counter fields. `spool_bytes` is the
logical serialized central-directory extent, excluding the replay buffer; for
the frozen corpus it is checked against
`member_count * (46 + len("ppt/slides/slide-00000.xml"))`.

The captured spool path is under `/tmp`, which this environment records as a
`tmpfs`. A tmpfs file consumes shared system memory and temporary-filesystem
capacity while it exists even when the library's process heap stays flat. The
benchmark removes each owned file after endpoint snapshots, so this is a
transient system-wide storage cost rather than a retained artifact. The
measurement therefore cannot be used to claim that spooling has no memory or
storage cost merely because allocator live bytes or process RSS is unchanged.

## Captured result and independent verification

The raw set contains 48 report JSON files, 48 `/usr/bin/time` resource files,
and 1,440 measured operation rows. An independent read-only cross-check
verified every report's 30 rows, selected policy, instrumentation identity,
two oracle records, exact control/spool output agreement, member reopening,
source identity, and successful measured-output oracle match. It also
recomputed the summary statistics for all 48 rows and recomputed the review
flag sets: 20 of 24 policy-pair rows have at least one positive field and 6
of 24 repeat-drift rows have at least one absolute field above 5%.

The measured spool extent scales exactly with the 72-byte central record used
by this corpus. The output bytes are identical between policies for each
method and count:

| members | logical spool bytes | Store output bytes | Deflate output bytes |
| ---: | ---: | ---: | ---: |
| 8 | 576 | 3,222 | 3,318 |
| 256 | 18,432 | 102,422 | 105,494 |
| 8,192 | 589,824 | 3,276,822 | 3,375,126 |

All 1,440 measured outputs matched the corresponding oracle byte count and
SHA-256. The six control/spool archive digests for those rows are retained in
the raw reports; no digest or sample was substituted during review.

## Observed normal timings

The following are the summary's exact per-process means and p95 values in
nanoseconds. Each arrow is control to spool; the parenthesized value is the
paired percentage change for that external repeat.

| members | method | repeat | mean control → spool | p95 control → spool |
| ---: | --- | ---: | ---: | ---: |
| 8 | Store | 1 | 3,275.667 → 7,687.333 (+134.680%) | 3,390 → 9,230 (+172.271%) |
| 8 | Store | 2 | 3,198.667 → 7,659.000 (+139.444%) | 3,290 → 8,130 (+147.112%) |
| 8 | Deflate | 1 | 81,811.667 → 87,573.767 (+7.043%) | 89,050 → 96,830 (+8.737%) |
| 8 | Deflate | 2 | 81,581.700 → 87,967.167 (+7.827%) | 86,761 → 97,221 (+12.056%) |
| 256 | Store | 1 | 96,177.100 → 135,036.600 (+40.404%) | 96,491 → 140,451 (+45.559%) |
| 256 | Store | 2 | 97,195.800 → 139,285.600 (+43.304%) | 111,771 → 149,781 (+34.007%) |
| 256 | Deflate | 1 | 3,485,331.933 → 3,557,197.867 (+2.062%) | 3,495,105 → 3,628,806 (+3.825%) |
| 256 | Deflate | 2 | 3,486,121.867 → 3,555,381.967 (+1.987%) | 3,498,475 → 3,564,516 (+1.888%) |
| 8,192 | Store | 1 | 3,155,941.467 → 4,360,752.067 (+38.176%) | 3,170,744 → 4,409,299 (+39.062%) |
| 8,192 | Store | 2 | 3,075,537.233 → 4,292,421.433 (+39.567%) | 3,098,653 → 4,306,088 (+38.966%) |
| 8,192 | Deflate | 1 | 109,401,029.233 → 113,265,122.200 (+3.532%) | 110,984,555 → 113,727,507 (+2.471%) |
| 8,192 | Deflate | 2 | 110,894,336.400 → 112,862,053.900 (+1.774%) | 111,699,710 → 114,369,902 (+2.391%) |

Store shows a repeatable positive elapsed difference at all three counts,
largest at eight members and approximately 38–44% at 256 and 8,192 members.
Deflate is approximately 2% at 256 and 2–4% at 8,192; its eight-member
rows are approximately 7–8%. These are observations of this low-level
synthetic operation. The protocol registers no latency claim and these values
must not be presented as a historical or default-configuration speedup.

## Memory, allocation, and scratch scaling

The `r1` process maximum RSS values below are KiB and are whole-process
high-water observations. They include matched oracle setup and output
materialization, as defined above; the allocator column is a separate binary
and is not pooled with normal timings.

| members | method | normal RSS control → spool | allocator RSS control → spool |
| ---: | --- | ---: | ---: |
| 8 | Store | 2,500 → 2,456 | 2,484 → 2,480 |
| 8 | Deflate | 2,724 → 2,720 | 2,740 → 2,804 |
| 256 | Store | 2,708 → 2,740 | 2,728 → 2,996 |
| 256 | Deflate | 2,960 → 2,968 | 2,956 → 2,976 |
| 8,192 | Store | 17,040 → 16,332 | 17,004 → 16,984 |
| 8,192 | Deflate | 16,868 → 16,884 | 17,124 → 16,876 |

The allocator means below are requested allocation calls/bytes and operation
region peak live-byte increases, all from the `r1` allocator reports:

| members | method | calls control → spool | requested bytes control → spool | incremental peak live bytes control → spool |
| ---: | --- | ---: | ---: | ---: |
| 8 | Store | 5 → 19 | 1,350 → 17,260 | 1,168 → 16,574 |
| 8 | Deflate | 7 → 21 | 414,310 → 430,220 | 414,128 → 429,534 |
| 256 | Store | 10 → 515 | 44,006 → 41,564 | 37,376 → 16,574 |
| 256 | Deflate | 12 → 517 | 456,966 → 454,524 | 450,336 → 429,534 |
| 8,192 | Store | 15 → 16,387 | 1,408,998 → 819,292 | 1,196,032 → 16,574 |
| 8,192 | Deflate | 17 → 16,389 | 1,821,958 → 1,232,252 | 1,608,992 → 429,534 |

The per-record spool `Vec` and its file writes are visible in the very large
increase in allocation calls: 19 versus 5 at eight Store members and 16,387
versus 15 at 8,192. Requested bytes and live high-water have different shapes
because the control retains its growing in-memory directory while the spool
lane repeatedly allocates and releases record buffers. The spool lane still
has a cumulative file extent and a fixed replay buffer; a lower process heap
high-water does not mean that the operation has no memory or storage cost.

## Sink and process I/O observations

The following means come directly from the normal `r1` raw operation rows.
Each cell is `output_write_calls / syscw / rchar / wchar / write_bytes` for
control or spool. `output_write_calls` counts calls accepted by the hashing
sink. The procfs fields include observer effects; `wchar` is logical process
write input and `write_bytes` is the kernel's physical-storage counter.

| members | method | control | spool |
| ---: | --- | --- | --- |
| 8 | Store | 55 / 0 / 1,972.500 / 0 / 0 | 40 / 8 / 2,551.867 / 576 / 0 |
| 8 | Deflate | 63 / 0 / 1,972.500 / 0 / 0 | 48 / 8 / 2,551.867 / 576 / 0 |
| 256 | Store | 1,543 / 0 / 1,976.667 / 0 / 0 | 1,033 / 256 / 20,406.133 / 18,432 / 0 |
| 256 | Deflate | 1,799 / 0 / 1,976.900 / 0 / 0 | 1,289 / 256 / 20,406.367 / 18,432 / 0 |
| 8,192 | Store | 49,159 / 0 / 1,982.467 / 0 / 0 | 32,811 / 8,192 / 591,811.700 / 589,824 / 0 |
| 8,192 | Deflate | 57,351 / 0 / 1,984.933 / 0 / 0 | 41,003 / 8,192 / 591,809.267 / 589,824 / 0 |

Spool `syscw` and `wchar` track one central-record write per member and the
logical spool extent. `write_bytes` stayed zero because this run used the
recorded tmpfs scratch path; that is evidence about this environment's cache
path, not a physical-disk result. The control `rchar` baseline is mostly
procfs observer activity, while the spool replay adds reads of the central
directory. The allocator lane reproduced the same member-scaled `syscw`,
`wchar`, and sink-call pattern; its procfs `rchar` differs only by small
observer/instrumentation effects.

## Every paired and drift review flag

The analyzer threshold is a screening rule: a paired field is listed when
spool is more than 5% above control, while a drift field is listed when the
absolute repeat-one to repeat-two change exceeds 5%. The 24 paired rows are
12 per instrumentation lane, and the 24 drift rows are 12 per lane. All flags
were independently recomputed from `summary.json`; the raw rows remain
unchanged.

### Positive policy-pair fields

| instrumentation | members | method | repeat | fields above 5% (spool vs control) |
| --- | ---: | --- | ---: | --- |
| normal | 8 | Store | 1 | elapsed mean +134.680%, p50 +124.159%, p95 +172.271%, p99 +338.192% |
| normal | 8 | Store | 2 | elapsed mean +139.444%, p50 +122.956%, p95 +147.112%, p99 +584.545% |
| normal | 8 | Deflate | 1 | elapsed mean +7.043%, p50 +6.485%, p95 +8.737%, p99 +6.837% |
| normal | 8 | Deflate | 2 | elapsed mean +7.827%, p50 +7.084%, p95 +12.056%, p99 +13.491%, RSS +10.815% |
| normal | 256 | Store | 1 | elapsed mean +40.404%, p50 +39.662%, p95 +45.559%, p99 +36.897% |
| normal | 256 | Store | 2 | elapsed mean +43.304%, p50 +43.724%, p95 +34.007%, p99 +41.619% |
| normal | 8,192 | Store | 1 | elapsed mean +38.176%, p50 +38.122%, p95 +39.062%, p99 +39.939% |
| normal | 8,192 | Store | 2 | elapsed mean +39.567%, p50 +39.591%, p95 +38.966%, p99 +38.354% |
| allocator | 8 | Store | 1 | allocation calls +280.000%, requested bytes +1,178.519%, incremental peak +1,319.007% |
| allocator | 8 | Store | 2 | allocation calls +280.000%, requested bytes +1,178.519%, incremental peak +1,319.007% |
| allocator | 8 | Deflate | 1 | allocation calls +200.000% |
| allocator | 8 | Deflate | 2 | allocation calls +200.000% |
| allocator | 256 | Store | 1 | allocation calls +5,050.000%, RSS +9.824% |
| allocator | 256 | Store | 2 | allocation calls +5,050.000% |
| allocator | 256 | Deflate | 1 | allocation calls +4,208.333% |
| allocator | 256 | Deflate | 2 | allocation calls +4,208.333%, RSS +7.266% |
| allocator | 8,192 | Store | 1 | allocation calls +109,146.667% |
| allocator | 8,192 | Store | 2 | allocation calls +109,146.667% |
| allocator | 8,192 | Deflate | 1 | allocation calls +96,305.882% |
| allocator | 8,192 | Deflate | 2 | allocation calls +96,305.882% |

The absence of a positive flag for normal Deflate at 256 and 8,192 members
does not establish equivalence; it only means those paired fields did not
cross this review threshold in these samples.

### Repeat-drift fields

| instrumentation | members | method | policy | absolute repeat drift above 5% |
| --- | ---: | --- | --- | --- |
| normal | 8 | Store | spool | p95 -11.918%, p99 +50.299% |
| normal | 8 | Deflate | spool | RSS +10.000% |
| normal | 256 | Store | control | p95 +15.836%, p99 +6.878% |
| normal | 256 | Store | spool | p95 +6.643%, p99 +10.564% |
| allocator | 256 | Store | spool | RSS -9.880% |
| allocator | 256 | Deflate | spool | RSS +9.140% |

The repeat flags are retained as stability evidence. They do not authorize
selecting the favorable repeat or treating a screening threshold as a
statistical regression test. The summary's `t(29)` intervals and nearest-rank
percentiles describe the 30 samples within each isolated process; the two
external repeats are separate process observations.

## Claim limits

The captures support a descriptive comparison of the explicit caller-owned
central-directory spool and the existing in-memory policy for this synthetic
low-level writer at 8, 256, and 8,192 members, under Store and Deflate, with
the stated limits and tmpfs scratch path. They support reporting exact ZIP
output parity, member reopen coverage, logical spool extent, requested
allocation activity, operation allocation high-water, process RSS, sink-call
counts, and procfs counter observations.

They do not support a historical/default speedup claim, a durability or
physical-I/O claim, or an end-to-end DOCX/PPTX/OPC memory guarantee. Higher
layers retain their own ZIP/OPC indexes and source structures, and the
benchmark's tmpfs file remains a shared system resource even when process
heap metrics are lower. Any final performance narrative must retain the
negative and positive flags and state these boundaries.
