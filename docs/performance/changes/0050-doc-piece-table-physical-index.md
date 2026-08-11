# Change 0050: native DOC PieceTable physical index

Date: 2026-08-11

Production control: `473c458ada4c684a5d642b9b907fccac8298504c`

Scope: private native DOC parsing only. iWork/IWA crates were explicitly
excluded.

## Hypothesis and change

Fresh profiling attributed 36.89% of large DOC-open self cycles to
`PieceTable::fc_range_to_cp_ranges`. Paragraph and character FKP parsing call
that method for every physical formatting run. It scanned every logical text
piece on every call and sorted each result even though most physical pieces
could not overlap the requested FC interval.

`PieceTable` now builds one private FC-ordered index when it parses the CLX.
Each index entry stores the original CP-ordered piece position and the maximum
physical end seen through that entry. A query uses two binary searches:

1. the prefix maximum finds the first entry whose preceding intervals cannot
   all end before the requested start;
2. the FC ordering finds the first entry beginning at or after the requested
   end.

Only the bounded candidate slice is scanned. Every actual intersection still
uses the existing ANSI/UTF-16 FC-to-CP conversion, and the returned CP ranges
are still sorted exactly as before. The prefix maximum is required because
fast-save documents may contain physically overlapping or disordered pieces;
a simple start-FC binary search would be incorrect for those inputs.

This adds one private `O(piece_count)` index allocation and an `O(n log n)`
sort per parsed piece table. It changes no public type, transaction contract,
dependency edge, runtime, lock, global state, output byte, or validation
boundary. PAPX and CHPX owners still parse every reachable FKP and the public
DOC reader still performs its complete readback.

## Matched latency evidence

The frozen control and candidate binaries have SHA-256:

- control: `c8f9597076a8f1a00b4a3001546f1b031a9f5643b4dcbd763d51cacb4dd37f7e`;
- candidate: `fc19181881e5d92479ecee39a4f5f9c9a56d09aa9538fc6682adc8ab81d6343f`.

Both use the unchanged standalone harness, release profile, Rust 1.95.0,
Linux 6.8.0-101-generic, the Rust system allocator, and CPU 2 pinned with
`taskset`. The generated large DOC contains 512 paragraphs, is 97,792 bytes,
and has SHA-256
`3d96764fe48e213b972ff5921df183dab9e8bfc8c8e751bcf3bf20190de4fec6`.
Its 81,920-byte `WordDocument` stream has SHA-256
`33e6cd70a45181c28d4a3e7bfa4e7817bd82d7b2e89e39437a589243abdc38eb`.

The primary direct public `doc_semantic_open` measurement used 50 warmups and
500 samples in each of five control/candidate and five candidate/control
pairs. Pooling raw samples gives 5,000 observations per state while balancing
binary order. Corpus construction and complete semantic verification remain
outside the timer.

| Large DOC open | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 790.727 us | 348.679 us | **-55.91%** |
| mean | 800.571 us | 354.032 us | **-55.78%** |
| p95 | 878.549 us | 391.679 us | **-55.42%** |
| p99 | 1,030.869 us | 448.297 us | **-56.51%** |

The approximate independent-sample 95% interval for the mean delta is
`[-55.99%, -55.57%]`. All ten paired p50 comparisons improve, spanning
54.97% to 57.23%. A preceding conventional A/B/B/A run with 1,000 samples per
leg produced the same conclusion (798.728 to 344.238 us pooled p50), but its
candidate legs differed by 6.4%; the balanced ten-pair result is the retained
primary evidence rather than hiding that process drift.

The secondary large `doc_semantic_one_edit_save` ABBA used 30 warmups and 500
samples per leg, or 1,000 pooled samples per state. It includes the changed
commit and output materialization; exact output comparison, forward patch,
inverse restoration, strict snapshot reopen, independent public DOC reopen,
and complete semantic verification remain outside timing.

| Large DOC one paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 1.379 ms | 0.950 ms | **-31.08%** |
| mean | 1.408 ms | 0.962 ms | **-31.68%** |
| p95 | 1.606 ms | 1.072 ms | **-33.22%** |
| p99 | 1.928 ms | 1.216 ms | **-36.95%** |

The approximate independent-sample 95% interval for the mean delta is
`[-32.22%, -31.13%]`. Both matched p50 comparisons improve by more than 29%.

Primary raw reports are the
[`forward/reverse`](../results/) `doc-piece-index-open-{forward,reverse}-*`
files. The canonical four-leg open reports and secondary edit reports are
`doc-piece-index-{open,edit}-{before-a,after-a,after-b,before-b}.json` in the
same directory. Their hashes are indexed in
[`doc-piece-index-sha256.txt`](../results/doc-piece-index-sha256.txt).

## Profile attribution and resources

Matched 3,000-sample `perf record` processes directly confirm the removed
owner:

| Self-cycle frame | Before | After |
|---|---:|---:|
| `PieceTable::fc_range_to_cp_ranges` | 36.89% | 4.17% |
| `PieceTable::build_physical_index` | absent | 0.11% |

The reports contain 10,140 and 6,215 samples with zero lost samples. Kernel
symbols are restricted on this host, but the userspace DOC frames are
resolved. See [`before`](../results/doc-piece-index-perf-before.txt) and
[`after`](../results/doc-piece-index-perf-after.txt).

Matched process-wide `perf stat` A/B/B/A runs used the same 50 warmups and
1,000 measured opens per leg. The two legs per state sum to:

| Counter | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 7,183.59 ms | 4,318.16 ms | -39.89% |
| cycles | 34,980,550,582 | 21,160,617,950 | -39.51% |
| instructions | 168,106,175,225 | 85,956,339,106 | -48.87% |
| branches | 37,204,386,019 | 21,016,013,225 | -43.51% |
| branch misses | 60,045,636 | 41,793,988 | -30.40% |
| cache references | 2,377,314,057 | 1,933,200,519 | -18.68% |
| cache misses | 164,026,946 | 147,932,597 | -9.81% |
| page faults | 20,673 | 20,762 | +0.43% |
| context switches | 98 | 83 | -15.31% |
| CPU migrations | 0 | 0 | unchanged |

These are complete-process counters, so their reduction is smaller than the
scoped timer. The instruction, branch, cycle, and cache movements all agree
with removing repeated scans. The 89 additional minor faults are 0.43%, with
zero major-fault or migration concern.

Heaptrack used two warmups and 20 measured large opens per state:

| Whole-process metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 724,760 | 724,826 | +0.009% |
| Temporary allocations | 400,098 | 400,098 | unchanged |
| Peak heap | 5.67 MiB | 5.67 MiB | unchanged |
| Heaptrack RSS | 17.31 MiB | 17.51 MiB | +1.16% |
| Leaked bytes | 544 B | 544 B | unchanged |

The additional calls are the bounded physical indexes; their count and peak
memory stay well below the 5% review threshold. Uninstrumented GNU Time ABBA
reported 30,848/30,976 KiB before and 30,848/30,848 KiB after, so maximum RSS
does not regress. Raw counter, Heaptrack, and GNU Time files use the
`doc-piece-index-{perf,heaptrack,time}-*` prefix.

## Guardrails and correctness

Large guards pool 600 samples per state. Tiny open and changed edit pool 2,000
samples per state.

| Guard | p50 delta | Mean delta | p95 delta | Disposition |
|---|---:|---:|---:|---|
| Large list paragraphs | +0.32% | -1.98% | -13.94% | neutral/better |
| Large one paragraph | +0.83% | +0.14% | -2.61% | neutral |
| Large full text | +10 ns / +1.04% | +119 ns / +11.65% | +100 ns / +7.62% | sub-microsecond noise; disclosed |
| Large exact no-op edit/save | +0.04% | -0.06% | +1.45% | neutral |
| Tiny open | -1.36% | -7.34% | -5.54% | neutral/better |
| Tiny one edit/save | +0.39% | -0.02% | -2.16% | neutral |

The full-text timer starts from an already opened model and measures only a
roughly 1 us join; its p99 also moves from 1.592 to 2.063 us. The absolute
movement is retained explicitly and does not offset the 442 us public-open
reduction. Complete semantic verification still runs after every sample.

New unit tests compare the index against the former scalar algorithm for
discontiguous, physically overlapping and duplicate-start pieces, mixed ANSI
and UTF-16 pieces, odd physical boundaries, empty/out-of-range queries,
zero-length and saturating intervals near `u32::MAX`, and 1,024 deterministic
adversarial queries over 256 physical intervals. Existing real Word/LibreOffice
fixtures and generated writer corpora exercise the same package reader through
the complete suite.

Verification completed:

- `litchi-doc --all-targets --all-features`: 958 unit tests passed, two
  fixture-dependent tests remained ignored, and all integration/example
  targets passed;
- warning-denied all-target/all-feature DOC Clippy passed;
- all 32 standalone harness tests and warning-denied all-target Clippy passed;
- the DOC libFuzzer target compiles;
- formatting, JSON parsing, artifact hashes, and `git diff --check` pass.

Warning-denied DOC rustdoc remains blocked by pre-existing broken/private links
in `section/columns`, `shape`, `mtef_extractor`,
`document/model/semantic`, and `parts/text`. None of those files changed; the
same limitation is recorded in change 0017.

## Remaining work

This index accelerates physical FKP-to-logical-piece mapping. It does not
change CFB publication, exact patches, encryption/signature/protection policy,
real-producer coverage, final result copies, or the retained strict owner and
public-reader reopens. Future native DOC work must profile a distinct remaining
owner instead of removing either validation boundary or reviving the rejected
terminal-render/shared-payload/recapture handoffs.
