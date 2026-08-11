# Change 0055: bounded RTF body style-block reservation

Date: 2026-08-11

Status: accepted

## Decision

Reserve the parser's root-body `StyleBlock` vector once, immediately before
the first retained block, when the existing structural preflight proves a
large ordinary body. Keep the original lazy geometric growth for small,
table-heavy, and deletion-heavy documents.

The reservation is private to `litchi-rtf`. It adds no public type, cache,
runtime, lock, dependency, unsafe code, or persisted state.

## Problem and attribution

The generated large plain corpus contains 10,000 paragraphs in 540,051 source
bytes. It produces exactly 10,000 `StyleBlock` values; each value is 984 bytes
in this build. The control parser let `Vec<StyleBlock>` grow to capacity 16,384,
requesting 16,121,856 bytes through 12 allocations per parse.

The exact control profile attributed 25.65% of cycles to `memmove`, including
the `RawVec::grow_one -> flush_text_buffer` ancestry. Across 22 instrumented
parses, that vector performed 264 growth allocations and consumed 16.12 MiB at
peak.

## Implementation

The parser now:

- counts nonempty root text tokens inside its existing full structural pass;
- enables the hint only for sources of at least 64 KiB;
- bounds capacity by the observed root-text count, `max_tokens`, 24 times the
  source byte length, and an absolute 16 MiB reservation ceiling;
- disables the hint when root table/nested-table or active-deletion controls
  make root text tokens a poor upper estimate of retained body blocks;
- consumes the hint only when the first body block is about to be retained;
- uses a fallible speculative reserve, then falls back to a fallible one-block
  reserve if the larger optional request fails.

The root-text counter is unconditional once a text token reaches the existing
root arm. This avoids a document-size-dependent branch on every medium text
token. The count cannot exceed the already materialized token vector length.

All three body insertion paths use the same helper: ordinary decoded text,
semantic text, and Unicode control output. Documents with no retained body
block allocate nothing.

## Safety and semantic boundaries

The change alters capacity only. Block values, order, formatting, paragraph
state, text lengths, revision ranges, source spans, and writer output are
unchanged.

The following boundaries remain:

- the lexer and parser still perform their complete validation passes;
- parser depth, token, text, table, revision, and allocation limits are
  unchanged;
- exact no-op source sharing and changed-candidate parse/readback remain;
- table and deletion inputs keep lazy growth rather than speculatively
  retaining a large body allocation;
- sources below 64 KiB keep the original allocation path;
- reservation failure cannot turn optional speculation into a large mandatory
  allocation: the parser retries only the one block it is about to append;
- compressed LZFu, raw CP-1252, watermark, tables, deletions, malformed input,
  durable patch/inverse, and transaction tests retain their existing behavior.

## Measurement method

Control revision: `3afe6c9610e01b56c385b92e41582f5ca7a9b9d5`.

Exact release binaries:

- before: `531af2133a9af3aafb51db0ac1c946649e66b4b9fe34651171d66645cc2322f5`
- after: `d23e562184fd71f6e9032de63d98707335ef61f76cc1afc7a1cc482707ea6836`

The headline open and one-edit/save measurements use six balanced pairs, one
case per process, pinned to CPU 9. Every leg has 100 warm-ups and 1,000 timed
samples. The order is before/after/after/before and is mirrored through all six
pairs, yielding 6,000 samples per state. The exact no-op guard uses two
balanced pairs and 2,000 samples per state.

The raw pooled samples, distribution statistics, environment, corpus digest,
and binary hashes are in
[`rtf-body-block-reservation-primary-summary.json`](../results/rtf-body-block-reservation-primary-summary.json).

## Latency result

| Case | Before p50 | After p50 | p50 | Mean | p95 | p99 |
|---|---:|---:|---:|---:|---:|---:|
| Large plain open | 2.073 ms | 1.634 ms | **-21.17%** | **-21.00%** | **-21.04%** | **-24.62%** |
| Large one-edit/save | 5.585 ms | 5.503 ms | **-1.46%** | **-1.75%** | **-1.87%** | **-4.11%** |
| Large exact no-op guard | 16.353 us | 15.461 us | **-5.45%** | **-3.13%** | +1.78% | +16.17% |

The no-op tail movement is disclosed. This microsegment is tens of
microseconds, its center improves, and the complete edit/save distribution also
improves through p99; no no-op claim is made from the noisy tail.

## Medium and variant guards

Each medium variant also uses six balanced pairs and 6,000 samples per state.
The complete distributions are in
[`rtf-body-block-reservation-medium-guards-summary.json`](../results/rtf-body-block-reservation-medium-guards-summary.json).

| Medium open | p50 | Mean | p95 | p99 |
|---|---:|---:|---:|---:|
| Plain | +0.49% | +1.16% | +4.80% | +5.10% |
| Raw CP-1252 | +2.84% | +1.01% | -8.84% | -1.35% |
| LZFu | -0.09% | -0.62% | -3.14% | -4.20% |

The small plain/CP-1252 center costs are retained and disclosed rather than
hidden. These sources are below the 64 KiB reservation threshold, so they keep
the prior geometric-growth behavior; the residual movement is the root-token
counter and compiler layout. A 25-sample capability smoke covers all supported
tiny plain/CP-1252/LZFu/watermark open/no-op/edit combinations in
[`rtf-body-block-reservation-tiny-variant-smoke.json`](../results/rtf-body-block-reservation-tiny-variant-smoke.json).

## Allocation and memory evidence

Matched Heaptrack runs use two warm-ups and 20 measured large opens, for 22
parses per process.

| Metric | Before | After | Change |
|---|---:|---:|---:|
| Body block-vector allocations | 264 | 22 | 12 geometric allocations become one exact reserve per parse |
| Process allocation calls | 962,144 | 961,844 | -300 (-0.031%) |
| Peak heap | 21.12 MiB | 14.84 MiB | **-29.73%** |
| Heaptrack-inclusive RSS | 39.67 MiB | 34.60 MiB | -12.78% |
| Leaked bytes | 544 | 544 | flat |

The retained block allocation is 9.84 MiB, exactly 10,000 blocks, instead of
the control vector's 16.12 MiB geometric capacity. Full reports are
[`before`](../results/rtf-body-block-reservation-before-heaptrack.txt) and
[`after`](../results/rtf-body-block-reservation-after-heaptrack.txt).

Uninstrumented matched GNU Time runs report 30,976/30,848 KiB before and
30,848/30,976 KiB after: peak RSS is flat at the tool's resolution. The four
raw reports are stored under `results/rtf-body-block-reservation-time-*.txt`.

## CPU evidence

Matched 500-sample `perf stat` processes report:

| Counter | Before, pooled A+B | After, pooled A+B | Change |
|---|---:|---:|---:|
| Task clock | 3,947.61 ms | 3,357.95 ms | **-14.94%** |
| Cycles | 19.290 billion | 16.414 billion | **-14.91%** |
| Instructions | 70.690 billion | 70.557 billion | -0.19% |
| Branches | 15.443 billion | 15.450 billion | +0.05% |
| Branch misses | 8.012 million | 8.174 million | +2.02% |
| Cache misses | 100.339 million | 68.031 million | **-32.20%** |

The exact `perf record` profiles reduce `memmove` from 25.65% of 9.60 billion
sampled cycles to 19.97% of 8.12 billion. The former style-block
`grow_one -> flush_text_buffer` stack is replaced by one
`try_reserve_additional` call per parse. Reports are
[`before`](../results/rtf-body-block-reservation-before-perf-report.txt) and
[`after`](../results/rtf-body-block-reservation-after-perf-report.txt); raw
counters are stored under `results/rtf-body-block-reservation-perf-stat-*.csv`.

## Rejected refinements

Two refinements were measured and removed before the final binary:

- guarding the root counter on every text token kept exact sizing but regressed
  medium plain open by 5.22% p50 / 6.56% mean;
- moving estimation into a second large-source token scan removed the medium
  branch but erased the edit/save center gain and moved its p99 +9.39%.

The accepted version keeps one structural pass and an unconditional root-text
increment, which preserves exact sizing without either rejected tradeoff.

## Validation

Passed on the final source:

- `cargo test -p litchi-rtf --release --all-targets --all-features` (303 unit
  tests plus every integration/example target);
- `cargo clippy -p litchi-rtf --all-targets --all-features -- -D warnings`;
- `RUSTDOCFLAGS="-D warnings" cargo doc -p litchi-rtf --all-features --no-deps`;
- RTF fuzz-bin compilation;
- the 32-test performance harness and its warning-denied all-target Clippy;
- formatter and whitespace checks.

The earlier ODF `generic-array::clone_from_slice` deprecation fix in commit
`1194fbc7f` was also rechecked with warning-denied all-target/all-feature
`litchi-odf-common` Clippy and rustdoc. Both pass.

## Limitations and next work

This is a generated text-heavy warm-memory corpus. It does not add formatting,
media, security/repair, cold-source, or broader real-producer CRUD coverage.
Table/deletion-heavy documents deliberately retain the previous allocation
path. The next RTF work should come from a new profile rather than increasing
the reservation multiplier or weakening those guards.
