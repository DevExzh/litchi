# Change 0170: XLSX streaming UTF-8 escape runs

Date: 2026-08-17

## Decision

Retain the narrow `StreamingWorkbookWriter` encoder optimization. Ordinary
UTF-8 text is now appended as contiguous runs between the five XML entity
boundaries instead of one scalar at a time. Text whose UTF-8 byte count is at
most Excel's 32,767-character limit skips a redundant scalar-count pass, and
the row number is formatted once and reused by every cell in that row.

The worksheet XML, Deflate input/write boundaries, ZIP topology, archive
bytes, limits, cancellation points outside private scratch construction, and
public API are unchanged.

## Measured selection

The exact control revision's retained large-shape profile from change 0169
attributes 4.35% of process samples to `push_escaped`; 2.01% of process samples
reach `memmove` through that function, while `append_cell` contributes another
0.99% `memmove` path. Deflate remains the larger sampled cost, but changing its
policy would alter compressed bytes and cross format boundaries. The encoder
loop was therefore the next bounded same-output target.

The previous implementation called the finite row-scratch extender for every
ordinary Unicode scalar. The retained implementation validates characters in
the same source order, flushes an ordinary byte run before each entity or
invalid scalar, and checks the same row limit at each flush. A run cannot exceed
the row limit unless scalar emission would also exceed it. Failed rows remain
private, publish nothing, and can be retried at the same row coordinate.

## Correctness gates

- A scalar reference encoder is compared with the batched encoder across
  ordinary ASCII, all five entities, allowed controls, multibyte/astral text,
  invalid XML scalars, and every scratch ceiling from 0 through 128 bytes.
  Errors and successful bytes match.
- A rejected invalid-text row is retried at the same row number with entity,
  Unicode, numeric, Boolean, error, blank, and `Z`/`AA`/`ZZ`/`AAA` coordinate
  boundaries; the completed workbook reopens through the public facade.
- Existing exact/one-over text-byte, scalar-count, row-scratch, worksheet,
  row, cell, object/work, output, cancellation, sink-failure, deterministic
  non-seek, and fixed-window tests remain green.
- The focused streaming suite passes 31/31. Strict library/test Clippy passes
  with `-D warnings -D deprecated`. An independent adversarial review returned
  SAFE.

## Clean release ABBA

Control `c51bb95c8` and candidate `478bd9b2a` were built from clean detached
worktrees with locked release dependencies. The candidate diff is byte-for-byte
identical to integrated production commit `f2279b121`. Every process was pinned
to CPU 2 on the AMD EPYC 9575F host. The accepted matrix ran in strict
`A1, B1, B2, A2` order with 20 warmups and 300 samples for each tiny, medium,
and large shape.

Positive values below mean the candidate is faster:

| Shape | Statistic | A1 -> B1 | B2 -> A2 | Decision |
|---|---:|---:|---:|---|
| tiny | p50 | 7.74% | 5.03% | accepted |
| tiny | mean | 9.38% | **-3.15%** | withheld |
| tiny | p95 | 16.91% | **-17.60%** | withheld |
| tiny | p99 | 21.99% | **-4.29%** | withheld |
| medium | p50 | 5.52% | 5.03% | accepted |
| medium | mean | 5.06% | 4.64% | accepted |
| medium | p95 | 4.68% | 4.45% | accepted |
| medium | p99 | **-4.73%** | 2.36% | withheld |
| large | p50 | 5.66% | 5.02% | accepted |
| large | mean | 5.96% | 5.19% | accepted |
| large | p95 | 6.99% | 5.10% | accepted |
| large | p99 | 6.52% | 6.63% | accepted |

The predeclared same-implementation gates are 5% for p50/mean, 10% for p95,
and 15% for p99. All medium and large drift values remain inside those gates.
Tiny p50 also remains inside its gate; candidate tiny mean/p95/p99 drift does
not, matching the withheld scope.

All twelve primary records are clean and revision-exact. Each leg carries three
300-sample shape vectors. Archive hashes, decompressed
worksheet hashes, rows/cells/text bytes, accepted sink bytes, logical sink
write topology, zero retained output, and the 4 KiB authoring window are exact
across all legs. Expected-artifact construction, complete `Workbook` reopen,
and every-cell semantic verification remain outside timing.

An earlier 200-sample pilot began immediately after two parallel release builds
and exceeded same-implementation drift gates; it is deliberately excluded from
the retained artifact set and all claims.

## Process resource boundary

Matched large-shape `perf stat` processes use three warmups and twenty samples:

| Counter | A1 -> B1 | B2 -> A2 |
|---|---:|---:|
| cycles | 3.90% | 2.66% |
| instructions | 6.19% | 6.15% |
| branches | 10.57% | 10.54% |
| branch misses | **-8.99%** | **-14.37%** |

Branch misses regress in both directions and are disclosed. GNU Time maximum
RSS is 252,144/252,524/252,272/252,260 KiB for A1/B1/B2/A2, respectively.
These are whole-process observations containing setup, warmups, untimed reopen,
and harness overhead. No operation-local allocation, cache, branch-prediction,
RSS, peak-memory, total-memory, physical-I/O, or cold-cache improvement is
claimed.

## Scope and remaining work

This is warm, in-memory, synthetic, one-sheet inline-scalar creation through a
non-seek hashing discard sink. It does not establish multi-sheet, shared-string,
style, formula, date, real-producer, filesystem, or remote-source performance.
Deflate remains the dominant sampled writer cost; any compression-policy work
requires a separate output-size, deterministic-byte, cross-format, and security
review rather than silently changing the current writer.

Artifacts:

- [summary](../results/xlsx-stream-escape-0170-summary.json)
- [manifest](../results/xlsx-stream-escape-0170-manifest.json)
- [primary statistics](../results/xlsx-stream-escape-0170-primary-stats.tsv)
- [comparisons](../results/xlsx-stream-escape-0170-comparisons.tsv)
- [canonical semantic projection](../results/xlsx-stream-escape-0170-semantic.json)
- raw primary and process-counter JSON/sidecars listed in the manifest
