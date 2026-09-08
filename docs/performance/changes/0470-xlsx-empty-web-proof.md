# Change 0470: reuse compaction to validate empty worksheet web bindings

`performance_claim: none; descriptive latency probe and whole-process allocation evidence`

`claim_authorized: false`

Production candidate `f0ab67b55` attaches a bounded success proof to the exact
compacted worksheet bytes. Ordinary worksheets with no possible web grammar
can return the proven empty binding collection without another XML traversal
or its per-element namespace/name allocations. Every ambiguous case invokes
the unchanged full reader after optional grid parsing and before style checks.
Output bytes, typed errors, limits, no-op behavior, atomic publication and the
4,096-cell/1 MiB Store handoff remain unchanged. See the
[source review and ADR mapping](../results/change-0470/source-review.md).

The complete XLSX suite passes 1,257 tests across 59 targets: 950 library,
305 integration and two documentation tests. New differential cases compare
exact output, bindings, error phase, debug variants and display messages, with
competing faults, extension markup, namespace aliases, unknown prefixes,
malformed text/CDATA, DTDs, quote normalization and size/depth boundaries.
Scoped formatting, workspace all-feature checking, warning-denied XLSX Clippy
and rustdoc, and crate boundaries pass. All 30 reviewed ADR hashes remain equal.

The control reuses the authenticated 0469 candidate binary at revision
`7cc58fc1b`. Both roles build from `/tmp/litchi-goal-0468/profile-tree` with Rust
1.98.1, release debug level 1, frame pointers and unwind tables. Both run on
CPU 2 with one worker. Inventories bind 6,992 control and 6,993 candidate files,
the same two compile-time fixtures, and exactly four changed XLSX source/test
files. The benchmark harness, corpus, output verification and operation
boundaries are identical. Builds, tests, captures and postprocessing run
serially. The frozen protocol precedes all measurements.

The 100-sample/five-warmup six-row ABBA is diagnostic evidence below the
unchanged 500-sample registered latency minimum. It adds no registry entry.
The 201-row short guard has 15 samples and three warmups. Whole-process
Heaptrack uses dense one-percent commit/save, five samples and one warmup;
it includes generation, expected output, warmups, verification and teardown.
Its timing and RSS cannot be compared with normal runs, and its rounded heap
display is not an exact byte count.

## Allocation and normal timing observations

Whole-process allocation calls fall from 39,892,490 to 32,012,510: 7,879,980 fewer (19.753%). Temporary allocations fall
from 10,633,496 to 6,693,551 (37.052%). Both rounded peak
heap displays are `104.38M`. These are process totals, not per-commit
counts or exact peak-byte measurements.

Median observations, milliseconds:

| Ordinary commit/save | Shape | A1 | B1 | B2 | A2 |
| --- | --- | ---: | ---: | ---: | ---: |
| One cell | tiny | 0.227741 | 0.217636 | 0.219586 | 0.228486 |
| One cell | medium | 2.568515 | 2.424060 | 2.424915 | 2.553891 |
| One cell | dense-wide | 176.262829 | 167.618582 | 166.916931 | 175.697926 |
| One percent | tiny | 0.416496 | 0.392976 | 0.396436 | 0.413727 |
| One percent | medium | 10.161826 | 9.609171 | 9.795822 | 10.337905 |
| One percent | dense-wide | 354.187557 | 349.138418 | 335.701967 | 365.809447 |

All six rows pass the diagnostic directional/drift checks for mean, median,
p95 and p99. Dense one-percent median reductions differ between pairs
(1.426% / 8.230%), so a single number would overstate consistency. The
remaining five rows have median reductions of roughly 3.9%–5.6%. No registered
latency claim follows from this 100-sample probe.

## RSS investigation

The original normal process RSS is 110,844 / 119,736 / 120,656 / 111,080 KiB
in ABBA order. Both candidate pairs exceed the five-percent RSS review trigger.
The whole-process heap display does not resolve that flag. A lifetime review
found no added worksheet-sized ownership; it could not attribute the RSS
difference to a specific allocator, layout or timing mechanism.

A separately frozen repeat uses the exact same six-row 100/5 workload and
returns 112,476 / 110,868 / 110,576 / 120,636 KiB. The candidate changes are
-1.430% / -8.339%, while the repeated control itself drifts +7.255%. The first
RSS increase is not consistently reproduced. Both sequences remain retained;
no peak-memory improvement or blanket absence of memory regression is claimed.
The repeated reports also pass strict ABBA identity checks without projection.
Full-guard RSS is 159,260 / 158,152 KiB.

## Regression review and retention decision

The 201-row guard compares 1,205 metrics and retains 75 latency policy flags;
there are no non-latency metric flags. Applying a uniform five-percent review
trigger to every mean/p50/p95/p99 produces 131 flagged cells across 56 rows.
The complete lists remain in the original full-guard summaries.

A separately frozen 25-row 100/5 ABBA covers CFB read-one/shared-read-one,
OPC no-op save/source-open, and payload-heavy PPT creation across the selected
many-small, few-large and wide-root corpora. All captures and strict identity
checks pass without erasing counters or projecting reports. Identity validity
is not a latency acceptance claim. Several adverse observations persist:

- Payload-heavy PPT write p50 is 2.935868 / 3.093584 / 3.073779 / 2.927863 ms.
  Candidate median changes are +5.372% / +4.984%, mean +5.324% / +5.444%,
  p95 +10.222% / +10.564%, and p99 +12.019% / +13.144%.
- Many-small CFB read-one medians are 540 / 580 / 580 / 540 ns for compressible
  data and 540 / 580 / 580 / 530 ns for incompressible data. Those small absolute
  costs still exceed the relative review trigger and remain flagged.
- Incompressible many-small OPC no-op save is 2.690 / 2.710 / 3.180 / 2.710 us:
  the original larger flag is not stable between candidate runs, but the second
  pair remains adverse. Incompressible few-large OPC source-open has median
  changes of +14.894% / +2.671%. The full targeted summary retains all tails,
  drift checks and individual rows.

Retain the bounded pass reuse for the substantial measured reduction in allocation
work and lower latency in all six ordinary XLSX rows. This is a scoped
tradeoff, not a claim that the binary is uniformly faster: the persistent PPT
and CFB flags limit cross-format conclusions. Their compiled source is unchanged,
but that alone does not dismiss observed process/compiler/layout effects. The
original RSS flag also remains visible alongside the contradictory repeat.
No geometric mean is used to conceal any individual result.

The guard's first candidate preflight found two truncated files in the temporary
checkout before starting the benchmark. The failure and diff were retained;
the files were restored from the bound revision and all 6,993 source hashes,
fixtures and the immutable binary were reauthenticated. Only the pending
B1/B2/A2 captures then ran. No successful capture was replaced, and the cause
of the temporary checkout truncation remains unassigned.

Further work should address larger eager-parser/snapshot traversals and measure
release of the no-longer-needed pre-compaction vector before grid parsing. The
lifetime review identifies that vector in both versions; it is an independent
memory opportunity, not an explanation claimed for this experiment's RSS.
The broader non-iWork program remains open. No new registered latency, exact
peak-memory, native Office, fuzz-campaign, physical-cold/range or scaling claim
is made.

See the [evidence bundle](../results/change-0470/README.md),
[original normal/heap results](../results/change-0470/summary.json),
[RSS repeat](../results/change-0470/rss-review-summary.json), and
[targeted guard](../results/change-0470/guard-review-summary.json).
