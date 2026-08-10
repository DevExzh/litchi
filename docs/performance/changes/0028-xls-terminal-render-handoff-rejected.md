# XLS terminal validated-render handoff is rejected

Date: 2026-08-11

Production base: `cd20a90ad0168ecb624e5d011328b91595f2db81`

Disposition: measured and fully reverted. OLE2, XLS, OOXML, RTF and ODF
production code are unchanged by this record. iWork/IWA was excluded.

## Hypothesis and prototype

An effective XLS cell transaction replaces the Workbook stream through the
common OLE2 object editor. `Editor::commit_candidate` renders the candidate,
reopens it, checks the container, recaptures every stream, reuses validated
stream allocations and rediscovers targets. `Snapshot::from_package_editor`
then calls `Editor::finish`, which renders the same recaptured package again
before the strict BIFF owner parse and independent public `Workbook` reopen.

The prototype returned the first validated rendering only to the immediately
following XLS snapshot constructor. It retained no editor-wide byte cache and
did not change ordinary `put_stream_shared`, DOC/PPT lifecycles, protection
checks, CFB reopen/recapture, strict BIFF parsing, public-reader validation,
typed readback, structural/resource readback, patch/inverse behavior or exact
no-op ownership.

Focused common-layer tests proved that the handed-off bytes exactly equaled a
fresh render of the recaptured editor, reopened as CFB, preserved the changed
stream, returned exact source bytes for a no-op and left a failed replacement
atomic. A genuine XLS test changed an existing BIFF `Number`, proved exact
render equality, reopened the supplied bytes through the complete snapshot
owner and retained an opaque stream. All focused tests passed before the
prototype was reverted.

## Profile and matched latency

The common-harness baseline executable SHA-256 is
`8c92c33ebe285bf5a3138c90e3eea5e04b4ffa8536da61cfe572f473c84ff40e`.
The rejected prototype executable SHA-256 is
`3df0b79d648172bd2c2d03a99999265576c2f7dbd92a9d8ab27647e87cd305a5`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 4 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic artifacts are byte-identical across
the executables: tiny is 4,096 bytes with SHA-256
`cdc133bd87aaa60a91ea5e94df6ff8da0eb6bb0f2432afa4bfdb13cf70c0298b`;
large is 163,840 bytes with SHA-256
`228c6585a4d26141aebfaf7b08844a2ee445b269d406006a1fdb0484619120fb`.

The baseline sampled profile used 30 warmups and 3,000 large changed-save
samples. The harness deliberately runs complete verification after every timed
operation, so the process profile is dominated by that verifier: 50.19% of
children are below the harness runner and 43.15% are below
`verify_semantic_xls`; unresolved call chains separately place 48.61% in the
ordinary semantic worksheet parser. The changed `Transaction::commit` subtree
is only 1.86% of the process and its `Snapshot::from_package_editor` subtree
is 1.26%. The renderer is inlined and has no trustworthy separate sampled
frame. Source ownership still proves the duplicated render; the direct
prototype is the materiality experiment rather than a claimed renderer-only
profile share.

Primary A-B-B-A used 30 warmups and 1,500 samples per leg. Pooling 3,000
samples per state gives:

| One-cell edit/save | Before p50 | Prototype p50 | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|---:|---:|
| Tiny | 24.275 us | 22.442 us | **-7.55%** | **-9.26%** | -20.75% | -17.59% |
| Large | 1.543 ms | 1.537 ms | -0.39% | -0.70% | -0.93% | +0.89% |

The tiny result is useful in isolation, but the large result is neutral and
does not justify a new cross-crate terminal API by itself.

## Rejection guard and memory

A 1,500-sample-per-leg open/no-op A-B-B-A kept ordinary open within the 5%
p50/mean gate: tiny p50/mean moved -0.10%/-1.81%, while large moved
+2.18%/+3.23%. Large open tails moved +8.50% p95 and +19.90% p99 and are
retained as a secondary variance trigger.

The large exact no-op guard was not acceptable. The high-sample A-B-B-A moved
p50 **+13.91%** and mean **+18.36%**. Four additional short A-B-B-A cycles,
which pool 2,000 samples per state, reproduced p50 **+22.00%** and mean
**+16.69%**; p95/p99 moved +8.83%/+5.07%. The absolute change is about
0.4-0.5 us and the no-op branch cannot call the terminal method, so binary
layout/allocation alignment is the likely mechanism. That does not make the
measured regression acceptable. The tiny no-op repeat stayed near neutral at
+0.91% p50/+4.35% mean.

Matched Heaptrack processes used five warmups and 100 large changed saves:

| Metric | Before | Prototype | Delta |
|---|---:|---:|---:|
| Allocation calls | 1,121,669 | 1,117,959 | -3,710 (-0.33%) |
| Temporary allocations | 98,189 | 97,659 | -530 (-0.54%) |
| Peak heap | 8.12 MiB | 8.12 MiB | unchanged |
| Heaptrack RSS | 19.95 MiB | 20.18 MiB | +1.15% |
| Leaked bytes | 544 B | 544 B | unchanged |

The allocation reduction supports the source-level hypothesis but does not
outweigh the no-op regression or the neutral large changed-save latency. No
hardware-counter or uninstrumented-RSS acceptance claim was pursued after the
latency rejection.

## Final state and next work

The common terminal method, XLS snapshot handoff and their tests were fully
reverted with `apply_patch`; the four touched production/test files are
byte-identical to the production base. No API, dependency, behavior, test
count, workflow matrix, ADR or iWork/IWA file changed. The unrelated existing
`docs/FORMAT_IMPLEMENTATION_REVIEW.md` worktree edit remains unstaged by this
batch.

Raw ABBA, repeat, Heaptrack and baseline profile evidence plus pooled JSON
summaries are under `docs/performance/results/`; digests are in
`xls-terminal-render-sha256.txt`.

The non-iWork queue is now:

1. RTF: add byte-1252, LZFu, LibreOffice watermark and relative-font-size
   coverage before another parser specialization.
2. OOXML: profile XLSX writer-local action regrouping on medium and dense-wide
   1% commits before flattening it.
3. ODF: add media-rich ODS corpora and attribute unchanged auxiliary-member
   inflate/deflate before considering validated raw-member transport.
4. OLE2: do not revive the XLS terminal-render handoff without explaining the
   exact no-op regression and providing a materially different attribution.
