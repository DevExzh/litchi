# Change 0063: atomic source-backed PPTX shape-text batch

Date: 2026-08-12

Production base: `2e7dc466f`

Status: accepted

## Contract and implementation

PPTX now exposes a borrowed `ShapeTextReplacement` value and an atomic
`set_shape_texts` operation on both the ordinary opened transaction and the
guarded source-backed selected-slide edit. A batch accepts at most 256
selector/value pairs and at most the existing aggregate text-byte limit.
Names and checked pre-order indices resolve against one immutable scene;
duplicate resolved identities and overlapping raw spans refuse before state
changes. Caller order is canonicalized by raw position.

The shared rewriter scans the slide XML once, validates all selected `a:t`
owners, computes the exact escaped output length with checked arithmetic,
reserves one bounded candidate, and emits all replacements in one pass. One
final scene parse proves unchanged shape count, no MCE projection, and exact
readback for every requested shape before the candidate is installed.

The source-backed editor still consumes one operation and publishes only one
existing slide Part through the accepted OPC overlay. Empty input does not
consume the capability. A nonempty all-equal batch retains the original XML
allocation, commits as an exact no-op, and remains byte-exact on signed input.
Changed signed sources, MCE-selected slides, stale/foreign closure, custom
output limits, and partial sequential sinks retain their existing refusals.
Package topology, relationships, content types, every unselected slide/media
Part, and the exact package/presentation/slide closure remain outside the edit.

## Matched corpus and protocol

The fixed corpus has 200 slides, eight text boxes per slide, eight deterministic
incompressible 2 MiB PNGs, 229 ordinary Parts, 445 ZIP members, and 17,568,429
logical Part bytes. Its SHA-256 is
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
Both cases replace all eight text boxes on slide 100 and emit the same
17,017,138-byte artifact with SHA-256
`8371618225b8478d7ea606f13d21d3453e7960f5cb09935975ac672320001755`.

The eager control materializes all 229 Parts, uses the same public atomic batch,
commits through the ordinary opened-package owner, and writes sequentially.
The source-backed case materializes only the mandatory presentation root and
selected slide, then raw-copies every other member. Corpus construction,
complete semantic/topology/relationship/media verification, output hashing,
and patch forward/inverse checks remain outside timing.

The prechange release binary was frozen as
`718297e1e86e716284d82f233ab09fd143d07e4bc6ad04c477f6d62b4f07344a`;
it cannot run the new batch cases, so the primary comparison is the matched
eager/source paths in measurement binary
`4ece3601613889ff99bcafdd5d1e227c702ca711b1e3311fa384bcdbb84bd080`.
After the final fallible-reservation and aggregate-limit tightening, release
binary
`3f4aea56e8c1921072afed0e2c2d21f87472eb4480babfee50e3f18354f4ef0e`
retained the exact output and 229-versus-two counts; a 25-sample confirmation
measured 320.495/8.422 ms p50 (-97.37%) and 321.104/8.426 ms mean (-97.38%).
Runs used Rust 1.95.0, Linux 6.8.0-101-generic, AMD EPYC 9575F, the system
allocator, and CPU 2 affinity.

The host was noisy: complete A/B/B/A attempts contained two-state latency
plateaus and eager long-tail contention despite 10 and 30 warmups. Those raw
legs are retained and explicitly excluded in the
[`summary`](../results/pptx-source-batch/summary.json). The accepted pool uses
two later legs per state whose p50/mean drift stayed within 2.58%, for 200
samples per state. This selection and every discarded filename are recorded;
no unstable leg is silently pooled.

## Results

| Eight-shape same-slide edit/save | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 200 | 200 | - |
| p50 | 322.306 ms | 8.206 ms | **-97.45%** |
| mean | 323.876 ms | 8.274 ms | **-97.45%** |
| p95 | 329.235 ms | 8.997 ms | **-97.27%** |
| p99 | 334.162 ms | 9.253 ms | **-97.23%** |
| semantic Part materializations | 229 | 2 | **-99.13%** |

Heaptrack whole-process profiles reported 3,220,146 versus 1,938,449
allocation calls (-39.80%), 622,673 versus 196,334 temporary allocations
(-68.47%), and 175.13 versus 159.49 MiB peak heap (-8.93%). Heaptrack RSS was
154.35/154.61 MiB. Uninstrumented GNU Time maximum RSS was 145,340/147,056 KiB
(+1.18%), within the 5% guard.

Three-repeat process-wide `perf stat` means were 6.359/3.203 billion cycles
(-49.63%), 16.607/8.232 billion instructions (-50.43%), and 40.670/14.906
million branch misses (-63.35%). These profiles include deterministic corpus
construction and complete untimed verification, so their deltas understate
the scoped publication improvement.

## Correctness and validation

Focused tests prove two-shape mixed name/index batches, caller-order-independent
bytes, duplicate identity refusal, invalid-late-item failure atomicity and
retry, all-equal signed identity, the one-operation guard, exact unselected
Part/relationship/media preservation, two semantic materializations, and full
ordinary reopen. The eager batch composes through the durable opened-package
transaction and final semantic readback. CI pins the corpus/output hashes,
Part/member/logical byte counts, sink bound, source reads, and 229-versus-two
materializations for smoke and release runs.

This tranche changes only OOXML/PPTX and the standalone benchmark. OLE2, RTF,
ODF, and all iWork/IWA crates are unchanged. The previously committed
ODF-common deprecation cleanup remains enforced by warning-denied Clippy and
rustdoc gates.
