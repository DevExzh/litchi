# Change 0045: coalesced ODT paragraph publication

Date: 2026-08-11

Production base: `a9bc0eb0ded7e255baefbe47016c007c333ceada`

Scope: consecutive plain-text paragraph replacements in one packaged ODT
transaction. OLE2, OOXML, RTF, iWork and IWA production code are unchanged.

## Problem and change

`Edit::commit` previously handled every `paragraph.replace` operation as a
complete package transaction. Each replacement rebuilt a `MutableDocument`
from the preceding candidate, regenerated and published `content.xml`, reopened
the ODT, and audited the changed XML before the next replacement started. A
large canonical 1% update therefore repeated whole-document work 100 times even
though plain-text replacement cannot alter paragraph topology.

Commit now coalesces only runs of at least two consecutive
`ReplaceParagraph` operations. It constructs one mutable candidate, applies
the replacements in their original order, publishes `content.xml` once,
reopens the candidate once, and performs one complete changed-XML compact
audit. A scalar replacement takes the old match arm unchanged. Any intervening
operation ends the run and retains its established publication boundary.

This is a private dispatch optimization, not a new API or patch vocabulary.
Every requested replacement still has one ordinary `paragraph.replace`
operation and one `OperationResult::Unit`; duplicate positions retain
last-write ordering. Invalid late positions fail the whole immutable edit
without publishing an intermediate package. No cache, runtime, lock, unsafe
code, dependency, limit, selector, error type or normalization path is added.

## Deterministic benchmark

The new opt-in `odt_semantic_one_percent_edit_save` case uses the existing
deterministic ODT corpus and stages evenly spaced scalar paragraph replacements
through the public transaction. Tiny, medium and large have 24, 200 and 10,000
paragraphs, selecting 1, 2 and 100 replacements respectively. Timing starts
after document open and covers edit creation, staging, commit and output-byte
materialization. Untimed verification reopens the candidate, checks every
paragraph and complete text, and checks the operation result count.

The fixed large source is 28,420 bytes, represents 490,000 authored text bytes,
and has SHA-256
`9d724c649cb5e4b4adce30c4ede2059ff9efc26109c1b84ac8460df00ecf89a9`.
The medium source is 2,648 bytes with SHA-256
`1d175098a7ffa42066bafb19e9a3e5ea44b73c646561227c0bde5bca30ac37ce`.

Control and candidate were built at the same detached worktree path and target
from the same harness source. Their executable SHA-256 values are
`65c5361fb7e82bbe42462518eeda28853ae90df4c261c1faa65135d3283b1497`
and
`4bcff9ea1e6e210ac76f3e1de03b02ec6874cd55e52cbcf44a404b03c99a8e4c`.
Matched A/B/B/A runs used three warmups and 15 samples per leg, yielding 30
pooled samples per state and shape.

| ODT 1%-edit/save | Before | After | Delta |
|---|---:|---:|---:|
| Medium p50 | 1.011 ms | 0.731 ms | **-27.62%** |
| Medium mean | 1.033 ms | 0.732 ms | **-29.11%** |
| Medium p95 | 1.342 ms | 0.881 ms | **-34.30%** |
| Large p50 | 906.439 ms | 15.615 ms | **-98.28% (58.05x)** |
| Large mean | 908.706 ms | 15.704 ms | **-98.27% (57.86x)** |
| Large p95 | 935.323 ms | 16.552 ms | **-98.23%** |
| Large p99 | 942.334 ms | 17.274 ms | **-98.17%** |

Both independent large legs reproduce the effect: before/after p50 is
905.801/15.630 ms in A and 907.078/15.601 ms in B. Medium likewise improves in
both legs. Raw records are [`before A`](../results/abba-odt-paragraph-batch-before-a.json),
[`after A`](../results/abba-odt-paragraph-batch-after-a.json),
[`after B`](../results/abba-odt-paragraph-batch-after-b.json), and
[`before B`](../results/abba-odt-paragraph-batch-before-b.json).

## CPU, allocation and guardrails

Three-repeat `perf stat` processes use one warmup and five samples. Task clock,
cycles, instructions and branches fall 95.87%, 95.93%, 96.90% and 97.10%.
Matched `perf record` reports show the repeated parser, element-clone and XML
publication frames leaving the after profile; the retained final compact audit
remains visible.

A single-sample Heaptrack process moves 16,649,546 -> 643,841 allocation calls
(-96.13%) and 3,133,789 -> 140,632 temporary allocations (-95.51%). Peak heap
is flat at 18.23 MiB. Tool-inclusive RSS rises 28.19 -> 30.99 MiB (+9.93%),
while the uninstrumented maximum is exactly 30,848 KiB in both processes; no
RSS improvement is claimed. See the complete
[`profile summary`](../results/odt-paragraph-batch-profile.txt), raw perf
reports, CSV counters and GNU Time records in `results/`.

The scalar/no-op guard A/B/B/A uses five warmups and 50 samples per leg. Pooled
medium one-edit p50/mean moves -0.29%/-0.45%, while large moves
-0.17%/+0.21%; medium p95 is -0.17% and large p95 is +3.07%. No-op medians
differ by 30 ns medium and -146 ns large. These sub-microsecond values are
disclosed as timer-level noise, not a speedup or regression claim.
Raw guard reports use the `abba-odt-paragraph-batch-guards` prefix.

## Preservation and correctness gates

- A batch is byte-identical to the same ordered replacements published through
  separate scalar commits. Duplicate positions, untouched neighbors, operation
  result count, durable deterministic JSON, replay, exact inverse and stale
  source authorization are checked.
- A late invalid position refuses atomically; a non-replacement operation ends
  a run; the scalar path remains separate.
- Two replacements preserve exact raw local and central records for mimetype,
  styles, metadata, manifest and a 1 MiB opaque media member.
- A two-replacement document whose regenerated `content.xml` exceeds the common
  16 MiB splice limit exercises and verifies the existing full rebuild path.
- Complete all-feature/all-target ODT tests, warning-denied Clippy and rustdoc,
  doctests, complete harness tests/Clippy, JSON/YAML parsing, formatting and
  diff checks are required before commit.
- CI includes the new case in the exact tiny smoke and tiny/large scheduled
  matrices, taking the harness to 121 selectable cases without changing its
  36-case / 198-record default.

All retained evidence hashes are indexed by
[`odt-paragraph-batch-sha256.txt`](../results/odt-paragraph-batch-sha256.txt).

## Remaining work

- The current implementation also coalesces contiguous model-backed paragraph
  insert, replace, remove, inline-run and hyperlink operations. Inline appends
  form a one-way boundary before a later plain/topology operation so scalar
  reopen semantics are preserved. This follow-on correctness extension is not
  represented by the retained 0045 latency or allocation measurements.
- XML-only line breaks, moves, notes, fields, RDF, forms, charts, resources and
  other package domains keep their existing publication boundaries.
- The final large candidate still performs one full mutable-document parse,
  content generation, package publication, package reopen, compact audit,
  final snapshot validation and complete semantic benchmark readback.
- Broader ODF 1%/bulk CRUD, source-backed selectors, real-producer and unknown
  extension corpora, security envelopes, resource addition and structural
  publication remain separate work.
- iWork/IWA remains deliberately out of scope while other agents modify those
  crates.
