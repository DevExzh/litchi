# Change 0012: coalesced DOCX paragraph replacements

Date: 2026-08-10

## Decision

Accept an additive canonical batch transaction for direct-body DOCX paragraph
text replacement. `Edit::replace_body_paragraph_texts` validates a non-empty,
strictly increasing set of unique paragraph positions, plans every disjoint
rewrite against the current immutable snapshot, emits the changed XML once,
parses one candidate snapshot, and performs complete semantic readback for
every selected paragraph before publishing the candidate.

The durable patch remains a sequence of ordinary `ReplaceParagraphText`
operations. Existing scalar callers and wire compatibility are unchanged.
Exact no-op batches retain the original shared XML allocation, while any
selector, authored-text, allocation, resource, parse, or readback failure
leaves the edit unchanged. Run formatting, drawings, unknown run XML, source
checks, limits, inverse operations, and package publication contracts remain
owned by the existing transaction machinery.

The performance harness uses the batch API only when the deterministic DOCX
one-percent selection contains more than one paragraph. The one-edit case
therefore remains a scalar-API guardrail. It can still observe the shared
range emitter's move from repeated copies to one forward copy, so it is
measured separately rather than claimed to be implementation-identical.

## Work removed

The former benchmark implementation called `replace_paragraph_text` once per
selected paragraph. On the large corpus that meant 100 complete main-document
XML rebuilds, 100 full `Snapshot::from_xml` parses, and 100 selected-paragraph
readbacks. The accepted path performs one complete rebuild and parse while
retaining per-paragraph validation and final readback. The shared disjoint
range emitter also precomputes its bounded final size, reserves once, and
copies source/replacement spans in one forward pass.

## Matched latency result

The release executables were frozen on production base `2250cb302` with the
same completed harness:

- before SHA-256:
  `b18d38cbe1886229d8db0f55ba86c995d03e38289b0f9eb78b1c50641bf536b5`
- after SHA-256:
  `7a20c6596f84b1faaff8b1128694085ee954f0f27ac38db294eb816d5ffbb681`

Both states ran pinned to CPU 2 in before-A, after-A, after-B, before-B order,
with three warmups and 15 measured samples per leg. The table pools both legs
(30 samples per state). Times are milliseconds; mean intervals are two-sided
Student's-t 95% intervals.

The reports mark the worktree dirty because the implementation, harness,
performance records, and an unrelated pre-existing documentation edit were
uncommitted. The before executable was frozen while the public batch method
still delegated to the scalar method for every selected paragraph; it already
contained the identical harness and API surface.

| DOCX one-percent edit/save | Before p50 / p95 / p99 | After p50 / p95 / p99 | p50 delta | Before mean (95% CI) | After mean (95% CI) | Mean delta |
|---|---:|---:|---:|---:|---:|---:|
| Medium, 200 paragraphs / 2 edits | 0.652 / 0.723 / 0.728 | 0.568 / 0.641 / 0.641 | **-12.98%** | 0.658 (0.648-0.669) | 0.571 (0.559-0.583) | **-13.27%** |
| Large, 10,000 paragraphs / 100 edits | 487.542 / 508.678 / 509.927 | 24.418 / 25.432 / 25.938 | **-94.99% (19.97x)** | 491.486 (486.909-496.063) | 24.492 (24.286-24.697) | **-95.02%** |

The matched large-corpus scalar one-edit guardrail is neutral: p50 moves from
24.095 to 23.788 ms (-1.28%), while mean moves from 24.315 ms (95% CI
24.017-24.614) to 24.508 ms (23.798-25.218), +0.79% with overlapping
intervals. Its after p95/p99 are 26.268/33.708 ms versus 25.413/27.691 ms
before, so no scalar tail-latency improvement is claimed.

Raw samples:
[`before A`](../results/abba-docx-batch-before-a.json),
[`after A`](../results/abba-docx-batch-after-a.json),
[`after B`](../results/abba-docx-batch-after-b.json), and
[`before B`](../results/abba-docx-batch-before-b.json).

Scalar one-edit guardrail samples:
[`before A`](../results/abba-docx-batch-one-before-a.json),
[`after A`](../results/abba-docx-batch-one-after-a.json),
[`after B`](../results/abba-docx-batch-one-after-b.json), and
[`before B`](../results/abba-docx-batch-one-before-b.json).

## Allocation and memory result

Heaptrack over one large sample reports allocation calls falling from
22,098,595 to 1,302,556 (**-94.11%**). Temporary allocations are effectively
flat at 222,895 versus 222,598 (-0.13%), peak heap is flat at 35.55 MB, and
profiler RSS falls from 38.11 to 37.70 MB (-1.08%). The remaining temporary
count and peak are dominated by corpus construction, package open/save, and
complete output verification.

A reverse-order uninstrumented GNU Time sample reports 33,796 KiB before and
33,920 KiB after maximum RSS (+0.37%). This is treated as flat measurement
noise, not a memory improvement. The decision rests on eliminated full parses,
the allocation-call reduction, and the matched latency distribution.

## Correctness and contract gates

- The batch result is byte-identical to the same strictly ordered scalar edit
  sequence. Its deterministic durable JSON patch applies to the source and its
  inverse restores the exact source bytes.
- Exact all-no-op input records zero operations and retains the original XML
  pointer. Empty, duplicate, out-of-order, and late invalid-text inputs refuse
  without modifying the edit.
- The medium harness case necessarily selects two paragraphs, publishes the
  edit to a complete DOCX package, serializes it, reopens it, and verifies every
  paragraph, full text, operation count, and sink behavior.
- The all-feature DOCX gate passes 841 unit tests, every integration suite, and
  74 doctests (31 ignored). The harness passes 21 tests. Warning-denied Clippy
  passes for the DOCX library and all harness targets; the broader DOCX
  all-target gate remains blocked by pre-existing test-only lints outside this
  change.
- No public archive type, dependency edge, unsafe code, ambient I/O, hidden
  runtime, global lock, or iWork/IWA change is introduced.

## Remaining limitations and next audits

The batch currently targets direct-body paragraphs with strictly increasing
positions; it does not add unordered, nested-story, structural, bulk-pattern,
or merge/split semantics. Complete XML validation and final readback remain
intentional fixed costs. Generated text corpora do not replace real-producer,
media-heavy, unknown-extension, signed, encrypted, malformed, cold-source, or
durable-patch timing matrices.

The next source-audited non-iWork candidates are deliberately separate:

1. Add public RTF open/list/full-text/stream-save/no-op/one-edit baselines, then
   measure replacing raw `RtfDocument::text`'s temporary `Vec<&str>` plus
   `join` with one pre-sized `String`.
2. Measure ODT snapshot handoff from the transaction package's existing shared
   immutable bytes instead of cloning the complete ZIP.
3. Add semantic DOC/XLS/PPT open/edit/save baselines before experimenting with
   any further CFB stream ownership change; the prior DOC move variant remains
   rejected because it regressed.

iWork remains deferred while the `iwa-*` crates are modified independently.
