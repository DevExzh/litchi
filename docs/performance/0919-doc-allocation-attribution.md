# 0919 — allocation attribution for repeated DOC edits

Status: measured diagnosis; no production change or adoption. This evidence
batch was finalized after the earlier read-only Git limitation cleared.
An [independent allocation audit](results/change-0919/profile-audit.md)
confirms scoped totals, phase accounting and source attribution. Its exact
replay commands and hash-bound external receipts are retained. The output-reservation prototype remains separate and unadopted.

## Boundary and conservation

The retained 0917 native probe is built from `bb19ca74bd52`. The pre-finalization HEAD,
`bf88f1da5f15ebf89b32709c3537b19519d89f2d`, differs only in performance
evidence for the relevant production/build paths. Its executable remains
55,146,832 bytes, SHA-256
`2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13`.
No new harness or library binary was built.

The [prospective protocol](results/change-0919/profile-protocol.json) specifies
six Heaptrack 1.5.0 captures: NoHeadFoot at 32 edits and FloatingPictures at
two/32 edits, each repeated, with the second ordering reversed. Every process
uses CPU 14, default retention, zero warmups and one sample. Heaptrack records
the whole process; analysis accepts only backtraces containing
`doc_retention_probe_0917::measure_one`. Preparation, reference routes and
post-sample oracles therefore remain outside the accepted allocation set.

All six children exit zero and pass the unchanged 0917 report validator's
twenty correctness gates. Two independent representations reconcile:
`heaptrack_print --filter-bt-function measure_one` folded-stack call counts,
and a decode of allocation events in the interpreted Heaptrack recording.
Both exactly match the separately instrumented 0917 allocator's call counts;
decoded requested bytes also exactly match that allocator in every run.

| Fixture | Edits | Lifecycle allocation calls | Requested bytes | Repetition |
| --- | ---: | ---: | ---: | --- |
| NoHeadFoot | 32 | 19,024 | 16,781,690 | exact |
| FloatingPictures | 2 | 18,014 | 17,369,691 | exact |
| FloatingPictures | 32 | 95,057 | 164,233,879 | exact |

The repeated filtered stack files are byte-identical for each fixture/count.
The full process has far more allocations because it also runs reference
routes and correctness checks; those totals are not lifecycle results.
Profiler elapsed time and whole-process/merged-backtrace peak numbers are
not native latency or lifecycle peak-live evidence. The separate 0917
allocator remains the source of peak-live measurements.

## Where requested bytes go

For FloatingPictures at 32 edits, the following categories partition all
164,233,879 requested bytes. They classify allocation backtraces, not timed
phases. Nested final-commit reopens belong to final commit, not initial open.
The exact ordered predicates and joint memberships are in
[phase-results.json](results/change-0919/phase-results.json).

| Exclusive allocation category | Calls | Requested bytes |
| --- | ---: | ---: |
| Initial snapshot open | 5,307 | 4,005,029 |
| Edit construction | 2,355 | 2,268,807 |
| Staged target resolution and guards | 21,009 | 13,173,117 |
| Staged mutation outside publication | 33,102 | 42,275,408 |
| Staged stream handoff outside common publication | 3,328 | 15,523,744 |
| Staged common publication, other | 2,816 | 802,310 |
| Staged rendering | 17,779 | 44,759,310 |
| Staged package recapture | 3,872 | 32,756,816 |
| Final commit and its nested validation/open | 5,487 | 7,647,386 |
| Probe input/output boundary | 2 | 1,021,952 |

The [top byte traces](results/change-0919/top-byte-traces.json) sharpen the
next investigation. Two `ReusePlan::emit` allocation stacks account for
21,245,952 bytes/73 calls and 18,808,832 bytes/32 calls. They are descendants
of `render_with_layout`, which still constructs `Cursor::new(Vec::new())` at
`litchi-ole-common/src/object/codec.rs`. The emitter writes the validated
layout to that growing destination. These 40,054,784 bytes include necessary
output storage as well as growth; they are not all an avoidable saving.

Package capture is the next large copy owner: two individual staged traces
request 15,886,728 bytes for stream ownership conversion and 15,837,080 bytes
through full FAT-stream reads. These are distinct allocations, not an
instruction-count estimate. Capture still needs complete validation and fresh
directory metadata.

Other overlapping subtrees are useful for ranking, but must not be summed:

| Subtree, full lifecycle | Calls | Requested bytes |
| --- | ---: | ---: |
| `parse_sprms` | 21,213 | 8,263,944 |
| `rewrite_chpx` | 10,046 | 8,133,040 |
| `source_paragraphs`, including decoding | 4,189 | 2,370,318 |
| `main_story_text` within that resolver | 285 | 1,574,694 |

Thus avoiding retained strings for unselected paragraphs is a smaller byte
opportunity than output growth and package recapture. High allocation-call
counts in SPRM parsing do not make it the largest requested-byte owner. Its
minimum eight-element reservation is visible as 19,781 allocations of 384
bytes, but this profile does not distinguish empty inputs from short nonempty
inputs and does not justify an empty-input prevalence claim.

## Next implementation decision

First investigate exact output reservation at the existing CFB plan boundary.
`ReusePlan::output_len` already checks the complete planned output length;
the candidate must reuse that plan rather than estimate from source size or
plan twice. It must preserve pre-emission validation, source-layout fallback,
all size/overflow checks, typed allocation failure, deterministic bytes and
ordinary sink behavior. Any retained-capacity changes need an explicit new
prospective protocol; 0918's exact-capacity equality gate cannot silently be
reused or waived for a capacity-changing candidate.

The second option is source-bound in-place comparison during recapture.
[0749](0749-cfb-reuse-plan-validation.md) already introduced the private CFB
`StreamComparer`, with full chain, tail, readability and zero-fill semantics.
Reusing its mechanism requires a reviewed owner boundary that cannot compare
against bytes unrelated to the parsed reader. It must not skip required
readback, directory capture, protection checks or object rediscovery.
[0036](changes/0036-ole-common-stage-attribution.md) already rejected merely
avoiding the Vec-to-Arc copy after materializing a complete stream; that is
not a new candidate.

The rejected [0918](0918-doc-stream-arc-handoff.md) stream handoff remains
unadopted. This attribution neither retries that experiment nor establishes
a speedup for either new direction. Broader CRUD, producer, distinct-object,
cold/range-source and concurrency requirements remain open; iWork is excluded.

## Custody and reproduction

[profile-results.json](results/change-0919/profile-results.json) records every
capture/analysis argv, timestamps, exit, validation summary, and artifact
filename/size/SHA-256. Raw `.zst` captures and 3.5–5.3 MB folded-stack files
stay outside the repository in `/tmp/litchi-performance-0919/`. Nothing under
the evidence packet is 1 MiB or larger, and no executable is included.

For a fresh empty external destination, run each protocol row with its
recorded capture argv, replacing only the output directory. Then run its
recorded `heaptrack_print` argv against that new recording. The stack filter
is required; unfiltered process totals include the untimed work. The
[analysis method](results/change-0919/profile-method.md) describes and gives
the independent event decoder used to reconcile calls and requested bytes.

The [method validation receipt](results/change-0919/method-validation.json)
binds the exact documented decoder and records successful re-execution on
all six recordings. Decoded totals and every reported inclusive membership
match the profile results; exclusive phase totals also conserve calls and
bytes. Root additionally verified 78 artifact descriptors across 60 distinct paths (37,821,896
unique bytes), the method-document hash and local document links. These root checks are
separate from the independent allocation audit linked above.

No temporary build, executable or source checkout was created for 0919. The
external capture, analysis and replay files are retained measurement evidence;
there are no disposable build files to clean from this batch. The completed
original seven-file evidence packet totals 127,651 bytes; every file is below 1 MiB.
Earlier command-startup and read-only Git limitations prevented committing it
on 2026-10-05. Current Git status succeeds; those limitations are historical.
Unrelated worktree changes are excluded from this batch.

The separate 0920 output-reservation prototype and draft measurement design
remain uncommitted work in progress. No comparative performance capture has
run and no production implementation is admitted. This 0919 evidence commit
does not depend on the prototype or its draft tooling.

Root subsequently replayed the audit's strict 0830 decoder over all six raw
captures. The concatenated result exactly matches the external 300-byte audit
receipt (SHA-256 `4f6e9e8405645491a245c6bf8cc4699deac9e705ee5e6612fe2ad9b1fe7b229e`).
The decoder and histogram receipt hashes also match, and the retained filtered
and unfiltered histograms are byte-identical. Their sums are whole-process
900,674 calls / 1,281,109,930 requested bytes, not the scoped lifecycle totals.
The transient sandbox startup quota error cleared on the next continuation.
A fresh 2026-10-06 audit rechecks capture custody and the strict replay before
committing this evidence; it makes no new timing or production claim.

The [finalization receipt](results/change-0919/finalization-audit.json) records
78 successful artifact-descriptor checks and a fresh exact replay of all six
captures. Production crates and Cargo manifests are unchanged between the
profiled baseline and the pre-commit HEAD. The hashed method document retains
its original pending-review wording; the later independent audit above
supersedes that review status without rewriting the historical receipt.
