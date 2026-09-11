# 0510 future queue: OLE2 and OOXML priority review

`scope: read-only evidence review; future queue only`

`reviewed revision: 869bbc3ed`

`performance_claim: none`

This review reads the retained hotspot and goal audits and the existing CFB,
XLSX, DOCX, OPC, and ordered-Part evidence. It does not change Rust, rebuild,
or start a workload. The queue below is for the work that follows the active
0510 capture. ODF work remains deferred until the full OLE2/OOXML optimization goal is
complete. Closing this three-item investigation queue does not establish that
completion; the full requirements still need an evidence audit.

## Decision

The next work should be a measurement-first sequence with three bounded
investigations, in this order:

1. **CFB/OLE2 operation-local attribution and one residual work-elimination
   gate.** This has the clearest legacy evidence and the largest unresolved
   substrate question.
2. **Dense XLSX commit/save operation-local attribution.** This follows the
   accepted repeated-attribute and empty-web-proof reductions and targets the
   larger remaining parser and snapshot paths.
3. **DOCX provider/publication operation-local attribution.** This connects
   the measured delayed-provider opportunity with the publication phase that
   dominates the current opened-edit lifecycle.

Each investigation should first bind an operation-only timer, allocator
region, source counter scope, and hardware-counter capture. A production
change is justified only by a named repeated operation in that scope and a
matched correctness/resource review. The newly available `perf stat`
permission is useful for this work: the grouped 0510 probe reports cycles, instructions, branches, and branch
misses at full running time, alongside software counters. The broader cache
probe had unreliable running time and its cache values are excluded. A
successful command alone does not establish counter quality or operation-local
scope; each workload capture must validate both.

## 1. CFB/OLE2: profile the residual chain and validation work

The 0412/0413 plain-source XLS profile found `SectorChainScratch::collect_exact`
and its `try_push<u32>` path at 24.97% and 24.87% of period-weighted samples
under observed source-open ancestors. 0413 replaced the proven per-sector
fallible push after exact reservation. The accepted warm XLS result reduced
plain owned-source p50 by 1.65%–3.18% and retained a reproducible roughly 3%
few-large CFB guard cost. The same profile still reported candidate source-open
work in FAT loading (16.90%), stream-allocation validation (14.56%), and
physical-layout validation (8.16%). These are inclusive diagnostic shares, not
operation timings, and the CFB guard selectors did not expose operation-local
allocation data.

The next capture should use the current production path on the fixed opaque-
heavy XLS corpus and keep the following arms distinct:

* plain owned positional source for CFB open, XLS open/list, and selected-cell
  access;
* a real `FileSource` arm for source construction and filesystem-backed
  behavior; and
* the existing instrumented source only as a logical-locality control, with
  its observer overhead outside the production attribution claim.

The operation report should split CFB catalog/FAT and MiniFAT work, Workbook
global parsing, selected-worksheet traversal, and final freshness fences. It
should record source calls and bytes, operation-local allocation/peak values,
and validated hardware counters. The tiny MiniFAT and few-large regular-FAT
guards should remain in the same review so a legacy optimization does not
silently move cost to a different chain shape. The old 0413 whole-process
profiles remain context; they cannot substitute for this slice.

If the profile shows repeated construction or validation for monotonic reads,
the highest-value implementation scope is a bounded same-stream span/cursor
reuse at the existing CFB owner boundary. It may reuse an already validated
chain position or exact bounded scratch for consecutive reads. It must retain
fallible reservations, source-version fences after I/O, cursor state
atomicity, cancellation, and the existing ownership and layout proofs. A
profile that finds no repeated work should close this candidate without a
rewrite. `ReadAt` API changes, a generic CFB cache, and broad concurrent chain
walking have no current causal evidence.

The required semantic gate covers FAT/MiniFAT and directory ownership,
cycle/overlap detection, early and late chain markers, truncated headers and
tails, FILEPASS/encryption refusal, duplicate-last BIFF behavior, worksheet
locality, source-change precedence over an I/O error, and unchanged error
ordering. Every retained payload and scratch reservation must remain bounded
by the existing execution context. The few-large CFB guard must remain visible
in any ABBA review; it cannot be folded into a favorable XLS result.

Relevant legacy opportunities remain available for later profile selection:
immutable positional reads without a shared cursor lock, compact chain indexes,
lazy loading of referenced MiniFAT/ministream ranges, contiguous sector runs
directly into the destination, file-backed stream views, bounded large-stream
spooling, and unchanged stream/sector copy-through. Each needs a measured
caller and a proof that all CFB validation continues to run.

## 2. XLSX: isolate the remaining dense commit/save work

0466's dense-wide profile placed inclusive commit-context weight at 40.01%
eager worksheet parsing, 24.41% snapshot scanning, 13.27% changed-XML
compaction, and 10.33% web metadata reading. 0467 removed five repeated
checked attribute scans and qualified a dense one-percent commit/save p50
reduction of 8.68% and 7.86% in its two matched directions. 0468 then showed
that unconditional `Event::into_owned()` was only 5.03% of the compaction
subtree and 0.348% of whole-process weight; 0469 accepted borrowed compaction
events. 0470 removed a later web-binding traversal after a bounded proof and
reported 19.753% fewer whole-process allocation calls, with no exact peak
claim. The larger eager-parser and snapshot traversals remain the stronger
measured lead.

The next diagnostic should bind the dense one-percent commit/save operation
and expose operation-local intervals for eager parse, snapshot scan, changed
XML compaction, web metadata proof, and package publication. Hardware counters
and allocator observations must be collected inside those intervals or from
equivalent operation-only child processes. Whole-process Heaptrack, RSS, and
the 0466 frame-pointer percentages should remain contextual evidence.

The first code candidate after that profile is whichever larger traversal
contains repeated work that can be removed while retaining one complete
validated pass. A separate low-risk memory probe may release the
pre-compaction vector as soon as its proof is no longer needed; 0470 identifies
that vector as an independent opportunity, and it needs operation-local peak
evidence before a memory claim. The failed unrestricted Store-handoff memory
experiment does not justify increasing the 4,096-cell/1 MiB retention limit.

The semantic differential must retain checked duplicate and malformed
attribute errors, entity normalization, namespace and MCE behavior, style and
metadata bounds, formula/type ordering, original-byte preservation, and full
publication validation. A complete pass cannot be skipped because the current
profile has inclusive overlap. No SIMD or parallel change is indicated by the
retained evidence.

## 3. DOCX: separate provider service from publication CPU

The DOCX provider baseline provides a concrete transport signal. 0494 records
about 377 nonempty reads and 16,799,430 logical bytes for the delayed and
warm-file arms; a 1 ms per-request model adds about 377 ms to the nominal
160.212 ms transfer service. Its warm medians are about 564–570 ms for the
delayed arm and 185 ms for the zero-delay range arm. 0492/0493 show that an
explicit 4 KiB/managed bounded window changes 19 physical fills to three and
reduces delayed medians by roughly 82%–85%, while increasing accepted bytes by
37.29% and operation allocation by 4,384 bytes. Local-source medians move by
about 2%, so the policy remains opt-in.

0496 adds phase clocks to the managed edit/save route. Publication is the
largest named normal phase in all 16 normal children, but its wall-clock phase
is not a CPU attribution. The 74 retained review flags remain unresolved,
including phase, allocation, RSS, and one unstable short-provider lifecycle
flag.

The next capture should use the current managed edit/save API with separate
operation-local hardware and allocator scopes for open, edit staging, commit,
XML identity diagnostics, publication, published-snapshot drop, and commit
drop. Keep owned, warm-file, short, zero-delay range, and delayed range arms
separate. Source-call service and publication serialization should be
reported independently so transport delay cannot be mistaken for document
CPU. A verified-cold arm and an independent producer can follow a selected
candidate; they should not be inferred from these warm synthetic sources.

If the operation profile identifies repeated XML audit or publication work,
the safe implementation scope is a proof-preserving reuse of an already
authenticated pass or bounded replay state. The candidate must preserve source
freshness, authored lexical policy, package/relationship validation, output
ordering, cancellation, partial-sink errors, inverse authorization, and all
input/memory/work limits. If provider service dominates, compare only an
explicit bounded read window or same-source range batching with physical
overfetch and retained-window memory charged to the execution context. A
default read-ahead policy for local sources is not supported by 0492/0493.

## Scheduler and publication guardrails carried forward

0499's operation-local worker reuse is retained as a low-level capability, not
as a reason to change format defaults. On many-small owned Parts, reused batch
p50 remains 186.04 microseconds at width 4 and 228.74 at width 8 versus
95.86 microseconds for the after serial route. Warm-file width-4 is 248.77
microseconds versus 182.04 serial. The delayed-provider width-8 p99 rises
30.79% despite a 3.67% median reduction. Whole-child profile evidence shows
fewer cycles alongside +251.49% context switches and +1,450% migrations, so
the scheduler cost is not operation-local yet.

Before any ordered-Part scheduler change, profile the current
`read_parts_ordered` interval around wave admission, command/reply transfer,
fences, joins, and payload publication. A small local candidate may remove a
proven redundant fence, wave-end signal, or reset, provided operation-local
evidence identifies it and the proof retains source/cancellation fences,
lowest-input-ordinal error selection, no later wave after failure, worker
stack/channel reservations, and monotonic `Work`/`InputBytes` accounting. The
operation-local worker set, explicit admission, and one-wave serial fast path
remain architectural constraints; no hidden pool or admission-policy change is
queued.

The 0500 managed paragraph-batch flags also remain active review inputs:
owned p128 K=1 lifecycle p50 is +7.34% for the batch route, p512 K=8
warm-file RSS is +5.77%, and the associated publication p50 observation is
+15.23% while edit p50 falls 0.35%. These observations do not establish
causality and must be rerun or retained in any DOCX publication comparison.
No favorable provider or publication result may erase those flags.

## Evidence references and stopping rule

The primary retained evidence is [0413 CFB scratch reuse](../../changes/0413-cfb-chain-scratch-reservation.md),
[0466 XLSX dense profiling](../../changes/0466-xlsx-dense-commit-profile.md),
[0468 remaining XLSX profiling](../../changes/0468-xlsx-remaining-commit-profile.md),
[0470 XLSX web-proof reuse](../../changes/0470-xlsx-empty-web-proof.md),
[0493 managed DOCX read-ahead](../../changes/0493-managed-opc-source-read-ahead.md),
[0494 DOCX provider baseline](../../changes/0494-docx-edit-provider-baseline.md),
[0496 DOCX phase attribution](../../changes/0496-docx-edit-phase-attribution.md),
[0498 Part-batch baseline](../../changes/0498-bounded-source-backed-part-batch.md),
[0499 worker reuse](../../changes/0499-operation-local-part-workers.md), and
[0500 paragraph batches](../../changes/0500-managed-paragraph-batches.md).

The current [HOTSPOTS queue](../../HOTSPOTS.md) and [goal audit](../../GOAL_AUDIT.md)
remain authoritative for broader coverage gaps. This file promotes only the
three operation-local investigations above. Stop after the first investigation
if its operation-local profile cannot separate the suspected work from setup,
verification, provider service, or shared-host scheduling; retain the evidence
and select the next queue item rather than guessing at a rewrite.
