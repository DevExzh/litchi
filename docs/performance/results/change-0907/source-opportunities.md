# 0907 — PPTX multi-slide source opportunities

This is a finite read-only source review. It makes no latency, allocation,
RSS, or throughput claim and adopts no candidate. The reviewed source is
`3d3650ac03be2cd8ba923a5a6e940e84b14f5eab` (`fix(perf): expose every
registered selector in CLI help`). No build, workload, profiler, or production
source edit was performed for this review.

The review applies the source-bound snapshot and patch rules in ADR 0003, the
lazy input, bounded cache, sequential output, and measurement rules in ADR
0005, and the preserve-or-refuse validation boundary in ADR 0006. In
particular, an opportunity is useful only if it keeps exact source
authorization, stale-source and cancellation checks, XML and aggregate limits,
unchanged-member raw copying, and the complete selected/unselected readback
oracle.

## What the current timer includes

The runner is `run_pptx_multi_slide_batch_edit_save` in
`tools/perf-baseline/src/lib.rs:51122-51219`. Each measured iteration starts at
`:51172`, calls `publish_pptx_multi_slide_batch_edit` at `:51173-51174`, and
stops at `:51175`. The following table separates the work in that interval.

| Timed work | Location | Classification |
| --- | --- | --- |
| Build one `Vec<String>` and one replacement vector for each selected slide | `lib.rs:51052-51065` (source) and `:51089-51098` (eager) | Harness input construction; useful for an end-to-end caller measurement, but not production edit work |
| Source-backed open and retained catalog, or eager source open/materialization into `OpcPackage` | `lib.rs:51051-51052`; eager `:51084-51088` | Production workflow work; keep in the open-inclusive endpoint |
| Selected-slide capture, transition/root checks, shape-text staging, commit, and eager transaction equivalent | `litchi-pptx/src/presentation/source.rs:1683-1687,1713-1718,2335-2360,2714-2786`; eager calls through `lib.rs:51088-51099` | Production validation and edit work |
| Forward replay and inverse replay of the source patch | `lib.rs:51070-51073,51076-51079` | Harness-only qualification work inside the current timer |
| Eager inverse/forward package replay | `lib.rs:51103-51110` | Harness-only qualification work inside the current timer |
| Source-backed overlay publication, or eager `to_bytes` and chunked sink writes | source `lib.rs:51074-51081`; eager `:51104-51119` | Production publication work, including required preservation planning and output |
| `cache_diagnostics`, changed-count checks, slide-count checks, and the timed helper's error guards | `lib.rs:51066-51075,51100-51114` | Harness observation and guards around production calls |

The expected-output preflight, full package reopen, topology/relationship
comparison, selected-slide semantic check, untouched-Part comparison, media
readback, output digest, and sink-bound checks occur outside the timer in
`lib.rs:51139-51158,51177-51196`. They are necessary correctness gates, not
latency work to remove. The source-backed guard remains the fixed
`PPTX_MULTI_SLIDE_BATCH_COUNT + 1` materializations and the eager guard remains
the complete manifest Part count; neither is a performance claim.

The current helper therefore cannot be read as a pure edit/publication timer.
In particular, an apparent source-path cost includes two patch replays per
iteration, and an apparent eager-path cost includes the corresponding package
replays. A future phase report should retain these replays as untimed
qualification or as a separately labelled diagnostic lane.

## Existing diagnostics do not provide the needed phase attribution

The optional OPC observer is not a publication phase vector. Its
`DiagnosticSnapshot` reports only content-free cache-state and same-Part
single-flight mutex acquisition boundaries (`crates/litchi-opc/src/source_backed.rs:3253-3307`),
and the public observer explicitly receives no package, Part, or source
identity (`:7532-7553`). It can supplement a future profile with cache-lock
contention evidence, but it cannot say whether an event belongs to catalog
open, selected-slide validation, source publication, or a harness replay.
The PPTX editor's public `cache_diagnostics` is likewise a content-free
counter snapshot with no member identities (`crates/litchi-pptx/src/presentation/source.rs:1647-1670`).

There are phase vectors for other PPTX selectors, but they do not answer this
question. `PptxOpenedTransactionPhaseSummary` and its runner
(`tools/perf-baseline/src/lib.rs:5421-5442,34463-34600`) cover one eager
one-shape opened-transaction edit; their `publication_ns` is the eager
`Package::to_bytes` call, and their package construction, semantic reopen,
digest, and checks are outside the clocks. The cross-copy summaries likewise
describe a different copy route. Neither vector is wired to
`publish_pptx_multi_slide_batch_edit`, and the current multi-slide runner
sets `operation_metrics: None` (`:51208-51218`).

Its `SourceSummary` records aggregate read/materialization counters after the
timed helper returns (`:51188-51198`; `:5483-5500,9233-9249`), so those
counters can qualify source activity but cannot attribute elapsed time to
production publication versus the replay calls at `:51070-51079` and
`:51103-51110`. Therefore the existing observer and other PPTX phase vectors
are useful supporting evidence only. The next profiling batch needs explicit,
bounded phase timestamps around this route's production helper stages, plus
separately labelled replay timing or replay gates, while retaining the observer
for lock/read context. It should not infer a phase cost from the number of
semantic materializations.

## Production path and concrete opportunities

### 1. Exact-source patch authorization has repeated comparisons

The source publication path must re-read the selected source Parts and prove
that the patch still names the exact source before writing overlays. That proof
is required. There are, however, avoidable repeated comparisons inside the
proof machinery:

* `SourceBackedSlidePatch::try_apply_fast` first calls
  `source.same_source(&self.before)` and then calls `self.apply(source)`, whose
  first action is the same `same_source` comparison
  (`crates/litchi-pptx/src/presentation/source.rs:2428-2436`). The batch form
  has the same shape at `:2893-2901`.
* `SourceBackedSlideBatchPatch::apply` then calls `is_changed`, which compares
  the complete before and after selected-slide snapshots again
  (`:2878-2882,2903-2917`). `SourceBackedSlidePatch::apply` has the analogous
  single-slide path (`:2413-2417,2438-2449`).
* `SourceBackedSlideSnapshot::same_source` compares the selected XML and all
  source closure fields, including `presentation_xml`
  (`:2246-2257`). A batch invokes that comparison once per selected slide
  (`:2848-2855`), even though the package relationships and presentation root
  are shared by every slide in the batch.

The first item is a direct candidate for an exact-match fast path that keeps
the execution checks and preserves the current no-op behavior of returning the
current source snapshot. A cached `changed` bit, established only by the
existing staging functions, could separately avoid recomputing before/after
byte equality. A batch-level provenance comparison could check shared closure
fields once and retain the per-slide XML and binding checks. These are
implementation hypotheses, not authorization to replace byte equality with a
version or digest alone. Any change must retain a byte-comparison fallback (or
an equally strong exact-handle proof) and must test changed, exact no-op, stale,
foreign-lineage, source-mutated, managed-cache, and cancelled cases.

This is the first production candidate to isolate in a phase profile. Count
the bytes compared and the number of `same_source` calls separately for one,
same-slide batch, and eight-slide batch edits. The harness replays exercise the
same methods but are not evidence of production latency; measure the
publication call with replay removed from its operation clock.

### 2. Overlay limit totals and preservation planning repeat immutable scans

After the selected slide replacements are assembled, the batch publisher calls
`write_part_overlays_shared_to_stream` (`source.rs:1870-1876`). Its
`validate_overlay_limits` implementation walks every catalog Part and every
physical archive member to recompute immutable source totals, then applies the
replacement deltas (`crates/litchi-opc/src/source_backed.rs:10705-10783`). The
same publication then builds a physical preservation index and scans its
entries once to count replacement/deletion matches
(`:11862-11884`) and again to construct the copy/regenerate/omit plan
(`:11927-11965`). The second pass is needed only because the current code
separates occurrence proof from plan construction; unchanged ZIP members still
must be physically copied by the preservation writer.

Two bounded candidates follow from this structure:

1. Retain validated immutable aggregate metadata from the source open, or from
   a publication preparation object, and validate only replacement deltas at
   save time. The source version, limits, overflow checks, and exact physical
   member set would remain authoritative.
2. Build replacement occurrence counts and the preservation actions in one
   indexed pass, then check that every selected member occurred exactly once
   before the first output byte. This must retain duplicate/case-folding,
   missing-member, omission collision, signed-source, trailing-byte, and
   source-freshness refusals.

These should be profiled as metadata-planning time separately from ZIP writing.
They are not permission to skip `preservation_index_with_limits`, aggregate
limit validation, XML replacement auditing, or the raw-copy plan. Exact-limit,
malformed-physical-layout, duplicate-name, signed-source, sink-failure, and
source-mutation tests are required before either candidate is considered.

### 3. Selected-slide validation has a measurable consolidation question

For each selected slide, `slide_snapshot_from_retained_catalog` loads the
deferred Part, checks source freshness, validates the transition source, builds
a `SourcePart`, and validates the slide root through `SlidePart::from_part`
(`crates/litchi-pptx/src/presentation/source.rs:2143-2192`). The subsequent
`stage_shape_texts` pass parses a `Scene`, resolves selectors and raw text
spans, rewrites only the selected text elements, reparses the candidate Scene,
and checks the selected semantic values and shape count
(`crates/litchi-pptx/src/opened/xml.rs:754-850`). Publication separately audits
changed XML through the OPC source replacement validator before output.

Some of these scans are deliberately different safety contracts: transition
validation handles its own supported/unsupported owner rules, the Scene/raw
span path proves shape locality, and the OPC audit owns XML publication safety.
They must not be collapsed by assuming that one parser covers another's
refusals. The concrete next step is phase instrumentation that reports root/MCE
validation, transition validation, Scene/span/rewrite/readback, and publication
XML audit separately. Only a differential proof that preserves error
precedence, markup-compatibility refusals, output limits, and selected-shape
locality should motivate a fused scan or a retained proof memo.

## Focused next profiling plan

Keep the existing guarded end-to-end result for continuity, but add a second
scope with the same frozen corpus and binary:

1. Prebuild the eight per-slide text strings and replacement descriptors before
   the operation clock. Keep the current input construction in a separately
   labelled caller/end-to-end lane so its cost is not silently discarded.
2. Move source and eager forward/inverse patch replays outside the production
   operation clock. Run the same replay, stale-source, inverse, digest, and
   preservation gates and retain their results as qualification evidence.
3. For the source route, record phase vectors for catalog open; selected Part
   loads; transition/root validation; Scene selection and raw-span planning;
   rewrite plus candidate readback; batch commit; exact-source authorization;
   aggregate-limit planning; preservation-index planning; and ZIP output. For
   the eager route, split open/materialization, transaction staging/commit, and
   `to_bytes`/sink output using the same labels where the work is comparable.
4. Keep both an open-inclusive endpoint and a fresh-editor operation endpoint.
   A pre-opened warm editor would answer a different question because
   publication consumes the editor and selected payload caches change the read
   state. Bind every endpoint to the source manifest, corpus archive/output
   identities, binary hash, toolchain, CPU placement, and source-read policy.
5. Requalify one-edit and same-slide batch controls beside the eight-slide
   batch. If exact-source comparison or overlay planning scales with selected
   slides or physical members, the controls will distinguish that scaling from
   the required raw-copy/output cost.

Adoption of either production candidate requires exact output and source
immutability, unchanged Part and relationship preservation, materialization
and source-read guards, all existing limits/refusal tests, and a fresh matched
timing result. This review supplies none of those performance results.

## Provenance and limits

The reviewed source blobs have these SHA-256 values:

```text
tools/perf-baseline/src/lib.rs                         f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea
crates/litchi-pptx/src/presentation/source.rs          56edd5712b7c711e835b9330b9ca2826cc10e2f2ab86266b51dcded1b6ec9350
crates/litchi-pptx/src/opened/xml.rs                   1a2c5e1f3858bae7e5a21f3c25a9bbd4838bb68079a16835e7d45d8522f57849
crates/litchi-pptx/src/opened/transaction.rs           01352c03f17c8cdbed0902ac582a9d94368ddce6ae3b6f79c157c89f6b7cb310
crates/litchi-pptx/src/presentation/transition.rs     6b2b2190381149ccdbae5b03b37a9ecb61948f71254a41af4cb44377af30e88e
crates/litchi-opc/src/source_backed.rs                 0de2778325c86d81a6a4d5c209d94d0c9b3452489c8c777e9528c79ee87ce8ad
```

The older 0829 PPTX phase profile and 0906 opportunity memo remain historical
or planning evidence. They do not supply current timings for this source or
authorize removing any validation. No conclusion here depends on a timing
number, an old executable, or a profiler attribution.
