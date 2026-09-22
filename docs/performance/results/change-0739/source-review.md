# 0739 PPTX cross-slide-copy lifecycle source review

status: bounded read-only source review; approved for descriptive harness
qualification

production_change: none

This review traces the current owned PPTX cross-slide-copy lifecycle from the
benchmark entry point through `CrossSlideCopyPlan::apply`, candidate archive
construction, the allocator region, and the output checks. It records what the
fresh 0739 baseline can attribute and what it cannot. It does not select a
candidate, claim a speedup, or treat the synthetic corpus as native Office
validation. No Rust source, Cargo file, test, fixture, or build configuration
was changed by this review.

The review is anchored to the source revision recorded in
[`source.json`](source.json), `65f82b0f8b7f301df71e42a1899715c90c1573b4`, and
the frozen workload description in [`plan.json`](plan.json). The source files
read for this review are pinned here:

| source | SHA-256 |
| --- | --- |
| [`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs) | `972376832a5ef2a96b735b335820c945429b5ffcca326dbeca59d4b75ee04567` |
| [`tools/perf-baseline/src/allocation_metrics.rs`](../../../../tools/perf-baseline/src/allocation_metrics.rs) | `0fcec3e5c4972031b23542fd359a38401e25f32a133c77ef4c6e75facf0710d5` |
| [`crates/litchi-pptx/src/opened/cross_copy_plan.rs`](../../../../crates/litchi-pptx/src/opened/cross_copy_plan.rs) | `9d7bfa07e9cbed9b529a2e56525a827469fbf3ccf4df818becff987691ba350c` |
| [`crates/litchi-pptx/src/package/model.rs`](../../../../crates/litchi-pptx/src/package/model.rs) | `919729e5689f41f14832a177f59aa1639e2020aaf7d14bcf647b19e5bfd2629f` |

## Corpus and one-time gates

`build_pptx_cross_copy_corpus` in `tools/perf-baseline/src/lib.rs:16403-16573`
creates both inputs in memory. It authors three source slides and two
destination slides with deterministic titles and text. The media-rich arm adds
eight deterministic 2 MiB PNG-shaped payloads to the selected source and
destination slides; the plain arm has no media closure. The archives are then
opened through `Package::from_vec`, so the measured lifecycle receives
source-preserving owned package bytes. Corpus creation, media authoring, the
first plan, and expected-output generation all occur before the per-sample
loop.

The corpus builder plans once and records the selected source/destination
names, insertion position, closure part count, planned bytes, relationship
count, and collision-remap count. It also applies that plan once to produce
`expected_output`. The output is checked by
`verify_pptx_cross_copy_output` (`lib.rs:18744-18930`) and the refusal controls
by `verify_pptx_cross_copy_refusal_gates` (`lib.rs:18932-19072`). Those controls
cover semantic output, package topology, copied dependency bytes and
relationships, source immutability, durable patch round-trip, borrowed-ingress
refusal, stale source/destination refusal, and foreign-source refusal.

The one-time gates are strong preconditions for running this fixed synthetic
workload, but they are not independent native-Office validation. The expected
archive is produced by the same Litchi implementation being measured, and the
oracle reopens it through the same Litchi PPTX/OPC stack. The checks compare
logical member payloads, content types, relationship fields, part sets,
selected slide name/text, and insertion count. The owned lifecycle oracle does
not independently render a presentation, invoke Office, or use a second PPTX
implementation. It also does not exercise the breadth of producer variation
represented by charts, notes, animations, hyperlinks, macros, signatures, or
arbitrary third-party ZIP metadata. Existing report flags therefore remain
benchmark diagnostics; they must not be described as native Office evidence.

The owned `verify_pptx_cross_copy_output` path uses
`pptx_cross_copy_archive_members` (`lib.rs:18451-18466`), which compares
decompressed member bytes. Its physical ZIP order, local/central record bytes,
extra fields, and archive comment are not part of this owned lifecycle oracle.
The source-backed helper has additional raw-member and physical-order checks
(`lib.rs:17520-17573` and `18066-18119`), but those checks are not called by
the owned lifecycle helper. This baseline therefore proves logical package
correctness and its selected preservation checks, not raw ZIP preservation for
every member.

The refusal gates run while the corpus is built, once per fresh process. Each
measured iteration repeats the fixed plan metadata checks, output-byte equality
with the prebuilt expected archive, sink shape checks, a reopen slide-count
check, the semantic/topology/dependency oracle, source byte immutability, and
the output digest. It does not rerun every stale/foreign/borrowed negative
control per sample. That separation is intentional and should remain visible
in the report.

## Lifecycle timer boundary

The boundary is `run_pptx_cross_copy_lifecycle` in
`tools/perf-baseline/src/lib.rs:50430-50699`. For every iteration:

```text
corpus source/destination Vec clones                 outside
bounded CountingSink reservation                     outside
process counter before                               outside the elapsed clock
allocator region begin                               immediately before clock
elapsed clock starts
  Package::from_vec(source_input)
  Package::from_vec(destination_input)
  source.opened_presentation()
  destination.opened_presentation()
  destination_snapshot.plan_cross_slide_copy(...)
  destination.apply_cross_slide_copy_plan(&source, &plan)
  destination.opc()?.to_stream(&mut sink)
elapsed clock stops
allocator region finish and process counter after
reopen, semantic/logical-member checks, digest, source immutability, and teardown       outside
```

The outer elapsed value is therefore an owned ingress-to-publication
lifecycle. The `plan_ns`, `commit_ns`, and `publication_ns` values are nested
diagnostics. `publication_ns` includes obtaining the OPC view and the full
sequential write to the bounded sink; the sink's complete output capacity was
reserved before the clock. `reopen_ns` is measured after the lifecycle and is
not included in it. The source text used by the semantic oracle is prepared
before the loop.

The public `apply_cross_slide_copy_plan` wrapper in
`crates/litchi-pptx/src/package/model.rs:269-305` is wholly inside
`commit_ns`. Its current-graph, mutation-policy, and physical-provenance
checks, the call to `opened::cross_copy_plan::apply_plan`, digest adoption, and
the final mutable-state update are all included. The source package remains a
local immutable input; the source immutability comparison occurs after the
clock.

The lifecycle clock stops before the post-operation checks and before the
`source`, `destination`, `published`, and sink locals are dropped. Thus the
reported elapsed time excludes post-operation oracle/reopen work and the
allocator observation excludes teardown of those retained locals. The clock is not an
isolated "copy bytes" timer: it includes package parsing, opened snapshot
construction, graph validation, candidate creation, revision checks, patch
validation, and publication.

`lifecycle_ns` is the authoritative total for this selector. The sum of the
nested phase values is not a decomposition of it. The residual includes both
`from_vec` ingress and opened snapshot construction, as well as ordinary call
and timer overhead; no pure open or graph fraction can be inferred from the
current fields.

## Planner, candidate, and apply path

`Snapshot::plan_cross_slide_copy` at
`crates/litchi-pptx/src/opened/cross_copy_plan.rs:565-582` enters
`prepare_cross_slide_copy_for_slides` (`:907-1155`). Within the measured plan
clock it:

1. validates source-preserving physical provenance, signatures/macros/unknown
   physical members, dialect, protected/MCE content, slide topology, limits,
   and the source/destination layout-master-theme inheritance graph;
2. collects and sorts the source slide's owned closure, assigns collision-free
   target names and relationship identities, and runs part preflight checks;
3. computes source and destination physical revisions;
4. calls `build_candidate` with `CandidateArchive::BuildAndRetain`; and
5. captures the durable patch and validates its descriptor.

`build_candidate` (`cross_copy_plan.rs:1157-1360`) clones the destination OPC
graph, creates copied parts with shared source payload handles, rewrites
internal relationship targets, retargets the copied slide to the destination
layout, and updates the destination presentation XML and relationships. It
then serializes the candidate through the bounded package writer and reopens
the serialized bytes. For an unmodified owned destination it can use
`from_vec_reusing_payloads`; if the serialized archive fits the intersected
retained-candidate budget, the plan keeps an exact shared handle to those
bytes. The result is a plan-time serialization/reopen cost that is inside
`plan_ns`, even when application can later reuse the retained archive.

`apply_plan` (`cross_copy_plan.rs:596-688`) starts with complete graph
fingerprints and physical package fingerprints for both source and
destination. It captures fresh snapshots, calls
`prepare_cross_slide_copy_for_slides` again, and compares all complete and
physical revisions plus the durable patch with the original plan. If a
retained archive is available, `build_candidate` takes the
`CandidateArchive::Reuse` branch (`cross_copy_plan.rs:1271-1300`): it allocates
a new Vec for the retained bytes and hashes them, avoiding fresh package
serialization/deflate. If retention was not admitted or was released, it
builds and serializes again. The fresh candidate then passes
`validate_application_candidate`, the published physical revision check, and
one final assignment to the destination. Every one of these operations is in
`commit_ns`.

Consequently, the existing three phase values cannot answer whether a future
change helps closure discovery, candidate graph construction, deflate and
reopen, patch capture, revision hashing, fresh application proof, or final
publication. A lower `plan_ns` would not by itself identify which of those
steps changed, and `commit_ns` may include a retained-candidate reuse decision
that is not represented in the JSON summary.

## Allocator and process boundary

The allocation lane uses the separate allocator-enabled executable. The
`allocation_metrics::begin` region takes absolute process counter snapshots;
it does not reset counters. `finish` publishes checked differences for
allocation/deallocation/reallocation calls and bytes, plus entry/exit live
bytes and an observer-ordered region high-water mark. The wrapper's own module
documentation and implementation (`tools/perf-baseline/src/allocation_metrics.rs:1-25,
273-444`) define the limits:

* the counters cover successful global-system-allocator events observed during
  the region, not allocations attributed to a particular internal function;
* `region_peak_live_bytes` includes the live bytes at entry and is callback
  order evidence, not RSS or a full-process heap peak;
* allocator-internal realloc overlap is excluded from that region peak, and
  instrumentation can perturb allocator scheduling; and
* a disabled or unavailable/overflow region is not a measured zero.

The region begins just before the lifecycle clock and finishes just after it.
Corpus clones, sink capacity reservation, all post-clock oracles, and package
and sink teardown are outside the measured allocation interval. There is no
phase allocation split in this harness. The native executable is
uninstrumented; its elapsed samples must not be compared as if they were
timings from the allocator executable. The three allocation processes with
one sample and zero warmups are descriptive allocation observations, not a
latency qualification or a mechanism proof.

## Attribution seam for a follow-up

The next useful diagnostic seam is inside the existing plan/apply boundary,
while retaining the exact same corpus, invocation, process schedule, and
publication behavior:

```text
plan_ns
  closure/topology/layout proof
  candidate graph and relationship rewrite
  bounded serialization/deflate
  candidate reopen/capture
  Patch::capture

commit_ns
  graph and physical fingerprints
  fresh snapshot capture
  fresh prepare/build or retained-archive reuse
  candidate validation and published-revision proof
  destination assignment/digest adoption
```

An opt-in diagnostic timer and allocator-region split at those boundaries
would establish where work and allocations occur without treating an outer
phase as a causal fraction. The existing `Region::split` API preserves the
combined operation interval while publishing non-overlapping segment samples,
so it is the appropriate mechanism for an allocation attribution probe if
the coordinator chooses to requalify the harness. Publication can be split
further only after the OPC writer's accepted-byte/deflate accounting is tied
to the same boundary; the current sink summary records logical accepted
writes, not compressed-source or physical-storage work.

This is a measurement seam, not an approved production optimization. Any
candidate that changes retained-candidate ownership, serialization, or
revision-proof order must keep the exact semantic/refusal/topology contract
and receive a new source review and qualification.

No source blocker was found for the planned descriptive baseline. The current
source is suitable for capture under the frozen 0739 schedule, with the
limitations above carried into the report and without a native Office or
historical cross-build performance claim.
