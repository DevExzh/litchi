# Change 0375: PPTX selected-slide retained snapshot

Status: implemented

performance_claim: none

## Scope and mechanism

Change 0375 retains the validated selected-slide semantic snapshots produced
during source-backed PPTX planning through publication. Publication reads the
current raw selected bytes and proves the same source with execution, version,
lineage, URI, limits, the complete retained selected-slide closure, and exact
selected bytes. It then applies the change against the retained snapshot
rather than reparsing the semantic slide.

An identity mismatch uses the exact old semantic recapture path. Later raw
errors never retry or publish, and the earlier retained bytes are validated in
their existing order, including the post-parse version fence. The
optimization is owned by litchi-pptx and uses the existing bounded OPC
publisher. No dependency edge, public API, archive/runtime handle, lock,
unsafe code, or parallel execution path is added.

The existing selectors are pptx_source_backed_one_edit_save,
pptx_source_backed_batch_edit_save, and
pptx_source_backed_multi_slide_batch_edit_save. No new selector or default
matrix claim is made here.

## Focused snapshot proof

changed_single_publication_reuses_semantic_snapshot passed with SlidePart and
Scene counters at 1/1 before publication and unchanged at 1/1 afterward.
changed_multi_slide_batch_publication_reuses_semantic_snapshots passed with
both counters at 2/2 before and after publication.
foreign_identical_source_recaptures_semantically_before_rejecting performed
one semantic recapture, returned StaleSource, and emitted zero bytes. These
focused tests and the independent source/safety reviews establish the reuse
and refusal behavior; timing is not used to infer reuse.

## Validation

Validation used one Cargo process at a time, CARGO_BUILD_JOBS=1,
CARGO_INCREMENTAL=0, an 8 GiB build virtual-memory cap, one process at a time
for runs with a 4 GiB virtual-memory cap, and no parallel build or worktree
lane. Process-global thread count was unavailable. A cold-requested run is a
request, not proof of a cold OS or filesystem cache.

The litchi-pptx library suite passed 533/533 with the exact unrelated
pre-existing exclusion
opened::tests::stale_and_unsupported_raw_xml_fail_before_publication.
Integration targets passed with the exact unrelated exclusions
pptx_malformed_presentation::malformed_presentation_children_are_reported_by_their_owner
and
pptx_table_styles::noncanonical_style_target_survives_transactional_raw_save;
source_backed_edit passed 21/21. All three exclusions are stale expectations
around stricter direct sldIdLst owner validation and never enter the Change
0375 publication path.

Production-library Clippy passed with only the explicit pre-existing
allowances clippy::nonminimal_bool in opened/xml.rs,
clippy::clone_on_copy in presentation/order.rs, and
clippy::needless_lifetimes in presentation/order.rs. The crate-boundary gate
passed for 64 packages, 240 declarations, and 14 explicit debt items.
Independent equivalence, safety, and test reviews were accepted.

## ABBA evidence

The clean A1/B1/B2/A2 release protocol used Rust 1.95.0 on Linux AMD EPYC
9575F, CPU affinity 2, one available logical CPU, one execution worker, 20
warmups, 500 samples per case, fresh isolated child processes, and no
instrumentation. Control was revision
53cbdbd085bd0f6905a3a8987acb7deb284ff717, with externally recorded git tree
0829a1006c1e351c6408f5c758e67de942807110 and binary SHA-256
d4679bd0b3791a8d1146da165858fbfe4ada32a0d77c1f39e73e57e69733c4f0.
Candidate was source commit
b4f83ae0c65738f99c605395bdb5020a0660257d, with externally recorded git tree
3bec111713710835bdc3fad8c39e6f2490d44e10 and binary SHA-256
91d334142f87596d5c267afa616ac66391e125ffe892d161ac303e34572ce3b5.
All binaries and raw report legs were clean.

The generated corpus uses litchi-pptx-source-edit-media-v1: 200 slides, 1,600
text boxes, eight inert 2 MiB PNG payloads, 229 OPC entries, 445 ZIP members,
17,017,139 archive bytes, and 17,568,429 uncompressed bytes. Its SHA-256 is
61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757.
Target slide 100/shape 0 has payload SHA-256
26c9e1dffc347568407eb56d54c87d5fddb8cdb895a867fffb3cd9f302be18e6.
Outputs and source counters were stable on every leg and output hashes were
equal between measured implementations. The exact compact evidence is
[the machine-readable ABBA result](../results/pptx-selected-slide-retained-snapshot-0375-abba.json).

The signed candidate-minus-control changes are stored for every case and
leg in the compact result. One-edit directions are -4.030781326% and
+1.757396997% at p50; batch directions are +4.083604925% and +18.464971741%;
multi-slide directions are -4.456289122% and +5.805660219%. The corresponding
mean directions are -3.795428917% and +1.174067002%, -3.176441963% and
+22.819940725%, and -5.486793615% and +14.877154731%.

The required same-implementation drift gates are p50/mean <=5%, p95 <=10%,
and p99 <=15%. Directions conflict for every mean and all one-edit
statistics; batch candidate p50/mean drift fails; multi-slide control p50 and
candidate mean drift fail. Some tail cells pass numerically, but they remain
observations only. performance_claim is therefore none.

The raw reports serialize zero correctness, semantic, preservation, refusal,
or error booleans across all 12 rows. Exit 0 proves only successful harness
completion. Output-hash equality proves byte identity for the generated
outputs, not semantic validity, reopen, preservation, refusal, or
reversibility.

## Claim boundary

No latency, allocation, RSS/heap, reads, decompression, materialization,
physical-I/O, cold-cache, throughput, scaling, fixed-memory, general-OOM,
all-PPTX, real-producer, topology/media/notes/theme/chart, or parallel claim
is made. The ABBA package supplies timing observations only and does not
establish semantic or refusal correctness. The retained-snapshot result is
limited to the focused tests, independent reviews, exact source-backed
selectors, and their existing bounded publication path. Temporary raw reports
are deleted after the compact evidence commit.
