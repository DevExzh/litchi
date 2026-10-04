# 0908 — PPTX multi-slide instruction attribution

The existing eager workflow spends about **69.67% of its collected guest
instructions in `Package::to_bytes`** and **21.73% in the initial opened
snapshot**. Source inspection explains the serialization cost: the measured
ReadAt-to-materialized-OPC route discards archive preservation provenance.
The owning `from_vec`/`from_reader` APIs already retain it. Qualifying that
missing owning-input control is the next step before proposing a production
change. This batch adds evidence only; it adopts no implementation change.

## Frozen scope and capture

Four serial Callgrind 3.26.0 children reuse the exact
[0907 executable and source](0907-pptx-multi-slide-current-baseline.md), built
from `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab` plus retained lockfiles.
Binary SHA-256 is
`e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84`.
No build or native timing capture was run in this batch. CPU 12, one worker,
zero warmups and one sample per child are fixed in eager/source/source/eager
order. The [plan](results/change-0908/plan-v2.json), command/start receipts,
raw reports and dumps retain exact arguments, tool identities and hashes.

The exact-symbol toggle is
`litchi_perf_baseline::publish_pptx_multi_slide_batch_edit`, with collection
initially off and the same full symbol supplied separately to `--zero-before`
and `--dump-after`. The literal string `same` is not a Callgrind alias.
Each child has one positive numbered dump: its single incoming owner edge
comes from measured source position `51174`. Expected-output position `51143`
has no collected owner edge. This observation qualifies the emitted binary;
no compiler-inlining explanation or subtraction of expected work is assumed.
The initially proposed two-dump protocol is superseded by the retained
[first-child admission](results/change-0908/first-child-admission.json) and
[amended review](results/change-0908/protocol-review.md).

Every owner self-plus-immediate-child sum and every full function-self sum
equals the dump summary and totals. Each final program-termination dump has
zero instructions. Raw stderr has `brk segment overflow` warnings and repeated
instruction diagnostics from the parent and a runtime-probe child. All four
workload processes exit zero and their semantic receipts pass. Stderr counts
are never summed; only the qualified target dumps supply the table below.
This run is not a memory-checking result.

## Collected instructions

| Leg | Route | Owner Ir |
| --- | --- | ---: |
| eager-a | materialized ReadAt | 4,577,995,784 |
| source-a | source-backed editor | 108,165,330 |
| source-b | source-backed editor | 108,181,418 |
| eager-b | materialized ReadAt | 4,578,507,829 |

The dominant immediate-owner edges are a disjoint partition. Other children
and owner self cost remain in the complete [analysis](results/change-0908/analysis.json).

| Route / direct callee | First leg Ir | Second leg Ir |
| --- | ---: | ---: |
| Eager `Package::to_bytes` | 3,189,476,327 | 3,189,730,670 |
| Eager `opened_presentation_with_limits` | 994,945,242 | 995,093,332 |
| Eager `into_opc_package_inner` | 98,544,032 | 98,553,345 |
| Source `publish_slide_batch_commit_to_stream` | 55,654,082 | 55,654,546 |
| Source `SourceBackedPresentationEditor::from_read_at` | 38,232,988 | 38,246,730 |
| Source `SourceBackedSlideBatchEdit::set_shape_texts` | 13,962,259 | 13,963,877 |

The source edges account for approximately 51.45%, 35.35% and 12.91% of
collected source-route instructions. Eager serialization reaches large Deflate
bodies; its initial snapshot reaches full-revision fingerprinting. Nested
function costs overlap their parents and cannot be added to this table.
Merged global function call counts also retain zero-cost edges: they do not
establish repeated serialization merely because a nested call count exceeds
one. Harness-labelled inclusive edges may contain production replay work;
classification by source path is not a pure production/overhead split.

`Ir` means guest instruction attribution under Valgrind. It is not native
elapsed time, cycles, allocated bytes, RSS or a release speedup. CPU dispatch
under instrumentation may differ from native execution. These fractions do
not explain the native latency ratio quantitatively.

## Semantics and timer boundary

The fixture and edit are unchanged from 0907: 200 slides, eight text boxes per
slide, eight deterministic 2 MiB media Parts, and 64 replacements across eight
selected slides. All four outputs are 17,017,145 bytes with SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.
Corpus identity, source counters, sink counters and generic operation metrics
match their native references after constant 200-sample vectors are reduced
to one sample. Semantic reopen and untouched logical OPC Part checks pass.
This does not establish independent raw ZIP header/comment/order fidelity.

The owner includes open, edit, commit, publication and the route-specific
inverse/forward or inverse/stale guards described in 0907. Expected-output
construction and full post-helper verification remain outside the collected
owner. Source materialization count is sampled after prepublication replay
but before publication; it is not an allocation or full-verification count.
The generic operation-metrics envelope exists, but route-specific phase
vectors are absent. No guard work is subtracted from the totals.

The first workload succeeded, but its original wrapper compared a 200-sample
native metrics envelope directly with the one-sample profile and failed.
The original helper, plan, failure and raw capture remain retained. The v2
validator checks native vector constancy before normalizing the sample axis;
all statuses and other values must match exactly. An intermediate validation
attempt that missed nested sink vectors is also retained. The first raw
capture was revalidated, not rerun; the next three use that corrected wrapper
with identical profiler options. These are evidence-validation corrections,
not workload or performance changes.

## Next bounded qualification

The [source audit](results/change-0908/eager-opportunity.md) follows the eager
route through `SourceBackedPackage::into_opc_package` into a fresh
`OpcPackage` without the retained physical archive. `PackageWriter` already
has an exact-source shortcut and a preservation writer, but this route cannot
supply their provenance. By contrast, owning `Package::from_vec`,
`from_reader` and `open` retain it; borrowed `from_bytes` does not.

Qualify an owning-input eager control with the same corpus and edits, explicit
input-copy timing, separate production/replay boundaries, and physical
copy/regenerate evidence. Preserve unknown members, stale-source refusal,
limits, durable inverse and sink-failure guarantees. This will determine
whether an eager provenance bridge is useful before duplicating existing
preservation machinery. Complete-revision snapshot validation remains
required; the profile does not authorize skipping payload fingerprints.

## Reproduction, review and cleanup

The [reducer](results/change-0908/analyze.py) reuses the retained 0845
Callgrind grammar reader and validates command, tool, binary, helper, plan,
raw artifact, owner and semantic bindings. Outputs are create-only; use a new
path with `--output` to replay. The existing binary is retained locally under
0907 and is not Git-tracked; rebuilding a fresh clone requires the 0907 source
and build receipts. The [source replay](results/change-0908/source-replay.json)
reconstructs its 9,350-file snapshot and rechecks the native packet.

Independent review and offline reducer checks are retained alongside the raw
captures. The owned `/tmp/litchi-0908-run` directory is removed; this batch
created no worktree or build root. Unrelated worktrees, files and the shared
iWork cherry-pick are untouched. This evidence batch is committed on a child
of the 0907 Office branch through an isolated index; target integration remains
pending. The broad OLE2/OOXML goal stays open; ODF is deferred and iWork excluded.
