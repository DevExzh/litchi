# 0908 PPTX Callgrind protocol review

This is a finite, read-only review of the planned current-source PPTX
instruction profile. It does not build, launch a workload, collect a profile,
or change production or harness code. The profile is an attribution lane for
the existing 0907 workflow; it is not a new timing capture or an optimization
claim.

## Frozen identity and schedule

Use the retained 0907 release executable with SHA-256
`e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84` (264,596,744
bytes). Its source is the complete `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab`
snapshot with the two retained lockfile overlays. The 0907 source receipt and
the binary receipt must be copied into the 0908 packet and checked before any
profile is admitted. The relevant harness source in that snapshot has SHA-256
`f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea`.
The dirty shared checkout is not a substitute for those bindings.

The owner is the single demangled symbol
`litchi_perf_baseline::publish_pptx_multi_slide_batch_edit`. A preflight
`nm -C --defined-only` check must find exactly one defined symbol with that
name in the retained binary. The raw Callgrind records must also resolve the
same owner name. If the symbol is absent, duplicated, merged into its caller,
or has no usable source callsite records, the profile is not admissible.

Run four separate processes serially on CPU 12, in this order:

| leg | selector | role |
| --- | --- | --- |
| eager-a | `pptx_eager_multi_slide_batch_edit_save` | first eager profile |
| source-a | `pptx_source_backed_multi_slide_batch_edit_save` | first source-backed profile |
| source-b | `pptx_source_backed_multi_slide_batch_edit_save` | second source-backed profile |
| eager-b | `pptx_eager_multi_slide_batch_edit_save` | second eager profile |

Each child uses `--workers 1 --warmup 0 --samples 1`, a fresh JSON output
path, and a fresh Callgrind output path. The external invocation is equivalent
to:

```text
taskset --cpu-list 12 valgrind \
  --tool=callgrind --collect-atstart=no \
  --toggle-collect=litchi_perf_baseline::publish_pptx_multi_slide_batch_edit \
  --zero-before=litchi_perf_baseline::publish_pptx_multi_slide_batch_edit \
  --dump-after=litchi_perf_baseline::publish_pptx_multi_slide_batch_edit \
  --callgrind-out-file=<fresh-path> <retained-binary> \
  --workers 1 --warmup 0 --samples 1 --case <selector> --json <fresh-path>
```

Keep the four raw dumps, child commands, environment, affinity, tool hashes,
binary hash, source/lockfile receipts, stdout, and stderr. Do not retry or
replace a child after seeing an inconvenient profile.

The fresh native 0907 result, 263 ms eager versus 7.8 ms source-backed, is
context only. It is not a Callgrind denominator and must not be combined with
the instruction counts below.

## The two owner calls in every child

The owner is a harness helper, and the selected runner calls it twice in the
frozen source even with zero warmups and one sample. The calls have the same
demangled caller and owner, so function name or dump order alone cannot tell
them apart.

| call | frozen source callsite | source and sink | clock role | profile role |
| --- | ---: | --- | --- | --- |
| expected-output construction | `tools/perf-baseline/src/lib.rs:51143` | `OwnedSource` and `Vec<u8>` | before the `Instant` at line 51172 | preflight/expected dump |
| measured iteration | `tools/perf-baseline/src/lib.rs:51174` | `InstrumentedSource` and bounded `CountingSink` | inside lines 51172–51175 | measured dump |

The expected call constructs `expected`, checks the expected materialization
count, and feeds the semantic/output gates. It is real execution of the same
helper, but it is outside the native elapsed interval and uses a different
source and writer. It must not be subtracted from, added to, or treated as a
repeat of the measured call.

The measured call is the call surrounded by `Instant::now()` and
`elapsed()`. Its Callgrind dump is the only profile that can be labelled as
the existing end-to-end operation. It still includes the helper's input
construction and qualification guards, because those occur inside
`publish_pptx_multi_slide_batch_edit`:

* source-backed mode opens the editor, stages eight slides times eight text
  replacements, commits, replays the patch and inverse before publication,
  publishes, then validates inverse success and stale forward refusal;
* eager mode opens and materializes the OPC package, stages the same 64
  replacements, commits, publishes, replays inverse and forward, checks the
  slide count, serializes with `to_bytes`, and writes 64 KiB chunks.

The callsites are distinguishable in raw Callgrind edge records. The reader
records the source position on the `calls=` edge; classify an owner edge only
when its source file is the frozen `tools/perf-baseline/src/lib.rs` and its
position is exactly `51143` or `51174`. The former is expected and the latter
is measured. A dump whose owner edge has `calls=2`, an unknown position, a
different source file, or an unknown positive incoming parent is ambiguous and
must fail admission. Do not classify by numbered dump order. If compiler line
tables coalesce both callsites, the protocol cannot prove the distinction and
the child must be rejected.

With this source and selector, admission should find two positive owner dumps
per child, one for each callsite, plus the zero-cost program-termination dump.
Preserve every dump even if the count differs. An extra owner call, a missing
expected or measured dump, a nonzero termination dump, or any unexplained
owner parent is evidence of a protocol mismatch rather than a reason to pick a
convenient subset.

## Raw profile admission

Use the retained 0845/0901 raw reader or an equivalently fail-closed reader;
do not make `callgrind_annotate` text the authority. For each positive dump,
require all of the following:

1. The header has only `Ir`, positive `summary`, equal `totals`, and the exact
   `--dump-after=<owner>` trigger. Its command and binary path match the
   packet's captured command.
2. There is exactly one owner function record and exactly one positive incoming
   edge for that owner. The edge caller is
   `litchi_perf_baseline::run_pptx_multi_slide_batch_edit_save`, its `calls`
   value is one, and its inclusive Ir equals the raw summary.
3. The edge is classified by `position` as the expected call at 51143 or the
   measured call at 51174. Both roles occur once. Unknown positive edges are
   not lifecycle rows to discard: they stop admission.
4. The owner's self Ir plus its immediate direct-child inclusive Ir exactly
   reconstructs the owner edge. The sum of all function self costs exactly
   reconstructs the raw summary. Nested child rows are retained as diagnostic
   detail but are not added to the immediate-child partition.
5. The terminal dump follows the positive dumps, has the termination trigger,
   and has zero summary and totals. Its zero cost is not a measured operation.

The JSON report for the measured child must independently pass its normal
schema, one-result, requested-selector, `samples=1`, `warmup=0`, binary,
corpus, output, sink, materialization, source-read, semantic, and digest
checks. Those gates validate the workload result; they do not identify which
Callgrind dump is measured. The raw callsite admission supplies that missing
identity.

`--collect-atstart=no` excludes setup outside an owner call, while
`--zero-before` and `--dump-after` make each owner invocation a separate
counter interval. They do not make the expected and measured calls equivalent:
their source adapters, writer types, and surrounding state differ. Report the
two Ir totals separately and use only the 51174 row for measured-operation
attribution.

## What the measured owner does and does not establish

Callgrind `Ir` is a guest-instruction count. It is not milliseconds, native
cycles, allocation count, RSS, cache misses, physical I/O, or a predicted
speedup. One profile sample per route cannot establish variance or a confidence
interval. CPU affinity to a shared CPU reduces placement variation but does not
establish an exclusive host, cold-cache result, filesystem or network latency,
cross-machine behavior, or native Office compatibility.

The owner is a broad harness boundary. Its measured dump includes both caller
overhead and production work, while semantic reopen, full selected/unselected
Part comparison, digesting, and output preservation checks after line 51175
are outside the owner and native timer. The profile therefore qualifies a
guarded workflow boundary, not a pure edit, pure publication, or pure ZIP
serialization cost. Source-backed and eager routes also intentionally have
different open/materialization and patch-guard contracts; their totals are
descriptive route profiles, not a causal proof of the 0907 wall-clock ratio.

LTO and inlining can remove or merge a candidate symbol. Missing symbols must
be recorded as unavailable; a nearby name or an inclusive parent is not a
substitute. Inclusive counts for nested functions overlap their parents, so a
phase table must use either a disjoint immediate-child partition or explicitly
label nested rows as diagnostics. It must never sum the owner, a direct child,
and that child's descendants.

## Candidate phase attribution targets

The owner dump can locate hypotheses, but it cannot create phase boundaries
that the helper does not expose. Retain self and inclusive Ir for the
following exact targets when they appear in the raw graph, together with their
source file and incoming callsite. Report a target as absent when LTO removes
it; do not infer its cost from a similarly named function.

| phase | source-backed targets | eager targets | interpretation |
| --- | --- | --- | --- |
| catalog open / materialization | `litchi_pptx::presentation::source::SourceBackedPresentationEditor::from_read_at` (`source.rs:1023` and its inner open) | `litchi_opc::source_backed::SourceBackedPackage::from_read_at` and `into_opc_package` (`source_backed.rs:6718,7853`), then `litchi_pptx::Package::from_opc_package` (`package/codec.rs:264`) | Open/materialization work; the eager route is expected to have a much larger materialization path. |
| selected-slide load and validation | `SourceBackedPresentationEditor::edit_slides`, `slide_snapshot_from_retained_catalog`, `SourceBackedSlideBatchEdit::set_shape_texts`, and `opened::xml::stage_shape_texts` (`source.rs:1713,2143,2720`; `opened/xml.rs:754`) | `Package::opened_presentation_transaction` and `opened::Transaction::set_shape_texts` (`package/model.rs:296`; `opened/transaction.rs:263`) | Selected slide parsing, selector resolution, raw-span planning, and rewrite/readback. |
| commit and source authorization | `SourceBackedSlideBatchEdit::commit`, `SourceBackedSlideBatchPatch::try_apply_fast`, `apply`, `is_changed`, and `same_source` (`source.rs:2790,2893,2880,2848`) | `Package::apply_opened_presentation_commit` and `apply_opened_presentation_patch` (`package/model.rs:550`; opened patch replay) | Exact-source, inverse, stale, and changed-state checks. The harness's explicit replay calls are included in the owner and must remain visible. |
| source publication planning and output | `SourceBackedPresentationEditor::publish_slide_batch_commit_to_stream` (`source.rs:1813`), `SourceBackedPackage::write_part_overlays_shared_to_stream`, `validate_overlay_limits`, `preservation_index_with_limits`, and XML replacement audit (`source_backed.rs:10505,10705,11772`) | `litchi_pptx::Package::to_bytes` (`package/codec.rs:360`), `litchi_opc::pkgwriter::to_bytes` (`pkgwriter.rs:1296`), and the package topology writer | Overlay/limit/preservation planning versus eager package serialization. These are candidate costs, not proof that a scan is removable. |
| harness boundary | `semantic_pptx_text`, `pptx_batch_replacements`, `cache_diagnostics`, and the inverse/stale guards in `lib.rs:51053-51079` | replacement construction and inverse/forward/slide-count guards in `lib.rs:51089-51114` | Caller and qualification overhead. Keep it in the current owner total and label it before comparing production hypotheses. |

The most useful next attribution is a disjoint owner partition separating
catalog open, selected-slide staging/commit, exact-source authorization,
publication planning, and output writing. A named child may be nested under a
publication parent, so it can be shown as a diagnostic row without being added
to that parent's inclusive Ir. If this profile shows repeated
`same_source`/`is_changed` work or repeated overlay-limit/preservation scans,
that is a bounded hypothesis for a later correctness-first change. It is not
permission to remove byte comparisons, stale-source checks, XML audits,
limits, raw-copy preservation, or inverse guards.

## Stop conditions and interpretation boundary

Stop the profile lane if the source/binary/lock bindings differ, the owner
symbol is not unique, either owner call cannot be identified by its frozen
callsite, any raw dump has an unknown positive owner edge, the expected and
measured roles do not each occur exactly once, or the semantic report fails.
Retain the failed receipt and raw artifacts; do not repair them by subtracting
the expected dump or by selecting the largest/smallest dump.

If all gates pass, the packet may publish two route-specific Callgrind Ir rows
for the measured call and separate expected-call diagnostics. It may describe
which named children account for those guest instructions. It may not call
those rows native latency, convert them into a release speedup, or claim that
the 0907 eager/source wall-clock separation has a demonstrated causal hotspot.
