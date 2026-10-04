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
snapshot with the two retained lockfile overlays. The 0908 plan must bind the
existing 0907 source and binary receipts by their recorded paths and hashes;
duplicating those receipts in this packet is unnecessary. The relevant harness
source in that snapshot has SHA-256
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

## The two source calls and the collected owner scope

The owner is a harness helper, and the selected runner has two syntactic calls
to it in the frozen source even with zero warmups and one sample. The calls
have the same demangled caller and owner name, so function name or dump order
alone cannot identify the timed call. The exact out-of-line owner filter in the
retained binary does establish a narrower scope: first-child raw admission
observed one positive owner edge at source position 51174 and no collected
owner edge or dump at 51143. The latter absence is valid for this exact
profile scope. The raw evidence does not establish the compiler reason for
that absence, so this review makes no such claim.

| call | frozen source callsite | source and sink | clock role | profile role |
| --- | ---: | --- | --- | --- |
| expected-output construction | `tools/perf-baseline/src/lib.rs:51143` | `OwnedSource` and `Vec<u8>` | before the `Instant` at line 51172 | no collected owner edge observed |
| measured iteration | `tools/perf-baseline/src/lib.rs:51174` | `InstrumentedSource` and bounded `CountingSink` | inside lines 51172–51175 | measured dump |

The expected call constructs `expected`, checks the expected materialization
count, and feeds the semantic/output gates. It is real execution of the same
source-level helper, but it is outside the native elapsed interval and uses a
different source and writer. It produced no collected owner dump in the
accepted binary. That is a scope observation, not evidence that the call did
no work, and it must not be reconstructed, subtracted from, added to, or
treated as a repeat of the measured call.

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

The reader records the source position on the `calls=` edge. Classify a
collected owner edge only when its source file is the frozen
`tools/perf-baseline/src/lib.rs` and its position is exactly `51143` or
`51174`. Position 51174 is the measured scope. Position 51143, if a future
capture emits it, is an unmeasured expected-call diagnostic. A dump whose
owner edge has `calls=2`, an unknown position, a different source file, or an
unknown positive incoming parent is ambiguous and must fail admission. Do not
classify by numbered dump order. The accepted first child has exactly one
positive measured owner dump, no expected-call dump, and a zero-cost
program-termination dump.

Preserve every dump and account for every positive owner edge. A later child
may not silently discard an additional edge; it must classify it by callsite
and keep it out of the measured row. The measured role must occur exactly
once. Absence of the expected role is valid under the retained exact-symbol
scope and is not a reason to manufacture a second diagnostic or to subtract
anything from the measured count.

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
   measured call at 51174. The accepted scope has exactly one measured 51174
   edge. A 51143 edge is retained as an unmeasured diagnostic if it appears;
   its absence is valid. Unknown positive edges are not lifecycle rows to
   discard: they stop admission.
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

`--collect-atstart=no` excludes setup outside a collected owner call, while
`--zero-before` and `--dump-after` make each collected owner invocation a
separate counter interval. They do not make the expected and measured calls
equivalent: their source adapters, writer types, and surrounding state differ.
The accepted child has one measured Ir total at 51174 and no expected Ir total.
If a later raw dump at 51143 appears, report it separately as an unmeasured
diagnostic and never subtract it from the 51174 result.

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
symbol is not unique, the measured owner edge cannot be identified at its
frozen 51174 callsite, the measured edge occurs more than once or not at all,
any raw dump has an unknown positive owner edge, or the semantic report fails.
Retain the failed receipt and raw artifacts; do not repair them by subtracting
an expected call or by selecting the largest/smallest dump. An absent 51143
expected-call dump is valid under the accepted exact-symbol scope.

If all gates pass, the packet may publish one measured Callgrind Ir row per
route and, only when raw evidence emits one, a separate 51143 expected-call
diagnostic. It may describe which named children account for the measured
guest instructions. It may not call those rows native latency, convert them
into a release speedup, or claim that the 0907 eager/source wall-clock
separation has a demonstrated causal hotspot.
