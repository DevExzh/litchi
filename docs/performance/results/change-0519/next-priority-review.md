# 0519 next-priority review: source-backed XLSX one-percent edit/save

`scope: read-only queue review after the 0511–0519 evidence`

`performance_claim: none`

## Decision

The next bounded capture should use the existing
`xlsx_source_backed_cell_values_one_percent_edit_save` selector over the
`medium` and `dense-sparse` cell-CRUD shapes. This is a common OOXML
edit-to-save path, and it exercises the source-backed route that is not
represented by a measured mapping in the current checked index. Keep this as
one source-backed route with two corpus shapes. Historical 0282 already ran a
12,000-row eager/source ABBA over these shapes and rejected selector-wide
speedup admission; the passing medium one-edit cell was diagnostic only. Do
not repeat that eager/source baseline as a speedup claim.

The proposed capture has a different purpose from 0282. That campaign compared
total eager and source-backed elapsed time in fresh children; it did not
compare phase vectors and explicitly made no allocation, physical-I/O, or
provider claim. The next capture should produce fresh current phase,
allocation, and provider attribution for the source-backed route before any
new optimization decision.

The first capture should establish stable phase, logical source-read, cache,
and correctness evidence. A production optimization should wait for a named
repeated owner in that evidence. If the result shows provider-sensitive work,
the follow-up should add one physical `FileSource` arm or one simulated
high-latency range arm, separately. The current runner does not support either
provider arm for this edit/save route.

## Why this is the highest-value remaining common scenario

The checked CRUD index has 15 categories and 34 scenario rows: 15 measured,
18 correctness-only, and one unsupported. The template-filling category
measures `xlsx_one_cell_commit` and `xlsx_one_percent_commit`, but those are
the older commit-only selectors. The index has no measured mapping for the
source-backed cell-values one-percent edit followed by publication and
reopen. [0093](../../changes/0093-xlsx-cell-values-matched-crud-evidence.md)
added this selector family as selectable correctness/phase evidence, and
[0282](../../changes/0282-xlsx-cell-values-abba-rejected.md) measured the
eager/source total interval, but neither provides the current operation-local
phase, allocation, or physical-provider attribution needed for a new
optimization decision. The selector remains opt-in and the 41-case default
omits it. This is therefore a current attribution and coverage-classification
gap, not a claim that the route has never been measured. The proposed work
supplies fresh evidence instead of repeating the rejected eager/source
baseline or the accepted commit-only measurement.

The corpus gives the case useful volume without requiring a new fixture. It
has four worksheets, eight untouched 512 KiB media entries, and 17 archive
members for the ordinary shapes. `medium` has 48 by 48 cells per sheet
(9,216 cells total); `dense-sparse` has 17,792 cells across its four sheets.
The update set is the harness-defined `ceil(1%)` set, and the source-backed
path tracks the selected worksheet ranges separately from unselected members.
The corpus is deterministic and generated, so it supplies reproducible
semantic and raw-member oracles while remaining explicitly short of a native
Office-producer claim.

The recent XLSX evidence makes this route actionable. 0512 attributed dense
one-percent commit work across source `Store`/validation parsing, rewrite,
and compaction; 0513 added operation-allocation evidence for the related
generic commit/save path; and 0515 showed that changed-output parsing and
compaction are substantial but that a shared reader alone targets only a
small part of the commit. Those reports do not attribute the current
source-backed cell-values edit/save route. They identify the work to measure,
not permission to repeat the rejected parser fusion.

## Why further DOCX publication microprofiling should wait

0519 reduced the named DOCX publication route to ZIP topology preservation and
raw source copying. The remaining Callgrind topology totals are about
1.31–1.35 million Ir, with `write_all_counted` around 533k Ir for p128 and
561k Ir for p512 and `read_exact_at` around 531k Ir. The fixture contains
roughly 512 KiB of raw media, so these leaves include required transfer work
near the memory-copy boundary. Their inclusive Ir does not identify a
removable algorithm.

The end-to-end share is already small after the 0517–0519 proof reuse. In the
`p512-k32-owned-repeated` route, candidate publication p50 is about 80–88 us
while candidate elapsed p50 is about 15.7–15.9 ms; the edit phase accounts for
nearly all of that lifecycle. In the batch routes, publication is roughly
75–77 us against 1.14–1.28 ms total. A further DOCX publication change would
need a concrete non-copy owner and an end-to-end opportunity before it outranks
the current attribution gap in the XLSX source-backed save path. 0511 CFB/FAT
work is complete,
and 0514/0516 XLSX fusion proposals remain rejected and are not requeued.

This does not close the DOCX provider, range, native-producer, or broader
topology requirements. It sets the next measurement order using the current
global lifecycle shares and the evidence already available.

## Existing case and code-owner boundary

The case is already declared and dispatchable in
[`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs):

| Concern | Existing evidence and owner |
| --- | --- |
| Selector | `Case::XlsxSourceBackedCellValuesOnePercentEditSave` is declared at `lib.rs:1215–1225` and named at `lib.rs:1745–1754`. It is recognized by the cell-values predicates at `lib.rs:3329–3357`; it is not in `Case::DEFAULT` at `lib.rs:1540–1582`. |
| Dispatch and timing | The cell-CRUD dispatch at `lib.rs:10062–10090` calls `run_xlsx_cell_values_edit_save`; the runner begins at `lib.rs:41630`. |
| Corpus | `build_xlsx_cell_crud_corpus` at `lib.rs:20578–20677` builds the four-sheet scalar grid, media payloads, fixed generator, archive manifest, and source ranges. |
| Source API | `crates/litchi-xlsx/src/cell_values/source.rs:222–244` exposes the `FileSource`-backed constructors; `:384–432` exposes the `ReadAt` constructors currently used by this runner; `:600–637` publishes a committed multi-sheet edit; `:818` and `:1122` own single- and multi-sheet commit construction. |
| Publication owner | `crates/litchi-opc/src/source_backed.rs:6573` owns `write_topology_to_stream`; the changed-overlay writer is at `:9514`. The XLSX source editor reaches this owner through `publish_multi_commit_to_stream`. |
| Provider harness | `tools/perf-baseline/src/filesystem.rs` has `xlsx_file_open`, `xlsx_file_open_lifecycle`, and `xlsx_file_selected_cell` arms, but no source-backed XLSX edit/save filesystem arm. `--filesystem-cache` therefore cannot be attached to this selector today. |

## Narrow missing measurement

The current runner already defines a useful phase boundary for the unmanaged
source-backed route:

* `open_ns` covers `SourceBackedEditor::from_read_at_with_limits_and_cache_limits`;
* `plan_ns` covers `edit_sheets`;
* `commit_ns` covers the staged `set` loop and `edit.commit()`;
* `publication_ns` covers `publish_multi_commit_to_stream` into the bounded
  non-seek `CountingSink`; and
* the reported overall elapsed vector is the sum of those four phases.

Corpus construction, expected eager output and semantic digest preparation,
the exact no-op/clear/remove/foreign/stale/inverse lifecycle gates, and sink
setup are outside that loop. Output reopen, semantic verification, package
identity, untouched-member raw ZIP identity, output hash, and source snapshot
collection happen after the timed phases. `reopen_ns` is recorded as evidence
but is excluded from the elapsed vector. This separation makes the next
capture an end-to-end edit/save measurement while keeping expensive proof
work from distorting the native phases.

The source counters and cache diagnostics are useful but have a defined limit.
They describe the logical `InstrumentedSource` and cache, not physical I/O,
device latency, or decompression. The current
`XlsxCellValuesIterationEvidence` records phase vectors, source calls/bytes,
per-region reads, materializations, cache hits/loads/evictions, retained
entries/bytes, and managed-budget fields, but it has no operation-local
allocator vector or hardware-counter vector. The 0513 allocation observations
belong to `xlsx_one_percent_commit_save` and must not be assigned to this
source-backed selector. A `perf stat` run around the current process would
also include setup and untimed oracles, so it is whole-process context rather
than operation-local evidence.

If allocation or hardware evidence is needed for admission, the minimal
harness-only extension is to bracket the same four phase calls in
`run_xlsx_cell_values_edit_save` (or one aggregate timed region first) and
label setup, reopen, and oracles outside the bracket. No production API or
source behavior needs to change for that measurement. Until such a bracket
exists, report allocator and hardware fields as unavailable.

## Recommended next capture

Run the existing selector with both fixed shapes and a larger sample count;
the command below is the proposed follow-up recipe and was not run during this
review:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 20 --samples 100 \
  --case xlsx_source_backed_cell_values_one_percent_edit_save \
  --xlsx-cell-crud-shape medium,dense-sparse \
  --json target/perf/next-xlsx-source-one-percent.json
```

Review the two shape rows separately. Require stable phase distributions,
unchanged source calls/bytes for the same source epoch, deterministic output
and semantic hashes, the expected touched worksheet set, the bounded 64 KiB
sink, and the existing cache/budget diagnostics. Keep the exact output and
semantic comparisons against the independently produced eager artifact. The
runner's existing gates also preserve exact no-op bytes, clear/remove state,
foreign-lineage refusal, stale-source refusal, inverse restoration, and
untouched media/ZIP-member identity.

Do not run the eager/source ABBA driver as part of this queue item. If a later
production candidate is proposed, compare the source-backed route against its
own baseline under the same source provider and corpus, then require a fresh
correctness/resource review. A separate provider follow-up can use
`SourceBackedEditor::from_path_with_limits_and_cache_limits` for one real
`FileSource` lane. A range-latency lane should use the existing range-provider
abstraction only after a source-backed edit/save adapter is added; the current
`--range-*` query controls do not cover this operation.

The proposed current capture should remain an opt-in correctness/phase
measurement. Only after its phase and resource boundaries are stable should
the checked catalog add a measured mapping for this selector or the 41-case
default grow. Native producer
coverage, formula/table/filter semantics, topology-changing edits, cold
filesystem behavior, and broader OOXML CRUD remain separate follow-ups.
