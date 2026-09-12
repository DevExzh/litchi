# 0520 profile review: source-backed XLSX one-percent edit/save

`scope: mechanism-level Callgrind evidence for one timed MultiSourceEdit::commit per profile`

## Result

The four-dump protocol is sound for the selected owner. Both repeats completed
for `medium` and `dense-sparse` with exit code 0, matching the frozen binary,
source manifest, native plan and shared capture helper recorded in the receipts.
The separate profile plan records the owner options used by those commands. The
checker reports `status: pass` in [`profile-analysis.json`](profile-analysis.json).

The external command used for each shape was:

```sh
taskset -c 2 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect=litchi_xlsx::cell_values::source::MultiSourceEdit::commit \
  --zero-before=litchi_xlsx::cell_values::source::MultiSourceEdit::commit \
  --dump-after=litchi_xlsx::cell_values::source::MultiSourceEdit::commit \
  --callgrind-out-file=profile.callgrind \
  /home/zhuhe/litchi-goal-0520-target/release/litchi-perf-baseline \
  --warmup 0 --samples 1 \
  --case xlsx_source_backed_cell_values_one_percent_edit_save \
  --xlsx-cell-crud-shape medium --json profile.json
```

The numbered dumps are scoped as follows: `.1` is the exact no-op gate, `.2`
is clear, `.3` is remove, and `.4` is the timed one-percent edit. Every dump
has one positive incoming edge to
`litchi_xlsx::cell_values::source::MultiSourceEdit::commit`; `.1`–`.3` come
from `run_xlsx_cell_value_lifecycle_gates`, and `.4` comes from
`run_xlsx_cell_values_edit_save`. The final unnumbered process dump has zero collected Ir
and is not operation evidence.

| Shape | `.4` inclusive Ir, repeat 1 / 2 | Owner self Ir, repeat 1 / 2 | `.4` incoming edge |
| --- | ---: | ---: | --- |
| `medium` | 220,444,234 / 219,884,708 | 5,692 / 5,692 | exactly 1 from `run_xlsx_cell_values_edit_save` |
| `dense-sparse` | 420,889,873 / 420,886,196 | 9,772 / 9,772 | exactly 1 from `run_xlsx_cell_values_edit_save` |

`MultiSourceEdit::commit` is the source owner at
[`source.rs:1122`](../../../../crates/litchi-xlsx/src/cell_values/source.rs#L1122).
The runner’s native `commit_ns` starts before the staged `edit.set` loop and
ends after `edit.commit()` at
[`lib.rs:41776`](../../../../tools/perf-baseline/src/lib.rs#L41776). The
Callgrind trigger starts only on entry to `commit`, so `.4` excludes the
staged `set`/`apply_batch` work even though that work is included in native
`commit_ns`. It also excludes publication, which starts at
[`lib.rs:41821`](../../../../tools/perf-baseline/src/lib.rs#L41821). This is
an owner profile, not a complete native commit-phase profile.

## Ranked direct callees

The rows below are direct inclusive callees of the timed owner, ranked within
each shape. Ranges cover the two repeats; call counts are stable across both
repeats. Ir means Callgrind guest instruction events.

| Rank | `medium` direct callee (Ir range; calls) | `dense-sparse` direct callee (Ir range; calls) |
| ---: | --- | --- |
| 1 | `litchi_xlsx::cell_values::snapshot::Snapshot::from_rewritten_source` — 144,117,632–144,706,503; 4x | `litchi_xlsx::cell_values::snapshot::Snapshot::from_rewritten_source` — 273,459,157–273,460,877; 4x |
| 2 | `litchi_xlsx::raw::worksheet::edit::package::rewrite` — 75,165,697–75,193,767; 4x | `litchi_xlsx::raw::worksheet::edit::package::rewrite` — 146,612,230–146,615,866; 4x |
| 3 | `litchi_xlsx::cell_values::snapshot::Snapshot::with_invalidated_workbook` — 158,156–159,168; 4x | `litchi_xlsx::cell_values::source::append_actions` — 225,579–225,923; 178x |
| 4 | `litchi_xlsx::cell_values::snapshot::Snapshot::invalidated_workbook_xml` — 138,756–141,450; 1x | `litchi_xlsx::cell_values::snapshot::Snapshot::with_invalidated_workbook` — 154,829–159,317; 4x |
| 5 | `litchi_xlsx::cell_values::source::append_actions` — 117,002–117,804; 93x | `litchi_xlsx::cell_values::snapshot::Snapshot::invalidated_workbook_xml` — 139,945–140,011; 1x |
| 6 | `litchi_xlsx::cell_values::source::effective_action_count` — 46,345–46,394; 4x | `litchi_xlsx::cell_values::source::effective_action_count` — 93,590–93,667; 4x |
| 7 | `litchi_xlsx::cell::Store::entry` — 21,855; 93x | `litchi_xlsx::cell::Store::entry` — 46,811; 178x |
| 8 | `litchi_xlsx::cell::Content::as_cell` — 21,391; 93x | `litchi_xlsx::cell::Content::as_cell` — 39,669–40,944; 178x |

The first two direct callees account for approximately 99.7–99.8% of the
owner’s inclusive Ir in both shapes. Within the first, the inclusive child
ranges are `raw::worksheet::parse` at 75,476,825–75,477,497 (`medium`) and
141,220,062–141,224,395 (`dense-sparse`), and
`cell_values::validation::validate_xml` at 68,631,527–69,221,071 and
132,225,956–132,232,117 respectively. Within the second,
`raw::worksheet::edit::codec::snapshot::scan::scan_with_limit` is
70,996,641–70,998,095 and 137,481,290–137,485,239 respectively. These are
inclusive child totals, not additional work to add to the direct-callee rows.

## Next owner priorities

1. Attribute allocations and inspect removable work in
   `Snapshot::from_rewritten_source`, using the already separated
   `raw::worksheet::parse` and `cell_values::validation::validate_xml` child
   costs. The implementation is at
   [`snapshot.rs:686`](../../../../crates/litchi-xlsx/src/cell_values/snapshot.rs#L686);
   worksheet validation begins at
   [`validation.rs:27`](../../../../crates/litchi-xlsx/src/cell_values/validation.rs#L27).
2. Attribute allocations and inspect removable work in
   `raw::worksheet::edit::package::rewrite` and its `scan_with_limit` child. Their implementation starts at
   [`package.rs:16`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/package.rs#L16)
   and [`scan.rs:200`](../../../../crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/scan.rs#L200).
3. Keep `append_actions`, `effective_action_count`, and cell-store bookkeeping
   below those parser, validation, and rewrite owners until a new profile
   shows a changed ranking.

For a follow-up owner capture, substitute one exact emitted owner name in the
same command and retain every numbered dump. Add
`--separate-callers=12` when the owner has shared lifecycle callers, then
admit a dump only after checking its raw incoming edge against the intended
runner parent. `edit_sheets` shares a monomorph with the no-op gate, and
`publish_multi_commit_to_stream` has several generic monomorphs, so a wildcard
toggle for either can conflate gates, the timed phase, or another route.
Function-name triggers also cannot bracket the un-named `edit.set` loop; a
full `commit_ns` Callgrind scope requires a new harness trigger around that
region or a separately validated staging profile.

## Limits

Callgrind Ir is not wall time, hardware instructions, cycles, allocations,
provider I/O, or latency. The retained stderr contains Valgrind `brk segment
overflow` notices even though all children exited successfully. The profile
therefore supports mechanism and owner ranking only; it does not support a
latency claim or an end-to-end speedup claim. The one-sample-per-shape profile
also does not replace the native statistical baseline.

The result follows the measurement boundary in [ADR 0005](../../../adr/0005-io-memory-and-performance.md)
and the XLSX/OPC ownership in [ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md)
and [ADR 0024](../../../adr/0024-current-topology.md). Do not resurrect the
0514/0516 proposals from these rows alone. Any revisit needs a new mechanism,
fresh source-backed preservation and resource evidence, and a new scoped
measurement.
