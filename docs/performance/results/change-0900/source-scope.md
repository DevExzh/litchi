# 0900 source scope

`performance_claim: none`. This is the source and oracle boundary for the
current native phase refresh. The native lane is the only lane qualified in
this batch; allocation and CPU profile lanes remain separate follow-ups after
the phase ranking. This note records no optimization or historical speedup
claim.

## Selected route and corpus

The selector is
`xlsx_source_backed_cell_values_one_percent_edit_save`. Dispatch reaches
`run_xlsx_cell_values_edit_save` at
`tools/perf-baseline/src/lib.rs:25328-25340`. It uses the fixed source-edit
generator `litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`
(`lib.rs:199-200`) and the two opt-in shapes in `XlsxCellCrudShape::ALL`
(`lib.rs:503-521`): `medium` and `dense-sparse`. Producer full-cell and
producer-edit cases are different routes and are outside this selector.

The inventory is built in `lib.rs:22504-22554` and one-percent positions are
chosen deterministically in `lib.rs:22557-22577`; the case requires the
manifest update count to equal `ceil(total cells / 100)`
(`lib.rs:23536-23573`). Therefore the current shapes contain:

| Shape | Inventory | One-percent updates |
| --- | ---: | ---: |
| medium | 4 sheets × 48 × 48 = 9,216 | 93 |
| dense-sparse | 16,384 + 1,024 + 256 + 128 = 17,792 | 178 |

`xlsx_update_sheet_selectors` deduplicates the worksheet numbers in the
updates (`lib.rs:23612-23623`); these one-percent vectors touch all four
worksheets. The corpus builder is `lib.rs:22723-22837`: it creates four-sheet
numeric workbooks, strips calculation properties, and adds **8** deterministic
incompressible PNG members of 512 KiB each (`lib.rs:317-318,
22580-22588, 22754-22765`). The current medium capture binds this accounting
to `native/b00-medium/report.json`: its corpus has 17 archive members and its
source oracle reports 12 untouched members. Thus the 12 untouched members are
8 media members plus 4 other unchanged members; they are not a 12-media count
and do not imply a corpus change from 0705.

The source is an `InstrumentedSource` over the in-memory corpus archive,
constructed outside the operation clock (`lib.rs:44578-44592`). Its counters
describe logical positional reads and cache behavior. They do not measure a
filesystem, physical storage, decompression, remote I/O, native Office, or a
cold cache.

## Clock and oracle boundaries

Untimed setup and preflight in `lib.rs:44494-44573` includes lifecycle gates,
the eager expected output and semantic digest, source range/budget derivation,
cache-limit construction, and managed refusal replay. For each source-backed
iteration, sink reservation and source cloning are also before the clock
(`lib.rs:44574-44602`).

The measured native total is the sum of four intervals:

* open: `SourceBackedEditor::from_read_at...`,
  `lib.rs:44603-44625`;
* planning: `editor.edit_sheets(selectors)`,
  `lib.rs:44626-44635`;
* commit: every `edit.set` followed by `edit.commit()`,
  `lib.rs:44648-44674`;
* publication: `publish_multi_commit_to_stream`; its returned publication
  snapshot is dropped before the timer stops, `lib.rs:44705-44716`.

The emitted output is checked after the operation clock. The reopen interval
starts at `lib.rs:44832-44853` and includes `verify_xlsx_cell_crud_output`,
semantic hashing against the eager expected output, and source-preservation
checks; it is retained as `reopen_ns` and excluded from the native total.
The largest sink write bound and digest determinism are also checked after the
clock (`lib.rs:44854-44894`). The recorded scope string is at
`lib.rs:44777-44789`.

The output oracle verifies semantic cell values and package identity
(`lib.rs:23632-23641, 23720-23825`), including content types, relationship
signatures, part set, untouched media, and the explicitly allowed calculation
invalidation. Raw ZIP member identity and the changed-member set are checked
by `xlsx_cell_crud_untouched_member_evidence`
(`lib.rs:23828-23949`). The semantic hash traverses the complete inventory
(`lib.rs:23951-23988`). The one-time lifecycle gates cover exact no-op,
clear/remove, source lineage, stale/foreign source, inverse patch, and related
refusal behavior (`lib.rs:23990-24045` and following gate body).

## Counters and follow-up instrumentation

Per-sample source evidence retains total, workbook, selected-worksheet and
unselected-worksheet read calls/bytes, payload materializations, cache hits,
cold and successful loads, waiter joins, failures, evictions, bypasses,
retained entries/bytes, in-flight loads, and budget snapshots
(`lib.rs:7101-7163`). These are evidence for this logical provider and are
not physical-I/O or process-RSS metrics.

The implementation exposes separate allocation regions for planning, staging,
commit core, combined commit, and publication
(`lib.rs:44627-44634, 44648-44672, 44705-44715`; evidence fields at
`lib.rs:7331-7360`). Meaningful values require the dedicated allocator-metrics
binary. Native runs must not interpret unavailable samples, allocator-instrumented
elapsed time, region peaks, or summed peaks as native latency or whole-process
memory. The case result deliberately has no generic operation metric
(`lib.rs:44885-44895`). Allocation and scoped Callgrind/perf work should be
started only after this native phase identifies the owner.

## Historical evidence and freshness

The retained medium report is
`docs/performance/results/change-0900/native/b00-medium/report.json` (SHA-256
`210d7dd642121972320c950cc10e60694b8078f233575f8d03daf48fa9626b54`); its
case, generator, shape, archive-member count, and untouched-member count are
the binding facts above. 0705 is the closest prior same-selector refresh. It documents the same
in-memory source and timer exclusions (`docs/performance/0705-xlsx-edit-save-baseline-refresh.md:15-21`),
with historical native p50s of 16.656 ms (medium) and 32.763 ms
(dense-sparse), and phase shares near planning 36–38%, commit 21%, publication
41–43% (`0705:43-64`). Its allocator table and publication Callgrind are
separate evidence, not native timing (`0705:72-131`). Its 12 untouched-member
observation is consistent with the current report; the relevant distinction is
the current eight-media/four-other breakdown and 17-member archive.

0707 profiles the same selector and identifies
`worksheet_xml_and_parse_source` as an inclusive planning owner (~96%) with
`Validator::observe` nested around 18–19%; nested rows overlap and Callgrind
guest instructions are not latency (`docs/performance/0707-xlsx-planning-attribution.md:77-106`).
0708 rejected the validator name-storage candidate: its best total p50 change
was 2.147%, while the corresponding planning p50 change was only 0.210%
(`docs/performance/0708-xlsx-validator-name-storage.md:16-42`).

Later source changes make a fresh current phase necessary. 0747 changed
publication audit reuse, but explicitly reports no usable window for scattered
one-percent edits and a one-percent publication result within noise
(`docs/performance/0747-xlsx-publication-audit-reuse.md:83-108,337-340,484-488`).
0750 tightened source-audit well-formedness and records a dense-sparse one-edit
publication cost; that is a different edit cardinality and audit revision
(`docs/performance/0750-xml-audit-well-formedness-gaps.md:21-45,750-764`).
0764 hardened XML/MCE attribute and namespace handling; its worksheet/MCE
measurements are historical to that diagnostic (`docs/performance/0764-xml-attribute-dos-hardening.md:263-299`).
0842 changed runtime compression dispatch and reports generic OPC results, not
semantic XLSX timing (`docs/performance/0842-opc-runtime-compression-dispatch.md:1-18,47-53,76-82`).

The 0891–0899 evidence is producer full-cell, selected-worksheet metadata, or
XML-validation attribution. It must not be mixed into this scalar cell-values
route; for example 0891's x14ac owner is explicitly a producer full-cell
profile (`docs/performance/0891-x14ac-namespace-resolution.md:12-24`).

The phase refresh should therefore rank current `open`, `plan`, `commit`, and
`publication` from matched native samples first. Only then should a separate
allocator-metrics capture and a narrowly scoped planning/publication profile
be selected, preserving the exact oracle and source identities above.
