# 0530 XLSX `edit_sheets` planning source review

This is a bounded source audit for the OLE2/OOXML workstream. It covers the
native planning interval in the XLSX source-backed cell-value harness and the
production call path below it. ODF remains deferred and iWork is out of scope.
No Rust source was edited, and this review ran no build, test, benchmark,
allocator capture, or profile capture. The retained 0522 cell-scanner/fusion,
0527 row-arena, and 0529 XML-attribute-probe mechanisms are rejected and are
not candidates here.

## Identity and dependency binding

The current revision is `655410c7b657dcf1b8fa4fe4dda071f37e6c3031`
(`655410c7b`), whose parent is `3f4c7be06159dc3d742d9c815800dc55bf7c2db6`.
The current commit adds four public XML-minifier stream-audit tests in
`crates/xml-minifier/tests/stream_audit.rs`, publication allocation metrics and
tests, and their documentation/results. Relative to the retained 0529 final
source baseline, the XLSX/OPC production files and the planning harness are
byte-identical; `tools/perf-baseline/src/lib.rs` has not changed since the 0529
final candidate. These hashes were read from the current checkout.

`docs/performance/results/change-0529/source-binding.json` is an XML-minifier
binding record only (SHA-256
`449d68eb8b80dd82024eaeadc748a4775c4d195f3763f4e4e3ec85cb690bbc16`); it has
no XLSX or OPC entries. The authoritative cross-format binding for this review
is `docs/performance/results/change-0529/final/source-manifest.json`,
SHA-256 `c5a7cfbd0d7da965c7d945aea3355d24e34685b91f9dbbb25baedbed97140a08`.
Its XLSX/OPC and harness entries match the current checkout.

| input | SHA-256 or identity |
| --- | --- |
| `crates/litchi-xlsx/src/cell_values/source.rs` | `10c9a99892dea1cf5c0dd529f628313126c4908c0fdc6e3694439c1c3889213b` |
| `crates/litchi-xlsx/src/cell_values/snapshot.rs` | `2f3f839adc91f0da204aefc83ebe2bb605cc02d269346759155b7bda54abc0e9` |
| `crates/litchi-xlsx/src/workbook/source.rs` | `7f2f8ac3cc417d7f3458b8fd281541155391e8e973eed763728a4996728f704d` |
| `crates/litchi-xlsx/src/source_payload.rs` | `a9d89ca9b4c5b09b4c8710d4cb322c7a45f0b1feb73b047b9ed82a60b34a47d0` |
| `crates/litchi-xlsx/src/cell_values/validation.rs` | `52e7d2d18f59e716c686f1c4c59b5835632fe6981dc6789ef7b06b555d136c0b` |
| `crates/litchi-xlsx/src/raw/catalog.rs` | `12318c600336a7331f47feed9ed89866987e8c6868e23dab1a5890d6945b7240` |
| `crates/litchi-xlsx/src/raw/worksheet/mod.rs` | `7ed9276c713b8fcf8c84722bc62e58696912a11a1f02360e011059b830cf9b73` |
| `crates/litchi-opc/src/source_backed.rs` | `6f37f9a1e2ecd0435cc811b56a3e92d4bbdb0b41bcdf755478d04f2455f8ea48` |
| `crates/litchi-opc/src/package.rs` | `b2910a82ceab2db28c3b22ab9afcdf857248cc5613612834af2fa43c74ee898d` |
| `tools/perf-baseline/src/lib.rs` | `bdeeeef53b276cb72d813ff3c7345a302e72e0d8575e13d6757d07875189585f` |
| `tools/perf-baseline/Cargo.toml` | `a04de024b9cbe9683cdb7307c3d7199daab7171bbdcb6bb8c30b361766857aca` |
| `tools/perf-baseline/Cargo.lock` | `13333f511914d8146c60282b5d6385693cf89c0f939fa52b12e4db821c9b8b36` |
| `crates/litchi-xlsx/Cargo.toml` | `d8139ad1d4abd39e3aa0a09c60fc7a70aa251ab026fa4f2e793b68228fcb3d9c` |
| `crates/litchi-opc/Cargo.toml` | `0572d27f2134b4d060e79ac2bbd57ae3d839ad768e57a944d5674538388b80fb` |
| `crates/litchi-core/Cargo.toml` | `f269d5d26b041dca9d589342f3b0f3cdd796ea4c9afb72b09bf9b1933001d8ac` |
| `Cargo.toml` | `911a52cf6932b81550bc9ffd6e522c327dec178297ee9bc46d2ccedb693d4885` |
| `Cargo.lock` | `9111221ee9d100daf90328a544613cb3f70287611dcc55a37d3b1b7a5d99c91a` |
| `rust-toolchain.toml` | `e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310` |
| `crates/litchi-ooxml-common/src/mce/codec.rs` | `e5911cdcd94116474b09062af96d639b8f08d80a09a26d90c86a931a7344fa74` |
| `crates/litchi-ooxml-common/Cargo.toml` | `e226484cae42a827ff8d37603870ebd088cfaba953338825664433290d04489e` |
| pinned `quick-xml` | `0.41.0`, registry checksum `e660451e55124f798a69a5af3f49ccfbefbd41910eefd25caf2393e1f3473ec1` |
| local `quick-xml-0.41.0` source tree | ordered-file digest `e3e17f7a072e642b2fe8d3418e4f95f0bc9accba8b880babe940202185414696` |
| pinned `memchr` | `2.8.3`, registry checksum `cf8baf1c55e62ffcace7a9f06f4bd9cd3f0c4beb022d3b367256b91b87513d98` |
| local `memchr-2.8.3` source tree | ordered-file digest `173488633e4f13447ff873388624886c89a7737a7ed69d0e691e4109eb25c002` |

The accepted policy inputs read for this review are `docs/GOAL.md`
(`bed4058bb76330daab8ce9d4bceff639ab3fbd7ea06634158bef41b133c4d1f1`),
ADR 0001 (`813bf4faf62cedac56fbfa3bf9b7600121b42a3ed5d4cec0028f8c8f2ac2798e`),
ADR 0002 (`73f4372ef6948f5a36c83fbf27a96f0b1f680225f331f201c470e04913417719`),
ADR 0003 (`9200d3546d91f1604d44a9ade7ad0a6bdea50b6442d12a0517b09db5e8e802ef`),
ADR 0005 (`34a6148a8fe77b3e90212810996667b654e25209fb49c56aaecd8d8fbe83f770`),
ADR 0006 (`b686465f342b2e051f856f094e38c2333aa05a2768bb255fc5b223cbba1f3381`),
ADR 0008 (`5c72d79a7ba78a044471b40ba22ab3f32f7b8f4f878614ab05b416e60b0163a2`),
ADR 0010 (`4e1af2c4019842d187804f8356299aa80ba8a6abd8d35a4a021e1bce3c2a6590`),
ADR 0011 (`3a1644536af0b66fed47c1a154e829c0e0008cccb89dec679a0328036d305e04`),
and ADR 0024 (`bdeefc52ba2486ca307c378e1081baafac5565c327b622981782f1aed74c55ce`).

Those records require correctness, preservation, bounded resources, typed
refusals, source-version and cancellation fences, and separately measured
planning evidence. They also place semantic XLSX ownership in `litchi-xlsx`
and physical OPC indexing and publication ownership in `litchi-opc`.

## What the harness actually times

The measured source-backed cell CRUD path is
`tools/perf-baseline/src/lib.rs:41760` (`run_xlsx_cell_values_edit_save`). In
the current source the relevant sequence is:

```text
open timer      41869..41891  SourceBackedEditor::from_read_at...()
selector build 41892        xlsx_update_sheet_selectors(&updates)
plan timer      41893..41897 editor.edit_sheets(selectors)
commit timer    41909..41927 edit.set(...) loop + edit.commit()
publication     41956..41978 publish_multi_commit_to_stream(...)
```

`xlsx_update_sheet_selectors` at `lib.rs:21568-21579` collects sheet
positions, sorts, deduplicates, and maps to `Selector<'static>`. It completes
before `plan_started`, so selector-vector allocation, sort, deduplication, and
mapping are outside `plan_ns`. They must not be charged to a planning
optimization. `edit_sheets` includes the wrapper execution check, source
catalog/snapshot construction, and the empty `MultiSourceEdit` state
construction; staging and commit are outside the plan interval.

The source-backed lifecycle runner at `lib.rs:41217` is a different timer:
`lib.rs:41280-41293` starts before `edit_sheets` and also includes address
construction plus `clear` or `remove` staging. Its `timing_scope` says
`selector planning/staging`; it cannot serve as an exact `edit_sheets` owner
measurement. Gate calls at `lib.rs:21952`, `21972`, `22024`, `41197`,
`43537`, `43586`, and `66888` are correctness/setup or refusal replays, not
the measured CRUD planning interval.

The measured CRUD cases dispatch through `run_case` at `lib.rs:23273-23285`.
The eight source-backed variants are one edit, one-percent, bounded batch,
two-worksheet, and the four managed equivalents. The managed and unmanaged
variants use the same `Vec<Selector<'static>>` planning path; only package
construction and execution/cache policy differ. The deterministic medium and
dense/sparse corpora select one to four worksheet owners. The prior 0529
retained phase context put planning near one third of its measured open/plan/
commit/publication phase sum, but that was rejected-pilot context and supplies
no current owner ranking or performance claim.

## Caller, generic, and owner ancestry

The production ancestry for the timed path is:

```text
tools::run_case
  -> run_xlsx_cell_values_edit_save
    -> SourceBackedEditor::from_read_at...       [open, outside plan]
      -> SourceBackedEditor::edit_sheets<'a, I>  [plan root]
        -> MultiSnapshot::load_source_backed<'a, I>
          -> load_source_catalog
          -> resolve_selectors
          -> Snapshot::from_source_selected       [once per selected sheet]
          -> MultiSnapshot::from_sheets
        -> MultiSourceEdit::new
```

`SourceBackedEditor::edit_sheets` is at
`crates/litchi-xlsx/src/cell_values/source.rs:470-476`. It calls
`package.check_execution()` and then forwards to
`MultiSnapshot::load_source_backed`; it does not itself construct the
selector vector. `MultiSourceEdit::new` at `source.rs:875-882` only creates an
empty `BTreeMap` and zero staged-cell count, so it is inside the timer but is a
small fixed tail.

The `I` monomorphizations relevant to attribution are distinct:

| caller | iterator type | timing/use |
| --- | --- | --- |
| `run_xlsx_cell_values_edit_save` | `Vec<Selector<'static>>` | measured source-backed CRUD planning |
| `run_xlsx_cell_lifecycle_edit_save` | `[Selector; 1]` | measured interval also includes staging |
| lifecycle/no-op/partial gates | `Vec<Selector<'static>>` or `[Selector; 1]` | untimed correctness paths |
| `SourceBackedEditor::edit_many` at `source.rs:479-527` | `Vec<Selector<'a>>` | production batch caller; staging follows planning |
| XLSX integration tests | arrays and `Vec`, including name selectors | correctness only |

`MultiSnapshot::load_source_backed` is generic over the same `I` and consumes
it into an owned selector vector. For Callgrind or `nm -C`, bind the concrete
demangled symbol emitted for the measured `Vec<Selector<'static>>` instance;
generic symbol suffixes may include compiler hashes. Do not use an array
instance or the lifecycle wrapper to infer the CRUD plan owner.

The ownership boundary is:

```text
litchi_core: Selector, SourceVersion, ExecutionContext/resource checks
    -> litchi_opc: SourceBackedPackage, immutable OPC catalog, ReadAt/cache
        -> litchi_xlsx: workbook/cell semantic catalog and source closure
            -> tools/perf-baseline: timer and evidence serialization
```

`SourceBackedEditor` owns a `SourceBackedPackage` by value. The format crate
does not own a concrete ZIP implementation. Any optimization must leave OPC
part lookup, physical source freshness, cache reservations, and typed OPC
errors below the existing `litchi_opc` boundary.

## Measured planning attribution

The retained analysis is
`docs/performance/results/change-0530/planning-analysis.json`, SHA-256
`a160d478f7b1e00a472d8ce907e2a404118bf754f97a35839b2363ff6308593c`, with
helper SHA-256 `5a5ba155c9eb775cd8653f41b10ef9b3e52e8fc871ea3df01166fba2e5c269a7`
and plan SHA-256
`18c516de7aa700c7e256de75c195fe183dd1fa8c1cdb277f4e8318f3912f9754`. It
reports `status: pass` for the selected `edit_sheets` method owner. The four
selected Callgrind dumps contain one measured CRUD `Vec<Selector<'static>>`
edge each; three numbered dumps in each group are lifecycle cases and a
separate unnumbered termination dump is zero-work. This review uses the
selected edge, not the lifecycle wrapper.

The following values are direct child edges from the selected owners unless
marked inclusive. They are instruction references from one generated medium
or dense/sparse corpus, not wall time, allocation counts, or a scaling law.
Nested inclusive rows overlap and must not be added together. The two repeats
are close enough to establish owner shape, but they do not prove a speedup.

| corpus/repeat | `edit_sheets` plan | catalog direct | selected worksheet raw parse direct | selected worksheet XML validation direct | `PartView::data` direct | raw parse -> `process_ooxml` direct |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| medium/r1 | 125,643,633 | 453,206 (0.36%) | 75,518,721 (60.1%) | 46,622,053 (37.1%) | 2,860,921 (2.28%) | 6,517,217 (5.19%) |
| medium/r2 | 125,655,087 | 452,744 (0.36%) | 75,530,690 (60.1%) | 46,621,647 (37.1%) | 2,861,139 (2.28%) | 6,517,120 (5.19%) |
| dense/sparse/r1 | 237,075,784 | 454,069 (0.19%) | 141,675,512 (59.8%) | 89,730,098 (37.9%) | 4,872,623 (2.06%) | 12,269,361 (5.18%) |
| dense/sparse/r2 | 237,102,562 | 454,319 (0.19%) | 141,699,182 (59.8%) | 89,732,515 (37.8%) | 4,872,586 (2.06%) | 12,269,284 (5.18%) |

The MCE detector owner
`litchi_ooxml_common::mce::codec::process_markup_compatibility` has inclusive
IR 6,466,820 (medium) and 12,219,020 (dense/sparse); its direct
`__memcmp_avx2_movbe` edge is 3,772,174 and 7,127,624 respectively. The
`process_ooxml` owner is 6,575,427 and 12,327,569 inclusive in the first
repeat, with the second repeat within the same range. These are the complete
call edges, including required capability setup and return/cleanup. Only the
raw substring-search implementation inside the MCE no-namespace branch is a
candidate; the full 5.2%-of-plan call is not removable. Likewise, the 5.2%
figure is a share of this planning profile, not a whole-workflow gain.

`load_source_catalog` is only about 0.36% of medium planning and 0.19% of the
dense/sparse planning in these samples. Catalog caching remains relevant to a
separate repeated-editor workload, but it is demoted from the primary
single-call opportunity. `Snapshot::from_source_selected` owns more than
99.6% inclusive of the selected loader in both shapes, and its raw parser and
full XML validator are required work under the current contracts. The profile
metadata also warns that collection-off call edges are not operation counts;
no allocation or call-count conclusion is drawn from them.

## Exact planning data path

### Selector capture and source fences

`MultiSnapshot::load_source_backed` at
`crates/litchi-xlsx/src/cell_values/snapshot.rs:179-232` first checks execution
and captures `package.source_version()`. It reserves bounded selector and
snapshot vectors (`MAX_SHEET_OWNERS == 64`), checks execution for each input
selector, rejects more than 64 owners, then calls `load_source_catalog`.
After every selected worksheet is loaded it checks the aggregate
`MAX_MULTI_WORKSHEET_BYTES == 64 MiB` bound, constructs `MultiSnapshot`, checks
execution again, captures the final source version, and rejects a changed
source. The per-selector checks and final version comparison are part of the
atomic source snapshot contract; removing them would change cancellation or
TOCTOU behavior.

### Workbook catalog and closure capture

`load_source_catalog` at `snapshot.rs:1845-1883` performs this ordered work:

1. `package.check_execution()` and `package.main_document_part()` resolve the
   unique package office-document owner. `main_document_part` is the OPC
   physical relationship contract, not just a workbook-name lookup.
2. `validate_package_relationships` rejects external/unsupported package
   relationships and requires one office-document owner. The workbook content
   type must be ordinary XLSX.
3. `workbook.data()` obtains the exact source payload through
   `SourcePayload::from_part_data`. `SourcePayload::Managed` retains the
   managed `PartData` reservation; the unmanaged path moves it into a shared
   `Arc<Vec<u8>>` without a semantic copy.
4. `validation::workbook_xml` performs bounded namespace, depth, text, root,
   and closing-element validation. `raw::parse_catalog` then parses the
   workbook sheet catalog and selector-visible metadata.
5. `validate_sheet_graph` at
   `crates/litchi-xlsx/src/workbook/source.rs:1318-1368` walks every workbook
   sheet, resolves its relationship and Part, checks relationship/content type
   and kind, and rejects duplicate Part targets.
6. `validate_workbook_relationships` checks all workbook relationship types and
   external flags and counts worksheet, styles, theme, and calculation-chain
   relationships. `unique_owner` checks the package owner contract again in the
   XLSX error domain.
7. `capture_auxiliary_source` reads and retains styles/theme Parts and parses
   the style count; `capture_calculation_chain_source` validates and retains a
   calculation-chain Part when present.
8. `capture_sheet_graph_source` at `snapshot.rs:1783-1812` walks every sheet
   again, copies names/IDs/visibility/kind/URI/content type, captures each
   sheet Part's relationship records, and stores the complete graph in an
   `Arc<[SheetGraphState]>`.
9. The catalog copies workbook/package/workbook relationships into sorted
   `SourceRelationship` arrays, retains workbook XML and auxiliary payloads,
   and records source lineage/version.

The full graph and auxiliary closure is retained even for a one-sheet edit.
`Snapshot::SourceState` at `snapshot.rs:1290-1304` compares and publishes
workbook, worksheet, owner, package/workbook relationships, calculation chain,
auxiliary Parts, and the complete graph. Skipping unrelated graph or auxiliary
state because the edit touches one sheet would weaken topology/provenance and
raw-preservation checks.

### Selector resolution and selected worksheet materialization

`resolve_selectors` at `snapshot.rs:1761-1780` resolves names or checked
positions, rejects an unresolved selector and duplicate position, and sorts
positions for deterministic snapshot order. Position selectors are O(1) after
the workbook catalog, while name selectors canonicalize names and linearly
scan the raw catalog. The current harness has already sorted and deduplicated
its position list outside the timer, but the public `edit_sheets` contract
still accepts arbitrary iterators and name selectors.

For each resolved position, `Snapshot::from_source_selected` at
`snapshot.rs:426-478` checks execution, rejects charts/dialog/macro/unknown
sheet kinds, obtains the worksheet Part, refuses worksheet relationships, reads
the exact worksheet payload, enforces the remaining aggregate byte bound, and
runs both `validation::worksheet_xml` and `raw::worksheet::parse`. It then
checks execution, validates cell/row/column style references, and rejects
unknown cells or cell/value metadata through `validate_scalar_cells`. The
result retains the parsed `Store`, selected worksheet bytes, workbook bytes,
graph, relationships, auxiliary Parts, calculation-chain state, source
lineage, and source version. Scalar cells, ordinary formulas, and shared-formula
readback therefore depend on this parse and refusal path.

`MultiSnapshot::from_sheets` at `snapshot.rs:148-177` checks nonempty and
64-owner bounds, sums worksheet bytes, sorts by sheet position, rejects
duplicates, and boxes the snapshots. `resolve_selectors` and the load loop
already provide sorted unique positions on this path, but `from_sheets` is
also used by other construction paths and is the final internal invariant
check.

## Mandatory work versus optimization hypotheses

| work in the plan interval | classification | reason a change would need proof |
| --- | --- | --- |
| execution checks, source current/version checks, owner and size limits | mandatory safety/resource fences | cancellation, source TOCTOU, typed limit errors, and managed budget behavior |
| workbook XML validation and catalog parse | mandatory | selector resolution, dialect/namespace/error policy, and workbook metadata |
| complete sheet graph validation and capture | mandatory state; duplicate traversal is a hypothesis | publication must retain all sheet descriptors, relationship targets, kinds, content types, and per-Part relationships |
| workbook/package relationship validation and capture | mandatory state; repeated scans may be reducible | external/unsupported relationship refusals and exact source topology are separate contracts |
| styles/theme and calculation-chain capture | mandatory when present | style-reference validation and exact source-preserving publication |
| selected worksheet read, XML validation, raw parse, style scan, scalar refusal | mandatory | parsed cells/formulas, unknown/metadata refusal, selected payload retention, and commit readback |
| selector resolution duplicate check and deterministic ordering | public contract | arbitrary name/position iterators, duplicate/error ordering, and stable output order |
| `MultiSnapshot::from_sheets` sort | likely redundant on this private load route, but shared invariant | other callers may be unsorted; a fast path needs a private precondition and duplicate check |
| `MultiSourceEdit::new` empty map/count | fixed tail, low ROI | transaction API state; no evidence of a material allocation owner |
| `xlsx_update_sheet_selectors` Vec/sort/dedup/map | outside `plan_ns` | timer begins after this helper at `lib.rs:41893` |

The fact that two passes or two checks touch the same metadata does not make
either one removable. Existing error order is observable: for example,
`validate_sheet_graph` reports a missing/external/wrong-kind/duplicate sheet
before later closure-copy failures, and workbook relationship validation
reports unsupported/external relationships before auxiliary capture. A fused
implementation must preserve those boundaries or explicitly establish a new
typed contract.

## Candidate opportunities and blockers

### 1. Replace the MCE no-namespace substring probe — selected next

The concrete hot edge is
`crates/litchi-ooxml-common/src/mce/codec.rs:530-549`, inside
`process_markup_compatibility`. After the input-byte limit at lines 535-537,
the current no-MCE fast path scans the complete byte slice with
`xml.windows(NAMESPACE.len()).any(|w| w == NAMESPACE.as_bytes())`. If it finds
no match, it applies the output-byte limit and returns the original slice as
`Cow::Borrowed`; only a match enters the bounded MCE parser. The profile
retains a large direct substring-comparison edge (`__memcmp`) and separate
owner self instructions, while
the complete `process_ooxml` call also includes capability setup, parser work,
and cleanup.

The narrow design is to replace only that predicate with the existing
`memchr` implementation, conceptually:

```rust
if memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_none() {
    // Keep the existing output-limit check and borrowed return unchanged.
}
```

`memchr` is already a direct dependency of `litchi-ooxml-common` and is pinned
by the hashes above, so this candidate needs no dependency or public-API
change. `memmem::find` returns a match exactly when the same nonempty,
contiguous byte sequence occurs. It therefore has the same result for empty
or shorter input, a match at offset zero or the last valid offset, repeated
matches, near-matches, and namespace bytes in text or comments. The latter are
deliberately included: the current probe is a raw byte test rather than an XML
semantic test. Keep the private `find_bytes` helper unchanged; it serves
active-offset paths and has not been measured as part of this candidate.

The branch order is part of the proof. Input-size rejection must remain before
the search; a no-MCE output-size rejection must remain after the search; and a
match must still enter the existing parser with the same `Capabilities`, MCE
directives, `Report`, output ownership, and bounded allocations. This change
does not bypass MCE validation, change namespace resolution, or alter source,
cancellation, provenance, or OPC resource fences in the XLSX caller.

The current owner ancestry is:

```text
Snapshot::from_source_selected
  -> raw::worksheet::parse
    -> mce::codec::process_ooxml
      -> mce::codec::process_markup_compatibility
        -> no-MCE substring predicate
```

`process_ooxml` and the underlying function are shared by DOCX, PPTX, and XLSX
callers. The planning profile shows the raw-parse edge at 5.18-5.19% of the
selected `edit_sheets` plan and the MCE owner at roughly 5.15% inclusive, but
only the search implementation is potentially reducible. These figures do not
justify a whole-workflow or XLSX-only speed claim. A follow-up must measure
no-MCE and MCE-containing parts for each materially affected format before
acceptance.

Required differential tests for a candidate implementation are:

* Empty, shorter-than-namespace, exact offset-zero, exact last-offset,
  repeated-match, false-prefix, and near-match inputs must choose the same
  branch. Inputs with a match in text/comment bytes must continue to choose the
  parser branch, matching the existing raw predicate.
* An input over `max_input_bytes` must return the input limit before either
  search implementation. A no-MCE input over `max_output_bytes` must return
  the output limit after the search. A matching document whose transformed
  output exceeds its limit must retain the bounded-output error and ordering.
* For valid and malformed MCE documents, compare output bytes, `Report`,
  `Cow::Borrowed` versus `Cow::Owned`, typed errors, and resource-limit
  behavior. Include `AlternateContent`, `Ignorable`, `ProcessContent`,
  `PreserveElements`, `PreserveAttributes`, `MustUnderstand`, unknown
  namespaces, deep nesting, DTD/PI, and custom-entity cases already covered by
  the MCE tests.
* Exercise the shared `process_ooxml`, `process_part`, and `process_str`
  callers, plus raw worksheet parsing with scalar, formula, unknown, and
  x14ac-extension cases. Existing worksheet and cross-format tests remain
  required; no DOCX/PPTX result may be inferred from the XLSX planning sample.

The unapplied draft is retained as `namespace-search.patch`, with exact source
and patch hashes in `next-pilot-source.json`. It has not been built or measured. It should be
profiled as one predicate substitution with the same release flags, corpus,
limits, and output checks. Do not combine it with parser fusion, validation
skipping, `find_bytes`, graph reuse, or any rejected prior mechanism.

### 2. Certified worksheet parser bypass — deferred, not proven safe

The successful value-only worksheet validator rejects MCE and x14ac elements
and attributes (while allowing namespace declarations), so it suggests that a
private validated-byte parser could avoid a second generic MCE pass. The raw
worksheet path cannot assume that fact today: `raw::worksheet::parse` first
checks `x14ac::may_contain_descent`, captures extension values when needed, and
then calls the generic `process_ooxml`; on rejected plain worksheets it repeats
the extension capture to preserve typed error precedence. The generic parser is
also shared with callers that may receive MCE content.

A certified route would need a private proof-bearing handoff tying the exact
bytes and validation limits to `validation::worksheet_xml`, while preserving
x14ac capture, malformed-input precedence, scalar/formula/shared-formula
behavior, source retention, and all resource fences. The validator result alone
does not establish those obligations. No bypass or validator/parser fusion is
recommended from this profile; the latter also overlaps the rejected 0514
fusion mechanism.

### 3. Graph reuse and catalog caching — measured secondary work

`load_source_catalog` still validates and later captures the complete sheet
graph, so a private validated-result reuse design may be worth a later audit.
However, its measured direct edge is only about 0.36% of medium planning and
0.19% of dense/sparse planning here. It is therefore not the primary
single-call opportunity. Any graph fusion must preserve validator error
precedence, duplicate-target checks, exact order/visibility/kind/content type,
per-Part relationships, fallible allocation and execution checks, source
version fences, and publication graph equality. The relevant ancestry remains:

```text
MultiSnapshot::load_source_backed
  -> load_source_catalog
    -> workbook::source::validate_sheet_graph
    -> capture_sheet_graph_source
      -> SourceBackedPackage::part
      -> capture_relationships
```

Catalog reuse has higher potential only for a separate repeated-editor
workload. A cache would need source-lineage/version invalidation, cancellation
and execution-context handling, managed `PartData` accounting, bounded parsed
store policy, and exact no-op/provenance behavior. The current fresh-editor
CRUD harness cannot demonstrate that ROI.

### Work explicitly rejected as an optimization

Do not remove workbook/worksheet XML validation, raw worksheet parsing, style
reference checks, unknown/metadata refusal, full graph or auxiliary capture,
calculation-chain capture, source-version checks, or per-selector cancellation
checks. They are required by the source-backed snapshot and publication
contracts. Do not attribute selector-vector construction to `plan_ns`. Do not
revive the 0522 cell-scanner/fusion, 0527 row-primary-arena, or 0529
XML-attribute probe; their dispositions are already rejected.

## Isolated Callgrind plan

No additional profile was run in this source audit. The retained profiles use
a current-HEAD release binary, the existing medium and dense/sparse
source-backed CRUD cases, and the measured `Vec<Selector<'static>>`
`edit_sheets` specialization. A candidate comparison must keep that timer and
corpus fixed and must not charge selector construction or lifecycle staging to
the plan.

For the selected MCE candidate, compare an unchanged baseline with the single
`windows`-to-`memmem::find` predicate substitution. Bind these demangled
owners independently so direct/self costs are not conflated:

```text
litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets
  litchi_xlsx::cell_values::snapshot::Snapshot::from_source_selected
    litchi_xlsx::raw::worksheet::parse
      litchi_ooxml_common::mce::codec::process_ooxml
        litchi_ooxml_common::mce::codec::process_markup_compatibility
          memchr::memmem::find       [candidate symbol, if emitted]
          __memcmp_avx2_movbe        [baseline comparison edge]
```

Run this lane for no-MCE worksheets, MCE-containing worksheets, and malformed
or limit-bound MCE inputs. Report direct/self and inclusive IR for the
predicate, `process_markup_compatibility`, `process_ooxml`, raw worksheet parse,
and the timed plan root. Also record selected worksheet count, workbook and
worksheet payload sizes, selector kind, managed/unmanaged policy, and
source/cache diagnostics as untimed evidence. A source-read count is not an
operation count and does not prove semantic materialization; the current
analysis explicitly excludes that interpretation.

Accept the predicate only if branch decisions and all output bytes, reports,
borrowed/owned results, typed errors, input/output limits, and MCE resource
behavior are identical across the differential matrix. Include the existing
XLSX scalar/formula/shared-formula and x14ac cases, then measure DOCX/PPTX
shared callers before making a cross-format claim. Do not treat a lower total
`process_ooxml` edge as proof that required MCE parsing or worksheet work was
removed.

Graph-result reuse can be revisited only in a separate lane if its direct
edges become material; the current catalog share does not justify combining
that work with the MCE predicate.

## Review disposition

The current `edit_sheets` timer is correctly bound to the source-backed
planning call, with selector-vector construction outside the interval and
lifecycle clear/remove staging kept separate. Most planning work is mandatory
source closure validation/materialization. The measured concrete next
opportunity is the single MCE no-namespace substring predicate substitution in
`process_markup_compatibility`; it has an exact byte-level equivalence proof
and a narrow owner boundary, but still needs differential tests and a fresh
candidate profile. The complete `process_ooxml` edge is not removable, and the
measured 5.2% share is not a whole-workflow gain. Catalog/graph reuse is
demoted to a repeated-workload or later isolated edge because its one-call
share is small. No production change is recommended by this source-only
audit.
