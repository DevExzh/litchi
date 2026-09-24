# XLSX pivot extension implementation reconciliation

This note reconciles the older “all zero matches” pivot-extension inventory in
`docs/report/spec-gap-audit.md` with committed implementation evidence, including
the reviewed scalar `pivotTableData` batch in `d077282c8`.

| Extension | Committed evidence | Remaining boundary |
| --- | --- | --- |
| `pivotTableServerFormats` | `4edeb3534` adds source-bound scalar edits; `363cbd9f5` adds ordered server-format leaf edits and known index-reference closure. | Container/Part lifecycle, refresh, calculation, and native producer interoperability are not established by this batch. |
| `cachedUniqueNames` | `dae6162fb` adds model-associated cache-field diagnostics, existing-name scalar edits, source-checked patches, and exact inverses. | The `CT_Items` index-bound anchor remains unresolved. Insertion, removal, and reindexing are refused; native producer interoperability remains unverified. |
| `pivotTableData` | `d077282c8` adds the ordinary Workbook sparse reader and existing-value/cell-extra scalar editor, caller-lowered semantic budgets, source-checked patches, and exact inverse. The [isolated gate](xlsx-pivot-table-data/final-gate/README.md) records 1,303 passing tests, strict Clippy/rustdoc, and independent semantic/resource review. | Structural matrix and container/Part lifecycle, native producer interoperability, refresh, calculation, and cube access remain unsupported or unverified. |
| Later pivot extension families | No completion claim in this reconciliation. | `implicitMeasureSupport`, `aggregationInfo`, `featureSupportInfo`, `autoRefresh`, and the 2025 family remain separate audit requirements. |

## Empty cached-name values

The committed contract distinguishes a missing required `name` attribute from
an empty `ST_Xstring` value. An existing name can be set to the empty string.
An earlier test-only checkpoint that expected rejection is superseded by the
committed contract; it is not a current correctness blocker.

This is demonstrated by the committed test
`existing_name_edits_allow_empty_xstring_but_reject_structure_and_unknown_index`
in `crates/litchi-xlsx/tests/cached_unique_names.rs` at `dae6162fb`. It checks
the empty value after reloading, preserves the other entry, and applies the
inverse to restore the exact original cache bytes. The same test refuses an
unknown entry index, insertion, removal, and reindexing.

The committed feature-matrix rows in
`crates/litchi-xlsx/docs/FEATURE_MATRIX.md` at `dae6162fb` describe these bounded
owners and their exclusions; the matrix now also records the C444 scalar owner.
The C444 gate additionally covers shared closure fixes for cached names and
server formats, including ownership-aware fragment caps and all recognized
duplicate payloads. This reconciliation refers to the recorded Git objects and
isolated test inputs. It does not extend their validation to concurrent
form-control work or the later pivot extension families.
