# XLSX pivot extension implementation reconciliation

This note reconciles the older “all zero matches” pivot-extension inventory in
`docs/report/spec-gap-audit.md` with committed implementation evidence. It does
not approve the current uncommitted `pivotTableData` implementation.

| Extension | Committed evidence | Remaining boundary |
| --- | --- | --- |
| `pivotTableServerFormats` | `4edeb3534` adds source-bound scalar edits; `363cbd9f5` adds ordered server-format leaf edits and known index-reference closure. | Container/Part lifecycle, refresh, calculation, and native producer interoperability are not established by this batch. |
| `cachedUniqueNames` | `dae6162fb` adds model-associated cache-field diagnostics, existing-name scalar edits, source-checked patches, and exact inverses. | The `CT_Items` index-bound anchor remains unresolved. Insertion, removal, and reindexing are refused; native producer interoperability remains unverified. |
| `pivotTableData` | The approved design is committed in `01a579c27`; implementation and tests are in progress in the shared worktree. | No production-readiness claim until implementation review and gates finish. |
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
owners and their exclusions. This reconciliation was checked against those Git
objects, not inferred from current dirty files or a stale test status. It does
not report a new Cargo run or extend prior validation to current work in progress.
