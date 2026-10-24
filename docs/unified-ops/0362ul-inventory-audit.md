# 0362ul — inventory owner coverage and disposition input audit

The inventory helper now reads the contract's `immutable_owners` policy paths
and resolves public aliases/re-exports by rustdoc nominal identity. It keeps
the historical owner-name heuristic, emits no duplicate owner rows, and still
excludes authoring builders. The disposition validator is a structural audit:
it checks a one-to-one inventory/disposition relation and concrete human
decisions, evidence and owners. It does not approve a disposition or close an
API-debt row. Item 22 is the only permitted deferral.

## Policy proof

This audit reads `tools/api_glossary.json` at the frozen source
`f6fe430a4d19d94aa360c0e9ed03ebd757265225`. The policy file is 847,191 bytes,
SHA-256 `0498ce055149c53c6982dedc9e3bfc142165edf8643c36e9291693fcbabdce9d`.
It contains 113 immutable-owner declarations, 91 unique package/path pairs,
and 32 declaration occurrences whose final type name is outside the legacy
`OWNERS` name heuristic. Those 32 occurrences represent 29 unique package/path
pairs across 11 packages; duplicates are the repeated feature/profile entries
in the policy, not additional public owners.

The 32 policy declarations are:

```text
litchi-xlsx: SourceBackedWorkbook, SourceWorksheet, Worksheet (each repeated once)
litchi-xls: SourceBackedWorkbook
litchi-ole-common: property_set::user_defined::Properties
litchi-ograph: PackageRef, WorkbookRef, Payload
litchi-ods: SourceBackedSpreadsheet
litchi-odp: SourceBackedPresentation
litchi-iwa-core: SnappyStream
litchi-iwa-archive: ComponentCatalog, SourceCatalog, OwnedSourceCatalog,
  LogicalSourceCatalog, FrozenDirectoryBundle, SemanticProjection,
  package::Catalog
litchi-iwa-structured: StructuredData
litchi-sheet-text: Worksheet
litchi-odraw: image::Blip, image::Entry, image::Store, image::Delay,
  shape::Shape, Parser, Container, prop::picture::View, prop::Props
```

The repeated XLSX declarations are intentionally counted as policy
declarations and deduplicated when resolving public paths. The implementation
does not infer that any of these owners is present in rustdoc; a missing or
ambiguous nominal path remains an inventory input error rather than an
accepted exception. The focused fixture proves a non-heuristic owner,
re-export and type alias produce one owner row and that a builder owner is
excluded from mutable-owner findings.

## Validator hardening

The TSV reader now rejects missing headers, short rows, extra columns, and
blank `crate`, `category` or `path` identity fields in both inventory and
disposition files. It pads malformed rows only to continue reporting all
structural errors, so malformed input cannot raise a `None.strip()` traceback
or hide extra columns. CRLF files remain valid, and empty `detail` and
`target` fields remain allowed.

The exact post-change test command was:

```text
python3 -B -m unittest tools.test_validate_api_disposition tools.test_api_debt_inventory tools.test_check_api_glossary
```

It ran 60 tests and passed. The complete stdout/stderr log is retained
outside Git at
`/home/zhuhe/code/litchi-worktrees/.coord/inventory-owner-coverage-55c562e9cf-followup.log`;
it is 159 bytes with SHA-256
`4e521a0f1bb41ae2a3457145b85068968ded05e68bb070a4f12a20c26aa363e5`.

Source commit `55c562e9cf4bc900c62787697eea974eef790201` contains the owner
coverage implementation and the initial validator/tests. Follow-up commit
`d2494012fc` adds the malformed-row hardening and this audit record. No
rustdoc capture, Cargo build, API-debt closure, or final acceptance claim is
made here.
