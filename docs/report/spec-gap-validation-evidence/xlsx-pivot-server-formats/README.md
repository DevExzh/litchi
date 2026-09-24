# XLSX pivotTableServerFormats owner

The public ordinary API is selector-first and workbook-contextual:

```rust
let view = workbook
    .pivot_table_server_formats(PivotTableSelector::Name("Pivot"))?;
let mut edit = workbook.edit_pivot_table(PivotTableSelector::Name("Pivot"))?;
edit.update_server_format(
    0,
    ServerFormatEdit {
        culture: PivotServerFormatAttributeEdit::Set("en-US".to_owned()),
        format: PivotServerFormatAttributeEdit::Clear,
    },
)?;
let commit = edit.commit()?;
let workbook = commit.workbook();
```

`Name` and `Position` are semantic workbook selectors. Ordinary callers do not
supply OPC relationship IDs, package URIs, XML ranges, or extension URIs. The
source-bound `Snapshot`, `Transaction`, `Patch`, and package functions remain
an advanced exact-source layer for callers that explicitly need those details.

This batch recognizes the C510 `pivotTableServerFormats` owner under a direct
PivotTable-definition extension and the exact non-worksheet
`pivotTableReferences` closure. The closure resolves the semantic
`PivotCacheId`, requires the external cache and ABF5 cache-version extension,
and validates the supported F057 source-connection route when present. The
payload is the qualified `x15:serverFormat` collection.

Reads expose the ordered typed collection and optional decoded `culture` and
`format` values. Scalar writes use `Keep`, `Set`, and `Clear` so either
optional attribute can be preserved, added/replaced, or removed. List writes
use `insert_server_format`, `push_server_format`, `remove_server_format`,
`move_server_format`, and `reorder_server_formats`; reorder inputs map each
final position to its prior position. Structural publication updates the
required `count` exactly, refuses removal of the final child, preserves known
`pivotValueCellExtra@in` associations, and refuses opaque or removed
associations. Opaque siblings, namespaces, comments, processing instructions,
and unchanged lexical bytes remain source material. Container and Part
lifecycle operations remain additional work in the wider PivotTable extension
audit. The owner does not refresh or calculate PivotTables or render them.

The implementation recognizes only the evidenced C510/983426/725AE2AE/ABF5
closure and the optional F057 owner. It does not guess newer extension URIs or
index mappings. Native producer compatibility remains unverified; synthetic
schema-shaped fixtures do not establish it.
