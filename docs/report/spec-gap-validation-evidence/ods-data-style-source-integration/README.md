# ODS source-qualified data-style integration

The v9 capture adds borrowed source scanning, source-qualified data-style queries, metadata edits, and extended graph authoring/replacement through the document transaction. Content automatic styles are editable; both styles.xml owners remain read-only. Effective cell data-style lookup uses content automatic styles followed by common styles. Metadata edits preserve opaque children and inherited namespace context; incompatible structural replacements refuse.

Ordinary Spreadsheet opening uses the shared normalized reader and retains the table title/description post-pass. Source tests cover XML grammar, schema-string whitespace, character-reference distinctions, literal CRLF normalization, exact semantic no-ops, and source-qualified byte ranges.

Graph replacement skips unchanged typed nodes, retaining their exact lexical bytes even in mixed changed/unchanged graphs. Replacement sizes use decoded content.xml ranges, not ZIP offsets. A Deflated-member regression puts the selected style beyond physical ZIP length and checks successful replacement plus exact inverse. Both replacement and insertion measure canonical markup and candidate XML before serialization. Physical package limits remain enforced by publication.

Root validation of `combined-validation-source.json`: 2,234 tests passed across 127 result groups, one ignored test; all-target Clippy with warnings denied and rustdoc with warnings denied passed. The public data-style vocabulary target contains 47 tests. Commands, raw logs, Cargo.lock and exact source hashes are retained. Common and ODT dependencies in the combined capture were committed separately.

Independent review approved the scanner. Owner review approved selectors, dependency closure, failed staging, stale/inverse checks, signature policy and facade integration, and the final v9 insertion preflight correction. The scoped owner review is approved. Earlier v7/v8 findings were corrected in this v9 capture; green results from earlier captures are not substituted for current validation.

Performance evidence is separate: the reviewed [source-workflow characterization](../ods-data-style-source-performance/README.md) is committed in `750b8d450`, and the [matched graph-preflight optimization](../ods-data-style-graph-preflight-performance/README.md) in `f2daff703` includes independent replay. Their workload and allocator caveats apply; the broader audit remains open.
