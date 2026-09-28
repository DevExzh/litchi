# 0806 visibility amendment source review

This archive repairs the API visibility regression found after applying the
0806 quality amendment. The retained quality check showed that
`litchi-crypto` imports the OLE helper's public `BytesStartExt` and
`CheckedAttributes`, while the candidate source had narrowed both declarations
to `pub(crate)`. This review is static: no Cargo, build, test, native capture,
or profiling command was run.

## Custody and exact scope

The parent source boundary is the recorded
`quality-amendment-application.json`, whose artifact hash is bound in
`manifest.json`. The amendment contains one production path:

`crates/litchi-ole-common/src/xml_attributes.rs`

The `before/` snapshot is byte-identical to the parent quality-amendment
snapshot. The `after/` snapshot differs by exactly two tokens:

```diff
-pub(crate) trait BytesStartExt {
+pub trait BytesStartExt {
...
-pub(crate) struct CheckedAttributes<'a> {
+pub struct CheckedAttributes<'a> {
```

The patch has no algorithm, field, method, test, comment, or formatting change.
It restores the two declarations to the visibility in the original 0806
baseline candidate source.

## Five-helper visibility audit

The original candidate `before/` snapshots and the current five helper sources
were inspected for all item declarations beginning with `pub`. Each helper has
only the `BytesStartExt` trait and `CheckedAttributes` struct as public item
declarations; their visibility comparison is:

| Helper | Baseline `BytesStartExt` | Current `BytesStartExt` | Baseline `CheckedAttributes` | Current `CheckedAttributes` |
| --- | --- | --- | --- | --- |
| `litchi-ole-common` | `pub` | `pub(crate)` | `pub` | `pub(crate)` |
| `litchi-opc` | `pub` | `pub` | `pub` | `pub` |
| `litchi-sign` | `pub(crate)` | `pub(crate)` | `pub(crate)` | `pub(crate)` |
| `litchi-xldm` | `pub(crate)` | `pub(crate)` | `pub(crate)` | `pub(crate)` |
| `xml-minifier` | `pub(crate)` | `pub(crate)` | `pub(crate)` | `pub(crate)` |

Thus the OLE helper's two declarations are the only visibility narrowings in
the five-file set. The amendment restores both and leaves the other four
helpers unchanged.

## Review conclusion and required gates

The source review finds the patch API-correct and behavior-neutral: it restores
external name accessibility without changing generated code inside either
declaration. The result is not a quality or performance result. The root
coordinator must obtain independent review, apply this patch only at the
recorded parent boundary, rerun the relevant quality checks, and bind any
adoption decision to the resulting source census and fresh evidence.
